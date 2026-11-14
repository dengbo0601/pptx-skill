"""Prune unused master/layout/theme chains from a PPTX template.

Keeps only the parts reachable from the retained slides (via each part's
.rels closure) plus presentation-level parts (notesMaster, handoutMaster,
tableStyles, pres/view props). Slide masters not referenced by any kept
layout are dropped, together with their layouts, themes, and media.

Also removes docProps/thumbnail.jpeg, which leaks the original deck preview
and serves no purpose in a template shell.

Usage:
    python prune_template.py input.pptx output.pptx [--slides slide1.xml,slide2.xml]
"""

from pathlib import Path
import argparse
import os
import posixpath
import re
import zipfile

from defusedxml import minidom


PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
THUMBNAIL_REL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"

ALWAYS_KEEP_PRESENTATION_RELS = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extLst",
}


def _parse_xml(data):
    return minidom.parseString(data)


def _rels_path(part):
    base, name = posixpath.split(part)
    return posixpath.join(base, "_rels", name + ".rels")


def _resolve(base, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base, target))


def _internal_targets(zf, part):
    """Internal parts referenced by part's .rels (external targets skipped)."""
    rels_path = _rels_path(part)
    targets = []
    if rels_path not in zf.namelist():
        return targets
    doc = _parse_xml(zf.read(rels_path))
    base = posixpath.dirname(part)
    for rel in doc.getElementsByTagName("Relationship"):
        if rel.getAttribute("TargetMode") == "External":
            continue
        target = rel.getAttribute("Target")
        if target:
            targets.append(_resolve(base, target))
    return targets


def _closure(zf, roots):
    """All parts reachable from roots through relationship chains."""
    keep = set(roots)
    stack = list(roots)
    while stack:
        part = stack.pop()
        for target in _internal_targets(zf, part):
            if target not in keep and target in zf.namelist():
                keep.add(target)
                stack.append(target)
    return keep


def _collect_slides(zf, slide_names):
    """Retained slides plus presentation-level parts they imply."""
    # Slide closure covers slides -> layouts -> masters -> themes/media.
    # presentation.xml is kept explicitly (its rels reference every master,
    # so it must not go through the closure; pres rels are managed below).
    keep = _closure(zf, {f"ppt/slides/{n}" for n in slide_names})
    keep.add("ppt/presentation.xml")
    # Package-level parts, referenced from the pack header, not via ppt rels.
    keep.update(
        name for name in zf.namelist()
        if name == "[Content_Types].xml"
        or name == "_rels/.rels"
        or name.startswith("docProps/")
    )

    # Presentation-level rels: keep slides + whitelist; drop other slide
    # masters that no kept layout references.
    pres_rels = "ppt/_rels/presentation.xml.rels"
    doc = _parse_xml(zf.read(pres_rels))
    for rel in list(doc.getElementsByTagName("Relationship")):
        rel_type = rel.getAttribute("Type")
        target = rel.getAttribute("Target")
        resolved = _resolve("ppt", target)
        if rel_type == SLIDE_REL_TYPE:
            if resolved in keep:
                continue
        elif rel_type in ALWAYS_KEEP_PRESENTATION_RELS:
            if resolved in zf.namelist():
                keep.add(resolved)
                if _rels_path(resolved) in zf.namelist():
                    keep.add(_rels_path(resolved))
                keep.update(_internal_targets(zf, resolved))
            continue
        rel.parentNode.removeChild(rel)
    keep.add(pres_rels)
    return keep, doc


def _prune_presentation_xml(files, keep):
    """Drop sldMasterId/sldId entries whose part was pruned."""
    name = "ppt/presentation.xml"
    doc = _parse_xml(files[name])
    rels_doc = _parse_xml(files["ppt/_rels/presentation.xml.rels"])
    target_by_rid = {
        n.getAttribute("Id"): n.getAttribute("Target")
        for n in rels_doc.getElementsByTagName("Relationship")
    }

    def _part_of(rid):
        target = target_by_rid.get(rid)
        return _resolve("ppt", target) if target else None

    for sld_master_id in list(doc.getElementsByTagName("p:sldMasterId")):
        part = _part_of(sld_master_id.getAttribute("r:id"))
        if part and part not in keep:
            sld_master_id.parentNode.removeChild(sld_master_id)

    for sld_id in list(doc.getElementsByTagName("p:sldId")):
        part = _part_of(sld_id.getAttribute("r:id"))
        if part and part not in keep:
            sld_id.parentNode.removeChild(sld_id)

    files[name] = doc.toxml(encoding="utf-8")


def _drop_thumbnail(files, keep):
    name = "docProps/thumbnail.jpeg"
    root_rels = "_rels/.rels"
    if name in files:
        del files[name]
        keep.discard(name)
    if root_rels in files:
        doc = _parse_xml(files[root_rels])
        changed = False
        for rel in list(doc.getElementsByTagName("Relationship")):
            if rel.getAttribute("Type") == THUMBNAIL_REL_TYPE:
                rel.parentNode.removeChild(rel)
                changed = True
        if changed:
            files[root_rels] = doc.toxml(encoding="utf-8")


def _prune_content_types(files, keep):
    name = "[Content_Types].xml"
    doc = _parse_xml(files[name])
    root = doc.documentElement
    for node in list(root.childNodes):
        tag = getattr(node, "tagName", None)
        if tag == "Override":
            part_name = node.getAttribute("PartName").lstrip("/")
            if part_name not in keep:
                root.removeChild(node)
    files[name] = doc.toxml(encoding="utf-8")


def prune_template(input_pptx, output_pptx, slide_names=None):
    input_pptx = Path(input_pptx)
    output_pptx = Path(output_pptx)
    temp_output = output_pptx.with_suffix(".tmp.pptx")

    with zipfile.ZipFile(input_pptx, "r") as zf:
        all_slides = [
            n for n in zf.namelist()
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)
        ]
        if slide_names:
            missing = [s for s in slide_names if f"ppt/slides/{s}" not in zf.namelist()]
            if missing:
                raise ValueError(f"Slides not found in template: {missing}")
        else:
            slide_names = [posixpath.basename(n) for n in sorted(all_slides)]
        if not slide_names:
            raise ValueError("Template has no slides to keep")

        keep, pres_rels_doc = _collect_slides(zf, slide_names)

        files = {}
        removed = []
        for name in zf.namelist():
            if name in keep:
                files[name] = zf.read(name)
            else:
                removed.append(name)
        # Part rels of kept parts are kept via closure only when referenced;
        # every kept part's own rels must survive regardless.
        for part in list(keep):
            rels = _rels_path(part)
            if rels in zf.namelist() and rels not in files:
                files[rels] = zf.read(rels)

        files["ppt/_rels/presentation.xml.rels"] = pres_rels_doc.toxml(encoding="utf-8")
        keep.add("ppt/_rels/presentation.xml.rels")
        _prune_presentation_xml(files, keep)
        _drop_thumbnail(files, keep)
        _prune_content_types(files, keep)

        removed = [n for n in removed if n not in files]
        with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as out:
            for name, data in files.items():
                out.writestr(name, data)

    os.replace(temp_output, output_pptx)
    return removed


def main():
    parser = argparse.ArgumentParser(description="Prune unused master/layout/theme chains from a PPTX template.")
    parser.add_argument("input_pptx", help="Template PPTX to prune")
    parser.add_argument("output_pptx", help="Pruned output PPTX")
    parser.add_argument(
        "--slides",
        default=None,
        help="Comma-separated slide filenames to keep (default: all slides)",
    )
    args = parser.parse_args()

    slides = [s.strip() for s in args.slides.split(",")] if args.slides else None
    removed = prune_template(args.input_pptx, args.output_pptx, slides)
    print(f"Removed {len(removed)} parts:")
    for name in sorted(removed):
        print(f"  - {name}")


if __name__ == "__main__":
    main()
