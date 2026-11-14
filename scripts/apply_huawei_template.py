"""Apply the bundled Huawei PPT master shell to a pptxgenjs content deck.

The template shell provides the master/layout chrome (header band, logo,
footer, page number); each content slide's top-level shapes are cloned into
the shell slide.

The shell keeps every resource a content slide references:

- Slide relationships (images, charts, hyperlinks, media, slide jumps) are
  merged into the output slide's .rels with fresh, non-conflicting rIds.
- Referenced parts (media files, chart parts, embedded workbooks, ...) are
  copied into the output package, including their own relationship chains,
  with automatic rename on filename conflicts.
- [Content_Types].xml is extended so every migrated part stays declared.

Usage:
    python apply_huawei_template.py content.pptx output.pptx [--template X.pptx]
"""

from pathlib import Path
import argparse
import os
import posixpath
import re
import zipfile

from defusedxml import minidom


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
DEFAULT_TEMPLATE = SKILL_DIR / "assets" / "huawei_ppt_master.pptx"

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
SLIDE_LAYOUT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

SLIDES_DIR = "ppt/slides"
SHAPE_TAGS = {"p:sp", "p:pic", "p:graphicFrame", "p:cxnSp", "p:grpSp"}

# Last-resort content types for migrated parts declared by neither the
# content package nor the template.
FALLBACK_CONTENT_TYPES = {
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif",
    "webp": "image/webp",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xls": "application/vnd.ms-excel",
}


def _parse_xml(data):
    return minidom.parseString(data)


def _clone(node, doc):
    return doc.importNode(node, deep=True)


def _rid_sort_key(rid):
    match = re.fullmatch(r"rId(\d+)", rid)
    return (int(match.group(1)) if match else 0, rid)


def _slide_numbers(zf):
    nums = []
    for name in zf.namelist():
        match = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name)
        if match:
            nums.append(int(match.group(1)))
    return sorted(nums)


def _next_rid(existing_ids):
    max_id = 0
    for rid in existing_ids:
        match = re.fullmatch(r"rId(\d+)", rid)
        if match:
            max_id = max(max_id, int(match.group(1)))
    return max_id + 1


def _is_rel_attr(attr):
    """True for attributes that hold a relationship id (r:embed, r:id, ...)."""
    return bool(attr.value) and (
        attr.namespaceURI == REL_NS or attr.name.startswith("r:")
    )


def _referenced_rids(doc):
    rids = set()
    for element in doc.getElementsByTagName("*"):
        attributes = element.attributes
        for i in range(attributes.length):
            attr = attributes.item(i)
            if _is_rel_attr(attr):
                rids.add(attr.value)
    return rids


def _rewrite_rids(doc, rid_map):
    if not rid_map:
        return
    for element in doc.getElementsByTagName("*"):
        attributes = element.attributes
        for i in range(attributes.length):
            attr = attributes.item(i)
            if attr.value in rid_map and _is_rel_attr(attr):
                element.setAttribute(attr.name, rid_map[attr.value])


def _resolve_part(base_dir, target):
    """Resolve a relationship Target (relative to base_dir) to a package path."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def _target_from(base_dir, part_path):
    """Inverse of _resolve_part: package path -> Target relative to base_dir."""
    return posixpath.relpath(part_path, base_dir)


def _part_rels_path(part_path):
    base, name = posixpath.split(part_path)
    return posixpath.join(base, "_rels", name + ".rels")


def _unique_part_name(files, desired_path):
    if desired_path not in files:
        return desired_path
    base, name = posixpath.split(desired_path)
    stem, dot, ext = name.rpartition(".")
    if not dot:
        stem, ext = name, ""
    counter = 2
    while True:
        candidate = posixpath.join(base, f"{stem}_c{counter}{ext}")
        if candidate not in files:
            return candidate
        counter += 1


def _scan_content_types(ct_doc):
    defaults = {}
    overrides = {}
    root = ct_doc.documentElement
    for node in root.childNodes:
        tag = getattr(node, "tagName", None)
        if tag == "Default":
            defaults[node.getAttribute("Extension").lower()] = node.getAttribute("ContentType")
        elif tag == "Override":
            overrides[node.getAttribute("PartName")] = node.getAttribute("ContentType")
    return defaults, overrides


class ContentTypes:
    """Mutable [Content_Types].xml for the output package."""

    def __init__(self, ct_xml):
        self.doc = _parse_xml(ct_xml)
        self.root = self.doc.documentElement
        self.defaults = set()
        self.overrides = set()
        for node in self.root.childNodes:
            tag = getattr(node, "tagName", None)
            if tag == "Default":
                self.defaults.add(node.getAttribute("Extension").lower())
            elif tag == "Override":
                self.overrides.add(node.getAttribute("PartName"))

    def drop_slides(self):
        for node in list(self.root.childNodes):
            if getattr(node, "tagName", None) == "Override" and re.fullmatch(
                r"/ppt/slides/slide\d+\.xml", node.getAttribute("PartName")
            ):
                self.root.removeChild(node)
                self.overrides.discard(node.getAttribute("PartName"))

    def add_slide(self, idx):
        part_name = f"/ppt/slides/slide{idx}.xml"
        if part_name in self.overrides:
            return
        node = self.doc.createElementNS(CT_NS, "Override")
        node.setAttribute("PartName", part_name)
        node.setAttribute("ContentType", SLIDE_CT)
        self.root.appendChild(node)
        self.overrides.add(part_name)

    def _add_default(self, ext, content_type):
        node = self.doc.createElementNS(CT_NS, "Default")
        node.setAttribute("Extension", ext)
        node.setAttribute("ContentType", content_type)
        first_override = next(
            (n for n in self.root.childNodes if getattr(n, "tagName", None) == "Override"),
            None,
        )
        self.root.insertBefore(node, first_override)
        self.defaults.add(ext)

    def _add_override(self, part_name, content_type):
        node = self.doc.createElementNS(CT_NS, "Override")
        node.setAttribute("PartName", part_name)
        node.setAttribute("ContentType", content_type)
        self.root.appendChild(node)
        self.overrides.add(part_name)

    def ensure(self, part_path, content_part_path, content_defaults, content_overrides):
        """Guarantee a content-type declaration for a migrated part.

        A specific Override from the content package always wins: a generic
        Default for the extension (e.g. xml -> application/xml) does not type
        chart/diagram parts correctly, so never treat it as covering those.
        """
        part_name = "/" + part_path
        if part_name in self.overrides:
            return
        if part_name in content_overrides:
            self._add_override(part_name, content_overrides[part_name])
            return
        basename = posixpath.basename(part_path)
        ext = basename.rpartition(".")[2].lower() if "." in basename else ""
        if ext and ext in self.defaults:
            return
        if ext and ext in content_defaults:
            self._add_default(ext, content_defaults[ext])
        elif ext and ext in FALLBACK_CONTENT_TYPES:
            self._add_default(ext, FALLBACK_CONTENT_TYPES[ext])
        else:
            raise ValueError(f"Cannot determine content type for migrated part {part_path}")

    def serialize(self):
        return self.doc.toxml(encoding="utf-8")


def _copy_part(content_zf, content_part, files, rename_map, ct, content_defaults, content_overrides):
    """Copy a part, its .rels, and every internal part it references.

    Identical-name-identical-bytes collisions with the template reuse the
    template part; differing bytes get a new name so template media is never
    clobbered. Returns the output package path of the part.
    """
    if content_part in rename_map:
        return rename_map[content_part]
    if content_part not in content_zf.namelist():
        raise KeyError(f"Content package is missing referenced part: {content_part}")

    data = content_zf.read(content_part)
    if content_part in files:
        if files[content_part] == data:
            out_part = content_part  # byte-identical: safe to reuse
        else:
            out_part = _unique_part_name(files, content_part)
            files[out_part] = data
    else:
        out_part = content_part
        files[out_part] = data
    rename_map[content_part] = out_part
    ct.ensure(out_part, content_part, content_defaults, content_overrides)

    content_rels_path = _part_rels_path(content_part)
    if content_rels_path in content_zf.namelist():
        rels_doc = _parse_xml(content_zf.read(content_rels_path))
        base = posixpath.dirname(content_part)
        for rel in list(rels_doc.getElementsByTagName("Relationship")):
            if rel.getAttribute("TargetMode") == "External":
                continue
            target = rel.getAttribute("Target")
            if not target:
                continue
            ref_part = _resolve_part(base, target)
            ref_out = _copy_part(
                content_zf, ref_part, files, rename_map, ct,
                content_defaults, content_overrides,
            )
            rel.setAttribute("Target", _target_from(posixpath.dirname(out_part), ref_out))
        new_rels = rels_doc.toxml(encoding="utf-8")
        out_rels_path = _part_rels_path(out_part)
        if files.get(out_rels_path) != new_rels:
            files[out_rels_path] = new_rels
    return out_part


def _merge_slide(
    template_slide_xml,
    template_rels_xml,
    content_slide_xml,
    content_rels_xml,
    content_zf,
    files,
    rename_map,
    ct,
    content_defaults,
    content_overrides,
    slide_num_to_idx,
):
    """Clone content shapes into the shell slide and merge its relationships."""
    template_doc = _parse_xml(template_slide_xml)
    content_doc = _parse_xml(content_slide_xml)

    sp_tree = template_doc.getElementsByTagName("p:spTree")[0]
    for node in list(sp_tree.childNodes):
        if getattr(node, "tagName", None) not in {"p:nvGrpSpPr", "p:grpSpPr"}:
            sp_tree.removeChild(node)

    content_sp_tree = content_doc.getElementsByTagName("p:spTree")[0]
    for node in content_sp_tree.childNodes:
        if getattr(node, "tagName", None) in SHAPE_TAGS:
            sp_tree.appendChild(_clone(node, template_doc))

    rels_doc = _parse_xml(template_rels_xml)
    rel_root = rels_doc.documentElement
    template_rids = [
        node.getAttribute("Id")
        for node in rel_root.childNodes
        if getattr(node, "tagName", None) == "Relationship"
    ]

    content_rels = {}
    if content_rels_xml is not None:
        content_rels_doc = _parse_xml(content_rels_xml)
        for node in content_rels_doc.getElementsByTagName("Relationship"):
            content_rels[node.getAttribute("Id")] = {
                "type": node.getAttribute("Type"),
                "target": node.getAttribute("Target"),
                "external": node.getAttribute("TargetMode") == "External",
            }

    rid_map = {}
    rid_num = _next_rid(template_rids)
    for rid in sorted(_referenced_rids(template_doc), key=_rid_sort_key):
        rel = content_rels.get(rid)
        # The output slide inherits the template's own layout relationship;
        # every other referenced relationship must be merged with a fresh id
        # (content rIds may numerically collide with template ones).
        if rel is None or rel["type"] == SLIDE_LAYOUT_REL_TYPE:
            continue
        new_rid = f"rId{rid_num}"
        rid_num += 1
        rid_map[rid] = new_rid

        new_rel = rels_doc.createElementNS(PKG_REL_NS, "Relationship")
        new_rel.setAttribute("Id", new_rid)
        new_rel.setAttribute("Type", rel["type"])
        if rel["external"]:
            new_rel.setAttribute("Target", rel["target"])
            new_rel.setAttribute("TargetMode", "External")
        elif rel["type"] == SLIDE_REL_TYPE:
            # Internal slide jump: retarget to the renumbered output slide.
            match = re.fullmatch(r"slide(\d+)\.xml", posixpath.basename(rel["target"]))
            out_idx = slide_num_to_idx.get(int(match.group(1))) if match else None
            if out_idx is None:
                raise ValueError(f"Slide jump target {rel['target']} is not a slide of the content deck")
            new_rel.setAttribute("Target", _target_from(SLIDES_DIR, f"ppt/slides/slide{out_idx}.xml"))
        else:
            content_part = _resolve_part(SLIDES_DIR, rel["target"])
            out_part = _copy_part(
                content_zf, content_part, files, rename_map, ct,
                content_defaults, content_overrides,
            )
            new_rel.setAttribute("Target", _target_from(SLIDES_DIR, out_part))
        rel_root.appendChild(new_rel)

    _rewrite_rids(template_doc, rid_map)

    return template_doc.toxml(encoding="utf-8"), rels_doc.toxml(encoding="utf-8")


def _update_presentation(files, slide_count):
    doc = _parse_xml(files["ppt/presentation.xml"])
    rels_doc = _parse_xml(files["ppt/_rels/presentation.xml.rels"])

    sld_id_lst = doc.getElementsByTagName("p:sldIdLst")[0]
    for node in list(sld_id_lst.childNodes):
        sld_id_lst.removeChild(node)

    rel_root = rels_doc.documentElement
    for node in list(rel_root.childNodes):
        if getattr(node, "tagName", None) == "Relationship" and node.getAttribute("Type") == SLIDE_REL_TYPE:
            rel_root.removeChild(node)

    existing_ids = [
        node.getAttribute("Id")
        for node in rel_root.childNodes
        if getattr(node, "tagName", None) == "Relationship"
    ]
    rid_num = _next_rid(existing_ids)

    for idx in range(1, slide_count + 1):
        rid = f"rId{rid_num + idx - 1}"

        sld_id = doc.createElementNS(P_NS, "p:sldId")
        sld_id.setAttribute("id", str(280 + idx))
        sld_id.setAttributeNS(REL_NS, "r:id", rid)
        sld_id_lst.appendChild(sld_id)

        rel = rels_doc.createElementNS(PKG_REL_NS, "Relationship")
        rel.setAttribute("Id", rid)
        rel.setAttribute("Type", SLIDE_REL_TYPE)
        rel.setAttribute("Target", f"slides/slide{idx}.xml")
        rel_root.appendChild(rel)

    files["ppt/presentation.xml"] = doc.toxml(encoding="utf-8")
    files["ppt/_rels/presentation.xml.rels"] = rels_doc.toxml(encoding="utf-8")


def _update_app_props(files, slide_count):
    name = "docProps/app.xml"
    if name not in files:
        return
    doc = _parse_xml(files[name])
    slides = doc.getElementsByTagName("Slides")
    if slides and slides[0].firstChild:
        slides[0].firstChild.nodeValue = str(slide_count)
    files[name] = doc.toxml(encoding="utf-8")


def apply_template(content_pptx, output_pptx, template_pptx=DEFAULT_TEMPLATE):
    content_pptx = Path(content_pptx)
    output_pptx = Path(output_pptx)
    template_pptx = Path(template_pptx)
    temp_output = output_pptx.with_suffix(".tmp.pptx")

    with zipfile.ZipFile(template_pptx, "r") as template_zf, zipfile.ZipFile(content_pptx, "r") as content_zf:
        files = {name: template_zf.read(name) for name in template_zf.namelist()}
        content_slide_nums = _slide_numbers(content_zf)
        if not content_slide_nums:
            raise ValueError(f"No slides found in {content_pptx}")

        ct = ContentTypes(files["[Content_Types].xml"])
        ct.drop_slides()
        content_ct_doc = _parse_xml(content_zf.read("[Content_Types].xml"))
        content_defaults, content_overrides = _scan_content_types(content_ct_doc)

        rename_map = {}
        slide_num_to_idx = {num: idx for idx, num in enumerate(content_slide_nums, start=1)}

        template_slide = template_zf.read("ppt/slides/slide1.xml")
        template_rels = template_zf.read("ppt/slides/_rels/slide1.xml.rels")

        for name in list(files.keys()):
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name) or re.fullmatch(r"ppt/slides/_rels/slide\d+\.xml\.rels", name):
                del files[name]

        for idx, slide_num in enumerate(content_slide_nums, start=1):
            content_rels_path = f"ppt/slides/_rels/slide{slide_num}.xml.rels"
            content_rels = (
                content_zf.read(content_rels_path)
                if content_rels_path in content_zf.namelist()
                else None
            )
            slide_xml, slide_rels = _merge_slide(
                template_slide,
                template_rels,
                content_zf.read(f"ppt/slides/slide{slide_num}.xml"),
                content_rels,
                content_zf,
                files,
                rename_map,
                ct,
                content_defaults,
                content_overrides,
                slide_num_to_idx,
            )
            files[f"ppt/slides/slide{idx}.xml"] = slide_xml
            files[f"ppt/slides/_rels/slide{idx}.xml.rels"] = slide_rels
            ct.add_slide(idx)

        _update_presentation(files, len(content_slide_nums))
        files["[Content_Types].xml"] = ct.serialize()
        _update_app_props(files, len(content_slide_nums))

        with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as out:
            for name, data in files.items():
                out.writestr(name, data)

    os.replace(temp_output, output_pptx)


def main():
    parser = argparse.ArgumentParser(description="Apply bundled Huawei PPT master to a pptxgenjs content deck.")
    parser.add_argument("content_pptx", help="Content-only PPTX generated by pptxgenjs")
    parser.add_argument("output_pptx", help="Final Huawei-style PPTX output")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Optional custom Huawei template PPTX")
    args = parser.parse_args()
    apply_template(args.content_pptx, args.output_pptx, args.template)


if __name__ == "__main__":
    main()
