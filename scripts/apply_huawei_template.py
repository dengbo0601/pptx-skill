from pathlib import Path
import argparse
import os
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
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"


def _parse_xml(data):
    return minidom.parseString(data)


def _clone(node, doc):
    return doc.importNode(node, deep=True)


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


def _merge_slide(template_slide_xml, content_slide_xml):
    template_doc = _parse_xml(template_slide_xml)
    content_doc = _parse_xml(content_slide_xml)

    sp_tree = template_doc.getElementsByTagName("p:spTree")[0]
    for node in list(sp_tree.childNodes):
        if getattr(node, "tagName", None) not in {"p:nvGrpSpPr", "p:grpSpPr"}:
            sp_tree.removeChild(node)

    content_sp_tree = content_doc.getElementsByTagName("p:spTree")[0]
    for node in content_sp_tree.childNodes:
        if getattr(node, "tagName", None) in {"p:sp", "p:pic", "p:graphicFrame", "p:cxnSp", "p:grpSp"}:
            sp_tree.appendChild(_clone(node, template_doc))

    return template_doc.toxml(encoding="utf-8")


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


def _update_content_types(files, slide_count):
    doc = _parse_xml(files["[Content_Types].xml"])
    root = doc.documentElement

    for node in list(root.childNodes):
        if (
            getattr(node, "tagName", None) == "Override"
            and re.fullmatch(r"/ppt/slides/slide\d+\.xml", node.getAttribute("PartName"))
        ):
            root.removeChild(node)

    for idx in range(1, slide_count + 1):
        override = doc.createElementNS(CT_NS, "Override")
        override.setAttribute("PartName", f"/ppt/slides/slide{idx}.xml")
        override.setAttribute("ContentType", SLIDE_CT)
        root.appendChild(override)

    files["[Content_Types].xml"] = doc.toxml(encoding="utf-8")


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

        template_slide = template_zf.read("ppt/slides/slide1.xml")
        template_rels = template_zf.read("ppt/slides/_rels/slide1.xml.rels")

        for name in list(files.keys()):
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name) or re.fullmatch(r"ppt/slides/_rels/slide\d+\.xml\.rels", name):
                del files[name]

        for idx, slide_num in enumerate(content_slide_nums, start=1):
            content_slide = content_zf.read(f"ppt/slides/slide{slide_num}.xml")
            files[f"ppt/slides/slide{idx}.xml"] = _merge_slide(template_slide, content_slide)
            files[f"ppt/slides/_rels/slide{idx}.xml.rels"] = template_rels

        _update_presentation(files, len(content_slide_nums))
        _update_content_types(files, len(content_slide_nums))
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
