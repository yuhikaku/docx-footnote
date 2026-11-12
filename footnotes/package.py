from lxml import etree

from .constants import CT_NS, FOOTNOTES_CT, FOOTNOTES_REL_TYPE, REL_NS
from .xml_utils import parse_xml, qn, to_xml_bytes


def ensure_content_type(data: bytes) -> bytes:
    """[Content_Types].xml に footnotes の Override を追加する"""
    root = parse_xml(data)
    for override in root.findall(f"{{{CT_NS}}}Override"):
        if override.get("PartName") == "/word/footnotes.xml":
            return data
    override = etree.SubElement(root, f"{{{CT_NS}}}Override")
    override.set("PartName", "/word/footnotes.xml")
    override.set("ContentType", FOOTNOTES_CT)
    return to_xml_bytes(root)


def ensure_relationship(data: bytes) -> bytes:
    """document.xml.rels に footnotes の Relationship を追加する"""
    root = parse_xml(data)
    existing_ids = set()
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        existing_ids.add(rel.get("Id"))
        if rel.get("Type") == FOOTNOTES_REL_TYPE:
            return data

    new_id = "rIdFootnotes"
    while new_id in existing_ids:
        new_id += "X"
    rel = etree.SubElement(root, f"{{{REL_NS}}}Relationship")
    rel.set("Id", new_id)
    rel.set("Type", FOOTNOTES_REL_TYPE)
    rel.set("Target", "footnotes.xml")
    return to_xml_bytes(root)


def ensure_styles(root: etree._Element) -> etree._Element:
    """styles に FootnoteText / FootnoteReference が無ければ追加する"""
    existing = {s.get(qn("styleId")) for s in root.findall(qn("style"))}

    if "FootnoteText" not in existing:
        style = etree.SubElement(root, qn("style"))
        style.set(qn("type"), "paragraph")
        style.set(qn("styleId"), "FootnoteText")
        etree.SubElement(style, qn("name")).set(qn("val"), "footnote text")
        etree.SubElement(style, qn("basedOn")).set(qn("val"), "Normal")
        rpr = etree.SubElement(style, qn("rPr"))
        etree.SubElement(rpr, qn("sz")).set(qn("val"), "20")

    if "FootnoteReference" not in existing:
        style = etree.SubElement(root, qn("style"))
        style.set(qn("type"), "character")
        style.set(qn("styleId"), "FootnoteReference")
        etree.SubElement(style, qn("name")).set(qn("val"), "footnote reference")
        rpr = etree.SubElement(style, qn("rPr"))
        etree.SubElement(rpr, qn("vertAlign")).set(qn("val"), "superscript")

    return root
