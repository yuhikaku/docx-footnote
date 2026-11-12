from lxml import etree

from .constants import XML_NS
from .xml_utils import qn


def build_footnote_element(footnote_id: int, text: str) -> etree._Element:
    """footnotes.xml に追加する w:footnote 要素を作る"""
    footnote = etree.Element(qn("footnote"))
    footnote.set(qn("id"), str(footnote_id))
    p = etree.SubElement(footnote, qn("p"))
    ppr = etree.SubElement(p, qn("pPr"))
    etree.SubElement(ppr, qn("pStyle")).set(qn("val"), "FootnoteText")

    ref_run = etree.SubElement(p, qn("r"))
    ref_rpr = etree.SubElement(ref_run, qn("rPr"))
    etree.SubElement(ref_rpr, qn("rStyle")).set(qn("val"), "FootnoteReference")
    etree.SubElement(ref_run, qn("footnoteRef"))

    text_run = etree.SubElement(p, qn("r"))
    t = etree.SubElement(text_run, qn("t"))
    t.text = f" {text}"
    t.set(f"{{{XML_NS}}}space", "preserve")
    return footnote


def build_footnotes_root(
    footnote_elements: list[etree._Element], nsmap: dict[str, str]
) -> etree._Element:
    """document.xml と同じ nsmap を持つ footnotes ルートを作る"""
    root = etree.Element(qn("footnotes"), nsmap=nsmap)
    for footnote_type, footnote_id, child_tag in (
        ("separator", "-1", "separator"),
        ("continuationSeparator", "0", "continuationSeparator"),
    ):
        footnote = etree.SubElement(root, qn("footnote"))
        footnote.set(qn("type"), footnote_type)
        footnote.set(qn("id"), footnote_id)
        p = etree.SubElement(footnote, qn("p"))
        r = etree.SubElement(p, qn("r"))
        etree.SubElement(r, qn(child_tag))
    for element in footnote_elements:
        root.append(element)
    return root
