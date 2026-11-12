import io
import zipfile

from lxml import etree

from .constants import R_NS, W_NS
from .markers import replace_markers_in_paragraph
from .package import ensure_content_type, ensure_relationship, ensure_styles
from .parts import build_footnote_element, build_footnotes_root
from .xml_utils import parse_xml, qn, to_xml_bytes


def _convert_document(data: bytes, on_marker) -> tuple[bytes, dict[str, str]]:
    """document.xml のマーカーを置換し、(新データ, nsmap) を返す"""
    root = parse_xml(data)
    nsmap = {prefix: uri for prefix, uri in root.nsmap.items() if prefix is not None}
    for paragraph in root.iter(qn("p")):
        replace_markers_in_paragraph(paragraph, on_marker)
    return to_xml_bytes(root), nsmap


def apply_footnotes(
    docx_bytes: bytes, mapping: dict[str, str]
) -> tuple[bytes, list[int]]:
    """docx のマーカーを脚注に置換し、(新docx, 使用した注番号) を返す"""
    used_numbers: list[int] = []
    footnote_elements: list[etree._Element] = []

    def _on_marker(key: str) -> int | None:
        if key not in mapping:
            return None
        footnote_id = len(footnote_elements) + 1
        footnote_elements.append(build_footnote_element(footnote_id, mapping[key]))
        used_numbers.append(int(key))
        return footnote_id

    src = zipfile.ZipFile(io.BytesIO(docx_bytes))
    names = set(src.namelist())
    out_buffer = io.BytesIO()
    document_nsmap: dict[str, str] = {"w": W_NS, "r": R_NS}

    with zipfile.ZipFile(out_buffer, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            name = item.filename
            if name == "word/footnotes.xml":
                continue

            data = src.read(name)
            if name == "word/document.xml":
                data, document_nsmap = _convert_document(data, _on_marker)
            if name == "[Content_Types].xml":
                data = ensure_content_type(data)
            if name == "word/_rels/document.xml.rels":
                data = ensure_relationship(data)
            if name == "word/styles.xml":
                data = to_xml_bytes(ensure_styles(parse_xml(data)))

            dst.writestr(item, data)

        if "word/styles.xml" not in names:
            styles_root = etree.Element(qn("styles"), nsmap={"w": W_NS})
            dst.writestr("word/styles.xml", to_xml_bytes(ensure_styles(styles_root)))

        footnotes_root = build_footnotes_root(footnote_elements, document_nsmap)
        dst.writestr("word/footnotes.xml", to_xml_bytes(footnotes_root))

    return out_buffer.getvalue(), used_numbers
