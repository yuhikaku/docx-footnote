from lxml import etree

from .constants import W_NS


def qn(tag: str) -> str:
    """w: 名前空間付きのタグ名を返す"""
    return f"{{{W_NS}}}{tag}"


def parse_xml(data: bytes) -> etree._Element:
    """空白を保持したまま XML を解析する"""
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    return etree.fromstring(data, parser)


def to_xml_bytes(root: etree._Element) -> bytes:
    """XML 宣言付きの bytes にシリアライズする"""
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
