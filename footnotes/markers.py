import copy
from collections.abc import Callable

from lxml import etree

from .constants import MARKER_PATTERN, XML_NS
from .xml_utils import qn


def make_text_run(template_run: etree._Element, text: str) -> etree._Element:
    """template_run の書式を引き継いだテキスト run を作る"""
    run = etree.Element(qn("r"))
    rpr = template_run.find(qn("rPr"))
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    t = etree.SubElement(run, qn("t"))
    t.text = text
    t.set(f"{{{XML_NS}}}space", "preserve")
    return run


def make_reference_run(footnote_id: int) -> etree._Element:
    """脚注参照 run を作る"""
    run = etree.Element(qn("r"))
    rpr = etree.SubElement(run, qn("rPr"))
    etree.SubElement(rpr, qn("rStyle")).set(qn("val"), "FootnoteReference")
    etree.SubElement(run, qn("footnoteReference")).set(qn("id"), str(footnote_id))
    return run


def collect_text_runs(paragraph: etree._Element) -> list[etree._Element]:
    """段落直下の run のうち、w:t のみを持つ単純な run を返す"""
    runs = []
    for run in paragraph.findall(qn("r")):
        children = [c for c in run if c.tag != qn("rPr")]
        if not children:
            continue
        if any(c.tag != qn("t") for c in children):
            continue
        runs.append(run)
    return runs


def run_text(run: etree._Element) -> str:
    """run 内の w:t を連結して返す"""
    return "".join(t.text or "" for t in run.findall(qn("t")))


def _split_run(
    run: etree._Element,
    text: str,
    run_start: int,
    marker_start: int,
    marker_end: int,
) -> tuple[str, str]:
    """run のテキストを、マーカー前後の文字列に分ける"""
    before = text[: max(0, marker_start - run_start)]
    after = text[min(len(text), marker_end - run_start) :]
    return before, after


def _replace_one_marker(
    paragraph: etree._Element,
    runs: list[etree._Element],
    texts: list[str],
    span: tuple[int, int],
    footnote_id: int,
) -> None:
    """マーカー1個にかかる run を、前後テキストと脚注参照 run に置換する"""
    start, end = span
    reference_inserted = False
    run_start = 0

    for run, text in zip(runs, texts):
        run_end = run_start + len(text)
        overlaps = run_start < end and start < run_end
        if not overlaps:
            run_start = run_end
            continue

        before, after = _split_run(run, text, run_start, start, end)
        index = paragraph.index(run)
        paragraph.remove(run)

        new_nodes = []
        if before:
            new_nodes.append(make_text_run(run, before))
        if not reference_inserted:
            new_nodes.append(make_reference_run(footnote_id))
            reference_inserted = True
        if after:
            new_nodes.append(make_text_run(run, after))

        for offset, node in enumerate(new_nodes):
            paragraph.insert(index + offset, node)

        run_start = run_end


def replace_markers_in_paragraph(
    paragraph: etree._Element, on_marker: Callable[[str], int | None]
) -> None:
    """
    段落内のマーカーを脚注参照 run に置換する。
    on_marker(key) は脚注IDを返すコールバック。未定義なら None を返し、
    そのマーカーは本文に残す。
    """
    search_from = 0
    while True:
        runs = collect_text_runs(paragraph)
        texts = [run_text(r) for r in runs]

        match = MARKER_PATTERN.search("".join(texts), search_from)
        if match is None:
            return

        footnote_id = on_marker(match.group(1))
        if footnote_id is None:
            search_from = match.end()
            continue

        _replace_one_marker(paragraph, runs, texts, match.span(), footnote_id)
        search_from = match.start()
