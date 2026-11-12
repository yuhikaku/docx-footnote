import csv
import sys
from pathlib import Path

import win32com.client

FINISHED_FILE_SUFFIX = "_脚注付"
MISSING_FILE_SUFFIX = "_欠番"


def find_unused_numbers(max_num: int, used: set[int]) -> list[int]:
    """
    max_num までの整数のうち、used に含まれないものを返す
    例: max_num=5, used=[1,3] の場合、[2,4,5] を返す
    例: max_num=5, used=[1,2,3,4,5] の場合、[] を返す
    例: max_num=5, used=[] の場合、[1,2,3,4,5] を返す
    """
    all_range = range(1, max_num + 1)
    return [n for n in all_range if n not in used]


def load_excel(excel_path: str) -> dict[str, str]:
    """
    Excelファイルから脚注マッピングを読み込む
    1列目: 脚注番号、2列目: 脚注テキスト
    """
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    try:
        workbook = excel.Workbooks.Open(excel_path, ReadOnly=True)
        sheet = workbook.Worksheets(1)  # 最初のシートを使用

        # 使用範囲を取得
        used_range = sheet.UsedRange
        rows = used_range.Rows.Count

        d = {}
        for i in range(1, rows + 1):
            key = sheet.Cells(i, 1).Value
            value = sheet.Cells(i, 2).Value

            if key is not None and value is not None:
                # 文字列化して前後の空白を削除
                d[str(int(key)).strip()] = str(value).strip()

        workbook.Close(SaveChanges=False)
        return d

    finally:
        excel.Quit()


def apply_footnotes(docx_path: Path) -> int:
    excel_path = docx_path.with_suffix(".xlsx")
    if not excel_path.exists():
        print(
            f"{docx_path.name}の脚注定義ファイル {excel_path.name} がありません。スキップします。"
        )
        return 0

    footnote_mapping = load_excel(str(excel_path.absolute()))

    output_path = docx_path.with_name(f"{docx_path.stem}{FINISHED_FILE_SUFFIX}.docx")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    footnote_counter = 0

    try:
        doc = word.Documents.Open(str(docx_path.absolute()))
        doc.TrackRevisions = False

        rng = doc.Content
        find = rng.Find
        find.ClearFormatting()
        find.MatchWildcards = True

        footnote_nums: list[int] = []
        left_markers: list[str] = []

        while find.Execute(FindText="【【[0-9]{1,}】】", Forward=True):
            key = str(rng.Text).replace("【", "").replace("】", "")
            footnote_text = footnote_mapping.get(key, None)

            if footnote_text is None:
                left_markers.append(key)
                rng.SetRange(rng.End, doc.Content.End)
                continue

            rng.Delete()
            footnote = doc.Footnotes.Add(Range=rng)
            footnote.Range.Text = footnote_text
            footnote_nums.append(int(key))
            footnote_counter += 1

            rng.SetRange(footnote.Reference.End, doc.Content.End)

        doc.SaveAs2(str(output_path.absolute()))

        if left_markers:
            unique_left = sorted(set(left_markers), key=int)
            print(
                f"==> Excel に注釈がないマーカーを本文に残しました: {', '.join(unique_left)}"
            )

        print(f"==> 処理完了: {output_path.name}")
        print(f"==> 追加した脚注数: {footnote_counter}")

        footnote_nums_in_excel = [int(k) for k in footnote_mapping if k.isdecimal()]
        missing_nums: list[int] = []
        if footnote_nums_in_excel:
            max_num = max(footnote_nums_in_excel)
            missing_nums = find_unused_numbers(max_num, set(footnote_nums))

        if 0 < len(missing_nums):
            print(
                f"==> 元のWord原稿に注番号の抜けが{len(missing_nums)}件ありました。ファイルに出力します"
            )
            skipped = [
                (num, footnote_mapping.get(str(num), "（Excel中になし）"))
                for num in missing_nums
            ]
            out_missing_path = output_path.with_name(
                f"{output_path.stem}{MISSING_FILE_SUFFIX}.txt"
            )
            out_missing_path.write_text(
                "\n".join([f"■{s[0]}\n{s[1]}\n" for s in skipped]), encoding="utf-8"
            )

        doc.Close()
        return footnote_counter

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise

    finally:
        word.Quit()


def main(data_dir: Path) -> None:
    lines = []
    for c in data_dir.glob("**/*.docx"):
        if c.is_dir():
            continue
        if c.name.startswith("~$"):
            continue
        if c.stem.endswith(FINISHED_FILE_SUFFIX):
            continue
        print(f"処理中：{c.name}")
        count = apply_footnotes(c)
        lines.append([c.name, count])

    out_csv_path = data_dir / "report.csv"
    with open(out_csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["ファイル名", "追加した脚注数"])
        writer.writerows(lines)

    input("\n\n処理が完了しました！\n何かキーを押すと終了します")


if __name__ == "__main__":
    args = sys.argv[1:]
    data_dir = Path("data") if len(args) < 1 else Path(args[0])
    main(data_dir)
