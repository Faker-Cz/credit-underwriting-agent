import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


PLACEHOLDERS = ("【待补充", "【待核验", "粘贴处", "XXX", "待定")
VAGUE_PHRASES = ("以实际为准", "以最终为准", "有望", "具备一定", "需持续关注", "待核验")


def document_stats(document):
    paragraphs = [paragraph.text for paragraph in document.paragraphs]
    cells = [cell.text for table in document.tables for row in table.rows for cell in row.cells]
    return {
        "paragraphs": len(document.paragraphs),
        "tables": len(document.tables),
        "images": len(document.inline_shapes),
        "sections": len(document.sections),
        "characters": len("\n".join(paragraphs + cells)),
    }


def main():
    parser = argparse.ArgumentParser(description="检查授信报告DOCX的残留客户、占位符和分页结构。")
    parser.add_argument("document")
    parser.add_argument("--baseline", help="用户修改基准版，用于检测内容或对象数量异常减少。")
    parser.add_argument("--forbidden", action="append", default=[])
    parser.add_argument("--output")
    args = parser.parse_args()

    path = Path(args.document)
    document = Document(path)
    paragraphs = [paragraph.text for paragraph in document.paragraphs]
    cells = [cell.text for table in document.tables for row in table.rows for cell in row.cells]
    full_text = "\n".join(paragraphs + cells)
    leftovers = {word: full_text.count(word) for word in args.forbidden if word and word in full_text}
    placeholders = {word: full_text.count(word) for word in PLACEHOLDERS if word in full_text}
    vague_phrases = {word: full_text.count(word) for word in VAGUE_PHRASES if word in full_text}
    empty_heading_like = []

    for index, text in enumerate(paragraphs):
        stripped = text.strip()
        if re.fullmatch(r"[（(]?[一二三四五六七八九十0-9]+[）).、]", stripped):
            empty_heading_like.append({"paragraph": index, "text": stripped})

    page_break_before = []
    keep_next = []
    keep_lines = []
    explicit_breaks = 0
    for index, paragraph in enumerate(document.paragraphs):
        if paragraph.paragraph_format.page_break_before:
            page_break_before.append(index)
        if paragraph.paragraph_format.keep_with_next:
            keep_next.append(index)
        if paragraph.paragraph_format.keep_together:
            keep_lines.append(index)
        explicit_breaks += len(paragraph._p.xpath('.//w:br[@w:type="page"]'))

    cant_split_rows = 0
    fixed_height_rows = 0
    for table in document.tables:
        for row in table.rows:
            properties = row._tr.find(qn("w:trPr"))
            if properties is not None and properties.find(qn("w:cantSplit")) is not None:
                cant_split_rows += 1
            if properties is not None and properties.find(qn("w:trHeight")) is not None:
                fixed_height_rows += 1

    report = {
        "document": str(path),
        "stats": document_stats(document),
        "explicit_page_breaks": explicit_breaks,
        "page_break_before_paragraphs": page_break_before,
        "keep_with_next_paragraphs": keep_next,
        "keep_together_paragraphs": keep_lines,
        "cant_split_table_rows": cant_split_rows,
        "fixed_height_table_rows": fixed_height_rows,
        "forbidden_leftovers": leftovers,
        "placeholders": placeholders,
        "vague_phrases": vague_phrases,
        "empty_heading_like": empty_heading_like,
        "warnings": [],
        "note": "结构检查不能替代逐页渲染检查。",
    }

    if args.baseline:
        baseline_path = Path(args.baseline)
        baseline = Document(baseline_path)
        baseline_stats = document_stats(baseline)
        report["baseline"] = str(baseline_path)
        report["baseline_stats"] = baseline_stats
        for key in ("paragraphs", "tables", "images", "sections", "characters"):
            current_value = report["stats"][key]
            baseline_value = baseline_stats[key]
            if current_value < baseline_value:
                report["warnings"].append(f"{key}由{baseline_value}减少为{current_value}，请确认未误删内容")

    output = Path(args.output) if args.output else path.with_suffix(".文档核验.json")
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "output": str(output),
                "forbidden_leftovers": leftovers,
                "placeholders": placeholders,
                "warnings": report["warnings"],
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
