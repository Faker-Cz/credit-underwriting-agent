import argparse
from collections import Counter
import json
import re
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn


PLACEHOLDERS = ("【待补充", "【待核验", "粘贴处", "XXX", "待定")
VAGUE_PHRASES = ("以实际为准", "以最终为准", "有望", "具备一定", "需持续关注", "待核验")
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def package_stats(path):
    with ZipFile(path) as archive:
        names = set(archive.namelist())

        def count_parts(prefix):
            return len([name for name in names if name.startswith(prefix) and not name.endswith("/")])

        comments = 0
        comment_ids = []
        if "word/comments.xml" in names:
            root = ElementTree.fromstring(archive.read("word/comments.xml"))
            comment_nodes = root.findall(f".//{{{WORD_NAMESPACE}}}comment")
            comments = len(comment_nodes)
            comment_ids = [node.get(f"{{{WORD_NAMESPACE}}}id") for node in comment_nodes]

        comment_anchors = 0
        comment_anchor_ids = []
        comment_end_ids = []
        if "word/document.xml" in names:
            root = ElementTree.fromstring(archive.read("word/document.xml"))
            anchor_nodes = root.findall(f".//{{{WORD_NAMESPACE}}}commentRangeStart")
            end_nodes = root.findall(f".//{{{WORD_NAMESPACE}}}commentRangeEnd")
            comment_anchors = len(anchor_nodes)
            comment_anchor_ids = [node.get(f"{{{WORD_NAMESPACE}}}id") for node in anchor_nodes]
            comment_end_ids = [node.get(f"{{{WORD_NAMESPACE}}}id") for node in end_nodes]

        return {
            "comments": comments,
            "comment_anchors": comment_anchors,
            "comment_ids": comment_ids,
            "comment_anchor_ids": comment_anchor_ids,
            "comment_end_ids": comment_end_ids,
            "headers": count_parts("word/header"),
            "footers": count_parts("word/footer"),
            "media_files": count_parts("word/media/"),
            "embedded_objects": count_parts("word/embeddings/"),
        }


def all_paragraphs(document):
    yield from document.paragraphs
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    for section in document.sections:
        yield from section.header.paragraphs
        yield from section.footer.paragraphs


def run_format_stats(document):
    fonts = Counter()
    sizes = Counter()
    bold = Counter()
    text_runs = 0
    for paragraph in all_paragraphs(document):
        for run in paragraph.runs:
            if not run.text.strip():
                continue
            text_runs += 1
            properties = run._r.rPr
            font_values = []
            if properties is not None and properties.rFonts is not None:
                for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                    font_values.append(properties.rFonts.get(qn(f"w:{key}")) or "继承")
            else:
                font_values = ["继承"] * 4
            fonts["|".join(font_values)] += 1
            sizes[str(run.font.size.pt if run.font.size else "继承")] += 1
            bold[str(run.bold if run.bold is not None else "继承")] += 1
    return {
        "text_runs": text_runs,
        "font_mappings": dict(fonts.most_common()),
        "font_sizes_pt": dict(sizes.most_common()),
        "bold_values": dict(bold.most_common()),
        "note": "字体统计用于定位异常运行；标题、正文和特殊对象是否适用统一字体仍须结合模板判断。",
    }


def dominant_key(counter_dict):
    if not counter_dict:
        return None
    return max(counter_dict, key=counter_dict.get)


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
        "package_stats": package_stats(path),
        "run_format_stats": run_format_stats(document),
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
        baseline_package_stats = package_stats(baseline_path)
        baseline_run_format_stats = run_format_stats(baseline)
        report["baseline"] = str(baseline_path)
        report["baseline_stats"] = baseline_stats
        report["baseline_package_stats"] = baseline_package_stats
        report["baseline_run_format_stats"] = baseline_run_format_stats
        for key in ("paragraphs", "tables", "images", "sections", "characters"):
            current_value = report["stats"][key]
            baseline_value = baseline_stats[key]
            if current_value < baseline_value:
                report["warnings"].append(f"{key}由{baseline_value}减少为{current_value}，请确认未误删内容")
        for key in ("comments", "comment_anchors", "headers", "footers", "media_files", "embedded_objects"):
            current_value = report["package_stats"][key]
            baseline_value = baseline_package_stats[key]
            if current_value < baseline_value:
                report["warnings"].append(f"{key}由{baseline_value}减少为{current_value}，请确认批注或文档对象未丢失")
        for key in ("comment_ids", "comment_anchor_ids", "comment_end_ids"):
            current_ids = report["package_stats"][key]
            baseline_ids = baseline_package_stats[key]
            missing_ids = [item for item in baseline_ids if item not in current_ids]
            if missing_ids:
                report["warnings"].append(
                    f"{key}缺少基准ID: {', '.join(missing_ids)}，请确认原批注及锚点未被删除或重建"
                )
            elif current_ids != baseline_ids:
                report["warnings"].append(f"{key}顺序发生变化，请逐条核对批注锚点")
        for key, label in (
            ("font_mappings", "主字体映射"),
            ("font_sizes_pt", "主字号"),
            ("bold_values", "主要字重"),
        ):
            baseline_dominant = dominant_key(baseline_run_format_stats[key])
            current_dominant = dominant_key(report["run_format_stats"][key])
            if baseline_dominant != current_dominant:
                report["warnings"].append(
                    f"{label}由{baseline_dominant}变为{current_dominant}，请确认未发生全局格式漂移"
                )

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
