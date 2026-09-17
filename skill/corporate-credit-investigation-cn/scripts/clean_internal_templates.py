"""一次性清理仓库内Word模板的真实信息、固定时点和预填结论。

脚本直接修改OOXML文本节点，不重建文档，尽量保留原模板的段落、表格、分页和对象。
运行前请保留版本库或其他可恢复副本。
"""

import argparse
import re
import tempfile
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DC = "http://purl.org/dc/elements/1.1/"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
STORY_PART = re.compile(r"^word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$")

REPLACEMENTS = [
    (r"浙\s*商\s*银\s*行", "【银行名称】"),
    (r"龙华支行", "【申报单位】"),
    (r"[（(]\s*(?:2018|2022)年版\s*[）)]", "（内部模板）"),
    (r"(?<!\d)1[3-9]\d{9}(?!\d)", "【填写】"),
    (r"2026年7月", "最新一期"),
    (r"截至2025年数据", "截至最近年度数据"),
    (r"2025年度", "最近年度"),
    (r"2023年", "前两年"),
    (r"2024年", "上一年"),
    (r"2025年", "最近年度"),
    (r"这里要列出合并报表的公司，具体看审计报告披露", "【按审计报告披露填写合并范围；不适用则删除或标注不适用】"),
    (r"流动资金类贷款授信额度测算表需要按照新公司数据添加然后粘贴在这", "【如适用，在此插入项目专用测算结果；不适用填“无”】"),
    (r"受信客户最新主体评级\s*[：:]\s*A3", "受信客户最新主体评级：【根据系统评级填写】"),
    (r"授信资金用途用于企业日常经营周转。经调查，xxx公司的企业经营、财务状况、产品和市场情况良好，确认贸易背景真实、合规。", "【根据正式授信方案、贸易背景材料及调查结果填写资金用途和真实性分析】"),
    (r"拟同意给予【受信客户全称】（次）低风险授信额度【金额】亿元，授信方案有效期1年，业务品种为流动性支持类、担保承诺类等产品，单笔业务期限按办法规定执行且不超过1年，授信由本行认可的一、二类合格缓释物有效缓释（含资产池入池缓释物）。", "【待补充：根据正式低风险授信方案和调查事实填写授信结论】"),
    (r"授信方案有效期为1年，额度项下品种为低风险业务，单笔业务期限为1年，授信缓释方式为低风险业务。", "【待补充：根据正式低风险授信方案和调查事实填写授信结论】"),
    (r"拟同意给予xxx（次）低风险授信额度xxx亿元", "拟同意给予【受信客户全称】（次）低风险授信额度【金额】亿元"),
    (r"无异常", "【根据查询结果填写】"),
    (r"债券1", "【融资品种】"),
    (r"(?i)(?<![A-Za-z])x{2,}(?![A-Za-z])", "【填写】"),
    (r"[☑☒■]", "□"),
]


SINGLE_CUSTOMER_TABLE_CORRECTIONS = {
    (12, 0, 0): ("保证人间接融资情况", "受信客户间接融资情况"),
    (12, 2, 0): ("保证人直接融资情况", "受信客户直接融资情况"),
    (21, 8, 0): ("受信客户融资评价", "保证人融资评价"),
    (21, 9, 0): ("受信客户融资变动原因分析", "保证人融资变动原因分析"),
}


def replace_in_paragraph(paragraph, pattern, replacement):
    nodes = list(paragraph.iter(f"{{{W}}}t"))
    if not nodes:
        return 0
    texts = [node.text or "" for node in nodes]
    combined = "".join(texts)
    matches = list(re.finditer(pattern, combined))
    for match in reversed(matches):
        starts = []
        cursor = 0
        for text in texts:
            starts.append(cursor)
            cursor += len(text)
        start_i = max(i for i, offset in enumerate(starts) if offset <= match.start())
        end_i = max(i for i, offset in enumerate(starts) if offset < match.end())
        start_offset = match.start() - starts[start_i]
        end_offset = match.end() - starts[end_i]
        prefix = texts[start_i][:start_offset]
        suffix = texts[end_i][end_offset:]
        texts[start_i] = prefix + replacement + (suffix if start_i == end_i else "")
        for index in range(start_i + 1, end_i):
            texts[index] = ""
        if end_i != start_i:
            texts[end_i] = suffix
    for node, text in zip(nodes, texts):
        node.text = text
    return len(matches)


def paragraph_text(element):
    return "".join(node.text or "" for node in element.iter(f"{{{W}}}t"))


def replace_exact_in_element(element, expected, replacement):
    paragraphs = list(element.iter(f"{{{W}}}p"))
    changes = 0
    for paragraph in paragraphs:
        if paragraph_text(paragraph).strip() == expected:
            changes += replace_in_paragraph(paragraph, re.escape(expected), replacement)
    return changes


def apply_document_corrections(root):
    changes = 0
    body = root.find(f"{{{W}}}body")
    tables = list(body.findall(f"{{{W}}}tbl")) if body is not None else []
    for (table_index, row_index, cell_index), (expected, replacement) in SINGLE_CUSTOMER_TABLE_CORRECTIONS.items():
        if table_index >= len(tables):
            continue
        rows = list(tables[table_index].findall(f"{{{W}}}tr"))
        if row_index >= len(rows):
            continue
        cells = list(rows[row_index].findall(f"{{{W}}}tc"))
        if cell_index < len(cells):
            changes += replace_exact_in_element(cells[cell_index], expected, replacement)

    for paragraph in list(root.iter(f"{{{W}}}p")):
        if paragraph_text(paragraph).strip() != "。":
            continue
        parent = next((node for node in root.iter() if paragraph in list(node)), None)
        if parent is not None:
            parent.remove(paragraph)
            changes += 1
    return changes


def clean_xml(data, document_part=False):
    root = ElementTree.fromstring(data)
    changes = 0
    for paragraph in root.iter(f"{{{W}}}p"):
        for pattern, replacement in REPLACEMENTS:
            changes += replace_in_paragraph(paragraph, pattern, replacement)
    if document_part:
        changes += apply_document_corrections(root)
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True), changes


def clean_core_properties(data):
    root = ElementTree.fromstring(data)
    title = root.find(f"{{{DC}}}title")
    if title is not None:
        title.text = "对公授信调查报告内部清洁模板"
    subject = root.find(f"{{{DC}}}subject")
    if subject is not None:
        subject.text = "仅供内部项目另存使用，不含客户事实"
    last_printed = root.find(f"{{{CP}}}lastPrinted")
    if last_printed is not None:
        root.remove(last_printed)
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)


def clean_document(path):
    path = Path(path)
    with tempfile.NamedTemporaryFile(suffix=".docx", dir=path.parent, delete=False) as handle:
        temporary = Path(handle.name)
    changes = 0
    try:
        with ZipFile(path) as source, ZipFile(temporary, "w", ZIP_DEFLATED) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if STORY_PART.match(item.filename):
                    data, count = clean_xml(data, document_part=item.filename == "word/document.xml")
                    changes += count
                elif item.filename == "docProps/core.xml":
                    data = clean_core_properties(data)
                target.writestr(item, data)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
    return changes


def main():
    parser = argparse.ArgumentParser(description="清理内部授信Word模板中的真实信息、固定时点和预填结论。")
    parser.add_argument("documents", nargs="+")
    args = parser.parse_args()
    for document in args.documents:
        print(f"{document}: {clean_document(document)}处替换")


if __name__ == "__main__":
    main()
