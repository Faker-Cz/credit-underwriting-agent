import argparse
import json
import math
import re
from pathlib import Path

from openpyxl import load_workbook


NUMBER_PATTERN = re.compile(r"^[+-]?(?:\d{1,3}(?:[,，]\d{3})+|\d+)(?:\.\d+)?%?$")
BALANCE_PAIRS = (
    ("资产总计", "负债和所有者权益总计"),
    ("资产总计", "负债及所有者权益总计"),
)


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def normalized_label(value):
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).replace("：", "").replace(":", "")


def find_labels(sheet):
    labels = {}
    duplicates = []
    for row in sheet.iter_rows():
        for cell in row[: min(3, len(row))]:
            label = normalized_label(cell.value)
            if not label:
                continue
            if label in labels and labels[label] != cell.row:
                duplicates.append({"label": label, "rows": [labels[label], cell.row]})
            labels.setdefault(label, cell.row)
    return labels, duplicates


def numeric_columns(sheet, first_row, second_row):
    checks = []
    for column in range(2, sheet.max_column + 1):
        first = sheet.cell(first_row, column).value
        second = sheet.cell(second_row, column).value
        if is_number(first) and is_number(second):
            difference = first - second
            checks.append(
                {
                    "column": column,
                    "first": first,
                    "second": second,
                    "difference": difference,
                    "status": abs(difference) <= 0.01,
                }
            )
    return checks


def main():
    parser = argparse.ArgumentParser(description="审计授信财务汇总工作簿的数值、公式和资产负债平衡。")
    parser.add_argument("workbook")
    parser.add_argument("--output")
    parser.add_argument("--pure-values", action="store_true", help="将任何公式视为交付错误。")
    args = parser.parse_args()

    path = Path(args.workbook)
    workbook = load_workbook(path, data_only=False, read_only=False)
    report = {"workbook": str(path), "pure_values": args.pure_values, "sheets": {}, "issues": []}

    for sheet in workbook.worksheets:
        labels, duplicate_labels = find_labels(sheet)
        text_numbers = []
        formulas = []
        errors = []
        hidden_rows = [index for index, dimension in sheet.row_dimensions.items() if dimension.hidden]
        hidden_columns = [key for key, dimension in sheet.column_dimensions.items() if dimension.hidden]

        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if cell.data_type == "f" or (isinstance(value, str) and value.startswith("=")):
                    formulas.append(cell.coordinate)
                elif isinstance(value, str):
                    stripped = value.strip()
                    if NUMBER_PATTERN.fullmatch(stripped):
                        text_numbers.append({"cell": cell.coordinate, "value": stripped})
                if cell.data_type == "e":
                    errors.append({"cell": cell.coordinate, "value": value})

        sheet_report = {
            "rows": sheet.max_row,
            "columns": sheet.max_column,
            "merged_ranges": [str(item) for item in sheet.merged_cells.ranges],
            "duplicate_labels": duplicate_labels,
            "numeric_values_stored_as_text": text_numbers,
            "formula_cells": formulas,
            "error_cells": errors,
            "hidden_rows": hidden_rows,
            "hidden_columns": hidden_columns,
        }

        for first_label, second_label in BALANCE_PAIRS:
            first = normalized_label(first_label)
            second = normalized_label(second_label)
            if first in labels and second in labels:
                checks = numeric_columns(sheet, labels[first], labels[second])
                sheet_report["balance_checks"] = checks
                if any(not item["status"] for item in checks):
                    report["issues"].append(f"{sheet.title}: 资产负债表不平")
                break

        if text_numbers:
            report["issues"].append(f"{sheet.title}: 存在{len(text_numbers)}个疑似文本数字")
        if errors:
            report["issues"].append(f"{sheet.title}: 存在{len(errors)}个Excel错误值")
        if args.pure_values and formulas:
            report["issues"].append(f"{sheet.title}: 纯数值交付仍含{len(formulas)}个公式")
        report["sheets"][sheet.title] = sheet_report

    output = Path(args.output) if args.output else path.with_suffix(".财务核验.json")
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output": str(output), "issues": report["issues"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
