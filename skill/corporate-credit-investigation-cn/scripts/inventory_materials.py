import argparse
import hashlib
import json
from pathlib import Path


CATEGORIES = {
    "template": ("模板", "学习版", "授信调查报告", "预审会"),
    "audit": ("审计", "合并", "财报"),
    "statement": ("资产负债表", "利润表", "现金流量表", "三表", "两张财务报表"),
    "detail": ("科目明细", "前五", "上下游", "合同", "订单"),
    "credit": ("授信明细", "融资", "征信", "中征码", "评级"),
    "guarantor": ("担保人", "保证人", "抵押", "质押"),
    "corporate": ("营业执照", "章程", "法人", "身份证", "股权", "验资"),
    "policy": ("政策", "准入", "禁入", "制度"),
}


def digest(path):
    hasher = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            hasher.update(block)
    return hasher.hexdigest()


def classify(name):
    return [key for key, words in CATEGORIES.items() if any(word in name for word in words)] or ["other"]


def main():
    parser = argparse.ArgumentParser(description="生成授信项目材料清单和重复文件识别底稿。")
    parser.add_argument("root")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    root = Path(args.root).resolve()
    records = []
    digests = {}

    for path in sorted(item for item in root.rglob("*") if item.is_file()):
        checksum = digest(path)
        digests.setdefault(checksum, []).append(str(path.relative_to(root)))
        records.append(
            {
                "path": str(path),
                "relative_path": str(path.relative_to(root)),
                "extension": path.suffix.lower(),
                "size": path.stat().st_size,
                "modified_time": path.stat().st_mtime,
                "sha256": checksum,
                "categories": classify(path.name),
            }
        )

    duplicates = [items for items in digests.values() if len(items) > 1]
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(
            {"root": str(root), "count": len(records), "duplicate_groups": duplicates, "files": records},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"indexed {len(records)} files, {len(duplicates)} duplicate groups -> {output}")


if __name__ == "__main__":
    main()
