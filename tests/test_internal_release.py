import hashlib
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "skill" / "corporate-credit-investigation-cn"
ASSETS = SKILL / "assets"


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


class InternalReleaseTests(unittest.TestCase):
    def test_root_and_skill_assets_match(self):
        for name in ("单一客户授信调查报告模版.docx", "低风险报告模版.docx", "流动资金类贷款授信额度测算表.xlsx"):
            self.assertEqual(sha256(ROOT / name), sha256(ASSETS / name), name)

    def test_docx_templates_are_clean(self):
        script = SKILL / "scripts" / "audit_docx.py"
        with tempfile.TemporaryDirectory() as directory:
            for name in ("单一客户授信调查报告模版.docx", "低风险报告模版.docx"):
                result = subprocess.run(
                    [sys.executable, str(script), str(ASSETS / name), "--mode", "template", "--strict", "--forbidden", "浙商银行", "--forbidden", "龙华支行", "--output", str(Path(directory) / f"{name}.json")],
                    capture_output=True,
                    text=True,
                )
                self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_loan_workbook_contract(self):
        script = SKILL / "scripts" / "audit_financial_workbook.py"
        with tempfile.TemporaryDirectory() as directory:
            result = subprocess.run(
                [sys.executable, str(script), str(ASSETS / "流动资金类贷款授信额度测算表.xlsx"), "--mode", "loan-template", "--strict", "--output", str(Path(directory) / "audit.json")],
                capture_output=True,
                text=True,
            )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_canonical_skill_name(self):
        guide = (ROOT / "授信报告智能体使用说明与操作指引.md").read_text(encoding="utf-8")
        self.assertNotIn("bank-corporate-credit-investigation-cn", guide)
        self.assertIn("$corporate-credit-investigation-cn", guide)


if __name__ == "__main__":
    unittest.main()
