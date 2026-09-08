"""运行：python -m unittest discover -s tests -v。所有测试写入临时目录。"""
import contextlib
import io
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

import clean_patient_records as cr
import generate_data as gen

ROOT = Path(__file__).resolve().parents[1]


def record(pid="TEST001", dates="入院日期:2026-02-20 出院日期:2026-03-05"):
    return (f"姓名:测试甲 性别:男 年龄:40岁 住院号:{pid}\n{dates}\n\n"
            "现病史:既往有高血压病史10年，否认糖尿病史。\n\n"
            "出院诊断:测试病\n\n出院带药:药物甲 5 mg 每日一次")


class ParsingTests(unittest.TestCase):
    def test_dose_formats(self):
        for dose, expected in [("5 mg", "5mg"), ("0.5 g", "0.5g"),
                               ("200 μg", "200μg"), ("5mg", "5mg"),
                               ("1 吸", "1吸"), ("1吸", "1吸"), ("0.5 片", "0.5片")]:
            with self.subTest(dose=dose):
                self.assertEqual(cr.parse_med_line(f"药物甲 {dose} 每日一次"),
                                 ("药物甲", expected, "每日一次"))

    def test_drug_name_and_missing_dose(self):
        self.assertEqual(cr.parse_med_line("碳酸钙 D3 片 600 mg 每日一次"),
                         ("碳酸钙D3片", "600mg", "每日一次"))
        self.assertEqual(cr.parse_med_line("药物甲 每周一次"),
                         ("药物甲", None, "每周一次"))

    def test_histories(self):
        cases = [
            ("既往有高血压病史10年。", ["高血压"]),
            ("既往有高血压病史，否认糖尿病史。", ["高血压"]),
            ("既往有高血压病史10年，否认糖尿病史。", ["高血压"]),
            ("既往有高血压病史10年、糖尿病病史5年。", ["高血压", "糖尿病"]),
            ("既往有脑梗死病史病史 15 余年。", ["脑梗死"]),
            ("既往有脑梗死病史、糖尿病病史 15 余年。", ["脑梗死", "糖尿病"]),
            ("平素体健，否认高血压、糖尿病等慢性病史。", []),
        ]
        for text, expected in cases:
            with self.subTest(text=text):
                self.assertEqual(cr.parse_comorbidities(text), expected)

    def test_advice_is_not_medication(self):
        self.assertEqual(cr.parse_meds("出院医嘱:监测血压 每日一次"), [])
        self.assertEqual(cr.parse_meds("出院医嘱:\n药物甲 5 mg 每日一次\n监测血压 每日一次"),
                         [("药物甲", "5mg", "每日一次")])


class PipelineTests(unittest.TestCase):
    def run_batch(self, cases):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        root = Path(temporary.name)
        data, output = root / "data", root / "output"
        data.mkdir()
        for name, text in cases.items():
            (data / name).write_text(text, encoding="utf-8")
        with patch.object(cr, "DATA_DIR", data), patch.object(cr, "OUTPUT_DIR", output):
            with contextlib.redirect_stdout(io.StringIO()):
                cr.main()
        self.assertEqual(len(list(output.iterdir())), 5)
        wide = pd.read_csv(output / "病历结构化_宽表.csv").set_index("来源文件")
        return wide

    def test_mixed_dates(self):
        wide = self.run_batch({
            "good.txt": record("GOOD"),
            "bad.txt": record("BAD", "入院日期:2026-02-30 出院日期:2026-03-05"),
            "missing.txt": record("MISSING", ""),
        })
        self.assertEqual(len(wide), 3)
        self.assertEqual(wide.loc["good.txt", "质量问题"], "无")
        self.assertEqual(wide.loc["good.txt", "住院天数"], 13)
        self.assertIn("日期无效", wide.loc["bad.txt", "质量问题"])
        self.assertEqual(wide.loc["missing.txt", "质量问题"], "日期缺失")

    def test_all_invalid_dates(self):
        wide = self.run_batch({"bad.txt": record(dates="入院日期:2026-02-30 出院日期:2026-03-05")})
        self.assertIn("日期无效", wide.loc["bad.txt", "质量问题"])

    def test_duplicate_ids_all_flagged(self):
        wide = self.run_batch({"a.txt": record(), "b.txt": record(), "c.txt": record("UNIQUE")})
        for name in ["a.txt", "b.txt"]:
            self.assertIn("住院号重复(TEST001)", wide.loc[name, "质量问题"])
        self.assertEqual(wide.loc["c.txt", "质量问题"], "无")

    def test_current_outputs_reproducible(self):
        # 对整批导出做内容比对，不只验证提取函数；不改仓库 output。
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp)
            with patch.object(cr, "OUTPUT_DIR", output), contextlib.redirect_stdout(io.StringIO()):
                cr.main()
            for source in (ROOT / "output").glob("*.csv"):
                with self.subTest(file=source.name):
                    pd.testing.assert_frame_equal(pd.read_csv(source), pd.read_csv(output / source.name))
            book = "病历结构化结果.xlsx"
            old = pd.read_excel(ROOT / "output" / book, sheet_name=None)
            new = pd.read_excel(output / book, sheet_name=None)
            self.assertEqual(set(old), set(new))
            for sheet in old:
                pd.testing.assert_frame_equal(old[sheet], new[sheet])


class VerifierTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        (self.root / "data").mkdir()
        for name in ["clean_patient_records.py", "verify_against_gold.py"]:
            shutil.copyfile(ROOT / name, self.root / name)
        (self.root / "data" / "sample.txt").write_text(record(), encoding="utf-8")
        self.gold = pd.DataFrame([{"file": "sample.txt", "主要诊断": "测试病",
                                   "合并症": "高血压", "药物": "药物甲"}])

    def run_verifier(self):
        self.gold.to_csv(self.root / "gold_standard.csv", index=False, encoding="utf-8-sig")
        # 实际子进程退出码；从项目以外运行，验证路径不依赖当前目录。
        return subprocess.run([sys.executable, "-B", str(self.root / "verify_against_gold.py")],
                              cwd=self.tmp.name, capture_output=True, encoding="utf-8",
                              env=dict(os.environ, PYTHONIOENCODING="utf-8"))

    def assert_rejected(self, message):
        result = self.run_verifier()
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn(message, result.stdout)
        self.assertNotIn("全部通过", result.stdout)

    def test_valid_non_fifty_dataset(self):
        result = self.run_verifier()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("1/1", result.stdout)

    def test_missing_file(self):
        extra = self.gold.copy()
        extra["file"] = "missing.txt"
        self.gold = pd.concat([self.gold, extra])
        self.assert_rejected("缺少病例文件")

    def test_extra_file(self):
        (self.root / "data" / "extra.txt").write_text(record(), encoding="utf-8")
        self.assert_rejected("病例缺少标准答案")

    def test_duplicate_conflicting_gold(self):
        extra = self.gold.copy()
        extra["主要诊断"] = "错误答案"
        self.gold = pd.concat([self.gold, extra])
        self.assert_rejected("file 重复")

    def test_no_matches(self):
        self.gold.loc[0, "file"] = "other.txt"
        self.assert_rejected("缺少病例文件")

    def test_empty_gold(self):
        self.gold = self.gold.iloc[:0]
        self.assert_rejected("没有完整的可验证样本")

    def test_blank_filename(self):
        self.gold.loc[0, "file"] = "  "
        self.assert_rejected("空文件名")

    def test_missing_column(self):
        self.gold = self.gold.drop(columns="药物")
        self.assert_rejected("缺少列")

    def test_blank_diagnosis(self):
        self.gold.loc[0, "主要诊断"] = ""
        self.assert_rejected("主要诊断存在空值")

    def test_wrong_answer(self):
        self.gold.loc[0, "主要诊断"] = "错误答案"
        self.assert_rejected("金标准=错误答案")

    def test_current_gold(self):
        result = subprocess.run([sys.executable, "-B", str(ROOT / "verify_against_gold.py")],
                                cwd=self.tmp.name, capture_output=True, encoding="utf-8",
                                env=dict(os.environ, PYTHONIOENCODING="utf-8"))
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        count = len(pd.read_csv(ROOT / "gold_standard.csv"))
        self.assertGreater(count, 0)
        for field in ["诊断", "合并症", "药物"]:
            self.assertIn(f"[PASS] {field}：{count}/{count}", result.stdout)


class GeneratorTests(unittest.TestCase):
    def test_shrinking_existing_batch_is_rejected_without_changes(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            with patch.object(gen, "N", 3), contextlib.redirect_stdout(io.StringIO()):
                gen.main(root)
            before = {p.relative_to(root): p.read_bytes() for p in root.rglob("*") if p.is_file()}
            with patch.object(gen, "N", 2), self.assertRaisesRegex(SystemExit, "已有数据"):
                gen.main(root)
            after = {p.relative_to(root): p.read_bytes() for p in root.rglob("*") if p.is_file()}
            self.assertEqual(before, after)

    def test_fresh_batches_reproducible_and_verifiable(self):
        with tempfile.TemporaryDirectory() as tmp:
            roots = [Path(tmp) / "first", Path(tmp) / "second"]
            for root in roots:
                with patch.object(gen, "N", 3), contextlib.redirect_stdout(io.StringIO()):
                    gen.main(root)
                self.assertEqual(len(list((root / "data").glob("*.txt"))), 3)
            snapshots = [{p.relative_to(root): p.read_bytes() for p in root.rglob("*") if p.is_file()}
                         for root in roots]
            self.assertEqual(*snapshots)
            for name in ["clean_patient_records.py", "verify_against_gold.py"]:
                shutil.copyfile(ROOT / name, roots[0] / name)
            result = subprocess.run([sys.executable, "-B", str(roots[0] / "verify_against_gold.py")],
                                    capture_output=True, encoding="utf-8",
                                    env=dict(os.environ, PYTHONIOENCODING="utf-8"))
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_invalid_count_writes_nothing(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "batch"
            with patch.object(gen, "N", 0), self.assertRaisesRegex(SystemExit, "正整数"):
                gen.main(root)
            self.assertFalse(root.exists())


if __name__ == "__main__":
    unittest.main()
