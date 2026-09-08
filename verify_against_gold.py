# -*- coding: utf-8 -*-
"""
一键验证脚本：程序提取结果 vs gold_standard.csv
================================================
用法（项目目录下）：
    python verify_against_gold.py

做什么：
    逐份病历跑 process_file，与金标准的 主要诊断 / 合并症 / 药物 三项对拍
    （集合比较：去空白、不看顺序），输出通过率与不一致的详细清单。

退出码：
    0 = 文件覆盖完整、标准答案有效、三项全部一致
    1 = 输入不完整、读取失败或内容不一致（打印具体原因）
"""

import re
import sys
from pathlib import Path

import pandas as pd

import clean_patient_records as cr

BASE = Path(__file__).resolve().parent


def norm(s):
    """比较口径：去掉所有空白，只比实质内容。"""
    return re.sub(r"\s+", "", str(s))


def main():
    try:
        gold = pd.read_csv(BASE / "gold_standard.csv", encoding="utf-8-sig",
                           dtype=str, keep_default_na=False)
    except (OSError, UnicodeError, pd.errors.ParserError, pd.errors.EmptyDataError) as exc:
        print(f"[FAIL] 标准答案读取失败：{exc}")
        sys.exit(1)
    files = sorted((BASE / "data").glob("*.txt"))
    required = {"file", "主要诊断", "合并症", "药物"}
    missing_columns = required - set(gold.columns)
    if missing_columns:
        print(f"[FAIL] 标准答案缺少列：{', '.join(sorted(missing_columns))}")
        sys.exit(1)

    errors = []
    if not files or gold.empty:
        errors.append("病例或标准答案为空，没有完整的可验证样本")
    if gold["file"].str.strip().eq("").any():
        errors.append("标准答案 file 存在空文件名")
    duplicates = gold.loc[gold["file"].duplicated(keep=False), "file"].unique()
    if len(duplicates):
        errors.append(f"标准答案 file 重复：{', '.join(duplicates)}")
    if gold["主要诊断"].str.strip().eq("").any():
        errors.append("标准答案主要诊断存在空值")
    actual_names = {path.name for path in files}
    expected_names = set(gold["file"])
    if expected_names - actual_names:
        errors.append(f"缺少病例文件：{', '.join(sorted(expected_names - actual_names))}")
    if actual_names - expected_names:
        errors.append(f"病例缺少标准答案：{', '.join(sorted(actual_names - expected_names))}")
    if errors:
        for error in errors:
            print(f"[FAIL] {error}")
        sys.exit(1)
    gold = gold.set_index("file")  # 上面已拒绝重复文件名

    stats = {"诊断": 0, "合并症": 0, "药物": 0}
    details = []          # 收集所有不一致：(文件名, 项目, 金标准, 提取)
    n = 0
    for path in files:
        row = gold.loc[path.name]
        n += 1
        try:
            rec = cr.process_file(path)
        except (OSError, UnicodeError, ValueError) as exc:
            print(f"[FAIL] 病例处理失败：{path.name}：{exc}")
            sys.exit(1)

        g_diag = norm(row["主要诊断"])
        g_combo = {norm(x) for x in str(row["合并症"]).split("；")
                   if norm(x)}
        g_med = {norm(x) for x in str(row["药物"]).split("；")
                 if norm(x)}

        d_list = [norm(x) for x in rec["_诊断列表"]]
        c_set = {norm(x) for x in rec["_合并症列表"]}
        m_set = {norm(x[0]) for x in rec["_用药明细"]}

        ok_diag = bool(d_list) and d_list[0] == g_diag
        ok_combo = c_set == g_combo
        ok_med = m_set == g_med

        stats["诊断"] += ok_diag
        stats["合并症"] += ok_combo
        stats["药物"] += ok_med
        if not ok_diag:
            details.append((path.name, "主要诊断", g_diag, "、".join(d_list)))
        if not ok_combo:
            details.append((path.name, "合并症", "、".join(sorted(g_combo)),
                            "、".join(sorted(c_set)) if c_set else "(空)"))
        if not ok_med:
            details.append((path.name, "药物", "、".join(sorted(g_med)),
                            "、".join(sorted(m_set)) if m_set else "(空)"))

    print(f"共对比 {n} 份（病例与标准答案一一对应）\n")
    for k, v in stats.items():
        mark = "PASS" if v == n else "FAIL"
        print(f"[{mark}] {k}：{v}/{n}")
    print()

    if details:
        print(f"共 {len(details)} 处不一致：")
        for f, item, gold_v, got_v in details:
            print(f"  {f} [{item}] 金标准={gold_v} | 提取={got_v}")
        sys.exit(1)
    print("全部通过：当前数据的主要诊断、合并症集合、药名集合均与标准答案一致")
    sys.exit(0)


if __name__ == "__main__":
    main()
