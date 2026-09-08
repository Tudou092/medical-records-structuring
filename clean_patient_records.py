# -*- coding: utf-8 -*-
"""
病历文本结构化清洗脚本（模拟出院小结 → 结构化表格）
=====================================================
流程：逐份解析字段（基本信息 / 日期 / 诊断 / 合并症 / 用药）→ 汇总为 DataFrame
      → 数据校验与质量问题标记 → 输出宽表 / 长表 / 质量报告 / Excel。

数据说明：data/ 为模拟生成的出院小结，含多种格式变体与脏数据，包括：
  1. 表头 3 种：  "姓名:xx 性别:xx 年龄:xx岁" / "患者 xx,男,52岁" / 住院号在下一行
  2. 日期写法：  2026-10-07 / 2026.02.05 / 2026年3月25日 / "住院时间:... 至 ..."
  3. 诊断写法：  编号列表(1. xxx) / 一句话式(考虑 A、B)
  4. 用药段标题：出院带药: / 带药:(无编号) / 出院医嘱:(混有医嘱话术)
  5. 个别记录缺剂量；部分记录无合并症(如"平素体健")

输出（output/ 目录）：
  - 病历结构化_宽表.csv        一行一个病人（诊断/合并症/用药为、连接的字符串）
  - 病历结构化_合并症长表.csv   一行一条合并症（方便统计频次）
  - 病历结构化_用药长表.csv     一行一条用药（药名/剂量/频次拆开）
  - 数据质量报告.csv            每个病人的质量问题标记
  - 病历结构化结果.xlsx         以上 4 张表合成一个 Excel

隐私保护：输出中姓名一律脱敏（李建国 → 李**），以住院号作为患者标识。
"""

import re
from datetime import datetime
import pandas as pd

# ========== 路径配置（按需修改） ==========
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent   # 脚本所在文件夹，任何机器 clone 下来都能直接跑
DATA_DIR = BASE_DIR / "data"                 # 模拟出院小结所在文件夹（data/）
OUTPUT_DIR = BASE_DIR / "output"             # 结果输出文件夹（output/，自动创建）

# 段落标题清单：用来判断"一个段落到哪里结束"
SECTION_HEADERS = [
    "主诉", "现病史", "入院诊断", "出院诊断", "诊断", "诊疗经过", "治疗经过",
    "住院期间", "出院带药", "带药", "出院医嘱", "出院指导", "注意", "住院号",
]


# ========== 第一步：基础提取函数 ==========
def extract_section(text, title):
    """取某个段落的内容：冒号后面同一行的字 + 下面跟着的行，直到空行或下一个标题。
    兼容"入院诊断:社区获得性肺炎"(内容在同一行)和"出院诊断:"(内容在下面几行)两种。"""
    lines = text.splitlines()
    content, collecting = [], False
    for line in lines:
        stripped = line.strip()
        # 段落开头的三种形态：标准标题、带序号的标题行
        header_match = re.match(rf"^({'|'.join(SECTION_HEADERS)})[：:]\s*(.*)$", stripped)
        if header_match:
            if collecting:
                break  # 碰到下一个标题，本段结束
            if header_match.group(1) == title:
                collecting = True
                if header_match.group(2):  # 同一行就有内容
                    content.append(header_match.group(2))
            continue
        if collecting:
            if not stripped:
                break   # 空行 = 段落结束
            content.append(stripped)
    return content


def parse_header(text):
    """从全文提取 姓名/性别/年龄/住院号。兼容 3 种表头格式。"""
    name = gender = age = pid = None
    # 格式1&2：姓名:李建国  性别:女  年龄:39岁  (住院号可在同行或下一行)
    m = re.search(r"姓名[：:](\S+?)\s+性别[：:](\S+?)\s+年龄[：:](\d+)岁", text)
    if m:
        name, gender, age = m.group(1), m.group(2), int(m.group(3))
    else:
        # 格式3：患者 周欣怡,男,52岁,住院号:ZY100074
        m = re.search(r"患者\s+(\S+?),\s*([男女]),\s*(\d+)岁", text)
        if m:
            name, gender, age = m.group(1), m.group(2), int(m.group(3))
    m = re.search(r"住院号[：:](\w+)", text)
    if m:
        pid = m.group(1)
    return name, gender, age, pid


def parse_dates(text):
    """提取入院/出院日期。兼容 4 种写法，统一成 YYYY-MM-DD：
       2026-10-07 / 2026.02.05 / 2026年3月25日 / "住院时间:... 至 ..." """
    # 日期的通用零件：年-月-日，分隔符可以是 - . 年月，日字结尾可有可无
    DATE = r"(\d{4})[-.年](\d{1,2})[-.月](\d{1,2})日?"

    def norm(y, mo, d):
        """统一成 YYYY-MM-DD。非法日期(如 2026-02-30)返回 None——
        让质量报告标记"日期无效"，而不是让整批任务中断。"""
        y, mo, d = int(y), int(mo), int(d)
        try:
            datetime(y, mo, d)          # 校验合法性，非法日期抛 ValueError
        except ValueError:
            return None
        return f"{y:04d}-{mo:02d}-{d:02d}"

    # 写法1：入院日期:2026-10-07  出院日期:2026-10-11
    m = re.search(rf"入院日期[：:]{DATE}[,，]?\s*出院日期[：:]{DATE}", text)
    if m:
        return norm(*m.group(1, 2, 3)), norm(*m.group(4, 5, 6))
    # 写法2：住院号:ZY100111  住院时间:2026.10.10 至 2026.10.17（或 2026年3月25日 至 ...）
    m = re.search(rf"住院时间[：:]{DATE}\s*至\s*{DATE}", text)
    if m:
        return norm(*m.group(1, 2, 3)), norm(*m.group(4, 5, 6))
    return None, None


def normalize_text(s):
    """去掉"2 型糖尿病"里数字和中文之间的怪空格 → "2型糖尿病"。"""
    return re.sub(r"(?<=[0-9])\s+(?=[\u4e00-\u9fa5A-Za-z])", "", s.strip())


def parse_diagnosis(text):
    """提取出院诊断列表。兼容编号列表、无编号、一句话式(考虑 A、B)。"""
    # 多数文件用"出院诊断:"，部分文件只有"诊断:"
    section = extract_section(text, "出院诊断") or extract_section(text, "诊断")
    items = []
    for line in section:
        line = re.sub(r"^\d+[.、]\s*", "", line)      # 去掉行首编号 "1. "
        line = re.sub(r"^(考虑|诊断为)", "", line)      # 去掉"考虑"等前缀词
        parts = [p.strip() for p in line.split("、") if p.strip()]  # 一句话式按、拆
        items.extend(normalize_text(p) for p in parts if p)
    # 去重且保持顺序
    return list(dict.fromkeys(items))


def parse_comorbidities(text):
    """从现病史中提取合并症。兼容多种写法：
      "既往有A、B病史 15 余年"（病史在句尾）
      "既往有A病史、B病史 15 余年"（病史在句中，需整句解析避免漏病）
      "既往有A病史病史 15 余年"（"病史"重复的脏数据）
    策略：取"既往有"到句号前的整句，统一去掉"病史"词与尾随年限，
    再按分隔符拆分病名；遇"否认"截断，只保留否认前的内容。
    未写"既往有"（如"平素体健,否认慢性病史"）→ 返回空列表。"""
    seg = re.search(r"既往有([^。\n]+)", text)
    if not seg:
        return []          # 未出现"既往有" → 视为无合并症
    raw = seg.group(1)
    raw = re.split(r"否认", raw)[0]     # "既往有X病史，否认Y史" → 只取否认之前的 X
    raw = re.sub(r"病史", "", raw)      # 统一去掉"病史"（句中/句尾/重复都处理）
    # 先按分隔符切分，再逐项剥除尾部年限（兼容"否认"后残留逗号的情况）
    items = [normalize_text(re.sub(r"\s*\d+(?:\.\d+)?\s*余?年\s*$", "", p))
             for p in re.split(r"[、，,]", raw) if p.strip()]
    return list(dict.fromkeys(i for i in items if i and len(i) <= 15))  # 过滤误抓的超长片段


# 用药行的"剂量"样子：0.25g / 30mg / 456mg / 200μg / 2IU ...
DOSE_PAT = re.compile(r"^\d+(?:\.\d+)?(?:mg|g|IU|ml|μg|ug|万单位)$", re.IGNORECASE)
# 特殊剂量单位（吸入剂等）：写成"1 吸""2 喷"这种 两个词 的形式
UNIT_WORDS = {"吸", "喷", "支", "片", "粒", "袋", "丸"}
# "频次"开头词：每日/每周/每晚/每 12 小时/隔日/必要时/按需吸入
FREQ_PAT = re.compile(r"^(每|隔日|必要时|按需)")


def parse_med_line(line):
    """把一行用药拆成 (药名, 剂量, 频次)。
    例：'头孢呋辛酯片 0.25g 每日两次'   → ('头孢呋辛酯片', '0.25g', '每日两次')
        '布地奈德福莫特罗吸入剂 1 吸 每日两次' → ('布地奈德福莫特罗吸入剂', '1吸', '每日两次')
        '甲氨蝶呤片 每周一次'           → ('甲氨蝶呤片', None,  '每周一次')
    注意：碳酸钙 D3 片里的"D3"带字母，不会误判成剂量。"""
    tokens = line.split()
    name_parts, dose, freq_parts = [], None, []
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if freq_parts or FREQ_PAT.match(tok):       # 碰到频次词，后面全是频次
            freq_parts.append(tok)
        elif (i + 1 < len(tokens) and re.fullmatch(r"\d+(?:\.\d+)?", tok)
              and name_parts
              and (tokens[i + 1] in UNIT_WORDS
                   or DOSE_PAT.fullmatch(tok + tokens[i + 1]))):
            dose = tok + tokens[i + 1]              # "1 吸"、"5 mg"、"0.5 g"
            i += 1
        elif re.match(r"^\d+(?:\.\d+)?[吸喷支片粒袋丸]$", tok) and name_parts:
            dose = tok                              # 无空格的"1吸"也认作剂量
        elif DOSE_PAT.match(tok) and name_parts:    # 前面已有药名 → 这是剂量
            dose = tok
        else:
            name_parts.append(tok)
        i += 1
    name = normalize_text("".join(name_parts))      # 碳酸钙 D3 片 → 碳酸钙D3片
    freq = normalize_text("".join(freq_parts)) or None
    return (name if name else None, dose, freq)


# 医嘱话术特征词：即使带频次(如"监测血压 每日一次")也不算药
MED_JUNK = ("监测", "复查", "复诊", "随访", "随诊", "戒烟", "限酒", "饮食", "活动",
            "休息", "运动", "饮水", "避免", "不适", "门诊", "定期")


def parse_meds(text):
    """提取用药清单。按 出院带药 → 带药 → 出院医嘱 顺序查找；
    出院医嘱可能混入医嘱话术(监测血压/戒烟限酒等)，仅保留符合药物特征的条目；
    同一行可能含多种药(以、连接)，先拆分再逐条解析。"""
    for title in ["出院带药", "带药", "出院医嘱"]:
        section = extract_section(text, title)
        if not section:
            continue
        meds = []
        for line in section:
            for part in line.split("、"):            # 一行多药：A 20mg 每日一次、B 10mg 每日两次
                part = re.sub(r"^\d+[.、]\s*", "", part)  # 去编号
                name, dose, freq = parse_med_line(part)
                # 药物判定：有药名且(有剂量或频次)，且不含医嘱话术特征词
                if (name and (dose or freq)
                        and not any(j in name for j in MED_JUNK)):
                    meds.append((name, dose, freq))
        if meds:
            return meds
    return []


# ========== 第二步：单文件处理 ==========
def process_file(path):
    with open(path, encoding="utf-8") as f:
        text = f.read()
    name, gender, age, pid = parse_header(text)
    admit, discharge = parse_dates(text)
    diagnosis = parse_diagnosis(text)
    comorbid = parse_comorbidities(text)
    meds = parse_meds(text)

    # 住院天数（日期解析失败时 stay_days 为 None，不中断整批）
    stay_days = None
    if admit and discharge:
        try:
            stay_days = (pd.Timestamp(discharge) - pd.Timestamp(admit)).days
        except ValueError:
            stay_days = None

    # 日期问题归类：文本里有日期字样但解析失败 → "日期无效"；根本没有 → "日期缺失"
    date_problem = None
    if not admit or not discharge:
        if re.search(r"入院日期|出院日期|住院时间", text):
            date_problem = "日期无效(格式无法解析)"
        else:
            date_problem = "日期缺失"

    return {
        "患者ID": pid,
        "姓名(脱敏)": (name[0] + "**") if name else None,   # 隐私保护：仅保留姓氏
        "性别": gender,
        "年龄": age,
        "入院日期": admit,
        "出院日期": discharge,
        "住院天数": stay_days,
        "主诉": (extract_section(text, "主诉") or [None])[0],
        "入院诊断": (extract_section(text, "入院诊断") or [None])[0],
        "出院诊断": "、".join(diagnosis),
        "合并症数量": len(comorbid),
        "合并症": "、".join(comorbid) if comorbid else None,
        "用药数量": len(meds),
        "用药": "、".join(m[0] for m in meds) if meds else None,
        "来源文件": path.name,          # 只要文件名，如 patient_001.txt
        "_日期问题": date_problem,      # "日期缺失"/"日期无效(...)"/None
        "_诊断列表": diagnosis,
        "_合并症列表": comorbid,
        "_用药明细": meds,   # [(药名,剂量,频次), ...]
    }


# ========== 第三步：批量处理、校验与输出 ==========
def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)   # 输出文件夹不存在就自动建
    files = sorted(DATA_DIR.glob("*.txt"))
    if not files:
        raise SystemExit(f"没找到 txt 文件，检查 DATA_DIR：{DATA_DIR}")

    records = [process_file(f) for f in files]
    df = pd.DataFrame(records)

    # ---- 清洗与校验 ----
    # 注意：进入 DataFrame 后缺失值以 NaN 表示，须用 pd.isna()/pd.notna() 判断；
    # NaN 在布尔判断中为真，不能直接用于 if 检查字段是否缺失。
    issues = {i: [] for i in range(len(df))}
    for i, row in df.iterrows():
        if pd.isna(row["患者ID"]):
            issues[i].append("缺住院号")
        if row["性别"] not in ("男", "女"):
            issues[i].append("性别缺失或异常")
        if pd.isna(row["年龄"]) or not (0 < row["年龄"] <= 120):
            issues[i].append("年龄缺失或异常")
        dq = row["_日期问题"]
        if pd.notna(dq):                          # 仅当存在日期问题时才追加标记
            issues[i].append(dq)
        elif pd.isna(row["入院日期"]) or pd.isna(row["出院日期"]):
            issues[i].append("日期缺失")           # 兜底分支
        elif not (0 < row["住院天数"] <= 365):
            issues[i].append("住院天数异常")
        if not row["_诊断列表"]:
            issues[i].append("无出院诊断")
        if row["用药数量"] == 0:
            issues[i].append("未提取到用药")
        if any(dose is None for _, dose, _ in row["_用药明细"]):
            issues[i].append("部分用药缺剂量")

    # 住院号重复检查：同一住院号的所有记录均标记
    dup_ids = df.loc[df["患者ID"].duplicated(keep=False) & df["患者ID"].notna(), "患者ID"]
    for pid in dup_ids.unique():
        for idx in df.index[df["患者ID"] == pid]:
            issues[idx].append(f"住院号重复({pid})")

    df["质量问题"] = ["；".join(v) if v else "无" for v in issues.values()]

    # ---- 拆长表 ----
    long_comorbid = df.explode("_合并症列表")[["患者ID", "_合并症列表"]] \
        .dropna().rename(columns={"_合并症列表": "合并症"})
    med_detail = df[["患者ID", "_用药明细"]].explode("_用药明细").dropna()
    long_med = pd.DataFrame({
        "患者ID": med_detail["患者ID"],
        "药名": [m[0] for m in med_detail["_用药明细"]],
        "剂量": [m[1] for m in med_detail["_用药明细"]],
        "频次": [m[2] for m in med_detail["_用药明细"]],
    })

    # ---- 输出 ----
    wide_cols = [c for c in df.columns if not c.startswith("_")]
    df_wide = df[wide_cols]

    df_wide.to_csv(OUTPUT_DIR / "病历结构化_宽表.csv", index=False, encoding="utf-8-sig")
    long_comorbid.to_csv(OUTPUT_DIR / "病历结构化_合并症长表.csv", index=False, encoding="utf-8-sig")
    long_med.to_csv(OUTPUT_DIR / "病历结构化_用药长表.csv", index=False, encoding="utf-8-sig")
    df[["患者ID", "来源文件", "质量问题"]].to_csv(
        OUTPUT_DIR / "数据质量报告.csv", index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(OUTPUT_DIR / "病历结构化结果.xlsx", engine="openpyxl") as writer:
        df_wide.to_excel(writer, sheet_name="宽表(一行一病人)", index=False)
        long_comorbid.to_excel(writer, sheet_name="合并症长表", index=False)
        long_med.to_excel(writer, sheet_name="用药长表", index=False)
        df[["患者ID", "来源文件", "质量问题"]].to_excel(writer, sheet_name="质量报告", index=False)

    # ---- 运行摘要 ----
    print(f"处理完成：{len(files)} 份病历")
    print(f"字段完整率：住院号 {df['患者ID'].notna().mean():.0%} | "
          f"日期 {(df['入院日期'].notna() & df['出院日期'].notna()).mean():.0%} | "
          f"诊断 {df['_诊断列表'].apply(bool).mean():.0%} | "
          f"用药 {df['用药数量'].gt(0).mean():.0%}")
    print(f"有质量问题：{(df['质量问题'] != '无').sum()} 人 "
          f"({(df['质量问题'] != '无').mean():.0%})")
    print(f"性别：男 {(df['性别'] == '男').sum()} / 女 {(df['性别'] == '女').sum()}；"
          f"年龄 {df['年龄'].min()}-{df['年龄'].max()} 岁，中位 {df['年龄'].median():.0f} 岁")
    print(f"住院天数：中位 {df['住院天数'].median():.0f} 天 "
          f"(范围 {df['住院天数'].min():.0f}-{df['住院天数'].max():.0f})")
    print("\n合并症 Top5：")
    print(long_comorbid["合并症"].value_counts().head(5).to_string())
    print("\n用药 Top5：")
    print(long_med["药名"].value_counts().head(5).to_string())


if __name__ == "__main__":
    main()
