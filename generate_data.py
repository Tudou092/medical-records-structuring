# -*- coding: utf-8 -*-
"""
模拟出院小结数据生成器(修正版)
==============================
用途:为"病历文本结构化"项目生成 N 份【模拟】出院小结文本。
    所有患者姓名、住院号均为随机虚构,不含任何真实病人信息。

v2 修正:
1. 出院日期 = 入院日期 + 3~14 天(不再出现出院早于入院);
2. 治疗经过与主诊断绑定(冠心病不会再写"抗感染");
3. 同一药物在带药区去重(一药治多病只列一次);
4. "考虑/待查"等脏点按临床书写习惯插入;
5. 日期混用格式(mixed)体现在同一份文件的入院/出院行不一致。

可复现:固定 random.seed(SEED);可扩量:改 N 即可。

生成结果:
- data/patient_XXX.txt      每份病历一个文本文件(共 N 份)
- gold_standard.csv         标准答案(诊断/合并症/药物),供程序跑完后自测准确率

运行方式(在项目目录下):
    python generate_data.py --output-dir ../new-batch
    目标批次包含 data/ 和 gold_standard.csv；已有数据时拒绝覆盖。
"""

import argparse
import csv
import random
from datetime import datetime, timedelta
from pathlib import Path

# ============ 可调参数 ============
N = 50                # 生成的份数
SEED = 42             # 随机种子(固定后可复现)
OUT_DIR = "data"      # 批次内的数据子目录
# ==================================

random.seed(SEED)

# ---------- 虚构姓名用字(避免真实人名) ----------
SURNAMES = ["张", "李", "王", "刘", "陈", "杨", "赵", "黄", "周", "吴"]
GIVEN = ["建国", "志强", "秀英", "桂芳", "丽华", "国庆", "玉兰", "海涛",
         "晓东", "春梅", "文博", "雨桐", "子涵", "欣怡", "浩然", "梓萱"]

# ---------- 主诊断库:诊断名 + 症状 + 治疗 + 用药 + 检查 ----------
MAIN_DX = [
    {"name": "高血压病 2 级(很高危)",
     "symptom": "反复头晕、头痛",
     "tx": "予降压、改善循环等对症治疗",
     "drugs": [("苯磺酸氨氯地平片", "5mg", "每日一次"),
               ("缬沙坦胶囊", "80mg", "每日一次")],
     "tests": [("血压", "172/104 mmHg")]},
    {"name": "2 型糖尿病",
     "symptom": "多饮、多尿、体重下降",
     "tx": "予控制血糖、改善代谢等综合治疗",
     "drugs": [("盐酸二甲双胍片", "0.5g", "每日两次"),
               ("阿卡波糖片", "50mg", "每日三次")],
     "tests": [("空腹血糖", "9.6 mmol/L"), ("糖化血红蛋白", "8.8%")]},
    {"name": "冠心病 不稳定型心绞痛",
     "symptom": "活动后胸闷、胸痛",
     "tx": "予抗血小板、调脂、扩冠等冠心病二级预防治疗",
     "drugs": [("阿司匹林肠溶片", "100mg", "每日一次"),
               ("阿托伐他汀钙片", "20mg", "每晚一次"),
               ("琥珀酸美托洛尔缓释片", "47.5mg", "每日一次")],
     "tests": [("心电图", "Ⅱ、Ⅲ、aVF 导联 ST 段压低 0.1mV")]},
    {"name": "社区获得性肺炎",
     "symptom": "发热、咳嗽、咳痰",
     "tx": "予抗感染、化痰等治疗",
     "drugs": [("头孢呋辛酯片", "0.25g", "每日两次"),
               ("盐酸氨溴索片", "30mg", "每日三次")],
     "tests": [("胸部 CT", "右肺下叶斑片状高密度影"),
               ("血常规 WBC", "13.2×10^9/L")]},
    {"name": "慢性阻塞性肺疾病(急性加重期)",
     "symptom": "咳嗽、喘息加重",
     "tx": "予解痉平喘、祛痰等治疗",
     "drugs": [("布地奈德福莫特罗吸入剂", "1 吸", "每日两次"),
               ("茶碱缓释片", "0.1g", "每日两次")],
     "tests": [("肺功能 FEV1/FVC", "58%")]},
    {"name": "急性脑梗死",
     "symptom": "突发口角歪斜、右侧肢体无力",
     "tx": "予抗血小板、调脂、改善脑循环等治疗",
     "drugs": [("阿司匹林肠溶片", "100mg", "每日一次"),
               ("硫酸氢氯吡格雷片", "75mg", "每日一次"),
               ("阿托伐他汀钙片", "20mg", "每晚一次")],
     "tests": [("头颅 MRI", "左侧基底节区急性脑梗死灶")]},
    {"name": "消化性溃疡(胃溃疡)",
     "symptom": "上腹部隐痛、反酸、嗳气",
     "tx": "予抑酸、保护胃黏膜等治疗",
     "drugs": [("奥美拉唑肠溶胶囊", "20mg", "每日一次"),
               ("胶体果胶铋胶囊", "200mg", "每日两次")],
     "tests": [("胃镜", "胃窦部可见 0.6cm 溃疡,HP(+)")]},
    {"name": "急性胆囊炎",
     "symptom": "右上腹持续性疼痛伴发热",
     "tx": "予抗感染、解痉止痛等治疗",
     "drugs": [("注射用头孢哌酮钠舒巴坦钠", "2g", "每 12 小时一次"),
               ("盐酸山莨菪碱注射液", "10mg", "必要时")],
     "tests": [("腹部 B 超", "胆囊增大,壁增厚毛糙,内见泥沙样结石")]},
    {"name": "腰椎间盘突出症",
     "symptom": "腰痛伴右下肢放射痛",
     "tx": "予脱水消肿、营养神经、止痛等保守治疗",
     "drugs": [("塞来昔布胶囊", "200mg", "每日一次"),
               ("甲钴胺片", "0.5mg", "每日三次")],
     "tests": [("腰椎 MRI", "L4/5 椎间盘向后突出,压迫右侧神经根")]},
    {"name": "类风湿关节炎",
     "symptom": "双手掌指关节肿痛、晨僵",
     "tx": "予抗风湿、抗炎止痛等治疗",
     "drugs": [("甲氨蝶呤片", "10mg", "每周一次"),
               ("来氟米特片", "20mg", "每日一次")],
     "tests": [("类风湿因子", "156 IU/mL"), ("抗 CCP 抗体", "阳性")]},
    {"name": "甲状腺功能亢进症",
     "symptom": "心悸、多汗、手抖、体重下降",
     "tx": "予抗甲状腺药物控制代谢等治疗",
     "drugs": [("甲巯咪唑片", "10mg", "每日三次"),
               ("盐酸普萘洛尔片", "10mg", "每日三次")],
     "tests": [("血清 TSH", "<0.01 mIU/L"), ("血清 FT4", "36.5 pmol/L")]},
    {"name": "右肾结石",
     "symptom": "突发右侧腰部绞痛、肉眼血尿",
     "tx": "予解痉止痛、促进排石等治疗",
     "drugs": [("盐酸坦索罗辛缓释胶囊", "0.4mg", "每日一次"),
               ("双氯芬酸钠栓", "50mg", "必要时")],
     "tests": [("泌尿系 CT", "右肾盏见直径约 5mm 高密度结石")]},
    {"name": "缺铁性贫血",
     "symptom": "乏力、头晕、活动后心悸",
     "tx": "予补铁、促进造血等纠正贫血治疗",
     "drugs": [("琥珀酸亚铁片", "0.1g", "每日三次"),
               ("维生素 C 片", "0.1g", "每日三次")],
     "tests": [("血常规 Hb", "78 g/L"), ("血清铁蛋白", "6.2 μg/L")]},
    {"name": "支气管哮喘(急性发作期)",
     "symptom": "反复发作性喘息、胸闷",
     "tx": "予解痉平喘、抗炎等治疗",
     "drugs": [("沙丁胺醇气雾剂", "200μg", "按需吸入"),
               ("布地奈德福莫特罗吸入剂", "1 吸", "每日两次")],
     "tests": [("呼气流速峰值(PEF)", "预计值 55%")]},
    {"name": "泌尿道感染",
     "symptom": "尿频、尿急、尿痛",
     "tx": "予抗感染、碱化尿液等治疗",
     "drugs": [("盐酸左氧氟沙星片", "0.5g", "每日一次"),
               ("热淋清颗粒", "4g", "每日三次")],
     "tests": [("尿常规", "白细胞满视野,亚硝酸盐(+)")]},
]

# ---------- 合并症库(含对应基础用药) ----------
COMORB = [
    ("高血压病", ("苯磺酸氨氯地平片", "5mg", "每日一次")),
    ("2 型糖尿病", ("盐酸二甲双胍片", "0.5g", "每日两次")),
    ("高脂血症", ("阿托伐他汀钙片", "20mg", "每晚一次")),
    ("冠状动脉粥样硬化性心脏病", ("阿司匹林肠溶片", "100mg", "每日一次")),
    ("心房颤动", ("利伐沙班片", "20mg", "每日一次")),
    ("慢性阻塞性肺疾病", ("噻托溴铵吸入粉雾剂", "1 吸", "每日一次")),
    ("慢性肾功能不全", ("百令胶囊", "2g", "每日三次")),
    ("骨质疏松症", ("碳酸钙 D3 片", "600mg", "每日一次")),
    ("非酒精性脂肪肝", ("多烯磷脂酰胆碱胶囊", "456mg", "每日三次")),
    ("高尿酸血症", ("非布司他片", "40mg", "每日一次")),
    ("脑梗死", ("阿司匹林肠溶片", "100mg", "每日一次")),
    ("轻度贫血", ("琥珀酸亚铁片", "0.1g", "每日两次")),
]

FOLLOW_UP = [
    "定期门诊复查",
    "1 个月后门诊复诊",
    "不适随诊,定期监测血压",
    "监测血糖,内分泌科门诊随访",
    "戒烟限酒,避免受凉感冒",
    "低盐低脂饮食,适度活动",
]

DURATIONS = ["3 天", "1 周", "2 周", "1 月", "半年", "2 年"]


def unique_drugs(drug_list):
    """按药名去重,保留首次出现的顺序(剂量/频次取第一次的)"""
    seen = {}
    for d in drug_list:
        if d[0] not in seen:
            seen[d[0]] = d
    return list(seen.values())


def fmt_date(dt, style, with_leading_zero=True):
    """日期样式:norm=2026-02-19, dot=2026.02.19, cjk=2026年2月19日"""
    if style == "dot":
        return dt.strftime("%Y.%m.%d")
    if style == "cjk":
        return f"{dt.year}年{dt.month}月{dt.day}日"
    return dt.strftime("%Y-%m-%d")


def render_drug_lines(drugs):
    """带药区渲染,埋 3 类脏点:无编号 / 缺剂量 / 挤成一行"""
    style = random.choices(["norm", "no_num", "no_dose", "one_line"],
                           weights=[68, 10, 12, 10])[0]
    if style == "one_line":
        items = []
        for name, dose, freq in drugs:
            item = name
            if dose:
                item += f" {dose}"
            item += f" {freq}"
            items.append(item)
        return "、".join(items)

    lines = []
    for i, (name, dose, freq) in enumerate(drugs, start=1):
        head = "" if style == "no_num" else f"{i}. "
        line = head + name
        if dose and not (style == "no_dose" and random.random() < 0.6):
            line += f" {dose}"
        line += f" {freq}"
        lines.append(line)
    return "\n".join(lines)


def make_patient(pid):
    """生成一份患者的全部结构化信息(含脏点设计)"""
    name = random.choice(SURNAMES) + random.choice(GIVEN)
    sex = random.choice(["男", "女"])
    age = random.randint(24, 84)
    hid = f"ZY{100000 + pid * 37}"          # 确定性唯一住院号

    dx = random.choice(MAIN_DX)

    # 合并症:多数 1~2 个,少数 0 个或 3 个(weights 保证分布)
    k = random.choices([0, 1, 2, 3], weights=[6, 40, 36, 18])[0]
    picked = random.sample(COMORB, k)
    comorb_names = [c[0] for c in picked]
    comorb_drugs = [c[1] for c in picked]

    # 带药 = 主诊断药 + 合并症基础药,药名去重
    all_drugs = unique_drugs(list(dx["drugs"]) + comorb_drugs)

    # ---------- 日期:出院 = 入院 + 3~14 天 ----------
    admit = datetime(2026, 1, 1) + timedelta(days=random.randint(0, 300))
    dis = admit + timedelta(days=random.randint(3, 14))
    date_style = random.choices(["norm", "dot", "cjk", "mixed"],
                                weights=[76, 10, 8, 6])[0]
    if date_style == "mixed":
        admit_str = fmt_date(admit, "norm")
        dis_str = fmt_date(dis, "dot")       # 同一份里两种格式(脏点)
    else:
        admit_str = fmt_date(admit, date_style)
        dis_str = fmt_date(dis, date_style)

    # ---------- 出院诊断(含脏点:10% 带"考虑") ----------
    pending = random.random() < 0.10
    if random.random() > 0.15:
        # 编号格式
        items = []
        main_line = ("考虑 " if pending else "") + dx["name"]
        items.append(f"1. {main_line}")
        for j, c in enumerate(comorb_names, start=2):
            items.append(f"{j}. {c}")
        dx_lines = "\n".join(items)
    else:
        # 顿号连排(脏点:无编号)
        names = [(("考虑 " if pending else "") + dx["name"])] + comorb_names
        dx_lines = "、".join(names)

    # ---------- 现病史 ----------
    dur = random.choice(DURATIONS)
    test_str = "；".join(f"{k}:{v}" for k, v in dx["tests"])
    if comorb_names:
        years = random.randint(2, 15)
        his = "既往有" + "、".join(comorb_names) + f"病史 {years} 余年。"
    else:
        his = "平素体健,否认高血压、糖尿病等慢性病史。"

    # ---------- 三种版式(制造段落结构差异) ----------
    drug_text = render_drug_lines(all_drugs)
    follow = random.choice(FOLLOW_UP)
    layout = random.randint(1, 3)

    if layout == 1:
        body = f"""出院小结

姓名:{name}  性别:{sex}  年龄:{age}岁  住院号:{hid}
入院日期:{admit_str}  出院日期:{dis_str}

主诉:{dx['symptom']}{dur}。

现病史:患者于{dur}前无明显诱因出现{dx['symptom']}。门诊查{test_str},为进一步诊治收治入院。{his}

入院诊断:{dx['name']}
出院诊断:
{dx_lines}

诊疗经过:入院后完善相关检查,{dx['tx']},患者症状明显好转,予出院。

出院带药:
{drug_text}

出院医嘱:{follow}。"""
    elif layout == 2:
        body = f"""出院小结

患者 {name},{sex},{age}岁,住院号:{hid}
入院日期:{admit_str},出院日期:{dis_str}

主诉:因"{dx['symptom']}{dur}"入院。

现病史:患者于{dur}前出现{dx['symptom']}。{his} 否认药物过敏史。

入院诊断:{dx['name']}
出院诊断:
{dx_lines}

住院期间:完善检查示{test_str}。{dx['tx']},症状缓解出院。

出院医嘱:
{drug_text}

出院指导:{follow}。"""
    else:
        body = f"""出院小结

姓名:{name}  性别:{sex}  年龄:{age}岁
住院号:{hid}  住院时间:{admit_str} 至 {dis_str}

主诉:{dx['symptom']}{dur}。

现病史:{dur}前出现{dx['symptom']}。查体及辅助检查:{test_str}。{his}

诊断:
{dx_lines}

治疗经过:{dx['tx']},好转出院。

带药:
{drug_text}

注意:{follow}。"""

    return {
        "file": f"patient_{pid:03d}.txt",
        "name": name, "sex": sex, "age": age, "hid": hid,
        "main_dx": dx["name"],
        "comorbs": comorb_names,
        "drugs": [d[0] for d in all_drugs],
        "body": body,
    }


def main(output_root=None):
    root = Path(output_root) if output_root is not None else Path(__file__).resolve().parent
    data_dir = root / OUT_DIR
    gold_path = root / "gold_standard.csv"
    if not isinstance(N, int) or isinstance(N, bool) or N <= 0:
        raise SystemExit("N 必须是正整数")
    # 写入任何内容前检查，避免缩量残留旧病例或覆盖旧版测试样本。
    if gold_path.exists() or (data_dir.exists() and
                              (not data_dir.is_dir() or any(data_dir.iterdir()))):
        raise SystemExit("目标批次已有数据或标准答案，请用 --output-dir 指定新的批次目录；未覆盖任何文件")
    data_dir.mkdir(parents=True, exist_ok=True)
    random.seed(SEED)
    patients = []
    for i in range(1, N + 1):
        p = make_patient(i)
        patients.append(p)
        with open(data_dir / p["file"], "w", encoding="utf-8") as f:
            f.write(p["body"])

    # 标准答案表(供跑完程序后自测准确率)
    with open(gold_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["file", "主要诊断", "合并症", "药物"])
        for p in patients:
            w.writerow([
                p["file"],
                p["main_dx"],
                "；".join(p["comorbs"]) if p["comorbs"] else "",
                "；".join(p["drugs"]),
            ])

    n_empty = sum(1 for p in patients if not p["comorbs"])
    n_pending = sum(1 for p in patients if "考虑" in p["body"])
    n_drugs = [len(p["drugs"]) for p in patients]
    print(f"生成完成:共 {N} 份,输出至 {data_dir}/")
    print(f"无合并症病历: {n_empty} 份 | 带'考虑'病历: {n_pending} 份")
    print(f"每份带药种类: {min(n_drugs)} ~ {max(n_drugs)} 种")
    print(f"标准答案已写入: {gold_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="在新批次目录生成模拟病例和标准答案，不覆盖已有数据")
    parser.add_argument("--output-dir", type=Path, help="批次目录，包含 data/ 与 gold_standard.csv；默认脚本目录")
    main(parser.parse_args().output_dir)
