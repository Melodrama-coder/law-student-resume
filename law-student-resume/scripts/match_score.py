#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
match_score.py - 简历与目标岗位匹配度雷达图

根据简历内容，对照内置的红圈所/央企法务 JD 画像，
在 6 个维度上打分并生成雷达图 PNG + 文本报告。

维度：
  1. 学历背景（目标院校层次、GPA、奖学金）
  2. 实习经历（律所/法院/企业法务数量、时长、相关性）
  3. 职业资格（法考、证券/基金/CPA 等）
  4. 语言能力（英语证书、其他语言）
  5. 检索与工具（北大法宝、Westlaw、办公软件等）
  6. 竞赛/科研（模拟法庭、科研项目、论文）

用法：
  python scripts/match_score.py \
    --data resume.json \
    --target law_firm        # 或 corporate_legal / foreign_firm \
    --out report.png
"""

import argparse
import json
import math
import os


def _check_dependencies():
    """matplotlib 未安装时给出友好提示，而非堆栈。"""
    try:
        import matplotlib
        return matplotlib
    except ImportError as e:
        print("[error] 缺少依赖：matplotlib。")
        print("        请运行：pip install matplotlib")
        raise SystemExit(1) from e


TARGET_PROFILES = {
    "law_firm": {
        "label": "红圈所 / 精品所",
        "weights": {
            "education": 0.15,
            "internship": 0.30,
            "qualification": 0.20,
            "language": 0.15,
            "tools": 0.10,
            "competition_research": 0.10,
        },
        "keywords": {
            "internship": ["律所", "红圈", "金杜", "中伦", "竞天", "世辉", "方达", "海问", "环球", "通商", "锦天城", "汉坤", "IPO", "尽调", "诉讼", "仲裁", "争议解决"],
            "qualification": ["法律职业资格", "法考", "A证", "证券从业", "基金从业", "CPA"],
            "language": ["CET-6", "雅思", "托福", "法律英语", "英文合同"],
            "tools": ["北大法宝", "威科", "Westlaw", "Lexis", "法律检索", "案例检索"],
            "competition_research": ["Jessup", "moot", "模拟法庭", " Vis ", "科研", "论文", "课题"],
        },
    },
    "corporate_legal": {
        "label": "央企 / 大型企业法务",
        "weights": {
            "education": 0.15,
            "internship": 0.25,
            "qualification": 0.15,
            "language": 0.10,
            "tools": 0.15,
            "competition_research": 0.20,
        },
        "keywords": {
            "internship": ["法务", "合规", "合同审查", "数据合规", "公司法务", "风险管理"],
            "qualification": ["法律职业资格", "法考", "A证", "企业合规师", "证券从业"],
            "language": ["CET-6", "雅思", "托福"],
            "tools": ["合同管理", "Office", "Excel", "北大法宝", "合规清单"],
            "competition_research": ["科研", "课题", "论文", "模拟法庭"],
        },
    },
    "foreign_firm": {
        "label": "外所 / 国际仲裁机构",
        "weights": {
            "education": 0.20,
            "internship": 0.25,
            "qualification": 0.15,
            "language": 0.25,
            "tools": 0.10,
            "competition_research": 0.05,
        },
        "keywords": {
            "internship": ["律所", "law firm", "international", "arbitration", "dispute resolution", "M&A", "due diligence"],
            "qualification": ["法考", "法律职业资格", "NY Bar", "UK Bar", "LLM"],
            "language": ["雅思", "托福", "CET-6", "法律英语", "英文", "legal English"],
            "tools": ["Westlaw", "Lexis", "Bloomberg", "英文检索"],
            "competition_research": ["Jessup", "Willem C. Vis", "Vis East", "国际法", " arbitration "],
        },
    },
}


def _text_of(data, lang="zh"):
    """提取 resume JSON 中所有可用于关键词匹配的文本。"""
    d, _ = data, None
    if lang == "en":
        d = data.get("en") or data
    texts = []
    p = d.get("personal", {})
    texts += [p.get("job_intention", "")]
    for e in d.get("education", []):
        texts += [str(e.get(k, "")) for k in ("school", "college", "major", "degree", "gpa", "rank", "honors", "courses")]
    for item in d.get("internship", []):
        texts += [str(item.get(k, "")) for k in ("org", "dept_role", "location")]
        texts += [str(b) for b in item.get("bullets", [])]
    for item in d.get("work", []):
        texts += [str(item.get(k, "")) for k in ("org", "dept_role", "location")]
        texts += [str(b) for b in item.get("bullets", [])]
    for item in d.get("research", []):
        texts += [str(item.get(k, "")) for k in ("title", "role")]
        texts += [str(b) for b in item.get("bullets", [])]
    for item in d.get("competition", []):
        texts += [str(item.get(k, "")) for k in ("title", "role")]
        texts += [str(b) for b in item.get("bullets", [])]
    s = d.get("skills", {})
    texts += [str(s.get(k, "")) for k in ("qualification", "language", "software", "extra")]
    texts.append(str(d.get("summary", "")))
    return "\n".join(t for t in texts if t)


def _score_dimension(text, keywords):
    """根据关键词命中数量给 0-100 分。"""
    if not keywords:
        return 0
    hits = sum(1 for kw in keywords if kw in text)
    return min(100, int(hits / max(1, len(keywords) * 0.3) * 100))


def score_resume(data, target="law_firm"):
    """返回 (各维度分数字典, 总分, 文本报告)。"""
    profile = TARGET_PROFILES.get(target, TARGET_PROFILES["law_firm"])
    text = _text_of(data)
    text_en = _text_of(data, lang="en")

    dims = {}
    # 1. 教育背景
    edu = data.get("education", [])
    edu_score = 0
    if edu:
        edu_score = 60
        for e in edu:
            gpa = str(e.get("gpa", ""))
            if gpa and ("3.7" in gpa or "3.8" in gpa or "3.9" in gpa or "4.0" in gpa):
                edu_score += 10
            rank = str(e.get("rank", ""))
            if rank and ("前" in rank or "Top" in rank or "/" in rank):
                edu_score += 10
            honors = str(e.get("honors", ""))
            if honors:
                edu_score += 5
            school = str(e.get("school", ""))
            if any(k in school for k in ["法大", "政法", "清华", "北大", "人大", "复旦", "交大", "华政", "China University", "LL.M.", "LL.B."]):
                edu_score += 5
    edu_score = min(100, edu_score)

    # 2. 实习经历
    it = data.get("internship", [])
    wk = data.get("work", [])
    exp_score = 0
    if it or wk:
        exp_score = 50
        exp_score += min(30, len(it) * 10)
        exp_score += min(20, len(wk) * 10)
    exp_score = min(100, exp_score)

    # 3. 职业资格
    qual_score = _score_dimension(text, profile["keywords"]["qualification"])

    # 4. 语言能力
    lang_score = _score_dimension(text + text_en, profile["keywords"]["language"])

    # 5. 工具/检索
    tools_score = _score_dimension(text, profile["keywords"]["tools"])

    # 6. 竞赛/科研
    cr_score = _score_dimension(text, profile["keywords"]["competition_research"])

    dims = {
        "education": edu_score,
        "internship": exp_score,
        "qualification": qual_score,
        "language": lang_score,
        "tools": tools_score,
        "competition_research": cr_score,
    }

    total = sum(int(dims[k] * profile["weights"][k]) for k in dims)
    return dims, total, profile["label"]


def render_radar(dims, total, target_label, out_path):
    """生成雷达图 PNG。"""
    matplotlib = _check_dependencies()
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    # 在 Windows 上优先使用系统中文字体，避免雷达图中文显示为方框
    import matplotlib.font_manager as fm
    candidates = ["Microsoft YaHei", "SimHei", "SimSun", "Noto Sans CJK SC", "WenQuanYi Micro Hei"]
    chinese_font = None
    for name in candidates:
        try:
            if fm.findfont(fm.FontProperties(family=name), fallback_to_default=False) != fm.findfont(fm.FontProperties(family="DejaVu Sans"), fallback_to_default=False):
                chinese_font = name
                break
        except Exception:
            continue
    if chinese_font:
        plt.rcParams["font.sans-serif"] = [chinese_font] + plt.rcParams.get("font.sans-serif", [])
        plt.rcParams["axes.unicode_minus"] = False

    labels = ["学历\n背景", "实习\n经历", "职业\n资格", "语言\n能力", "检索/工具", "竞赛/科研"]
    keys = ["education", "internship", "qualification", "language", "tools", "competition_research"]
    values = [dims[k] for k in keys]
    values += values[:1]  # 闭合

    angles = [i * 2 * math.pi / len(labels) for i in range(len(labels))]
    angles += angles[:1]

    fig, ax = plt.subplots(figsize=(6, 5.5), subplot_kw=dict(polar=True))
    ax.set_theta_offset(math.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_thetagrids([a * 180 / math.pi for a in angles[:-1]], labels, fontsize=11)
    ax.set_ylim(0, 100)
    ax.set_rlabel_position(0)
    ax.plot(angles, values, color="#2E7CC2", linewidth=2, marker="o")
    ax.fill(angles, values, color="#2E7CC2", alpha=0.25)
    ax.set_title(f"简历-岗位匹配度：{target_label}\n综合得分 {total}/100", y=1.08, fontsize=13, weight="bold")
    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"✅ 雷达图已保存：{out_path}")


def text_report(dims, total, target_label):
    """生成文本诊断。"""
    lines = [f"\n=== 简历与「{target_label}」匹配度诊断 ==="]
    lines.append(f"综合得分：{total}/100")
    name_map = {
        "education": "学历背景",
        "internship": "实习经历",
        "qualification": "职业资格",
        "language": "语言能力",
        "tools": "检索与工具",
        "competition_research": "竞赛/科研",
    }
    for k, label in name_map.items():
        score = dims[k]
        bar = "█" * (score // 10) + "░" * (10 - score // 10)
        lines.append(f"  {label:10s} {score:3d} {bar}")
    # 找短板
    weakest = sorted(dims.items(), key=lambda x: x[1])[:2]
    lines.append("\n提升建议：")
    for k, _ in weakest:
        if k == "internship":
            lines.append("  - 实习经历：补充 1 段与目标岗位强相关的 3 个月以上实习，保留量化成果。")
        elif k == "language":
            lines.append("  - 语言能力：突出雅思/托福/CET-6 分数，增加英文合同审阅或英文检索经验。")
        elif k == "qualification":
            lines.append("  - 职业资格：法考 A 证是基本盘；如目标含证券/基金/合规，可补充对应证书。")
        elif k == "tools":
            lines.append("  - 检索与工具：明确列出北大法宝、威科、Westlaw、Lexis、企查查等使用深度。")
        elif k == "competition_research":
            lines.append("  - 竞赛/科研：增加 Jessup/Vis/模拟法庭或科研课题、发表论文等经历。")
        elif k == "education":
            lines.append("  - 学历背景：突出 GPA、排名、奖学金、核心课程，弱化无关课程。")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="简历-岗位匹配度雷达图")
    ap.add_argument("--data", required=True, help="简历 JSON 路径")
    ap.add_argument("--target", default="law_firm",
                    choices=list(TARGET_PROFILES.keys()),
                    help="目标岗位画像：law_firm/corporate_legal/foreign_firm")
    ap.add_argument("--out", default="match_score.png", help="输出 PNG 路径")
    args = ap.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    dims, total, label = score_resume(data, args.target)
    print(text_report(dims, total, label))
    render_radar(dims, total, label, args.out)


if __name__ == "__main__":
    main()
