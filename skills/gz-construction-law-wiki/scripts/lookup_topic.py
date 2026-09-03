"""
按业务主题查 - 根据业务关键词映射到相关条款。
用法:
  python lookup_topic.py --topic "工资"
  python lookup_topic.py --topic "消防"
  python lookup_topic.py --topic "招标"
  python lookup_topic.py --topic "安全"
"""
import argparse
import re
from pathlib import Path

WIKI_ROOT = Path.home() / "wiki"

# 业务主题 → 关键词映射（多关键词取并集）
TOPIC_KEYWORDS = {
    "工资":      ["工资", "人工费", "欠薪", "专户", "保证金", "实名制", "劳资", "用工"],
    "消防":      ["消防", "防火", "灭火", "应急疏散"],
    "招标":      ["招标", "投标", "评标", "中标", "开标", "保证金", "规模", "必须招标"],
    "规划":      ["规划", "选址", "用地", "容积率", "建筑密度", "建设用地", "建设工程规划"],
    "安全":      ["安全", "事故", "隐患", "危大", "三同时", "防护"],
    "资质":      ["资质", "资格", "注册", "执业", "特种作业", "证书", "持证"],
    "监理":      ["监理", "总监", "见证", "旁站"],
    "环保":      ["环保", "环境", "排污", "噪声", "扬尘", "环评"],
    "农民工":    ["农民工", "务工", "劳务", "工人", "工资册", "考勤"],
    "建设单位":  ["建设单位", "业主", "甲方", "发包人"],
    "施工":      ["施工", "总包", "分包", "承包", "现场", "在建"],
    "竣工":      ["竣工", "验收", "备案", "核实"],
}

def main():
    parser = argparse.ArgumentParser(description="按业务主题查 wiki 法规")
    parser.add_argument("--topic", required=True, help="业务主题（部分匹配）")
    parser.add_argument("--show_lines", action="store_true", help="显示含关键词的行（限制 5 行/文件）")
    args = parser.parse_args()

    # 匹配主题
    matched_topics = {}
    for topic, kws in TOPIC_KEYWORDS.items():
        if args.topic in topic or topic in args.topic:
            matched_topics[topic] = kws

    if not matched_topics:
        # 模糊匹配
        for topic, kws in TOPIC_KEYWORDS.items():
            if any(args.topic in kw for kw in kws):
                matched_topics[topic] = kws

    if not matched_topics:
        print(f"❌ 未找到匹配主题: {args.topic}")
        print("支持的主题:")
        for t in TOPIC_KEYWORDS:
            print(f"  - {t}")
        return 1

    # 合并关键词
    all_kws = set()
    for kws in matched_topics.values():
        all_kws.update(kws)

    print(f"🎯 匹配主题: {', '.join(matched_topics.keys())}")
    print(f"🔑 关键词 ({len(all_kws)}): {', '.join(sorted(all_kws))}")
    print("─" * 60)

    # 搜索 wiki
    md_files = sorted(WIKI_ROOT.rglob("*.md"))
    hits_by_file = {}

    for f in md_files:
        try:
            text = f.read_text(encoding="utf-8")
        except Exception:
            continue
        lines = text.splitlines()
        hit_lines = []
        for i, line in enumerate(lines):
            if any(kw in line for kw in all_kws):
                hit_lines.append((i + 1, line.strip()))
        if hit_lines:
            hits_by_file[f] = hit_lines

    if not hits_by_file:
        print(f"⚠️ 在 wiki 中未找到任何相关条款")
        return 0

    total = sum(len(v) for v in hits_by_file.values())
    print(f"📊 命中 {total} 行 / {len(hits_by_file)} 个文件\n")

    for f, hits in hits_by_file.items():
        rel = f.relative_to(WIKI_ROOT)
        print(f"📄 {rel} ({len(hits)} 处)")
        if args.show_lines:
            for lineno, line in hits[:5]:
                print(f"  L{lineno}: {line[:150]}")
            if len(hits) > 5:
                print(f"  ... 还有 {len(hits)-5} 处")
        print()

    return 0

if __name__ == "__main__":
    raise SystemExit(main())
