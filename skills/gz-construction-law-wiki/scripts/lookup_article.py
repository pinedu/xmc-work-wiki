"""
按条款号查 - 在 ~/wiki/ 下找指定法规的指定条款。
用法:
  python lookup_article.py --law "保障农民工工资支付条例" --article "第二十四条"
  python lookup_article.py --law "消防法" --article "第十条"   # 简写匹配
"""
import argparse
import re
from pathlib import Path

WIKI_ROOT = Path.home() / "wiki" / "concepts"

# 法规名到目录的映射（支持简写）
LAW_DIRS = {
    "保障农民工工资支付条例":       WIKI_ROOT / "上位法",
    "农民工工资支付条例":           WIKI_ROOT / "上位法",
    "中华人民共和国消防法":         WIKI_ROOT / "上位法",
    "消防法":                       WIKI_ROOT / "上位法",
    "必须招标的工程项目规定":       WIKI_ROOT / "上位法",
    "招标规定":                     WIKI_ROOT / "上位法",
    "贵州省工程建设领域农民工实名制管理暂行办法": WIKI_ROOT / "贵州省地方法规",
    "农民工实名制":                 WIKI_ROOT / "贵州省地方法规",
    "实名制":                       WIKI_ROOT / "贵州省地方法规",
    "贵州省建筑市场管理条例":       WIKI_ROOT / "贵州省地方法规",
    "建筑市场管理条例":             WIKI_ROOT / "贵州省地方法规",
    "贵州省贯彻落实《保障农民工工资支付条例》实施意见": WIKI_ROOT / "贵州省地方法规",
    "黔府办发32号":                 WIKI_ROOT / "贵州省地方法规",
    "实施意见":                     WIKI_ROOT / "贵州省地方法规",
    "贵州省招标投标条例":           WIKI_ROOT / "贵州省地方法规",
    "招标投标条例":                 WIKI_ROOT / "贵州省地方法规",
    "贵州省城乡规划条例":           WIKI_ROOT / "贵州省地方法规",
    "城乡规划条例":                 WIKI_ROOT / "贵州省地方法规",
    "贵州省安全生产条例":           WIKI_ROOT / "贵州省地方法规",
    "安全生产条例":                 WIKI_ROOT / "贵州省地方法规",
}

def find_note(law_name: str) -> Path | None:
    """根据法规名定位笔记文件"""
    for key, d in LAW_DIRS.items():
        if key in law_name or law_name in key:
            for f in d.glob("*.md"):
                if f.name != "_index.md" and key.replace(WIKI_ROOT.parent.name, "") in f.name or law_name in f.name:
                    return f
            # 如果是子目录
            for f in d.iterdir():
                if f.is_file() and f.suffix == ".md" and f.name != "_index.md":
                    return f
    return None

def extract_article(text: str, article: str) -> list[str]:
    """从笔记中提取指定条款及其原文（含块引用）"""
    lines = text.splitlines()
    captured = []
    found = False
    # 宽松匹配：条款号后面可能跟（中文括号）或顿号、空格
    article_re = re.compile(r"^#{2,4}\s+" + re.escape(article) + r"[\s（(、\-\u3001]")
    for line in lines:
        # 找到条款起始（支持 ### 第二十四条（人工费拨付周期） 或 第七条 xxx）
        if not found:
            stripped = line.strip()
            if (article_re.match(stripped) or
                stripped.startswith(article + " ") or
                stripped.startswith(article + "（") or
                (stripped.startswith(article) and not re.match(r"^第[一二三四五六七八九十百千零\d]+条", stripped[len(article):]))):
                found = True
                captured.append(stripped)
                continue
        if found:
            # 块引用中
            if line.startswith(">"):
                captured.append(line)
            # 下一个 ## 标题（章节）或下一个 ### 子章节
            elif line.strip().startswith("## "):
                break
            elif line.strip().startswith("### "):
                # 如果是另一条条款，也结束
                if re.match(r"^###\s+第[一二三四五六七八九十百千零\d]+条", line.strip()):
                    break
                captured.append(line)
            # 空行保留
            elif not line.strip():
                captured.append("")
            else:
                # 普通段落
                captured.append(line)
    return captured

def main():
    parser = argparse.ArgumentParser(description="按条款号查 wiki 法规")
    parser.add_argument("--law", required=True, help="法规名称（支持简写）")
    parser.add_argument("--article", required=True, help="条款号，如'第二十四条'")
    args = parser.parse_args()

    note = find_note(args.law)
    if not note:
        print(f"❌ 未找到法规: {args.law}")
        print("支持的法规（简写示例）:")
        for k in sorted(set(LAW_DIRS.keys())):
            print(f"  - {k}")
        return 1

    text = note.read_text(encoding="utf-8")
    captured = extract_article(text, args.article)

    print(f"📖 法规: {note.stem}")
    print(f"📁 路径: {note}")
    print(f"🔍 条款: {args.article}")
    print("─" * 60)
    if captured:
        for line in captured:
            print(line)
    else:
        print(f"⚠️ 未找到条款 {args.article}")
        # 列出所有条款
        all_articles = re.findall(r"^第[一二三四五六七八九十百千零\d]+条", text, re.MULTILINE)
        unique = sorted(set(all_articles))
        print(f"\n该法规共有 {len(unique)} 个条款:")
        for a in unique[:30]:
            print(f"  {a}")
        if len(unique) > 30:
            print(f"  ... 还有 {len(unique)-30} 条")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
