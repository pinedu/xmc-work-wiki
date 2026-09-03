"""
按关键词查 - 在 ~/wiki/ 下全文搜索关键词。
用法:
  python search_keyword.py --keyword "专户"
  python search_keyword.py --keyword "实名制" --context 3
  python search_keyword.py --keyword "罚款" --files_only
"""
import argparse
import re
from pathlib import Path

WIKI_ROOT = Path.home() / "wiki"

def main():
    parser = argparse.ArgumentParser(description="按关键词查 wiki 法规")
    parser.add_argument("--keyword", required=True, help="关键词")
    parser.add_argument("--context", type=int, default=2, help="显示前后行数（默认 2）")
    parser.add_argument("--files_only", action="store_true", help="只显示含关键词的文件名")
    parser.add_argument("--ignore_case", action="store_true", default=True, help="忽略大小写（默认）")
    args = parser.parse_args()

    md_files = list(WIKI_ROOT.rglob("*.md"))
    total_hits = 0
    results_by_file = {}

    flags = re.IGNORECASE if args.ignore_case else 0
    pattern = re.compile(re.escape(args.keyword), flags)

    for f in md_files:
        try:
            text = f.read_text(encoding="utf-8")
        except Exception:
            continue
        lines = text.splitlines()
        hits = []
        for i, line in enumerate(lines):
            if pattern.search(line):
                start = max(0, i - args.context)
                end = min(len(lines), i + args.context + 1)
                ctx = lines[start:end]
                hits.append((i + 1, ctx))
        if hits:
            results_by_file[f] = hits
            total_hits += len(hits)

    if not results_by_file:
        print(f"❌ 未找到关键词: {args.keyword}")
        return 1

    print(f"🔍 关键词: {args.keyword}")
    print(f"📊 命中: {total_hits} 处 / {len(results_by_file)} 个文件")
    print("─" * 60)

    for f, hits in results_by_file.items():
        rel = f.relative_to(WIKI_ROOT)
        if args.files_only:
            print(f"📄 {rel} ({len(hits)} 处)")
        else:
            print(f"\n📄 {rel} ({len(hits)} 处):")
            for lineno, ctx in hits:
                print(f"  L{lineno}:")
                for line in ctx:
                    if pattern.search(line):
                        print(f"    >>> {line[:200]}")
                    else:
                        print(f"        {line[:200]}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
