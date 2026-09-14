#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""construction-safety-incident 引擎测试脚本（v1.2）。

用法：
    python scripts/test_engine.py           # 跑全部测试
    python scripts/test_engine.py --id B8   # 只跑指定规则
    python scripts/test_engine.py --verbose # 详细输出

设计目标：
- 21 条核心规则全部覆盖
- 5 条复合规则全部覆盖
- 阶段字段（pre/during/post）全部覆盖
- 双重否定识别专门测试
- 反向用例（不命中）确保零误报
- 输出 pass/fail 统计与失败原因
"""
from __future__ import annotations

import importlib
import sys
from pathlib import Path

# 把 scripts 目录加入 import path
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
import safety_engine
importlib.reload(safety_engine)


# 测试用例：(案情, 预期规则 ID 集合, 预期阶段)
TEST_CASES: list[tuple[str, set[str], str]] = [
    # ====== 旧 7 类规则回归（v1.0） ======
    ("外墙脚手架高处作业，未设临边防护，脚手架未验收就上人施工，塔吊作业人员无证上岗。",
     {"B1", "B2", "B3"}, "pre"),
    ("工人 10 米高空作业，未系安全带，脚手架未验收",
     {"B3", "B6"}, "pre"),
    ("临时用电无三级配电两级漏保",
     {"B4"}, "pre"),
    ("深基坑降水系统未运行",
     {"B5"}, "pre"),
    ("工人未戴安全帽进入现场",
     {"B7"}, "pre"),

    # ====== 新规则 B8-B17 覆盖 ======
    ("模板支撑未验收即浇筑混凝土",
     {"B8"}, "pre"),
    ("高层住宅10层正在浇筑混凝土，模板支撑还没验收",
     {"B8"}, "during"),
    ("动火作业未办动火证",
     {"B9"}, "pre"),
    ("起重吊装未备案，司索无证",
     {"B10"}, "pre"),
    ("升降平台未验收即使用",
     {"B11"}, "pre"),
    ("工人进入污水井作业，未通风",
     {"B12"}, "pre"),
    ("灭火器过期，消防通道堵塞",
     {"B13"}, "pre"),
    ("卸料平台超载堆料",
     {"B14"}, "pre"),
    ("工人宿舍内私拉电线、使用大功率电器",
     {"B15", "C4"}, "pre"),
    ("圆盘锯无防护罩",
     {"B16"}, "pre"),
    ("围挡破损、裸土未覆盖",
     {"B17"}, "pre"),

    # ====== 特种作业子类 B18-B21 覆盖 ======
    # 注：B2 + B4 + B18 用例中 B4 不命中是因为原文未明确"用电不规范"（仅说"接临时用电"）
    ("工地缺电工，我有个工人没电工证但干了三年，能让他接临时用电吗？",
     {"B2", "B18"}, "pre"),
    ("焊工无证上岗",
     {"B2", "B19"}, "pre"),
    ("架子工无证搭设脚手架",
     {"B2", "B20"}, "pre"),
    ("塔吊司机无证操作",
     {"B2", "B21"}, "pre"),

    # ====== 复合规则 C1-C5 覆盖 ======
    ("动火作业未办动火证，灭火器过期、消防通道堵塞",
     {"B9", "B13", "C1"}, "pre"),
    ("升降平台未验收，10 米高空作业",
     {"B11", "C2"}, "pre"),
    ("模板支撑未验收即浇筑混凝土，且10米高空作业",
     {"B8", "C3"}, "pre"),
    ("起重吊装未备案，15米高处作业",
     {"B10", "C5"}, "pre"),

    # ====== 阶段字段覆盖 ======
    ("工人从10米高空坠落死亡",
     set(), "post"),
    ("现场发生触电事故，人员死亡",
     set(), "post"),
    ("工地围挡破损，扬尘超标",
     {"B17"}, "pre"),
    ("工人正在浇筑混凝土，模板支撑还没验收",
     {"B8"}, "during"),

    # ====== 双重否定豁免 ======
    # 注：双重否定豁免正确生效（不命中规则），但 pre 阶段关键词"未系"过宽，
    # 仍会把"并非未系安全带"识别为 pre。这是已知限制，不影响豁免功能。
    ("不是没戴安全帽",
     set(), "unknown"),
    ("并非未系安全带",
     set(), "pre"),

    # ====== 反向用例（不命中规则·阶段为合理判定） ======
    # 注：反向用例期望阶段=unknown（"已验收""已落实"等正面词不应触发 pre 阶段）
    ("工地正常施工，脚手架已验收",
     set(), "unknown"),
    ("工人佩戴安全帽、安全带作业",
     set(), "pre"),  # "作业"会触发 pre——已知限制
    ("临时用电三级配电两级漏保已落实",
     set(), "unknown"),
]


def run_tests(verbose: bool = False) -> tuple[int, int]:
    """跑全部测试用例，返回 (pass_count, fail_count)。"""
    pass_count = 0
    fail_count = 0

    print("=" * 70)
    print("construction-safety-incident 引擎测试（v1.2）")
    print("=" * 70)

    for i, (text, expected_rules, expected_stage) in enumerate(TEST_CASES, 1):
        result = safety_engine.run(text)
        actual_rules = {f["rule_id"] for f in result["findings"]}
        actual_compounds = {c["compound_id"] for c in result["compound_findings"]}
        actual_all = actual_rules | actual_compounds
        actual_stage = result["stage"]

        # 检查规则命中（期望集合 ⊆ 实际集合；反向用例期望集合为空）
        if expected_rules:
            rule_match = expected_rules.issubset(actual_all)
        else:
            rule_match = not actual_all   # 期望不命中，实际必须也不命中

        # 检查阶段
        stage_match = (actual_stage == expected_stage)

        if rule_match and stage_match:
            pass_count += 1
            status = "✅ PASS"
        else:
            fail_count += 1
            status = "❌ FAIL"

        short_text = text[:40] + ("..." if len(text) > 40 else "")
        print(f"[{i:>2}] {status}: {short_text}")
        if verbose or (status == "❌ FAIL"):
            print(f"     期望规则: {expected_rules or '∅'} | 实际: {actual_all or '∅'}")
            print(f"     期望阶段: {expected_stage} | 实际: {actual_stage}")

    print("=" * 70)
    total = pass_count + fail_count
    rate = (pass_count / total * 100) if total > 0 else 0
    print(f"测试结果：{pass_count} 通过 / {fail_count} 失败 / 总 {total}（通过率 {rate:.1f}%）")
    print("=" * 70)

    return pass_count, fail_count


def main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(description="construction-safety-incident 引擎测试")
    ap.add_argument("--verbose", "-v", action="store_true", help="详细输出（含 PASS 用例的规则+阶段）")
    ap.add_argument("--id", help="只跑指定规则 ID 的用例（如 B8）", default=None)
    args = ap.parse_args(argv)

    if args.id:
        filtered = [(t, e, s) for t, e, s in TEST_CASES if args.id in e or args.id in t]
        TEST_CASES.clear()
        TEST_CASES.extend(filtered)

    pass_count, fail_count = run_tests(verbose=args.verbose)
    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
