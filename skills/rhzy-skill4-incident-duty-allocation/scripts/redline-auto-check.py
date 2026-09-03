#!/usr/bin/env python3
"""
红线一票否决自动校验脚本

输入：企业 merged JSON 文件（来自 rhzy-contractor-due-diligence）
输出：是否触发一票否决 + 触发的红线项 + 综合评级

用法：
    python redline-auto-check.py <企业JSON文件路径>
    python redline-auto-check.py <企业JSON文件路径> --json  # 输出 JSON 格式

退出码：
    0 = 不触发一票否决
    1 = 触发一票否决
    2 = 输入文件错误
"""

import sys
import json
import argparse
from datetime import datetime, date
from pathlib import Path


def parse_date(date_str):
    """解析 YYYY-MM-DD 格式的日期字符串"""
    if not date_str:
        return None
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def days_until(target_date):
    """距离 target_date 还有多少天（负数=已过期）"""
    if not target_date:
        return None
    return (target_date - date.today()).days


def check_red_lines(company):
    """
    检查企业是否触发一票否决红线

    Args:
        company: 企业 merged JSON（来自 rhzy-contractor-due-diligence）

    Returns:
        dict: {
            "triggered": bool,
            "red_lines": list of triggered red lines,
            "risk_grade": "red" | "orange" | "yellow" | "blue",
            "suggestion": str
        }
    """
    red_lines = []
    today = date.today()

    # 红线 1：安全生产许可证已过期
    safety_license_expiry = parse_date(
        company.get("qualification", {}).get("safety_license_expiry")
    )
    if safety_license_expiry and safety_license_expiry < today:
        red_lines.append({
            "id": "RL-001",
            "name": "安全生产许可证已过期",
            "value": str(safety_license_expiry),
            "action": "一票否决·禁止使用"
        })

    # 红线 2：安全生产许可证剩余有效期 < 2 个月（即将过期）
    elif safety_license_expiry:
        days = days_until(safety_license_expiry)
        if days is not None and days < 60:  # 60 天 ≈ 2 个月
            red_lines.append({
                "id": "RL-002",
                "name": "安全生产许可证即将过期（剩余 < 2 个月）",
                "value": f"剩余 {days} 天 / 到期日 {safety_license_expiry}",
                "action": "一票否决·禁止使用"
            })

    # 红线 3：近 1 年行政处罚 ≥ 4 次
    penalty_items = company.get("administrative_penalty", {}).get("items", [])
    one_year_penalty_count = 0
    one_year_ago = date.today().replace(year=date.today().year - 1)
    for item in penalty_items:
        decision_date = parse_date(item.get("decision_date"))
        if decision_date and decision_date >= one_year_ago:
            one_year_penalty_count += 1

    if one_year_penalty_count >= 4:
        red_lines.append({
            "id": "RL-003",
            "name": "近 1 年行政处罚 ≥ 4 次",
            "value": f"{one_year_penalty_count} 次",
            "action": "一票否决·禁止使用"
        })

    # 红线 4：失信被执行人
    is_dishonesty = company.get("judicial", {}).get("is_dishonesty", False)
    if is_dishonesty:
        red_lines.append({
            "id": "RL-004",
            "name": "失信被执行人",
            "value": "是",
            "action": "一票否决·禁止使用"
        })

    # 红线 5：重大税收违法
    major_tax_violation = company.get("finance", {}).get("major_tax_violation", False)
    if major_tax_violation:
        red_lines.append({
            "id": "RL-005",
            "name": "重大税收违法",
            "value": "是",
            "action": "一票否决·禁止使用"
        })

    # 红线 6：资质缺失
    main_qualification = company.get("qualification", {}).get("main_qualification")
    if not main_qualification:
        red_lines.append({
            "id": "RL-006",
            "name": "主要资质缺失",
            "value": "无",
            "action": "一票否决·禁止使用"
        })

    # 计算七维评分（简化版，仅用于不触发一票否决时的参考）
    risk_score = 0
    if not red_lines:
        # 司法风险 20%
        litigation_by_type = company.get("judicial", {}).get("litigation_by_type", {})
        labor_dispute_ratio = litigation_by_type.get("labor_dispute", 0) / max(
            sum(litigation_by_type.values()), 1
        )
        risk_score += labor_dispute_ratio * 20

        # 行政处罚 20%
        if one_year_penalty_count == 3:
            risk_score += 20  # =3 次触发高风险（<4 次不触发一票否决）
        elif one_year_penalty_count == 2:
            risk_score += 13
        elif one_year_penalty_count == 1:
            risk_score += 7

        # 财务健康 15%
        paid_in_ratio = company.get("finance", {}).get("paid_in_capital_ratio", 1.0)
        if paid_in_ratio < 0.3:
            risk_score += 15
        elif paid_in_ratio < 0.6:
            risk_score += 10

    # 综合评级
    if red_lines:
        risk_grade = "red"
    elif risk_score >= 60:
        risk_grade = "orange"
    elif risk_score >= 30:
        risk_grade = "yellow"
    else:
        risk_grade = "blue"

    # 建议
    if red_lines:
        suggestion = "🚨 一票否决·禁止使用——出具《关于排除该单位投标资格的说明》+ 启动备选单位核查"
    elif risk_grade == "orange":
        suggestion = "⚠️ 橙色高风险——专项尽调后复议 / 合同附额外约束条款（履约保证金/银行保函）"
    elif risk_grade == "yellow":
        suggestion = "🟡 黄色中风险——可招标 + 合同附额外约束条款"
    else:
        suggestion = "🔵 蓝色低风险——可正常招标"

    return {
        "triggered": len(red_lines) > 0,
        "red_lines": red_lines,
        "risk_grade": risk_grade,
        "risk_score": round(risk_score, 2),
        "suggestion": suggestion
    }


def main():
    parser = argparse.ArgumentParser(description="企业一票否决红线自动校验")
    parser.add_argument("json_file", help="企业 merged JSON 文件路径")
    parser.add_argument("--json", action="store_true", help="输出 JSON 格式")
    args = parser.parse_args()

    json_path = Path(args.json_file)
    if not json_path.exists():
        print(f"❌ 文件不存在：{json_path}", file=sys.stderr)
        return 2

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            company = json.load(f)
    except json.JSONDecodeError as e:
        print(f"❌ JSON 解析失败：{e}", file=sys.stderr)
        return 2

    result = check_red_lines(company)

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        company_name = company.get("name", "未知企业")
        print("=" * 60)
        print(f"企业一票否决红线核查报告")
        print("=" * 60)
        print(f"企业名称：{company_name}")
        print(f"核查日期：{date.today()}")
        print()
        print(f"一票否决触发：{'🚨 是' if result['triggered'] else '✅ 否'}")
        print(f"综合评级：{result['risk_grade'].upper()}")
        if not result['triggered']:
            print(f"风险评分：{result['risk_score']}")
        print()
        if result['red_lines']:
            print("触发的红线：")
            for rl in result['red_lines']:
                print(f"  • [{rl['id']}] {rl['name']} = {rl['value']}")
                print(f"    → {rl['action']}")
            print()
        print(f"建议：{result['suggestion']}")
        print("=" * 60)

    return 1 if result['triggered'] else 0


if __name__ == "__main__":
    sys.exit(main())


# === 测试用例 ===
# 1. 触发一票否决：安全生产许可证 2 个月内到期
# {
#   "name": "贵州宏达建筑",
#   "qualification": {
#     "main_qualification": "建筑工程施工总承包二级",
#     "safety_license_expiry": "2026-09-15"
#   },
#   "administrative_penalty": {"items": [
#     {"decision_date": "2025-12-01"},
#     {"decision_date": "2026-01-15"},
#     {"decision_date": "2026-03-20"},
#     {"decision_date": "2026-05-10"}
#   ]}
# }
# 预期：触发 RL-002 + RL-003 → 一票否决
