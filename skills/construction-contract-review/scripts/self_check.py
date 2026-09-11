"""确定性自测：同一份合同双跑，输出必须逐字节一致。

三项断言：
  1. 进程内连续两次 review_contract 结果完全一致；
  2. 两次独立子进程跑 CLI 的 JSON 输出逐字节一致；
  3. 样例合同的命中规则集符合黄金基线（防规则库被意外改坏）。

用法：python self_check.py
退出码 0 = 全部通过；非 0 = 有失败项（stdout 会打印细节）。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.join(HERE, "review_engine.py")
SAMPLE = os.path.join(HERE, "..", "assets", "sample_contract.txt")

# 黄金基线：样例合同必须命中的规则号（按现行 27 条规则库）
GOLDEN_RULE_IDS = {"C01", "C02", "C03", "C05", "C13", "C14", "C18", "C22", "C23", "C27"}

SAMPLE_CONTRACT_TEXT = """某住宅项目施工分包合同（样例）

第一条 工程概况
乙方承建某住宅楼主体结构工程。

第二条 合同价款
本合同采用固定总价包干，合同价款一次性包死，材料价格涨跌不予调整。

第三条 付款方式
以业主付款为前提，总包收到业主款项后向乙方支付进度款。
工程验收后支付，具体期限另行协商。

第四条 质量保证金
质量保证金按结算总额的 5% 预留，保修期满后无息退还。

第五条 保修责任
乙方对本工程承担终身保修责任。

第六条 优先受偿权
乙方放弃建设工程价款优先受偿权。

第七条 垫资与资金
乙方自筹资金组织施工，全额垫资至主体封顶。

第八条 用工管理
工人工资由乙方负责，与甲方无关。

第九条 工程分包
甲方同意乙方将本工程转包给具备资质的第三方施工。

第十条 结算依据
补充协议不一致以补充协议为准。
"""

failures: list = []


def check(name: str, ok: bool, detail: str = "") -> None:
    print(("PASS  " if ok else "FAIL  ") + name + (("  -> " + detail) if detail and not ok else ""))
    if not ok:
        failures.append(name)


def canonical(result: dict) -> str:
    return json.dumps(result, ensure_ascii=False, sort_keys=True)


def main() -> int:
    sys.path.insert(0, HERE)
    import review_engine as eng

    # 1) 进程内双跑一致
    r1 = eng.review_contract(SAMPLE_CONTRACT_TEXT)
    r2 = eng.review_contract(SAMPLE_CONTRACT_TEXT)
    check("进程内双跑结果一致", canonical(r1) == canonical(r2))

    # 2) 两次独立子进程 CLI JSON 输出逐字节一致
    py = sys.executable
    outs = []
    for _ in range(2):
        p = subprocess.run(
            [py, ENGINE, SAMPLE, "--format", "json"],
            capture_output=True, text=True, encoding="utf-8", cwd=HERE,
        )
        if p.returncode != 0:
            check("CLI 子进程运行", False, p.stderr[-300:])
            outs = None
            break
        outs.append(p.stdout)
    if outs is not None:
        check("CLI 双跑输出逐字节一致", outs[0] == outs[1])

    # 3) CLI md 输出双跑一致
    if outs is not None:
        mds = []
        for _ in range(2):
            p = subprocess.run(
                [py, ENGINE, SAMPLE, "--format", "md"],
                capture_output=True, text=True, encoding="utf-8", cwd=HERE,
            )
            mds.append(p.stdout if p.returncode == 0 else "ERR")
        check("CLI md 报告双跑一致", mds[0] == mds[1] and mds[0] != "ERR")

    # 4) 黄金基线：命中规则号与预期一致
    hit = {f["rule_id"] for f in r1["findings"]}
    check("命中规则集符合黄金基线", hit == GOLDEN_RULE_IDS,
          f"实际命中: {sorted(hit)}  缺少: {sorted(GOLDEN_RULE_IDS - hit)}  多出: {sorted(hit - GOLDEN_RULE_IDS)}")

    # 5) 豁免词生效：禁止性表述不误报（C23/C24）
    clean = eng.review_contract("第一条 分包管理\n本工程不得转包，禁止借用资质承接工程。\n")
    hit2 = {f["rule_id"] for f in clean["findings"]}
    check("禁止性表述不误报（无 C23/C24）", not ({"C23", "C24"} & hit2), f"误报: {sorted(hit2)}")

    print()
    if failures:
        print(f"自测失败 {len(failures)} 项：{'、'.join(failures)}")
        return 1
    print("全部通过：同一合同输入，输出确定性成立。")
    return 0


if __name__ == "__main__":
    sys.exit(main())