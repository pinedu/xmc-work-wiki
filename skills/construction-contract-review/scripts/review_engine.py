"""建筑领域合同条款审查引擎（确定性，27 条规则）。

设计铁律：
  - 纯标准库 + 固定规则，不依赖任何模型 → 同一份合同输入，输出逐字节一致；
  - 规则迭代按固定顺序遍历、发现按固定顺序追加 → 无集合/字典乱序；
  - 判定只依据：关键词命中（any_kw / force_kw）、阈值（max_pct / max_year）、
    等级上调词（high_kw）、豁免词（exclude_kw），全部显式、无随机性。

命令行：
    python review_engine.py <合同.txt | 合同.docx | -> [--format json|md]
                                           [--stage1-only] [--stage1-output FILE]
                                           [--stage2-input verdicts.json]
                                           [--no-stage2]

    - 表示从标准输入读文本；--format md 输出固定格式的审查报告。
    v1.3.0 起支持 LLM 复核（Stage 2）：
      --stage1-only 输出含 findings_raw + clauses 的 JSON，供平台 agent 调 LLM 复核；
      --stage2-input verdicts.json 读入 LLM 复核 verdicts，合并后出最终报告。

产出结构：{"clauses", "findings", "summary", "meta"}，JSON 输出按键排序。
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import List, Optional

ENGINE_NAME = "construction-contract-review"
ENGINE_VERSION = "1.3.0"
RULES_VERSION = "27条/2026-09-10 + v1.3.0 LLM 复核接口"

# ---------------------------------------------------------------- 条款切分

# 「第X条」形式优先；其次行首编号（1. / 1、/ (1) / 一、）
_RE_ARTICLE = re.compile(r"第[一二三四五六七八九十百千零〇\d]+条")
_RE_NUMBERED = re.compile(r"^\s*(?:\(?\d+\)?[.、]|[一二三四五六七八九十]+[、.])")

# 子条款编号：22.2 / 22.3 / 22、3 / (2) 等，匹配"第X条"正文内的子项
# 优先级：数字点（22.2）> 数字顿（22、3）> 括号数字（(2)）> 圈数字（②）
_RE_SUB_DOT = re.compile(r"(?<!\d)(\d{1,3})\.(\d+)(?!\d)")           # 22.2 / 22.10
_RE_SUB_COMMA = re.compile(r"(?<!\d)(\d{1,3})[、、](?=\d)")          # 22、3
_RE_SUB_PAREN = re.compile(r"(?<!\w)\((\d+)\)(?!\w)")                # (2) / (10)
_RE_SUB_CIRCLE = re.compile(r"[②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯]")              # ②③④


def _split_sub_clauses(article_id: str, body: str) -> List[dict]:
    """把一个整条（body）进一步切分为子条款，返回 [{"id","title","text"}]。

    子条款 id 格式：「第22.2条」（条号 + 子编号）。
    若正文内无任何子条款编号模式，则返回整条本身（id 不变）。
    支持嵌套：先按 . 点号拆分，再处理 (2) / ② 等子项。
    """
    # 优先检测是否存在子条款编号
    has_dot = _RE_SUB_DOT.search(body)
    has_comma = _RE_SUB_COMMA.search(body)
    has_paren = _RE_SUB_PAREN.search(body)
    has_circle = _RE_SUB_CIRCLE.search(body)

    if not (has_dot or has_comma or has_paren or has_circle):
        # 无子条款，整条返回
        return [{"id": article_id, "title": body.split("\n", 1)[0].strip()[:40], "text": body.strip()}]

    # 将 body 按行拆分，识别每行是否以子条款编号开头
    lines = body.split("\n")
    sub_clauses: List[dict] = []
    cur_buf: List[str] = []
    cur_sub_id: str = ""
    cur_sub_title: str = ""

    def _flush():
        nonlocal cur_buf, cur_sub_id, cur_sub_title
        if cur_buf and cur_sub_id:
            text = "\n".join(cur_buf).strip()
            sub_clauses.append({"id": cur_sub_id, "title": cur_sub_title[:40], "text": text})
        cur_buf, cur_sub_title = [], ""

    def _make_sub_id(article_id: str, sub_str: str) -> str:
        """生成子条款 id，如「第22条」+「2」→「第22.2条」。

        sub_str 的格式取决于来源：
          - _RE_SUB_DOT：完整 article.sub，如 "21.1"（含文章号），直接拼入；
          - _RE_SUB_COMMA：完整 article.sub，如 "21.3"；
          - _RE_SUB_PAREN：纯子编号，如 "2"（无文章号），用 article_id 拼接；
          - _RE_SUB_CIRCLE：纯子编号，如 "2"。
        判定方式：若 sub_str 包含 '.'，则已经是 article.sub，直接用；否则拼 article_id。
        """
        stripped_article = article_id.rstrip("条")
        if "." in sub_str:
            # 来自 _RE_SUB_DOT / _RE_SUB_COMMA，已含完整 article.sub
            return f"{stripped_article}.{sub_str}条"
        else:
            # 来自 _RE_SUB_PAREN / _RE_SUB_CIRCLE，纯子编号
            return f"{stripped_article}.{sub_str}条"

    def _is_sub_start(line: str) -> tuple:
        """判断一行是否开始新的子条款。返回 (bool, sub_str)。

        sub_str 格式：仅含子编号（无文章号），由 _make_sub_id 统一拼入文章号。
        例：第21.1 → _is_sub_start 返回 sub_str="1"（不是"21.1"）。
        """
        stripped = line.strip()
        if not stripped:
            return False, ""
        # 22.2 形式：group(1)=文章号，group(2)=子编号；_make_sub_id 统一拼接，只传子编号
        m = _RE_SUB_DOT.match(stripped)
        if m:
            return True, m.group(2)   # 只传子编号 "1"，不是 "21.1"
        # 22、3 形式：group(1)=文章号；只传子编号部分
        m = _RE_SUB_COMMA.match(stripped)
        if m:
            # sub_str 来自逗号：group(0) 是 "21、3"，取最后一个数字字符
            return True, stripped[len(m.group(1)):].lstrip("、")  # 提取 "3"
        # (2) 形式
        m = _RE_SUB_PAREN.match(stripped)
        if m:
            return True, m.group(1)
        # ②③④ 形式（单独一行时）
        if _RE_SUB_CIRCLE.match(stripped) and len(stripped) <= 3:
            circle_map = {"②": "2", "③": "3", "④": "4", "⑤": "5",
                          "⑥": "6", "⑦": "7", "⑧": "8", "⑨": "9",
                          "⑩": "10", "⑪": "11", "⑫": "12",
                          "⑬": "13", "⑭": "14", "⑮": "15", "⑯": "16"}
            ch = stripped[0]
            return True, circle_map.get(ch, "")
        return False, ""

    for ln in lines:
        is_new, sub_str = _is_sub_start(ln)
        if is_new and sub_str:
            _flush()
            cur_sub_id = _make_sub_id(article_id, sub_str)
            cur_sub_title = ln.strip()
            cur_buf = [ln]
        elif cur_sub_id:
            cur_buf.append(ln)
        else:
            # 第一个子条款出现之前的开头内容（条款标题行等），合并进第一个子条款
            if cur_buf or sub_clauses:
                cur_buf.append(ln)
            else:
                cur_buf.append(ln)

    _flush()

    if not sub_clauses:
        # 保底：整条返回
        return [{"id": article_id, "title": body.split("\n", 1)[0].strip()[:40], "text": body.strip()}]

    return sub_clauses


def split_clauses(text: str) -> List[dict]:
    """把合同文本切成条款列表 [{"id","title","text"}]。

    两级切分（条 → 子条款）：
      第一步：按「第X条」切分为整条；
      第二步：若整条正文内存在子条款编号（22.2 / 22、3 / (2) / ② 等），
              则进一步拆分为子条款，子条款 id 格式为「第22.2条」。
    若无「第X条」形式，回退至行首编号 → 整篇作为一条。
    """
    text = (text or "").strip()
    if not text:
        return []

    marks = list(_RE_ARTICLE.finditer(text))
    if marks:
        out: List[dict] = []
        for i, m in enumerate(marks):
            end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
            body = text[m.start():end].strip()
            article_id = m.group(0)
            # 第二级：子条款拆分
            sub_items = _split_sub_clauses(article_id, body)
            out.extend(sub_items)
        if out:
            return out

    lines = text.split("\n")
    out, cur_id, buf = [], None, []
    for ln in lines:
        if _RE_NUMBERED.match(ln) and ln.strip():
            if buf and cur_id:
                out.append({"id": cur_id, "title": buf[0][:40], "text": "\n".join(buf).strip()})
            cur_id = ln.strip()[:20]
            buf = [ln]
        elif cur_id:
            buf.append(ln)
    if buf and cur_id:
        out.append({"id": cur_id, "title": buf[0][:40], "text": "\n".join(buf).strip()})
    if out:
        return out

    return [{"id": "全文", "title": "合同全文", "text": text}]


# ---------------------------------------------------------------- 规则库

@dataclass
class ClauseRule:
    """一条审查规则。

    any_kw 任一命中即进入判定；exclude_kw 任一命中则豁免（排除"不得转包"
    这类合规表述误报）；force_kw / max_pct / max_year 细化判定；
    high_kw 命中则等级上调为 level_high。
    """
    id: str
    name: str
    any_kw: tuple
    risk_type: str
    level: str                      # 一般风险 / 较大风险 / 重大风险
    basis: str                      # 法规或惯例依据（法规用全称+条号）
    suggestion: str                 # 修订建议
    exclude_kw: tuple = ()          # 命中则豁免本规则
    max_pct: Optional[float] = None  # 提取到百分比 > max_pct 才算问题
    max_year: Optional[float] = None # 提取到"X年" > max_year 才算问题
    force_kw: tuple = ()             # 命中则直接判定（不依赖阈值）
    high_kw: tuple = ()              # 命中则等级上调一档
    level_high: str = "较大风险"


RULES: List[ClauseRule] = [
    # ---------- 原有 21 条（与项目 clause_review.py 保持一致） ----------
    ClauseRule(
        id="C01", name="背靠背付款条款",
        any_kw=("背靠背", "以业主付款为前提", "待业主支付", "业主未付款", "收到业主款项后"),
        risk_type="合同收款风险",
        level="重大风险",
        basis="《保障农民工工资支付条例》及合同公平原则：总包不得将业主付款风险完全转嫁分包",
        suggestion="改为「总包收到业主款项后 X 日内支付」，并增设最迟付款期限（如验收后 90 日），"
                   "避免以业主未付为由无限期拖延",
    ),
    ClauseRule(
        id="C02", name="付款节点/期限约定不明",
        any_kw=("验收后支付", "结算后支付", "竣工后支付", "审计后支付"),
        risk_type="合同收款风险",
        level="较大风险",
        basis="《中华人民共和国民法典》第五百一十一条：履行期限不明确的，债务人可随时履行、"
              "债权人也可随时请求履行",
        suggestion="明确具体期限与比例，如「验收合格后 30 日内支付至结算价的 97%」，"
                   "并约定逾期利息",
    ),
    ClauseRule(
        id="C03", name="质量保证金预留比例超标",
        any_kw=("质保金", "质量保证金", "保留金", "质量保留金"),
        risk_type="合同资金占用风险",
        level="一般风险",
        basis="《建设工程质量保证金管理办法》第七条：预留比例不得高于工程价款结算总额的 3%",
        suggestion="将预留比例降至 3% 以内，或以银行保函替代现金预留，释放现金流",
        max_pct=3.0,
        # v1.2.1：3% 是法定上限，遇到"5%/8%/10%"等才真报；fix 3% 不误报
        # 同时排除含"质保金按 ... 3% 预留 / 缺陷责任期满 ... 返还"等合规表述
        exclude_kw=("按工程价款结算总额的 3%", "按结算价款的 3%",
                    "按结算总额的 3%", "按 3% 预留", "3% 作为质量保证金",
                    "质量保证金按工程价款结算总额的 3%"),
        force_kw=("5%", "5％", "8%", "8％", "10%", "10％", "15%", "15％", "20%", "20％"),
    ),
    ClauseRule(
        id="C04", name="甲供材约定不明",
        any_kw=("甲供材", "甲方供应材料", "发包人供应材料", "甲供材料"),
        risk_type="合同计价风险",
        level="一般风险",
        basis="行业惯例：甲供材的计价、损耗、保管责任须明确，否则结算时易生争议",
        suggestion="明确甲供材的计价方式（是否计入合同价）、损耗率上限、超耗承担方、"
                   "到场验收与保管责任归属",
        # v1.2.1：仅在条款含"损耗率/超耗/保管/结算方式"等真"约定不明"特征时触发；
        # 条款已约定计价+损耗率+保管责任的不报
        exclude_kw=("损耗率按定额", "超耗部分", "保管责任", "甲供材计入合同价款",
                    "到场验收", "损耗率按", "计入合同价款"),
    ),
    ClauseRule(
        id="C05", name="材料价格不予调整（无调差机制）",
        any_kw=("固定总价", "总价包干", "不予调整", "不做调整", "一次性包死"),
        risk_type="材料价格波动风险",
        level="较大风险",
        basis="《中华人民共和国民法典》第五百三十三条情势变更：重大变化继续履行显失公平的，"
              "可请求变更或解除",
        suggestion="增设主要材料价格调差条款，约定涨跌幅超过 ±5% 时按造价信息调整，"
                   "并明确调差基准期与计算方式",
    ),
    ClauseRule(
        id="C06", name="变更/签证计价原则缺失",
        any_kw=("现场签证", "工程变更", "设计变更", "变更估价"),
        risk_type="变更结算风险",
        level="较大风险",
        basis="《建设工程价款结算暂行办法》：变更价款应按合同约定的计价原则确定",
        suggestion="明确变更估价顺序：参照合同已有单价 → 类似单价 → 定额组价 → 市场价，"
                   "并约定签证确认时限与逾期视为认可规则",
        # v1.2.1：条款已明确估价顺序四档 + 签证确认时限 + 逾期视为认可的应排除
        # 第1条"工期条款"含"设计变更"四字但只谈顺延不谈计价 → 应排除
        # 第5条已含完整估价顺序 + 14日时限 + 视为认可 → 应排除
        exclude_kw=("合同已有适用单价", "合同类似单价", "定额组价", "市场询价",
                    "14 日内", "14日内", "逾期未答复", "视为认可"),
    ),
    ClauseRule(
        id="C07", name="工期索赔权被排除",
        any_kw=("工期不予顺延", "放弃工期索赔", "任何情况不顺延", "不得要求顺延",
                "工期一律不顺延"),
        risk_type="工期风险",
        level="重大风险",
        basis="《中华人民共和国民法典》第四百九十七条：不合理免除己方责任、排除对方主要权利的"
              "格式条款无效；第五百九十条：不可抗力可部分或全部免责",
        suggestion="限定为「因承包方自身原因不予顺延」，保留业主原因、不可抗力、"
                   "设计变更等情形下的工期顺延与费用索赔权",
        # v1.2.2：条款已约定发包人原因/设计变更/不可抗力顺延权的应排除
        # 区分"完整条款（顺延权已保留）"与"真排除条款（一刀切不顺延）"
        # ⚠️ 注意排除词必须区分"保留顺延权"vs"放弃顺延权"——
        # "放弃工期索赔"在 any_kw 里命中，但同时会误命中 exclude_kw
        # 因此排除词只用"正面的保留表述"，不用含"放弃/不/不得"的词
        exclude_kw=("相应顺延", "经确认后顺延", "经发包人书面确认后顺延",
                    "工期相应顺延", "有权主张工期", "有权顺延"),
    ),
    ClauseRule(
        id="C08", name="违约责任不对称或违约金畸高",
        any_kw=("违约金", "滞纳金", "逾期付款", "逾期竣工"),
        risk_type="违约风险",
        level="一般风险",
        basis="《中华人民共和国民法典》第五百八十五条：违约金过分高于造成的损失的，"
              "可请求人民法院或仲裁机构适当减少",
        suggestion="约定对等的双向违约责任，比例控制在合理区间（一般日万分之五以内），"
                   "并设置赔偿上限",
        high_kw=("千分之", "每日千分之", "日千分之", "百分之十", "10%"),
        level_high="较大风险",
        # v1.2.1：日万分之三/万分之五在合理区间不报；日千分之一及以上、累计 ≥10% 才报
        # 实现：仅靠 force_kw 命中，不走 max_pct 路径
        # （max_pct 不适用：因为同条款常含"累计不超过 X%"，与日费率无法区分）
        force_kw=("千分之", "千分之二", "千分之五", "千分之十",
                  "百分之十", "百分之十五", "百分之二十"),
    ),
    ClauseRule(
        id="C09", name="争议解决/管辖约定不利",
        any_kw=("甲方所在地", "发包人所在地", "仲裁", "管辖"),
        risk_type="争议解决风险",
        level="一般风险",
        basis="《中华人民共和国民事诉讼法》：合同纠纷可由合同履行地法院管辖",
        suggestion="争取约定工程所在地人民法院管辖，降低异地维权成本；"
                   "如选仲裁，明确仲裁机构名称与适用规则",
        # v1.2.1：已约定工程所在地法院管辖的应排除（这是甲方最优安排）
        exclude_kw=("工程所在地有管辖权", "工程所在地人民法院", "工程所在地法院",
                    "合同履行地"),
    ),
    ClauseRule(
        id="C10", name="结算以单方审计结论为准",
        any_kw=("以甲方审计", "以发包人审计", "最终结算以审计", "审计结果为准"),
        risk_type="结算风险",
        level="较大风险",
        basis="审计结论不当然等同于双方结算合意，单方定价易显失公平",
        suggestion="约定审计期限（如收到结算资料后 60 日）、异议与复核程序，"
                   "以及逾期未出具审计意见视为认可送审价",
    ),
    ClauseRule(
        id="C11", name="安全责任全额转嫁",
        any_kw=("安全事故责任均由", "安全责任均由乙方", "甲方不承担任何安全",
                "一切安全事故由承包方承担"),
        risk_type="安全生产风险",
        level="重大风险",
        basis="《中华人民共和国安全生产法》：生产经营单位对安全生产负主体责任，"
              "不得通过协议免除法定责任",
        suggestion="按过错划分责任，明确发包方依法应履行的安全管理与协调职责，"
                   "删除全额转嫁表述（该约定可能因违法而无效）",
    ),
    ClauseRule(
        id="C12", name="赔偿责任无上限/无限兜底",
        any_kw=("承担一切损失", "无条件承担", "全部损失", "一切责任"),
        risk_type="责任范围风险",
        level="较大风险",
        basis="《中华人民共和国民法典》第五百八十四条：赔偿以可预见的可得利益损失为限",
        suggestion="限定赔偿范围为直接损失，设置责任上限（如不超过合同价款的 10%），"
                   "并排除间接损失与可得利益",
    ),
    ClauseRule(
        id="C13", name="放弃建设工程价款优先受偿权",
        any_kw=("放弃优先受偿权", "放弃建设工程价款优先受偿权", "不主张优先受偿权",
                "优先受偿权由", "让渡优先受偿权", "不享有优先受偿权"),
        risk_type="工程款保障风险",
        level="重大风险",
        basis="《中华人民共和国民法典》第八百零七条（承包人优先受偿权）；"
              "最高人民法院《关于审理建设工程施工合同纠纷案件适用法律问题的解释（一）》"
              "第四十一条（行使期限 18 个月）、第四十二条（损害建筑工人利益的放弃约定无效）",
        suggestion="不得预先放弃优先受偿权；如确需限制须约定不损害建筑工人工资权益，"
                   "且保留在发包人欠付范围内主张的权利",
    ),
    ClauseRule(
        id="C14", name="垫资施工",
        any_kw=("垫资", "带资承包", "带资", "乙方自筹资金", "全额垫资", "垫付款"),
        risk_type="资金周转风险",
        level="重大风险",
        basis="《保障中小企业款项支付条例》第九条：机关、事业单位和大型企业不得要求施工单位"
              "垫资；垫资在司法上原则有效但风险极高",
        suggestion="明确垫资额度上限、垫资利息（不低于同期贷款市场报价利率 LPR）、回收节点与"
                   "逾期违约责任；政府/国企项目不得要求施工单位垫资",
    ),
    ClauseRule(
        id="C15", name="停工权/抗辩权限制",
        any_kw=("不得停工", "放弃停工", "无论是否付款均不得停工", "放弃抗辩权",
                "不得行使抗辩权", "不得要求发包人承担"),
        risk_type="履约抗辩风险",
        level="重大风险",
        basis="《中华人民共和国民法典》第五百二十五条（同时履行抗辩权）、第五百二十七条"
              "（不安抗辩权）：发包人未按约付款时承包人依法享有停工与顺延抗辩权",
        suggestion="保留承包人在发包人逾期付款时的停工/顺延抗辩权，约定「发包人逾期付款超过 X 日，"
                   "承包人有权暂停施工且工期顺延、费用由发包人承担」",
    ),
    ClauseRule(
        id="C16", name="发票与税务转嫁",
        any_kw=("先开票后付款", "以发票作为付款前提", "发票作为付款前提", "一切税费由乙方承担",
                "乙方负责代扣代缴", "税费由乙方承担"),
        risk_type="税务与票据风险",
        level="一般风险",
        basis="纳税义务法定，不得通过约定转嫁；《保障中小企业款项支付条例》第九条：不得强制以"
              "审计、发票等为由延期付款",
        suggestion="付款义务与开票义务分离——付款不以开票为前提；税费按法定各自承担，"
                   "不得约定「一切税费由乙方承担」",
        high_kw=("先开票后付款", "以发票作为付款前提", "发票作为付款前提"),
        level_high="较大风险",
    ),
    ClauseRule(
        id="C17", name="索赔逾期失权（28天条款）",
        any_kw=("逾期视为放弃", "逾期作废", "视为放弃索赔", "未在约定期限内提出视为放弃",
                "28天内未提出索赔视为放弃"),
        risk_type="索赔时效风险",
        level="较大风险",
        basis="《建设工程施工合同（示范文本）》（GF-2017-0201）通用条款第 19 条（索赔期限）；"
              "逾期失权条款显失公平的可被调整",
        suggestion="保留合理索赔期限并设兜底：因发包人原因致承包人未能按期索赔的不构成失权；"
                   "期限不宜过短并约定逾期异议程序",
    ),
    ClauseRule(
        id="C18", name="保修期/保修责任超法定",
        any_kw=("保修期", "保修期限", "质量保修", "保修"),
        risk_type="质量保修风险",
        level="一般风险",
        basis="《建设工程质量管理条例》第四十条：最低保修期限（防水 5 年、供热供冷 2 个周期、"
              "电气/给排水/装修 2 年、地基基础与主体结构为设计文件合理使用年限）",
        suggestion="保修期不得低于法定最低期限；超出部分（如「整体终身保修」）须限定在法定范围内，"
                   "明确保修金返还节点",
        max_year=5.0,
        force_kw=("终身保修", "永久保修"),
    ),
    ClauseRule(
        id="C19", name="履约/投标保证金异常",
        any_kw=("履约保证金", "投标保证金", "保证金不予退还", "高额保证金"),
        risk_type="保证金风险",
        level="一般风险",
        basis="《保障中小企业款项支付条例》第十二条：保证金不得超过合同金额 10%，"
              "推行银行保函替代现金保证金",
        suggestion="明确保证金比例（履约不超过 10%）、退还节点与条件，优先采用银行保函替代现金，"
                   "约定逾期退还违约责任",
    ),
    ClauseRule(
        id="C20", name="工程量清单错漏由承包人承担",
        any_kw=("清单漏项", "量差", "投标人自行复核", "清单错误由承包人承担", "一切以现场为准"),
        risk_type="清单计价风险",
        level="较大风险",
        basis="《建设工程工程量清单计价规范》（GB 50500-2013）第 9.3、9.4 条；"
              "最高人民法院《关于审理建设工程施工合同纠纷案件适用法律问题的解释（一）》"
              "第十九条：工程量有争议按施工过程证据认定",
        suggestion="约定「发包人提供的清单错漏、项目特征不符按实调整合同价款」，"
                   "不以「投标人已复核」免除发包人责任",
    ),
    ClauseRule(
        id="C21", name="竣工资料移交卡结算/备案",
        any_kw=("未移交资料不予结算", "以资料齐全为付款前提", "不配合备案", "未提交竣工资料"),
        risk_type="结算抗辩风险",
        level="较大风险",
        basis="《中华人民共和国民法典》第七百九十九条：建设工程验收合格后发包人应支付价款；"
              "资料移交是义务但不得作为拒付工程款的抗辩",
        suggestion="结算不以资料移交为前提；约定资料移交时限与发包人配合备案义务，逾期视为认可",
    ),
    # ---------- 新增 6 条（2026-09-10 扩充定稿） ----------
    ClauseRule(
        id="C22", name="农民工工资支付责任转嫁",
        any_kw=("工资支付责任由乙方", "工资支付责任由分包", "由分包自行支付", "分包自行负责支付",
                "工资由分包承担", "工人工资由乙方负责", "工资由乙方自行解决"),
        risk_type="用工与工资支付风险",
        level="重大风险",
        basis="《保障农民工工资支付条例》第二十四条、第三十条：施工总承包单位对分包单位"
              "劳动用工和工资支付负监督责任，对拖欠农民工工资先行清偿，不得通过约定转嫁",
        suggestion="删除转嫁表述；按条例实行农民工工资专用账户管理与施工总承包单位代发，"
                   "总包先行清偿责任不得约定排除",
    ),
    ClauseRule(
        id="C23", name="违法转包/再分包",
        any_kw=("转包", "再分包", "二次分包", "层层分包", "允许分包单位再分包"),
        exclude_kw=("不得转包", "禁止转包", "严禁转包", "禁止再分包", "不得再分包",
                    "严禁再分包", "不得违法转包", "禁止违法转包"),
        risk_type="合同效力风险",
        level="重大风险",
        basis="《中华人民共和国建筑法》第二十八条、第二十九条；《中华人民共和国民法典》"
              "第七百九十一条：禁止转包和违法分包，转包约定无效",
        suggestion="删除转包/再分包安排；确需专业分包的，限于总承包合同约定或经建设单位认可"
                   "的专业工程，且不得再分包（劳务作业除外）",
    ),
    ClauseRule(
        id="C24", name="借用资质/挂靠",
        any_kw=("挂靠", "借用资质", "以他人名义投标", "以其他单位名义承揽",
                "资质证书借用", "出借资质"),
        exclude_kw=("禁止挂靠", "不得挂靠", "严禁挂靠", "禁止借用资质", "不得借用资质",
                    "严禁借用资质", "禁止出借", "不得出借"),
        risk_type="合同效力风险",
        level="重大风险",
        basis="《中华人民共和国建筑法》第二十六条：禁止以他人名义承揽工程、禁止出借资质；"
              "借用资质签订的施工合同无效",
        suggestion="删除挂靠/借资质安排，由实际施工主体以自有资质签约；"
                   "涉及资质联合经营的须符合《建筑法》关于联合体承包的规定",
    ),
    ClauseRule(
        id="C25", name="发包人单方随时变更权",
        any_kw=("甲方有权随时变更", "发包人有权随时变更", "无需承包人同意即可变更",
                "无需乙方同意即可调整", "甲方可单方变更", "发包人可单方调整",
                "单方调整设计方案"),
        risk_type="变更与工期价款风险",
        level="较大风险",
        basis="《中华人民共和国民法典》第五百零九条：按约定全面履行义务并遵循诚信原则；"
              "变更涉及工程量、工期、价款的，应按约定程序确认并相应顺延、调整",
        suggestion="删除「随时/单方/无需同意」表述，变更须经书面签证程序确认，"
                   "并同步约定工期顺延与费用、价款调整机制",
    ),
    ClauseRule(
        id="C26", name="不可竞争费用让利",
        any_kw=("安全文明施工费让利", "安全文明施工费下浮", "规费让利", "规费下浮",
                "税金下浮", "不可竞争费用让利", "不可竞争费让利"),
        risk_type="计价合规风险",
        level="较大风险",
        basis="《建设工程工程量清单计价规范》（GB 50500-2013）第 3.1.4 条、第 3.1.5 条："
              "安全文明施工费、规费和税金为不可竞争费用，不得作为竞争性让利内容",
        suggestion="删除让利/下浮表述；安全文明施工费按规定足额计取、专款专用，"
                   "规费与税金按法定标准计取",
    ),
    ClauseRule(
        id="C27", name="背离中标合同实质性内容",
        any_kw=("补充协议不一致以补充协议为准", "以补充协议为准", "黑白合同", "阴阳合同",
                "补充协议优先于", "补充协议效力优先"),
        risk_type="结算依据风险",
        level="重大风险",
        basis="《中华人民共和国招标投标法》第四十六条；最高人民法院"
              "《关于审理建设工程施工合同纠纷案件适用法律问题的解释（一）》第二条："
              "招标工程另行订立的背离中标合同实质性内容的协议，不得作为结算依据",
        suggestion="补充协议不得背离中标合同的工期、价款、质量等实质性内容；"
                   "确需调整的依法定程序办理，结算依据以中标合同为准",
    ),
]

_PCT = re.compile(r"(\d+(?:\.\d+)?)\s*[%％]")
_YEAR = re.compile(r"(\d+(?:\.\d+)?)\s*年")

_LEVEL_RANK = {"一般风险": 1, "较大风险": 2, "重大风险": 3}


def _extract_pcts(text: str) -> List[float]:
    return [float(x) for x in _PCT.findall(text or "")]


def _extract_years(text: str) -> List[float]:
    return [float(x) for x in _YEAR.findall(text or "")]


# ---------------------------------------------------------------- 审查

def review_clauses(clauses: List[dict]) -> List[dict]:
    """对条款逐条跑规则库，返回条款级发现列表（确定性：条款顺序 × 规则顺序）。"""
    findings: List[dict] = []
    for c in clauses or []:
        text = c.get("text") or ""
        if not text:
            continue
        for r in RULES:
            if not any(k in text for k in r.any_kw):
                continue
            # 豁免词：条文明确禁止性表述（如"不得转包"）不构成风险
            if r.exclude_kw and any(k in text for k in r.exclude_kw):
                continue
            triggered = False
            if r.force_kw and any(k in text for k in r.force_kw):
                triggered = True
            if not triggered and r.max_pct is not None:
                pcts = _extract_pcts(text)
                if any(p > r.max_pct for p in pcts):
                    triggered = True
            if not triggered and r.max_year is not None:
                yrs = _extract_years(text)
                if any(y > r.max_year for y in yrs):
                    triggered = True
            if not (r.force_kw or r.max_pct is not None or r.max_year is not None):
                triggered = True
            if not triggered:
                continue
            level = r.level
            if r.high_kw and any(k in text for k in r.high_kw):
                level = r.level_high
            findings.append({
                "clause_id": c.get("id", ""),
                "clause_title": c.get("title", ""),
                "rule_id": r.id,
                "rule_name": r.name,
                "risk_type": r.risk_type,
                "level": level,
                "basis": r.basis,
                "suggestion": r.suggestion,
                "evidence": [f"条款文本命中「{'/'.join(r.any_kw[:3])}」"],
            })
    return findings


def summarize(findings: List[dict]) -> dict:
    """把条款级发现聚合成审查概览（按等级计数 + 按风险类型归并最高等级）。"""
    by_level = {"一般风险": 0, "较大风险": 0, "重大风险": 0}
    by_type: dict = {}
    for f in findings:
        lv = f.get("level", "一般风险")
        by_level[lv] = by_level.get(lv, 0) + 1
        rt = f.get("risk_type", "其他")
        cur = by_type.get(rt)
        if cur is None or _LEVEL_RANK.get(lv, 0) > _LEVEL_RANK.get(cur, 0):
            by_type[rt] = lv
    top = max(by_level, key=lambda k: (by_level[k], _LEVEL_RANK[k])) if findings else "无风险"
    return {
        "finding_count": len(findings),
        "by_level": by_level,
        "by_risk_type": by_type,
        "overall": top if findings else "未发现明显风险条款",
    }


def review_contract(text: str) -> dict:
    """便捷入口：切分 + 审查 + 概览。返回 {"clauses","findings","summary","meta"}。"""
    clauses = split_clauses(text)
    findings = review_clauses(clauses)
    summary = summarize(findings)
    summary["clause_count"] = len(clauses)
    return {
        "clauses": clauses,
        "findings": findings,
        "summary": summary,
        "meta": {
            "engine": ENGINE_NAME,
            "engine_version": ENGINE_VERSION,
            "rules_version": RULES_VERSION,
            "rule_count": len(RULES),
        },
    }


# ---------------------------------------------------------------- v1.3.0 Stage 2: LLM 复核合并

# verdict 取值（与 references/architecture.md / stage2_verdict_schema.json 保持一致）：
#   - "confirmed"      — LLM 复核确认是真问题（原 finding 保留）
#   - "false_positive" — LLM 判定为误报（原 finding 移除）
#   - "uncertain"      — LLM 存疑（原 finding 保留，加【复核存疑】标签）
#   - "missed"         — LLM 判定规则引擎未输出但合同确有此问题（补报到 findings）
#   - "correctly_absent" — LLM 复核确认无问题（仅用于审计日志，不进 findings）
_VALID_VERDICTS = {"confirmed", "false_positive", "uncertain", "missed", "correctly_absent"}


def _merge_stage2(result: dict, verdicts_path: str) -> dict:
    """读入平台 agent 调 LLM 后产出的 verdicts JSON，与 Stage 1 findings 合并。

    verdicts JSON 格式（由 SKILL.md + references/architecture.md 定义，由平台 agent 写入）：
        {
          "engine_version": "1.3.0",
          "input_path": "<原合同路径>",
          "stage1_finding_count": N,
          "verdicts": [
            {
              "verdict": "confirmed" | "false_positive" | "uncertain"
                         | "missed" | "correctly_absent",
              "rule_id": "C07",            # confirmed/false_positive/uncertain 对应原 finding 的 rule_id
              "clause_id": "第1条",         # missed 必填；其他可选
              "clause_title": "工期与进度管理",  # missed 必填
              "clause_text": "...",        # missed 必填
              "reason": "LLM 复核理由..."   # 必填：false_positive / missed / uncertain 必填理由
            },
            ...
          ]
        }

    合并规则（确定性）：
      - verdict=confirmed → 原 finding 保留，加 verdict 字段
      - verdict=false_positive → 原 finding 从 findings 中移除（保留在 suppressed_by_stage2 中供审计）
      - verdict=uncertain → 原 finding 保留，加【复核存疑】前缀
      - verdict=missed → 新增 finding，按 (clause_id, rule_id) 唯一；引用 RULES 中规则定义
      - verdict=correctly_absent → 不动 findings（仅审计日志记录：LLM 复核确认无问题）

    fail-safe：
      - 文件不存在 → 报错退出（exit code 2）
      - JSON 解析失败 → 报错退出（exit code 2）
      - verdict 值非法 → 报错退出（exit code 2）
      - 缺失 rule_id/clause_id/clause_title/clause_text/reason → 报错退出（exit code 2）
    """
    import json as _json
    import sys

    if not os.path.exists(verdicts_path):
        sys.stderr.write(
            f"[stage2] verdicts 文件不存在: {verdicts_path}\n"
            f"[stage2] LLM 复核必走模式（v1.3.0）：未传 verdicts 应报错退出。\n"
            f"[stage2] 如不要 LLM 复核，请不要传 --stage2-input（默认 v1.2.2 行为）。\n"
        )
        sys.exit(2)

    try:
        with open(verdicts_path, "r", encoding="utf-8") as fh:
            verdicts_doc = _json.load(fh)
    except _json.JSONDecodeError as e:
        sys.stderr.write(f"[stage2] verdicts JSON 解析失败: {e}\n")
        sys.exit(2)

    verdicts = verdicts_doc.get("verdicts", [])
    if not isinstance(verdicts, list):
        sys.stderr.write(f"[stage2] verdicts 必须是 list，当前类型: {type(verdicts).__name__}\n")
        sys.exit(2)

    # 校验每条 verdict
    for i, v in enumerate(verdicts):
        if not isinstance(v, dict):
            sys.stderr.write(f"[stage2] verdicts[{i}] 不是 dict\n")
            sys.exit(2)
        verdict = v.get("verdict")
        if verdict not in _VALID_VERDICTS:
            sys.stderr.write(
                f"[stage2] verdicts[{i}].verdict 非法值: {verdict!r}，"
                f"必须是 {_VALID_VERDICTS} 之一\n"
            )
            sys.exit(2)
        # missed 必须带 clause_id/title/text/reason
        if verdict == "missed":
            for k in ("clause_id", "clause_title", "clause_text", "reason"):
                if not v.get(k):
                    sys.stderr.write(
                        f"[stage2] verdicts[{i}].verdict=missed 必填字段缺失: {k}\n"
                    )
                    sys.exit(2)
            if not v.get("rule_id"):
                sys.stderr.write(
                    f"[stage2] verdicts[{i}].verdict=missed 必填字段缺失: rule_id\n"
                    f"[stage2] LLM 只能补报已存在规则，不能发明新规则\n"
                )
                sys.exit(2)
        # false_positive / uncertain 必须带 reason
        if verdict in ("false_positive", "uncertain"):
            if not v.get("reason"):
                sys.stderr.write(
                    f"[stage2] verdicts[{i}].verdict={verdict} 必填 reason\n"
                )
                sys.exit(2)
        # 校验 rule_id 是否在 27 条规则内
        rule_id = v.get("rule_id")
        if rule_id is not None:
            known_rule_ids = {r.id for r in RULES}
            if rule_id not in known_rule_ids:
                sys.stderr.write(
                    f"[stage2] verdicts[{i}].rule_id={rule_id!r} 不在 27 条规则库内\n"
                    f"[stage2] LLM 只能复核已有规则，不能发明新规则\n"
                )
                sys.exit(2)

    # ---- 合并 ----
    findings = list(result.get("findings", []))
    suppressed: List[dict] = []  # 误报清单（审计用）
    appended_missed: List[dict] = []  # 补报清单
    review_log: List[dict] = []  # 全部复核记录

    # 索引原始 findings：(rule_id, clause_id) → finding
    raw_index: dict = {}
    for f in findings:
        key = (f.get("rule_id", ""), f.get("clause_id", ""))
        raw_index[key] = f

    # 处理每条 verdict
    consumed_keys: set = set()
    for v in verdicts:
        verdict = v["verdict"]
        rule_id = v.get("rule_id", "")
        clause_id = v.get("clause_id", "")
        reason = v.get("reason", "")

        if verdict == "correctly_absent":
            review_log.append({"verdict": verdict, "rule_id": rule_id,
                               "clause_id": clause_id, "reason": reason})
            continue

        if verdict in ("confirmed", "false_positive", "uncertain"):
            key = (rule_id, clause_id)
            f = raw_index.get(key)
            if f is None:
                sys.stderr.write(
                    f"[stage2] verdicts 引用了 Stage 1 没输出的 (rule_id={rule_id}, "
                    f"clause_id={clause_id})，conflicted verdict 无对应原 finding\n"
                )
                sys.exit(2)
            consumed_keys.add(key)
            # 复制后修改原 findings 里的对象（保证后续 false_positive 过滤生效）
            f = dict(f)
            if verdict == "confirmed":
                f["stage2_verdict"] = verdict
                review_log.append({"verdict": verdict, "rule_id": rule_id,
                                   "clause_id": clause_id, "reason": reason})
            elif verdict == "false_positive":
                f["stage2_verdict"] = verdict
                f["stage2_reason"] = reason
                suppressed.append(f)
                review_log.append({"verdict": verdict, "rule_id": rule_id,
                                   "clause_id": clause_id, "reason": reason})
            elif verdict == "uncertain":
                f["stage2_verdict"] = verdict
                f["stage2_reason"] = reason
                review_log.append({"verdict": verdict, "rule_id": rule_id,
                                   "clause_id": clause_id, "reason": reason})
            # 把 modified f 替换回 findings（必须，否则后续筛选不起作用）
            for i, orig in enumerate(findings):
                if (orig.get("rule_id") == rule_id
                        and orig.get("clause_id") == clause_id):
                    findings[i] = f
                    break

        elif verdict == "missed":
            # 补报：必须在 RULES 内、必须带 clause 信息
            rule_def = next((r for r in RULES if r.id == rule_id), None)
            if rule_def is None:
                # 已在前面校验过，这里是兜底
                sys.stderr.write(f"[stage2] rule_id={rule_id} 不在 RULES 内\n")
                sys.exit(2)
            new_finding = {
                "clause_id": clause_id,
                "clause_title": v["clause_title"],
                "rule_id": rule_id,
                "rule_name": rule_def.name,
                "risk_type": rule_def.risk_type,
                "level": rule_def.level,
                "basis": rule_def.basis,
                "suggestion": rule_def.suggestion,
                "evidence": [f"Stage 2 LLM 复核补报：{reason}"],
                "stage2_verdict": "missed",
                "stage2_reason": reason,
            }
            findings.append(new_finding)
            appended_missed.append(new_finding)
            review_log.append({"verdict": verdict, "rule_id": rule_id,
                               "clause_id": clause_id, "reason": reason})

    # ---- 移除 false_positive ----
    final_findings = []
    for f in findings:
        if f.get("stage2_verdict") == "false_positive":
            continue  # 已在 suppressed 中记录
        final_findings.append(f)

    # ---- 检查：所有 Stage 1 findings 都必须有对应 verdict ----
    missing = []
    for f in result.get("findings", []):
        key = (f.get("rule_id", ""), f.get("clause_id", ""))
        if key not in consumed_keys:
            missing.append({"rule_id": key[0], "clause_id": key[1]})
    if missing:
        sys.stderr.write(
            f"[stage2] Stage 1 有 {len(missing)} 条 finding 未在 verdicts 中复核：\n"
        )
        for m in missing:
            sys.stderr.write(f"  - rule_id={m['rule_id']}, clause_id={m['clause_id']}\n")
        sys.stderr.write(
            "[stage2] verdicts 必须覆盖 Stage 1 全部 findings（confirmed/false_positive/uncertain 三选一）\n"
        )
        sys.exit(2)

    # ---- 重算 summary ----
    summary = summarize(final_findings)
    summary["clause_count"] = len(result.get("clauses", []))
    summary["stage2"] = {
        "suppressed_count": len(suppressed),
        "missed_count": len(appended_missed),
        "review_log_count": len(review_log),
    }

    return {
        "clauses": result.get("clauses", []),
        "findings": final_findings,
        "suppressed_by_stage2": suppressed,
        "appended_by_stage2": appended_missed,
        "stage2_review_log": review_log,
        "summary": summary,
        "meta": result.get("meta", {}),
    }


# ---------------------------------------------------------------- docx 文本提取（纯标准库）

_ENTITIES = {"&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&apos;": "'"}
_RE_WT = re.compile(r"<w:t[^>]*>([^<]*)</w:t>")


def _unescape(s: str) -> str:
    for k, v in _ENTITIES.items():
        s = s.replace(k, v)
    return s


def extract_docx_text(path: str) -> str:
    """从 .docx（OOXML zip）提取纯文本，按段落换行。确定性、纯标准库。"""
    import zipfile
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    lines: List[str] = []
    for p in xml.split("</w:p>"):
        parts = _RE_WT.findall(p)
        if parts:
            lines.append(_unescape("".join(parts)))
    return "\n".join(lines)


# ---------------------------------------------------------------- Markdown 报告（固定格式）

_LEVEL_ORDER = {"重大风险": 0, "较大风险": 1, "一般风险": 2}


def _strip_clause_id_prefix(clause_id: str, clause_title: str) -> str:
    """去掉 clause_title 中与 clause_id 重复的前缀，使输出不出现"第21条　第21条 工程款支付节点"这样的重复。
    规则：若 clause_title 以 clause_id 开头（允许中间有空格/冒号等分隔），则去掉该前缀后返回；否则原样返回 clause_title。
    """
    if not clause_id or not clause_title:
        return clause_title
    # clause_id 形如"第21条"，clause_title 形如"第21条 工程款支付节点"
    # 去掉 clause_title 开头与 clause_id 相同的部分（忽略大小写/全角半角）
    if clause_title.startswith(clause_id):
        rest = clause_title[len(clause_id):]
        # 去掉开头的空白字符（可能有空格、冒号、全角空格等）
        rest = rest.lstrip(" \t　:：")
        return rest if rest else clause_title
    return clause_title


def render_markdown(result: dict, title: str = "建筑合同审查报告") -> str:
    """把审查结果渲染为固定格式报告（确定性：先按等级、再按规则号稳定排序）。"""
    findings = sorted(
        result.get("findings", []),
        key=lambda f: (_LEVEL_ORDER.get(f.get("level", "一般风险"), 3),
                       f.get("rule_id", "")),
    )
    s = result.get("summary", {})
    bl = s.get("by_level", {})
    lines: List[str] = []
    ap = lines.append
    ap(f"# {title}")
    ap("")
    ap(f"> 引擎版本：{ENGINE_NAME} v{ENGINE_VERSION}　|　规则版本：{RULES_VERSION}")
    ap("")
    ap("## 一、审查概览")
    ap("")
    ap(f"- 条款总数：{s.get('clause_count', 0)}")
    ap(f"- 风险发现总数：{s.get('finding_count', 0)}")
    ap(f"- 重大风险：{bl.get('重大风险', 0)}　较大风险：{bl.get('较大风险', 0)}　"
       f"一般风险：{bl.get('一般风险', 0)}")
    ap(f"- 总体判断：**{s.get('overall', '未发现明显风险条款')}**")
    by_type = s.get("by_risk_type", {})
    if by_type:
        ap("")
        ap("### 风险类型分布")
        ap("")
        ap("| 风险类型 | 最高等级 |")
        ap("| --- | --- |")
        for rt in sorted(by_type):
            ap(f"| {rt} | {by_type[rt]} |")
    ap("")
    ap("## 二、风险明细（按严重程度排列）")
    if not findings:
        ap("")
        ap("未发现命中规则库的风险条款。")
    for i, f in enumerate(findings, 1):
        ap("")
        ap(f"### {i}. [{f.get('level')}] {f.get('rule_name')}（{f.get('rule_id')}）")
        ap("")
        # 所在条款：条款号 + 去重后的条款名称，避免"第21条　第21条 工程款支付节点"
        clause_id = f.get("clause_id", "")
        clause_title = f.get("clause_title", "")
        display_title = _strip_clause_id_prefix(clause_id, clause_title)
        ap(f"- **所在条款**：{clause_id}　{display_title}")
        ap(f"- **风险类型**：{f.get('risk_type')}")
        ap(f"- **法规依据**：{f.get('basis')}")
        ap(f"- **修订建议**：{f.get('suggestion')}")
        ev = "; ".join(f.get("evidence", []))
        if ev:
            ap(f"- **判定依据**：{ev}")
    ap("")
    ap("## 三、说明")
    ap("")
    ap("- 本报告由确定性规则引擎生成：同一份合同，结论恒定不变。")
    ap("- 审查范围为规则库覆盖的风险点，不构成对合同全部法律问题的判断；"
       "重大事项建议咨询执业律师。")
    return "\n".join(lines)


# ---------------------------------------------------------------- 命令行入口

def _read_input(argv: List[str]) -> str:
    path = argv[0]
    if path == "-":
        import sys
        return sys.stdin.read()
    if path.lower().endswith(".docx"):
        return extract_docx_text(path)
    with open(path, "r", encoding="utf-8", errors="ignore") as fh:
        return fh.read()


def main(argv: List[str]) -> None:
    import argparse
    import json
    import sys

    parser = argparse.ArgumentParser(description="建筑合同确定性审查引擎（v1.3.0：可选 LLM 复核）")
    parser.add_argument("input", help="合同文件路径（.txt/.docx），或 - 读标准输入")
    parser.add_argument("--format", choices=["json", "md"], default="json",
                        help="输出格式：json（默认）或 md（固定格式报告）")
    parser.add_argument("--title", default="建筑合同审查报告", help="md 报告标题")
    # v1.3.0：Stage 1 only — 只跑规则引擎，结果写到 --stage1-output（默认 stdout）。
    # 输出包含 findings_raw 数组（Stage 1 原始发现）+ clauses（条款全文），
    # 由调用方（平台 agent）读完后再调 LLM 复核，写入 verdicts.json 后用 --stage2-input 合并。
    parser.add_argument("--stage1-only", action="store_true",
                        help="只跑 Stage 1：输出含 findings_raw + clauses 的 JSON，"
                             "供 LLM 复核介入（不输出 md 报告）")
    parser.add_argument("--stage1-output", default=None,
                        help="Stage 1 JSON 写到指定文件（与 --stage1-only 配合）")
    # v1.3.0：Stage 2 input — 读入 LLM 复核 verdicts（由平台 agent 调用 LLM 后产出），
    # 与 Stage 1 原始 findings 合并：confirmed 保留、false_positive 删除、uncertain 保留+标签、
    # missed 补报到 findings。
    parser.add_argument("--stage2-input", default=None,
                        help="LLM 复核 verdicts JSON 路径（由平台 agent 调 LLM 后产出）；"
                             "合并后输出最终报告")
    # v1.3.0：no-stage2 — 显式声明不启用 Stage 2（默认行为即向后兼容 v1.2.2）。
    # 仅保留用于意图清晰。
    parser.add_argument("--no-stage2", action="store_true",
                        help="显式声明不启用 Stage 2（默认即不启用）")
    args = parser.parse_args(argv)

    text = _read_input([args.input])
    result = review_contract(text)

    # ---- Stage 1 only ----
    if args.stage1_only:
        out = {
            "engine": ENGINE_NAME,
            "engine_version": ENGINE_VERSION,
            "rules_version": RULES_VERSION,
            "stage": "stage1",
            "input_path": args.input,
            "findings_raw": result.get("findings", []),
            "clauses": [
                {"id": c["id"], "title": c["title"], "text": c["text"]}
                for c in result.get("clauses", [])
            ],
            "summary_raw": result.get("summary", {}),
        }
        serialized = json.dumps(out, ensure_ascii=False, sort_keys=True, indent=2)
        if args.stage1_output:
            with open(args.stage1_output, "w", encoding="utf-8") as fh:
                fh.write(serialized + "\n")
        else:
            sys.stdout.write(serialized + "\n")
        return

    # ---- Stage 2 input：合并 LLM 复核 verdicts ----
    if args.stage2_input:
        result = _merge_stage2(result, args.stage2_input)

    # ---- 默认输出（向后兼容 v1.2.2 行为） ----
    if args.format == "md":
        sys.stdout.write(render_markdown(result, args.title) + "\n")
    else:
        out = dict(result)
        out["clauses"] = [
            {"id": c["id"], "title": c["title"], "text": c["text"]} for c in out["clauses"]
        ]
        sys.stdout.write(json.dumps(
            out, ensure_ascii=False, sort_keys=True, indent=2) + "\n")


if __name__ == "__main__":
    import sys
    main(sys.argv[1:])