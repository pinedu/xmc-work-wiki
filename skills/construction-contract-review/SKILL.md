---
name: construction-contract-review
version: 1.3.3
description: "建筑领域施工/分包合同条款风险审查。Stage 1：基于 27 条确定性规则引擎（背靠背付款、质保金超标、无调差机制、放弃优先受偿权、垫资、违法转包、挂靠、农民工工资转嫁、不可竞争费让利等）输出条款级风险发现。Stage 2（v1.3.0 新增）：由平台 agent 调 LLM 复核 Stage 1 输出，发现规则引擎的误报与漏报（限 27 条规则内，不发明新规则）。LLM 调用不写进 skill——各平台 agent 自行接入。v1.3.1 收紧 LLM 端交付边界：LLM 复核只产出 verdicts.json，不在对话里另写自由格式总结/解读/建议段；最终交付物=引擎 render_markdown() 输出。v1.3.2 调整交付与清理：审计结果（引擎 --format md 输出）直接输出到聊天面板交付，不专门生成报告文件；stage1.json、verdicts.json 等中间文件一律放系统临时目录，审计输出后立即删除。v1.3.3 新增强制 Stage 2 复核操作规程（第五节）：两遍复核法——第一遍对 Stage 1 全部 findings 逐条读完整条款纠误报，第二遍对 27 条规则全量防漏报（不允许跳过任何一条），复核理由全量留痕进 verdicts。本 skill 应当在用户提供合同文本或合同文件（.txt/.docx）并要求审查合同风险、审合同、找合同风险点时使用。核心保证：Stage 1 同一份合同无论审查多少次，风险结论完全一致（确定性引擎，不依赖模型判断）。★ 子条款精度（v1.1.0）：引擎对「第X条」正文内含「X.1 / X.2 / X.3」编号的合同，可精确将风险定位至「第X.Y条」粒度（如第22.2条），而非整条（第22条）。"
agent_created: true
---

# 建筑合同审查（v1.3.3：Stage 1 确定性 + Stage 2 强制复核规程 + 聊天直出 + 中间文件清理）

## 一、触发场景

用户说"审合同 / 合同审查 / 合同风险 / 帮我审一下" → 触发本 skill。

## 二、两阶段架构（v1.3.0 起）

```
用户合同
   ↓
[Stage 1] review_engine.py 跑规则引擎（确定性）
   输出 stage1.json = findings_raw + clauses
   ↓
[平台 agent] 读 stage1.json + 合同原文，调本平台 LLM
   产出 verdicts.json（按 assets/stage2_verdict_schema.json）
   ↓
[Stage 2] review_engine.py 读 verdicts.json 合并出最终报告
   渲染遵循 references/report_template.md 固定结构（model 不参与排版）
   ↓
[交付] 引擎 --format md 的原样输出 = 直接打印到聊天面板（不生成报告文件）
   ↓
[清理] 删除 stage1.json / verdicts.json 等全部中间文件
```

### 关键边界

- **Stage 1 确定性**——同一份合同任何时候跑，输出字节相同
- **Stage 2 由平台 agent 实现**——skill 不写 LLM 调用代码
- **LLM 只能复核 27 条规则内**——不能发明新风险点
- **fail-safe**：verdicts 违规 → exit 2，禁止输出报告
- **交付边界**：给用户看的报告**只能是引擎 `render_markdown()` 的输出**，LLM 不再额外写任何总结/解读/建议段。LLM 的推理只体现在 verdicts.json 的 `reason` 字段里，由引擎在 `suppressed_by_stage2` / `uncertain_notes` 等审计区块中按固定措辞引用。LLM 若在 verdicts 之外另写一段自由格式回复，视为违反本 skill 边界。
- **不要混用上下文**：LLM 复核时**禁止引用会话上下文**（其他项目、其他案件、用户偏好），只以本合同的条款文本 + `references/rules.md` 为输入。
- **聊天直出（v1.3.2）**：审计结果以引擎 `--format md` 输出**直接打印在聊天面板**交付，不保存为报告文件、不调用文件展示功能。
- **中间文件即用即删（v1.3.2）**：stage1.json / verdicts.json 一律写到系统临时目录，审计输出完成后**立即删除**；任何中间产物不得留在工作区或交付给用户。
- **复核规程强制（v1.3.3）**：Stage 2 复核必须按**第五节《Stage 2 复核操作规程》**执行——两遍复核法（纠误报 + 防漏报全量覆盖），27 条规则每条必须有落点结论，理由全量留痕；未按规程复核视为本次审查无效，不得输出报告。

## 三、典型调用

### 3.1 不启用 LLM 复核（默认 = v1.2.2 行为）

```bash
python review_engine.py contract.docx --format md
```

输出与 v1.2.2 字节相同。

### 3.2 启用 LLM 复核

```bash
# Step 1：跑 Stage 1
python review_engine.py contract.docx --stage1-only --stage1-output stage1.json

# Step 2：（平台 agent 调 LLM 产出 verdicts.json）
# ... agent 自行实现 ...

# Step 3：Stage 2 合并
python review_engine.py contract.docx --stage2-input verdicts.json --format md
```

### 3.3 交付与清理（v1.3.2）

1. stage1.json、verdicts.json 等中间文件**一律写到系统临时目录**（不进工作区）；
2. Stage 2 合并拿到 `--format md` 输出后，**原样打印到聊天面板**交付；
3. 交付后**立即删除**全部中间文件（stage1.json / verdicts.json）；
4. 不向用户交付任何 .json / 报告文件；聊天回复中除引擎报告原文外，不得另写自由格式总结/解读/建议段。

## 四、verdict 5 个值

| verdict | 含义 | 对原 finding 的动作 |
|---|---|---|
| `confirmed` | LLM 复核确认是真问题 | 保留 |
| `false_positive` | LLM 判定为误报 | 移除（移入 suppressed_by_stage2） |
| `uncertain` | LLM 存疑 | 保留 + 标签 |
| `missed` | LLM 判定规则引擎漏报（合同里有但 Stage 1 没报） | 补报到 findings |
| `correctly_absent` | LLM 复核确认无问题 | 不动（仅审计） |

详细 schema 见 `assets/stage2_verdict_schema.json`。

## 五、Stage 2 复核操作规程（v1.3.3 起，强制）

平台 agent 执行 LLM 复核时必须遵守本节全部要求。违反任何一条视为本次审查未按 skill 执行，不得输出报告。

### 5.1 输入与隔离

- 复核输入**仅限三样**：Stage 1 的 `stage1.json`（findings_raw + clauses）、合同原文、`references/rules.md`；
- **禁止引用会话上下文**（其他项目、其他合同、用户偏好、历史审查结论）——与「关键边界」条款一致，此处为执行细则；
- 中间文件一律写系统临时目录；「Stage 1 → 产出 verdicts → Stage 2 合并 → 清理」须在**同一次执行流程内串行完成**（部分平台沙箱会在多次工具调用间清理临时目录，跨调用存放会丢文件）。

### 5.2 两遍复核法（强制顺序，缺一不可）

**第一遍：纠误报（对象 = Stage 1 findings_raw，逐条、全量）**

1. 对每条 finding，**读完整条款文本**（整条，必要时相邻条款），不得只看命中关键词所在句子；
2. 按以下判据下 verdict：
   - 条款前文或同条其他款项已实质**保留权利 / 设有兜底**的（如「因发包人原因、设计变更或不可抗力……工期相应顺延」在前、「工期不予顺延」在后）→ `false_positive`，理由必须**引用保留权利的原文**；
   - 命中且上下文无减免 → `confirmed`；
   - 风险是否成立**取决于合同文本之外的事实**（如是否属依法必须招标项目、是否经过招投标程序）→ 一律 `uncertain`，理由写明「取决于什么事实、如何核实」，**不得武断 confirm 或否掉**。

**第二遍：防漏报（对象 = 27 条规则，逐条、全量，不允许跳过任何一条）**

1. 把 `references/rules.md` 的 27 条规则当作**检查清单**，每一条都在合同全文中寻找事实对应（如：C01 找付款前提安排、C03 找质保金比例数字并与 3% 法定上限比对、C14 找垫资安排、C22 找工资支付责任表述）；
2. 每条规则必须落一个结论：合同中存在该规则情形但 Stage 1 未报 → `missed`（按 schema 附 clause_id / clause_title / clause_text / reason）；不存在 → `correctly_absent`；
3. 规则触发词看似出现但属**合规表述**的（如 28 天失权条款带「发包人原因导致未能按期主张的不在此限」兜底、质保金恰好 3%），落 `correctly_absent`，理由中写明排除依据。

### 5.3 理由与留痕（强制）

- `false_positive` / `uncertain` / `missed` 理由必填（schema 强制）；
- `correctly_absent` **也必须写理由**：说明该规则为何不适用、或被何种合规表述排除——**全部 27 条的核对理由都要进 verdicts 留痕**，不允许空过、不允许合并笼统带过；
- 理由须引用合同原文措辞或写明核对对象条款号；不接受「无此风险」「未发现」式空泛理由。

### 5.4 产出与校验

- verdicts.json 严格按 `assets/stage2_verdict_schema.json`：`engine_version` 取自 stage1.json，`stage1_finding_count` 与 findings_raw 长度一致；
- 提交 Stage 2 合并前**自查两项**：Stage 1 每条 finding 恰有一个对应 verdict；全部 rule_id 在 27 条枚举内；
- 校验失败（exit 2）= fail-safe：禁止输出报告，修正 verdicts 后重跑，不得降级为「不带复核直接出报告」。

## 六、典型场景示例

### 场景 A：规则引擎误报，LLM 复核纠正

合同第1条："因发包人原因、设计变更或不可抗力……工期相应顺延；因承包人自身原因……工期不予顺延。"

- Stage 1：报 C07 工期索赔权被排除（重大风险）
- LLM 复核：前半句已保留顺延权 → `verdict: false_positive`
- Stage 2 合并：移除 C07，最终报告 0 条发现

### 场景 B：规则引擎漏报，LLM 复核补报

合同某条："乙方放弃建设工程价款优先受偿权。"

- Stage 1 v1.2.2：可能漏报（如果其他规则已写 keep_all）
- LLM 复核：合同确实写了 → `verdict: missed, rule_id: C13`
- Stage 2 合并：补报 C13 到 findings

## 七、文件清单

- `scripts/review_engine.py`：Stage 1 + Stage 2 合并
- `scripts/self_check.py`：Stage 1 基线测试
- `assets/sample_contract.txt`：自测样例
- `assets/stage2_verdict_schema.json`：verdicts 格式契约
- `references/rules.md`：27 条规则库
- `references/architecture.md`：架构、数据契约、跨平台迁移清单

## 八、跨平台迁移

复制整个 skill 目录到新平台即可。**LLM 调用不写进 skill**——各平台 agent 自行实现"读 Stage 1 JSON → 调 LLM → 写 verdicts.json"。

详见 `references/architecture.md` 第四节。

## 九、扩充规则库的流程

1. 读 `references/rules.md`，确认新规则不与现有 27 条重复；
2. 在 `review_engine.py` 的 `RULES` 末尾追加（id 接续编号，法规全称+条号，明确 any_kw，必要时配 exclude_kw 防合规表述误报）；
3. 更新 `RULES_VERSION`（条数/日期）与 `references/rules.md`；
4. 若新规则应被样例命中，更新样例合同与 `self_check.py` 的 `GOLDEN_RULE_IDS`；
5. 跑 `self_check.py`，全过为止；
6. ⚠️ 新增规则同时更新 `assets/stage2_verdict_schema.json` 的 `rule_id` 枚举值。