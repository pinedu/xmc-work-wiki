# v1.3.0 架构说明：Stage 1 确定性 + Stage 2 LLM 复核

> **目的**：解释 skill 内部 Stage 1 / Stage 2 的边界、数据契约、跨平台迁移清单。

---

## 一、为什么需要 Stage 2？

### Stage 1 的固有限制

Stage 1（Python `review_engine.py`）用关键词匹配判定风险：

- ✅ 同一份合同任何时候审都一致（**确定性 / 零漂移**）
- ✅ 离线运行（不依赖 LLM API）
- ❌ **上下文盲**：只看关键词是否出现，不看上下文是否已保留权利
- ❌ **同义不同字**：条款用别的措辞表达同义权利时，引擎不知道

**实际案例**（HCJY-ZCB-2026-018 第1条）：

```
合同原文："因发包人原因、设计变更或不可抗力导致工期延误的……工期相应顺延；
          因承包人自身原因导致工期延误的，工期不予顺延。"
```

这是**完整对等条款**，但 Stage 1 v1.2.2 之前会报 C07（重大风险），因为它只看到"工期不予顺延"五个字就触发，没读前半句"工期相应顺延"。

### Stage 2 的作用

**LLM（由各平台 agent 调用）读 Stage 1 JSON + 合同全文**：

| Stage 1 报的问题 | LLM 复核 | Stage 2 输出 |
|---|---|---|
| 关键词命中但合同已保留权利 | **false_positive** | 移除，移入 `suppressed_by_stage2` |
| 关键词命中，LLM 确认是真问题 | **confirmed** | 保留 |
| 关键词命中，LLM 拿不准 | **uncertain** | 保留 + 标签 |
| 合同里写了某规则但 Stage 1 没报 | **missed** | 补报到 findings |
| Stage 1 没输出，LLM 复核确认没漏 | **correctly_absent** | 不动（仅审计） |

**核心边界**：LLM **只能在 27 条规则范围内**打 verdict，不能发明新风险点。

---

## 二、数据契约

### Stage 1 输出（`--stage1-only`）

平台 agent 调用：

```bash
python review_engine.py <合同> --stage1-only --stage1-output stage1.json
```

`stage1.json` 结构：

```json
{
  "engine": "construction-contract-review",
  "engine_version": "1.3.0",
  "rules_version": "...",
  "stage": "stage1",
  "input_path": "...",
  "findings_raw": [
    {
      "clause_id": "第1条",
      "clause_title": "工期与进度管理",
      "rule_id": "C07",
      "rule_name": "工期索赔权被排除",
      "risk_type": "工期风险",
      "level": "重大风险",
      "basis": "...",
      "suggestion": "..."
    }
  ],
  "clauses": [
    {"id": "第1条", "title": "...", "text": "..."}
  ],
  "summary_raw": {...}
}
```

### Stage 2 输入（平台 agent 产出 verdicts.json）

平台 agent 读完 `stage1.json` + 合同原文后，调平台 LLM 产出 verdicts.json：

```json
{
  "engine_version": "1.3.0",
  "input_path": "...",
  "stage1_finding_count": 1,
  "verdicts": [
    {
      "verdict": "false_positive",
      "rule_id": "C07",
      "clause_id": "第1条",
      "reason": "合同前半句已明确'因发包人原因、设计变更或不可抗力……工期相应顺延'，保留顺延权，不构成排除主要权利。"
    }
  ]
}
```

verdict 完整定义见 `assets/stage2_verdict_schema.json`。

### Stage 2 合并输出（`--stage2-input verdicts.json`）

```bash
python review_engine.py <合同> --stage2-input verdicts.json --format md
```

输出含：
- `findings` —— 经 Stage 2 修正后的最终 findings
- `suppressed_by_stage2` —— 被 LLM 判为误报的清单（审计用）
- `appended_by_stage2` —— LLM 补报的清单
- `stage2_review_log` —— 全部复核记录
- `summary` —— 含 `stage2.suppressed_count / missed_count / review_log_count`

---

## 三、verdict 校验规则（review_engine.py 强制）

平台 agent 写 verdicts.json 时必须遵守：

| 规则 | 说明 |
|---|---|
| verdict 必须是 5 个枚举值之一 | `confirmed / false_positive / uncertain / missed / correctly_absent` |
| rule_id 必须在 27 条规则内 | LLM 不能发明新规则 |
| `false_positive / uncertain` 必填 `reason` | 必须解释为什么 |
| `missed` 必填 `clause_id / clause_title / clause_text / reason` | 用于复审 |
| `confirmed / false_positive / uncertain` 必须对应 Stage 1 真实 findings | 否则 exit 2 |
| **Stage 1 全部 findings 必须有对应 verdict** | 否则 exit 2（防止 LLM 漏复核） |
| 失败处理 | exit 2（**fail-safe：禁止输出**） |

---

## 四、跨平台迁移清单

把 skill 复制到新平台时：

| 文件 | 必带 | 说明 |
|---|---|---|
| `SKILL.md` | ✅ | 主入口，触发条件 + Stage 1/2 流程 |
| `scripts/review_engine.py` | ✅ | Stage 1 + Stage 2 合并 |
| `scripts/self_check.py` | ✅ | Stage 1 基线 |
| `references/rules.md` | ✅ | 27 条规则库 |
| `references/architecture.md` | ✅ | 本文件 |
| `assets/sample_contract.txt` | ✅ | 自测样例 |
| `assets/stage2_verdict_schema.json` | ✅ | verdicts 格式契约 |

**LLM 调用代码 → 不带**。由各平台 agent 实现：

| 平台 | Stage 2 实现 |
|---|---|
| 桌面（Hermes Agent） | 我（me）作为 agent，读 Stage 1 JSON → 调 LLM（我自己）→ 写 verdicts.json |
| 微信 bot | bot 进程读 Stage 1 JSON → 调 bot 接的 LLM → 写 verdicts.json |
| Web | web session 读 Stage 1 JSON → 调 web 端 LLM → 写 verdicts.json |

**SKILL.md 在每个平台的角色**：告诉 agent "拿到 Stage 1 JSON 后，按这个 schema 产出 verdicts.json，再调用 Stage 2 合并"。

---

## 五、典型调用流程（桌面平台示例）

```bash
# 1. Stage 1（中间文件一律写系统临时目录，不进工作区）
python review_engine.py contract.docx --stage1-only --stage1-output stage1.json

# 2. （平台 agent）读 stage1.json + 合同全文，调 LLM，产出 verdicts.json
#    这一步由 agent 决定如何调 LLM，skill 不管

# 3. Stage 2 合并：--format md 输出原样打印到聊天面板交付（不保存报告文件）
python review_engine.py contract.docx --stage2-input verdicts.json --format md

# 4. 清理（v1.3.2）：交付后立即删除 stage1.json / verdicts.json 等全部中间文件
```

**v1.3.2 交付与清理规则**：

- 报告**不落盘**——`--format md` 输出直接进聊天面板，不写报告文件；
- 中间文件**即用即删**——一律放系统临时目录，审计输出完成后立即删除，工作区不留任何中间产物；
- 用户事后明确要求留档时才例外保存，属用户确认的例外路径。

---

## 六、向后兼容

- **v1.2.2 用户升级后**：`python review_engine.py contract.docx --format md` 输出与 v1.2.2 字节相同
- **新功能可选**：不传 `--stage2-input` 即按 Stage 1 输出
- **fail-safe**：传 `--stage2-input` 后任何违规（如 verdicts 不全、rule_id 非法）都 exit 2，不输出报告

---

## 七、为什么 LLM 调用不写进 Python？

1. **跨平台差异**：每个平台调 LLM 的客户端不同（Hermes Agent / 微信 bot / web 端）
2. **配置解耦**：LLM API key、prompt 模板、temperature 等由各平台 agent 控制，不写死
3. **可审计**：verdicts JSON 是显式契约，可人工 review LLM 的输出
4. **可替换**：以后想换 LLM 厂商只动 agent，不动 skill

**结论**：skill 只负责"确定性判定 + 数据契约"，不负责"调 LLM"。后者是平台 agent 的事。