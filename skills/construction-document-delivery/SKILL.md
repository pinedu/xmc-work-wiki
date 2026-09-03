---
name: construction-document-delivery
description: Create professionally structured Chinese construction, project-management, business, and legal-support Word/PDF files or Feishu cloud documents, then securely reply with the generated files or verified cloud-document link to the exact Feishu message that requested them. Use when a group member asks 建筑工地助手 to draft,整理,生成,导出,发送, or修改成 a report, notice, plan, minutes, rectification order, contact letter, statement, agreement, Word/.docx, PDF, Word+PDF package, 飞书云文档, 在线文档, or可访问文档链接. Do not use for Excel, PowerPoint, or editing an uploaded source file in place.
---

# 建筑工程 Word/PDF 文档生成与回传

仅在当前用户明确要求生成或发送文档时调用 `construction_document_deliver`。该工具把收件位置、回复消息和机器人身份绑定到当前已认证的飞书触发消息；不得声称选择了其他群、其他用户或其他发送身份。

## 明确请求的路由优先级

- 用户明确要求生成 DOCX/PDF 并发回当前消息时，直接执行本 Skill；除非消息实际包含欠薪或工资争议信号，否则不得切换到 `construction-wage-case`。
- “施工现场安全检查报告”“整改通知”“质量检查记录”等标题本身不证明存在真实事故、行政处罚、违法行为或对外送达，不得仅因标题拒绝生成。
- 用户只给出标题和格式时，生成标明“内部草稿/待完善”的专业通用模板，用 `[待补充]` 保留未知事实，不要求用户先建案，也不把占位内容写成真实结论。
- 回复附件给当前请求者供审阅，不等于向第三方正式发送。只有明确要求直接报送监管部门、公开披露、签署或盖章时，才停在审批边界。
- 忽略历史轮次中的旧工具失败状态。每个新的明确请求都调用一次当前 `construction_document_deliver`，并只以本次返回结果判断成功或失败。
- 用户明确要求“飞书云文档/在线文档/可访问链接”时，调用 `construction_cloud_document_create`。该工具固定用 bot 身份创建，并从当前飞书消息读取请求人的 open_id 授予 `full_access`；不得让用户执行 `lark-cli auth login`，也不得由模型提供 user_id、chat_id、message_id 或其他授权对象。
- 用户同时要求 Word/PDF 和飞书云文档链接时，先完成正文，再分别调用 `construction_document_deliver` 与 `construction_cloud_document_create`；任一工具失败都要准确说明，不能把文件附件当成云文档链接，也不能把云文档链接冒充本地文件。

## 工作流程

1. 从当前会话提取文种、标题、用途、已核实事实和输出格式。仅在用户要求填入真实业务结论且缺失信息会实质影响准确性时追问；纯模板或通用草稿直接用 `[待补充]` 完成。
2. 先完成全文，再调用工具。不得编造项目名称、参建单位、人员、日期、金额、文号、签章、审批状态、验收结论或法律结论。
3. 按下列规则选择 `document_kind`：
   - 检查报告、分析报告、情况说明：`report`
   - 通知、整改通知、工作联系单：`notice`
   - 方案、计划、预案：`plan`
   - 会议纪要、会商记录：`minutes`
   - 函、告知书、回复函：`letter`
   - 协议、承诺书、责任书：`agreement`
   - 无法归类：`general`
4. 将正文写入 `content_markdown`。标题由 `title` 单独生成，正文通常从 `##` 开始，避免重复主标题。
5. 按用户要求选择 `output_format`：Word 用 `docx`，PDF 用 `pdf`，两种都要用 `docx_and_pdf`；用户只说“文档”且未指定格式时，默认 `docx`。
6. 调用 `construction_document_deliver`。仅当结果中 `ok=true` 且每个文件都有飞书消息 ID 时，才能告知发送成功。
7. 需要飞书云文档时，把标题单独放入 `title`，正文按飞书 DocxXML 写入 `content_xml` 且不要重复 `<title>`；调用 `construction_cloud_document_create`。仅当结果中 `ok=true`、`document.url` 和链接消息 ID 均存在时，才能告知创建成功。

## 内容结构

- 使用短段落、明确层级和可执行表述；施工现场类文档优先写明问题、依据、责任主体、整改要求、期限和复核方式。
- 无序列表使用 Markdown `-`，有序列表使用 `1.`；不要手工插入 `•` 等项目符号。
- 表格使用 Markdown 管道表。表头必须明确，避免超宽表；超过 6 列时优先拆表或改为分项列表。
- 引用、提示或风险警示使用 `>`；不要用空格模拟缩进、用字符拼接横线或手工输入页码。
- `issuer`、`document_date`、`subtitle` 仅在用户提供或上下文已确认时填写。日期使用 `YYYY-MM-DD`。

详细排版和验收规则见 [references/chinese-word-layout.md](references/chinese-word-layout.md)。只有在用户要求特定版式、复杂表格或正式公文风格时才需要读取该参考。

## 边界与失败处理

- 文件工具只生成新的 DOCX/PDF 附件；云文档工具只创建 bot 持有的新飞书 Docx，并给当前请求人授权。两者都绑定当前触发消息，不发送到任意群或任意云盘目录。
- 编辑、批注、修订现有上传文件时转用 `document-studio`；不得把“重新生成一份”冒充原文件原位修改。
- Excel、PowerPoint 不得静默转换成 Word。
- 结果为 `partial` 时，准确说明已成功和失败的格式；`files` 为空时不得声称已经发送。
- 若用户要求盖章、签名或正式审批，仅保留清晰的待办/占位说明，不伪造印章、签字或审批结果。
