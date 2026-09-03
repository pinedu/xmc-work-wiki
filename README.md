# xmc-work-wiki

薪莫愁训练知识仓库。

## 仓库定位

本仓库是训练后 Skill 的版本化知识库。每个 Skill 存放在 `skills/<skill-name>/` 目录，入口文件为 `SKILL.md`；仓库根目录的 `manifest.json` 提供 Skill 名称、版本、描述和输入/输出 schema，便于程序发现和加载。

## 与 ToolPlane 的关系

当前 Agent 会将训练中沉淀并完成落档的 Skill 推送到本仓库。ToolPlane 通过读取本仓库中的 Skill，将合适的能力分配给对应 Agent 调用：先从 `manifest.json` 识别可用能力，再读取目标 Skill 的 `SKILL.md` 及其引用资源，最后由对应 Agent 按 Skill 约束执行任务。

## 加载与调用方式

ToolPlane 通过 GitHub MCP 连接加载本仓库地址：

`https://github.com/pinedu/xmc-work-wiki.git`

加载流程如下：

1. ToolPlane 的 MCP GitHub 连接访问仓库。
2. 读取仓库根目录的 `manifest.json`，按名称、版本、描述和 schema 选择匹配的 Skill。
3. 加载 `skills/<skill-name>/SKILL.md`，以及该 Skill 所引用的资源文件。
4. 将 Skill 上下文和任务输入分配给对应 Agent，由该 Agent 完成调用并返回结果。

## 微信训练群归档

在微信训练群中形成并已落档的 Skill，需要发布到训练知识库时，发送触发关键词“归档”。归档通道会将对应 Skill 自动推送到本 GitHub 仓库。

“归档”是显式发布动作：普通聊天消息或仅上传文件不会自动推送；训练完成后发送“归档”，才会进入仓库发布流程。

## Skill 维护约定

- 每个 Skill 的 `SKILL.md` frontmatter 必须声明 `name`、`version` 和 `description`。
- 新增或更新 Skill 时，同步维护仓库根目录的 `manifest.json`，使其与 Skill 的名称、版本、描述和 schema 一致。
- 首次正式发布使用 `1.0.0`；后续版本按变更程度递增。
