---
title: docx 法规汇编落档 SOP（v2.1 沉淀）
type: sop
created: 2026-07-13
source: 老板本会话 3 次丢 docx 的实战沉淀
tags: [落档, docx, 双存, rhzy, 操作SOP]
---

# docx 法规汇编落档 SOP

> **触发场景**：老板丢一份 .docx（如 `建设工程全流程法律法规汇编_XX.docx`）并说"落档"。
> **本档用途**：把"docx → wiki 落档"这套 5 步流水线固化，下次同类任务不再问、不要再摸索。
> **本会话实战**：2026-07-13 老板丢 2 份 docx（实名制汇编 + 贵州省地方法规汇编），本 SOP 就是从这 2 次实战总结。

## 🚨 老板在做的事情（铁律）

老板说 **"对该文件内的法律文本进行落档"** 或 **"对该文件内的法律文本进行落档"** = **直接干，不要问。**

**绝对禁止**：
- ❌ "路线 A/B/C/D" 选项题（你给的，老板要看 4 个文件就不看了）
- ❌ "您拍板下一步"——本来落档就是直路
- ❌ 等老板补充信息——除非**完全不可解析**（比如 docx 是扫描件无 OCR 结果）

## 🚨 老板在同一文件上重复发 docx（铁律）

> 本会话老板发了**同一份** `贵州省实名制管理相关法律法规汇编_全文版.docx` **两次** —— md5 哈希相同。

**遇到这种情况**：
1. 第一反应：**先 md5 比对**（不是看文件名）—— 复用上次结果
2. 第二反应：**直接用上次的脚本**（`extract_laws.py` 已存在）—— 不要重写
3. 第三反应：**只更新 _index.md**（加 updated 戳），不要重复全部操作

## 📋 5 步落档流水线（铁律）

### 第 1 步：识别文档结构（探测）

**目标**：找到"按法规切分的分界点"，常见有 3 种分法：
- **a. 文件名编号法**："法规一"、"法规二"、...（实名制汇编就是这种）
- **b. 元数据标签法**：【法规名称】+【发文字号】（贵州省地方法律法规就是这种）
- **c. 章节标题法**：每部法规都有"第 X 章 总则"

**统一方案**：先用 2 个正则找分界点：
```python
# 找"法规 X"
text.startswith("法规") and len(text) < 8
# 找 【法规名称】
text.startswith("【法规名称】") or text.startswith("【条例名称】")
```
命中其中一个就出分界；都不命中再换规则。

### 第 2 步：写 `drop_laws.py` 脚本（铁律）

> 这是 rhzy 体系**首次出现的核心沉淀**——把"5 步"沉淀成可复用的脚本模板。

**模板位置**（在 `gz-construction-law-wiki/scripts/drop_laws_template.py`）：

```python
"""
落档：把 docx 里的 N 部法规原文逐一写入 ~/wiki/concepts/某目录/
每部为一个 .md 文件（含 frontmatter + 元数据块 + 原文）

USAGE: 替换 SRC / DST_DIR / LAWS 三个变量即可。
"""
import docx
from pathlib import Path

SRC = Path(r"...)  # 待改：传入老板给的 docx 路径

LAWS = [
    # 待改：每部法规一行
    # {
    #     "id": 1, "name": "...", "doc_num": "...",
    #     "start": 3, "end": 75,  # 在 docx.paragraphs 中的起止
    #     "file_out": "..._全文.md",
    # },
]

DST_DIR = Path(r"~/wiki/concepts/某目录_全文版")
DST_DIR.mkdir(parents=True, exist_ok=True)

doc = docx.Document(str(SRC))
paragraphs = [p.text.strip() for p in doc.paragraphs]

def make_frontmatter(law):
    return f"""---
title: {law['name']}
law_id: {law['law_id']}
doc_num: {law['doc_num']}
type: regulation/fulltext
publisher: {law['publisher']}
issued: {law['issued']}
effective: {law['effective']}
source_ref: {law['source_ref']}
created: 2026-07-13
source: {law.get('source_doc', '建设工程...docx')}
tags: [贵州省, 地方法规, rhzy, 合规]
---

# {law['name']}

> **法规层级**：{law['type']}
> **文号**：`{law['doc_num']}`
> ...
"""

for law in LAWS:
    text = "\n\n".join(paragraphs[law['start']:law['end']+1])
    (DST_DIR / law['file_out']).write_text(
        make_frontmatter(law) + text, encoding="utf-8"
    )
    print(f"✅ [{law['id']}] {law['name']}")
```

### 第 3 步：核对"法条数"（完整性自检）

**绝不**落档后直接走——必须自检：
```bash
grep -cE "^第[一二三四五六七八九十百]+条" ~/wiki/concepts/<dir>/*.md
```
非 0 才算成功。**空文件 = 失败**。

### 第 4 步：写 `_index.md`（必含双存指针 + 字节数）

> 本会话沉淀出的"6 部 × 2 版本 = 12 个档案"的命名规则：
> - **摘要版**：`<dir>/<法规名>.md`
> - **全文版**：`<dir>_全文版/<法规名>_全文.md`
> - 两者**互相引用**——索引表里每个法规**必须双链接**

索引表头加：
- 总数（部）+ 法条数（条）+ 字节数（bytes）—— 给老板一眼看体量

### 第 5 步：升级"上级目录 _index.md" + 报告

> **铁律**：落档后**必须**把上级目录（如 `~/wiki/concepts/贵州省地方法规/_index.md`）的索引同步升级——加双存版指针列 + 法条数表 + 数据完整性声明。

报告必须给：
- ✅ 落档位置（绝对路径）+ 实测字节数
- ✅ 法条数（grep 计数）
- ✅ 上级 _index.md 升级状态

## 🗂 双存版本的命名规则（铁律）

> **核心规则**：已有摘要版 → 补全文版 = `*_fulltext` / `*_全文` 后缀；如果两份都在 `_index.md` 同一行写出来。

| 类型 | 摘要版 | 全文版 |
|---|---|---|
| **国家级上位法** | `~/wiki/concepts/上位法/<法规>.md` | `~/wiki/concepts/实名制管理汇编/<法规>_全文.md` |
| **贵州省地方法规** | `~/wiki/concepts/贵州省地方法规/<法规>.md` | `~/wiki/concepts/贵州省地方法规全文版/<法规>_全文.md` |

✅ **统一**：用目录名区分"国家级"和"地方法规"（不是同目录混存）
✅ **`_全文` 后缀**：让全文版在文件管理器一眼可见
✅ **`_index.md` 必含双链接**：每一行都有 `[摘要]` + `[全文]` 两个链接

## 📌 落档清单总览（2026-07-13 实测）

| 序 | 源 docx | 目标目录 | 落档部数 | 总字节 | 法条总数 |
|---|---|---|---|---|---|
| 第 1 次 | 实名制汇编 | `~/wiki/concepts/实名制管理汇编/` | 4 | 40,372 | 99 条（57+28+10+4）|
| 第 2 次 | 贵州省地方法律法规汇编 | `~/wiki/concepts/贵州省地方法规全文版/` | 6 | 135,954 | 285 条 + 10 制度 |

**累计**：10 部法规 + 2 个 _index.md = **12 份 .md 文件**

## ⚠️ 老板已知指令（2026-07-13）

老板两次明确：
1. **"对该文件内的法律文本进行落档"** → 直接干 5 步流水线
2. **"列出你已经落档的法律文件"** → 用 ls 实证，给双存版统计

**违反任一 → 重做**。老板 7/13 还说过：
- 涉法必给原文 = 4 件套门控（已并入 gz-construction-law-wiki::SKILL.md §涉法引用纪律）
- 避免自创框架（如"rhzy-law-citation-strict"）→ 没拍板就不写

## 🛠 关联 skill / 工具

- **`rhzy-citation-linker-toolkit::SKILL.md`** —— `format_citation.py` 用于在回答末尾挂 4 层标识 + OSS URL，**新落档档案可点开**
- **`rhzy-citation-linker-toolkit::references/tier3-html-rendering-tdd.md`** —— 上传到 OSS 后必跑 HTML 渲染 + 锚点真实 HTTP GET 验证
- **`~/AppData/Local/hermes/skills/rhzy/gz-construction-law-wiki/scripts/drop_laws_template.py`** —— 落档 SOP 模板（本次沉淀）

## 🎯 SOP 校验 Checklist

下次老板丢 docx 落档前**自检 5 件事**：
- [ ] 老板发的文件是 .docx / .pdf / .xlsx（不是 .jpg / 手写）
- [ ] docx 文本可解析（`python-docx` 能跑通，不是扫描件）
- [ ] `md5sum` 比对——同文件复用上次结果
- [ ] 不要问"是否落档"——直接干
- [ ] 落档后 grep `第N条` 自检 + 更新 _index.md
