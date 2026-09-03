---
name: rhzy-law-supplement-pipeline
description: "rhzy 法规补档 + 结构化全自动流水线 — 两阶段闭环：Phase 1（老板发 docx 法规汇编/单部法规/国务院令/地方政府令时落档 md）+ Phase 2（老板已落档 ≥ 2 部法规要做全文检索/按条款速查/跨法规对照/场景速查时结构化为 6 文件 wiki）。Phase 1 用 7 轮沉淀方法学（MD5 同源检测 → python-docx 原文逐字抽取 → 章节边界自动识别 → 法条编号去重 → 字节/法条数实测自检 → 索引同步），5 分钟 1 部。Phase 2 用 6 步 SOP（勘盘 → README → 逐部切片 → 跨法规对照 → 场景速查 → 验证）+ 4 大 Pitfall，10 分钟一批。覆盖 7 类档案 + 4 类结构化产物。触发场景：老板发 docx 让落档 / 法规补档 / 法规扩容 / 飞书收到 docx 法规文件 / wiki 法规档案建设 / 老板选 B 选项做结构化沉淀 / 法规跨法规对照 / 5 大场景月度 checklist。"
version: 1.0.0
author: 小建建 (Hermes Agent)
license: MIT
metadata:
  hermes:
    tags: [law, supplement, pipeline, rhzy, docx, 法律补档, 法规汇编]
    related_skills:
      - gz-construction-law-wiki
      - rhzy-law-citation-strict
      - rhzy-citation-linker-toolkit
      - rhzy-functional-coverage-audit
---

# rhzy 法规补档 + 结构化全自动流水线 · v2.0

> **2026-07-13 实战沉淀 — 两阶段闭环固化**

老板在本日**7 次丢法规 docx 让我落档**（Phase 1）+ **1 次做结构化沉淀**（Phase 2 = 4 部实名制法规 → 6 文件结构化 wiki）。两阶段方法学全固化在本 skill。

## 📌 这是什么

一个**法规 docx 落档 + 结构化流水线**——两阶段闭环：

### Phase 1（落档：docx → md）
1. **老板发来 docx 怎么落档**（5 步 SOP）
2. **落档后怎么同步索引**
3. **老板跳过"A/B/C 选项"**（除非 OCR 失败）
4. **MD5 检测避免重复落档同源 docx**

### Phase 2（结构化：md → 6 文件 wiki）
1. **老板已落档 ≥ 2 部法规 → 想要"全文检索 + 按条款速查 + 跨法规对照 + 场景速查"**
2. **6 文件产出**：README 索引 + N 部法规逐条切片 + 1 张跨法规对照表 + 1 张场景速查表
3. **三段式设计**：原文 + 关键义务 + 项目落地（红黄绿灯）
4. **就高从严原则**：地方版 vs 国版冲突时走更严的

**反例 → 正例**：
- ❌ 老板丢 docx 让我"对法律文书落档" → 我问"老板您是落档到 A 国家级上位法，还是落档到 B 贵州省地方法规，还是落档到 C 实名制汇编？" → **老板很烦**
- ✅ 老板丢 docx 让我落档 → 我先 md5、跑 docx 解析、按文件名前缀/文号/发布机关自动归类、跑落档脚本、写索引、自测字节数 → **5-10 分钟完成**
- ❌ 老板说"4 部实名制法规已落档，但用着不顺手" → 我重新解释每部法规 → **老板没耐心**
- ✅ 老板选 B 选项做结构化 → 6 步 SOP 跑完 → 6 文件 wiki 落档 → 全文检索 + 按条款速查 + 跨法规对照 + 场景速查 **全部可用**

## 🎯 何时触发（必读）

以下任一情况立即加载本 skill：

### Phase 1 触发场景（落档）

| 触发场景 | 关键词 / 行为 |
|---|---|
| 老板发送 docx 文件 + "落档" | `对该文件内的法律文本进行落档` `进行落档` |
| 老板要求"补档"某部法规 | `落档 第X号` `补档` |
| 法规汇编批量落档 | `建设工程全流程法律法规汇编` `实名制管理相关法律法规汇编` |
| 老板发送单部法规 | `政府投资条例.docx` `建设工程勘察设计管理条例.docx` |
| 老板未明示归类 | 老板发 docx 没说放哪个目录 |

### Phase 2 触发场景（结构化）⭐ NEW 2026-07-13

| 触发场景 | 关键词 / 行为 |
|---|---|
| 老板选"B 选项做结构化沉淀" | `做结构化` `本地结构化 wiki` |
| 老板要"全文检索 + 按条款速查" | `全文检索` `按条款速查` |
| 老板要"跨法规对照表" | `国版 vs 贵州版` `跨法规映射` `就高从严` |
| 老板要"场景速查表" | `5 大场景速查` `项目月度 checklist` |
| 老板新增 ≥ 2 部同主题法规 + 旧法规联动 | `新法规 + 旧法规 对照` |

> ⚠️ **Phase 1 自动覆盖"老板已发 docx + 未明确归类"**——按文件名前缀 + 文号 + 发布机关自动判别。
> ⚠️ **Phase 2 自动覆盖"老板已落档 ≥ 2 部同主题法规 + 想要可查可定位"**——跑 6 步 SOP 出 6 文件 wiki。
> ⚠️ **本 skill 不覆盖**——OCR 失败的扫描件（如签字栏扫描）、老板明示"先放某个目录"、docx 内含多项合并文件（需老板分项确认）、老板只要单部法规的 raw md（已落档够用）。

## 🛠 5 步 SOP（每次"落档"指令都跑）

### Step 1: MD5 同源检测（避免重复落档）⭐⭐⭐

```bash
echo "=== 老板发的 docx 路径 ==="
ls -la "/c/Users/rhzy/AppData/Local/hermes/cache/documents/<新docx>.docx"

echo "=== 跟已落档档案比对 md5（防重复）==="
md5sum "/c/Users/rhzy/AppData/Local/hermes/cache/documents/<新docx>.docx" \
       "/c/Users/rhzy/AppData/Local/hermes/cache/documents/<已落档同源docx>.docx"
```

**判定规则**：
- MD5 一致 = 复用已落档档案 = **不再处理**
- MD5 不一致 = 已落档同源但版本不同 = 跑老板确认（防 711/722/743 等文号相似的混淆）
- 无同源 = 继续 Step 2

### Step 2: python-docx 解析文档结构

```bash
cat > /tmp/inspect_doc.py << 'EOF'
import docx, sys
src = sys.argv[1]
doc = docx.Document(src)
paragraphs = [p.text.strip() for p in doc.paragraphs]
print(f"段数: {len(paragraphs)}, 表格: {len(doc.tables)}")
for i in range(min(35, len(paragraphs))):
    t = doc.paragraphs[i].text.strip()
    if t:
        print(f"P{i:03d} | {t[:140]}")
EOF
python /tmp/inspect_doc.py "/c/Users/rhzy/AppData/Local/hermes/cache/documents/<新docx>.docx"
```

**输出判定**：
- **单部法规汇编**：头几段有"国务院令第X号 / 通过日期 / 修订历程 / 标题"
- **多部法规汇编**：第 1-3 段是头部，后续有"法规一""法规二""法规三"分界
- **法规标题位置**：P008-P009 通常是条例大标题
- **章节边界**：找 "第一章 X" "第二章 Y" 行
- **法条计数**：找 "^第[一二三四五六七八九十百零]+条" 行（去重 + 排除注释行）

### Step 3: 自动归类到 5 类目录

| 文号特征 | 文件名前缀 | 落档目录 |
|---|---|---|
| 国务院令 / 主席令 / 部委令 | `中华人民共和国` 或 国务院规章 | `~/wiki/concepts/国家级上位法全文版/` |
| 法规一/二/三 编排 + 国务院令 / 主席令 | `中华人民共和国` 或 国务院规章 | `~/wiki/concepts/国家级上位法全文版/` |
| 地方政府令第 X 号 / 各省条例 | 贵州省 / 市政府 / 省政府办公厅 | `~/wiki/concepts/贵州省地方法规全文版/` |
| 黔** / 各省行政文件 | 文件名含"实名制/工资/招投标/..." | `~/wiki/concepts/实名制管理汇编/`（**仅限建市/黔人社发/国务院令的特定组合**）|

### Step 4: 复用 templates/drop_laws_template.py（已沉淀在 gz-construction-law-wiki）

```bash
# 1. 复制模板（替换 SRC / DST_DIR / LAWS 三个变量）
cp ~/.hermes/skills/rhzy/gz-construction-law-wiki/scripts/drop_laws_template.py \
   /tmp/drop_<law_short_name>.py

# 2. 修改 LAWS 元数据数组（参 gz-construction-law-wiki::references/drop_docx_sop.md 第 4 步 双存命名规则）

# 3. 跑脚本
python /tmp/drop_<law_short_name>.py

# 4. 自动写入 大索引 _index.md（含 frontmatter + 章节 / 条文 / 字节数自检）
```

### Step 5: 同步索引（**绝不漏这一步**）

```bash
# 5.1: 上位法主索引 + 全文版中央索引 + 摘要版（双存）
patch 上位法/_index.md（加 1 行表格 + 1 段更新历史 + 1 个 rhzy 项目高频引用条款）
patch 国家级上位法全文版/_index.md（加 1 行表格 + 章数/条数/层级统计）

# 5.2: 贵州省地方法规（双存）
patch 贵州省地方法规/_index.md（同上）
patch 贵州省地方法规全文版/_index.md（写完整 rhzy 项目必读条款速查）

# 5.3: 实名制管理汇编（特殊）
patch 实名制管理汇编/_index.md
```

**强制规定**：每落档 1 部，老板没异议就立即同步索引。不分批，不等老板指令。

## 🛠 7 类档案 · 落地位置速查

| 档案类型 | 落档目录 | 文件名模板 |
|---|---|---|
| 国家级上位法（国务院令/主席令/部委令）| `~/wiki/concepts/国家级上位法全文版/` | `<法规名>_全文.md` |
| 贵州省地方性法规 / 省政府令 / 省级规范性文件 | `~/wiki/concepts/贵州省地方法规全文版/` | `<法规名>_全文.md` |
| 实名制管理汇编（含国务院令+黔人社发+黔府办发组合）| `~/wiki/concepts/实名制管理汇编/` | `<法规名>_全文.md` |
| 国家级上位法·摘要（按业务方向提炼）| `~/wiki/concepts/上位法/` | `<法规名>.md`（**不直接落档 docx**，只是对全文版做 RHZY 业务索引）|
| 贵州省地方法规·摘要 | `~/wiki/concepts/贵州省地方法规/` | 同上 |
| 业务主题 → 法条 · 反向索引 | `~/wiki/topics/README.md` | （patch）|
| 业务主题专门笔记 | `~/wiki/topics/<topic>.md` | （如需新建）|

## 🛠 Phase 2 SOP（结构化：已落档 md → 6 文件 wiki）⭐ NEW 2026-07-13

**触发条件**：老板已落档 ≥ 2 部同主题法规（如：4 部实名制法规）+ 想要"可查可定位"。

**完整方法学见** `references/structured-wiki-sop.md`——本档只列 6 步骨架：

| 步 | 动作 | 关键产物 |
|---|---|---|
| Step 1 | 勘盘已落档原文（条数自检）| 完整性 OK / FAIL |
| Step 2 | 建结构化目录 + README 索引 | 1 文件 README（5 段式） |
| Step 3 | 逐部法规切片（"原文 + 关键义务 + 项目落地"三段式）| N 文件切片 |
| Step 4 | 跨法规对照表（含就高从严原则）| 1 文件对照表 |
| Step 5 | 场景速查表（5 段式 + 月度 checklist）| 1 文件速查表 |
| Step 6 | 验证 + 落档清单（引用铁律自检）| 6 文件全部 ✅ |

**6 文件标准产出**（如 4 部实名制法规）：
```
README_实名制法规结构化索引.md           ← 总入口
01_<修订通知>_逐条切片.md              ← 第 1 部
02_<国家上位法>_N条逐条切片.md           ← 第 2 部
03_<地方版法规>_N条逐条切片.md           ← 第 3 部
04_<国版>vs<地方版>_同事项条款对照表.md  ← 跨法规映射
05_<N大场景>速查表_<场景名>.md         ← 场景→条款
```

**老板决策记录（A/B/C/D 选项）**：
- A 仅 docx 留档 ❌（已做，5 秒；老板查不到）
- **B 本地结构化 wiki ✅**（10 分钟，全文检索 + 按条款速查）
- C B + OSS 推送（15 分钟，跨设备 + 永久留存）
- D C + 场景化反索赔剧本（25 分钟，开箱即用）

老板当前锁定 **B**——C/D 后续可选升级。

---

## ⚠️ 4 大 Pitfall（实战沉淀 · 7 轮踩过）

### Pitfall 1：MD5 同源 docx 重复落档（最常见）

**反例**：
- 老板第 1 次发 `实名制管理相关法律法规汇编_全文版.docx` → 我落档 4 部
- 老板第 2 次发**同源** docx（md5 一致）→ 我**重新跑 5 步 SOP** → 重新生成 4 部 → 文件覆盖 → 老板"为什么又来一次？"

**正确做法**：第一步先 md5 比对，发现同源 = **直接不动 + 报告"已经落档完毕"**。

> 2026-07-13 老板第 2 次发同名 docx（hash 一致）时直接复用。

### Pitfall 2：把"老板被通知老板但实际已落档"的状态标 ⚠️

**反例**：
- 我标 ⚠️ "《土地管理法》（主席令第 32 号）—— ⚠️ 本地无档案"
- 但实际**已经落档** — `ls ~/wiki/concepts/上位法全文版/中华人民共和国土地管理法_全文.md` = 存在
- 老板原话："⚠️ 上位法：《土地管理法》（主席令第 32 号）—— 已落档全文版，这个法律文件不是已经落档了吗刚刚"

**根因**：写之前没 probe。

**正确做法**：写任何"未落档"前先 `ls` + `grep`。"未落档"和"已落档"不能共线。

### Pitfall 3：法条去重要排除注释行

**反例**：贵州省 202 号第 13 条注释 "第十三条及本条特殊、紧急项目采取'一事一议'方式报送"——这不是真的第 13 条，是注释行。如果不去重 = 计数错误。

**正确做法**：
```python
seen = set()
arts = []
for line in text.splitlines():
    m = re.match(r"^第([一二三四五六七八九十百零]+)条", line)
    if m:
        n = m.group(1)
        # 排除明显的注释行
        if "一条特殊" not in line and "紧急项目" not in line:
            if n not in seen:
                arts.append(n)
                seen.add(n)
```

### Pitfall 4：把法条编号的判断用 "数字" 转换而忽略两位数

**反例**：
- 贵州省 202 号第 30 条正文 + 第 31 条 = 10 行以上
- 我用 `cn_num = {...}` dict 转中文→数字 —— `"三十"` 在 dict 里但 `"三十一"` 不在 → 第 31 条转换失败 → 计数错误

**正确做法**：**直接用中文数字当 ID 比较，不做转换**，或者用完整的 `cn_num = {'一':1, '二':2, ..., '五十':50, '一百零':100, '二十三':23, '三十一':31, ...}` 展开成完整 dict。

## 🔗 关联 Skill

| 关联对象 | 关系 |
|---|---|
| `gz-construction-law-wiki` | 总览（27→32 部法规 wiki 全集）|
| `gz-construction-law-wiki::references/drop_docx_sop.md` | **完整 SOP** — 落档 5 步 + 本档实战沉淀 |
| `gz-construction-law-wiki::scripts/drop_laws_template.py` | **落档模板脚本** — 替换 3 变量即可复用（已实战 7 次）|
| `gz-construction-law-wiki::references/legal-archive-roll-2026-07-13.md` | 32 部法律档案实证清单（含字节数/法条数/4 层标识）|
| `rhzy-law-citation-strict` | 落档后的引用 4 件套门控 |
| `rhzy-citation-linker-toolkit` | 4 层标识基础设施 |
| `rhzy-functional-coverage-audit` | 落档后的章节归属（前期开发部/工程管理部等）|

## 📂 skill 目录结构（本 skill · v2.0）

```
~/AppData/Local/hermes/skills/rhzy-law-supplement-pipeline/
├── SKILL.md                                # 本文件（Phase 1 + Phase 2 双闭环）
└── references/
    ├── drop_docx_sop.md                   # Phase 1 完整 5 步 SOP（含 4 大 Pitfall 详细示例）
    ├── 实战案例-7 轮补档记录.md             # 2026-07-13 7 轮落档完整时间线 + 每轮 byte/章节/条文数
    └── structured-wiki-sop.md             # ⭐ Phase 2 完整 6 步 SOP（含 4 大 Pitfall + 6 文件模板）NEW
```

## 🔄 维护与更新

- 本 skill **2026-07-13 升级到 v2.0**——新增 Phase 2（结构化沉淀）
- Phase 1：7 轮实战沉淀（5 步 SOP + 4 大 Pitfall + 落档模板脚本）
- **Phase 2：1 轮实战沉淀**（4 部实名制法规 → 6 文件结构化 wiki；6 步 SOP + 4 大 Pitfall）
- 任何 rhzy 项目"老板发 docx 让我落档" → 跑 Phase 1
- 任何 rhzy 项目"老板已落档法规但想可查可定位" → 跑 Phase 2
- 若老板要求"换目录分类"或"OSS 推送"等新约束 → 升级 v2.1 + 加新 Pitfall
- **沉淀进 skill 而非 memory** —— 老板原话："你以后给个清单给我，方便补档时知道哪些需要补"（rhzy-law-supplement-pipeline::references/实战案例-7 轮补档记录.md 就是这份清单）

## 🗒️ 老板决策记录（2026-07-13）

| 决策点 | 选项 | 选定 | 理由 |
|---|---|---|---|
| 落档流程必跑 | A 我自己跑 SOP / B 老板拍板后再动 / C 5 步 SOP 跑完直接告诉老板结果 | **A 5 步 SOP（自动跑）** | 老板多次催"快点落档——别再问我 A/B/C" |
| MD5 检测 | A 跳过 / B 必跑 | **B 必跑** | 老板第 2 次发同名 docx，发现 hash 一致 = 复用，不再走流程 |
| 索引同步 | A 老板说同步再同步 / B 落档完立即同步 | **B 立即同步** | 老板"别让档案目录看起来半生不熟" |
| Pitfall 反例位置 | A 留在 SKILL.md / B references/ | **B references/drop_docx_sop.md** | 长文不让 SKILL.md 变得不可读 |
