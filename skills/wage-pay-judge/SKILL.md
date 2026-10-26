---
name: wage-pay-judge
description: 判定工人工资是否超付/欠付。上传（或指定）当月工人工资册与考勤后，计算工人日薪（当月实发÷当月考勤），与工种市场单价区间比对；日薪高于工种上限×1.30判超付，低于工种下限×0.85判可能欠付并进入第二轮——用该工人全部工作历史的平均日薪再比对工种下限，低于下限坐实为欠付。适用于建设单位/甲方对农民工工资合规的批量核查。
---

# 工资超付/欠付两轮判定 Skill

## 一、定位

建设单位（甲方/业主）视角下，对农民工工资发放的**超付、欠付**做批量、可审计的合规判定。
输入：当月工资册 + 当月考勤（及历史库、花名册、工种单价表）；输出：逐人判定报告（Markdown）。

## 二、判定逻辑（状态机）

对每一位当月发放工资的工人，先算 **日薪 = 当月实付 ÷ 当月考勤天数**，然后：

**第一轮（工种市场单价区间卡）**
- 日薪 > 该工种**上限 × 1.30** → **超付**
- 日薪 < 该工种**下限 × 0.85** → **可能欠付**（进入第二轮）
- 介于两者之间 → **正常**

**第二轮（仅“可能欠付”人员）**
- 历史平均日薪 = 该工人**全部工作历史**（分析月之前所有可用月份）的“实付÷考勤”均值
- 历史平均日薪 < 该工种**下限** → **欠付**（坐实）
- 历史平均日薪 ≥ 该工种**下限** → 维持 **可能欠付（历史未佐证·待复核）**
- 该工人无可用历史月份 → 维持 **可能欠付（无历史数据·待复核）**

最终四态 + 两异常态：**正常 / 超付 / 可能欠付 / 欠付 / 无法计算 / 缺工种·需人工**。

## 三、数据关联与来源

| 数据 | 来源 | 关键字段 |
|---|---|---|
| 当月工资册 | `<data>/工资/<YYYYMM>-*.xlsx`（按月多文件，按身份证汇总实付） | 身份证号、姓名、实付、是否短期工、班组名称 |
| 当月考勤 | `<data>/黔薪保考勤/考勤人员信息<YYYYMM>.xlsx` | 姓名、班组、考勤天数 |
| 花名册（主索引） | `<data>/花名册/花名册合并数据-总表.xlsx`（sheet=总表） | 身份证→工种、进退场 |
| 工种单价区间（判定锚） | `<data>/工种单价预估范围.xlsx` 或 `references/工种单价预估范围.md` | 工种→底线/上线（元/工日） |
| 历史库 | 自动读取 `<data>/工资/` 与 `<data>/黔薪保考勤/` 中分析月之前所有月份 | 同上 |

关联链：
- 工资册 ↔ 花名册：用 **身份证号**
- 工资册 ↔ 考勤：用 **(姓名, 班组)**
- 工种名 → 单价区间：先查 `references/工种别名映射.md` 的【别名】行，再精确匹配，都无则标“未匹配·需人工”

## 四、处理流程

1. 加载当月工资册（多文件按身份证汇总实付）+ 当月考勤，按 (姓名,班组) 内连接 → 实付、考勤天数 → 日薪
2. 用身份证挂花名册 → 工种、进退场状态
3. 工种名匹配《工种单价预估范围》→ 底线、上线
4. 第一轮判定（区间卡）
5. 仅对“可能欠付”者：算其全部历史月份的日薪均值，再卡下限
6. 产出 Markdown 报告（口径说明 + 汇总 + 逐人明细 + 待确认清单）

## 五、边界与异常（都会在报告中单列，不静默丢弃）

- 考勤天数 = 0 / 空 / 工资册在考勤查不到 (姓名,班组) → **无法计算**
- 工资册身份证在花名册查不到工种 → **缺工种·需人工**
- 工种名匹配不上单价表 → 标“未匹配·需人工”
- 临时工（是否短期工=是）、已离场（进退场=离场）：**全量纳入判定**，报告中打标签，不预先剔除
- 工资册同身份证出现多行（含“存在重复人员”备注）→ 按身份证汇总实付，并打“工资册重复记录×N”标签，供人工复核
- 考勤/工资班组不一致时，退化为按姓名匹配并打“按姓名匹配(班组不一致)”标签

## 六、运行方式

统一数据源（qxb 黔薪保平台 MCP，唯一外部数据源）：

```bash
python scripts/judge.py --source qxb --target 26.8 \
    --project-name <项目名> \
    --config judge_config.sample.json \
    --history all \
    --json out.json
# --data 仅用于定位本地《工种单价预估范围.xlsx》；缺省回退 references 的 md；多项目需 --project-name
```

参数说明：
- `--target`：分析月 YYYYMM 或 `YY.M`（必填）
- `--source`：qxb（唯一外部数据源；2026-10-13 起 standard/zhongyi/db 文件源与 HTTP 接口源已移除）
- `--project-name`：qxb 项目名，运行时从**当前账号**项目列表匹配解析 project_id（单项目可省略；多项目必填；也可 `--project-id` 显式覆盖）
- `--data`：仅用于定位本地工种单价表（默认当前目录；找不到则回退 skill `references/工种单价预估范围.md`）
- `--config`：判定口径 json（仅阈值/工种解析等，**不含 project_id**）；不传用标准口径
- `--alias`：工种别名映射 md，默认 `references/工种别名映射.md`
- `--history`：`all`=该工人全部历史（默认）；或整数 N=最近 N 月
- `--json`：结构化结果 JSON 输出（供月报组装层调用）
- `--out`：Markdown 报告输出路径

依赖：Python 3 + openpyxl（仅读取 xlsx，不修改源数据）。

## 七、输出报告结构（Markdown）

1. **判定口径**：分析月、日薪公式、两轮阈值、历史窗口取值、关联链
2. **汇总**：超付 X 人/估算金额、欠付 Y 人/金额、可能欠付 Z 人、正常 U 人、无法计算/缺工种 V 人
3. **逐人明细**：姓名、班组、工种、考勤天、实付、日薪、下限×0.85、上限×1.30、第一轮、历史均值、下限、最终状态、标签/备注
4. **待确认清单**：无法计算 / 缺工种 的人员及原因

> 金额（超付/欠付）仅按区间阈值的粗估，用于排序与预警，**不作为最终结算依据**。

## 八、已知数据质量注记（本项目）

- 工资册按班组拆成多文件，同身份证可能跨文件出现（已汇总处理）
- 部分工人工种被记为“其它”（兜底工种，区间 150–300 偏低），其“超付”判定可能是**工种填错**导致的假性超付，需结合花名册/合同复核
- 花名册个别工人（如徐金宜、朱洪章）未进入花名册主档，会落“缺工种·需人工”，需补录
- 单价表已规范化：同工种唯一区间；此前 16 个漏区间工种已按住建劳务日工资指导价/招聘市场价补齐

## 九、数据源政策与 qxb 数据源（黔薪保平台 MCP）

**数据源政策（2026-10-13 老哥拍板）：外部数据源只保留 qxb（黔薪保平台 MCP）**。
standard/zhongyi 文件源、db HTTP 接口源、custom 自定义适配器已全部移除（`datasource_db.py`、
`datasource_template.py` 已删除，`make_datasource` 仅接受 `qxb`/`mcp`，传入其他值直接报错）。
skill 的 `references/`（工种别名映射、工种单价预估范围）不受影响，保留。

`--source qxb` 经 `QxbMcpDataSource`（`scripts/datasource_qxb.py`）走 **streamable HTTP MCP**（标准库 JSON-RPC，无 SDK 依赖）从黔薪保平台取数：

```bash
python judge.py --source qxb --target 26.8 --project-name <项目名> \
    --config judge_config.sample.json --json out.json
# 单项目账号可省略 --project-name（自动选用）；多项目用 --project-name 指定其一；也可 --project-id 显式覆盖
```

- **连接**：默认复用 WorkBuddy `~/.workbuddy/mcp.json` 的 `qxb` server（url+token 不落第二处）；可在 judge_config `qxb` 段覆盖（url/headers/page_size/timeout）。
- **项目解析（重要）**：`project_id` **不在任何配置文件里固化**——运行时由 `qxb_projects_page` 列出**当前登录用户**名下的项目，按 `--project-name` 匹配或单项目自动选用（`resolve_project()`）；多项目未指定则报错列出可选全称。这样 skill 多用户/多项目复用时不串数据。
- **映射**：02/12 工资→`qxb_salary_details_list`（**该工具无月份参数**，上游对额外传入的 payMonth/month 静默忽略——全量翻页拉回后本地按 payMonth 筛）；03 考勤→`qxb_attendance_records_list`（按 projectId+begin+end 拉打卡流水，不依赖分包 subId）；06 产值→`qxb_monthly_outputs_list`；花名册→`qxb_workers_list`（降级 `qxb_contract_persons_list`，再降级工资册合成）。
- **人的唯一锚（2026-10-21 升级"用身份证"）**：工资明细证件为掩码（前 6 位+12 星+后 4 位）、花名册/考勤矩阵为全号 → **掩码夹逼匹配**（键=前 6+"|"+后 4，勿写死星号数；**必须"证件匹配+姓名一致"双保险**——实测存在掩码撞车：同区县+同后 4 的两个真人，5/105 对完全不同姓名，单靠掩码会挂错工种/错并工资）。锚优先级：证件号 → 姓名 → 班组名反推；花名册同名同姓自动告警。**考勤矩阵同证件号多姓名**（平台录入笔误，8 月实测 6 对如 李红海/李洪海、宠/庞治强）→ 按"在职行优先、打钩权重"归并为同一人（canonical 姓名），kaoqin 返回三元组 `(by_nt, by_name, id_days)`，id_days={全号:并集天数} 供判定侧掩码夹逼直查天数（姓名笔误断链时兜底）；流水口径无证件号，id_days 空、仍姓名锚。**矩阵 0 天行也必须登记证件号→姓名**（否则在职行 0 天+离场行有天数时 canonical 会错取离场拼写）。**工资册同掩码号多姓名**（1 月 10+ 对，如 申学俊/申恒强 两日期金额完全相同）→ **不自动归并**（疑似重复发放或掩码撞车），仅告警交人工核实——自动归并会把两人工资相加制造假超付。判定主键：工资册内部=personId；salary_agg 另带 `idcard_masked` 字段供锚定。
- **降级**：平台无对应工具（04/05/07/08/09/11）或端点故障（workers_list/subcontractors_list 的 zod 校验 bug、exit_unsettled 404、wage_accounts 权限；~~project_get zod bug~~ **2026-10-21 平台修复**——project_meta 已改为首选 `qxb_project_get`，一次直出项目档案+五方单位中文名+11 类岗位人员，附件 OCR 降为兜底）→ 返回空 → 月报对应板块【待填】+《待补充数据清单》；**服务端修复后本类无需改码自动生效**。
- **配置**：judge_config 仅承载判定口径（`threshold_mode`/`second_round`/`gongzhong_resolve`/`team_gong_map` 等），**绝不含 project_id / 具体项目路径**；price 锚默认就近《工种单价预估范围.xlsx》，回退 skill 自带 `references/工种单价预估范围.md`（市场基准，随 skill 分发，非项目数据）。
- **判定口径**：由 `--config judge_config.sample.json` 提供（如中医药 v4：warn + hold + team_fallback）；不传 config 时用 `JudgeConfig.standard()` 标准口径。
- **附件扫描兜底通道**（同源第二取数途径，仍取自 qxb 平台，非第二数据源）：MCP 工具拉不到的字段，用同平台的**附件扫描**兜底——
  `scripts/scan_attachments.py`（调 `QxbMcpDataSource.attachment_scan()`）枚举 `project_files_list` /
  `project_contract_files_list` / `settlement_files_list` 的项目公开附件（施工合同扫描件、农民工维权
  信息告示牌、工伤险完税证明等），下载图片到**系统临时目录**（不落项目/ skill 目录，不外传第三方）并生成
  `manifest.json`；**文字提取由 agent 读图完成（半自动）**，提取结果经 `--attach-data` **内联**传给
  `build_full_report.py`，**不写回任何项目/ skill 文件**。

### ⚠️ qxb MCP 串行调用铁律（强制约束，2026-10-16 老哥拍板）

探测黔薪保平台 MCP 时，**严禁并发调用工具**——必须**串行**，且每次调用之间**间隔 5 秒**（2026-10-19 由 3 秒上调）。
否则平台会触发风控（限流 / 封 IP / 拒绝服务），所有 40+ 个工具全部拉空。

- **适用范围**：所有 qxb MCP 工具（salary_details_list / attend_table / projects_page / workers_list / subcontractors_list / project_get / monthly_outputs_list / attachment_scan / 等），无例外
- **节奏**：取上一条结果 → `sleep(3)` → 调下一条
- **实现位置**：`scripts/datasource_qxb.py` 的 `_McpHttp.call()` 入口（所有工具调用都过这里，单进程天然串行）
- **副作用**：单次月报生成耗时会显著拉长（每多一次 MCP 调用 +5 秒），但项目/ skill/第三方数据完整性优先
- **例外**：`_initialize()`（建连握手）不走 `call()`，无需间隔

### ⚠️ 分页与对账（2026-10-16 实测新增，P0 级）

- **上游 hasMore 会中段误报 False**：salary_details_list 实测 total=3758，但翻到第 20 页（恰好 2000 条）时
  hasMore=False，21 页起仍有数据；且各月份**乱序混排**，截断砍到哪个月看运气（2026-07 曾被砍半：
  105 人只剩 49 人、70.7 万只剩 34.5 万）。`_items()` 已重写：**按 total 硬拉到底 + 按记录 id 去重**，
  不再信 hasMore；pageSize 上游硬上限 100（请求 200 也只回 100）。
- **对账锚（salary_anchor_check）**：判定前必做——目标月「明细 SUM(实发)」vs「`qxb_payrolls_list` 批次 SUM(实发)」，
  对平才出判定；对不平报告/JSON 亮 ⚠️，结论仅作参考。这是防静默截断的最后防线。
- **payrolls_list 的 payMonth 过滤同样失效**：入参名义要 yyyy-MM，传了返回 0 条（数据里 payMonth 是
  "202607" 无横杠格式）——按月查批次也是全量翻页本地筛。
- **死路接口（勿再试，2026-10-16 实测）**：`qxb_salary_detail_statistics` 任何 subId×月份恒返回空；
  `qxb_salary_files_list` 需 all_projects 全量授权（单项目授权账号 PERMISSION_DENIED）；
  `qxb_subcontractors_list` 仍因 endDate=null 的 zod 故障挂掉。
- **subId 备用来源**：`qxb_teams_list` 目前可用且带 subId（16 班组 8 分包），若恢复 attend_table 路径
  可从这里取 subId，不必等 subcontractors_list 修复。
- 2026-10-16 复核：`qxb_projects_list` 两条通道（SDK 直连 / WorkBuddy 托管）均正常，
  此前"projects_list 被风控"的注记已过时（user_info.projectIds 优先路径保留，无害）。

## 十、月报数据契约（12 类，报告侧消费口径）

月报需要的 12 类数据（对应业务方清单 02-12）在 qxb 源下的取数映射与内部标准结构如下；**凡拉取不到的板块自动降级【待填】并进《待补充数据清单》，不报错**：

| # | 资源（endpoints 键） | 内容 | 内部字段 |
|---|---|---|---|
| 02 | `salary` | 当月工资册 | idcard/name/team/pay/short |
| 03 | `attendance` | 当月考勤册 | name/team/days |
| — | `roster` | 花名册 | id/name/status/entry/labor/team/gong/contract |
| — | `price` | 工种单价表 | gong/low/high |
| 04 | `offline_attendance` | 线下考勤册 | name/team/offline_days |
| 05 | `fund_flow` | 专户进账流水 | date/payer/amount（万元） |
| 06 | `output_value` | 产值数据 | ym/cur_reported/cum_reported/cum_approved |
| 07 | `six_indicator_status` | 六项指标状态 | name/status（绿/黄/红）/note |
| 08 | `complaint_ledger` | 投诉台账 | date/port/people/amount/status/handling |
| 09 | `inspections` | 迎检记录 | date/unit/matter/rectify |
| 10 | `entry_exit` | 进退场与结算 | name/team/entry/leave/contract/settle_amount/settle_status |
| 11 | `next_attendance` | 次月已发生考勤 | name/team/days |
| 12 | `yearly_salary` | 全年逐月工资册 | ym/name/team/gong/days/pay |
| — | `salary_months`/`att_months` | 可用月份列表 | [ym,...] |
| — | `meta` | 项目主数据（封面/概况/判定说明类） | 键=报告占位标签 |
| — | `history` | 历史日薪（可选） | hist/prior |

**聚合下沉原则**：接口只吐明细，汇总由月报生成器算——2.3 三端口表与走势图从 08 台账聚合、2.4 资金表与拨付率图从 05+06 聚合、5.1 欠薪事件从 08 聚合、3.2.1 差异表从 04 关联 03 算、3.1.3 次月预估用 11、3.1.4 全年分布用 12。单一出处，后端不备汇总接口。
- **工种解析三级兜底（2026-10-19）**：花名册 by_id（ID 空间与明细不一致时落空）→ 姓名匹配（明细与花名册唯一可用 join 键，202608 实测 84/88）→ **班组反推**（工资明细不带班组，从考勤键 (name,team) 找回班组名，按 DEFAULT_TEAM_TO_GONG 映射反推；缺工种 52→2）。打标「工种按姓名匹配」/「工种按班组反推（班组名）」可追溯。
- **QXB_ENGINE_JSON 预载**：引擎子进程在生成器内偶发瞬崩（同命令手动重放必成功，根因未明），支持先手动跑 judge.py 产出 JSON 再由生成器直载，绕开子进程。
