#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工人工资数据监理服务月报（全量）Word 生成器
================================
结构基准：用户提供的标准月报成品《云锦华府二期项目2026年8月数据监理服务月报.html》
（按《薪莫愁服务平台运营项目服务报告板块及其计算逻辑（项目端）》思维导图落地），
**字段纪律：标题/指标名/字段名/措辞一律以该 HTML 标准件为准，禁止自行改写**。
pos 导图降为逻辑底层，冲突处 HTML 标准件优先。

覆盖范围：
  封面：标题（{项目名}工人工资数据监理服务月报）+ 信息表九行 + 数据来源声明
  一、项目概况（项目类别/拨付比例、责任划断声明）
  二、项目健康度分析（2.1 五项指标+综合评级 / 2.2 六项制度落实情况 / 2.3 项目欠薪投诉情况（三端口）/
     2.4 项目资金情况 / 2.5 本年度工资发放统计（截至报告期））
  三、工资审核（3.1 主要审核点：3.1.1 日工资离散分析 / 3.1.2 超欠付预警清单 / 3.1.3 次月预估工资 /
     3.1.4 全年日工资水平分布；3.2 次要审核点：3.2.1 线下考勤与黔薪保考勤比对 / 3.2.2 进场合同抽查 /
     3.2.3 离场结算抽查）
  四、人员情况（4.1 当月用工情况）
  五、其他特殊情况（5.1 欠薪事件处理 / 5.2 迎检情况）
  六、建议与反馈（6.1 对建设单位 / 6.2 对施工单位 / 6.3 五级反馈机制启动情况）
  落款：编制/审核/报送 + 报告编号/出具日期

图表（9 张，形态对齐标准件）：图1 六项制度圆角色条 / 图2 三端口投诉组合图（柱合计+三线）/
  图3 资金拨付双轴 / 图4 分账拨付率走势（带政策线）/ 图5 工资双轴 / 图6 当月离散散点（双参考线）/
  图8 全年分布（0-87 分档渐变柱）/ 图7 班组水平条 / 图9 班组折线。

联动契约：
  - 超付/欠付判定全部交给 wage-pay-judge 引擎（判定阈值/金额沿用引擎预警值口径）。
  - 本脚本只调用引擎拿 JSON，再按上述结构落进 docx，不重复实现判定逻辑。
  - 原稿中本应由平台/人工填写、当前无数据源的字段，统一用【待填】就地占位（不单列清单板块）；
    qxb 源经 MCP 工具拉取（项目主数据走 project_meta；拉不到的进《待补充数据清单》）。

用法：
  python build_full_report.py --target 26.8 --data <项目根目录> \
      [--engine <judge.py>] [--source qxb] \
      [--judge-config <判定口径.json>] [--history all|N] \
      [--out <xxx.docx>] \
      [--project-name <项目名> | --project-id <32位ID>] \
      [--attach-data <附件扫描提取.json>]

说明：project_id 默认由 qxb MCP 会话（当前登录用户）解析——单项目自动选用，多项目用
--project-name 指定；附件扫描提取结果经 --attach-data 内联传入（落临时目录，不进项目/ skill 文件）。
"""

import argparse
import glob
import json
import os
import re
import statistics
import subprocess
import sys
import tempfile
from datetime import datetime

# 封面固定信息（2026-10-13 老哥拍板：生成单位/报送对象为定值，编号自动生成）
REPORT_UNIT = "薪莫愁（贵州）数据服务有限公司 · 运营部"
SUBMIT_TO = "建设单位、施工单位、监理例会"
REPORT_NO_PREFIX = "XMC"  # 报告编号前缀，编号格式：XMC-{YYYYMM}-{项目简码}

import openpyxl

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 统一数据源（判定引擎与月报共用，单一来源；见 datasources.py 桥接）
import datasources as DS

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_ENGINE = os.path.normpath(
    os.path.join(HERE, "..", "..", "wage-pay-judge", "scripts", "judge.py")
)
DEFAULT_PRICE = "工种单价预估范围.xlsx"

# ----------------------------- 中文字体 -----------------------------
_CANDIDATES = ["SimHei", "Microsoft YaHei", "SimSun", "Arial Unicode MS",
               "Noto Sans CJK JP", "WenQuanYi Micro Hei", "Source Han Sans SC"]
_AVAIL = {f.name for f in font_manager.fontManager.ttflist}
_CHOSEN = next((c for c in _CANDIDATES if c in _AVAIL), None)
if _CHOSEN:
    plt.rcParams["font.sans-serif"] = [_CHOSEN]
plt.rcParams["axes.unicode_minus"] = False

# 颜色
C_GREEN = "#2E8B57"
C_YELLOW = "#E1A100"
C_RED = "#C0392B"
C_GRAY = "#9AA0A6"
C_BLUE = "#2F6FB0"


# ----------------------------- 配置 / 占位 -----------------------------
# 待补充数据项 → 补充途径提示（qxb 源按平台工具/接口端点；未列出的走通用提示）
_PENDING_HINTS = {
    "投诉走势图数据": "系统接口 complaint_ledger（投诉台账）未配置或无数据",
    "历史累计投诉(数量/金额)": "系统接口 complaint_ledger（投诉台账）",
    "当月投诉(数量/金额)": "系统接口 complaint_ledger（投诉台账）",
    "已解决投诉数量": "系统接口 complaint_ledger（投诉台账）",
    "线下考勤对比": "系统接口 offline_attendance（线下考勤册）未配置或无数据",
    "报审产值": "系统接口 output_value（产值数据）",
    "审定产值": "系统接口 output_value（产值数据）",
    "拨付工程款": "系统接口暂未提供该口径，需项目主数据补充",
    "专户进账金额": "系统接口 fund_flow（专户进账流水）",
    "合规分账进账": "系统接口 fund_flow 剔除非建设单位打款后计算",
    "专户剩余资金": "项目主数据（专户余额）",
    "专户够发判定": "项目主数据（专户余额与次月预估比对）",
    "分账拨付率": "系统接口 fund_flow + output_value 聚合计算",
    "保证金状态": "系统接口 six_indicator_status（六项指标状态）",
    "分账拨付状态": "系统接口 six_indicator_status（六项指标状态）",
    "离场结算状态": "系统接口 six_indicator_status（六项指标状态）",
    "合同抽查份数": "合同抽查台账（系统接口未提供，需项目补充）",
    "低工价合同提示": "合同工价核对记录（系统接口未提供，需项目补充）",
    "离场结算抽查份数": "离场结算抽查台账（系统接口未提供，需项目补充）",
    "离场结算不合规数": "系统接口 entry_exit（离场结算状态）",
    "超30日未结算离场": "系统接口 entry_exit（离场结算状态）",
    "预估月份发放金额(缺工种)": "工种解析失败人员的单价表补齐",
    "欠薪事件处置说明": "投诉处置过程记录（系统接口 complaint_ledger.handling）",
    "五级反馈启动情况": "五级反馈机制台账（系统接口未提供，需项目补充）",
}


class Placeholders:
    """集中管理【待填】项。规则（2026-10-12 老哥要求）：凡本次未能拉取到的数据，
    一律落【待填】就地占位，并在报告末尾汇总成《待补充数据清单》（板块/数据项/补充途径）。

    方案甲：need() 优先从 meta（MCP 接口/project_meta 回填 + 内联 attach_data 兜底）取值，
    缺项才落【待填】——qxb 源经平台工具取全后报告零占位、无附录；兜底链：
    MCP 工具 → 附件扫描（agent 读图内联）→ 待填。"""
    def __init__(self):
        self.items = []  # [{"section","item","hint"}]
        self.meta = {}
        self.source_kind = "file"  # "db" 或 "file"：决定补充途径提示口径

    def set_meta(self, d):
        self.meta = d or {}

    def _hint(self, item):
        if item in _PENDING_HINTS:
            return _PENDING_HINTS[item]
        return f"系统接口/项目主数据未回填「{item}」（或附件扫描亦未取到）"

    def need(self, section, item):
        v = self.meta.get(item)
        if v not in (None, "", [], {}):
            return v
        if not any(x["section"] == section and x["item"] == item for x in self.items):
            self.items.append({"section": section, "item": item, "hint": self._hint(item)})
        return "【待填】"


PH = Placeholders()


def load_config():
    """报告头字段统一从 PH.meta（MCP project_meta + 内联 attach_data）取值；
    取不到一律【待填】。不再读取任何落盘的项目配置文件（避免跨用户/项目串数据）。"""
    out = {
        "项目名称": PH.need("封面", "项目名称"),
        "建设单位": PH.need("封面", "建设单位"),
        "施工单位": PH.need("封面", "施工单位"),
        "监理单位": PH.need("封面", "监理单位"),
        "现场负责人": PH.need("项目概况", "现场负责人"),
        "项目经理": PH.need("项目概况", "项目经理"),
        "劳资专管员": PH.need("项目概况", "劳资专管员"),
        "开工日期": PH.need("项目概况", "实际开工时间"),
        "预计完工日期": PH.need("项目概况", "预计完工时间"),
    }
    return out


# ----------------------------- 数据读取 -----------------------------
def find_header_row(ws, must_contain):
    for r in range(1, min(ws.max_row, 8) + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any((v or "").strip() in must_contain for v in vals if v):
            return r, vals
    return 1, [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]


def col_index(header_vals, name):
    for i, v in enumerate(header_vals):
        if v and str(v).strip() == name:
            return i + 1
    return None


def load_roster(ds):
    """花名册（统一数据源，单一来源）。返回 (persons, by_id, by_nt)。"""
    return ds.roster()


def load_price_table(ds):
    """工种单价区间（统一数据源，单一来源）。"""
    return ds.price_table()


def month_list_from_files(ds):
    """返回 (工资月份集合, 考勤月份集合)，统一从数据源读取。"""
    return set(ds.list_salary_months()), set(ds.list_att_months())


def load_salary_month(ds, ym):
    """返回 list[{id,name,pay,team}]。"""
    agg = ds.salary_agg(ym)
    rows = []
    for r in agg.values():
        rows.append({
            "id": r.get("idcard", "") or "",
            "name": r.get("name", ""),
            "pay": r.get("pay", 0.0),
            "team": r.get("team", ""),
        })
    return rows


# 快速模式（QXB_FAST_MODE=1，2026-10-19 应对平台限流）：考勤只拉目标月，历史月不取
_FAST_ATT_MONTHS = None  # None=全量口径；set=仅这些月份拉考勤（由 build 按环境变量设置）
_OFFLINE_ATT_PATH = None  # 3.2.1 线下考勤册 xlsx（工资册即考勤册：工作天数列=线下出勤）


def load_attendance_month(ds, ym):
    """返回 list[{name,team,days}]，统一从数据源读取。"""
    if _FAST_ATT_MONTHS is not None and _norm_ym(ym) not in _FAST_ATT_MONTHS:
        return []
    return ds.attendance_target(ym)


def load_offline_attendance(path):
    """2026-10-23：3.2.1 线下考勤册=工资册 xlsx。表头自适应——扫前 8 行找含"姓名"+"工作天数"
    的表头行，按列名定位（2026.6 表头在第 4 行、2026.7 在第 3 行且多"欠薪宝考勤天数"列，
    只认"工作天数"为线下出勤口径；"欠薪宝考勤天数"是平台侧回写值不作线下口径）。
    工作天数即线下出勤天数；加班天数不计入考勤比对（另行发放口径）。尾部"合计"行姓名为空自动跳过。
    返回 list[{name, team, offline_days}]，与数据源 offline_attendance 契约对齐并多带 name。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    head = list(ws.iter_rows(min_row=1, max_row=8, values_only=True))
    hdr_i = col_team = col_name = col_days = None
    for i, r in enumerate(head):
        cells = [str(v).strip() if v is not None else "" for v in (list(r) + [None] * 24)[:24]]
        if "姓名" in cells and any("工作天数" in c for c in cells):
            hdr_i = i
            col_name = cells.index("姓名")
            col_days = next(j for j, c in enumerate(cells) if "工作天数" in c)
            col_team = cells.index("班组") if "班组" in cells else 1
            break
    if hdr_i is None:
        raise ValueError(f"{path}：前 8 行未找到含 姓名/班组/工作天数 的表头行")
    out, bad = [], []
    for r in ws.iter_rows(min_row=hdr_i + 2, values_only=True):
        cells = (list(r) + [None] * (col_days + 2))[: max(col_days + 2, 14)]
        team = cells[col_team] if col_team < len(cells) else None
        name = cells[col_name] if col_name < len(cells) else None
        days = cells[col_days] if col_days < len(cells) else None
        if name is None or not str(name).strip():
            continue  # 合计行/空行
        try:
            d0 = float(days)
        except (TypeError, ValueError):
            bad.append(str(name).strip())
            continue
        out.append({"name": str(name).strip(), "team": (str(team).strip() if team else ""),
                    "offline_days": d0})
    if bad:
        print(f"[offline-att] 工作天数非数值、已跳过 {len(bad)} 人: {bad[:8]}", flush=True)
    return out


# ----------------------------- 引擎联动 -----------------------------
def run_engine(engine, target, data, source, judge_config, history, project_id=None, project_name=None):
    # 2026-10-19：引擎子进程在生成器内反复瞬崩（同命令手动重放必成功，疑与子进程环境相关）。
    # 支持预生成直载：先手动跑 judge.py 产出 JSON，再以 QXB_ENGINE_JSON 指给它，跳过子进程。
    _pre = os.environ.get("QXB_ENGINE_JSON")
    if _pre and os.path.exists(_pre):
        print(f"[step] 引擎结果预生成，直接载入 {_pre}", flush=True)
        with open(_pre, encoding="utf-8") as f:
            return json.load(f)
    tmp = tempfile.NamedTemporaryFile(
        prefix=f"judge_{target}_", suffix=".json", delete=False, mode="w", encoding="utf-8"
    )
    tmp.close()
    cmd = [sys.executable, engine, "--target", target, "--data", data,
           "--source", source, "--json", tmp.name, "--history", history]
    if judge_config:
        cmd += ["--config", judge_config]
    if project_id:
        cmd += ["--project-id", project_id]
    if project_name:
        cmd += ["--project-name", project_name]
    import time as _time
    last_err = None
    for attempt in (1, 2):  # 2026-10-19：引擎子进程偶发瞬态崩溃，自动重试一次
        try:
            # stderr 不再吞掉（继承父进程日志），否则引擎崩溃原因永远不可见
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL)
            last_err = None
            break
        except subprocess.CalledProcessError as e:
            last_err = e
            print(f"[step] 引擎子进程第 {attempt} 次失败（exit {e.returncode}）"
                  + ("，10 秒后重试" if attempt == 1 else "，放弃"), flush=True)
            if attempt == 1:
                _time.sleep(10)
    if last_err is not None:
        raise last_err
    with open(tmp.name, encoding="utf-8") as f:
        payload = json.load(f)
    os.unlink(tmp.name)
    return payload


# ----------------------------- 业务计算 -----------------------------
def _roster_gong(roster_by_nt, name, team):
    """从花名册 by_nt 取工种；兼容标准布局（dict）与中医药冲突布局（list）。"""
    rec = roster_by_nt.get(name + "|" + team)
    if not rec:
        return ""
    if isinstance(rec, list):
        for r in rec:
            if r.get("gong"):
                return r.get("gong")
        return rec[0].get("gong", "") if rec else ""
    return rec.get("gong", "")


def parse_ym_month(s):
    m = re.search(r"(\d{4})[-/]?(\d{1,2})", s or "")
    if m:
        return f"{int(m.group(1)):04d}{int(m.group(2)):02d}"
    return None


def compute_monthly(ds, salary_months, att_months, price_table, roster_by_nt):
    """返回按月份排序的聚合 list[dict]。"""
    all_months = sorted(salary_months | att_months)
    result = []
    for ym in all_months:
        sal = load_salary_month(ds, ym)
        att = load_attendance_month(ds, ym)
        pay_sum = sum(r["pay"] for r in sal)
        sal_persons = len({(r["id"] or r["name"] + "|" + r["team"]) for r in sal})
        days_sum = sum(r["days"] for r in att)
        att_persons = len({(r["name"], r["team"]) for r in att})
        # 预估发放金额（仅考勤无工资时）= Σ avg(下限,上限)×考勤（缺工种者不计入）
        est = None
        if not sal and att:
            est = 0.0
            for a in att:
                gong = _roster_gong(roster_by_nt, a["name"], a["team"])
                if gong and gong in price_table:
                    lo, hi = price_table[gong]
                    est += (lo + hi) / 2 * a["days"]
            if est == 0.0:
                est = None
        result.append({
            "ym": ym, "has_salary": bool(sal), "has_att": bool(att),
            "pay_sum": pay_sum, "sal_persons": sal_persons,
            "days_sum": days_sum, "att_persons": att_persons, "est": est,
        })
    return result


def compute_yearly_daily(ds, salary_months, att_months, roster_by_id=None, year=None):
    """返回 list[(daily_wage, team)]。
    2026-10-19 老哥拍板改口径：统计粒度从"人员-月"改为"人"——全年口径为
    年实发合计 ÷ 年考勤合计（同一人各月工资与考勤先汇总再相除），分布图每根柱
    数的是人而不是人次；统计范围限报告年度（year=报告年），覆盖该年全部有工资+考勤的月份。
    2026-10-16 修复：qxb 工资明细不带班组（team=""），(姓名,班组) 匹配考勤必然落空 →
    按姓名兜底匹配（考勤侧 by-name 已跨班组合并）。"""
    out = []
    agg = {}  # name -> [pay_sum, days_sum]
    for ym in sorted(salary_months):
        if year and not str(ym).startswith(str(year)):
            continue
        sal = load_salary_month(ds, ym)
        att = load_attendance_month(ds, ym)
        att_map = {(a["name"], a["team"]): a["days"] for a in att}
        att_by_name = {}
        for a in att:
            att_by_name[a["name"]] = att_by_name.get(a["name"], 0.0) + a["days"]
        for s in sal:
            days = att_map.get((s["name"], s["team"]))
            if days is None:
                days = att_by_name.get(s["name"])
            if days and days > 0 and s["pay"] > 0:
                rec = agg.setdefault(s["name"], [0.0, 0.0])
                rec[0] += s["pay"]
                rec[1] += days
    for name, (pay_sum, days_sum) in agg.items():
        if days_sum > 0 and pay_sum > 0:
            out.append((pay_sum / days_sum, ""))
    return out


def compute_timeline(persons, salary_months, att_months, ds):
    """为每人推断 进场月/离场月，用于全年进退场图。"""
    # 先汇总每人出现在哪些月份（工资或考勤）
    appear = {}
    for ym in salary_months | att_months:
        sal = load_salary_month(ds, ym)
        att = load_attendance_month(ds, ym)
        for s in sal:
            k = s["id"] or s["name"] + "|" + s["team"]
            appear.setdefault(k, set()).add(ym)
        for a in att:
            k = a["name"] + "|" + a["team"]
            appear.setdefault(k, set()).add(ym)
    recs = []
    for p in persons:
        k = p["id"] if p["id"] else p["name"] + "|" + p["team"]
        months = sorted(appear.get(k, set()))
        entry = parse_ym_month(p["entry"]) or (months[0] if months else None)
        if p["status"] == "离场" and months:
            leave = months[-1]
        else:
            leave = None
        recs.append({"name": p["name"], "team": p["team"], "entry": entry, "leave": leave})
    return recs


# ----------------------------- 图表 -----------------------------
def ensure_assets(data, target):
    d = os.path.join(data, "_report_assets", target)
    os.makedirs(d, exist_ok=True)
    return d


def fig_to_file(fig, path):
    fig.tight_layout()
    fig.savefig(path, dpi=130, bbox_inches="tight")
    plt.close(fig)


def chart_salary_trend(monthly, assets, target):
    """图5：本年度工资发放金额与人数（柱=当月实发工资，左轴；方块折线=发放人数，右轴）。
    仅绘制报告期所在年份、且不晚于报告月的月份（2026-10-19 老哥拍板：只统计到报告月）。"""
    fig, ax1 = plt.subplots(figsize=(9.5, 4.2))
    _ry, _rm = _split_ym(target)
    ym, pay, people = [], [], []
    for m in monthly:
        _y, _m = _split_ym(m["ym"])
        if _y != _ry or (_y, _m) > (_ry, _rm):
            continue
        _y, _m = _split_ym(m["ym"])
        ym.append(f"{_y}-{_m:02d}")
        pay.append(m["pay_sum"] / 10000 if m["has_salary"] else 0)
        people.append(m["sal_persons"] if m["has_salary"] else (m["att_persons"] if m["has_att"] else 0))
    x = range(len(ym))
    ax1.bar(x, pay, color=C_BLUE, alpha=0.9, label="当月实发工资（万元）", width=0.6)
    for xi, pv in zip(x, pay):
        if pv:
            ax1.text(xi, pv + max(pay) * 0.015, f"{pv:.2f}", ha="center", fontsize=8, color="#333333")
    ax1.set_ylabel("实发工资（万元）")
    ax1.set_xticks(list(x))
    ax1.set_xticklabels(ym, rotation=45)
    ax1.set_ylim(0, max(pay) * 1.15 if any(pay) else 1)
    ax2 = ax1.twinx()
    ax2.plot(x, people, color=C_RED, marker="s", markersize=5, linewidth=1.8, label="发放人数（人）")
    ax2.set_ylabel("发放人数（人）", color=C_RED)
    ax2.tick_params(axis="y", labelcolor=C_RED)
    ax1.set_title("本年度工资发放金额与人数")
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(h1 + h2, l1 + l2, fontsize=8, loc="upper left")
    path = os.path.join(assets, "salary_trend.png")
    fig_to_file(fig, path)
    return path


def chart_daily_disp(persons, assets, target):
    """图6：当月全项目人员日工资离散分布（x=当月考勤天数，y=当月日工资；
    区间内=蓝、高于预警上限=黄、低于预警下限=红；双参考线取全员预警值中位数）。"""
    pts = [p for p in persons if p.get("days") and p.get("daily")]
    if not pts:
        return None
    import statistics as st
    lows = [p["under_th"] for p in pts if p.get("under_th")]
    highs = [p["over_th"] for p in pts if p.get("over_th")]
    lo_ref = st.median(lows) if lows else 0
    hi_ref = st.median(highs) if highs else 0
    fig, ax = plt.subplots(figsize=(9.5, 4.6))
    for p in pts:
        if p.get("over_th") and p["daily"] > p["over_th"]:
            c = C_YELLOW
        elif p.get("under_th") and p["daily"] < p["under_th"]:
            c = C_RED
        else:
            c = C_BLUE
        ax.scatter(p["days"], p["daily"], s=16, color=c, alpha=0.65, edgecolors="none")
    if hi_ref:
        ax.axhline(hi_ref, color=C_RED, linestyle="--", linewidth=1.2)
        ax.text(ax.get_xlim()[1], hi_ref, " 项目适用上限区间参考线", fontsize=8.5,
                color=C_RED, va="bottom", ha="right")
    if lo_ref:
        ax.axhline(lo_ref, color=C_GREEN, linestyle="--", linewidth=1.2)
        ax.text(ax.get_xlim()[1], lo_ref, " 项目适用下限区间参考线", fontsize=8.5,
                color=C_GREEN, va="top", ha="right")
    ax.set_xlabel("当月黔薪保考勤天数（天）")
    ax.set_ylabel("当月日工资（元/天）")
    ax.set_title(f"当月全项目人员日工资离散分布（n={len(pts)}）")
    path = os.path.join(assets, "daily_disp.png")
    fig_to_file(fig, path)
    return path


def chart_complaint_trend(complaints, assets):
    """图2：三端口欠薪投诉走势（总表=三端口合计柱状；分表=三端口各自折线，一张组合图）。
    complaints: {ym: {"政府端": n, "项目端": n, "薪莫愁端": n}}，无数据返回 None。"""
    if not complaints:
        return None
    months = sorted(complaints.keys())
    ports = ["政府端", "项目端", "薪莫愁端"]
    port_color = {"政府端": C_BLUE, "项目端": C_YELLOW, "薪莫愁端": C_RED}
    totals = [sum(complaints[m].values()) for m in months]
    fig, ax = plt.subplots(figsize=(9.5, 4.2))
    x = list(range(len(months)))
    ax.bar(x, totals, width=0.55, color="#B8D4EA", label="三端口合计", zorder=1)
    for xi, tv in zip(x, totals):
        if tv:
            ax.text(xi, tv + 0.06, str(tv), ha="center", fontsize=9, color="#555555")
    for p in ports:
        vals = [complaints[m].get(p, 0) for m in months]
        ax.plot(x, vals, marker="o", markersize=5, linewidth=1.8,
                color=port_color[p], label=p, zorder=3)
    ax.set_xticks(x)
    ax.set_xticklabels([m[2:].replace("-", "/") for m in months])
    ax.set_ylabel("投诉件数（件）")
    ax.set_xlabel("月份")
    ax.set_title("三端口欠薪投诉走势（柱=三端口合计，线=各端口）")
    ax.legend(fontsize=8, ncol=4, loc="upper left")
    ax.set_ylim(bottom=0)
    path = os.path.join(assets, "complaint_trend.png")
    fig_to_file(fig, path)
    return path


def _six_status_text(status, note):
    """六项制度图条内文字：达标显示"正常"，否则显示核心原因（取第一个分句）。"""
    if status == "达标":
        return "● 正常"
    if status == "待填":
        return "● 待平台回填"
    short = str(note or "").split("；")[0].split("（")[0].strip()
    return f"● {short}" if short else f"● {status}"


def chart_fund(fund_history, assets):
    """图3：项目累计资金拨付与专户进账（分组柱=报审/审定/拨付工程款，左轴；
    红线=专户累计进账，右轴）。fund_history: {ym: {报审, 审定, 拨付, 专户}}（万元）。"""
    if not fund_history:
        return None
    months = sorted(fund_history.keys())
    rep = [fund_history[m].get("报审", 0) for m in months]
    appr = [fund_history[m].get("审定", 0) for m in months]
    pay = [fund_history[m].get("拨付", 0) for m in months]
    acc = [fund_history[m].get("专户", 0) for m in months]
    fig, ax = plt.subplots(figsize=(9.5, 4.2))
    x = list(range(len(months)))
    w = 0.26
    ax.bar([i - w for i in x], rep, width=w, color="#B8D4EA", label="累计报审产值")
    ax.bar(x, appr, width=w, color="#6FA3D8", label="累计审定产值")
    ax.bar([i + w for i in x], pay, width=w, color="#2F6FB0", label="累计拨付工程款")
    ax.set_ylabel("产值 / 拨付（万元）")
    ax.set_xticks(x)
    ax.set_xticklabels([m[2:].replace("-", "/") for m in months])
    ax2 = ax.twinx()
    ax2.plot(x, acc, color=C_RED, marker="o", markersize=5, linewidth=1.8, label="专户累计进账")
    ax2.set_ylabel("专户进账（万元）", color=C_RED)
    ax2.tick_params(axis="y", labelcolor=C_RED)
    ax.set_title("项目累计资金拨付与专户进账情况")
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, fontsize=8, ncol=4, loc="upper left")
    path = os.path.join(assets, "fund_flow.png")
    fig_to_file(fig, path)
    return path


def chart_fund_rate(fund_history, assets, policy=0.20):
    """图4：分账拨付率走势（专户累计进账 ÷ 累计审定产值），带政策要求参考线。"""
    if not fund_history:
        return None
    months = sorted(fund_history.keys())
    rates = []
    for m in months:
        appr = fund_history[m].get("审定", 0) or 0
        acc = fund_history[m].get("专户", 0) or 0
        rates.append(acc / appr * 100 if appr else 0)
    fig, ax = plt.subplots(figsize=(9.5, 3.8))
    x = list(range(len(months)))
    ax.plot(x, rates, color=C_BLUE, marker="o", markersize=6, linewidth=2)
    for xi, rv in zip(x, rates):
        ax.text(xi, rv + 0.06, f"{rv:.2f}%", ha="center", fontsize=8.5, color="#333333")
    ax.axhline(policy * 100, color=C_RED, linestyle="--", linewidth=1.4)
    ax.text(len(months) - 1, policy * 100 + 0.05, f"政策要求 {policy*100:.0f}%",
            ha="right", fontsize=9, color=C_RED)
    ax.set_xticks(x)
    ax.set_xticklabels([m[2:].replace("-", "/") for m in months])
    ax.set_ylabel("分账拨付率")
    ax.set_xlabel("月份")
    ax.set_title("分账拨付率走势（专户累计进账 ÷ 累计审定产值）")
    ax.set_ylim(bottom=min(rates + [policy * 100]) - 0.5)
    path = os.path.join(assets, "fund_rate.png")
    fig_to_file(fig, path)
    return path


def chart_six_indicators(indicators, assets):
    """图1：六项制度落实达标情况（横向圆角色条，绿=达标/黄=限期整改/红=严重违规）。"""
    fig, ax = plt.subplots(figsize=(9.5, 4.2))
    n = len(indicators)
    # 2026-10-22：补"严重违规"红色键，"不达标"保留为红色兜底（历史词防御）
    status_color = {"达标": C_GREEN, "提示": C_YELLOW, "限期整改": C_YELLOW, "预警": C_YELLOW,
                    "警告": C_RED, "严重违规": C_RED, "不达标": C_RED, "待填": C_GRAY}
    for idx, i in enumerate(indicators):
        y = n - 1 - idx
        color = status_color.get(i["status"], C_GRAY)
        ax.add_patch(FancyBboxPatch((0.02, y - 0.28), 1.5, 0.56,
                                    boxstyle="round,pad=0.02,rounding_size=0.08",
                                    linewidth=0, facecolor=color))
        ax.text(0.08, y, _six_status_text(i["status"], i["note"]),
                va="center", ha="left", fontsize=9.5, color="white")
        ax.text(1.68, y, i["name"], va="center", ha="left", fontsize=10.5, color="#333333")
    ax.set_xlim(0, 3.4)
    ax.set_ylim(-0.6, n - 0.4)
    ax.axis("off")
    ax.set_title("六项制度落实达标情况（绿=达标 / 黄=限期整改 / 红=严重违规）",
                 fontsize=11, pad=12)
    path = os.path.join(assets, "six_indicators.png")
    fig_to_file(fig, path)
    return path


def chart_daily_dist(yearly, assets):
    """图8：全项目人员全年日工资水平分布（年实发÷年考勤）。
    分档：0-87 / 88-200 / 201-300 / 301-400 / 401-500 / >500，柱顶标人数。"""
    bins_def = [("0-87", lambda v: v <= 87), ("88-200", lambda v: 88 <= v <= 200),
                ("201-300", lambda v: 201 <= v <= 300), ("301-400", lambda v: 301 <= v <= 400),
                ("401-500", lambda v: 401 <= v <= 500), (">500", lambda v: v > 500)]
    vals = [d for d, _ in yearly if d is not None]
    if not vals:  # 2026-10-16 防空：无有效日工资样本时返回 None，报告侧跳过，不画零柱空图
        return None
    counts = [sum(1 for v in vals if f(v)) for _, f in bins_def]
    colors = ["#6FA3D8", "#2F6FB0", "#E8B4BC", "#D98880", "#C0392B", "#7D3C98"]
    fig, ax = plt.subplots(figsize=(9, 4))
    x = range(len(bins_def))
    ax.bar(x, counts, color=colors, width=0.62)
    for xi, cv in zip(x, counts):
        if cv:
            ax.text(xi, cv + max(counts) * 0.02, f"{cv}人", ha="center", fontsize=10)
    ax.set_xticks(list(x))
    ax.set_xticklabels([b for b, _ in bins_def])
    ax.set_xlabel("全年日工资区间（元/天）")
    ax.set_ylabel("人数（人）")
    ax.set_title("全项目人员全年日工资水平分布（年实发 ÷ 年考勤）")
    ax.set_ylim(0, max(counts) * 1.12 if any(counts) else 1)
    path = os.path.join(assets, "daily_dist_all.png")
    fig_to_file(fig, path)
    return path


def chart_team_workforce(att_target, assets, target):
    """图7：当月各班组用工人数（水平条形，按人数降序，条端标人数）。"""
    from collections import defaultdict
    cnt = defaultdict(int)
    for a in att_target:
        cnt[a["team"]] += 1
    teams = sorted(cnt.keys(), key=lambda t: cnt[t])
    fig, ax = plt.subplots(figsize=(9, 4.6))
    vals = [cnt[t] for t in teams]
    ax.barh(teams, vals, color=C_BLUE, height=0.62)
    for i, v in enumerate(vals):
        ax.text(v + max(vals) * 0.01, i, f"{v}人", va="center", fontsize=9.5)
    ax.set_xlabel("当月用工人数（人）")
    ax.set_xlim(0, max(vals) * 1.12 if vals else 1)
    ax.set_title(f"当月各班组用工人数（{target}）")
    path = os.path.join(assets, "team_workforce.png")
    fig_to_file(fig, path)
    return path


def chart_team_attendance(att_target, assets, target):
    """图9：当月各班组人员考勤总天数（折线+数据点标签，按考勤降序）。"""
    from collections import defaultdict
    days = defaultdict(float)
    for a in att_target:
        days[a["team"]] += a["days"]
    teams = sorted(days.keys(), key=lambda t: -days[t])
    if not teams:
        return None
    fig, ax = plt.subplots(figsize=(9.5, 4.2))
    x = list(range(len(teams)))
    vals = [days[t] for t in teams]
    ax.plot(x, vals, color=C_BLUE, marker="o", markersize=7, linewidth=2)
    for xi, v in zip(x, vals):
        ax.text(xi, v + max(vals) * 0.02, f"{v:.0f}", ha="center", fontsize=9.5)
    ax.set_xticks(x)
    ax.set_xticklabels(teams, rotation=30, ha="right")
    ax.set_ylabel("当月考勤总天数（天）")
    ax.set_title(f"当月各班组人员考勤天数（{target}）")
    ax.set_ylim(0, max(vals) * 1.12)
    path = os.path.join(assets, "team_attendance.png")
    fig_to_file(fig, path)
    return path


# ----------------------------- Word 渲染 -----------------------------
def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def style_header_row(row, hex_color="2F6FB0"):
    for cell in row.cells:
        set_cell_bg(cell, hex_color)
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(9)


def add_table(doc, headers, rows, col_widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = str(h)
    style_header_row(t.rows[0])
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = "" if v is None else str(v)
            for p in cells[i].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in t.rows:
                row.cells[i].width = Inches(w)
    return t


def h(doc, text, level):
    p = doc.add_heading(text, level=level)
    return p


def para(doc, text, bold=False, color=None, size=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if size:
        run.font.size = Pt(size)
    return p


def fmt_money(v):
    if v is None:
        return "—"
    return f"¥{v:,.2f}"


def fmt_num(v):
    if v is None:
        return "—"
    return f"{v:,.2f}"


# ----------------------------- 主组装 -----------------------------
def load_judge_cfg(path):
    if path and os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def _split_ym(ym):
    """统一解析报告月：支持标准 'YYYYMM' 与中医药 'YY.M' 两种写法，返回 (year, month)。"""
    if "." in str(ym):
        y, m = (int(x) for x in str(ym).split("."))
        if y < 100:
            y += 2000
        return y, m
    s = str(ym)
    return int(s[:4]), int(s[4:6])


def _next_ym(ym):
    """返回 (年, 月) 的下一月。"""
    y, m = _split_ym(ym)
    if m == 12:
        return y + 1, 1
    return y, m + 1


def _step(msg):
    """进度打点（2026-10-16 加）：长任务全程可见，卡哪一步一目了然。"""
    print(f"[step {datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def build(target, data, engine, source, judge_config, history, config_path=None,
          out=None, project_id=None, project_name=None, attach_extracted=None):
    jcfg = load_judge_cfg(judge_config)
    # 待补充清单的提示口径：qxb=平台接口未回填（外部数据源仅 qxb，2026-10-13 起）
    PH.source_kind = "db" if source in ("qxb", "mcp") else "file"
    # 运行时 project_id / project_name 注入顶层配置，供 make_datasource 的 qxb 分支解析项目
    if project_id:
        jcfg["project_id"] = project_id
    if project_name:
        jcfg["project_name"] = project_name
    _step("创建数据源（qxb MCP）")
    global _FAST_ATT_MONTHS
    # 2026-10-22（老哥拍板）：极速模式默认开启，QXB_FAST_MODE=0 显式关闭才走全量模式
    if os.environ.get("QXB_FAST_MODE", "1") != "0":
        # 2026-10-19 老哥拍板：考勤放行报告年 1 月~报告月（2.5 天数列 + 3.1.4 全年分布都要逐月数据），
        # 报告年之前的月份仍不拉（历史日薪只算当年）
        _ty_f, _tm_f = _split_ym(target)
        _FAST_ATT_MONTHS = {f"{_ty_f}{m:02d}" for m in range(1, _tm_f + 1)}
        # 跨进程磁盘缓存（引擎复用生成器已拉数据，避免同量数据拉两遍触发限流）
        # 2026-10-22（老哥拍板）：快速模式不再强制 QXB_SKIP_MATRIX——考勤坚持矩阵
        # 优先（限流/验证码隔 2 分钟退避试探），跳过矩阵会整月回退慢速打卡流水
        if os.environ.get("QXB_SKIP_MATRIX") == "1":
            os.environ.pop("QXB_SKIP_MATRIX", None)
        if not os.environ.get("QXB_CACHE_FILE"):
            # 2026-10-21：缓存默认落 <data>/qxb_cache/（随项目在 E 盘）——原 Temp 路径
            # 会被系统清理且有 C 盘空间压力；已有缓存的月份直接复用不重拉
            # 2026-10-23：单缓存文件 qxb_cache.json（跟项目走、不跟报告月走）
            _cache_dir = os.path.join(data, "qxb_cache")
            os.makedirs(_cache_dir, exist_ok=True)
            os.environ["QXB_CACHE_FILE"] = os.path.join(_cache_dir, "qxb_cache.json")
        print(f"[step] 极速模式：考勤拉报告年 1~{_tm_f} 月（{sorted(_FAST_ATT_MONTHS)}），"
              f"缓存={os.environ['QXB_CACHE_FILE']}", flush=True)
    # 2026-10-22（老哥拍板）：磁盘缓存不再绑定快速模式——拉到的数据一律缓存进
    # 工作空间 <data>/qxb_cache/（工资全量/花名册/逐月考勤），重启与引擎子进程直接复用
    if not os.environ.get("QXB_CACHE_FILE"):
        _cache_dir = os.path.join(data, "qxb_cache")
        os.makedirs(_cache_dir, exist_ok=True)
        os.environ["QXB_CACHE_FILE"] = os.path.join(_cache_dir, "qxb_cache.json")
        print(f"[step] qxb 数据缓存（工作空间）={os.environ['QXB_CACHE_FILE']}", flush=True)
    # 2026-10-23（老哥拍板）：单缓存文件——缓存跟项目走、不跟报告月走。月报生成只补拉
    # 缺失月份，历史数据全吃缓存。若单文件不存在而工作空间还有旧的按月缓存
    # （qxb_cache_YYYYMM.json），自动合并迁移：kaoqin 按月收编（须含 by_nt 才算有效）、
    # salary/contract_persons 取条数最多的一份；迁移完成后旧文件改名 .bak 退役
    # （不再参与合并、可人工恢复）。
    _cache_path = os.environ["QXB_CACHE_FILE"]
    _cache_dir2 = os.path.dirname(_cache_path) or "."
    _legacy = [os.path.join(_cache_dir2, f) for f in sorted(os.listdir(_cache_dir2))
               if f.startswith("qxb_cache_") and f.endswith(".json")]
    if _legacy and not os.path.exists(_cache_path):
        _merged = {}
        for _p in _legacy:
            try:
                with open(_p, encoding="utf-8") as f:
                    _src = json.load(f)
            except Exception:
                continue
            if not isinstance(_src, dict):
                continue
            for _ym, _e in (_src.get("kaoqin") or {}).items():
                if isinstance(_e, dict) and _e.get("by_nt"):
                    _merged.setdefault("kaoqin", {}).setdefault(_ym, _e)
            for _k in ("salary", "contract_persons"):
                _s = _src.get(_k)
                _t = _merged.get(_k)
                if isinstance(_s, list) and (not isinstance(_t, list) or len(_s) > len(_t)):
                    _merged[_k] = _s
        if _merged:
            with open(_cache_path, "w", encoding="utf-8") as f:
                json.dump(_merged, f, ensure_ascii=False)
            for _p in _legacy:
                try:
                    os.replace(_p, _p + ".bak")
                except OSError:
                    pass
            print(f"[step] 缓存迁移：{len(_legacy)} 个按月缓存并入单文件 qxb_cache.json"
                  f"（考勤 {len(_merged.get('kaoqin') or {})} 月、工资 "
                  f"{len(_merged.get('salary') or [])} 条、花名册 "
                  f"{len(_merged.get('contract_persons') or [])} 条），旧文件已改名 .bak", flush=True)
    ds = DS.make_datasource(source, data, jcfg)
    meta = ds.project_meta() or {}
    # 2026-10-21：封面三方/管理人员主来源=qxb_project_get（平台五方接口）；附件扫描降为
    # 兜底，只补 MCP 没取到的键（如维权电话等告示牌专属字段），不再覆盖工具值
    _att_json = os.path.join(data, "_attach_extracted.json")
    if os.path.exists(_att_json):
        try:
            with open(_att_json, encoding="utf-8") as _f:
                _att = json.load(_f)
            if isinstance(_att, dict):
                _fill = {k: v for k, v in _att.items() if v and not meta.get(k)}
                meta.update(_fill)
                _step(f"附件提取兜底回填 {len(_fill)} 项（工具值优先，附件仅补缺）")
        except Exception as _e:
            print(f"[step] 附件提取文件读取失败：{_e}", flush=True)
    _step("项目主数据就绪")
    # 2026-10-19 新工具：合规快照平铺进 meta（PH.need 按条目名直取）
    _comp = ds.compliance() if hasattr(ds, "compliance") else None
    if _comp:
        _IND = _comp.get("indicators") or {}
        _ST = {"met": "达标", "warning": "预警", "unmet": "未落实"}

        def _flat(key):
            v = _IND.get(key) or {}
            s = _ST.get(v.get("status"), "")
            note = v.get("note") or ""
            if s and note:
                return f"{note}（平台状态：{s}）"
            return note or s or None

        for _k, _item in (("保证金存储", "保证金状态"), ("人工费分账拨付", "分账拨付状态"),
                          ("离场结算", "离场结算状态")):
            _v = _flat(_k)
            if _v:
                meta[_item] = _v
        if _comp.get("account_balance_yuan") is not None:
            meta["专户剩余资金"] = _comp["account_balance_yuan"] / 10000.0  # 3.1.3 口径为万元
            meta["专户余额说明"] = (f"平台快照余额 {_comp['account_balance_yuan']:,.2f} 元"
                                   f"（{_comp.get('balance_asof') or '时点未知'}），不代表银行实时余额")
        if _comp.get("deposit_type"):
            meta["保证金详情"] = (f"{_comp['deposit_type']} {_comp.get('deposit_amount_wan') or 0} 万元，"
                                 f"有效期 {_comp.get('deposit_validity') or '待填'}")
    # 取数链（优先级严格）：MCP project_meta（平台工具，最优先） → 内联 attach_extracted（附件扫描兜底） → 待填
    # 注意 merge 顺序：MCP 取到的值必须在最后合并，确保工具优先于附件扫描。
    PH.set_meta({**(attach_extracted or {}), **meta})
    cfg = load_config()
    persons, roster_by_id, roster_by_nt = load_roster(ds)
    _step(f"花名册就绪（{len(persons)} 人）")
    price_table = load_price_table(ds)
    salary_months, att_months = month_list_from_files(ds)
    _step(f"月份清单就绪（工资 {len(salary_months)} 月 / 考勤 {len(att_months)} 月）")
    monthly = compute_monthly(ds, salary_months, att_months, price_table, roster_by_nt)
    _step("逐月聚合完成")
    yearly = compute_yearly_daily(ds, salary_months, att_months, year=_split_ym(target)[0])
    timeline = compute_timeline(persons, salary_months, att_months, ds)
    _step("全年日工资与进退场推断完成")
    # 引擎启动前把已拉数据写入跨进程缓存（引擎子进程直接复用，不再重复拉取触发限流）
    if hasattr(ds, "_save_disk_cache"):
        ds._save_disk_cache()
    _step("启动判定引擎子进程（工资明细全量 + 逐月考勤 + 对账，最耗时）")
    payload = run_engine(engine, target, data, source, judge_config, history, project_id, project_name)
    anchor = payload.get("salary_anchor")
    _step(f"引擎完成（对账锚 ok={anchor.get('ok') if isinstance(anchor, dict) else '无'}）")
    # 对账锚（2026-10-16）：引擎侧「当月明细合计 vs 批次总额」，不平则超欠付数字仅作参考
    anchor = payload.get("salary_anchor")

    # 项目名称：配置优先；无配置时从数据源 meta 提取（qxb 等支持 project_name 的数据源）
    proj_name = cfg["项目名称"]
    if "待填" in str(proj_name) and hasattr(ds, "project_name"):
        pn = ds.project_name()
        if pn:
            proj_name = pn
            PH.items = [x for x in PH.items if not (x["section"] == "封面" and x["item"] == "项目名称")]

    assets = ensure_assets(data, target)
    # ---- 12 类明细数据：接口优先，文件源默认空（基类空实现）→ 对应板块降级 ----
    # 3.2.1 线下考勤：--offline-att 指定工资册 xlsx 优先；否则走数据源（基类空实现）
    if _OFFLINE_ATT_PATH:
        off_att = load_offline_attendance(_OFFLINE_ATT_PATH)
        print(f"[offline-att] 线下考勤册载入 {len(off_att)} 人（{_OFFLINE_ATT_PATH}）", flush=True)
    else:
        off_att = ds.offline_attendance(target) or []
    fund_rows = ds.fund_flow() or []
    output_rows = ds.output_value() or []
    six_status_rows = ds.six_indicator_status(target) or []
    complaint_rows = ds.complaint_ledger()  # None=接口故障（待填）；[]=平台确认 0 条（显示 0）
    inspect_rows = ds.inspections() or []
    entry_exit_rows = ds.entry_exit(target) or []
    _ny_, _nm_ = _next_ym(target)
    next_att_rows = ds.next_attendance(f"{_ny_}{_nm_:02d}") or []
    yearly_rows = ds.yearly_salary() or []

    # 投诉：08 台账聚合优先，meta 手工汇总兜底
    # 语义（2026-10-16 老哥拍板）：接口正常 0 条 = 无投诉 → 显示 0；接口故障/无工具 → 待填
    _complaint_unavailable = complaint_rows is None
    complaints, complaints_monthly = _agg_complaints(complaint_rows or [], target,
                                                     zero_fill=not _complaint_unavailable)
    if not complaints:
        complaints = (meta.get("complaint_summary") if isinstance(meta, dict) else None) or {}
        complaints_monthly = (meta.get("complaints") if isinstance(meta, dict) else None) or {}
    # 资金：05+06 明细聚合优先，meta fund_history 兜底
    fund_history = _agg_fund(fund_rows, output_rows)
    if not fund_history:
        fund_history = (meta.get("fund_history") if isinstance(meta, dict) else None) or {}
    # 全年日工资：12 全年工资册接口优先，文件源现算兜底
    # 2026-10-19 口径统一：两条路都聚合成"逐人年"（年实发合计 ÷ 年考勤合计），限报告年度
    _ry_y = _split_ym(target)[0]
    if yearly_rows:
        _agg_y = {}
        for r in yearly_rows:
            if not str(r.get("ym", "")).startswith(str(_ry_y)):
                continue
            if r.get("days") and r["days"] > 0 and r.get("pay"):
                rec = _agg_y.setdefault(r.get("name") or "?", [0.0, 0.0])
                rec[0] += float(r["pay"])
                rec[1] += float(r["days"])
        _yr = [(p / d, "") for p, d in _agg_y.values() if d > 0 and p > 0]
        if _yr:  # 接口数据有效才覆盖（2026-10-16 防空：days 全 0 时保留现算结果）
            yearly = _yr

    charts = {}
    charts["salary_trend"] = chart_salary_trend(monthly, assets, target)
    att_target = load_attendance_month(ds, target)
    # 2026-10-19：次月预估与专户够发判定必须在此早算——2.x 资金行与末尾汇总表④的 need()
    # 都在本段之前渲染，晚了注入 meta 也来不及（v6 曾因此④行仍待填）
    _early_base = next_att_rows if next_att_rows else att_target
    _est_next_e, _est_days_e = _next_month_estimate(_early_base, roster_by_nt, price_table)
    _left_e = PH.meta.get("专户剩余资金")
    if isinstance(_left_e, (int, float)) and _est_next_e:
        # 2026-10-20 老哥拍板：判定列只放短判定词，数字细节进说明列
        _ok_e = _left_e * 10000 >= _est_next_e
        PH.meta["专户够发判定"] = "足够" if _ok_e else "不足"
        PH.meta["专户够发说明"] = (f"平台快照专户余额 {_left_e * 10000:,.2f} 元（{_left_e:,.2f} 万元）"
                                   + (" ≥ " if _ok_e else " < ")
                                   + f"次月预估工资 {_est_next_e:,.2f} 元"
                                   + ("" if _ok_e else f"，缺口 {_est_next_e - _left_e * 10000:,.2f} 元"))
    if six_status_rows:
        _ind_api = [{"name": r["name"], "status": _SIX_FROM_API.get(r["status"], "待填"),
                     "note": r["note"]} for r in six_status_rows]
    else:
        _ind_api = build_indicators(persons, att_target, target, salary_months, att_months, ds=ds)
    charts["six"] = chart_six_indicators(_ind_api, assets)
    charts["complaint"] = chart_complaint_trend(complaints_monthly, assets)
    charts["fund"] = chart_fund(fund_history, assets)
    charts["fund_rate"] = chart_fund_rate(fund_history, assets)
    charts["dist_all"] = chart_daily_dist(yearly, assets)
    charts["disp"] = chart_daily_disp(payload.get("persons", []), assets, target)
    charts["workforce"] = chart_team_workforce(att_target, assets, target)
    charts["attendance"] = chart_team_attendance(att_target, assets, target)

    doc = Document()
    # 默认正文字体
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    style.font.size = Pt(10.5)

    # ---- 封面（数据监理服务月报：标题 + 信息表 + 数据来源声明）----
    _ry, _rm = _split_ym(target)
    _now = datetime.now().strftime("%Y-%m-%d")                 # 数据截止时间=报告生成时时间，精确到年月日（老哥拍板）
    _ym6 = f"{_ry:04d}{_rm:02d}"
    _code = ((getattr(ds, "project_id", "") or project_id or "")[:4] or "XXXX")  # 编号简码：运行时解析出的项目 ID 派生，不固化
    _report_no = f"{REPORT_NO_PREFIX}-{_ym6}-{_code}"          # 报告编号自动生成（方案 B）
    title = doc.add_heading(f"{proj_name}工人工资数据监理服务月报", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_rows = [
        ["项目名称", proj_name],
        ["报告期间", f"{_ry} 年 {_rm} 月（月度服务报告）"],
        ["数据截止时间", _now],
        ["报告生成单位", REPORT_UNIT],
        ["建设单位", cfg["建设单位"]],
        ["施工单位", cfg["施工单位"]],
        ["监理单位", cfg["监理单位"]],
        ["报告编号", _report_no],
        ["报送对象", SUBMIT_TO],
    ]
    add_table(doc, ["项目", "内容"], cover_rows, col_widths=[2.2, 4.4])
    para(doc, "本报告数据来源于黔薪保实名制管理平台、薪莫愁服务平台及项目报送资料。", size=9, color="666666")

    # ============ 一、项目概况 ============
    h(doc, "一、项目概况", 1)
    # 已运行（2026-10-19 老哥定口径）= 报告年月 − 项目开工年月；开工日期取
    # 平台档案实际开工时间（actualBeginDate，本项目 2024-12-15），无则回退薪莫愁介入时间（人工）
    _ty0, _tm0 = _split_ym(target)
    _begin_src = PH.meta.get("实际开工时间") or PH.meta.get("薪莫愁介入时间")
    _bg_ym = parse_ym_month(str(_begin_src)) if _begin_src else None
    if _bg_ym:
        _by, _bm = _split_ym(_bg_ym)
        run_months = (_ty0 * 12 + _tm0) - (_by * 12 + _bm) + 1
    else:
        run_months = PH.need("项目概况", "已运行时间")
    # 预计剩余 = 预计完工年月 − 报告年月（平台档案 endDate，本项目 2026-10-29）
    _end_v = PH.meta.get("预计完工时间")
    _end_ym = parse_ym_month(str(_end_v)) if _end_v else None
    if _end_ym:
        _ey, _em = _split_ym(_end_ym)
        remain = max((_ey * 12 + _em) - (_ty0 * 12 + _tm0), 0)
    else:
        remain = PH.need("项目概况", "预计剩余时间")
    # 项目类别 + 适用拨付比例（住建领域人工费分账拨付比例政策线）
    _ptype = PH.need("项目概况", "项目类别")
    _ratio = ""
    for _kw, _r in (("房屋建筑", "20%"), ("市政", "15%"), ("水利", "10%"),
                    ("公路", "5%"), ("水运", "5%"), ("铁路", "2%"), ("民航", "2%")):
        if _kw in str(_ptype):
            _ratio = _r
            break
    _ptype_disp = f"{_ptype}（适用拨付比例：不低于 {_ratio}）" if _ratio else _ptype
    rows = [
        ["项目名称", proj_name],
        ["项目类别 / 适用拨付比例", _ptype_disp],
        ["建设单位", f"{cfg['建设单位']}（现场负责人：{cfg['现场负责人']}）"],
        ["施工单位", f"{cfg['施工单位']}（项目经理：{cfg['项目经理']}　劳资专管员：{cfg['劳资专管员']}）"],
        ["实际开工 / 预计完工", f"{cfg['开工日期']}　/　{cfg['预计完工日期']}"],
        ["项目已运行 / 预计剩余", f"已运行 {run_months} 个自然月　/　预计剩余 {remain} 个自然月"],
        # 2026-10-20 老哥拍板：当前形象进度为纯人工报送字段，报告中不再出现（不占待填）
        ["薪莫愁数据监理服务介入时间", f"{PH.need('项目概况', '薪莫愁介入时间')}（项目第一次扫码登记入库时间，人工填报）"],
    ]
    add_table(doc, ["字段", "内容"], rows, col_widths=[2.8, 3.8])
    para(doc, "责任划断声明：薪莫愁数据监理服务于上表所列介入时间正式介入本项目。对介入前已产生的数据，本报告仅作梳理呈现，"
              "不对其真实性、合规性负责；介入后产生的数据按《薪莫愁数据监理服务标准工作流程》核查并对核查结论负责。",
         size=9, color="666666")

    # ============ 二、项目健康度分析 ============
    h(doc, "二、项目健康度分析", 1)

    # 2.1 健康度汇总表（五项指标 + 综合评级）
    h(doc, "2.1 健康度汇总表（五项指标）", 2)
    hs = health_summary(persons, att_target, target, salary_months, att_months, payload, indicators=_ind_api)
    add_table(doc, ["序", "健康度指标", "判定", "说明"], hs, col_widths=[0.5, 2.0, 1.1, 2.9])

    # 2.2 六项制度落实情况
    h(doc, "2.2 六项制度落实情况", 2)
    indicators = _ind_api
    doc.add_picture(charts["six"], width=Inches(6.0))
    add_table(doc,
              ["指标", "状态", "未达标原因 / 处置进展"],
              [[i["name"], _six_display(i["status"]), i["note"] or ""] for i in indicators],
              col_widths=[1.8, 1.1, 3.6])
    para(doc, "说明：绿=达标、黄=限期整改、红=严重违规、灰=待平台回填。未达标须写明原因。", size=9, color="888888")

    # 2.3 项目欠薪投诉情况（三端口）
    h(doc, "2.3 项目欠薪投诉情况（三端口）", 2)
    para(doc, "由政府端、项目端、薪莫愁端三个端口分别报送收到投诉情况，然后汇总到薪莫愁服务平台统计。", size=9)
    cs = complaints or {}

    def _cs_row(dim):
        v = cs.get(dim)
        if not v:
            return ["【待填】", "【待填】"]
        return [f"{v[0]:g} 件", f"{v[1]:g} 万元"]

    rows = [["历史累计（介入以来）"] + _cs_row("历史累计"),
            ["其中：政府端"] + _cs_row("政府端"),
            ["项目端"] + _cs_row("项目端"),
            ["薪莫愁端"] + _cs_row("薪莫愁端"),
            [f"本月（{_rm} 月）新增"] + _cs_row("本月新增"),
            ["已解决"] + _cs_row("已解决"),
            ["未解决"] + _cs_row("未解决")]
    add_table(doc, ["统计维度", "投诉件数", "涉及金额"], rows, col_widths=[2.6, 1.6, 1.6])
    if isinstance(cs.get("历史累计"), list) and not cs["历史累计"][0] and not complaints_monthly:
        para(doc, "截至报告期，政府端、项目端、薪莫愁端三端口均未收到本项目欠薪投诉记录（平台台账 0 条）。",
             size=9, color="888888")
    if charts["complaint"]:
        doc.add_picture(charts["complaint"], width=Inches(6.2))
        para(doc, "图：三端口欠薪投诉走势（柱=三端口合计，线=各端口）。", size=9, color="888888")
        _months = sorted(complaints_monthly.keys())
        _tot = [sum(complaints_monthly[m].values()) for m in _months]
        _peak = _months[_tot.index(max(_tot))]
        para(doc, f"文字分析：历史累计投诉覆盖 {len(_months)} 个月，峰值在 {_peak}（{max(_tot):g} 件）；"
                  "结合本月新增与已解决数量，评估欠薪风险趋势与处置进度。", size=9)
    else:
        para(doc, "历史累计投诉做成走势图：总表三个端口收到的投诉情况、分表在分别做三个端口自己的情况。"
                  "当前无投诉月度数据源，待平台回填后绘制。" + PH.need("欠薪情况", "投诉走势图数据"), size=9, color="888888")

    # 2.4 项目资金情况
    h(doc, "2.4 项目资金情况", 2)
    total_paid_all = sum(m["pay_sum"] for m in monthly if m["has_salary"])

    def _wan(v):
        """meta 数值按万元保留 2 位；字符串原样。"""
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return f"{v:,.2f}"
        return v

    # 05+06 接口聚合优先（末月累计值），meta 手工值兜底
    if fund_history:
        _last = fund_history[max(fund_history.keys())]
        _rep = _last.get("报审") or PH.need("资金情况", "报审产值")
        _appr = _last.get("审定") or PH.need("资金情况", "审定产值")
        _acc = _last.get("专户") or PH.need("资金情况", "专户进账金额")
        if _last.get("审定"):
            _v = _last["专户"] / _last["审定"] * 100
            _flag = "低于" if _v < 20 else "达到"
            _rate_note = f"{_v:.2f}%（{_flag}房建 20% 要求）"
        else:
            _rate_note = PH.need("资金情况", "分账拨付率")
    else:
        _rep = PH.need("资金情况", "报审产值")
        _appr = PH.need("资金情况", "审定产值")
        _acc = PH.need("资金情况", "专户进账金额")
        _rate_note = PH.need("资金情况", "分账拨付率")
    _pay = PH.need("资金情况", "拨付工程款")
    if isinstance(_appr, (int, float)) and isinstance(_pay, (int, float)) and _appr:
        _progress = f"拨款进度率 {_pay / _appr * 100:.2f}%"
    else:
        _progress = "拨款进度率待数据"
    rows = [
        ["累计报审产值", _wan(_rep), PH.need("资金情况", "报审产值说明")],
        ["累计审定产值", _wan(_appr), "—"],
        ["建设单位累计拨付工程款", _wan(_pay), _progress],
        ["农民工工资专户累计进账", _wan(_acc), f"其中非建设单位打款 {PH.need('资金情况', '非建设单位打款情况')}"],
        ["合规分账进账（剔除违规）", _wan(PH.need("资金情况", "合规分账进账")), "—"],
        ["分账拨付率", _rate_note, "专户累计进账 ÷ 累计审定产值；政策要求：房建20%、市政/水利/电力15%、交通10%"],
        ["专户累计发放工资", f"{total_paid_all / 10000:,.2f}", "—"],
        ["专户剩余资金", _wan(PH.need("资金情况", "专户剩余资金")), PH.need("资金情况", "专户够发判定")],
    ]
    add_table(doc, ["指标", "金额（万元）", "说明"], rows, col_widths=[2.3, 1.3, 3.0])
    if charts["fund"]:
        doc.add_picture(charts["fund"], width=Inches(6.2))
        para(doc, "图：项目累计资金拨付与专户进账情况。", size=9, color="888888")
    if charts["fund_rate"]:
        doc.add_picture(charts["fund_rate"], width=Inches(6.2))
        para(doc, "图：分账拨付率走势（专户累计进账 ÷ 累计审定产值）。", size=9, color="888888")

    # 2.5 本年度工资发放统计（截至报告期）
    h(doc, "2.5 本年度工资发放统计（截至报告期）", 2)
    para(doc, "按本年度计算，所有月份统计在一张表上：已发生月份统计发放金额、发放人数、考勤总天数；"
              "未发生月份统计发放人数、考勤总天数、预估月份发放金额（预估工资计算逻辑见工资审核板块）。"
              "本月出具上月报告，本月有考勤数据、只预估本月情况，下月为空白值。", size=9)
    _ry, _rm = _split_ym(target)
    rows = []
    for m in monthly:
        _yy, _mm2 = _split_ym(m["ym"])
        if _yy != _ry:
            continue
        # 2026-10-19 老哥拍板：只统计到报告月，报告月之后的月份不进表
        if (_yy, _mm2) > (_ry, _rm):
            continue
        _ym_disp = f"{_yy}-{_mm2:02d}"
        if m["has_salary"]:
            kind = "已发生" if m["ym"] != target else "报告期"
            pay = f"{m['pay_sum'] / 10000:,.2f}"; est = "—"
        elif m["has_att"]:
            kind = "预估"
            pay = "—"; est = f"{m['est'] / 10000:,.2f}" if m["est"] is not None else PH.need("工资发放", "预估月份发放金额(缺工种)")
        else:
            kind = "空白值"
            pay = "—"; est = "—"
        rows.append([_ym_disp, pay, m["sal_persons"] if m["has_salary"] else (m["att_persons"] if m["has_att"] else "—"),
                     f"{m['days_sum']:,.0f}" if m["has_att"] else "—", kind])
    # 2026-10-20：累计行与行显示同口径——只统计到报告月（9 月 154 元曾混入累计）
    _in_range = lambda m: (_split_ym(m["ym"])[0], _split_ym(m["ym"])[1]) <= (_ry, _rm)
    _sum_pay = sum(m["pay_sum"] for m in monthly
                   if m["has_salary"] and _split_ym(m["ym"])[0] == _ry and _in_range(m))
    _sum_days = sum(m["days_sum"] for m in monthly
                    if m["has_att"] and _split_ym(m["ym"])[0] == _ry and _in_range(m))
    # 2026-10-20 老哥拍板补累计发放人数：按（姓名+掩码证件号）去重——personId 会因离场重进场
    # 重复发放（本项目 429 个 ID 实为 412 人），纯姓名又合并同名不同人（410）；括注各月人次合计
    _sal_yr = [it for it in (ds._salary_all() if hasattr(ds, "_salary_all") else [])
               if str(it.get("payMonth", "")).startswith(str(_ry))
               and (int(str(it.get("payMonth"))[:4]), int(str(it.get("payMonth"))[4:6])) <= (_ry, _rm)]
    _ppl = {(str(it.get("name") or "").strip(), str(it.get("idCardNo") or "").strip())
            for it in _sal_yr if str(it.get("name") or "").strip()}
    _prs_sum = sum(m["sal_persons"] for m in monthly
                   if m["has_salary"] and _split_ym(m["ym"])[0] == _ry and _in_range(m))
    rows.append(["本年累计", f"{_sum_pay / 10000:,.2f}",
                 f"{len(_ppl)}（{_prs_sum:,} 人次）", f"{_sum_days:,.0f}", "—"])
    add_table(doc, ["月份", "实发工资（万元）", "发放人数（人）", "考勤总天数（天）", "备注"], rows,
              col_widths=[0.9, 1.5, 1.3, 1.4, 1.0])
    doc.add_picture(charts["salary_trend"], width=Inches(6.2))
    para(doc, "图：本年度工资发放金额与人数。", size=9, color="888888")
    # 当月工资整体审核情况
    _s = payload["summary"]; _fc = payload["four_category"]
    _att_n = len(att_target)
    _sal_n = _s.get("total_workers", 0)
    _rate = f"{_sal_n / _att_n * 100:.2f}%" if _att_n else "—"
    para(doc, f"当月工资整体审核情况：应发人数（考勤在册）{_att_n} 人，实发人数 {_sal_n} 人，发放率 {_rate}。"
              f"出现以下异常（出现一个即重点关注）：有考勤无工资 {len(_fc['attendance_no_salary'])} 人；"
              f"有工资无考勤 {len(_fc['salary_no_attendance'])} 人；日工资高于审核区间 {_s['over']} 人（超付 {fmt_money(_s['over_amount'])}）；"
              f"日工资低于审核区间 {_s['under'] + _s['possible_under']} 人，其中坐实欠付 {_s['under']} 人"
              f"（欠付 {fmt_money(_s['under_amount'])}）、可能欠付 {_s['possible_under']} 人"
              f"（预警待复核，经第二轮历史均值比对未坐实）。详见第三章。", size=9)
    if anchor:
        if anchor.get("ok"):
            para(doc, f"数据对账：当月发放明细合计 ¥{anchor['detail_sum']:,.2f}（{anchor['detail_count']} 人）"
                      f"与工资册批次总额 ¥{anchor['batch_sum']:,.2f}（{anchor['batch_count']} 笔）对平，数据完整。",
                size=9, color="888888")
        else:
            para(doc, f"⚠️ 数据对账不平：当月明细合计 ¥{anchor['detail_sum']:,.2f}（{anchor['detail_count']} 人）"
                      f"≠ 工资册批次总额 ¥{anchor['batch_sum']:,.2f}（{anchor['batch_count']} 笔）——"
                      f"当月取数不完整，上述人数与金额仅作参考（详见 3.1.2）。",
                size=9, color="C0392B")

    # ============ 三、工资审核 ============
    h(doc, "三、工资审核", 1)

    # 3.1 主要审核点
    h(doc, "3.1 主要审核点", 2)
    para(doc, "审核依据：当月黔薪保实名制系统工资册与考勤册。测算单价统一采用薪莫愁《工种行业单价预估范围》行情单价，"
              "不采用黔薪保合同单价——合同单价由施工方与班组自行约定，存在人为压低工价以“合同价”掩盖欠付的可能，"
              "第三方行情单价才是真实反映欠付的客观基准；合同单价仅在本报告 3.2.2 合同抽查中作偏离度辅助校验。"
              "当月日工资 = 当月实发 ÷ 当月黔薪保考勤；预警区间 = [行情下限 × (1−15%)，行情上限 × (1+30%)]；"
              "超付金额 = 实发 − 行情区间上限，欠付金额 = 行情区间下限 − 实发。", size=9)

    # 3.1.1 日工资离散分析
    h(doc, "3.1.1 日工资离散分析", 3)
    if charts["disp"]:
        doc.add_picture(charts["disp"], width=Inches(6.2))
        para(doc, "图：当月全项目人员日工资离散分布（x=当月考勤天数，y=当月日工资；蓝=区间内、黄=高于上限、红=低于下限）。", size=9, color="888888")
    render_daily_discrete(doc, payload.get("persons", []))

    # 3.1.2 超欠付预警清单
    h(doc, "3.1.2 超欠付预警清单", 3)
    if anchor and not anchor.get("ok"):
        para(doc, f"⚠️ 数据对账不平（明细 ¥{anchor['detail_sum']:,.2f}/{anchor['detail_count']} 人 "
                  f"≠ 批次 ¥{anchor['batch_sum']:,.2f}/{anchor['batch_count']} 笔）：当月取数疑不完整，"
                  f"本清单人数与金额仅作参考，须复核取数完整性后重跑再采信。",
            size=9, color="C0392B")
        if not any(x["section"] == "工资审核" and x["item"] == "工资明细与批次对账" for x in PH.items):
            PH.items.append({"section": "工资审核", "item": "工资明细与批次对账",
                             "hint": f"目标月明细合计 ¥{anchor['detail_sum']:,.2f} ≠ 批次总额 "
                                     f"¥{anchor['batch_sum']:,.2f}（疑上游分页截断），须重跑取数至对平"})
    # 2026-10-19 新工具：平台月度工资考勤汇总三重核验（汇总 vs 明细，独立于批次锚）
    if anchor and anchor.get("ok"):
        try:
            _ms_list = ds.monthly_summary(target[:4]) if hasattr(ds, "monthly_summary") else []
            _t6 = _norm_ym(target)
            _ms = next((r for r in _ms_list if r["ym"] == _t6), None)
            if _ms:
                _diff = abs(_ms["pay"] - anchor["detail_sum"])
                _dp = _ms["persons"] - anchor["detail_count"]
                _ok3 = (_diff < 0.01 and _dp == 0)
                _verdict = (f"平台月度工资考勤汇总交叉核验：{target[:2]}年{target[3:]}月发放 "
                            f"¥{_ms['pay']:,.2f} / {_ms['persons']} 人"
                            f"（发放方式 {_ms['pay_type'] or '—'}），与明细对账结果"
                            + ("一致 ✓" if _ok3 else
                               f"不一致（差 ¥{_diff:,.2f} / {_dp:+d} 人），须复核取数完整性"))
                para(doc, _verdict, size=9, color="1E8449" if _ok3 else "C0392B")
        except Exception:
            pass
    # 2026-10-23 老哥拍板：身份一致性提示补当月两类——黔薪保考勤一人多班组、工资册同名多行
    # （render_four_category 只有 payload，可及性数据在此算好塞进去）
    _ia_aug = payload.setdefault("identity_anomalies", {})
    _mt_tmp = {}
    try:
        for (_an, _at), _ad in (ds.kaoqin(target)[0] or {}).items():
            _mt_tmp.setdefault(_an, []).append((_at, _ad))
    except Exception:
        pass
    _ia_aug["month_multi_team"] = [
        {"name": n, "teams": v} for n, v in sorted(_mt_tmp.items()) if len(v) > 1]
    _odup_tmp = {}
    if off_att and "name" in off_att[0]:
        for _o in off_att:
            _odup_tmp.setdefault(_o["name"], []).append(
                (_o["team"] or "未分班组", _o["offline_days"]))
    _ia_aug["offline_dup_rows"] = [
        {"name": n, "rows": v} for n, v in sorted(_odup_tmp.items()) if len(v) > 1]
    render_four_category(doc, payload)

    # 3.1.3 次月预估工资
    h(doc, "3.1.3 次月预估工资", 3)
    _ny, _nm = _next_ym(target)
    _est_base = next_att_rows if next_att_rows else att_target  # 11 次月已发生考勤优先，退回当月考勤
    _est_next, _est_days = _next_month_estimate(_est_base, roster_by_nt, price_table)
    _avg_daily = _est_next / _est_days if _est_days else None
    _left = PH.need("资金情况", "专户剩余资金")
    # 口径：meta 资金字段为万元，预估为元；缺口统一折万元
    _left_wan = _left if isinstance(_left, (int, float)) else None
    _gap = (_est_next / 10000 - _left_wan) if _left_wan is not None else None
    # 2026-10-20：专户够发判定/说明已早算注入（判定=短词、说明=数字细节），此处不再覆盖
    rows = [
        [f"{_nm} 月在册人数（预估）", len(att_target)],
        [f"{_nm} 月考勤天数", f"{_est_days:,.0f}"],
        [f"{_nm} 月预估工资 = Σ[（工种上限 + 工种下限）÷ 2 × 考勤]", fmt_money(_est_next)],
        ["加权平均日工资", f"{_avg_daily:.2f} 元/天" if _avg_daily else "—"],
        ["专户剩余资金（万元）", f"{_left_wan:,.2f}" if _left_wan is not None else _left],
        ["资金缺口（万元）", f"{_gap:,.2f}" if _gap is not None else "【待填】"],
    ]
    add_table(doc, ["项目", "数值"], rows, col_widths=[3.6, 2.8])
    para(doc, "注：预估仅按行情单价粗估，实际发放以劳动合同约定单价为准，本数仅供资金体量参考。", size=9, color="888888")

    # 3.1.4 全年日工资水平分布
    h(doc, "3.1.4 全年日工资水平分布", 3)
    if charts["dist_all"]:
        doc.add_picture(charts["dist_all"], width=Inches(6.2))
        para(doc, "图：全项目人员全年日工资水平分布（年实发 ÷ 年考勤），区间按 0-87、88-200 起分档。", size=9, color="888888")
    else:
        para(doc, "全年日工资分布暂无可计算样本（缺考勤或工资金额），待数据补全后自动出图。", size=9, color="888888")
    render_daily_discrete(doc, payload.get("persons", []), yearly=yearly, monthly=False)

    # 3.2 次要审核点
    h(doc, "3.2 次要审核点", 2)
    h(doc, "3.2.1 线下考勤与黔薪保考勤比对", 3)
    if off_att:
        def _nteam(t):
            # 2026-10-23：班组名归一——工资册"水电/机械/外架"与平台"水电班组/机械班组/外架班组"同指；
            # 其余班组两册同名（混凝土工/模板工/砌筑工/杂工班组/防水班组/墙体砌筑/墙板砌筑）
            t = (t or "").strip()
            return t[:-2] if t.endswith("班组") and len(t) > 2 else (t or "未分班组")

        _per_person = "name" in off_att[0]  # 工资册解析路径带姓名，可按人关联；数据源路径仅班组（旧口径）
        _off_team = {}   # 归一班组 -> [展示名, 人数, 线下总天, 黔薪保匹配天, 差异≥3天人数]
        _unmatched = []  # 线下册有、黔薪保无当月考勤（重点核查：有工资无平台考勤）
        _diff_persons = []  # 差异≥3天逐人名单（name, 班组, 线下天, 平台天, 差异）
        if _per_person:
            _qx = {}
            # 2026-10-23：用原始 (姓名,班组) 口径——attendance_target 会把跨班组人员的天数
            # 按姓名合并累加（如刘显军=水电29+混凝土27=56），逐人比对必须分班组取数
            try:
                _bnt = ds.kaoqin(target)[0] or {}
            except Exception:
                _bnt = {}
            if _bnt:
                for (_qn, _qt), _qd in _bnt.items():
                    _qx.setdefault(_qn, []).append({"name": _qn, "team": _qt, "days": _qd})
            else:
                for a in att_target:
                    _qx.setdefault(a["name"], []).append(a)
            _seen_off = set()
            for o in off_att:
                _t = _nteam(o["team"])
                _disp = o["team"] or "未分班组"
                _seen_off.add(o["name"])
                rec = _off_team.setdefault(_t, [_disp, 0, 0.0, 0.0, 0])
                rec[1] += 1
                rec[2] += o["offline_days"]
                _cands = _qx.get(o["name"])
                if not _cands:
                    _unmatched.append((o["name"], _disp, o["offline_days"]))
                    continue
                # 同名跨班组消歧：优先取班组一致的记录，否则取天数最大的一条
                _pick = next((a for a in _cands if _nteam(a["team"]) == _t), None) \
                    or max(_cands, key=lambda a: a["days"])
                rec[3] += _pick["days"]
                _diff = o["offline_days"] - _pick["days"]
                if abs(_diff) >= 3:
                    rec[4] += 1
                    _diff_persons.append((o["name"], _disp, o["offline_days"], _pick["days"], _diff))
            _only_qx = {}  # 黔薪保有当月考勤、线下册无（team -> [(name, days)]）
            for a in att_target:
                if a["name"] not in _seen_off:
                    _only_qx.setdefault(_nteam(a["team"]), []).append((a["name"], a["days"]))
        else:
            for o in off_att:
                _t = o["team"] or "未分班组"
                rec = _off_team.setdefault(_t, [_t, 0, 0.0, 0.0, 0])
                rec[1] += 1
                rec[2] += o["offline_days"]

        def _fd(v):
            return f"{v:,.2f}".rstrip("0").rstrip(".")

        rows = []
        for _t in sorted(_off_team):
            _disp, _on, _od, _qd, _big = _off_team[_t]
            rows.append([_disp, _on, _fd(_od), _fd(_qd), f"{_od - _qd:+,.2f}".rstrip("0").rstrip(".")])
        _n_all = sum(r[1] for r in _off_team.values())
        _od_all = sum(r[2] for r in _off_team.values())
        _qd_all = sum(r[3] for r in _off_team.values())
        rows.append(["合计", _n_all, _fd(_od_all), _fd(_qd_all),
                     f"{_od_all - _qd_all:+,.2f}".rstrip("0").rstrip(".")])
        add_table(doc, ["班组", "人数", "线下考勤（天）", "黔薪保考勤（天）", "差异（天）"], rows,
                  col_widths=[1.8, 0.9, 1.3, 1.4, 1.0])
        _extra = []
        if _per_person:
            _n_big = sum(r[4] for r in _off_team.values())
            _big_teams = "、".join(f"{t}{r[4]} 人" for t, r in sorted(_off_team.items()) if r[4])
            if _n_big:
                _extra.append(f"两册差异≥3天共 {_n_big} 人（{_big_teams}）")
            if _unmatched:
                _head = "、".join(f"{n}（{t}，{d:g}天）" for n, t, d in _unmatched[:6])
                _more = f"等共 {len(_unmatched)} 人" if len(_unmatched) > 6 else ""
                _extra.append(f"线下册有而黔薪保无当月考勤 {len(_unmatched)} 人：{_head}{_more}，"
                              f"需核查系遗漏打卡还是虚列用工")
            if _only_qx:
                _sum = sum(len(v) for v in _only_qx.values())
                _detail = "、".join(f"{t}{len(v)} 人" for t, v in sorted(_only_qx.items()))
                _extra.append(f"黔薪保有当月考勤而线下册无 {_sum} 人（{_detail}，其中管理/监理岗位不属用工册范围属正常）")
        para(doc, "两个考勤册差异超过三天则提示：有多少人、分布在哪些班组、线下考勤天数、黔薪保考勤天数、对比差异天数；"
                  "已按黔薪保考勤为准核定工资。"
                  + ("；".join(_extra) + "。" if _extra else ""), size=9)
        # —— 差异人员名单（2026-10-23 老哥拍板：具体到人，对齐异常人员名单格式）——
        if _per_person:
            if _diff_persons:
                _diff_persons.sort(key=lambda x: -abs(x[4]))
                para(doc, f"（1）两册差异≥3天人员（{len(_diff_persons)} 人）：按差异绝对值降序", bold=True, size=9)
                add_table(doc,
                          ["序号", "姓名", "班组", "线下考勤(天)", "黔薪保考勤(天)", "差异(天)"],
                          [[str(i), _n0, _t0, _fd(_od), _fd(_qd), f"{_df:+,.2f}".rstrip("0").rstrip(".")]
                           for i, (_n0, _t0, _od, _qd, _df) in enumerate(_diff_persons, 1)],
                          col_widths=[0.5, 1.0, 1.2, 1.2, 1.3, 1.0])
            else:
                para(doc, "两册差异均未超过三天。", size=9, color="888888")
            if _unmatched:
                para(doc, f"（2）线下册有而黔薪保无当月考勤人员（{len(_unmatched)} 人）："
                          "需核查系遗漏打卡还是虚列用工", bold=True, size=9)
                add_table(doc,
                          ["序号", "姓名", "班组", "线下考勤(天)", "说明"],
                          [[str(i), _n0, _t0, _fd(_od), "黔薪保无当月考勤记录"]
                           for i, (_n0, _t0, _od) in enumerate(_unmatched, 1)],
                          col_widths=[0.5, 1.0, 1.2, 1.2, 2.6])
            if _only_qx:
                _oq_rows = []
                for _t, _lst in sorted(_only_qx.items()):
                    for _n0, _qd in _lst:
                        _note = ("管理/监理岗位，不属用工册范围属正常" if _t in ("管理", "监理")
                                 else "工资册无此人，需核查用工与工资是否漏报")
                        _oq_rows.append([str(len(_oq_rows) + 1), _n0, _t, _fd(_qd), _note])
                para(doc, f"（3）黔薪保有当月考勤而线下册无人员（{len(_oq_rows)} 人）", bold=True, size=9)
                add_table(doc,
                          ["序号", "姓名", "班组", "黔薪保考勤(天)", "说明"],
                          _oq_rows,
                          col_widths=[0.5, 1.0, 1.2, 1.3, 2.4])
    elif (meta.get("offline_attendance") if isinstance(meta, dict) else None):
        off = meta["offline_attendance"]
        rows = [[o.get("班组", ""), o.get("人数", ""), o.get("线下", ""), o.get("黔薪保", ""), o.get("差异", "")]
                for o in off]
        _n = sum(o.get("人数", 0) or 0 for o in off)
        _d = sum(o.get("差异", 0) or 0 for o in off)
        rows.append(["合计", _n, "—", "—", f"+{_d}"])
        add_table(doc, ["班组", "人数", "线下考勤（天）", "黔薪保考勤（天）", "差异（天）"], rows,
                  col_widths=[1.8, 0.9, 1.3, 1.4, 1.0])
        para(doc, "两个考勤册差异超过三天则提示：有多少人、分布在哪些班组、线下考勤天数、黔薪保考勤天数、对比差异天数；"
                  "已按黔薪保考勤为准核定工资。", size=9)
    else:
        para(doc, "有线下考勤则比对，无线下考勤册则不对比。当前无线下考勤数据源，待补充。" + PH.need("工资审核", "线下考勤对比"),
             size=9, color="888888")
    h(doc, "3.2.2 进场合同抽查", 3)
    _ty_, _tm_ = _split_ym(target)
    _tnorm = f"{_ty_:04d}{_tm_:02d}"

    def _same_month(d):
        return bool(d) and _norm_ym(d) == _tnorm

    _CONTRACT_OK = ("是", "已签订", "已签", "有")
    if entry_exit_rows:
        _entry_rows = [r for r in entry_exit_rows if _same_month(r.get("entry"))]
        _left_rows = [r for r in entry_exit_rows if _same_month(r.get("leave"))]
        _entry_bad = sum(1 for r in _entry_rows if r.get("contract") not in _CONTRACT_OK)
        _settle_bad = sum(1 for r in _left_rows if r.get("settle_status") and "未" in str(r.get("settle_status")))
    else:
        _entry_rows = [p for p in persons if parse_ym_month(p["entry"]) == _tnorm]
        _left_rows = [p for p in timeline if p["leave"] == target]
        _entry_bad = sum(1 for p in _entry_rows if p["contract"] != "是")
        _settle_bad = None
    _check_n = PH.need("工资审核", "合同抽查份数")
    para(doc, f"当月新进场 {len(_entry_rows)} 人，抽查劳动合同 {_check_n} 份，不合规 {_entry_bad} 份"
              f"（缺工价、无工人签字、无印章等要素不全），已通知 10 日内补正。"
              + PH.need("工资审核", "低工价合同提示"), size=9)
    h(doc, "3.2.3 离场结算抽查", 3)
    para(doc, f"当月离场 {len(_left_rows)} 人，抽查离场结算单 {PH.need('工资审核', '离场结算抽查份数')} 份，"
              f"不合规 {_settle_bad if _settle_bad is not None else PH.need('工资审核', '离场结算不合规数')} 份。"
              + PH.need("工资审核", "超30日未结算离场"), size=9)

    # ============ 四、人员情况 ============
    h(doc, "四、人员情况", 1)
    h(doc, "4.1 当月用工情况", 2)
    from collections import defaultdict as _dd
    _team_people = _dd(int); _team_days = _dd(float)
    for a in att_target:
        _team_people[a["team"]] += 1
        _team_days[a["team"]] += a["days"]
    rows = [[t, _team_people[t], f"{_team_days[t]:,.0f}", f"{_team_days[t] / _team_people[t]:.1f}"]
            for t in sorted(_team_people, key=lambda x: -_team_people[x])]
    rows.append(["合计", len(att_target), f"{sum(_team_days.values()):,.0f}",
                 f"{sum(_team_days.values()) / len(att_target):.1f}" if att_target else "—"])
    add_table(doc, ["班组", "用工人数（人）", "考勤天数（天）", "人均考勤（天）"], rows,
              col_widths=[1.9, 1.5, 1.5, 1.5])
    doc.add_picture(charts["workforce"], width=Inches(6.0))
    para(doc, "图：当月各班组用工人数。", size=9, color="888888")
    if charts["attendance"]:
        doc.add_picture(charts["attendance"], width=Inches(6.0))
        para(doc, "图：当月各班组人员考勤天数。", size=9, color="888888")
    para(doc, "用工变化分析：当月用工以各班组的用工人数与考勤总量体现，人均考勤显著偏低的班组结合欠付预警跟踪核查。", size=9)

    # ============ 五、其他特殊情况 ============
    h(doc, "五、其他特殊情况", 1)
    h(doc, "5.1 欠薪事件处理", 2)
    if complaints:
        rows = [
            ["历史累计投诉", f"{complaints['历史累计'][0]:g} 人", f"{complaints['历史累计'][1]:g} 万元"],
            ["已解决", f"{complaints['已解决'][0]:g} 人", f"{complaints['已解决'][1]:g} 万元"],
            ["未解决", f"{complaints['未解决'][0]:g} 人", f"{complaints['未解决'][1]:g} 万元"],
        ]
    else:
        rows = [
            ["历史累计投诉", PH.need("其他", "欠薪事件历史累计(人数/金额)"), PH.need("其他", "欠薪事件历史累计(人数/金额)")],
            ["已解决", PH.need("其他", "欠薪事件已解决(人数/金额)"), PH.need("其他", "欠薪事件已解决(人数/金额)")],
            ["未解决", PH.need("其他", "欠薪事件未解决(人数/金额)"), PH.need("其他", "欠薪事件未解决(人数/金额)")],
        ]
    add_table(doc, ["统计维度", "人数（人）", "金额（万元）"], rows, col_widths=[2.0, 1.8, 1.8])
    para(doc, "注：投诉事件定性、协调处置与证据链出具情况据实填写。" + PH.need("其他", "欠薪事件处置说明"), size=9, color="888888")
    h(doc, "5.2 迎检情况", 2)
    if inspect_rows:
        _detail = "；".join(f"{r['date']} {r['unit']}（{r['matter']}，要求整改：{r['rectify']}）"
                           for r in inspect_rows if r.get("date") or r.get("unit"))
        para(doc, f"本月迎检 {len(inspect_rows)} 次。{_detail}", size=9)
    else:
        para(doc, f"本月迎检 {PH.need('其他', '迎检次数')} 次。要求整改点及完成情况：{PH.need('其他', '迎检详情')}。", size=9)
    para(doc, "原则：有就写、没有就不管，用句式写。", size=9, color="888888")

    # ============ 六、建议与反馈 ============
    h(doc, "六、建议与反馈", 1)
    render_suggestions(doc, payload, meta, _est_next, _left, _rate_note)

    # ---- 落款 ----
    para(doc, f"编制：薪莫愁运营部　　审核：运营部负责人　　报送：{SUBMIT_TO}", size=10)
    para(doc, f"报告编号：{_report_no}　　出具日期：{_now}", size=10)

    # ---- 附：待补充数据清单（无待填项则不输出）----
    render_pending_supplement(doc, PH.items)

    doc.save(out)
    return out, payload, monthly, yearly


_SIX_DISPLAY = {"达标": "达标", "提示": "限期整改", "警告": "严重违规",
                "严重违规": "严重违规", "不达标": "严重违规", "待填": "待平台回填"}


def _six_display(status):
    """内部状态 → 六项制度展示词（达标/限期整改/严重违规/待平台回填）。"""
    return _SIX_DISPLAY.get(status, status)


# 07 接口状态词（绿/黄/红）→ 内部状态
_SIX_FROM_API = {"绿": "达标", "达标": "达标", "正常": "达标",
                 "黄": "提示", "限期整改": "提示", "提示": "提示",
                 "红": "警告", "严重违规": "警告", "警告": "警告"}


def _norm_ym(s):
    """日期/月份串 → YYYYMM（支持 YYYY-MM-DD / YYYY-MM / YYYYMM / YYMM / YY.M / YYYY.M）。"""
    s = str(s or "").strip()
    if "." in s:  # YY.M / YYYY.M
        parts = s.split(".")
        try:
            y = int(parts[0])
            m = int(parts[1])
        except (ValueError, IndexError):
            return ""
        if y < 100:
            y += 2000
        return f"{y:04d}{m:02d}"
    digits = s.replace("-", "").replace("/", "")
    if len(digits) >= 6:
        return digits[:6]
    if len(digits) == 4:  # YYMM
        return "20" + digits
    return ""


def _agg_complaints(rows, target_ym, zero_fill=False):
    """08 投诉台账 → (complaint_summary, complaints_monthly)。
    summary: {维度: (件数, 金额万元)}；monthly: {YYYYMM: {端口: 件数}}（驱动走势图）。
    zero_fill（2026-10-16 老哥拍板）：接口正常但 0 条 → 全维度显示 0（"无投诉"），不是待填；
    待填仅用于接口故障/无数据源。"""
    if not rows:
        if not zero_fill:
            return {}, {}
        ports = ["政府端", "项目端", "薪莫愁端"]
        summary = {"历史累计": [0.0, 0.0]}
        for p in ports:
            summary[p] = [0.0, 0.0]
        summary["已解决"] = [0.0, 0.0]
        summary["未解决"] = [0.0, 0.0]
        summary["本月新增"] = [0.0, 0.0]
        return summary, {}
    ports = ["政府端", "项目端", "薪莫愁端"]
    summary = {"历史累计": [0.0, 0.0]}
    for p in ports:
        summary[p] = [0.0, 0.0]
    summary["已解决"] = [0.0, 0.0]
    summary["未解决"] = [0.0, 0.0]
    monthly = {}
    _tgt = _norm_ym(target_ym)
    for r in rows:
        n = r.get("people") or 0
        amt = r.get("amount") or 0
        summary["历史累计"][0] += n
        summary["历史累计"][1] += amt
        port = r.get("port") or ""
        if port in summary:
            summary[port][0] += n
            summary[port][1] += amt
        st = str(r.get("status") or "")
        if "未解决" in st:
            summary["未解决"][0] += n
            summary["未解决"][1] += amt
        elif "已解决" in st or "解决" in st:
            summary["已解决"][0] += n
            summary["已解决"][1] += amt
        ym = _norm_ym(r.get("date"))
        if ym:
            monthly.setdefault(ym, {})
            monthly[ym][port or "其他"] = monthly[ym].get(port or "其他", 0) + (n or 1)
    # 本月新增 = target 月条目单列
    if _tgt and _tgt in monthly:
        summary["本月新增"] = [float(sum(monthly[_tgt].values())),
                               float(sum((r.get("amount") or 0) for r in rows if _norm_ym(r.get("date")) == _tgt))]
    else:
        summary["本月新增"] = [0.0, 0.0]
    return summary, monthly


def _agg_fund(fund_rows, output_rows):
    """05 专户进账流水 + 06 产值 → fund_history {YYYYMM: {报审, 审定, 拨付, 专户}}（万元，累计）。"""
    fh = {}
    for r in output_rows or []:
        ym = _norm_ym(r.get("ym"))
        if not ym:
            continue
        rec = fh.setdefault(ym, {"报审": 0.0, "审定": 0.0, "拨付": 0.0, "专户": 0.0})
        rec["报审"] = float(r.get("cum_reported") or 0)
        rec["审定"] = float(r.get("cum_approved") or 0)
    run = 0.0
    for r in sorted(fund_rows or [], key=lambda x: str(x.get("date") or "")):
        ym = _norm_ym(r.get("date"))
        run += float(r.get("amount") or 0)
        rec = fh.setdefault(ym or "unknown", {"报审": 0.0, "审定": 0.0, "拨付": 0.0, "专户": 0.0})
        rec["专户"] = run
    return fh


def health_summary(persons, att_target, target, salary_months, att_months, payload, indicators=None):
    """2.1 健康度汇总表：五项指标（①-⑤）+ 综合评级行（拨付率入综合评级说明）。
    indicators：07 接口状态或 build_indicators 现算（调用方传入，接口优先）。"""
    if indicators is None:
        indicators = build_indicators(persons, att_target, target, salary_months, att_months)
    ok_cnt = sum(1 for i in indicators if i["status"] == "达标")
    # 2026-10-20：状态改用平台原文（达标/限期整改等），综合判定按原文关键词归档
    _warn_kw = ("警告", "限期整改", "未落实", "不达标", "预警")
    if any(any(k in i["status"] for k in _warn_kw) for i in indicators):
        six_status = "警告"
    elif any(i["status"] == "提示" for i in indicators):
        six_status = "提示"
    else:
        six_status = "合格"
    s = payload["summary"]
    fc = payload["four_category"]
    abn = (f"有考勤无工资 {len(fc['attendance_no_salary'])}人；有工资无考勤 {len(fc['salary_no_attendance'])}人；"
           f"超付 {s['over']}人/{fmt_money(s['over_amount'])}；欠付 {s['under']}人/{fmt_money(s['under_amount'])}")
    abn_status = "有异常" if (s['over'] or s['under'] or fc['attendance_no_salary'] or fc['salary_no_attendance']) else "无异常"
    rate_note = PH.need("汇总表", "分账拨付率")
    return [
        ("①", "六项制度综合判定", six_status,
         f"六项中可计算项达标 {ok_cnt}/{len(indicators)}；红/黄项明细见 2.2"),
        ("②", "欠薪事件解决率", PH.need("汇总表", "已解决欠薪占比"), ""),
        ("③", "是否存在非建设单位打款", PH.need("汇总表", "非建设单位打款情况"), ""),
        ("④", "专户余额是否覆盖次月预估工资", PH.need("汇总表", "专户够发判定"),
         PH.need("汇总表", "专户够发说明")),
        ("⑤", "工资发放四项异常指标", abn_status, abn),
        ("综合评级", "", six_status,
         f"分账拨付率 {rate_note}（政策要求：房建20%、市政/水利/电力15%、交通10%）"),
    ]


def render_daily_discrete(doc, persons, yearly=None, monthly=True):
    """3.1.1 当月日工资离散文字分析；yearly 提供时附 3.1.4 全年分布文字分析。
    monthly=False 时只出全年段（2026-10-20 老哥拍板：全年图下不许串当月分析）。"""
    vals = [p.get("daily") for p in persons if p.get("daily")]
    if vals and monthly:
        n = len(vals)
        low = sum(1 for v in vals if v < 150)
        mid = sum(1 for v in vals if 201 <= v <= 300)
        high = sum(1 for v in vals if v > 430)
        para(doc, f"文字分析：当月 {n} 名有薪人员日工资呈双峰离散分布——低端（<150 元/天）{low} 人"
                  f"（{low/n*100:.1f}%），多为杂工/小工与新进场学徒；高端（>430 元/天）{high} 人"
                  f"（{high/n*100:.1f}%），多为架子工、电焊工等技术工种；主体集中在 201～300 元/天"
                  f"（{mid} 人，{mid/n*100:.1f}%）。两端人数较少属正常形态；若某一端人数异常增多，"
                  "须排查压低工价或补贴计入口径问题。", size=9)
    if yearly:
        yv = [d for d, _ in yearly if d is not None]
        if yv:
            b0 = sum(1 for v in yv if v <= 87)
            b1 = sum(1 for v in yv if 201 <= v <= 300)
            para(doc, f"文字分析：全年 {len(yv)} 名有薪人员日工资分布（年实发合计 ÷ 年考勤合计），首区间 0-87 元 {b0} 人"
                      f"（{b0/len(yv)*100:.1f}%），201-300 元区间 {b1} 人（{b1/len(yv)*100:.1f}%）。"
                      "首区间人数若异常增多，须排查新进场学徒集中或压低工价规避分账的行为。", size=9)


def _meta_status(val):
    """平台回填值 → 达标/提示/警告（未回填=待填）。"""
    if val in (None, "", "【待填】"):
        return "待填"
    s = str(val)
    if any(k in s for k in ("未", "异常", "拖欠", "缺口", "违规")):
        return "警告"
    if any(k in s for k in ("部分", "待", "补充", "完善", "提升", "整改", "低于", "不足")):
        return "提示"
    return "达标"


def build_indicators(persons, att_target, target, salary_months, att_months, ds=None):
    """六项制度（HTML 标准件指标名与顺序）：工资保证金存储/劳动合同签署/实名制打卡/
    人工费分账拨付/工资银行代发/离场结算。

    2026-10-16 修复与增强：
    - 劳动合同签署：hasContract 上游返回布尔 True，旧代码比对 "是" 永不相等 → 签约率恒 0.0%。
      改为兼容 True/"true"/"1"/"是" 多形态。
    - 工资银行代发：从当月明细 payWay 字段算代发占比（payWay=1 为专户代发），
      ds 未提供时退回"当月有工资记录即达标"的旧口径。
    """
    total = len(persons) or 1
    _TRUE = ("true", "1", "是")
    # 2026-10-19 新工具：合规快照官方状态优先（met/warning/unmet → 达标/提示/警告），本地推断兜底
    comp = None
    if ds is not None and hasattr(ds, "compliance"):
        try:
            comp = ds.compliance()
        except Exception:
            comp = None
    comp_ind = (comp or {}).get("indicators") or {}
    _CMP = {"met": "达标", "warning": "提示", "unmet": "警告", "unknown": "待填"}

    def _official(key):
        v = comp_ind.get(key) or {}
        # 2026-10-20 老哥拍板：以平台为准。rawStatus 是"是/预警"问答口径，
        # 按标准状态词集转译（是→达标、预警→限期整改、否→严重违规）。
        # 2026-10-22 修复：原"否→不达标"是自造词，色条图颜色表查不到 → 离场结算误显示灰色；
        # "否"即平台最差档，按标准词映射"严重违规"（红）
        _RAW2UI = {"是": "达标", "预警": "限期整改", "否": "严重违规"}
        st = (_RAW2UI.get(str(v.get("raw") or "").strip())
              or _CMP.get(v.get("status"), ""))
        note = str(v.get("note") or "").strip()
        return (st, note) if st else None

    contract_ok = sum(1 for p in persons if str(p.get("contract", "")).strip().lower() in _TRUE)
    contract_any = sum(1 for p in persons if str(p.get("contract", "")).strip())
    _ct_o = _official("劳动合同签署")
    rate = (contract_ok / total) if total else 0
    if _ct_o:
        contract_status = _ct_o[0]
        contract_note = _ct_o[1]  # 2026-10-20：平台原文，不再附加本地自算签约率
    elif contract_any == 0:
        # 2026-10-19：花名册若整体未返回合同字段（降级合成），签约率无从谈起 → 待平台回填，不出"0% 警告"假数字
        contract_status, contract_note = "待填", "花名册未返回合同签署字段（接口降级），待平台回填"
    else:
        contract_status = "达标" if rate >= 1 else ("提示" if rate >= 0.8 else "警告")
        contract_note = f"进场7日内签订并上传平台；签约率 {rate*100:.1f}%"
    _dep_o = _official("保证金存储")
    _spl_o = _official("人工费分账拨付")
    _set_o = _official("离场结算")
    deposit = _dep_o[1] if _dep_o else PH.need("六项指标", "保证金状态")
    deposit_status = _dep_o[0] if _dep_o else _meta_status(deposit)
    split = _spl_o[1] if _spl_o else PH.need("六项指标", "分账拨付状态")
    split_status = _spl_o[0] if _spl_o else _meta_status(split)
    settle = _set_o[1] if _set_o else PH.need("六项指标", "离场结算状态")
    settle_status = _set_o[0] if _set_o else _meta_status(settle)
    # 银行代发：官方状态优先，否则按当月明细 payWay 占比
    _tgt6 = _norm_ym(target)
    _pay_o = _official("工资银行代发")
    payway_note, payway_status = (_pay_o[1], _pay_o[0]) if _pay_o else ("", "待填")
    if not _pay_o and ds is not None and hasattr(ds, "_salary_all"):
        try:
            rows = [it for it in ds._salary_all()
                    if str(it.get("payMonth", "")).replace("-", "") == _tgt6]
            if rows:
                via_bank = sum(1 for it in rows if str(it.get("payWay") or "").strip() in ("1", "true", "True"))
                _r = via_bank / len(rows)
                payway_status = "达标" if _r >= 0.95 else ("提示" if _r >= 0.8 else "警告")
                payway_note = f"当月发放 {len(rows)} 人次，专户银行代发 {via_bank} 人次，代发率 {_r*100:.1f}%"
        except Exception:
            payway_note, payway_status = "", "待填"
    if payway_status == "待填":
        payway_status = "达标" if target in salary_months else "待填"
        payway_note = "工资专户由银行直接发放到农民工本人银行卡" if target in salary_months else ""
    _att_o = _official("实名制打卡")
    if _att_o:
        att_status, att_note = _att_o
    else:
        att_status = "达标" if target in att_months else "待填"
        att_note = "工人人脸实名进场打卡，考勤数据实时上传黔薪保" if target in att_months else ""
    indicators = [
        {"name": "工资保证金存储", "status": deposit_status, "note": deposit},
        {"name": "劳动合同签署", "status": contract_status, "note": contract_note},
        {"name": "实名制打卡", "status": att_status, "note": att_note},
        {"name": "人工费分账拨付", "status": split_status, "note": split},
        {"name": "工资银行代发", "status": payway_status, "note": payway_note},
        {"name": "离场结算", "status": settle_status, "note": settle},
    ]
    return indicators


def _next_month_estimate(att_target, roster_by_nt, price_table):
    """次月预估工资 = Σ[（工种上限+工种下限）÷2 × 考勤]，返回 (总额, 考勤天数)。"""
    tot = 0.0
    days = 0.0
    for a in att_target:
        g = _roster_gong(roster_by_nt, a["name"], a["team"])
        if g and g in price_table and a["days"] > 0:
            lo, hi = price_table[g]
            tot += (lo + hi) / 2 * a["days"]
            days += a["days"]
    return tot, days


def render_four_category(doc, payload):
    """3.1.2 超欠付预警清单：四类汇总表（类型/人数/金额/分布与说明）+ 四类异常逐人名单 + 全年追溯说明。"""
    fc = payload["four_category"]
    s = payload["summary"]
    para(doc, "当月工资整体审核分四类：有考勤无工资、有工资无考勤、超付、欠付。出现任意一类即重点关注。", size=9)
    _no_sal_amt = sum((r.get("est_pay_low") or 0) for r in fc["attendance_no_salary"])
    _no_att_amt = sum((r.get("pay") or 0) for r in fc["salary_no_attendance"])
    rows = [
        ["有考勤无工资", f"{len(fc['attendance_no_salary'])} 人", f"{_no_sal_amt / 10000:,.2f}",
         "重点核查并单独列出，按价格下限×当月黔薪保考勤天数计算工资"],
        ["有工资无考勤", f"{len(fc['salary_no_attendance'])} 人", f"{_no_att_amt / 10000:,.2f}", "重点列出"],
        ["超付（高于预警上限）", f"{s['over']} 人", f"{s['over_amount'] / 10000:,.2f}",
         "日工资高于行情区间上限，计算超出多少钱"],
        ["欠付（低于预警下限）", f"{s['under']} 人", f"{s['under_amount'] / 10000:,.2f}",
         "日工资低于行情区间下限，计算少出多少钱；低于下限者拉全年复核"],
    ]
    add_table(doc, ["类型", "人数", "金额（万元）", "分布与说明"], rows, col_widths=[1.7, 0.9, 1.2, 2.8])

    # —— 异常人员名单（2026-10-20 老哥拍板：四类异常逐人列出，置于 3.1.2）——
    para(doc, "异常人员名单（逐人）", bold=True, size=10)
    para(doc, "注：考勤天数按“人×日期”并集去重计天——同一日多次打卡或跨班组出勤只计 1 天；跨班组人员"
              "的分班构成见备注/说明列（格式：班组：天数），分项之和可能大于合计（重合日只计一次）。", size=8, color="888888")
    _persons = payload.get("persons", [])
    _possible = [r for r in _persons if str(r.get("final") or "").startswith("可能欠付")]

    def _tags(r):
        return "；".join(r.get("tags") or []) or "—"

    def _n(v, nd=2):
        return f"{v:,.{nd}f}" if isinstance(v, (int, float)) else "—"

    def _days_cell(r):
        """考勤天数列：直接显示合计；跨班组人员的分班构成在备注/说明列展开。"""
        return _n(r.get("days"), 0)

    # （1）超付
    _over = fc.get("over") or []
    para(doc, f"（1）超付人员（{len(_over)} 人）：日工资高于行情区间上限 × 1.30", bold=True, size=9)
    if _over:
        add_table(doc,
                  ["序号", "姓名", "班组", "工种", "考勤(天)", "实发(元)", "日薪(元)", "超付金额(元)", "备注"],
                  [[str(i), r.get("name") or "—", r.get("team") or "—", r.get("gong") or "—",
                    _days_cell(r), _n(r.get("pay")), _n(r.get("daily")), _n(r.get("over_amt")),
                    _tags(r)] for i, r in enumerate(_over, 1)],
                  col_widths=[0.35, 0.75, 0.95, 0.7, 0.85, 0.9, 0.65, 0.9, 0.8])
    else:
        para(doc, "本月无超付人员。", size=9, color="888888")

    # （2）欠付 + 可能欠付（v4 hold 口径：低于下限统一“可能欠付(待复核)”，不自动坐实）
    _under_all = list(fc.get("under") or []) + _possible
    para(doc, f"（2）欠付及可能欠付人员（{len(_under_all)} 人）：日工资低于行情区间下限 × 0.85；"
              "按 v4 口径，低于下限者统一列为“可能欠付(待复核)”，不自动坐实欠付",
         bold=True, size=9)
    if _under_all:
        def _u_desc(r):
            _parts = []
            td = r.get("team_days") or {}
            if len(td) > 1:
                _fmt = lambda v: int(v) if v == int(v) else v
                _parts.append("考勤构成：" + "，".join(f"{t}：{_fmt(d)}天" for t, d in sorted(td.items()))
                              + f"（重合日去重后合计{_fmt(r.get('days'))}天）")
            if r.get("under_amt") is not None:
                _parts.append(f"欠付金额 {_n(r.get('under_amt'))} 元")
            if r.get("hist_avg") is not None:
                _parts.append(f"历史均值 {_n(r.get('hist_avg'), 1)} 元/天")
            elif r.get("hist_note"):
                _parts.append(r.get("hist_note"))
            elif r.get("under_amt") is None:
                # 2026-10-22：hold 口径 + 无当年历史（如报告月为 1 月/新人进场）时明说，不留"—"
                _parts.append("无当年历史可比对")
            return "；".join(_parts) or "—"
        add_table(doc,
                  ["序号", "姓名", "班组", "工种", "考勤(天)", "实发(元)", "日薪(元)", "判定", "说明"],
                  [[str(i), r.get("name") or "—", r.get("team") or "—", r.get("gong") or "—",
                    _days_cell(r), _n(r.get("pay")), _n(r.get("daily")),
                    r.get("final") or "—", _u_desc(r)]
                   for i, r in enumerate(_under_all, 1)],
                  col_widths=[0.35, 0.75, 0.9, 0.7, 0.85, 0.85, 0.6, 1.0, 0.95])
    else:
        para(doc, "本月无欠付及可能欠付人员。", size=9, color="888888")

    # （3）有考勤无工资
    _ans = fc.get("attendance_no_salary") or []
    para(doc, f"（3）有考勤无工资人员（{len(_ans)} 人）：当月有黔薪保考勤但工资册无发放记录，"
              "应发按下限 × 考勤天数估算", bold=True, size=9)
    if _ans:
        add_table(doc,
                  ["序号", "姓名", "班组", "工种", "考勤(天)", "按下限估算应发(元)"],
                  [[str(i), r.get("name") or "—", r.get("team") or "—", r.get("gong") or "—",
                    _days_cell(r), _n(r.get("est_pay_low"))] for i, r in enumerate(_ans, 1)],
                  col_widths=[0.5, 1.1, 1.6, 1.2, 1.1, 1.2])
    else:
        para(doc, "本月无有考勤无工资人员。", size=9, color="888888")

    # （4）有工资无考勤
    _sna = fc.get("salary_no_attendance") or []
    para(doc, f"（4）有工资无考勤人员（{len(_sna)} 人）：当月工资册有发放记录但黔薪保无考勤记录", bold=True, size=9)
    if _sna:
        add_table(doc,
                  ["序号", "姓名", "班组", "工种", "实发(元)", "备注"],
                  [[str(i), r.get("name") or "—", r.get("team") or "—", r.get("gong") or "—",
                    _n(r.get("pay")), _tags(r)] for i, r in enumerate(_sna, 1)],
                  col_widths=[0.5, 1.1, 1.5, 1.1, 1.1, 1.3])
    else:
        para(doc, "本月无有工资无考勤人员。", size=9, color="888888")

    # （5）身份一致性异常提示（2026-10-21 老哥要求：同证件号多姓名/同名同姓，供人工复核）
    # 2026-10-21 17:26 老哥两次收紧：只显示当月；掩码造成的异常（工资册同掩码号多姓名，
    # 平台证件号脱敏无法确证是否同人）不出现——保留的是有全号实锤的矩阵笔误归并 + 花名册同名
    _ia = payload.get("identity_anomalies") or {}
    _mi = _ia.get("matrix_id_conflicts") or []
    _dn = _ia.get("roster_dup_names") or []
    # 2026-10-23 老哥拍板：当月同名、一人多班组也写入身份一致性提示
    _mt = _ia.get("month_multi_team") or []
    _od = _ia.get("offline_dup_rows") or []
    _n5 = len(_mi) + len(_dn) + len(_mt) + len(_od)
    para(doc, f"（5）身份一致性异常提示（共 {_n5} 项）："
              "以下情况涉及人员身份唯一性，请人工复核后处置", bold=True, size=9)
    if not _n5:
        para(doc, "本月未发现身份一致性异常（无同证件号多姓名、无同名同姓、无一人多班组）。", size=9, color="888888")
    else:
        def _mask_fid(fid):
            fid = str(fid)
            if "*" in fid:
                return fid
            return fid[:6] + "********" + fid[-4:] if len(fid) == 18 else fid

        def _fm(ym):
            s = str(ym)
            return s[:4] + "年" + s[4:6] + "月" if len(s) == 6 else s

        _rows = []
        for r in _mi:
            _rows.append([str(len(_rows) + 1), "考勤矩阵·同一证件号多个姓名",
                          _mask_fid(r.get("idcard")), "、".join(r.get("names") or []),
                          f"{_fm(r.get('month'))}；平台录入笔误，系统已按证件号归并为同一人"
                          "（考勤天数取各姓名日期并集）"])
        for _n in _dn:
            _rows.append([str(len(_rows) + 1), "花名册同名同姓", "—", _n,
                          "存在多名同姓名人员，按姓名关联的考勤/工资无法区分，需人工核实"])
        for r in _mt:
            _rows.append([str(len(_rows) + 1), "黔薪保考勤·一人多班组", "—", r.get("name") or "—",
                          "当月考勤登记在多个班组：" + "；".join(f"{t} {d:g} 天" for t, d in r.get("teams") or [])
                          + "；需核实系同一人跨班组用工还是同名同姓两人（考勤按人关联时无法自动区分）"])
        for r in _od:
            _rows.append([str(len(_rows) + 1), "工资册·同名多行", "—", r.get("name") or "—",
                          "工资册出现多行：" + "；".join(f"{t} {d:g} 天" for t, d in r.get("rows") or [])
                          + "；需核实系同一人跨班组还是同名同姓两人"])
        add_table(doc,
                  ["序号", "类别", "证件号（掩码）", "涉及姓名", "说明"],
                  _rows, col_widths=[0.4, 1.9, 1.5, 1.6, 2.6])

    para(doc, "全年追溯结果：低于下限人员已拉取全年发放情况核对，全年仍不在区间内的须累计计算低于下限值金额；"
              "涉及合同工价低于行情下限的，列入《低工价合同提示清单》报送建设单位。", size=9, color="888888")


# 待补充清单的板块排序（按报告章节顺序归组）
_SECTION_ORDER = ["封面", "项目概况", "汇总表", "六项指标", "欠薪情况", "资金情况",
                  "工资发放", "工资审核", "人员情况", "其他", "建议与反馈"]


def render_pending_supplement(doc, items):
    """附：待补充数据清单（2026-10-12 规则：凡未拉取到的数据，集中列出并给出补充途径）。
    无待填项时不输出该附录。按板块顺序归组。"""
    if not items:
        return
    items = sorted(items, key=lambda x: (_SECTION_ORDER.index(x["section"])
                                         if x["section"] in _SECTION_ORDER else 99,))
    h(doc, "附：待补充数据清单", 1)
    para(doc, "以下数据本次未能拉取，已在报告对应位置以【待填】占位；请按补充途径补全数据后重新生成，"
              "补全后本附录自动消失。", size=9, color="888888")
    rows = [[str(i), it["section"], it["item"], it["hint"]] for i, it in enumerate(items, 1)]
    add_table(doc, ["序号", "所属板块", "待补充数据项", "补充途径"], rows,
              col_widths=[0.5, 1.2, 1.9, 2.9])


def render_suggestions(doc, payload, meta, est_next, left, rate_note):
    """六、建议与反馈：6.1 对建设单位 / 6.2 对施工单位 / 6.3 五级反馈机制启动情况（数据驱动）。"""
    s = payload["summary"]
    fc = payload["four_category"]
    h(doc, "6.1 对建设单位的建议", 2)
    tips = []
    if isinstance(left, (int, float)) and est_next and left < est_next / 10000:
        tips.append(f"保障次月工资发放资金：专户剩余 {left:,.2f} 万元不足以覆盖次月预估工资 "
                    f"{est_next / 10000:,.2f} 万元（缺口 {est_next / 10000 - left:,.2f} 万元），建议在拨付窗口一次性补足。")
    if "低于" in str(rate_note):
        tips.append(f"补足分账拨付缺口：当前分账拨付率 {rate_note}，低于政策要求，建议按政策比例补足人工费拨付。")
    if s['over'] or s['under'] or fc['attendance_no_salary']:
        tips.append("督促施工单位对超付/欠付预警人员清单逐项核实整改，并对有考勤无工资人员重点核查。")
    if not tips:
        tips.append("本月资金与工资发放未见重大异常，请按常规履行拨款与监管职责。")
    for t in tips:
        para(doc, "· " + t, size=10)
    h(doc, "6.2 对施工单位的建议", 2)
    tips2 = []
    if fc['salary_no_attendance']:
        tips2.append(f"规范工资发放渠道：{len(fc['salary_no_attendance'])} 人有工资无考勤，须统一考勤补录规则，工资一律走专户银行代发。")
    if s['under'] or s['possible_under']:
        tips2.append(f"{s['under'] + s['possible_under']} 人日工资低于预警下限，须按合同与考勤核实工价，存在低工价的补签核增。")
    if fc['attendance_no_salary']:
        tips2.append(f"{len(fc['attendance_no_salary'])} 人有考勤无工资，须逐人说明并限期发放。")
    if not tips2:
        tips2.append("本月工资发放未见违规事项，请保持专户代发与考勤上传的规范性。")
    for t in tips2:
        para(doc, "· " + t, size=10)
    h(doc, "6.3 五级反馈机制启动情况", 2)
    para(doc, "五级反馈（逐级）：劳资员 → 项目经理 → 建设单位现场负责人 → 建设单位集团负责人 → 行管部门。"
              "仅在存在问题但未整改时启动，一级一级汇报；情况据实写，没有则不写。", size=10)
    para(doc, "本月启动情况：" + str(PH.need("建议与反馈", "五级反馈启动情况")), size=10)


def main():
    ap = argparse.ArgumentParser(description="工人工资数据监理服务月报（全量）Word 生成器")
    ap.add_argument("--target", required=True, help="报告月：标准 YYYYMM / 中医药 YY.M")
    ap.add_argument("--data", default=".", help="项目根目录（仅用于定位本地工种单价表，可指向任意目录）")
    ap.add_argument("--engine", default=DEFAULT_ENGINE, help="wage-pay-judge 的 judge.py")
    ap.add_argument("--source", default="qxb",
                    choices=["qxb", "mcp"],
                    help="数据源：qxb=黔薪保平台 MCP（唯一外部数据源，2026-10-13 起）")
    ap.add_argument("--judge-config", default=None, help="判定策略 json（判定口径；不含任何项目数据）")
    ap.add_argument("--history", default=None,
                    help="历史窗口（all=全部历史 / N=最近 N 月）。2026-10-22 老哥拍板：默认=报告年 1 月~报告月前"
                         "（欠付第二轮必须跑，历史日薪只算当年，任何模式一致）；显式传参可覆盖")
    ap.add_argument("--out", default=None, help="输出 docx 路径")
    ap.add_argument("--project-id", default=None,
                    help="qxb 项目标识（32 位项目 ID；或设置环境变量 QXB_PROJECT_ID；留空则从 qxb MCP 会话解析）")
    ap.add_argument("--project-name", default=None,
                    help="qxb 项目名（运行时从当前账号项目列表匹配，避免固化项目 ID）")
    ap.add_argument("--attach-data", default=None,
                    help="附件扫描提取结果 JSON 路径（由 agent 读图后内联传入，不落项目目录）")
    ap.add_argument("--offline-att", default=None,
                    help="3.2.1 线下考勤册 xlsx（月度工资册，工作天数=线下出勤；优先于数据源接口）")
    args = ap.parse_args()

    global _OFFLINE_ATT_PATH
    if args.offline_att:
        _OFFLINE_ATT_PATH = args.offline_att

    out = args.out or os.path.join(args.data, f"数据监理月报_{args.target}.docx")
    # 2026-10-22（老哥拍板）：欠付第二轮必须跑，历史日薪只算当年——默认历史窗口=
    # 报告年 1 月~报告月前（如 26.6 → 最近 5 个月）；显式传 --history 可覆盖
    _ty_h, _tm_h = _split_ym(args.target)
    history = args.history if args.history else str(max(_tm_h - 1, 0))
    attach = None
    if args.attach_data:
        with open(args.attach_data, encoding="utf-8") as f:
            attach = json.load(f)
    path, payload, monthly, yearly = build(
        args.target, args.data, args.engine, args.source, args.judge_config, history,
        None, out, args.project_id, args.project_name, attach
    )
    s = payload["summary"]
    fc = payload["four_category"]
    anchor = payload.get("salary_anchor")
    print(f"月报已生成：{path}")
    if anchor:
        if anchor.get("ok"):
            print(f"  对账锚：明细 ¥{anchor['detail_sum']:,.2f}（{anchor['detail_count']} 人）"
                  f"= 批次 ¥{anchor['batch_sum']:,.2f}（{anchor['batch_count']} 笔）✓ 对平")
        else:
            print(f"  ⚠️ 对账不平：明细 ¥{anchor['detail_sum']:,.2f}（{anchor['detail_count']} 人）"
                  f"≠ 批次 ¥{anchor['batch_sum']:,.2f}（{anchor['batch_count']} 笔）——报告内已标注，须复核")
    print(f"  超付 {s['over']} 人 / 欠付 {s['under']} 人 / 可能欠付 {s['possible_under']} 人 / "
          f"有考勤无工资 {len(fc['attendance_no_salary'])} 人 / 有工资无考勤 {len(fc['salary_no_attendance'])} 人")
    print(f"  待补充项（【待填】）共 {len(PH.items)} 类（报告末尾附清单）")
    print(f"  工资月份 {len([m for m in monthly if m['has_salary']])} 个，考勤月份 {len([m for m in monthly if m['has_att']])} 个")


if __name__ == "__main__":
    main()
