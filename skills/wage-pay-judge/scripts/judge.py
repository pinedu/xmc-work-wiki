#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工资超付/欠付两轮判定引擎 (wage-pay-judge)
=========================================
判定逻辑（建设单位/甲方视角，面向农民工工资合规）：

  对每一位当月发放工资的工人：
    日薪 = 当月实付 ÷ 当月考勤天数

  第一轮（工种市场单价区间卡）：
    日薪 > 工种[上限] × over_mult(默认1.30)        -> 超付
    日薪 < 工种[下限] × under_ratio(默认0.85)       -> 可能欠付（进入第二轮）
    介于两者之间                                -> 正常

  第二轮（仅“可能欠付”人员）：
    history_window=None 时取该工人【全部工作历史】（分析月之前所有可用月份）的“实付÷考勤”均值
    second_round=seat  (标准) ：历史均值 < 下限 -> 欠付（坐实）；≥下限 -> 可能欠付(历史未佐证·待复核)；无历史 -> 可能欠付(无历史数据·待复核)
    second_round=hold  (v4)   ：一律维持“可能欠付(待复核)”，不自动坐实

数据来源（通过 DataSource 适配器注入，不再写死路径）：
  - QxbMcpDataSource : 黔薪保平台 MCP（streamable HTTP）——唯一外部数据源（2026-10-13 起，
    standard/zhongyi/db 文件源与 HTTP 接口源已移除；连接与 token 复用 WorkBuddy mcp.json 的 qxb server）

用法：
  python judge.py --target 26.8 --data <项目根目录> --source qxb --config judge_config.json
"""

# 注：--config 仅承载判定口径（阈值/工种解析等），不含任何项目数据；project_id 在运行时
# 由 qxb MCP 会话解析（--project-name 或单项目自动），不固化进配置。

import argparse
import os
import json
import statistics
from collections import defaultdict

import openpyxl

from datasources import make_datasource, team_to_gong, DEFAULT_TEAM_TO_GONG

# ----------------------------- 兼容常量（render_report 仍引用默认值） -----------------------------
OVER_MULT = 1.30       # 超付倍数：日薪 > 上限×1.30
UNDER_RATIO = 0.85     # 欠付下限比：日薪 < 下限×0.85


# ----------------------------- 判定策略配置（可注入，不再写死） -----------------------------
class JudgeConfig:
    """判定口径策略对象。所有阈值/规则均可由外部 config json 覆盖。默认即标准口径（零破坏）。"""

    def __init__(self, **kw):
        self.over_mult = kw.get("over_mult", 1.30)                 # 超付倍数
        self.under_ratio = kw.get("under_ratio", 0.85)             # 欠付下限比
        # 阈值模式：warn=预警值（上限×over_mult / 下限×under_ratio）；hard=硬区间（直接上限/下限）
        self.threshold_mode = kw.get("threshold_mode", "warn")
        # 第二轮：seat=坐实（历史均值<下限则欠付）；hold=一律待复核（v4 口径，不自动坐实）
        self.second_round = kw.get("second_round", "seat")
        # 工种解析：roster_only=花名册唯一即用，空/其它→缺工种；team_fallback=冲突/多/空→按班组反推
        self.gongzhong_resolve = kw.get("gongzhong_resolve", "roster_only")
        self.team_gong_map = kw.get("team_gong_map", DEFAULT_TEAM_TO_GONG)
        self.attendance_source = kw.get("attendance_source", "kaoqin")
        self.alias_path = kw.get("alias_path")
        self.history_window = kw.get("history_window")   # None=全部历史
        self.price = kw.get("price")

    @classmethod
    def from_file(cls, path):
        with open(path, encoding="utf-8") as f:
            return cls(**json.load(f))

    @classmethod
    def standard(cls):
        return cls()


# ----------------------------- 工具函数 -----------------------------
def find_col(header, candidates, exact=False):
    for c in candidates:
        for i, h in enumerate(header):
            if h is None:
                continue
            hs = str(h).strip()
            if exact:
                if hs == c:
                    return i
            else:
                if c in hs:
                    return i
    return None


def find_header_row(rows, candidates):
    best, best_score = 0, 0
    for idx, r in enumerate(rows[:5]):
        score = sum(1 for c in candidates if any(c in str(v) for v in r if v is not None))
        if score > best_score:
            best, best_score = idx, score
    return best


def to_num(v):
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def load_xlsx(path, sheet=None):
    # 勿用 read_only=True：依赖文件内声明的维度标记，WPS/Excel 编辑后可能陈旧导致只读前几行。
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    sh = sheet if (sheet and sheet in wb.sheetnames) else wb.sheetnames[0]
    ws = wb[sh]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows


def load_price_table(price_path):
    """工种单价区间：{工种: (底线, 上线)}。"""
    rows = load_xlsx(price_path)
    hdr_idx = find_header_row(rows, ["工种", "底线", "上线"])
    hdr = rows[hdr_idx]
    ci = find_col(hdr, ["工种"], exact=True)
    cl = find_col(hdr, ["底线"])
    ch = find_col(hdr, ["上线"])
    table = {}
    for r in rows[hdr_idx + 1:]:
        name = str(r[ci]).strip() if (ci is not None and r[ci]) else ""
        if not name:
            continue
        lo = to_num(r[cl])
        hi = to_num(r[ch])
        if lo == 0 and hi == 0:
            continue
        table[name] = (lo, hi)
    return table


def load_alias(alias_path):
    alias = {}
    if not alias_path or not os.path.exists(alias_path):
        return alias
    with open(alias_path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line.startswith("|"):
                continue
            cols = [c.strip() for c in line.strip("|").split("|")]
            if len(cols) < 3:
                continue
            src, tgt, typ = cols[0], cols[1], cols[2]
            if src in ("源工种名", "") or "精确匹配" in typ:
                continue
            if typ == "别名" and src and tgt:
                alias[src] = tgt
    return alias


def resolve_gongzhong(name, price_table, alias):
    """工种名 -> 单价表工种（先别名，再精确匹配）。返回 (canonical, matched)。"""
    if name in alias:
        tgt = alias[name]
        return tgt, tgt in price_table
    if name in price_table:
        return name, True
    return name, False


# ----------------------------- 工种解析辅助（v4 多候选 + 班组反推） -----------------------------
def _candidate_gongs(by_id, key):
    """从花名册索引收集有效工种（排除空；保留“其它”作兜底标志）。
    返回 (valid_gongs, has_other)。key 可能是身份证或 姓名|班组。"""
    cands = by_id.get(key)
    if cands is None:
        return [], False
    if isinstance(cands, dict):
        cands = [cands]
    valid, has_other = [], False
    for c in cands:
        g = (c.get("gong") or c.get("gongzhong") or "")
        if g == "其它":
            has_other = True
        elif g:
            valid.append(g)
    return list(dict.fromkeys(valid)), has_other


def _resolve_multi(valid_gongs, has_other, price_table, config, team, alias):
    # 别名优先（仅对有效工种）
    for g in valid_gongs:
        if g in alias:
            tgt = alias[g]
            if tgt in price_table:
                return tgt, True
    if len(valid_gongs) == 1:
        g = valid_gongs[0]
        return g, g in price_table
    if config.gongzhong_resolve == "team_fallback":
        # 冲突/多个/空 -> 按班组名反推（v4 口径）
        mapped = team_to_gong(team, config.team_gong_map)
        if mapped:
            return mapped, mapped in price_table
        if valid_gongs:
            g = valid_gongs[0]
            return g, g in price_table
        # 反推失败且只有“其它”：v4 不回退“其它”，视作缺工种
        return None, False
    # roster_only（标准）：有效工种取第一个；若仅有“其它”且单价表含“其它”则兜底命中
    if valid_gongs:
        g = valid_gongs[0]
        return g, g in price_table
    if has_other and "其它" in price_table:
        return "其它", True
    return None, False


# ----------------------------- 判定核心 -----------------------------
def judge(config, ds, target_ym):
    price_table = ds.price_table()
    alias = load_alias(config.alias_path)
    persons_list, by_id, by_nt = ds.roster()
    # 2026-10-21（老哥拍板"用身份证"）：工资明细证件为掩码(前6+后4)、花名册为全号 →
    # 建掩码索引做夹逼匹配；花名册同名单独告警（姓名锚的风险护栏）
    by_masked = {}
    for _cs in by_id.values():
        for _c in (_cs if isinstance(_cs, list) else [_cs]):
            _fid = str(_c.get("id") or "").strip()
            if len(_fid) == 18:
                # 键=前6|后4（星号数无关：平台实际掩码为 12 星，勿写死 8 星）
                by_masked.setdefault(_fid[:6] + "|" + _fid[-4:], []).append(_c)
    _name_cnt = {}
    for _p in persons_list:
        if _p.get("name"):
            _name_cnt[_p["name"]] = _name_cnt.get(_p["name"], 0) + 1
    dup_names = {n for n, c in _name_cnt.items() if c > 1}
    if dup_names:
        print(f"[warn] 花名册存在同名同姓 {len(dup_names)} 人（{'、'.join(sorted(dup_names))}）"
              "——姓名锚匹配须人工复核", flush=True)
    hist, prior = ds.history(target_ym, config.history_window)
    sal = ds.salary_agg(target_ym)
    # 对账锚（2026-10-16 新增）：明细合计 vs 批次总额，对不平则数据不完整，判定结论不可信
    anchor = None
    if hasattr(ds, "salary_anchor_check"):
        try:
            anchor = ds.salary_anchor_check(target_ym)
        except Exception:
            anchor = None
    by_nt_k, by_name_k, id_days_k = ds.kaoqin(target_ym)
    # 2026-10-19（老哥拍板"用班组映射反推工种"）：工资明细不带班组 → 从考勤键 (name, team) 找回
    att_team = {}
    for (_nm, _tm) in by_nt_k.keys():
        if _nm and _tm:
            att_team.setdefault(_nm, set()).add(_tm)

    # —— 反向缺失：有考勤无工资（供月报“工资发放情况”四类使用）——
    sal_keys = {(r["name"], r["team"]) for r in sal.values()}
    sal_names = {r["name"] for r in sal.values() if r["name"]}
    name_gong = {p["name"]: p.get("gong") for p in persons_list if p.get("name")}
    att_only = []
    for (name, team), days in by_nt_k.items():
        if not days or days <= 0:
            continue
        if (name, team) in sal_keys or name in sal_names:
            continue
        gong = name_gong.get(name)
        ag, aoth = ([gong] if (gong and gong != "其它") else [], gong == "其它")
        gong_c, gong_ok = _resolve_multi(ag, aoth, price_table, config, team, alias)
        low = price_table[gong_c][0] if gong_ok else None
        att_only.append({
            "name": name, "team": team, "days": days,
            "gong": gong_c, "gong_ok": gong_ok,
            "est_pay_low": round(low * days, 2) if low is not None else None,
        })

    results = []
    for key, rec in sal.items():
        idc, name, team = rec["idcard"], rec["name"], rec["team"]
        idc_masked = rec.get("idcard_masked") or ""  # 掩码证件号（前6+后4），用于夹逼锚定
        pay = rec["pay"]

        # 2026-10-20（老哥拍板）：工资明细接口（qxb_salary_details_list）无班组字段，工资侧班组全空；
        # 班组只存在于考勤键 (name, team) → 从考勤回填：唯一班组直接填，多班组用斜杠并列
        _sal_team_empty = not team
        _tb_teams = sorted(att_team.get(name) or set())
        _team_backfilled = ""
        if _sal_team_empty and _tb_teams:
            team = "/".join(_tb_teams)
            _team_backfilled = "班组按考勤回填"

        # 考勤天数（salary_agg 不含 days，从 kaoqin 取）
        days = by_nt_k.get((name, team))
        att_note = ""
        if days is None:
            days = by_name_k.get(name)
            if days is not None:
                att_note = ("工资册无班组·按姓名匹配考勤" if _sal_team_empty
                            else "按姓名匹配(班组不一致)")
        if days is None and idc_masked and id_days_k:
            # 2026-10-21（老哥发现 李红海/李洪海 同一证件号双姓名）：姓名 join 断链时
            # 用工资掩码号夹逼矩阵全号直接取天数——姓名只作显示，不参与锚定
            _same = [fid for fid in id_days_k
                     if fid[:6] == idc_masked[:6] and fid[-4:] == idc_masked[-4:]]
            if len(_same) == 1:
                days = id_days_k[_same[0]]
                att_note = "考勤按证件号匹配"
            elif len(_same) > 1:
                att_note = "考勤证件号多候选·需人工核实"
        if days is None and not att_note:
            att_note = "缺考勤"
        daily = (pay / days) if (days not in (None, 0) and days > 0) else None

        # 2026-10-21（老哥拍板）：跨班组人员考勤天数按班组拆明细，写入 team_days，
        # 供月报标注“相加构成”（如 混凝土工22 + 水电班组20 = 42），避免误读为单月实际出勤
        _team_days = {}
        for _t in sorted(att_team.get(name) or set()):
            _d = by_nt_k.get((name, _t))
            if _d:
                _team_days[_t] = _d

        # 工种解析
        valid_gongs, has_other = _candidate_gongs(by_id, idc)
        if not valid_gongs and not has_other and name:
            valid_gongs, has_other = _candidate_gongs(by_id, name + "|" + team)  # 标准布局 name|team 键
        # 2026-10-21：掩码证件号夹逼匹配（唯一候选才采信，多候选自动弃用）——同名同姓也可锚定；
        # 工资掩码号 vs 花名册全号此前判"无法互 join"过于悲观，前6+后4夹逼在百人级项目近乎唯一
        _gong_via = "id"
        _idc_mismatch = False
        if not valid_gongs and not has_other and idc_masked:
            _cands = by_masked.get(idc_masked[:6] + "|" + idc_masked[-4:]) or []
            if _cands:
                # 2026-10-21：掩码锚必须"证件匹配+姓名一致"双保险——实测存在掩码撞车
                # （同区县+同后4 的两个真人，5/105 对完全不同姓名），单靠掩码会挂错工种
                _nm_hit = [c for c in _cands if str(c.get("name") or "") == name]
                if len(_nm_hit) == 1:
                    _g = str(_nm_hit[0].get("gong") or "").strip()
                    if _g:
                        valid_gongs, has_other = (([_g] if _g != "其它" else []), _g == "其它")
                        _gong_via = "idcard"
                else:
                    _idc_mismatch = True  # 证件号与姓名不符：宁缺毋滥，交人工核实
        if not valid_gongs and not has_other and name:
            _g = name_gong.get(name)
            if _g:
                valid_gongs, has_other = (([_g] if _g != "其它" else []), _g == "其它")
                _gong_via = "name"
        # 班组反推（2026-10-19 老哥拍板）：工种仍缺失或花名册登记为"其它"时，
        # 用考勤找回的班组名按 班组→工种 映射反推（工资明细不带班组，考勤带）
        if (not valid_gongs or has_other) and name:
            _mg = None
            for _tm in sorted(att_team.get(name) or set()):
                _mg = team_to_gong(_tm, config.team_gong_map)
                if _mg:
                    break
            if _mg:
                valid_gongs, has_other = [_mg], False
                _gong_via = "team:" + "/".join(sorted(att_team.get(name) or set()))
        gong, gong_ok = _resolve_multi(valid_gongs, has_other, price_table, config, team, alias)
        low = price_table[gong][0] if gong_ok else None
        high = price_table[gong][1] if gong_ok else None

        # 花名册状态（离场标签等）——掩码锚同样要求姓名一致
        status = None
        cands = by_id.get(idc)
        if not cands and idc_masked:
            cands = [c for c in (by_masked.get(idc_masked[:6] + "|" + idc_masked[-4:]) or [])
                     if str(c.get("name") or "") == name]
        if isinstance(cands, dict):
            cands = [cands]
        for c in (cands or []):
            if c.get("status"):
                status = c["status"]
                break

        tags = []
        if rec.get("short") and "是" in rec["short"]:
            tags.append("短期工")
        if rec.get("rows") and rec["rows"] > 1:
            tags.append(f"工资册重复记录×{rec['rows']}")
        if _team_backfilled:
            tags.append(_team_backfilled)
        if att_note:
            tags.append(att_note)
        if len(_team_days) > 1:
            _fmt = lambda v: int(v) if v == int(v) else v
            _td_parts = "，".join(f"{t}：{_fmt(d)}天" for t, d in sorted(_team_days.items()))
            _td_sum = sum(_team_days.values())
            if days is not None and _td_sum > days:
                tags.append(f"考勤跨{len(_team_days)}班组（{_td_parts}；重合日去重后合计{_fmt(days)}天）")
            else:
                tags.append(f"考勤跨{len(_team_days)}班组（{_td_parts}，合计{_fmt(_td_sum)}天）")
        if _gong_via == "name" and (valid_gongs or has_other):
            tags.append("工种按姓名匹配" + ("（花名册同名·需人工区分）" if name in dup_names else ""))
        elif _gong_via == "idcard" and valid_gongs:
            tags.append("工种按证件号匹配")
        elif _idc_mismatch:
            tags.append("证件号与花名册姓名不符·需人工核实")
        elif _gong_via.startswith("team:") and valid_gongs:
            tags.append(f"工种按班组反推（{_gong_via.split(':', 1)[1]}）")
        if status and "离场" in status:
            tags.append("已离场")
        if not valid_gongs and not has_other:
            tags.append("花名册缺工种")
        elif not gong_ok:
            tags.append("未匹配工种")

        # 阈值（按 threshold_mode）
        if config.threshold_mode == "hard":
            under_th, over_th = low, high
        else:
            under_th = (low * config.under_ratio) if gong_ok else None
            over_th = (high * config.over_mult) if gong_ok else None

        r1 = None
        final = None
        reason = ""
        over_amt = under_amt = None
        if daily is None:
            final = "无法计算"
            reason = att_note or "考勤为0/缺考勤"
        elif not gong_ok:
            final = "缺工种·需人工"
            reason = "花名册无工种" if not valid_gongs and not has_other else "工种未匹配单价表"
        else:
            if daily > over_th:
                r1 = "超付"
                over_amt = round(pay - over_th * days, 2)
            elif daily < under_th:
                r1 = "可能欠付"
            else:
                r1 = "正常"
            final = r1

        hist_avg = hist_note = None
        if r1 == "可能欠付":
            hv = hist.get(idc, [])
            if hv:
                hist_avg = round(statistics.mean(hv), 2)
            if config.second_round == "hold":
                final = "可能欠付(待复核)"
                reason = "低于下限×%.2f" % config.under_ratio
                if hist_avg is not None:
                    reason += "·历史均值%.1f" % hist_avg
                else:
                    reason += "·无历史/未佐证"
            else:
                if hist_avg is not None and hist_avg < low:
                    final = "欠付"
                    under_amt = round(low * days - pay, 2)
                elif hist_avg is not None:
                    final = "可能欠付(历史未佐证·待复核)"
                else:
                    final = "可能欠付(无历史数据·待复核)"
                    hist_note = "该工人无可用历史月份"

        results.append({
            "name": name, "team": team, "gong": gong, "status_label": status,
            "days": days, "pay": round(pay, 2), "daily": round(daily, 2) if daily else None,
            "low": low, "high": high, "under_th": round(under_th, 2) if under_th else None,
            "over_th": round(over_th, 2) if over_th else None,
            "r1": r1, "hist_avg": hist_avg, "hist_note": hist_note,
            "final": final, "over_amt": over_amt, "under_amt": under_amt,
            "tags": tags, "att_note": att_note, "team_days": _team_days,
        })

    # 2026-10-21：身份一致性异常（同证件号多姓名/同名同姓）——供月报人工复核清单
    # （17:12 老哥要求：只显示当月的，历史月份不进报告）
    ym6 = ds._norm(target_ym)
    ds._identity_anomalies = {
        "roster_dup_names": sorted(dup_names),
        "salary_mask_conflicts": [
            {"month": _ym, "masked": _mk, "names": _ns}
            for _ym, _mp in sorted(getattr(ds, "_mask_conflicts", {}).items()) if _ym == ym6
            for _mk, _ns in sorted(_mp.items())],
        "matrix_id_conflicts": [
            {"month": _ym, "idcard": _idc, "names": _ns}
            for _ym, _mp in sorted(getattr(ds, "_idc_name_conflicts", {}).items()) if _ym == ym6
            for _idc, _ns in sorted(_mp.items())
            # 2026-10-21 17:45 老哥要求：只列当月有出勤的笔误对，零出勤的名单残留不占报告
            if (id_days_k.get(_idc) or 0) > 0],
    }

    return results, att_only, prior, len(price_table), anchor


# ----------------------------- 报告 -----------------------------
STATUS_ORDER = ["超付", "欠付", "可能欠付", "正常", "无法计算", "缺工种·需人工"]


def main_status(s):
    """提取主状态：去掉括号后缀。"""
    return s.split("(")[0].strip()


def render_report(target_ym, results, prior, price_count, out_path, config=None, anchor=None):
    counts = defaultdict(int)
    over_sum = under_sum = 0.0
    for r in results:
        counts[main_status(r["final"])] += 1
        if r["over_amt"]:
            over_sum += r["over_amt"]
        if r["under_amt"]:
            under_sum += r["under_amt"]

    om = config.over_mult if config else OVER_MULT
    ur = config.under_ratio if config else UNDER_RATIO
    tm = config.threshold_mode if config else "warn"

    lines = []
    lines.append(f"# 工人工资超付/欠付判定报告（{target_ym}）\n")
    lines.append("## 一、判定口径")
    lines.append(f"- 分析月：**{target_ym}**")
    lines.append(f"- 日薪公式：**当月实付 ÷ 当月考勤天数**（实发取“实付”列）")
    if tm == "hard":
        lines.append(f"- 第一轮阈值（硬区间）：日薪 > 工种**上限** → 超付；日薪 < 工种**下限** → 可能欠付；区间之间 → 正常")
    else:
        lines.append(f"- 第一轮阈值（预警值）：日薪 > 工种**上限 × {om}** → 超付；日薪 < 工种**下限 × {ur}** → 可能欠付；区间之间 → 正常")
    lines.append(f"- 第二轮（仅“可能欠付”）：历史平均日薪 = **该工人全部工作历史**（分析月之前所有可用月份，本次取 {prior}）的“实付÷考勤”均值")
    if (config.second_round if config else "seat") == "hold":
        lines.append(f"  - 本口径（hold）：低于下限者统一维持“可能欠付(待复核)”，不自动坐实；无历史 → “可能欠付(无历史数据·待复核)”")
    else:
        lines.append(f"  - 历史均值 < 工种**下限** → **欠付**（坐实）；≥ 下限 → 维持“可能欠付(历史未佐证·待复核)”；无历史 → “可能欠付(无历史数据·待复核)”")
    lines.append(f"- 工种单价参照：内置《工种单价预估范围》（共 {price_count} 个工种区间）")
    lines.append(f"- 关联链：工资册↔花名册 用身份证；工资册↔考勤 用 (姓名,班组)；工种名→单价区间先查别名映射再精确匹配")
    if anchor is not None:
        if anchor.get("ok"):
            lines.append(f"- 对账锚：✅ 明细合计 ¥{anchor['detail_sum']:,.2f}（{anchor['detail_count']} 人）"
                         f"= 工资册批次总额 ¥{anchor['batch_sum']:,.2f}（{anchor['batch_count']} 笔），数据完整")
        else:
            lines.append(f"- **⚠️ 对账不平（数据不完整）**：明细合计 ¥{anchor['detail_sum']:,.2f}"
                         f"（{anchor['detail_count']} 人）≠ 批次总额 ¥{anchor['batch_sum']:,.2f}"
                         f"（{anchor['batch_count']} 笔）。本报告判定结论**仅作参考，不可直接采用**，"
                         f"须先排查取数完整性（上游分页/月份过滤故障）后重跑")
    lines.append("")

    lines.append("## 二、汇总")
    lines.append(f"- 判定总人数（当月有工资发放）：**{len(results)}**")
    lines.append(f"- 超付：**{counts['超付']}** 人，估算超付金额合计 **¥{round(over_sum,2):,.2f}**")
    lines.append(f"- 欠付：**{counts['欠付']}** 人，估算欠付金额合计 **¥{round(under_sum,2):,.2f}**")
    lines.append(f"- 可能欠付（待复核）：**{counts['可能欠付']}** 人")
    lines.append(f"- 正常：**{counts['正常']}** 人")
    lines.append(f"- 无法计算（缺考勤/考勤为0）：**{counts['无法计算']}** 人")
    lines.append(f"- 缺工种·需人工：**{counts['缺工种·需人工']}** 人\n")

    lines.append("## 三、逐人明细")
    header = ["姓名", "班组", "工种", "考勤天", "实付", "日薪", "下限×0.85", "上限×1.30",
              "第一轮", "历史均值", "下限", "最终状态", "标签/备注"]
    lines.append("| " + " | ".join(header) + " |")
    lines.append("|" + "|".join(["---"] * len(header)) + "|")

    def sort_key(r):
        try:
            return STATUS_ORDER.index(main_status(r["final"]))
        except ValueError:
            return 99

    for r in sorted(results, key=sort_key):
        note = "；".join(r["tags"]) if r["tags"] else ""
        if r["hist_note"]:
            note = (note + "；" + r["hist_note"]) if note else r["hist_note"]
        lines.append("| " + " | ".join([
            str(r["name"]), str(r["team"]), str(r["gong"] or "—"),
            str(r["days"]) if r["days"] is not None else "—",
            f'{r["pay"]:,.2f}' if r["pay"] is not None else "—",
            f'{r["daily"]:,.2f}' if r["daily"] is not None else "—",
            f'{r["under_th"]:,.2f}' if r["under_th"] is not None else "—",
            f'{r["over_th"]:,.2f}' if r["over_th"] is not None else "—",
            str(r["r1"] or "—"),
            f'{r["hist_avg"]:,.2f}' if r["hist_avg"] is not None else "—",
            f'{r["low"]:,.2f}' if r["low"] is not None else "—",
            str(r["final"]), note,
        ]) + " |")
    lines.append("")
    lines.append("## 四、待确认清单")
    pending = [r for r in results if r["final"] in ("无法计算", "缺工种·需人工")]
    if not pending:
        lines.append("- 无（所有人员均已正常完成判定）")
    else:
        for r in pending:
            lines.append(f"- {r['name']}（{r['team']}）：{r['final']} —— {'；'.join(r['tags'])}")
    lines.append("")
    lines.append("> 说明：超付金额 = 实付 − 上限×1.30×考勤天数；欠付金额 = 下限×考勤天数 − 实付。"
                 "金额仅按区间阈值的粗估，用于排序与预警，不作为最终结算依据。")

    text = "\n".join(lines)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(text)
    return text, counts, over_sum, under_sum


# ----------------------------- 四象限分类（月报“工资发放情况”用） -----------------------------
def build_four_category(results, att_only):
    def ms(s):
        return s.split("(")[0].strip()

    attendance_no_salary = att_only
    salary_no_attendance = [r for r in results if r["final"] == "无法计算"]
    over = [r for r in results if r["final"] == "超付"]
    under = [r for r in results if ms(r["final"]) == "欠付"]
    return {
        "attendance_no_salary": attendance_no_salary,
        "salary_no_attendance": salary_no_attendance,
        "over": over,
        "under": under,
    }


# ----------------------------- 入口 -----------------------------
def main():
    ap = argparse.ArgumentParser(description="工资超付/欠付两轮判定")
    ap.add_argument("--target", required=True, help="分析月：标准 YYYYMM / 中医药 YY.M")
    ap.add_argument("--data", default=".", help="项目根目录（仅用于定位本地工种单价表，可指向任意目录）")
    ap.add_argument("--source", default="qxb", choices=["qxb", "mcp"],
                    help="数据源：qxb=黔薪保平台 MCP（唯一外部数据源，2026-10-13 起）")
    ap.add_argument("--config", default=None, help="判定策略 json（覆盖 JudgeConfig）")
    ap.add_argument("--price", default=None, help="工种单价 xlsx 路径")
    here = os.path.dirname(os.path.abspath(__file__))
    ap.add_argument("--alias", default=os.path.join(here, "..", "references", "工种别名映射.md"),
                    help="工种别名映射 md 路径")
    ap.add_argument("--history", default="all", help="历史窗口：all=全部历史，或整数N=最近N月")
    ap.add_argument("--out", default=None, help="报告输出 md 路径")
    ap.add_argument("--json", default=None, help="结构化结果 JSON 输出路径（供月报组装层调用）")
    ap.add_argument("--project-id", default=None,
                    help="qxb 项目标识（32 位项目 ID；或设置环境变量 QXB_PROJECT_ID；留空则从 qxb MCP 会话解析）")
    ap.add_argument("--project-name", default=None,
                    help="qxb 项目名（运行时从当前账号项目列表匹配，避免固化项目 ID）")
    args = ap.parse_args()

    if args.config:
        cfg = JudgeConfig.from_file(args.config)
    else:
        cfg = JudgeConfig.standard()
    cfg.alias_path = args.alias
    cfg.history_window = None if args.history in ("all", "") else int(args.history)
    if args.price:
        cfg.price = args.price
    # 运行时 project_id 注入顶层配置（make_datasource 的 qxb 分支读取）
    if args.project_id:
        cfg.project_id = args.project_id

    # 原始配置透传（project_id / qxb 连接段不在 JudgeConfig 字段内，工厂需要）
    raw_extra = {}
    if args.config and os.path.exists(args.config):
        try:
            with open(args.config, encoding="utf-8") as _f:
                raw_extra = json.load(_f)
        except Exception:
            raw_extra = {}
    ds_cfg = dict(cfg.__dict__)
    for k in ("project_id", "qxb"):
        if raw_extra.get(k):
            ds_cfg[k] = raw_extra[k]
    if args.project_name:
        ds_cfg["project_name"] = args.project_name

    ds = make_datasource(args.source, args.data, ds_cfg)
    out = args.out or os.path.join(args.data, f"样例报告_{args.target}.md")

    results, att_only, prior, pc, anchor = judge(cfg, ds, args.target)
    text, counts, over_sum, under_sum = render_report(args.target, results, prior, pc, out, cfg,
                                                      anchor=anchor)
    four = build_four_category(results, att_only)

    if args.json:
        payload = {
            "target": args.target,
            "source": args.source,
            "prior_months": prior,
            "price_count": pc,
            "salary_anchor": anchor,
            "summary": {
                "total_workers": len(results),
                "over": counts["超付"], "under": counts["欠付"],
                "possible_under": counts["可能欠付"],
                "normal": counts["正常"],
                "cannot_compute": counts["无法计算"],
                "missing_gongzhong": counts["缺工种·需人工"],
                "over_amount": round(over_sum, 2),
                "under_amount": round(under_sum, 2),
            },
            "four_category": four,
            "persons": results,
            "attendance_only": att_only,
            "identity_anomalies": getattr(ds, "_identity_anomalies", {}),
        }
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print(f"结构化结果已写出：{args.json}")

    print(f"分析月 {args.target}｜历史月份 {prior}")
    if anchor is not None:
        if anchor.get("ok"):
            print(f"对账锚：明细 {anchor['detail_count']} 人 ¥{anchor['detail_sum']:,.2f} "
                  f"= 批次 {anchor['batch_count']} 笔 ¥{anchor['batch_sum']:,.2f} ✓ 对平")
        else:
            print(f"⚠️ 对账不平：明细 {anchor['detail_count']} 人 ¥{anchor['detail_sum']:,.2f} "
                  f"≠ 批次 {anchor['batch_count']} 笔 ¥{anchor['batch_sum']:,.2f} —— 数据不完整，"
                  f"判定结论仅作参考，须复核后使用！")
    print(f"判定人数 {len(results)}｜超付 {counts['超付']}(¥{over_sum:,.2f})｜"
          f"欠付 {counts['欠付']}(¥{under_sum:,.2f})｜可能欠付 {counts['可能欠付']}｜"
          f"正常 {counts['正常']}｜无法计算 {counts['无法计算']}｜缺工种 {counts['缺工种·需人工']}")
    print(f"报告已写出：{out}")


if __name__ == "__main__":
    main()
