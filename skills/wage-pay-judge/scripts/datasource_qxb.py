# -*- coding: utf-8 -*-
"""
QxbMcpDataSource（黔薪保平台 MCP 数据源）
==========================================
数据源第 5 种实现：经 qxb MCP（streamable HTTP，JSON-RPC 2.0）从黔薪保平台取数。
协议客户端用标准库 urllib 实现（initialize + tools/call），不引入 MCP SDK 依赖。

连接配置（judge_config 的 `qxb` 段，或内置默认）：
- 默认从 WorkBuddy 的 mcp.json（~/.workbuddy/mcp.json）读取 `qxb` server 的 url + headers，
  复用当前登录用户已配置的连接与 token，不在第二处落盘。
- 也可在 qxb 段直接给 url / headers 覆盖。

**项目解析（重要，2026-10-14 起）**：project_id 一律在运行时由**当前登录用户的 qxb 会话**解析，
绝不固化进任何配置/脚本——`qxb_projects_page` 列出该账号名下的项目，按 `--project-name` 匹配
或单项目自动选用；多个项目未指定则报错列出可选全称。这样 skill 多用户/多项目复用时不串数据。

降级约定（2026-10-13 老哥拍板）：凡 MCP 端点拉取不到的数据（未配置/上游故障/平台无数据），
一律返回空 → 月报对应板块就地【待填】并进《待补充数据清单》，流程不中断。

实测状态（2026-10-14 重探+接口对齐，运行时按当前账号会话）：
- ✅ qxb_projects_page（富档案：类别/开工/完工/五方ID）/ qxb_projects_list
- ✅ qxb_salary_details_list（工资明细，身份证脱敏、personId 完整）
- ✅ qxb_attendance_records_list（**考勤流水按 projectId+begin+end 查 8月 4521 条**，作为 kaoqin 主路径，替代 attend_table）
- ✅ qxb_attend_rates_list（**69 条按月出勤率**，作为 list_att_months 主路径）
- ✅ qxb_contract_persons_list（**136 人含 workType/teamName/hasContract/hasWorkInsurance**，作为 roster 兜底，替代 workers_list 故障）
- ✅ qxb_settlements_list（**11 条合同级结算**，作为 entry_exit 替代 exit_unsettled_workers_list 404）
- ⚠️ qxb_workers_list / qxb_subcontractors_list / qxb_project_get：服务端 zod 校验 bug（endDate null、state number），
  roster 仍优先尝试，待服务端修复后自动生效（2026-10-16 复核：subcontractors_list 仍挂）
- ⚠️ 2026-10-16 新增实测：salary_details_list 分页 hasMore 在中段误报 False（_items 已改为按 total 硬拉+id 去重）；
  salary_details/payrolls 的月份过滤均失效（全量拉回本地筛）；salary_detail_statistics 恒空、
  salary_files_list 需 all_projects 权限（均死路）；新增 salary_anchor_check 对账锚（明细 vs 批次，对平才可信）；
  teams_list 可用且带 subId（attend_table 备用来源）；projects_list 两条通道均正常，风控注记过时
- ❌ 04 线下考勤 / 05 专户进账流水（wage_accounts_list 需 all_projects 权限）/ 07 六项状态 / 08 投诉台账 / 09 迎检：
  平台无对应工具或权限不足 → 走附件扫描 / 待补充清单

**主键约定**：工资明细的 idCardNo 为脱敏串（前 6 后 4），不作为主键；以 personId（完整 ID）为
人员唯一键（idcard 字段承载 personId），报告不展示完整证件号，无影响。
"""

import json
import os
import tempfile
import time
import urllib.request
import urllib.error
from datetime import date, timedelta

from datasources import DataSource, load_xlsx, find_header_row, find_col, to_num

# ⚠️ qxb MCP 串行调用铁律：必须串行，禁止并发（防平台风控）。
# 模块级全局时间戳：单进程天然串行；多 _McpHttp 实例也共享同一节奏。
# 2026-10-16 拍板 3 秒；同日实测长会话（约 150 次连续调用）即便 3 秒也触发账号级风控
# （"账号异常，请稍后再试"，冷却约 30 分钟解限）；2026-10-19 老哥拍板：**默认间隔改为 5 秒**。
# 特殊场合可用环境变量 QXB_CALL_INTERVAL 覆盖。
_LAST_MCP_TS = 0.0
_QXB_CALL_INTERVAL = float(os.environ.get("QXB_CALL_INTERVAL", "5.0"))  # 秒

DEFAULT_QXB_CONFIG = {
    "mcp_json": os.path.join(os.path.expanduser("~"), ".workbuddy", "mcp.json"),
    "server_key": "qxb",      # 从 WorkBuddy mcp.json 复用连接配置
    "url": "",                # 直接指定则优先于 mcp_json
    "headers": {},            # 直接指定则优先于 mcp_json
    "project_id": "",         # 32 位项目 ID（可选）；留空则在运行时从 qxb MCP 会话解析（见 resolve_project）
    "project_name": "",       # 项目名（可选）；用于运行时从账号项目列表匹配，避免固化 ID
    "price": "",              # 工种单价表路径（qxb 平台无单价接口，委托本地表；缺省回退 skill references）
    "team_gong_map": None,    # 班组→工种映射（缺省用内置词典）
    "page_size": 100,  # 上游分页硬上限 100（请求更大也只回 100，2026-10-16 实测）
    "timeout": 90,
}


def _load_connection(cfg):
    """url/headers 解析：显式配置 > mcp.json 的 server_key 段。"""
    url = cfg.get("url") or ""
    headers = dict(cfg.get("headers") or {})
    if url and headers:
        return url, headers
    mj = cfg.get("mcp_json")
    key = cfg.get("server_key") or "qxb"
    if mj and os.path.exists(mj):
        try:
            with open(mj, encoding="utf-8") as f:
                srv = (json.load(f).get("mcpServers") or {}).get(key) or {}
            url = url or srv.get("url") or ""
            headers = headers or dict(srv.get("headers") or {})
        except Exception:
            pass
    if not url:
        raise ValueError("qxb 配置缺少 url（请在 judge_config 的 qxb 段或 mcp.json 中配置）")
    return url, headers


class _McpHttp:
    """极简 streamable HTTP MCP 客户端：initialize → tools/call（标准库实现）。"""

    def __init__(self, url, headers, timeout=60):
        self.url, self.headers, self.timeout = url, dict(headers), timeout
        self._sid = None
        self._id = 0
        self._initialize()

    def _post(self, payload):
        h = {"Content-Type": "application/json",
             "Accept": "application/json, text/event-stream", **self.headers}
        if self._sid:
            h["mcp-session-id"] = self._sid
        req = urllib.request.Request(self.url, data=json.dumps(payload).encode("utf-8"),
                                     headers=h, method="POST")
        with urllib.request.urlopen(req, timeout=self.timeout) as r:
            sid = r.headers.get("mcp-session-id")
            if sid:
                self._sid = sid
            raw = r.read().decode("utf-8")
        obj = None
        for line in raw.splitlines():
            if line.startswith("data:"):
                obj = json.loads(line[5:])
                break
        if obj is None:
            obj = json.loads(raw)
        return obj

    def _initialize(self):
        self._post({"jsonrpc": "2.0", "id": 0, "method": "initialize", "params": {
            "protocolVersion": "2025-03-26", "capabilities": {},
            "clientInfo": {"name": "month-report-qxb", "version": "1.0"}}})
        try:
            self._post({"jsonrpc": "2.0", "method": "notifications/initialized"})
        except Exception:
            pass

    def call(self, tool, arguments=None):
        """tools/call；返回 content[0].text 解析后的 JSON（失败抛 RuntimeError）。

        ⚠️ 串行铁律：每次调用前 sleep 到距离上一次至少 3 秒（防 qxb 平台风控）。
        首调用不等待（_LAST_MCP_TS=0 走短路径）；后续每次兜底 sleep。
        """
        global _LAST_MCP_TS
        now = time.time()
        if _LAST_MCP_TS > 0:
            gap = now - _LAST_MCP_TS
            if gap < _QXB_CALL_INTERVAL:
                time.sleep(_QXB_CALL_INTERVAL - gap)
        _LAST_MCP_TS = time.time()
        self._id += 1
        obj = self._post({"jsonrpc": "2.0", "id": self._id, "method": "tools/call",
                          "params": {"name": tool, "arguments": arguments or {}}})
        if obj.get("error"):
            raise RuntimeError(f"qxb MCP {tool}: {obj['error']}")
        res = obj.get("result", {})
        for c in res.get("content", []):
            if c.get("type") == "text":
                txt = c.get("text", "")
                try:
                    return json.loads(txt)
                except (ValueError, TypeError):
                    return txt
        return res.get("structuredContent")


def _norm_ym(s):
    """任意月份串 → YYYYMM（兼容 YYYY-MM / YYYYMM / YY.M / YYMM）。"""
    s = str(s or "").strip()
    if "." in s:
        parts = s.split(".")
        try:
            y, m = int(parts[0]), int(parts[1])
        except (ValueError, IndexError):
            return ""
        if y < 100:
            y += 2000
        return f"{y:04d}{m:02d}"
    digits = s.replace("-", "").replace("/", "")
    if len(digits) >= 6:
        return digits[:6]
    if len(digits) == 4:
        return "20" + digits
    return ""


class QxbMcpDataSource(DataSource):
    def __init__(self, data, cfg=None):
        cfg = {**DEFAULT_QXB_CONFIG, **(cfg or {})}
        self.data = data
        self._project_name = cfg.get("project_name") or ""
        url, headers = _load_connection(cfg)
        self.mcp = _McpHttp(url, headers, cfg.get("timeout", 60))
        self.page_size = int(cfg.get("page_size", 200))
        # 工种单价表：qxb 平台无单价接口，优先本地《工种单价预估范围.xlsx》，
        # 回退 skill 自带 references/工种单价预估范围.md（市场基准，非项目数据，可随 skill 分发）
        self.price_path = cfg.get("price") or os.path.join(data, "工种单价预估范围.xlsx")
        self._team_gong_map = cfg.get("team_gong_map")
        self._salary_cache = None
        self._kaoqin_cache = {}  # ym6 -> (by_nt, by_name, id_days)；一次运行内同月考勤只拉一遍（每次 MCP 调用间隔 3 秒）
        # 2026-10-21：身份一致性异常采集（供月报人工复核清单）
        self._mask_conflicts = {}       # 月份 -> {掩码证件号: [姓名...]}（工资册同掩码号多姓名）
        self._idc_name_conflicts = {}   # 月份 -> {证件号: [姓名...]}（考勤矩阵同证件号多姓名）
        self._sub_ids_cache = None  # 分包关系 ID 列表（teams_list）
        self._stats_attempts = 0  # salary_detail_statistics 连续不可用计数（防Broken期间反复空试）
        # 跨进程磁盘缓存（2026-10-19）：生成器与引擎子进程共享已拉数据，避免同量数据拉两遍
        # 触发平台配额限流。QXB_CACHE_FILE 指向缓存 JSON；QXB_SKIP_MATRIX=1 跳过矩阵试探。
        self._disk_cache = None
        # project_id：显式配置 > 环境变量 > 运行时从 qxb MCP 会话解析（默认）
        pid = cfg.get("project_id") or os.environ.get("QXB_PROJECT_ID") or ""
        if not pid:
            pid = self.resolve_project(self._project_name)
        self.project_id = pid

    # ----------------------------- 跨进程磁盘缓存 -----------------------------
    def _load_disk_cache(self):
        if getattr(self, "_disk_cache", None) is not None:
            return self._disk_cache
        path = os.environ.get("QXB_CACHE_FILE") or ""
        data = {}
        if path and os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                if not isinstance(data, dict):
                    data = {}
            except Exception:
                data = {}
        self._disk_cache = data
        return data

    def _save_disk_cache(self):
        path = os.environ.get("QXB_CACHE_FILE") or ""
        data = getattr(self, "_disk_cache", None)
        if not path or not data:
            return
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            # 2026-10-23：单缓存文件承载全部家当（工资/花名册/逐月考勤），写坏即全丢——
            # 临时文件写完原子替换，杜绝半截 JSON
            _tmp = path + ".tmp"
            with open(_tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
            os.replace(_tmp, path)
        except Exception:
            pass

    def _disk_get(self, key):
        return self._load_disk_cache().get(key)

    def _disk_put(self, key, value):
        self._load_disk_cache()[key] = value
        self._save_disk_cache()

    # ----------------------------- 内部工具 -----------------------------
    def _opt(self, tool, args=None):
        """可选工具调用：status=error / 结构异常 → None（降级，不中断）。"""
        try:
            r = self.mcp.call(tool, args)
        except (RuntimeError, urllib.error.URLError, OSError):
            return None
        if not isinstance(r, dict) or r.get("status") != "ok":
            return None
        return r.get("data") or {}

    def _items(self, tool, args, max_pages=300, label=""):
        """分页明细拉取（2026-10-16 重写，修 P0 截断缺陷）：
        ⚠️ 上游 hasMore 会在中段误报 False（实测 salary_details 第 20 页/2000 条处，实际 total=3758，
        21 页起仍有数据），且各月份乱序混排——**绝不能以 hasMore 判断取完**。
        规则：① 以 total 为准，硬拉到去重后条数 >= total；② 按记录 id 去重（翻页可能重叠）；
        ③ 上游对 pageSize 有硬上限 100（请求 200 也只回 100）；④ 连续 3 页零新增且未达 total 才放弃。
        label：页级进度打点前缀（长任务排障用）。"""
        import sys as _sys
        items, page, seen, dup_streak = [], 1, set(), 0
        empty_retries = 0  # 空页重试：上游偶发抖动/软限流会回 ok+空页，直接 break 会静默截断
        net_retries = 0  # 网络异常重试（2026-10-19 平台响应慢，60 秒超时频发，不兜住会崩整轮生成）
        total = None
        tag = label or tool
        while page <= max_pages:
            try:
                r = self.mcp.call(tool, {**args, "pageNum": page, "pageSize": self.page_size})
            except (TimeoutError, OSError, urllib.error.URLError, RuntimeError) as e:
                # RuntimeError：MCP 层 JSON-RPC 错误（限流期频发）；不兜住会崩整轮生成
                net_retries += 1
                _sys.stdout.write(f"[pull {tag} p{page}] 调用异常（{type(e).__name__}），重试 {net_retries}/2\n")
                _sys.stdout.flush()
                if net_retries >= 3:
                    _sys.stdout.write(f"[pull {tag}] 连续异常，放弃剩余页（已取 {len(items)} 条）\n")
                    _sys.stdout.flush()
                    break
                continue  # 同页重试
            net_retries = 0
            if not isinstance(r, dict) or r.get("status") != "ok":
                # 2026-10-23：页级 status!=ok 不再直接截断——26.7 生成实测 salary 第 31 页撞
                # "达到限制次数"，直接停会静默丢 758/3758 条明细（20%），历史日薪与对账全被污染。
                # 同页退避重试：限流/验证码类隔 2 分钟，其他 5 秒；5 次仍败才放弃剩余页
                _msg5 = " ".join(str(r.get(k) or "") for k in ("message", "msg", "error", "errMsg", "code")) \
                    if isinstance(r, dict) else ""
                net_retries += 1
                if net_retries >= 6:
                    _sys.stdout.write(f"[pull {tag}] 页级连续失败 5 次，放弃剩余页"
                                      f"（已取 {len(items)}/{total} 条）——数据不完整，需排查后重跑！\n")
                    _sys.stdout.flush()
                    break
                _backoff = 120 if (("限制" in _msg5) or ("稍后再试" in _msg5)
                                   or ("频繁" in _msg5) or ("验证码" in _msg5)) else 5
                _sys.stdout.write(f"[pull {tag} p{page}] status!=ok（{_msg5[:80] or '未知原因'}），"
                                  f"退避 {_backoff} 秒后重试同页 {net_retries}/5\n")
                _sys.stdout.flush()
                import time as _t5
                _t5.sleep(_backoff)
                continue
            data = r.get("data") or {}
            # 2026-10-19：平台周末升级后部分端点（如 contract_persons_list）返回键由 items 改为 list，二者兼容
            page_items = data.get("items") or data.get("list") or []
            if isinstance(data.get("total"), int):
                total = data["total"]
            new = 0
            for it in page_items:
                key = it.get("id") if isinstance(it, dict) else None
                if key is not None:
                    if key in seen:
                        continue
                    seen.add(key)
                items.append(it)
                new += 1
            _sys.stdout.write(f"[pull {tag} p{page}] +{new} uniq={len(items)}/{total} hasMore={data.get('hasMore')}\n")
            _sys.stdout.flush()
            if total and len(items) >= total:
                break
            if not page_items:
                empty_retries += 1
                if empty_retries >= 3:
                    break
                page += 1
                continue
            empty_retries = 0
            dup_streak = dup_streak + 1 if new == 0 else 0
            if dup_streak >= 3:
                _sys.stdout.write(f"[pull {tag}] 连续 3 页零新增且未达 total，停止（疑上游翻页异常）\n")
                _sys.stdout.flush()
                break
            page += 1
        return items

    def _salary_all(self):
        """工资明细全量（分页缓存，全项目所有月份）；优先读跨进程磁盘缓存。"""
        if self._salary_cache is None:
            cached = self._disk_get("salary")
            if cached:
                self._salary_cache = cached
            else:
                self._salary_cache = self._items("qxb_salary_details_list", {"projectId": self.project_id},
                                                 label="salary")
                self._disk_put("salary", self._salary_cache)
        return self._salary_cache

    def salary_anchor_check(self, ym):
        """对账锚（2026-10-16 新增）：目标月「工资册批次总额」vs「发放到人明细合计」。
        上游分页曾发生静默截断（hasMore 误报 False），明细合计与批次总额对平才算数据完整；
        对不平必须拒绝出判定结论（宁出【待复核】，不出错误判定）。
        返回 {ym, batch_count, batch_sum, detail_count, detail_sum, ok}。"""
        ym6 = self._norm(ym)
        batches = self._items("qxb_payrolls_list", {"projectId": self.project_id})
        b = [p for p in batches if str(p.get("payMonth", "")).replace("-", "") == ym6]
        d = [it for it in self._salary_all()
             if str(it.get("payMonth", "")).replace("-", "") == ym6]
        batch_sum = round(sum(float(p.get("finalPay") or 0) for p in b), 2)
        detail_sum = round(sum(float(it.get("finalPay") or 0) for it in d), 2)
        # 2026-10-16 堵洞：两边都是 0（0 人 0 笔）不是"对平"，是"数据没拿到"——
        # 只有批次与明细都有数据且金额一致才算 ok（限流期 0=0 假通过的教训）
        ok = len(b) > 0 and len(d) > 0 and abs(batch_sum - detail_sum) < 0.01
        return {"ym": ym6, "batch_count": len(b), "batch_sum": batch_sum,
                "detail_count": len(d), "detail_sum": detail_sum,
                "ok": ok}

    def _norm(self, ym):
        return _norm_ym(ym)

    # ----------------------------- 判定引擎需要 -----------------------------
    def _salary_via_statistics(self, ym6):
        """主路径（2026-10-16 老哥拍板）：qxb_salary_detail_statistics 按分包×月份到人统计
        （人名/证件号/出勤天数/金额），平台修复后每 月 8 次调用即可替代 38 页全量。
        任一分包失败，或全部分包 0 行（当前平台态：接口返回骨架但恒空）→ 返回 None 回退明细。
        连续 2 个月不可用则本运行内不再尝试（Broken 期间防 152 次空试）。"""
        if self._stats_attempts >= 2:
            return None
        subs = self._sub_ids()
        if not subs:
            return None
        ym_dash = f"{ym6[:4]}-{ym6[4:6]}"
        agg = {}
        any_data = False
        for sid in subs:
            try:
                r = self.mcp.call("qxb_salary_detail_statistics",
                                  {"projectId": self.project_id, "subId": sid,
                                   "payMonth": ym_dash})
            except Exception:
                self._stats_attempts += 1
                return None
            if not isinstance(r, dict) or r.get("status") != "ok":
                self._stats_attempts += 1
                return None
            d = r.get("data") or {}
            rows = d.get("list") or d.get("items") or d.get("rows") or []
            for row in rows:
                name = str(row.get("name") or row.get("personName") or "").strip()
                if not name:
                    continue
                pid = str(row.get("personId") or row.get("id") or row.get("idCardNo") or "").strip()
                pay = float(row.get("finalPay") or row.get("shouldPay") or row.get("amount") or 0)
                key = pid or ("__N__", name, "")
                rec = agg.get(key)
                if rec is None:
                    rec = {"idcard": pid or str(row.get("idCardNo", "")), "name": name,
                           "team": "", "pay": 0.0, "rows": 0, "short": ""}
                    agg[key] = rec
                rec["pay"] += pay
                rec["rows"] += 1
                any_data = True
        if not any_data:
            self._stats_attempts += 1
            return None
        return agg

    def salary_agg(self, ym):
        """当月到人工资聚合：明细全量为基准，statistics 仅当**人数不少于明细**时才采信。
        2026-10-19 实测教训：statistics 平台侧修复不完整（同月只回 7~10 人 vs 明细 37~306 人），
        若直接采信会静默少人——故以明细聚合为完整性闸门，statistics 人数 >= 明细人数才替换。"""
        ym6 = self._norm(ym)
        import sys as _sys
        agg = {}
        _mask_names = {}  # 2026-10-21：掩码证件号 -> set(姓名)，仅用于告警（不自动归并——
        # 实测存在掩码撞车：同区县+同后4的两个真人，5/105 对完全不同姓名，自动归并会相加两人工资）
        for it in self._salary_all():
            if str(it.get("payMonth", "")).strip() != ym6:
                continue
            pid = str(it.get("personId") or it.get("id") or "").strip()
            name = str(it.get("name", "")).strip()
            if not pid and not name:
                continue
            # 主键 = personId（idCardNo 平台侧脱敏，不作主键；报告不展示完整证件号）
            key = pid or ("__N__", name, "")
            _idc_raw = str(it.get("idCardNo", ""))
            _masked = _idc_raw if "*" in _idc_raw else ""
            rec = agg.get(key)
            if rec is None:
                rec = {"idcard": pid or _idc_raw,
                       # 2026-10-21：掩码证件号单独带出（供判定引擎做 证件+姓名双保险锚定）
                       "idcard_masked": _masked,
                       "name": name,
                       "team": "", "pay": 0.0, "rows": 0, "short": ""}
                agg[key] = rec
            if _masked:
                _mask_names.setdefault(_masked, set()).add(name)
            rec["pay"] += float(it.get("finalPay") or 0)
            rec["rows"] += 1
        for _mk, _ns in _mask_names.items():
            if len(_ns) > 1:
                # 2026-10-21：登记身份一致性异常（月报人工复核清单用）
                self._mask_conflicts.setdefault(ym6, {})[_mk] = sorted(_ns)
                _sys.stdout.write(f"[warn] 同一掩码证件号在工资册挂有多个姓名 {_mk}: "
                                  f"{'/'.join(sorted(_ns))} ——疑似录入笔误或掩码撞车，"
                                  "未自动归并，需人工核实\n")
                _sys.stdout.flush()
        stats = self._salary_via_statistics(ym6)
        if stats is not None and len(stats) >= len(agg):
            _sys.stdout.write(f"[salaryagg {ym6}] 口径=statistics，人数={len(stats)}（明细 {len(agg)}，采信 statistics）\n")
            _sys.stdout.flush()
            return stats
        # 2026-10-19：statistics 返回残缺数据（人数 < 明细）也计入熔断——半修状态的接口
        # 每次都"有响应"但不采信，不计数的话熔断永不触发，19 个月 × 9 次调用全白烧（实测 14 分钟）
        if stats is not None:
            self._stats_attempts += 1
        _sys.stdout.write(f"[salaryagg {ym6}] 口径=明细全量（statistics 人数 "
                          f"{len(stats) if stats is not None else '不可用'} < 明细 {len(agg)}，不采信），人数={len(agg)}\n")
        _sys.stdout.flush()
        return agg

    def _sub_ids(self):
        """分包关系 ID 列表：qxb_teams_list（2026-10-16 实测可用且带 subId；
        subcontractors_list 因 endDate=null zod 故障不可用）。"""
        if self._sub_ids_cache is None:
            subs = []
            try:
                r = self.mcp.call("qxb_teams_list", {"projectId": self.project_id, "pageSize": 200})
                data = (r or {}).get("data") or []
                if isinstance(data, dict):
                    data = data.get("items") or data.get("list") or []
                for it in data:
                    sid = str((it or {}).get("subId") or "").strip()
                    if sid and sid not in subs:
                        subs.append(sid)
            except Exception:
                subs = []
            self._sub_ids_cache = subs
        return self._sub_ids_cache

    def _kaoqin_via_matrix(self, ymd, by_nt, by_name, id_days):
        """主路径：qxb_attend_table 月度考勤矩阵（人×日）。
        2026-10-21（老哥拍板）：考勤以"人×日期"为唯一标准——平台存在同人在同班组重复两行
        （在职行+离场行，王付会 21+21=42 实证）、以及跨班组同日出勤，直接累加 attendDays 会
        重复计天；改用 attDays 每日打钩标记按日期并集去重：同一天无论打几次卡、在几个班组
        出勤，总数只计 1 天（分班组明细各取本班组去重天数，重合日会使分项之和大于总数）。
        按分包逐个查（subId 来自 teams_list）；任一分包失败或全部分包无打钩 → 返回 False
        （回退流水路径，其本身按 (personId, 日期) 去重）。"""
        import re as _re
        subs = self._sub_ids()
        if not subs:
            return False
        ym_dash = f"{ymd[:4]}-{ymd[4:6]}"
        nt_dates = {}   # (姓名, 班组) -> 出勤日号集合
        idc_dates = {}  # 证件号 -> 出勤日号集合（同证件的在职/离场重复行在此自然合并）
        idc_meta = {}   # 证件号 -> {"names": {姓名: 行权重}, "active": 在职行姓名集合,
        #                   "team_dates": [(班组, 日集合)]}（2026-10-21 改造，见下）
        for sid in subs:
            # 2026-10-22（老哥拍板）：考勤坚持用矩阵，不轻易回退流水——
            # ① 命中验证码：隔 2 分钟再试探一次，最多 10 次（约 20 分钟，防死循环）；
            # ② 非验证码瞬时抖动：保留原 5 秒 ×3 次短重试；
            # ③ 仍失败才返回 False 回退流水（保数据完整性）。
            import sys as _sys
            import time as _time
            r = None
            _CAPTCHA_ATTEMPTS = 5    # 限流/验证码退避总次数（5×2 分钟 = 10 分钟，2026-10-22 老哥拍板）
            _TRANSIENT_LIMIT = 5     # 两次退避之间的瞬时失败额度（超时/抖动）
            _transient = 0
            for _att in range(1, _CAPTCHA_ATTEMPTS + 1):
                try:
                    r = self.mcp.call("qxb_attend_table",
                                      {"projectId": self.project_id, "subId": sid,
                                       "queryType": "month", "yearMonth": ym_dash})
                except Exception as _e:  # MCP 层错误文本里也可能带验证码提示
                    r = {"status": "error", "message": str(_e)}
                if isinstance(r, dict) and r.get("status") == "ok":
                    break
                _msg = " ".join(str(r.get(_k) or "") for _k in
                                ("message", "msg", "error", "errMsg", "code")) \
                    if isinstance(r, dict) else str(r or "")
                _low = _msg.lower()
                # 2026-10-22 实测：矩阵失败真实原因是"当前接口达到限制次数，请稍后再试"
                # （平台限流），与验证码同属"需退避等待"——统一隔 2 分钟再试探；
                # 每次退避后瞬时失败计数清零（限流解除后上游常伴随读超时，需独立额度）
                if (("验证码" in _msg) or ("captcha" in _low)
                        or ("限制" in _msg) or ("稍后再试" in _msg)
                        or ("频繁" in _msg) or ("429" in _low)):
                    _sys.stdout.write(f"[kaoqin {ymd}] 分包 {sid} 触发限流/验证码"
                                      f"（{_msg[:80].strip() or '未知原因'}），"
                                      f"隔 2 分钟再试探（第 {_att}/{_CAPTCHA_ATTEMPTS} 次）...\n")
                    _sys.stdout.flush()
                    _time.sleep(120)
                    _transient = 0
                    continue
                _transient += 1
                if _transient < _TRANSIENT_LIMIT:
                    _sys.stdout.write(f"[kaoqin {ymd}] 分包 {sid} 矩阵调用失败"
                                      f"（{_msg[:120] or '未知原因'}），5 秒后重试 "
                                      f"{_transient}/{_TRANSIENT_LIMIT}...\n")
                    _sys.stdout.flush()
                    _time.sleep(5)  # 与 QXB_CALL_INTERVAL 默认一致
                else:
                    _sys.stdout.write(f"[kaoqin {ymd}] 分包 {sid} 矩阵连续失败"
                                      f"（{_msg[:120] or '未知原因'}），本月回退流水\n")
                    _sys.stdout.flush()
                    break
            if not isinstance(r, dict) or r.get("status") != "ok":
                return False
            tl = ((r.get("data") or {}).get("tableList")) or []
            for row in tl:
                raw_name = str(row.get("name") or "")
                name = _re.sub(r"\s*[（(]\s*离场\s*[)）]\s*$", "", raw_name).strip()
                team = str(row.get("teamName") or "").strip()
                if not name:
                    continue
                dates = set()
                for _dk, _dv in (row.get("attDays") or {}).items():
                    if _dv:  # 有打钩（如 ["√"]）才计出勤
                        try:
                            dates.add(int(str(_dk)))
                        except (ValueError, TypeError):
                            pass
                if not dates:
                    continue
                idc = str(row.get("idCardNo") or "").strip() or f"{name}|{team}"
                # 2026-10-21：0 天行也必须登记证件号→姓名（否则在职行 0 天+离场行有天数时，
                # canonical 姓名会错取离场行的笔误拼写，工资侧按姓名 join 不到考勤）
                meta = idc_meta.setdefault(idc, {"names": {}, "active": set(), "team_dates": []})
                meta["names"][name] = meta["names"].get(name, 0) + len(dates)
                if raw_name == name:
                    meta["active"].add(name)
                meta["team_dates"].append((team, dates))
                if not dates:
                    continue
                idc_dates.setdefault(idc, set()).update(dates)
        if not idc_dates:
            return False
        # 2026-10-21（老哥发现 李红海/李洪海 同一证件号）：平台录入笔误会让同一证件号
        # 挂多个姓名（全项目 8 月实测 6 对，如 宠/庞治强、罩/覃永翠）——同证件号强制
        # 归并为同一人，canonical 姓名取在职行中打钩权重最高者，避免同人被劈成
        # "有考勤无工资 + 有工资无考勤"两个假人
        canonical = {}
        for _idc, _meta in idc_meta.items():
            _names = _meta["names"]
            if len(_names) > 1:
                import sys as _sys
                # 2026-10-21：登记身份一致性异常（月报人工复核清单用）
                self._idc_name_conflicts.setdefault(ymd, {})[_idc] = sorted(_names)
                _sys.stdout.write(f"[warn] 同一证件号在考勤矩阵挂有多个姓名 {_idc}: "
                                  f"{'/'.join(sorted(_names))} ——已按在职行姓名归并为同一人\n")
                _sys.stdout.flush()
            _pref = [n for n in _names if n in _meta["active"]] or list(_names)
            canonical[_idc] = max(_pref, key=lambda n: _names[n])
        name_dates = {}  # 姓名(canonical) -> 出勤日号并集（跨班组/跨重复行/跨笔误姓名）
        for _idc, _ds in idc_dates.items():
            name_dates.setdefault(canonical[_idc], set()).update(_ds)
        for _idc, _pairs in idc_meta.items():
            _cn = canonical[_idc]
            for _tm, _ds in _pairs["team_dates"]:
                _k = (_cn, _tm)
                nt_dates[_k] = nt_dates.get(_k, set()) | _ds
        for (_nm, _tm), _ds in nt_dates.items():
            by_nt[(_nm, _tm)] = float(len(_ds))
        for _nm, _ds in name_dates.items():
            by_name[_nm] = float(len(_ds))
        # 2026-10-21：证件号→天数直查表（判定侧用工资掩码号夹逼匹配，绕开姓名笔误）
        for _idc, _ds in idc_dates.items():
            if len(_idc) == 18:
                id_days[_idc] = float(len(_ds))
        return True

    def _kaoqin_via_records(self, ymd, by_nt, by_name, id_days):
        """回退路径：qxb_attendance_records_list 打卡流水，按（姓名×日期）去重计天。
        2026-10-21：去重键从 (personId, 日期) 改为（姓名×日期）——平台流水里同一人
        在职/离场两个身份是不同 personId（王付会 21+21=42 实证），(pid,日期) 会把
        同一天计两次；与工资侧按姓名 join 同口径，且洗掉姓名"(离场)"后缀后再合。
        流水不带证件号（字段实测仅 8 个），id_days 保持为空。"""
        import re as _re
        try:
            y, m = int(ymd[:4]), int(ymd[4:6])
            first = date(y, m, 1)
            if m == 12:
                last = date(y + 1, 1, 1) - timedelta(days=1)
            else:
                last = date(y, m + 1, 1) - timedelta(days=1)
            begin = f"{first} 00:00:00"
            end = f"{last} 23:59:59"
        except (ValueError, TypeError):
            return
        import sys as _sys
        _sys.stdout.write(f"[kaoqin {ymd}] {begin} ~ {end} 拉取打卡流水（矩阵不可用，回退流水）...\n")
        _sys.stdout.flush()
        items = self._items("qxb_attendance_records_list",
                            {"projectId": self.project_id,
                             "begin": begin, "end": end},
                            label=f"attendance {ymd}")
        name_dates, nt_dates = set(), set()
        for it in items:
            pid = str(it.get("personId") or "").strip()
            name = _re.sub(r"\s*[（(]\s*离场\s*[)）]\s*$", "",
                           str(it.get("personName") or "")).strip()
            team = str(it.get("teamName") or "").strip()
            t = str(it.get("clockingTime") or "").strip()
            if not t:
                continue
            d = t[:10]
            if not d or len(d) < 10:
                continue
            if not name:
                continue  # 无姓名的流水无法与工资侧按姓名关联，若以 pid 充名会污染"有考勤无工资"名单
            name_dates.add((name, d))          # 总数口径：姓名×日期并集
            if team:
                nt_dates.add((name, team, d))  # 分班口径：姓名×班组×日期并集（两维度独立去重）
        _nm, _nt = {}, {}
        for (n, d) in name_dates:
            _nm.setdefault(n, set()).add(d)
        for (n, t, d) in nt_dates:
            _nt.setdefault((n, t), set()).add(d)
        for (n, t), ds in _nt.items():
            by_nt[(n, t)] = float(len(ds))
        _has_team = {n for (n, _t) in _nt}
        for n, ds in _nm.items():  # 全程无班组信息的人兜底挂空班组，避免从分班明细里消失
            if n not in _has_team:
                by_nt[(n, "")] = float(len(ds))
        for n, ds in _nm.items():
            by_name[n] = float(len(ds))

    def kaoqin(self, ym):
        """当月考勤（2026-10-16 老哥拍板：attend_table 矩阵为主路径）：
        ① 优先 attend_table 月度矩阵（按分包 subId 逐个查，每 月仅 8 次调用，量级远小于流水）；
        ② 矩阵不可用（上游返回空骨架/故障/验证码）时自动回退 attendance_records_list
           打卡流水（按 (姓名, 日期) 去重计天）。返回 (by_nt, by_name, id_days)：
        2026-10-21 新增第三元素 id_days={全号证件: 并集天数}——平台录入笔误（李红海/李洪海
        同一证件号两个姓名）会让姓名 join 断链，判定侧用工资掩码号夹逼全号直接取天数，
        姓名仅作显示；流水路径无证件号，id_days 为空。"""
        ymd = self._norm(ym)
        by_nt, by_name, id_days = {}, {}, {}
        if not ymd:
            return by_nt, by_name, id_days
        cached = self._kaoqin_cache.get(ymd)
        if cached is not None:
            return cached
        # 跨进程磁盘缓存：引擎子进程直接复用生成器已拉的考勤（不触发任何 MCP 调用）
        cache_all = self._load_disk_cache()
        disk = cache_all.setdefault("kaoqin", {})  # ⚠️ 必须挂到缓存对象上，or {} 会造成游离字典写丢
        if ymd in disk:
            entry = disk[ymd]
            # 2026-10-21 起考勤口径升级为"按日并集去重"(v2)；旧缓存为 attendDays 直接相加口径，作废重拉
            if entry.get("v") == 2:
                by_nt = {(n, t): d for n, t, d in entry.get("by_nt", [])}
                by_name = {n: d for n, d in entry.get("by_name", [])}
                id_days = {fid: d for fid, d in entry.get("id_days", [])}
                # 2026-10-21：同证件号多姓名冲突随缓存恢复（离线重建的月份也能进复核清单）
                for _idc, _ns in entry.get("id_conflicts", []):
                    self._idc_name_conflicts.setdefault(ymd, {})[_idc] = list(_ns)
                self._kaoqin_cache[ymd] = (by_nt, by_name, id_days)
                return by_nt, by_name, id_days
        import sys as _sys
        path_used = None
        if os.environ.get("QXB_SKIP_MATRIX") != "1":
            try:
                if self._kaoqin_via_matrix(ymd, by_nt, by_name, id_days):
                    path_used = "attend_table"
            except Exception:
                by_nt, by_name, id_days = {}, {}, {}
        if path_used is None:
            self._kaoqin_via_records(ymd, by_nt, by_name, id_days)
            path_used = "attendance_records"
        _sys.stdout.write(f"[kaoqin {ymd}] 口径={path_used}，人数={len(by_name)}\n")
        _sys.stdout.flush()
        self._kaoqin_cache[ymd] = (by_nt, by_name, id_days)
        disk[ymd] = {"v": 2,
                     "by_nt": [[n, t, d] for (n, t), d in by_nt.items()],
                     "by_name": [[n, d] for n, d in by_name.items()],
                     "id_days": [[fid, d] for fid, d in id_days.items()],
                     "id_conflicts": [[_idc, _ns] for _idc, _ns
                                      in self._idc_name_conflicts.get(ymd, {}).items()]}
        self._save_disk_cache()
        return by_nt, by_name, id_days

    def roster(self):
        """花名册：workers_list 优先（待服务端 zod 修好后自动生效）→ 失败回退
        contract_persons_list（136 人含 workType/teamName/hasContract/hasWorkInsurance/beginDate）→
        再回退用工资册人员合成（status/工种空）。"""
        rows = None
        r = self._opt("qxb_workers_list", {"projectId": self.project_id, "pageNum": 1,
                                           "pageSize": self.page_size})
        if r and (r.get("items") or r.get("list")):
            rows = r.get("items") or r.get("list")
        # 兜底 1：contract_persons_list（合同人员花名册，含完整工种/班组/合同/工保/进场日期）
        # 2026-10-16 改分页拉全（138 人 > 100/页上限，单页会截断）；
        # 2026-10-19 加重试：该端点偶发返回 ok+空（2026-10-19 晨实测），空结果隔一次调用间隔再试一轮
        if not rows:
            rows = self._disk_get("contract_persons") or []
        if not rows:
            for _attempt in (1, 2):
                rows = self._items("qxb_contract_persons_list", {"projectId": self.project_id},
                                   label=f"contract_persons 试{_attempt}")
                if rows:
                    break
            if rows:
                self._disk_put("contract_persons", rows)
        persons, by_id, by_nt = [], {}, {}
        if rows:
            for w in rows:
                rec = {"name": str(w.get("name", "")).strip(),
                       "id": str(w.get("personId") or w.get("idCardNo") or w.get("id") or "").strip(),
                       "status": str(w.get("state") or w.get("status") or "").strip(),
                       "entry": str(w.get("entryDate") or w.get("hireDate")
                                    or w.get("beginDate") or "").strip(),
                       "labor": str(w.get("laborCompanyName") or w.get("entName") or "").strip(),
                       "team": str(w.get("teamName") or "").strip(),
                       "gong": str(w.get("workType") or w.get("workKindName") or "").strip(),
                       "contract": str(w.get("hasContract") or w.get("contractSigned") or "").strip()}
                if rec["name"]:
                    persons.append(rec)
        if not persons:
            for it in self._salary_all():
                pid = str(it.get("personId") or it.get("id") or "").strip()
                name = str(it.get("name", "")).strip()
                if not pid or pid in by_id:
                    continue
                rec = {"name": name, "id": pid, "status": "", "entry": "", "labor": "",
                       "team": "", "gong": "", "contract": ""}
                persons.append(rec)
                by_id[pid] = [rec]
                by_nt.setdefault(name + "|" + rec["team"], []).append(rec)
            return persons, by_id, by_nt
        for rec in persons:
            key = rec["id"] or rec["name"] + "|" + rec["team"]
            by_id.setdefault(key, []).append(rec)
            by_nt.setdefault(rec["name"] + "|" + rec["team"], []).append(rec)
        return persons, by_id, by_nt

    def history(self, target_ym, window=None):
        """历史日薪：qxb 无专用接口 → 客户端按 prior 月份现算（与 db 源降级路径同语义）。"""
        t6 = self._norm(target_ym)
        months = sorted({m for m in (self.list_salary_months()) if m < t6})
        if window is not None:
            months = months[-window:] if window > 0 else []  # window=0 → 无历史（快速模式）
        hist = {}
        for m in months:
            sal = self.salary_agg(m)
            if not sal:
                continue
            by_nt, by_name = self.kaoqin(m)[:2]
            for key, rec in sal.items():
                days = by_nt.get((rec["name"], rec["team"]))
                if days is None:
                    days = by_name.get(rec["name"])
                if days and days > 0:
                    hist.setdefault(rec["idcard"], []).append(rec["pay"] / days)
        return hist, months

    def price_table(self):
        """工种单价区间：qxb 平台无此接口。优先本地《工种单价预估范围.xlsx》，
        回退 skill 自带 references/工种单价预估范围.md（市场基准，非项目数据，可随 skill 分发）。"""
        table = {}
        path = self.price_path
        if path and os.path.exists(path):
            try:
                rows = load_xlsx(path)
                hr = find_header_row(rows, ["工种", "底线", "上线"])
                hdr = rows[hr]
                ci = find_col(hdr, ["工种"], exact=True)
                cl = find_col(hdr, ["底线"])
                ch = find_col(hdr, ["上线"])
                for r in rows[hr + 1:]:
                    name = str(r[ci]).strip() if (ci is not None and r[ci]) else ""
                    if not name:
                        continue
                    lo, hi = to_num(r[cl]), to_num(r[ch])
                    if lo == 0 and hi == 0:
                        continue
                    table[name] = (lo, hi)
            except Exception:
                table = {}
        if not table:
            mpath = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "references", "工种单价预估范围.md")
            if os.path.exists(mpath):
                table = self._parse_price_md(mpath)
        return table

    @staticmethod
    def _parse_price_md(path):
        """解析 references/工种单价预估范围.md 的表格：取首列序号、第2列工种、末两列底线/上线，
        对说明列含 | 的情况稳健（首尾数字列定位）。"""
        table = {}
        with open(path, encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if not s.startswith("|"):
                    continue
                parts = [p.strip() for p in s.strip("|").split("|")]
                if len(parts) < 5:
                    continue
                try:
                    lo = float(parts[-2]); hi = float(parts[-1])
                except ValueError:
                    continue
                gong = parts[1]
                if not gong or (lo == 0 and hi == 0):
                    continue
                table[gong] = (lo, hi)
        return table

    # ----------------------------- 月报基础 -----------------------------
    def list_salary_months(self):
        return sorted({str(it.get("payMonth", "")).strip()
                       for it in self._salary_all() if it.get("payMonth")})

    def list_att_months(self):
        """真实考勤月份：取 qxb_attend_rates_list 的 time 字段（YYYY-MM-DD）去重 → YYYYMM。
        旧版本以工资月份近似，缺月时无法对齐；本路径不依赖分包 subId。"""
        months = set()
        try:
            items = self._items("qxb_attend_rates_list",
                                {"projectId": self.project_id})
        except Exception:
            items = []
        for it in items:
            t = str(it.get("time") or it.get("timeStr") or "").strip()
            if len(t) >= 7:
                ym = t[:4] + t[5:7]
                if ym.isdigit() and len(ym) == 6:
                    months.add(ym)
        if not months:
            return self.list_salary_months()
        return sorted(months)

    def attendance_target(self, target_ym):
        by_nt, _ = self.kaoqin(target_ym)[:2]
        per = {}
        for (name, team), d in by_nt.items():
            rec = per.setdefault(name, {"name": name, "team": team, "days": 0.0})
            rec["days"] += d
        return list(per.values())

    # ----------------------------- 项目解析（运行时，不落配置） -----------------------------
    def resolve_project(self, project_name=None):
        """从**当前登录用户的 qxb 会话**解析项目 ID，绝不读固化配置。

        2026-10-15：qxb 平台对 `projects_list` / `projects_page` 工具调用有风控（同一 session
        下 `user_info` 通但 `projects_list` 报"账号异常"），改为从 `qxb_user_info` 的
        `projectIds` 字段直接拿项目 ID 列表，绕过被风控的工具。

        - 单项目账号：自动选用 user_info.projectIds[0]；
        - 多项目：按 fullName 模糊匹配（仍走 projects_page 兜底，若可用）；
        - 多项目未指定：抛错并提示 --project-name。
        """
        # 路径 1：user_info.projectIds（绕过被风控的 projects_list/page）
        info = self._opt("qxb_user_info", {})
        if isinstance(info, dict):
            ids = info.get("projectIds") or []
            if isinstance(ids, list) and ids:
                if len(ids) == 1:
                    return str(ids[0])
                # 多项目：尝试 projects_page 拿全名做匹配
                raw = self._opt("qxb_projects_page", {"pageNum": 1, "pageSize": 100})
                items = (raw or {}).get("items") or []
                if not items:
                    raw2 = self._opt("qxb_projects_list", {})
                    items = (raw2 or {}).get("projects") or []
                if project_name:
                    name = str(project_name).strip()
                    for p in items:
                        fn = str(p.get("fullName") or p.get("projectName") or "")
                        if name == fn or name in fn:
                            return str(p.get("id"))
                    # projects 工具全风控时，匹配唯一含关键字的 id
                    return str(ids[0]) if len(ids) == 1 else None
                # 未指定：报错
                avail = "、".join(str(i) for i in ids)
                raise ValueError(f"账号下有多个项目，请用 --project-name 指定其一（ID 候选：{avail}）")
        # 路径 2：projects_page 兜底（极少数连接器对 user_info 风控而对 projects_page 通行）
        raw = self._opt("qxb_projects_page", {"pageNum": 1, "pageSize": 100})
        items = (raw or {}).get("items") or []
        if not items:
            raw2 = self._opt("qxb_projects_list", {})
            items = (raw2 or {}).get("projects") or []
        if not items:
            raise ValueError("qxb MCP 未返回任何项目：请确认当前 qxb 账号下存在项目，"
                             "或显式用 --project-id 指定 32 位项目 ID")
        if project_name:
            name = str(project_name).strip()
            for p in items:
                fn = str(p.get("fullName") or p.get("projectName") or p.get("name") or "")
                pn = str(p.get("projectName") or p.get("fullName") or "")
                if name == fn or name == pn or name in fn or name in pn:
                    return str(p.get("id"))
            avail = "、".join(str(p.get("fullName") or p.get("id")) for p in items)
            raise ValueError(f"未匹配到项目「{name}」；当前账号可用项目：{avail}")
        if len(items) == 1:
            return str(items[0].get("id"))
        avail = "、".join(str(p.get("fullName") or p.get("id")) for p in items)
        raise ValueError(f"账号下有多个项目，请用 --project-name 指定其一：{avail}")

    def project_meta(self):
        """项目主数据 + 五方主体（2026-10-21 起首选 qxb_project_get：平台修复 zod 故障并重写为
        「项目看板→项目基础信息→五方主体」一接口——一次返回项目档案、五方单位中文名与 11 类
        岗位人员姓名/证件/电话，封面三方与管理人员不再依赖附件 OCR）。
        project_get         失败时回退旧链：projects_page（富档案）→ projects_list 补名 → project_get 补缺。"""
        meta = {}

        def _p(v):
            return str(v).strip() if v not in (None, "") else ""

        # ---- 首选：qxb_project_get（档案 + 五方 + 人员，一接口全齐）----
        try:
            pg = self.mcp.call("qxb_project_get", {"projectId": self.project_id})
        except Exception:
            pg = None
        if isinstance(pg, dict) and pg.get("status") == "ok":
            d = pg.get("data") or {}
            pj = d.get("project") or {}
            if _p(pj.get("fullName")):
                meta["项目名称"] = _p(pj["fullName"])
            if _p(pj.get("projectType")):
                meta["项目类别"] = _p(pj["projectType"])
            if _p(pj.get("actualBeginDate")) or _p(pj.get("beginDate")):
                meta["实际开工时间"] = _p(pj.get("actualBeginDate")) or _p(pj.get("beginDate"))
            if _p(pj.get("endDate")):
                meta["预计完工时间"] = _p(pj["endDate"])
            if _p(pj.get("filingNo")):
                meta["备案编号"] = _p(pj["filingNo"])
            if _p(pj.get("contractPrice")):
                meta["合同价"] = _p(pj["contractPrice"])
            if _p(pj.get("builderLicence")):
                meta["施工许可证号"] = _p(pj["builderLicence"])
            parties = {str(pp.get("role") or "").strip(): pp
                       for pp in (d.get("fiveParties") or []) if isinstance(pp, dict)}
            for role in ("建设单位", "施工单位", "监理单位"):
                unit = _p((parties.get(role) or {}).get("unitName"))
                if unit:
                    meta[role] = unit
            people = {}
            for role in ("建设单位", "施工单位", "监理单位"):
                for pp in ((parties.get(role) or {}).get("people") or []):
                    if isinstance(pp, dict) and pp.get("key"):
                        people[str(pp["key"])] = pp

            def _pname(k):
                return _p(((people.get(k) or {}).get("person") or {}).get("name"))

            if _pname("partyASuperintendent"):
                meta["现场负责人"] = _pname("partyASuperintendent")
            if _pname("manager"):
                meta["项目经理"] = _pname("manager")
            if _pname("labour"):
                meta["劳资专管员"] = _pname("labour")
            if _pname("supervisionEngineer"):
                meta["总监理工程师"] = _pname("supervisionEngineer")
            detail = {}
            for k, pp in people.items():
                ps = pp.get("person") or {}
                if _p(ps.get("name")):
                    detail[k] = {"岗位": _p(pp.get("position")), "姓名": _p(ps.get("name")),
                                 "电话": _p(ps.get("phone"))}
            if detail:
                meta["_五方人员明细"] = detail
            if meta:
                meta["source"] = "project_get"
                return meta

        # ---- 兜底：旧链（projects_page 富档案 → projects_list 补名 → project_get 补缺）----
        r = self._opt("qxb_projects_page", {"pageNum": 1, "pageSize": 100})
        items = (r or {}).get("items") or []
        proj = None
        for p in items:
            if str(p.get("id")) == str(self.project_id):
                proj = p
                break
        if proj is None and items:
            proj = items[0]
        if proj:
            if proj.get("fullName"):
                meta["项目名称"] = str(proj["fullName"]).strip()
            if proj.get("projectType"):
                meta["项目类别"] = str(proj["projectType"]).strip()
            for ext, label in (("actualBeginDate", "实际开工时间"),
                               ("beginDate", "实际开工时间"),
                               ("endDate", "预计完工时间")):
                v = proj.get(ext)
                if v and label not in meta:
                    meta[label] = str(v).strip()
        if "项目名称" not in meta:
            r2 = self._opt("qxb_projects_list", {})
            for p in (r2 or {}).get("projects", []):
                if str(p.get("id")) == str(self.project_id):
                    meta["项目名称"] = str(p.get("fullName") or "").strip()
                    break
        pg = self._opt("qxb_project_get", {"projectId": self.project_id})
        if isinstance(pg, dict):
            d = pg.get("data") or {}
            _pj = d.get("project") or d  # 2026-10-21 起 data.project 为档案载体；兼容旧扁平结构
            for ext, label in (("fullName", "项目名称"), ("actualBeginDate", "实际开工时间"),
                               ("beginDate", "实际开工时间"), ("endDate", "预计完工时间")):
                v = _pj.get(ext)
                if v and label not in meta:
                    meta[label] = str(v)
        return meta

    # ----------------------------- 附件扫描兜底数据源 -----------------------------
    @staticmethod
    def _extract_urls(obj, out=None):
        """递归从任意 JSON 结构提取所有 http(s) URL（无视字段名）。"""
        if out is None:
            out = []
        if isinstance(obj, str):
            if obj.startswith("http://") or obj.startswith("https://"):
                out.append(obj)
        elif isinstance(obj, dict):
            for v in obj.values():
                QxbMcpDataSource._extract_urls(v, out)
        elif isinstance(obj, list):
            for v in obj:
                QxbMcpDataSource._extract_urls(v, out)
        return out

    @staticmethod
    def _safe_name(idx, url, tag):
        import re
        from urllib.parse import unquote
        seg = unquote(url.split("?")[0].rstrip("/").split("/")[-1])
        ext = ""
        m = re.search(r"\.([A-Za-z0-9]+)$", seg)
        if m:
            ext = "." + m.group(1).lower()
        tag = re.sub(r"[^\w\u4e00-\u9fff-]", "_", tag)[:20] or "file"
        return f"{idx:02d}_{tag}{ext}"

    @staticmethod
    def _tag_from_url(url):
        from urllib.parse import unquote
        return unquote(url.split("?")[0].rstrip("/").split("/")[-1])

    @staticmethod
    def _hint_for(tag):
        if "合同" in tag:
            return "参建三方名称、合同价"
        if "告示" in tag or "维权" in tag:
            return "建设/施工项目负责人、项目经理、劳资专管员、维权电话"
        if "工伤" in tag or "保险" in tag:
            return "参保单位、参保人数（六项制度-工保）"
        if "开户" in tag or "专户" in tag or "账户" in tag or "协议" in tag:
            return "专户账号、开户行、进账/代扣（资金05）"
        if "工保" in tag or "保证金" in tag:
            return "工资保证金存储状态（六项制度）"
        if "分账" in tag or "拨付" in tag:
            return "人工费分账拨付状态（六项制度）"
        if "开工" in tag or "施工许可" in tag:
            return "实际开工/完工日期"
        return "（待 agent 读图识别）"

    def attachment_scan(self, dest_dir=None):
        """枚举 + 下载项目全部公开附件到**系统临时目录**（不落项目/ skill 目录，跑完即弃），
        生成 manifest.json，返回清单。文字提取由 agent 读图完成（半自动），
        提取结果由调用方以 --attach-data 内联传入报告生成器，**不写回任何项目/ skill 文件**。"""
        import re
        from datetime import datetime
        if dest_dir is None:
            dest_dir = tempfile.mkdtemp(prefix="qxb_attach_")
        os.makedirs(dest_dir, exist_ok=True)
        sources = [
            "qxb_project_files_list",
            "qxb_project_contract_files_list",
            "qxb_settlement_files_list",
        ]
        seen, atts = set(), []
        idx = 0
        for tool in sources:
            try:
                raw = self.mcp.call(tool, {"projectId": self.project_id, "pageSize": 200})
            except Exception:
                continue
            for u in self._extract_urls(raw):
                if u in seen:
                    continue
                if not re.search(r"\.(png|jpe?g|webp|bmp|gif)($|\?)", u, re.I):
                    continue
                seen.add(u)
                idx += 1
                tag = self._tag_from_url(u)
                atts.append({"source": tool, "tag": tag, "url": u,
                             "local": self._safe_name(idx, u, tag),
                             "hint": self._hint_for(tag)})
        for a in atts:
            a["local_path"] = os.path.join(dest_dir, a["local"])
            try:
                from urllib.parse import quote
                dl = quote(a["url"], safe=":/")
                req = urllib.request.Request(dl, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=60) as resp:
                    blob = resp.read()
                with open(a["local_path"], "wb") as f:
                    f.write(blob)
                a["size"] = len(blob)
            except Exception as e:
                a["error"] = str(e)[:200]
        manifest = {"project_id": self.project_id,
                    "scanned_at": datetime.now().strftime("%Y-%m-%d"),
                    "dest_dir": dest_dir,
                    "attachments": atts}
        with open(os.path.join(dest_dir, "manifest.json"), "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        return manifest

    def project_name(self):
        return self.project_meta().get("项目名称")

    # ----------------------------- 月报明细（04-12） -----------------------------
    # _subcontractor_ids() 已删除（2026-10-14）：考勤改走 attendance_records_list 后不再依赖 subId；
    # subcontractors_list 当前服务端 zod 故障（endDate null）链路无法跑通，保留此方法不必要。

    def output_value(self):
        """06 产值：monthly_outputs_list（该账号当前无数据 → 空降级）。"""
        rows = self._items("qxb_monthly_outputs_list", {"projectId": self.project_id})
        out = []
        for r in rows:
            ym = _norm_ym(r.get("month") or r.get("yearMonth"))
            if not ym:
                continue
            out.append({"ym": ym,
                        "cur_reported": float(r.get("curOutput") or r.get("output") or 0),
                        "cum_reported": float(r.get("cumOutput") or 0),
                        "cum_approved": float(r.get("approved") or 0)})
        return out

    def entry_exit(self, target_ym):
        """10 进退场与结算：settlements_list（合同级结算 11 条，含 settlementDate/settlementPrice）。
        注意：settlements_list 是合同级数据（无工人名单），name/team 留空；报告侧可标"项目结算概况"语义。
        旧路径 exit_unsettled_workers_list 仍 404 不可用。"""
        out = []
        r = self._opt("qxb_settlements_list", {"projectId": self.project_id,
                                                "pageSize": 100})
        for s in ((r or {}).get("list") or (r or {}).get("items") or []):
            out.append({"name": "",
                        "team": "",
                        "entry": "",
                        "leave": str(s.get("settlementDate") or "").strip(),
                        "contract": "",
                        "settle_amount": float(s.get("settlementPrice") or 0),
                        "settle_status": "已结算"})
        return out

    def complaint_ledger(self):
        """08 投诉台账：qxb_service_records_list（项目级投诉/服务记录）。
        返回值语义（2026-10-16 老哥拍板）：**None = 接口故障/无工具（降级待填）**；
        **[] = 接口正常、平台确认 0 条投诉（报告应显示 0，不是待填）**。"""
        try:
            r = self.mcp.call("qxb_service_records_list", {"projectId": self.project_id,
                                                           "pageSize": 200})
        except Exception:
            return None
        if not isinstance(r, dict) or r.get("status") != "ok":
            return None
        rows = ((r.get("data") or {}).get("items")
                or (r.get("data") or {}).get("list") or [])
        out = []
        for it in rows:
            out.append({
                "date": str(it.get("date") or it.get("createTime") or "").strip()[:10],
                "port": str(it.get("port") or it.get("source") or "").strip(),
                "people": int(it.get("people") or it.get("workerCount") or 0),
                "amount": float(it.get("amount") or 0),
                "status": str(it.get("status") or it.get("handleState") or "").strip(),
                "handling": str(it.get("handling") or it.get("handleResult") or "").strip(),
            })
        return out

    def compliance(self):
        """项目合规快照（2026-10-19 平台新工具 qxb_project_compliance_get）：
        六项制度官方状态 + 专户余额 + 保证金详情 + 在册/合同/离场人数指标。
        该工具实测可用（2026-10-19），一举补齐此前"无接口"的保证金/分账拨付/离场结算与
        权限卡死的专户余额。失败返回 None（消费方走原有兜底链）。"""
        d = self._opt("qxb_project_compliance_get", {"projectId": self.project_id})
        if not d:
            return None
        det = (d.get("sources") or {}).get("indicatorDetail") or {}
        st = det.get("statuses") or {}

        def _one(key):
            v = st.get(key) or {}
            # 2026-10-20 老哥拍板：平台怎么写就怎么报——raw 保留上游原文（达标/限期整改等），
            # status 保留规范化值（met/warning/unmet）供旧消费方兼容
            return {"status": str(v.get("status") or "").strip(),
                    "raw": str(v.get("rawStatus") or "").strip(),
                    "note": str((v.get("reason") or {}).get("sanitized") or "").strip()}

        bal = det.get("accountBalance") or {}
        dep = det.get("deposit") or {}
        dep_amt = dep.get("depositAmount") or {}
        metrics = det.get("metrics") or {}
        try:
            bal_yuan = float(bal.get("amount") or 0)
        except (ValueError, TypeError):
            bal_yuan = None
        try:
            dep_wan = float(dep_amt.get("amount") or 0)
        except (ValueError, TypeError):
            dep_wan = None
        return {
            "indicators": {label: _one(key) for label, key in (
                ("保证金存储", "isStoreWageDeposit"), ("劳动合同签署", "isSignContract"),
                ("实名制打卡", "isRealNameAttendance"), ("人工费分账拨付", "isSplitAppropriation"),
                ("工资银行代发", "isAgentPayment"), ("离场结算", "isExitSettlement"))},
            "account_balance_yuan": bal_yuan,
            "balance_asof": str(bal.get("sourceUpdatedAt") or "").strip(),
            "deposit_type": str(dep.get("depositType") or "").strip(),
            "deposit_amount_wan": dep_wan,
            "deposit_validity": str(dep.get("depositValidity") or "").strip(),
            "metrics": {k: metrics.get(k) for k in ("onJobNum", "signContractNum",
                                                    "lastMonthPayNum", "exitWithSettlementNum",
                                                    "exitWithNoSettlementNum")
                        if metrics.get(k) is not None},
        }

    def monthly_summary(self, year):
        """月度工资考勤汇总（2026-10-19 平台新工具 qxb_project_salary_attendance_summary）：
        逐月 payAmount/attendPersonNum/payType（月度汇总口径，非逐人矩阵）。
        实测与明细对账精确一致（2026-07=707,319/105人、2026-08=650,803/88人），
        可作第三重对账锚。失败返回 []。"""
        d = self._opt("qxb_project_salary_attendance_summary",
                      {"projectId": self.project_id, "year": str(year)})
        if not d:
            return []
        out = []
        for it in (d.get("items") or []):
            try:
                out.append({"ym": f"{year}{str(it.get('whichMonth', '')).zfill(2)}",
                            "pay": float(it.get("payAmount") or 0),
                            "persons": int(float(it.get("attendPersonNum") or 0)),
                            "attend_times": int(float(it.get("attendTimes") or 0)),
                            "pay_type": str(it.get("payType") or "").strip()})
            except (ValueError, TypeError):
                continue
        return out

    def next_attendance(self, next_ym):
        """11 次月已发生考勤：attend_table 次月矩阵（依赖分包接口，当前降级为空）。"""
        ymd = self._norm(next_ym)
        ym_dash = f"{ymd[:4]}-{ymd[4:6]}" if ymd else ""
        by_nt, _ = self.kaoqin(ym_dash)[:2]
        return [{"name": n, "team": t, "days": d} for (n, t), d in by_nt.items()]

    def yearly_salary(self):
        """12 全年逐月工资册：salary_details 全量 + 按月 join 考勤算 days。
        2026-10-16 修复：此前 days 写死 0.0，消费侧「yearly_rows 覆盖现算」分支判 days 恒假，
        导致 3.1.4 全年日工资分布在 qxb 源下恒为空图。明细无班组 → days 按姓名匹配考勤
        （kaoqin 的 by_name 已跨班组合并）；无考勤月份 days=0，消费侧自然跳过。"""
        out = []
        for it in self._salary_all():
            name = str(it.get("name", "")).strip()
            if not name:
                continue
            out.append({"ym": str(it.get("payMonth", "")), "name": name, "team": "",
                        "gong": "", "days": 0.0, "pay": float(it.get("finalPay") or 0)})
        by_month = {}
        for rec in out:
            by_month.setdefault(rec["ym"], []).append(rec)
        for ym, recs in by_month.items():
            # 2026-10-19 配额保护：只为**已缓存**的考勤月份做 join（快速模式下仅目标月有缓存），
            # 未缓存的月份不触发新的 MCP 拉取——否则 19 个月 × 9 次失败调用空烧配额
            if ym not in self._kaoqin_cache:
                continue
            try:
                _by_nt, by_name = self.kaoqin(ym)[:2]
            except Exception:
                by_name = {}
            if not by_name:
                continue
            for rec in recs:
                d = by_name.get(rec["name"])
                if d and d > 0:
                    rec["days"] = float(d)
        return out
