# -*- coding: utf-8 -*-
"""
scan_attachments.py —— qxb 平台附件扫描数据源（枚举 + 下载到临时目录）

用途：月报生成前，先把 qxb 平台项目公开附件（合同扫描件、维权告示牌、工伤
保险、结算文件等）枚举并下载到**系统临时目录**（不落项目/ skill 目录，不外传第三方），
生成 manifest.json 并返回清单。文字提取由 agent 读图完成（半自动）——agent 读取
临时目录里的图片、提取单位/人名/日期/金额，再以 --attach-data 内联传给
build_full_report.py，**不写回任何项目/ skill 文件**，保证 qxb 平台数据不落本地。

用法：
    python scan_attachments.py --judge-config <判定口径.json> \
        [--project-name <项目名> | --project-id <id>] [--out <dest_dir>]
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from datasource_qxb import QxbMcpDataSource


def main():
    ap = argparse.ArgumentParser(description="qxb 平台附件扫描（枚举+下载到临时目录）")
    ap.add_argument("--judge-config", default=None, help="判定口径 json（仅取 qxb 连接段；不含项目数据）")
    ap.add_argument("--project-name", default="", help="qxb 项目名（运行时从账号项目列表匹配）")
    ap.add_argument("--project-id", default="", help="覆盖 project_id（32 位）")
    ap.add_argument("--out", default="", help="下载目录（默认系统临时目录 qxb_attach_xxx）")
    args = ap.parse_args()

    cfg = {}
    if args.judge_config and os.path.exists(args.judge_config):
        with open(args.judge_config, encoding="utf-8") as f:
            jc = json.load(f)
        cfg = dict(jc.get("qxb") or {})
    if args.project_id:
        cfg["project_id"] = args.project_id
    if args.project_name:
        cfg["project_name"] = args.project_name

    # data 仅用于定位本地工种单价表，可指向任意目录（默认当前目录）
    data_dir = os.path.dirname(os.path.abspath(args.judge_config)) if args.judge_config else "."
    ds = QxbMcpDataSource(data_dir, cfg)
    manifest = ds.attachment_scan(args.out or None)

    n = len(manifest["attachments"])
    ok = sum(1 for a in manifest["attachments"] if a.get("size") is not None and "error" not in a)
    print(f"附件扫描完成：共发现 {n} 个图片附件，成功下载 {ok} 个（临时目录：{manifest['dest_dir']}）")
    print(f"manifest: {os.path.join(manifest['dest_dir'], 'manifest.json')}")
    print("请 agent 用 Read 读取下列临时图片并提取单位/人名/日期/金额，再以 --attach-data 内联传入生成器：")
    for a in manifest["attachments"]:
        sz = a.get("size", "ERR")
        print(f"  [{a['source']}] {a['tag']}  -> {a['local_path']}  ({sz}B)")
        print(f"       建议提取: {a['hint']}")


if __name__ == "__main__":
    main()
