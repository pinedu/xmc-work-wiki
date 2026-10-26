# -*- coding: utf-8 -*-
"""
月报生成器侧的数据源桥接（不重复实现，统一复用 wage-pay-judge 的 datasources）。
================================================================================
判定引擎与月报生成器必须共用同一份数据读取逻辑，否则会出现“判定一套、报告一套”的分叉。
本文件仅做路径桥接，真正的数据读取实现位于：
    <skills>/wage-pay-judge/scripts/datasources.py

注意：本文件与真实实现同名（datasources.py），若直接 `from datasources import ...`
会解析到自身造成循环导入。因此用 importlib 按文件绝对路径显式加载真实模块，
并以别名 `wpj_datasources` 注册，再 re-export 给 build_full_report.py 使用。
"""
import os
import sys
import importlib.util

_WAGE_SCRIPTS = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                 "wage-pay-judge", "scripts")
)
_WAGE_DS = os.path.join(_WAGE_SCRIPTS, "datasources.py")

_spec = importlib.util.spec_from_file_location("wpj_datasources", _WAGE_DS)
wpj_ds = importlib.util.module_from_spec(_spec)
sys.modules["wpj_datasources"] = wpj_ds
_spec.loader.exec_module(wpj_ds)

# re-export（真实实现在 wage-pay-judge/scripts/datasources.py；
# 外部数据源仅 qxb，QxbMcpDataSource 实现在 datasource_qxb.py，经 make_datasource 工厂加载。
# 注意：datasource_qxb.py 以 `from datasources import ...` 取 helper，本进程内该名字指向本桥接模块，
# 因此 helper 也必须齐数 re-export，否则 qxb 源在生成器进程内 ImportError）
DataSource = wpj_ds.DataSource
make_datasource = wpj_ds.make_datasource
team_to_gong = wpj_ds.team_to_gong
DEFAULT_TEAM_TO_GONG = wpj_ds.DEFAULT_TEAM_TO_GONG
load_xlsx = wpj_ds.load_xlsx
find_header_row = wpj_ds.find_header_row
find_col = wpj_ds.find_col
to_num = wpj_ds.to_num
