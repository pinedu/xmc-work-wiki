# 数据采集 → 风险评估 流水线图

```
┌─────────────────────────────────────┐
│ rhzy-contractor-due-diligence (采集层) │
│                                      │
│ 输入：公司名 + 信用代码              │
│ 输出：merged_<日期>.json             │
└──────────────┬───────────────────────┘
               │
               ▼  ~/hermes/due-diligence/<公司>/<日期>/
┌──────────────────────────────────────┐
│ S1 gsxt  │ S2 zxgk │ S3 wenshu       │
│ S4 jzsc  │ S5 贵州  │ S6 tax          │
│ (6 个分源 JSON + 1 个 merged)        │
└──────────────┬───────────────────────┘
               │
               ▼
┌──────────────────────────────────────┐
│ rhzy-skill1-contractor-risk-         │
│ assessment (分析层)                  │
│                                      │
│ 输入：merged JSON                    │
│ 输出：红/橙/黄/蓝 + 清单 + 建议       │
└──────────────────────────────────────┘
```

## 关键设计点

1. **采集层与分析层解耦**：可以独立测试、独立维护
2. **每源独立失败**：部分源抓不到不影响其他
3. **离线模式**：网络不通时生成 placeholder + TODO，让用户手动补
4. **JSON 是事实层**：不做 LLM 语义清洗，留给 skill1

## 真实运行产物（2026-07-09 贵州建工一公司测试）

```
~/hermes/due-diligence/贵州建工集团第一建筑工程有限责任公司/2026-07-09/
├── merged_2026-07-09.json   (4.2KB - 给 skill1 消费)
├── S1_gsxt.json             (skipped - 待 cookie)
├── S2_zxgk.json             (needs_patch)
├── S3_wenshu.json           (needs_patch)
├── S4_jzsc.json             (needs_patch)
├── S5_guizhou.json          (needs_patch)
└── S6_tax.json              (needs_patch)
```

测试命令：`python collect.py --company "贵州建工集团第一建筑工程有限责任公司" --code "91520000214401234X" --offline`

输出符合 schema，6 源 + 1 merged 全部落盘。