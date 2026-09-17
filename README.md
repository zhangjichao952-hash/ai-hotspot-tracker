# AI Hotspot Tracker

AI 行业热点信息搜集与短视频选题规划技能（Skill）。

## 功能

从 25 个 AI 信息源网站搜集热点新闻 → 去重 → 六维度评估 → 生成短视频选题方案，输出为 Excel 文件。

六维度评估体系：
- 发布时间
- 讨论度（极高/高/中/低）
- 普通用户可理解性（高/中/低）
- 画面可演示性（极高/高/中/低）
- 同质化程度（极高/高/中/低）
- 普通用户理解门槛（低/中/高）

## 安装

```bash
git clone git@github.com:zhangjichao952-hash/ai-hotspot-tracker.git ~/.stepfun/skills/ai-hotspot-tracker
```

## 使用

在 AMOO 或兼容的 AI 助手中，说出类似以下指令即可自动触发：

- "帮我追踪 AI 热点并规划选题"
- "搜集 AI 行业最新信息并整理短视频选题"
- "从 AI 信息源搜集热点并评估"

## 结构

```
ai-hotspot-tracker/
├── SKILL.md                          # 主手册：五步工作流程
├── references/
│   └── ai-info-sources.md            # 25 个 AI 信息源网站清单及分类
└── scripts/
    └── generate_report.py            # Excel 报告生成脚本
```

## 依赖

```bash
pip3 install openpyxl
```
