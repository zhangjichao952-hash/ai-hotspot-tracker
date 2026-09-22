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

## 开发者信息

- 开发者：张继超
- GitHub：https://github.com/zhangjichao952-hash
- 版本：1.0.0
- 开源协议：MIT License（详见 LICENSE 文件）

## 安装

```bash
git clone git@github.com:zhangjichao952-hash/ai-hotspot-tracker.git ~/.stepfun/skills/ai-hotspot-tracker
```

## 卸载方法

```bash
rm -rf ~/.stepfun/skills/ai-hotspot-tracker
```

卸载后不会在系统中残留任何数据。

## 安装及运行所需权限

| 权限 | 用途 |
|---|---|
| 网络访问（WebSearch） | 检索 25 个 AI 信息源的公开互联网内容 |
| 文件写入 | 在用户指定路径下创建 Excel 报告文件 |
| Python 执行 | 运行 generate_report.py 脚本生成 Excel |

## 依赖

```bash
pip3 install openpyxl
```

openpyxl 遵循 MIT License，仅用于本地 Excel 文件生成，不涉及网络请求或数据收集。

## 使用

在 AMOO 或兼容的 AI 助手中，说出类似以下指令即可自动触发：

- "帮我追踪 AI 热点并规划选题"
- "搜集 AI 行业最新信息并整理短视频选题"
- "从 AI 信息源搜集热点并评估"

## 结构

```
ai-hotspot-tracker/
├── SKILL.md                          # 主手册：五步工作流程
├── PRIVACY_POLICY.md                 # 用户协议与隐私政策
├── LICENSE                           # MIT 开源协议
├── references/
│   └── ai-info-sources.md            # 25 个 AI 信息源网站清单及分类
└── scripts/
    └── generate_report.py            # Excel 报告生成脚本
```

## 隐私政策

详见 [PRIVACY_POLICY.md](PRIVACY_POLICY.md)。本技能不收集、不存储、不上传用户任何个人信息。
