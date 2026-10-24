# gszy-thesis-format

甘肃中医药大学 **医学信息工程学院** 本科毕业论文（设计）排版与格式审核工具。

把一份 Markdown 稿件排成符合学院《撰写规范》的 Word 文档；也能反过来体检别人交来的 `.docx` 有多少处不合规。

作者不是学院官方，规范条文全部来自学院下发的《医学信息工程学院本科毕业论文（设计）撰写规范建议及示例》，原件保存在 `assets/reference-template.docx`，有争议时以学院最新文件为准。

## 它解决什么

手工套论文模板最费时间的几件事，都交给脚本：

- 封面、题目内页、授权书、中英文摘要、目录、正文、总结、致谢、参考文献、附录 —— 一次生成，顺序和装订要求一致
- 前置部分用 Ⅰ Ⅱ Ⅲ 页码，正文从第 1 章重新从 1 编号（分节实现，不是手动插分页符）
- 页眉「甘肃中医药大学本科毕业论文（设计）」宋体五号带双线，页脚页码宋体小五居中
- 三级标题样式（黑体小三居中 / 黑体小四左对齐）、正文宋体小四首行缩进 2 字符 1.5 倍行距
- 三线表（上下 1.5 磅、表头下 0.5 磅、无竖线）、图题图下、表题表上、公式编号右对齐
- 目录是真目录域，按 Ctrl+A 再 F9 就能更新出带页码的目录

## 快速开始

```bash
pip install python-docx

git clone https://github.com/xiao-jie520/gszy-thesis-format.git
cd gszy-thesis-format

# 用示例稿件跑一遍
python scripts/build_thesis.py assets/sample-manuscript.md -o demo.docx

# 只要一个空白但格式正确的骨架
python scripts/build_thesis.py --skeleton 空白论文.docx

# 检查一份已有论文
python scripts/check_thesis.py 我的论文.docx -o 检查报告.md
```

生成后打开 Word，按 `Ctrl+A` → `F9` 更新域（目录由域生成，不更新只显示占位文字）。

## 稿件怎么写

````markdown
---
title: 基于Spring Boot的|药房管理系统的设计与实现   # | 指定标题换行位置
title_en: Design and Implementation of a Pharmacy|Management System Based on Spring Boot
advisor: 张三  教授
major: 医学信息工程
grade: 2022级
author: 李四
date: 二〇二六年六月
---

# 摘　　要

摘要正文……

【关键词】药房管理；Spring Boot；B/S架构；MySQL；信息系统

# Abstract

English abstract……

【Keywords】Pharmacy Management  Spring Boot  B/S Architecture  MySQL  Information System

# 第1章 引言

## 1.1 研究背景与意义

正文段落，自动首行缩进 2 字符。

系统架构如图2-1所示。

![图2-1 系统总体架构](fig/arch.png)

各模块职责如表3-1所示。

表3-1 系统功能模块划分

| 模块名称 | 主要功能 | 使用角色 |
| --- | --- | --- |
| 药品管理 | 药品信息增删改查 | 管理员 |

预警条件如式（3-1）所示。

```formula (3-1)
Q_current < Q_safe
```

# 参考文献

[1] 吴功宜.计算机网络(第三版)[M].北京:清华大学出版社,2011:5-26.
````

完整语法说明见 [`references/manuscript-syntax.md`](references/manuscript-syntax.md)，可运行的示例见 [`assets/sample-manuscript.md`](assets/sample-manuscript.md)。

## 规范速查

| 项目 | 要求 |
| --- | --- |
| 纸张 | A4，上 3.5 / 下 2.5 / 左右 2.5 cm，左侧装订 |
| 页眉 | 「甘肃中医药大学本科毕业论文（设计）」，宋体五号居中，距边界 2.5 cm |
| 页脚 | 页码宋体小五居中，距边界 2.0 cm |
| 行距 | 全文 1.5 倍，段前段后 0 |
| 封面标题 | 黑体二号加粗，不超过 30 字 |
| 一级标题 | 黑体小三，居中，段前段后各 1 行，每章另起一页 |
| 二、三级标题 | 黑体小四，左对齐，段前段后各 6 磅 |
| 正文 | 宋体小四，首行缩进 2 字符；英文与数字 Times New Roman |
| 图题 / 表题 | 宋体五号居中；图题在图下，表题在表上 |
| 参考文献 | 宋体五号，按正文引用顺序编号 |
| 章节编号 | 第1章、1.1、1.1.1，最多三级；图表公式分章编号：图3-1、表2-2、式(3-1) |

逐条细则、著录格式与最常被判不合格的十六条，见 [`references/format-spec.md`](references/format-spec.md)。

## 审核报告长什么样

```markdown
**结果：3 项错误，5 项警告**

## 必须修改
- ❌ 上边距 2.20 cm，应为 3.5 cm
- ❌ 一级标题字号 [14.0]，应为 15 pt（小三）：第2章 关键技术
- ❌ 中文关键词 2 个，规范为 3～5 个

## 建议核对
- ⚠️ 行距不是 1.5 倍：共 6/26 段（如第 48 段「目　　录」→1.0）
- ⚠️ 表题下方未紧跟表格：表3-1 系统功能模块划分
```

覆盖页面尺寸与页边距、页码分节、页眉、正文字体字号行距缩进、标题层级、图表题位置与编号、参考文献著录与字号、关键词个数与分隔符、中文段落混用半角标点、空段过多等检查项。

## 目录结构

```
.
├── SKILL.md                          # 技能说明（WorkBuddy / Claude Skill 入口）
├── scripts/
│   ├── build_thesis.py               # 稿件 → 规范 .docx
│   └── check_thesis.py               # .docx → 检查报告
├── references/
│   ├── format-spec.md                # 完整规范条文与常见错误
│   └── manuscript-syntax.md          # 稿件语法
└── assets/
    ├── sample-manuscript.md          # 示例稿件
    ├── fig_sample.png                # 示例插图
    └── reference-template.docx       # 学院原始规范文件
```

顶层 `SKILL.md` 用 WorkBuddy / Claude 的 Skill 机制组织，因此整个仓库也可直接放进 `~/.workbuddy/skills/` 或 `~/.claude/skills/` 作为技能使用；不用 Skill 机制时，直接调 `scripts/` 里的两个脚本即可。

## 已知限制

- **公式**：只做占位排版（1×2 无框表格，公式居中、编号右对齐）。复杂公式请在 Word 里用 MathType 或公式编辑器重排，保留编号位置即可。脚本不生成 OMML。
- **版面微调**：脚本判断不了"图跨页""表格挤在半页"，生成后需要人工扫一眼图表所在页。
- **目录**：由 Word 域生成，必须在 Word 里更新一次；WPS 打开同样支持 F9 更新。
- **只排版不写作**：不会改动原文用词与标点风格，规范性问题只在报告里指出。

## 技术备注

两个踩过的坑，供改代码的人参考：

1. Word 会把标题 run 的字符格式抄进目录条目。所以标题只用段落样式、不写字符级格式，并且样式的 `rFonts` 必须删掉 `asciiTheme` / `eastAsiaTheme` 等主题属性 —— 否则标题和目录会渲染成等线、微软雅黑。
2. 「目录」这个词自己不要出现在目录里：目录标题用内置 `TOC Heading` 样式（大纲级别=正文），而不是 `Heading 1` 再改大纲级别。

## 许可

脚本与文档可自由使用、修改。规范内容版权归甘肃中医药大学医学信息工程学院。
