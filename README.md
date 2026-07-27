# Excel Reconciliation Tool

> A professional Excel reconciliation tool for financial and accounting data.

Version: V1.0 (Planning)

---

# Project Overview

Excel Reconciliation Tool 是一个用于财务、会计及数据分析场景的自动对账工具。

程序能够自动比较两个 Excel Sheet 中的金额数据，并根据预设规则寻找对应关系，生成带颜色标记的对账结果。

项目目标是：

> **保持原始数据不变，实现智能金额匹配。**

---

# Features

支持以下匹配方式：

✅ 1 ↔ 1（Exact Match）

✅ 1 ↔ 1（Tolerance Match，±0.01）

✅ 1 ↔ 2

✅ 2 ↔ 1

✅ 1 ↔ 3

✅ 3 ↔ 1

✅ 2 ↔ 2

后续版本将继续支持：

- 2 ↔ 3
- 3 ↔ 2
- 3 ↔ 3
- 更复杂组合

---

# Input Format

Excel 文件包含两个 Sheet。

例如：

## Sheet1

| Row | Content |
|------|----------|
|1|Total|
|2|Header|
|3~End|Amount|

## Sheet2

结构完全一致。

程序默认：

- 第1行为总计（不参与匹配）
- 第2行为标题（不参与匹配）
- 第3行开始参与计算

---

# Matching Rules

按照固定顺序执行：

## Stage 1

### Exact Match

金额完全一致。

颜色：

🟩 Green

---

## Stage 2

### Tolerance Match

允许误差：

±0.01

颜色：

🟦 Blue

---

## Stage 3

### One to Two

1 ↔ 2

颜色：

🟨 Yellow

---

## Stage 4

### Two to One

2 ↔ 1

颜色：

🟨 Yellow

---

## Stage 5

### One to Three

1 ↔ 3

颜色：

🟧 Orange

---

## Stage 6

### Three to One

3 ↔ 1

颜色：

🟧 Orange

---

## Stage 7

### Two to Two

2 ↔ 2

颜色：

🟪 Purple

---

## Unmatched

所有未匹配数据：

🟥 Red

---

# Design Principles

## Preserve Original Data

程序绝不：

- 删除数据
- 排序数据
- 移动数据

Excel 原始顺序保持不变。

---

## Output

程序生成：

```
result.xlsx
```

原始 Excel 不做任何修改。

---

# Output Summary

程序结束后输出统计：

```
========================================

Sheet1 Records

Sheet2 Records

Exact Match

Tolerance Match

1 -> 2

2 -> 1

1 -> 3

3 -> 1

2 -> 2

Unmatched Sheet1

Unmatched Sheet2

========================================
```

---

# Technology Stack

Programming Language

- Python 3.12+

Libraries

- openpyxl
- decimal
- itertools

---

# Planned Project Structure

```
Excel-Reconciliation/

│
├── README.md
├── CHANGELOG.md
├── ROADMAP.md
├── requirements.txt
│
├── main.py
├── matcher.py
├── excel_io.py
├── models.py
│
├── data.xlsx
└── result.xlsx
```

---

# Version Roadmap

## Version 1.0

- Excel Read
- Exact Match
- Tolerance Match
- 1↔2
- 2↔1
- 1↔3
- 3↔1
- 2↔2
- Color Highlight
- Result.xlsx
- Statistics

---

## Version 2.0

Performance Optimization

- Faster matching algorithm
- Large Excel support
- Better memory usage

---

## Version 3.0

Desktop Application

- GUI
- Drag & Drop
- EXE Package
- PDF Report

---

# Development Principles

整个项目遵循以下原则：

1. 正确性优先
2. 原始数据不可修改
3. 代码保持可维护
4. 功能逐步迭代
5. 先完成 V1，再优化性能

---

# License

MIT License

---

# Author

Steven Sun

Project Start:

2026

Current Status:

Planning / Design Phase

# Core Philosophy

本项目不是传统意义上的 Excel 比较工具。

它的核心目标是：

> **寻找金额之间的业务对应关系（Reconciliation），而不是简单比较单元格。**

因此：

- 一笔金额可以对应多笔金额
- 多笔金额也可以对应一笔金额
- 匹配遵循固定规则与优先级
- 所有匹配过程均保留原始数据

项目强调的是**可解释、可追踪、可维护**的自动对账流程，而不是单纯的数据差异比较。