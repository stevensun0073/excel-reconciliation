# Excel Reconciliation V1.0

Automatic Bank Reconciliation Tool

Author: Steven Sun

---

# 1. Project Overview

本系统用于自动完成银行流水（Sheet1）与银行对账单（Sheet2）的金额匹配。

设计目标：

- 尽可能自动完成匹配
- 不强行凑组合
- 保留业务逻辑
- 不确定的记录留给人工审核

整个系统坚持：

> 宁可少匹配，也不要错误匹配。

---

# 2. Workflow

整个流程共六个阶段：

```
company.xlsx
        │
        ▼
prepare_data.py
        │
        ▼
data.xlsx（生成 Key word）
        │
        ▼
main.py
        │
        ├── Stage 1  Exact Matching
        ├── Stage 2  Keyword Difference
        ├── Stage 3  Final Yellow Matching
        ├── Stage 4  Validation
        ├── Stage 5  Difference Analysis
        ▼
result_reconciliation.xlsx
```

---

# 3. Program Files

## prepare_data.py

负责数据预处理。

功能：

### Sheet1

读取：

- company.xlsx
- 文本列

生成：

Key word

规则：

- Company Keyword
- KZ
- GR

其中：

GR 包括：

- GR
- GREAT RESOURCES
- Company-EL
- Company-ACMV
- Company-FP
- Company-PS

"-" 两边允许有空格。

company.xlsx 由人工维护。

程序不会自动修改 company.xlsx。

---

### Sheet2

根据：

Recipient's Account Name

自动提取：

Key word

---

## matcher.py

负责：

所有 Keyword 相同情况下的金额匹配。

包括：

1↔1

1↔2

2↔1

...

一直到

1↔10

10↔1

以及：

2↔2

2↔3

...

等组合。

原则：

金额必须完全一致。

---

## keyword_difference_matcher.py

负责：

Keyword 不同，

但业务规则可以确认的匹配。

包括：

### Rule 1

唯一金额

1↔1

Keyword 不同。

---

### Rule 2

重复金额

先消除 Keyword 相同。

如果剩余数量一致，

逐条对应。

---

### Rule 3

金额差额补齐

允许：

最多三条空白 Keyword

补足金额。

例如：

2↔1

3↔1

这些记录：

Excel 中显示：

Keyword Difference (2-to-1)

Keyword Difference (3-to-1)

---

## final_yellow_matcher.py

负责：

最后剩余黄色记录。

特点：

不要求：

Key word 相同。

仅要求：

金额一致。

支持：

1↔1

1↔2

...

一直到

1↔6

及反向。

不支持：

2↔2

以及更复杂组合。

所有结果：

蓝色。

Match Type：

Final Yellow Match (1-to-4)

等。

---

## validate_matches.py

负责：

结果验证。

检查：

✓ Partner Rows 双向一致

✓ 不引用不存在记录

✓ 每条记录只能匹配一次

验证失败：

停止生成结果文件。

---

## difference_analyzer.py

负责：

最终剩余记录。

计算：

Sheet1

减

Sheet2

剩余差额。

寻找：

最小组合。

用于：

人工最终审核。

---

## excel_io.py

负责：

Excel 输入输出。

以及：

颜色。

---

## models.py

数据结构定义。

Record

等。

---

# 4. Color Rules

绿色

Key word 相同。

可靠匹配。

---

浅棕色

Keyword 不同。

但业务规则确认。

例如：

Keyword Difference

---

蓝色

Final Yellow Match

最后金额匹配。

不要求 Keyword 相同。

---

黄色

最终未匹配。

需要人工处理。

---

# 5. Matching Principles

整个系统坚持：

优先业务规则。

其次金额。

最后人工审核。

绝不为了提高匹配率而强行组合。

---

# 6. Validation

程序生成 Excel 前，

必须通过：

Validation。

只有：

Validation Passed

才生成：

result_reconciliation.xlsx

---

# 7. Current Matching Rules

Keyword 相同：

最高：

1↔10

Keyword 不同：

Final Yellow：

最高：

1↔6

Difference Analysis：

最多：

10 条组合。

---

# 8. Future Improvements

计划：

- 自动日志（Run Log）
- Match Statistics Dashboard
- HTML Report
- Batch Processing
- Rule Configuration
- SQLite Database
- Performance Optimization

---

# Version

Current Version

V1.0
