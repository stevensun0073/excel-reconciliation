# Excel Reconciliation Engine

A Python-based Excel reconciliation engine for bank reconciliation.

The project focuses on automatic matching between a bank ledger and a bank statement, supporting multiple reconciliation strategies while preserving the original Excel files.

---

# Current Status

**Version:** V0.2

Completed

- ✅ Project structure
- ✅ Record data model
- ✅ Excel loading module
- ✅ One-to-One matching engine
- ✅ Amount tolerance (±0.01)
- ✅ Git version control
- ✅ .gitignore configuration

In Progress

- One-to-Two Match
- Two-to-One Match
- One-to-Three Match
- Three-to-One Match
- Two-to-Two Match
- Excel result output

---

# Project Structure

```
excel-reconciliation/

├── README.md
├── ROADMAP.md
├── CHANGELOG.md
├── .gitignore
├── requirements.txt
│
├── main.py
├── models.py
├── excel_io.py
├── matcher.py
│
├── data.xlsx              (Local only)
└── result_reconciliation.xlsx
```

---

# Current Architecture

```
Excel

    │

    ▼

excel_io.py

    │

    ▼

Record Objects

    │

    ▼

Matcher

    ├── ✅ One-to-One Match
    ├── ⏳ One-to-Two Match
    ├── ⏳ Two-to-One Match
    ├── ⏳ One-to-Three Match
    ├── ⏳ Three-to-One Match
    └── ⏳ Two-to-Two Match

    │

    ▼

Excel Output
```

---

# Matching Rules

## Implemented

### One-to-One Match

One ledger record matches one bank statement record.

Amount comparison allows a tolerance of **±0.01**.

---

## Planned

- One-to-Two Match
- Two-to-One Match
- One-to-Three Match
- Three-to-One Match
- Two-to-Two Match

---

# Amount Comparison Rule

All amount comparisons use the same rule.

```python
abs(a - b) <= Decimal("0.01")
```

This rule is applied consistently throughout the project.

---

# Data Model

Each transaction is represented as a Record object.

Fields include:

- Excel row number
- Amount
- Matched status
- Match type
- Partner row(s)

---

# Design Principles

- Never modify the original Excel file.
- Preserve original row order.
- Use Decimal for all amount calculations.
- Every record can only be matched once.
- Every matching algorithm uses the same tolerance rule (±0.01).
- Develop incrementally and verify every stage before moving forward.

---

# Development Roadmap

## V0.1

- Project initialization
- Excel loading
- Record model

✅ Completed

---

## V0.2

- One-to-One matching engine

✅ Completed

---

## V0.3

- One-to-Two matching

---

## V0.4

- Two-to-One matching

---

## V0.5

- One-to-Three matching

---

## V0.6

- Three-to-One matching

---

## V0.7

- Two-to-Two matching

---

## V0.8

- Excel result export

---

## V1.0

Complete Excel Reconciliation Engine

---

# Current Test Result

Using real reconciliation data:

```
Sheet1 Records : 1561

Sheet2 Records : 1374

One-to-One Matched : 1218

Remaining Sheet1 : 343

Remaining Sheet2 : 156
```

---

# License

Private project.