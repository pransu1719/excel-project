<div align="center">

# -- ! Excel Fundamentals Booster ! --
### *A Practice Workbook for Core Excel Formulas & Functions*

[![Excel](https://img.shields.io/badge/Excel-2019%2F365-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/microsoft-365/excel)
[![Formulas](https://img.shields.io/badge/Formulas-Lookup%20%26%20Logic-FF6F00?style=for-the-badge&logo=googlesheets&logoColor=white)](https://support.microsoft.com/excel)
[![Sheets](https://img.shields.io/badge/Sheets-4%20Datasets-4CAF50?style=for-the-badge&logo=readthedocs&logoColor=white)](#-sheets-overview)
[![Level](https://img.shields.io/badge/Level-Beginner%20to%20Intermediate-9C27B0?style=for-the-badge&logo=bookstack&logoColor=white)](#-purpose)

<br/>

> *"A formula is only as good as the data it points to — master lookups, and the rest follows."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [📊 Sheets Overview](#-sheets-overview)
- [🧮 Function Reference](#-function-reference)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [✅ Notes](#-notes)
- [📄 License](#-license)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Excel Fundamentals Booster** is a sample workbook built for practicing core Excel formulas and functions. It contains **4 realistic datasets** — students, sales, employees, and lookup/master data — each paired with demo formulas showing how different Excel functions are used in practice.

This workbook is designed to:
- Strengthen understanding of conditional, lookup, and text functions
- Practice real-world formula patterns across multiple linked sheets
- Demonstrate both modern (`XLOOKUP`, `FILTER`) and legacy-safe (`INDEX/MATCH`) approaches
- Provide a single, self-contained reference file for Excel practice

---

## 🎯 Problem Statement

> **Objective:** Build a single workbook that demonstrates beginner-to-intermediate Excel functions across realistic, linked datasets.

Learners often practice Excel functions in isolation, without seeing how they connect across sheets in a real workbook. This project solves that by building four interlinked sheets — student records, a lookup/master table, sales transactions, and employee data — where formulas in one sheet actively reference another, just like in a real spreadsheet system.

| 📂 Sheet | 📄 Type | 🔍 Description |
|----------|---------|----------------|
| Students | Dataset | Marks, totals, averages, and grade logic |
| Lookup_Data | Reference | Master product & salesperson data |
| Sales | Dataset | Transactions, pricing, discount & tax |
| Employees | Dataset | Department, dates, salary & bonus logic |

The goal is to demonstrate **fundamental-to-intermediate Excel skills** through practical, interconnected, real-world-style examples.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📊 **4 Linked Sheets** | Students, Lookup_Data, Sales, and Employees, cross-referencing one another |
| 🔢 **Conditional Logic** | Grades, bonus rules, and multi-condition checks via `IF` / `AND` |
| 🔍 **Lookup Toolkit** | `VLOOKUP`, `XLOOKUP`, `INDEX/MATCH`, `XMATCH`, `OFFSET`, `INDIRECT` |
| 📐 **Aggregate Functions** | `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` for filtered summaries |
| 📅 **Date Logic** | Age and tenure calculations using `DATEDIF`-style formulas |
| 🔤 **Text Functions** | `LEFT`, `UPPER` for name parsing and formatting |
| 🎯 **Rounding Functions** | `ROUND`, `CEILING`, `FLOOR` for bonus calculations |
| 🛡️ **Legacy-Safe Alternatives** | `INDEX/MATCH` fallback provided wherever `XLOOKUP`/`XMATCH` is used |

---

## 🏗️ Project Structure

```
📦 excel-fundamentals-booster/
│
├── 📄 Excel_Fundamentals_Booster.xlsx   ← Main workbook (4 sheets)
│
└── 📄 README.md                          ← Project documentation
```

---

## 🔄 Project Workflow

```
Open Workbook
      │
      ▼
┌─────────────────────────────┐
│   Lookup_Data (Master)      │  ← Product Master + Salesperson sales
└────────────┬────────────────┘
             │ referenced by
     ┌───────┼────────────────┐
     ▼                        ▼
┌─────────────┐        ┌──────────────────┐
│  Students   │        │      Sales        │
│  Marks/Grade│        │  Price/Tax/Region  │
└──────┬──────┘        └────────┬──────────┘
       │                        │
       ▼                        ▼
┌─────────────┐        ┌──────────────────┐
│ Totals, Avg │        │  SUMIFS, VLOOKUP  │
│ IF / AND    │        │  OFFSET, INDIRECT │
└──────┬──────┘        └────────┬──────────┘
       │                        │
       └───────────┬────────────┘
                    ▼
           ┌──────────────────┐
           │    Employees      │
           │ Age / Bonus /     │
           │ XLOOKUP Salary    │
           └────────┬──────────┘
                    │
                    ▼
            Formula Results
            Displayed In-Sheet
```

---

## 📊 Sheets Overview

### 1️⃣ Students
Student-wise Math and Science marks, with Total, Average, and Grade calculated from them.

**Covered functions:**
- `SUM`, `AVERAGE` — Total and average score
- `IF` — Assigning grades (A/B/C/D/F)
- `AND` — Checking two conditions together (`Above80_Both`)
- Text functions — `LEFT` for First Name, `UPPER` for capitalized names
- `COUNTIFS` — Counting students matching given conditions
- `AVERAGEIFS` — Conditional average
- `XMATCH` / `MATCH` — Finding a record's position in a list
- `FILTER` — Filtering records that match a condition

### 2️⃣ Lookup_Data
Reference (master) tables — Product Master and Salesperson-wise monthly sales.

This sheet is mainly used as the **lookup source** for formulas in other sheets (`VLOOKUP`, `INDEX/MATCH`, `OFFSET`, `INDIRECT`, etc.).

### 3️⃣ Sales
Sales transaction data — order-wise product, quantity, price, discount, and tax calculations.

**Covered functions:**
- `VLOOKUP` — Fetching Unit Price from Product Code
- `SUMIFS` — Total sales filtered by Region + Product
- `INDEX`/`MATCH` and `XLOOKUP` — Finding a salesperson's sales value for a specific month
- `XMATCH` / `MATCH` — Position lookup
- `OFFSET` — Dynamic 3-month rolling sum
- `INDIRECT` — Building a dynamic range reference from text
- Absolute reference (`$`) — Using a fixed value like Tax Rate

### 4️⃣ Employees
Employee records — Department, Date of Birth, Date of Joining, Salary, and Bonus calculations.

**Covered functions:**
- Date formulas (`DATEDIF`-style) — Calculating Age and Days Since Joining
- `ROUND`, `CEILING`, `FLOOR` — Rounding, ceiling, and flooring the bonus amount
- `XLOOKUP` and `INDEX/MATCH` (safe alternative) — Search-tool demo to find Salary by entering an EmpID

---

## 🧮 Function Reference

| Category | Functions Used |
|----------|-----------------|
| 🔍 **Lookup** | `VLOOKUP`, `XLOOKUP`, `INDEX/MATCH`, `XMATCH`, `MATCH`, `OFFSET`, `INDIRECT` |
| 🔢 **Conditional** | `IF`, `AND`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `FILTER` |
| 📅 **Date** | `DATEDIF`-style Age and tenure calculations |
| 🎯 **Rounding** | `ROUND`, `CEILING`, `FLOOR` |
| 🔤 **Text** | `LEFT`, `UPPER` |
| ➕ **Aggregate** | `SUM`, `AVERAGE` |

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | 2019 / 365 | Core spreadsheet application |
| 🔍 **Lookup Functions** | Built-in | Cross-sheet data retrieval |
| 🔢 **Logical Functions** | Built-in | Conditional grading and bonus rules |
| 📅 **Date Functions** | Built-in | Age and tenure calculations |
| 🎯 **Math Functions** | Built-in | Rounding and threshold logic |
| 📌 **Absolute References** | Built-in | Fixed values like Tax Rate |

---

## 📈 Results & Insights

After opening and exploring the workbook, the following are available:

- ✅ **4 Interlinked Sheets** — Students, Lookup_Data, Sales, and Employees
- 🔢 **Automated Grading** — Every student's grade calculated from Total/Average via `IF`
- 💰 **Sales Calculations** — Price, discount, and tax resolved through `VLOOKUP` + `SUMIFS`
- 📅 **Age & Tenure** — Automatically computed for every employee record
- 🛡️ **Version-Safe Formulas** — Legacy `INDEX/MATCH` alternatives alongside modern functions

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Beginner Friendly** | Core-to-intermediate functions in one linked workbook |
| 🔄 **Realistic Data Flow** | Sheets reference each other like a real business workbook |
| 📚 **Educational** | Each sheet reinforces a different function category |
| 🖥️ **No Add-ins Needed** | Runs with standard Excel — no plugins required |
| ⚡ **Self-Contained** | Single-file workbook, ready to open and explore |
| 🧪 **Extensible** | Easy to add new sheets or formula demos |
| 🛡️ **Fallback Formulas** | Safe alternatives included for older Excel versions |

---

## ✅ Notes

- Before editing any formula, check the related references in the **Lookup_Data** sheet — many formulas depend on it.
- Functions like `XLOOKUP`, `XMATCH`, and `FILTER` only work in Excel 365 / Excel 2021+. On older Excel versions (2019 or earlier), these may return a `#NAME?` error — use the "safe alternative" `INDEX/MATCH` formulas instead in that case.
- The Tax Rate (Sales sheet) and other key assumptions are highlighted right in the header, so they're easy to find when updating formulas.

---

## 📄 License

This project is licensed under the **MIT License** — free to use, modify, and distribute with attribution.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 🙏 Acknowledgements

Special thanks to the following resources that made this project possible:

- 📚 [Microsoft Excel Support](https://support.microsoft.com/excel) — Official function reference
- 🔍 [Exceljet](https://exceljet.net/) — Formula examples and explanations
- 📐 [GeeksForGeeks — Excel Functions](https://www.geeksforgeeks.org/) — Function usage examples
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---
pransu


</div>
