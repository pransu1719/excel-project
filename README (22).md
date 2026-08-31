# Excel Fundamentals Booster 📊

A hands-on Excel workbook for practicing real-world spreadsheet formulas — built around four themed datasets (students, sales, employees, and lookup tables) instead of dry, generic exercises.

## 🎯 What This Is

Most "learn Excel" resources teach formulas in isolation. This workbook flips that: every function is demonstrated inside a realistic mini-scenario — grading students, calculating sales tax, computing employee bonuses — so you see *why* a formula matters, not just its syntax.

## 📁 Workbook Structure

| Sheet | Scenario | What It Covers |
|---|---|---|
| **Students** | Grade analysis for 15 students across Math & Science | `COUNTIFS`, `AVERAGEIFS`, `AND`, `TEXT`, `UPPER`, `XMATCH` / `MATCH`, `FILTER` |
| **Lookup_Data** | Product master list + salesperson monthly performance | Reference/master tables used by other sheets |
| **Sales** | 10 order transactions with tax and discounts | `VLOOKUP`, `INDEX`/`MATCH`, `XLOOKUP`, `XMATCH`, `SUMIFS`, `OFFSET`, `INDIRECT`, absolute references (`$`) |
| **Employees** | 8 employee records with salary & bonus logic | Date math (age, tenure), `ROUND`, `CEILING`, `FLOOR`, interactive `XLOOKUP` search box |

## 🧠 Key Concepts Practiced

- **Lookups:** `VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH`, `XMATCH` — including "safe alternative" formulas for versions of Excel that don't support the newer functions
- **Conditional aggregation:** `COUNTIFS`, `AVERAGEIFS`, `SUMIFS`
- **Dynamic ranges:** `OFFSET` for rolling sums, `INDIRECT` for building references from text
- **Logic & text:** `AND`, `TEXT`, `UPPER`
- **Dates & rounding:** age/tenure calculations, `ROUND`, `CEILING`, `FLOOR`
- **Absolute vs. relative referencing** (`$` locking) for tax and discount calculations
- **Interactive lookup tool** — type an Employee ID and watch `XLOOKUP` pull the salary live

## 🚀 How to Use It

1. Open `Excel_Fundamentals_Booster.xlsx` in Excel (or LibreOffice Calc / Google Sheets).
2. Pick a sheet and read the "Analysis / Formula Demos" section at the bottom — each row names the function being demonstrated.
3. Click into any result cell to see the actual formula and reverse-engineer how it works.
4. Try changing an input value (e.g., the EmpID search box on **Employees**, or the Tax Rate on **Sales**) and watch the dependent formulas recalculate.
5. For extra practice, delete a formula and try to rebuild it from scratch using the pattern shown elsewhere in the sheet.

## ✅ Status

Downloaded and verified — ready to use as a self-paced Excel practice reference or a teaching aid.

## 💡 Ideas for Extending It

- Add a **Dashboard** sheet with charts summarizing grades, sales by region, or bonus totals by department
- Introduce `PivotTable`s built from the Sales sheet
- Add data validation dropdowns (e.g., restrict Region or Grade inputs)
- Convert a few static formulas into named ranges for readability
