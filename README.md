# Task 6 – Excel Formulas & Functions Fundamentals

**Dataset:** Sample Superstore (real dataset). The `Data` tab uses the first 500 of 9,994 rows. `Lookups` tables (49 states -> region, 17 sub-categories -> category) are derived from the full dataset.
**Workbook:** `Task6_Excel_Formulas_Functions.xlsx`

| Tab | Content |
|---|---|
| Formula_Demo | Every function on live data; yellow input cells (dropdowns), formula shown beside each result |
| Data | 500 real rows + 10 calculated columns (UPPER/LEFT, PROPER/TRIM, VLOOKUP, INDEX/MATCH, IF, nested IF, OR, ROUND) |
| Lookups | State->Region and Sub-Category->Category tables |
| Notes | When to use each formula, watch-outs, interview answers |

**Functions covered:** SUMIF, SUMIFS, COUNTIF, COUNTIFS, COUNTBLANK, AVERAGEIFS, VLOOKUP, XLOOKUP, INDEX/MATCH, IF, nested IF, AND/OR, IFERROR, TRIM, PROPER, UPPER, LEN, LEFT, MID, FIND, SUBSTITUTE, TEXT, VALUE, ISNUMBER, N, named ranges.

**Edge cases tested:** extra spaces before lookup, number stored as text, blank cell, division by zero, lookup value not found.

**Interview answers**
- *SUMIF vs SUMIFS:* SUMIF = one condition, SUMIFS = many; sum_range is last in SUMIF, first in SUMIFS.
- *XLOOKUP vs VLOOKUP:* looks left or right, exact match by default, no column index, built-in not-found value.
- *INDEX/MATCH vs VLOOKUP:* when the return column is left of the key, for older Excel, or when columns may change.

Note: XLOOKUP needs Excel 365/2021+ or Google Sheets.
