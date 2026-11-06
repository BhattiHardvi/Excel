# PR. 1 Fundamental Booster -- Excel Project

## Project Overview

This project is an Excel/Google Sheets practice workbook that
demonstrates spreadsheet formulas and data-analysis features. It
contains three datasets: **Students Grade**, **Sales Data**, and
**Employee Data**, plus a **Formula Guide** sheet.

## Workbook Sheets

### 1. Students Grade

Contains 20 student records with Student ID, Name, Math, Science,
English, and Enrollment Date. Formula columns calculate total marks,
average percentage, grade, and conditional checks.

**Formulas demonstrated:** - `SUM` and `AVERAGE` -- calculate total and
average marks. - Nested `IF` -- assign grades based on average marks. -
`AND` and `OR` -- test multiple conditions. - `COUNTIF` and `COUNTIFS`
-- count records matching conditions. - `AVERAGEIF` and `AVERAGEIFS` --
calculate averages using criteria. - `FILTER` -- extract students whose
average is above 80.

### 2. Sales Data

Contains 20 sales records with Sales ID, Product, Region, Salesperson,
Amount, and Date. Additional columns demonstrate month extraction, sales
classification, rounding, and lookup.

**Formulas demonstrated:** - `TEXT` -- display the month from a date. -
`IF` -- classify sales as High Sales or Low Sales. - `ROUND` -- round
sales amounts. - `SUM`, `SUMIFS`, and `AVERAGE` -- analyze sales. -
`VLOOKUP`, `XLOOKUP`, `INDEX`, and `MATCH` -- retrieve matching
information. - `XMATCH` -- find a product's position. - `INDIRECT` and
`OFFSET` -- demonstrate dynamic references and ranges. - Date
subtraction -- calculate the number of days between dates.

**Product price note:** Product prices were not included in the supplied
Sales Data. The Product Lookup area is provided for entering actual
prices. Add prices in the blank input cells before using the
product-price `VLOOKUP` example.

### 3. Employee Data

Contains 20 employee records with Employee ID, Name, Department, Salary,
and Joining Date. Additional columns calculate years of service, salary
category, uppercase name, salary lookup, joining month, and days since
joining.

**Formulas demonstrated:** - `DATEDIF` and `TODAY` -- calculate
completed years of service. - `IF` -- classify salary. - `UPPER` --
convert employee names to uppercase. - `VLOOKUP` and `XLOOKUP` --
retrieve employee details. - `INDEX` and `MATCH` -- find employee
information. - `COUNTIF`, `AVERAGE`, and `SUMIFS` -- analyze employee
data. - `TEXT` -- format joining dates. - `INDIRECT` and `OFFSET` --
demonstrate dynamic references.

### 4. Formula Guide

Includes example formulas and explanations for relative and absolute
references, nested `IF`, `AND`/`OR`, counting and summing functions,
lookup functions, text functions, date functions, dynamic references,
rounding functions, and `FILTER`.

## How to Use the Workbook

1.  Download `PR1_Fundamental_Booster_Completed.xlsx`.
2.  Open it in Microsoft Excel or import it into Google Sheets.
3.  Select a worksheet using the tabs at the bottom.
4.  Review the formulas in the calculated columns and analysis sections.
5.  Change a source value to practice and observe how formula results
    update.
6.  Enter actual product prices in the Product Lookup area of the Sales
    Data sheet to test product-price lookup formulas.
7.  Refer to the Formula Guide sheet for formula examples.

## Notes

-   Student, sales, and employee records were transcribed from the
    screenshots supplied for this project.
-   Formula results recalculate when opened in a spreadsheet application
    that supports the functions used.
-   `FILTER` and `XLOOKUP` may require a recent version of Microsoft
    Excel or Google Sheets.
-   Results using `TODAY()` change according to the current date.

## Files

-   `PR1_Fundamental_Booster_Completed.xlsx` -- completed workbook.
-   `README.md` -- project documentation.

## Conclusion

This project provides practical experience with spreadsheet formulas,
conditional logic, lookups, text functions, date calculations, dynamic
references, and filtering. It builds a foundation for working with and
analyzing data in Excel.
