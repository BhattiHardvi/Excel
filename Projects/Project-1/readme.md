# PR. 1 Fundamental Booster – Excel Project

## Introduction
This project demonstrates the use of Microsoft Excel formulas and functions for working with student grades, sales records, and employee information. It provides practical experience with calculations, logical conditions, lookup functions, text functions, date functions, and dynamic references.

## Workbook Structure

### 1. Students Grade
This worksheet demonstrates:
- `SUM` – Calculates total marks.
- `AVERAGE` – Calculates average marks.
- `IF` – Assigns grades or pass/fail results.
- `AND` and `OR` – Check multiple conditions.
- `COUNTIF` – Counts records matching a condition.
- `AVERAGEIF` – Calculates an average based on a condition.
- `FILTER` – Displays rows matching a condition.

**Example formulas:**
```excel
=SUM(C2:E2)
=AVERAGE(C2:E2)
=IF(C2>50,"Pass","Fail")
=IF(AND(C2>80,D2>80),"Yes","No")
=COUNTIF(G2:G13,">60")
=FILTER(A2:J13,G2:G13>80,"No matches")
```

### 2. Sales Data
This worksheet demonstrates:
- `VLOOKUP` and `XLOOKUP` – Retrieve product information.
- `INDEX`, `MATCH`, and `XMATCH` – Find matching values.
- `SUMIFS` – Calculate sales based on multiple conditions.
- `INDIRECT` and `OFFSET` – Create dynamic references.
- `TEXT`, `LEFT`, `UPPER`, and `LOWER` – Manipulate text.
- `ROUND`, `CEILING`, and `FLOOR` – Round numerical values.

**Example formulas:**
```excel
=VLOOKUP("P103",L3:N6,3,FALSE)
=XLOOKUP("P104",L3:L6,M3:M6,"Not found")
=INDEX(H2:H11,MATCH(2006,A2:A11,0))
=XMATCH("P103",L3:L6)
=SUMIFS(H2:H11,C2:C11,"West",E2:E11,"Laptop")
=TEXT(B2,"mmm yyyy")
=LEFT(I2,3)
=UPPER(I2)
=LOWER(I2)
=ROUND(H2,0)
=CEILING(G2,100)
=FLOOR(G2,100)
```

### 3. Employee Data
This worksheet demonstrates:
- `VLOOKUP` and `XLOOKUP` – Find employee information.
- `DATEDIF` and `TODAY` – Calculate age and years of service.
- Date subtraction – Calculate the difference between dates.
- `LEFT` and `FIND` – Extract employee first names.
- `FILTER` – Display employees matching a condition.
- `INDIRECT` and `OFFSET` – Create dynamic references.

**Example formulas:**
```excel
=XLOOKUP(503,A2:A7,E2:E7,"Not found")
=VLOOKUP(502,A2:H7,2,FALSE)
=DATEDIF(G2,TODAY(),"Y")
=DATEDIF(D2,TODAY(),"Y")
=ABS(D3-D2)
=TODAY()-D2
=FILTER(A2:H7,E2:E7>50000,"No matches")
=INDIRECT(K9)
=SUM(OFFSET(E2,0,0,3,1))
=LEFT(B2,FIND(" ",B2)-1)
```

### 4. Formula Guide
The Formula Guide worksheet provides a summary of the Excel functions used in this project with their purposes and examples.

## Relative and Absolute Cell References

**Relative Reference**
```excel
=C2+D2
```
Cell references change automatically when the formula is copied.

**Absolute Reference**
```excel
=C2*$P$3
```
The `$P$3` reference remains fixed when the formula is copied.

## How to Use the Workbook
1. Open the Excel workbook in Microsoft Excel.
2. Select the required worksheet.
3. Enter or update the records.
4. Apply the required formulas.
5. Copy formulas down to other rows.
6. Verify the results.

## Compatibility
Functions such as `XLOOKUP`, `XMATCH`, and `FILTER` require a recent version of Microsoft Excel.

## Note About Sample Data
The solution workbook contains illustrative sample records because the assignment screenshots provided instructions but did not include the original datasets. Replace the sample records and adjust formula ranges according to the actual assignment data.

## Conclusion
This project provides practical knowledge of Excel formulas and functions used for calculations, grade classification, conditional analysis, data lookup, text manipulation, date calculations, and data filtering.