# Bar-Chart-Adjustment
# Excel VBA Macros for Fiscal Year & Shape Positioning

This repository contains VBA macros designed for Excel to work with fiscal years, quarters, and dynamically position shapes based on cell values within an Excel workbook.

## 1. FiscalYearAndQuarter Function

### Description

The `FiscalYearAndQuarter` function calculates the fiscal year and quarter for a given date. The fiscal year starts in July and ends in June. The function returns a string in the format "FY[Year] [Quarter]".

### Usage

- **Input**: Date value for which fiscal year and quarter need to be determined.
- **Output**: String in the format "FY[Year] [Quarter]".

## 2. CalculateColumnOffset Function

### Description

The `CalculateColumnOffset` function calculates the column offset based on a given date, the start fiscal year, and the starting column index. This is useful for mapping dates to specific columns based on fiscal year and quarter.

### Usage

- **Input**:
  - `dateValue`: Date for which column offset is calculated.
  - `startFY`: The starting fiscal year.
  - `startColumn`: The column index where the fiscal year begins.
- **Output**: Integer representing the column offset.

## 3. ReportAndPositionShapesUntilBlankRows Subroutine

### Description

The `ReportAndPositionShapesUntilBlankRows` subroutine dynamically positions shapes in an Excel worksheet based on the start and end dates specified in rows. The shapes represent different phases like "Design", "Procurement", and "Construction". The subroutine stops when it encounters four consecutive empty rows.

### Usage

- Automatically positions shapes based on the start and end date columns for each phase.
- The shapes are named and positioned dynamically, adapting to the date values.

### Assumptions

- The worksheet is named "Sheet1".
- The fiscal year starts in July.
- Phases and corresponding columns are predefined in the `phaseColumns` array.
- The subroutine starts from row 8 and checks for four consecutive empty rows to terminate.

---

These macros are part of a larger project to automate and enhance data visualization and reporting in Excel using VBA. They are designed to work together but can be adapted for separate use.

