# Excel Files Analysis - Output Generation Logic

## Executive Summary

This document analyzes the Excel files in the `/input` directory to understand how the output "Back down and Non compliance" sheet is generated, based on the actual files and handwritten calculation notes.

---

## File Structure and Relationships

### 1. **Calculation Sheet** (`calculation sheet for BD and non compliance of HNPCL for Jan 26.xlsx`)

**Purpose**: Main output/template file showing the final calculated results.

**Sheet: "New method"**
- **Total Rows**: 1,694 data rows
- **Column Structure**:
  - Column B: Date
  - Column C: From (time)
  - Column D: To (time)
  - Column E: DC (MW) - Declared Capacity
  - Column F: As per SLDC Scada in MW (yellow highlighted)
  - Column G: Diff (MW) - Back down difference
  - Column H: Mus - Back down MUs
  - Column J: MW as per ramp
  - Column K: Diff - Non-compliance difference
  - Column L: MU - Non-compliance MUs

**Key Observations**:
- Date appears only on the first row of each instruction period
- Subsequent rows have empty date cells
- Time intervals are in 15-minute blocks (e.g., 20:15-20:30, 20:30-20:45)

---

### 2. **HNPCL Revised DC File** (`HNPCL revised DC for the month January 2026 SLDC.xlsx`)

**Purpose**: Source for Column E (DC MW) values.

**Structure**:
- 32 sheets (one per day: 01.01.2026 through 31.01.2026, plus Sheet1)
- Each sheet contains:
  - Column A: TB No (Time Block Number)
  - Column B: From (time)
  - Column C: To (time)
  - Column D: Day Ahead (initial DC value)
  - Column E: Final Revision (revised DC value) ← **Used in output**
  - Column F: Remarks

**Example** (01.01.2026 sheet):
- Row 3: TB No 1, From 00:00, To 00:15:00, Day Ahead 492.7, Final Revision 492.7
- Row 4: TB No 2, From 00:15:00, To 00:30:00, Day Ahead 492.7, Final Revision 492.7

**Relationship**: The "Final Revision" column values are copied to Column E (DC MW) in the calculation sheet, matched by date and time interval.

---

### 3. **Daily BD LR Files** (31 files: `BD LR 01-01-2026.xlsx` through `BD LR 31-01-2026.xlsx`)

**Purpose**: Source for Column F (As per SLDC Scada in MW) values.

**Key Sheet: "SCADA Grid"**
- Contains actual SCADA readings
- Column A: Time (datetime format: 2026-01-01 00:00:00)
- Column B: SYSCA_AT.SYSTEM.GRID_DMD_
- Column C: SCHED_PG.SYSTEM.SR_FREQ.H
- Column D: SCHED_PG.SYSTEM.AP_UI.MW ← **Likely source for SCADA values**
- Other columns: Solar, Wind, Block No, etc.

**Other Relevant Sheets**:
- "BD Start": Contains 15-minute time intervals (00:00-00:15, 00:15-00:30, etc.)
- Multiple other sheets for different calculations

**Relationship**: SCADA values from these daily files are matched by date and time to populate Column F in the calculation sheet.

---

## Output Generation Logic (Based on Handwritten Notes and File Analysis)

### Column J: "MW as per ramp" Calculation

**Formula Logic** (from handwritten notes):

1. **First Cell of Instruction Period**:
   ```
   J = (X - A)
   ```
   Where:
   - X = DC (MW) value from Column E
   - A = Variable based on starting time within instruction period:
     - **15 minutes duration**: A = 40 MW
     - **10 minutes duration**: A = 27.5 MW
     - **5 minutes duration**: A = 15 MW

2. **Subsequent Cells**:
   ```
   J(n) = J(n-1) - 40
   ```
   Each subsequent 15-minute interval reduces by 40 MW.

3. **Minimum Value Constraint**:
   - If calculated value < 270 MW, set to **270 MW**
   - Once 270 is reached, maintain 270 for all remaining cells until the end of the instruction period

**Actual Implementation** (from calculation sheet):
- Row 5: `=E5-40` (first cell, assuming 15-min period)
- Row 6: `=J5-40`
- Row 7: `=J6-40`
- Row 8: `=J7-40`
- Row 9: `=J8-40`
- Row 10 onwards: Hardcoded `270` (minimum value applied)

**Example Calculation**:
- Row 5: DC = 492.7, J = 492.7 - 40 = 452.7
- Row 6: J = 452.7 - 40 = 412.7
- Row 7: J = 412.7 - 40 = 372.7
- Row 8: J = 372.7 - 40 = 332.7
- Row 9: J = 332.7 - 40 = 292.7
- Row 10: J = 292.7 - 40 = 252.7 → **Set to 270** (minimum constraint)
- Row 11 onwards: J = 270 (maintained)

---

### Other Column Calculations

**Column G: Diff (MW)** - Back down difference
```
G = E - F
```
Difference between Declared Capacity and SCADA reading.

**Column H: Mus** - Back down MUs
```
H = G / 4000
```
Convert MW difference to MUs (Mega Units).

**Column K: Diff** - Non-compliance difference
```
K = F - J
```
Difference between SCADA reading and ramp value.

**Column L: MU** - Non-compliance MUs
```
L = IF(K/4000 > 0, K/4000, 0)
```
Conditional calculation: only positive differences converted to MUs.

---

## Data Flow Diagram

```
┌─────────────────────────────────────────────────────────────┐
│  Instruction Periods (Date, From Time, To Time)              │
│  - Defines time ranges for analysis                          │
└───────────────────────┬───────────────────────────────────────┘
                        │
        ┌───────────────┴───────────────┐
        │                               │
        ▼                               ▼
┌──────────────────┐          ┌──────────────────┐
│ HNPCL Revised DC │          │ Daily BD LR     │
│ File             │          │ Files            │
│                  │          │                  │
│ Column E:        │          │ Column F:        │
│ Final Revision   │          │ SCADA readings   │
│ (DC MW values)   │          │ (15-min intervals)│
└────────┬─────────┘          └────────┬─────────┘
         │                            │
         └────────────┬───────────────┘
                      │
                      ▼
         ┌────────────────────────────┐
         │  Calculation Sheet         │
         │  "New method" sheet        │
         │                            │
         │  Column E: DC (MW)        │
         │  Column F: SCADA (MW)     │
         │  Column J: Ramp (calculated)│
         │  Column G: Diff (E-F)      │
         │  Column K: Diff (F-J)      │
         │  Column H, L: MUs          │
         └────────────────────────────┘
```

---

## Key Patterns Identified

### 1. Time Alignment
- All files use **15-minute intervals**
- Format: HH:MM-HH:MM (e.g., 00:00-00:15, 00:15-00:30)
- Times are matched across files by date and interval

### 2. Date Grouping
- Date shown only on **first row** of each instruction period
- Subsequent rows have **empty date cell**
- This groups related time intervals under one date

### 3. Ramp Calculation Pattern
- Starts with DC - A (where A depends on start time)
- Decreases by 40 MW per 15-minute interval
- Floors at 270 MW minimum
- Can also increase (e.g., `=J22+40` seen in some rows)

### 4. Formula Dependencies
- Column J depends on Column E (DC) and previous J values
- Column G depends on Columns E and F
- Column K depends on Columns F and J
- Column L depends on Column K

---

## Data Matching Logic

To generate the output, data must be matched across files:

1. **Match by Date**: 
   - Instruction period date → Daily BD LR file
   - Instruction period date → HNPCL Revised DC sheet

2. **Match by Time Interval**:
   - Instruction period From/To → 15-minute block in source files
   - Example: 20:15-20:30 matches TB No corresponding to 20:15-20:30

3. **Extract Values**:
   - DC (MW) from HNPCL Revised DC → Column E
   - SCADA reading from Daily BD LR → Column F
   - Calculate Column J based on ramp logic
   - Calculate other columns using formulas

---

## Summary

The output generation process involves:

1. **Reading instruction periods** (Date, From, To)
2. **Fetching DC values** from HNPCL Revised DC file (matched by date/time)
3. **Fetching SCADA values** from Daily BD LR files (matched by date/time)
4. **Calculating ramp values** (Column J) using the formula: (DC - A) for first cell, then decreasing by 40 MW per interval, with 270 MW minimum
5. **Calculating differences and MUs** using the formulas defined above
6. **Formatting output** with proper header structure, date grouping, and cell formatting

The handwritten notes provide the specific logic for calculating Column J (MW as per ramp), which is the most complex calculation involving variable A based on instruction start time and a minimum value constraint of 270 MW.
