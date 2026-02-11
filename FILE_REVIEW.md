# Excel Files Review and Relationships

## Overview
This document reviews the Excel files in the `/input` directory and their relationships for processing Back Down (BD) and Non-compliance data for HNPCL.

## File Inventory

### 1. **INSTRUCTION.xlsx** (Expected but not found)
- **Purpose**: Source file containing instruction periods with Date, From Time, and To Time
- **Structure**: Should have columns for Date, From Time, To Time
- **Status**: Not present in current directory (script searches for it)
- **Usage**: Script reads this to extract time periods for BD/non-compliance analysis

### 2. **Daily BD LR Files** (31 files: `BD LR 01-01-2026.xlsx` through `BD LR 31-01-2026.xlsx`)
- **Purpose**: Daily detailed reports for each day of January 2026
- **Key Sheet: "BD Start"**
  - Contains 15-minute time blocks (00:00-00:15, 00:15-00:30, etc.)
  - Columns: Block, HOUR, RTPP Stg - I, RTPP Stg - II, RTPP Stg - III, RTPP Stg - IV, K'Patnam Stg - I, VTS (1-6), VTPS - IV, K'Patnam Stg - II, Hinduja, VTPS - V, SEIL-P2-125, SEIL-P2-502
  - Contains power values (MW) for each generator at each 15-minute interval
- **Other Sheets**: BD&LR, Individual BD, CGS BD CAL, Actual, Entitlement, ISGS, IND-GEN, DATA-CMD, SCADA Grid, GNA, TGNA, RTM, IEX, HP-DAM, TPCIL, CHECK, ISGS MOD, SEIL-P2 EBC MRI, SEIL-P2 check, ISGS BD
- **Relationship**: Provides actual power generation data in 15-minute intervals

### 3. **Calculation Sheet** (`calculation sheet for BD and non compliance of HNPCL for Jan 26.xlsx`)
- **Purpose**: Main calculation workbook for BD and non-compliance analysis
- **Key Sheet: "New method"**
  - **Header Structure** (matches desired output format):
    - Row 2: "Back down and Non compliance of HNPCL for Oct 2025" (title)
    - Row 3: "Back down" (merged A-G) | "Non compliance" (merged H-J)
    - Row 4: Column headers:
      - **Back down**: Date, From, To, DC (MW), As per SLDC Scada in MW, Diff (MW), Mus
      - **Non compliance**: MW as per ramp, Diff, MU
  - **Data Structure**:
    - Date column shows full date only on first row of each day
    - From/To columns show 15-minute time intervals
    - DC (MW): Declared Capacity values
    - As per SLDC Scada in MW: Actual SCADA readings (yellow highlighted column)
    - Diff (MW): =E-F (difference between DC and SCADA)
    - Mus: =G/4000 (MUs calculation)
    - MW as per ramp: Calculated ramp values
    - Diff: =F-J (difference between SCADA and ramp)
    - MU: =IF(K/4000>0,K/4000,0) (conditional MU calculation)
- **Relationship**: This is the template/output format that should be generated

### 4. **HNPCL Revised DC File** (`HNPCL revised DC for the month January 2026 SLDC.xlsx`)
- **Purpose**: Contains revised Declared Capacity (DC) values for each day
- **Structure**: 
  - One sheet per day (01.01.2026, 02.01.2026, etc.)
  - Each sheet contains:
    - TB No (Time Block Number)
    - From, To (15-minute intervals)
    - Day Ahead (initial DC value)
    - Final Revision (revised DC value)
    - Remarks
- **Example**: Row 3 shows TB No 1, From 00:00, To 00:15:00, Day Ahead 492.7, Final Revision 492.7
- **Relationship**: Source for "DC (MW)" column values

### 5. **Monthly Summary File** (`jan 2026.xlsx`)
- **Purpose**: Monthly aggregated data
- **Sheets**: monthstate, RTPP1, KPTM, VTPS, HNPCL, SEIL, Sheet1, Sheet2, Sheet3
- **Status**: Appears to be mostly empty or formatted differently
- **Relationship**: May contain monthly summaries or aggregations

## Data Flow and Relationships

```
┌─────────────────────────────────────────────────────────────┐
│                    INSTRUCTION.xlsx                         │
│  (Date, From Time, To Time - instruction periods)           │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│         Daily BD LR Files (01-31 Jan 2026)                  │
│  - BD Start sheet: 15-min intervals with power values      │
│  - SCADA Grid sheet: Actual SCADA readings                  │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│      HNPCL Revised DC File (January 2026 SLDC)              │
│  - Daily sheets with revised DC values                      │
│  - 15-minute intervals                                     │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│    Calculation Sheet (Template/Output Format)                │
│  - "New method" sheet with header structure                 │
│  - Formulas for Diff, Mus, MU calculations                 │
└─────────────────────────────────────────────────────────────┘
```

## Key Relationships

1. **Time Alignment**: All files use 15-minute intervals (00:00-00:15, 00:15-00:30, etc.)

2. **Data Sources**:
   - **DC (MW)**: From HNPCL Revised DC file → "Final Revision" column
   - **As per SLDC Scada in MW**: From Daily BD LR files → SCADA Grid sheet or similar
   - **Date, From, To**: From INSTRUCTION.xlsx (when periods are specified) or calculated from daily intervals

3. **Calculation Logic** (from Calculation Sheet):
   - Diff (MW) = DC (MW) - As per SLDC Scada in MW
   - Mus = Diff (MW) / 4000
   - MW as per ramp = Previous ramp value - 40 (or specific ramp calculation)
   - Diff (Non-compliance) = As per SLDC Scada in MW - MW as per ramp
   - MU (Non-compliance) = IF(Diff/4000 > 0, Diff/4000, 0)

4. **Output Format**:
   - Header structure matches Calculation Sheet "New method" sheet
   - Date column: Shows date only on first row of each day
   - From/To columns: Show 15-minute intervals
   - Yellow highlight on "As per SLDC Scada in MW" column

## Current Script Functionality

The `read_instructions.py` script:
1. Reads INSTRUCTION.xlsx (when available)
2. Extracts Date, From Time, To Time
3. Calculates 15-minute slots
4. Creates output Excel file with proper header structure
5. Fills Date, From, To columns
6. Leaves other columns empty for manual/data source filling

## Recommendations

1. **Data Integration**: The script should be enhanced to:
   - Read DC values from HNPCL Revised DC file
   - Read SCADA values from Daily BD LR files
   - Automatically populate all columns

2. **File Matching**: Need to match:
   - Instruction periods → Daily files
   - Dates → Corresponding daily sheets
   - Time intervals → 15-minute blocks

3. **Missing INSTRUCTION.xlsx**: 
   - Either create this file with instruction periods
   - Or modify script to work directly with daily files and extract periods from them

4. **Output Enhancement**:
   - Add formulas for calculated columns (Diff, Mus, MU)
   - Maintain date grouping (show date only on first row per day)
   - Apply proper formatting (yellow highlight, borders, etc.)
