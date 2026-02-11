# Back Down and Non-compliance Report Automation

## Overview

This script automates the generation of the "Back Down and Non-compliance" Excel report that was previously created manually. It reads input files and generates a fully formatted output Excel file with all calculations.

## Requirements

- Python 3.6+
- openpyxl library: `pip install openpyxl`

## Input Files Required

1. **Reference calculation sheet** (OUTPUT/REFERENCE - manually generated):
   - File pattern: `*calculation*sheet*.xlsx`
   - Sheet "New method" with Date, From, To columns
   - Each date in Column B marks start of a new instruction period
   - This file is used as REFERENCE to extract instruction periods and understand output format

2. **HNPCL revised DC file** - Contains Declared Capacity values:
   - File pattern: `*HNPCL*revised*DC*.xlsx`
   - One sheet per day (e.g., "01.01.2026", "02.01.2026", etc.)
   - Each sheet has: TB No, From, To, Day Ahead, **Final Revision** (used), Remarks

3. **Daily BD LR files** - Contains SCADA readings:
   - File pattern: `BD LR DD-MM-YYYY.xlsx` or `BD LR_MBED DD-MM-YYYY.xlsx`
   - One file per day
   - Must have "SCADA Grid" sheet with time and SCADA values

## Usage

### Basic Usage

```bash
python generate_bd_report.py input/
```

This will:
- Search for reference calculation sheet in the input directory
- Extract instruction periods from the reference file (these were manually entered)
- Search for HNPCL revised DC file in the input directory (INPUT)
- Search for Daily BD LR files in the input directory (INPUT)
- Generate output file matching reference format: `input/Back_down_and_Non_compliance_output.xlsx`

### Advanced Usage

```bash
python generate_bd_report.py input/ \
    --reference "input/calculation sheet for BD and non compliance of HNPCL for Jan 26.xlsx" \
    --dc-file "input/HNPCL revised DC for the month January 2026 SLDC.xlsx" \
    --output output/report.xlsx
```

### Options

- `input_directory` (required): Directory containing input files and reference file
- `--reference`: Path to reference calculation sheet file (default: auto-search for `*calculation*sheet*.xlsx`)
- `--dc-file`: Path to HNPCL revised DC file (default: auto-search)
- `--output`: Output file path (default: `input_directory/Back_down_and_Non_compliance_output.xlsx`)

## Output Format

The generated Excel file contains:

### Header Structure
- Row 1: Title "Back down and Non compliance of HNPCL for Oct 2025" (merged A-J)
- Row 2: Category headers "Back down" (A-G) and "Non compliance" (H-J)
- Row 3: Column headers

### Columns

**Back down section:**
- Column B: Date
- Column C: From (time)
- Column D: To (time)
- Column E: DC (MW) - from HNPCL Revised DC file
- Column F: As per SLDC Scada in MW - from Daily BD LR files (yellow highlighted)
- Column G: Diff (MW) - Formula: `=E-F`
- Column H: Mus - Formula: `=G/4000`

**Non compliance section:**
- Column J: MW as per ramp - Formula:
  - First cell: `=E{row}-A` (where A = 40 for 15-min, 27.5 for 10-min, 15 for 5-min)
  - Subsequent: `=MAX(J{prev}-40, 270)`
- Column K: Diff - Formula: `=F-J`
- Column L: MU - Formula: `=IF(K/4000>0, K/4000, 0)`

## Calculation Logic

### Column J (MW as per ramp) Calculation

Based on handwritten notes:

1. **First cell of instruction period:**
   - Formula: `=E{row} - A`
   - Variable A depends on instruction duration:
     - 15 minutes: A = 40 MW
     - 10 minutes: A = 27.5 MW
     - 5 minutes: A = 15 MW

2. **Subsequent cells:**
   - Formula: `=MAX(J{prev_row}-40, 270)`
   - Decreases by 40 MW per 15-minute interval
   - Minimum value: 270 MW

### Other Calculations

- **Column G (Diff MW)**: Difference between DC and SCADA
- **Column H (Mus)**: Back down MUs = Diff MW / 4000
- **Column K (Diff)**: Difference between SCADA and ramp
- **Column L (MU)**: Non-compliance MUs (only positive values)

## Data Matching

The script matches data across files by:

1. **Date matching**: Instruction period date → Daily BD LR file → HNPCL Revised DC sheet
2. **Time matching**: 15-minute intervals aligned across all files
3. **Slot generation**: Instruction periods are split into 15-minute slots

## Date Grouping

- Date appears only on the first row of each instruction period
- Subsequent rows have empty date cells (grouped under same date)

## Example

```bash
# Generate report from input directory
python generate_bd_report.py input/

# Output:
# Reading instruction periods from: INSTRUCTION.xlsx
# Found 8 instruction periods
# Generated 54 data rows
# Fetching DC and SCADA values...
# Output saved to: input/Back_down_and_Non_compliance_output.xlsx
# Total rows created: 54
```

## Troubleshooting

### Error: "No reference calculation sheet found"
- Ensure reference calculation sheet exists in the input directory
- File name should contain "calculation" and "sheet"
- This is the OUTPUT/REFERENCE file that was manually generated
- Or specify path with `--reference` option

### Error: "No HNPCL revised DC file found"
- Ensure the DC file exists in the input directory
- File name should contain "HNPCL" and "revised" and "DC"
- Or specify path with `--dc-file` option

### Missing SCADA values
- Check that Daily BD LR files exist for all dates in instruction periods
- Verify files have "SCADA Grid" sheet
- Check that time formats match (15-minute intervals)

### Missing DC values
- Verify HNPCL Revised DC file has sheets for all dates
- Check sheet names match date format (DD.MM.YYYY)
- Verify "Final Revision" column (Column E) has values

## Notes

- The script generates formulas in Excel, so values will recalculate when the file is opened
- Date format in output matches Excel date format
- Time values are formatted as HH:MM
- Yellow highlighting is applied to Column F (As per SLDC Scada in MW)
