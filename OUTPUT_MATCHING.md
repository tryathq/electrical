# Output File Matching Reference

## Answer to: "matched to which given file?"

**The output file should match the REFERENCE FILE format.**

## Reference File

**File:** `calculation sheet for BD and non compliance of HNPCL for Jan 26.xlsx`
**Sheet:** "New method"

This file was **manually generated** and is provided as a **reference/template** showing:
- Expected output format
- Instruction periods (Date, From, To)
- Formula structure
- Formatting (headers, merged cells, colors, borders)

## Output File Structure (Must Match Reference)

### Header Structure:
- **Row 1:** Empty
- **Row 2:** Title "Back down and Non compliance of HNPCL for [Month Year]" 
  - Merged: B2:J2
- **Row 3:** Category headers
  - "Back down" merged: B3:I3
  - "Non compliance" merged: J3:M3
- **Row 4:** Column headers (starting from Column B)
  - Back down: Date, From, To, DC\n(MW), As per SLDC Scada in MW, Diff (MW), Mus
  - Non compliance: MW as per ramp, Diff , MU

### Data Rows:
- **Row 5 onwards:** Data rows
- Column B: Date (only on first row of each period)
- Column C: From (time)
- Column D: To (time)
- Column E: DC (MW) - values from HNPCL DC file
- Column F: As per SLDC Scada in MW - values from Daily BD LR files (yellow highlighted)
- Column G: Diff (MW) - Formula: =E-F
- Column H: Mus - Formula: =G/4000
- Column J: MW as per ramp - Formulas: =E-A (first), =MAX(J{prev}-40,270) (subsequent)
- Column K: Diff - Formula: =F-J
- Column L: MU - Formula: =IF(K/4000>0,K/4000,0)

## Key Differences

| Aspect | Reference File | Output File |
|--------|---------------|-------------|
| **Data Source** | Manually entered | Fetched from input files |
| **DC Values** | Manually entered | From HNPCL revised DC file |
| **SCADA Values** | Manually entered | From Daily BD LR files |
| **Format/Structure** | Template | **Must match exactly** |
| **Formulas** | Same | Same |
| **Formatting** | Same | Same |

## Summary

The **OUTPUT FILE** should have:
- ✅ **Same structure** as reference file (rows, columns, merged cells)
- ✅ **Same formatting** as reference file (fonts, colors, borders)
- ✅ **Same formulas** as reference file
- ✅ **Same header layout** as reference file
- ✅ **Fresh data** from input files (DC file, BD LR files)
- ✅ **Same instruction periods** extracted from reference file

The reference file is the **template/format guide** - the output should look identical in structure and format, but contain fresh data automatically fetched from the input files.
