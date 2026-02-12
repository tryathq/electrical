# Find Station Rows Script

## Installation

```bash
pip install openpyxl
```

## Usage

### Complete Command

```bash
python find_station_rows.py \
  --instructions-file "data/january/instructions.xlsx" \
  --station HINDUJA \
  --dc-file "data/january/HNPCL revised DC for the month January 2026 SLDC.xlsx" \
  --bd-folder "data/january/BD" \
  --scada-column "HNJA4_AG.STTN.X_BUS_GEN.MW" \
  --bd-sheet "DATA-CMD"
```

## Output

Output file is saved to `output/` folder with format: `{STATION}_{DATE}_{TIME}.xlsx`

Example: `HINDUJA_12-Feb-2026_2-39-40-PM.xlsx`
