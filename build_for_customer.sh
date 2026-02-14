#!/bin/bash
# Run this ONCE on your Mac/Linux. It creates a folder you can zip and send to the customer.
# Customer (Mac/Linux): unzip and run ./START_APP.sh (no Python install needed on their side).

set -e
cd "$(dirname "$0")"
OUT="BackDownCalculator"
VENV="$OUT/venv"

echo "Building customer package in: $OUT"
rm -rf "$OUT"
mkdir -p "$OUT"

echo "Creating virtual environment..."
python3 -m venv "$VENV"
source "$VENV/bin/activate"
pip install -r requirements.txt -q

echo "Copying app files..."
cp app.py config.py reports_store.py url_utils.py instructions_parser.py excel_builder.py find_station_rows.py requirements.txt "$OUT/"
[ -f CUSTOMER_README.txt ] && cp CUSTOMER_README.txt "$OUT/README.txt"
[ -f .streamlit/config.toml ] && mkdir -p "$OUT/.streamlit" && cp .streamlit/config.toml "$OUT/.streamlit/"
mkdir -p "$OUT/reports"

echo "Creating one-click launcher..."
cat > "$OUT/START_APP.sh" << 'LAUNCHER'
#!/bin/bash
cd "$(dirname "$0")"
echo "Starting app... browser will open shortly."
./venv/bin/streamlit run app.py
echo "App stopped. You can close this window."
LAUNCHER
chmod +x "$OUT/START_APP.sh"

echo ""
echo "Done. Folder: $OUT"
echo "Next: zip it (e.g. zip -r BackDownCalculator.zip $OUT) and send to customer."
echo "Customer: unzip, then run ./START_APP.sh. No other steps."
