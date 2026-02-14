# Streamlit Desktop App - Find Station Rows

A user-friendly GUI version of the Find Station Rows command-line tool.

## Installation

```bash
pip install -r requirements.txt
```

## Running the App

### Option 1: Using the shell script
```bash
./run_app.sh
```

### Option 2: Direct Streamlit command
```bash
streamlit run app.py
```

### Option 3: With custom port
```bash
streamlit run app.py --server.port 8501
```

## Usage

1. **Start the app** - Run one of the commands above
2. **Upload Files** - Use the sidebar to:
   - Upload Instructions Excel file (required)
   - Upload DC file (optional)
   - Enter BD folder path (optional)
3. **Configure Settings**:
   - Enter Station Name (e.g., "HINDUJA")
   - Optionally specify Sheet Name, Column Name
   - Configure SCADA settings if needed
4. **Process** - Click the "🚀 Process" button
5. **Download** - Download the generated output file

## Features

- ✅ File upload interface (no need to specify file paths)
- ✅ Real-time progress tracking
- ✅ Summary statistics after processing
- ✅ Direct download of output file
- ✅ All features from command-line version

## Requirements

- Python 3.7+
- streamlit >= 1.28.0
- openpyxl >= 3.1.0
