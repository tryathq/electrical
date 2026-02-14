# PyInstaller spec: build .exe + folder for customer. Zip the folder; customer unzips and double-clicks the .exe.
# Run: pyinstaller BackDownCalculator.spec   (or use build_exe.bat which also zips)

import os
from PyInstaller.utils.hooks import collect_all

SPEC_DIR = os.path.dirname(os.path.abspath(SPEC))

# App files to ship next to the exe (so streamlit can run app.py from cwd)
app_datas = [
    (os.path.join(SPEC_DIR, "app.py"), "."),
    (os.path.join(SPEC_DIR, "config.py"), "."),
    (os.path.join(SPEC_DIR, "reports_store.py"), "."),
    (os.path.join(SPEC_DIR, "url_utils.py"), "."),
    (os.path.join(SPEC_DIR, "instructions_parser.py"), "."),
    (os.path.join(SPEC_DIR, "excel_builder.py"), "."),
    (os.path.join(SPEC_DIR, "find_station_rows.py"), "."),
]
streamlit_config = os.path.join(SPEC_DIR, ".streamlit", "config.toml")
if os.path.isfile(streamlit_config):
    app_datas.append((streamlit_config, ".streamlit"))
readme = os.path.join(SPEC_DIR, "CUSTOMER_README.txt")
if os.path.isfile(readme):
    app_datas.append((readme, "."))

# Streamlit and its dependencies (data + binaries + hidden imports)
streamlit_datas, streamlit_binaries, streamlit_hidden = collect_all("streamlit")
try:
    altair_datas, altair_binaries, altair_hidden = collect_all("altair")
except Exception:
    altair_datas = altair_binaries = altair_hidden = []

a = Analysis(
    [os.path.join(SPEC_DIR, "run_app.py")],
    pathex=[SPEC_DIR],
    datas=app_datas + streamlit_datas + altair_datas,
    binaries=streamlit_binaries + altair_binaries,
    hiddenimports=[
        "streamlit",
        "streamlit.web.cli",
        "pandas",
        "openpyxl",
        "st_aggrid",
    ] + streamlit_hidden + altair_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="BackDownCalculator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,  # Keep console so user sees "You can now view your Streamlit app in your browser"
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="BackDownCalculator",
)
