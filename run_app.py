#!/usr/bin/env python3
"""
Launcher for the Back Down Calculator Streamlit app.
Used as the PyInstaller entry point so the .exe runs the app.
Works both when frozen (PyInstaller) and when run from source.
"""
import os
import sys

def _base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def main():
    base = _base_dir()
    os.chdir(base)
    if base not in sys.path:
        sys.path.insert(0, base)

    # Run Streamlit programmatically (same as: streamlit run app.py --server.headless true)
    import streamlit.web.cli as stcli
    sys.argv = [
        "streamlit", "run",
        os.path.join(base, "app.py"),
        "--server.headless", "true",
        "--server.port", "8501",
        "--browser.gatherUsageStats", "false",
    ]
    stcli.main()

if __name__ == "__main__":
    main()
