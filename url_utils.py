"""URL query parameter helpers for navigation."""

import streamlit as st


def url_reports_list() -> None:
    """Set URL to reports list (?view=reports)."""
    if getattr(st, "query_params", None) is not None and hasattr(st.query_params, "from_dict"):
        st.query_params.from_dict({"view": "reports"})


def url_report_file(filename: str) -> None:
    """Set URL to single report (?view=report&file=...)."""
    if getattr(st, "query_params", None) is not None and hasattr(st.query_params, "from_dict"):
        st.query_params.from_dict({"view": "report", "file": filename})


def url_main() -> None:
    """Set URL to main page (clear view/file params)."""
    if getattr(st, "query_params", None) is not None:
        if hasattr(st.query_params, "clear"):
            st.query_params.clear()
        elif hasattr(st.query_params, "from_dict"):
            st.query_params.from_dict({})
