"""Streamlit Cloud entry point that explicitly executes the page on every rerun."""

from pathlib import Path
import runpy


runpy.run_path(
    str(Path(__file__).with_name("line_control_app.py")),
    run_name="__main__",
)
