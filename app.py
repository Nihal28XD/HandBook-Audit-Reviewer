"""
RBI Handbook Enterprise Audit & Verification Suite - Main Deployment Entry Point.

This module serves as the primary entry point (`app.py`) for Streamlit and Render.
It securely locates and executes the production application defined in
`pdf_analyser_Final.py` without code duplication or circular dependencies.
"""

from pathlib import Path
import runpy
import sys

# Resolve absolute path to project root and target application script
BASE_DIR = Path(__file__).resolve().parent
TARGET_SCRIPT = BASE_DIR / "pdf_analyser_Final.py"

if not TARGET_SCRIPT.is_file():
    raise FileNotFoundError(
        f"Critical startup error: Target script '{TARGET_SCRIPT.name}' not found at {TARGET_SCRIPT}"
    )

# Ensure project root is present in sys.path for cross-platform imports
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

# Execute the application script in-place with top-level __main__ context
runpy.run_path(str(TARGET_SCRIPT), run_name="__main__")
