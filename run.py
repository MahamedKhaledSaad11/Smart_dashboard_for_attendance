import os
import sys

# Ensure clean UTF-8 environment and direct local home
script_dir = os.path.dirname(os.path.abspath(__file__))
os.environ["USERPROFILE"] = script_dir
os.environ["HOME"] = script_dir
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"

os.chdir(script_dir)

# Run streamlit directly via Python entrypoint
from streamlit.web import cli as stcli

if __name__ == "__main__":
    sys.argv = [
        "streamlit",
        "run",
        "app.py",
        "--browser.gatherUsageStats=false",
        "--server.headless=false"
    ]
    sys.exit(stcli.main())
