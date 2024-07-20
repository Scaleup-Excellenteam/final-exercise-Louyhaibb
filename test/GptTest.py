import sys
import json
from pathlib import Path
import subprocess

# Configure file locations
ROOT_DIR = Path(__file__).parent.parent
MAIN_SCRIPT = ROOT_DIR / 'main.py'
SAMPLE_PPT = ROOT_DIR / 'logging_debugging_getting_into_a_large_codebase.pptx'

def verify_script_execution_and_output():
    # Execute the main script and capture its output
    process = subprocess.run([sys.executable, str(MAIN_SCRIPT)], 
                             capture_output=True, text=True)

    # Ensure the script executed successfully
    if process.returncode != 0:
        raise AssertionError(f"Script execution failed.\nStdout: {process.stdout}\nStderr: {process.stderr}")

    # Verify the existence of the output JSON file
    expected_output = SAMPLE_PPT.with_suffix('.json')
    if not expected_output.exists():
        raise FileNotFoundError(f"Expected output file not found: {expected_output}")

if __name__ == "__main__":
    verify_script_execution_and_output()