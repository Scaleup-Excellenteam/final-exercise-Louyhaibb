from pathlib import Path
import subprocess
import pytest

# Locate relevant files
ROOT_DIR = Path(__file__).parent.parent
MAIN_SCRIPT = ROOT_DIR / 'main.py'
SAMPLE_PRESENTATION = ROOT_DIR / 'logging_debugging_getting_into_a_large_codebase.pptx'

def verify_script_execution():
    # Execute the main script and capture its output
    completed_process = subprocess.run(['python', str(MAIN_SCRIPT)], 
                                       capture_output=True, text=True)
    
    # Ensure the script executed successfully
    if completed_process.returncode != 0:
        error_message = f"Script execution failed.\nStdout: {completed_process.stdout}\nStderr: {completed_process.stderr}"
        pytest.fail(error_message)
    
    # Verify the existence of the output JSON file
    expected_output = SAMPLE_PRESENTATION.with_suffix('.json')
    assert expected_output.exists(), f"Expected output file not found: {expected_output}"

def test_script_functionality():
    verify_script_execution()