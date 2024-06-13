import os
import subprocess

def test_script():
    demo_pptx = os.path.join(os.path.dirname(__file__), 'presentation.pptx')
    output_json = os.path.splitext(demo_pptx)[0] + '.json'
    subprocess.run(['python', 'explainer.py', demo_pptx])
    assert os.path.exists(output_json)
