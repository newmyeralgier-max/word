import latex2mathml.converter
from lxml import etree
import os
import sys

# Add the script's directory to sys.path so we can import word_utils
sys.path.insert(0, r"d:\1. Project\Word\WORD\execution")
import word_utils as wu

test_str = r"B_{\text{л}}"
print(f"Testing LaTeX: {test_str}")

try:
    omml = wu.latex_to_omml(test_str)
    if omml is not None:
        print("Success! OMML generated.")
        print(etree.tostring(omml, encoding='unicode'))
    else:
        print("Failure: latex_to_omml returned None")
except Exception as e:
    print(f"Error: {e}")
