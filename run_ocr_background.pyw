#!/usr/bin/env pythonw
# -*- coding: utf-8 -*-
"""
Run OCR service in background without console window
Use .pyw extension to run without console on Windows
"""

import sys
import os
import subprocess
from pathlib import Path

def run_ocr_service():
    """Run the OCR service in background"""
    script_dir = Path(__file__).parent
    ocr_script = script_dir / "background_ocr_service_refactored.py"
    
    # Check if the script exists
    if not ocr_script.exists():
        # Try complete_ocr_service.py as fallback
        ocr_script = script_dir / "complete_ocr_service.py"
    
    if ocr_script.exists():
        # Run the script using pythonw (no console window)
        subprocess.Popen([sys.executable, str(ocr_script)], 
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        stdin=subprocess.DEVNULL,
                        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0)

if __name__ == "__main__":
    run_ocr_service()