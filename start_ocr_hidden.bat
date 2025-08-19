@echo off
:: Start OCR Service in hidden mode (no CMD window)

:: Run the VBS script which starts the OCR service hidden
wscript.exe "%~dp0run_ocr_hidden.vbs"

:: Exit immediately
exit