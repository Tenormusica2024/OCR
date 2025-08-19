' VBScript to run OCR service without visible CMD window
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the current directory
strCurrentDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Build the command to run
strCommand = "cmd /c """

' Check if virtual environment exists and activate it
If objFSO.FolderExists(strCurrentDir & "\ocr_env") Then
    strCommand = strCommand & "cd /d """ & strCurrentDir & """ && "
    strCommand = strCommand & "call ocr_env\Scripts\activate.bat && "
End If

' Add the Python command
strCommand = strCommand & "python """ & strCurrentDir & "\background_ocr_service_refactored.py"""""

' Run the command hidden (0 = hidden window)
WshShell.Run strCommand, 0, False

' Exit silently
WScript.Quit