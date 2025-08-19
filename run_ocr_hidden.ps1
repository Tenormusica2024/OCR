# PowerShell script to run OCR service without visible window
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonScript = Join-Path $scriptPath "background_ocr_service_refactored.py"
$venvPath = Join-Path $scriptPath "ocr_env"

# Build arguments for the hidden process
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = "cmd.exe"
$psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
$psi.CreateNoWindow = $true
$psi.UseShellExecute = $false

# Check if virtual environment exists
if (Test-Path "$venvPath\Scripts\activate.bat") {
    $psi.Arguments = "/c `"cd /d `"$scriptPath`" && call `"$venvPath\Scripts\activate.bat`" && python `"$pythonScript`"`""
} else {
    $psi.Arguments = "/c `"python `"$pythonScript`"`""
}

# Start the process
$process = [System.Diagnostics.Process]::Start($psi)

Write-Host "OCR service started in hidden mode (Process ID: $($process.Id))"