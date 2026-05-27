Set-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (Test-Path ".\.venv\Scripts\python.exe") {
    .\.venv\Scripts\python.exe main.py
} else {
    python main.py
}
