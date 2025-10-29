# if necessary, set up virtual environment and download dependencies
if (-not (Test-Path .venv)) {
    Write-Output "Setting up environment..."
    ./update.ps1
}
# activate virtual environment
./.venv/Scripts/Activate.ps1

# run script
python main.py
# deactivate virtual environment
deactivate
# wait for user input to end script (keeps window open if launched through desktop shortcut)
pause
