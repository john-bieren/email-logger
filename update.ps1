# pull recent changes
git pull

# activate virtual environment and upgrade pip
if (-not (Test-Path .venv)) { python -m venv .venv }
./.venv/Scripts/Activate.ps1
python -m pip install --upgrade pip

# install dependencies
python -m pip install -r requirements.txt

# deactivate virtual environment
deactivate
