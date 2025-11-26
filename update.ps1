# pull recent changes
git pull

# activate virtual environment and upgrade pip
if (-not (Test-Path .venv)) { python -m venv .venv }
./.venv/Scripts/Activate.ps1
python -m pip install --upgrade pip

# install dependencies
$Dependencies = @(
    "openpyxl>=3.1.5,<4.0.0",
    "pandas>=2.3.3,<3.0.0",
    "pypdf>=6.4.0,<7.0.0",
    "tqdm>=4.67.1,<5.0.0"
)
foreach ($item in $Dependencies) {
    python -m pip install $item
}

# deactivate virtual environment
deactivate
