# Email Logger

A Python script that uses .eml files to create a log (in Excel .xlsx format) of sender, recipients (including CCs), subject line, and time sent for a given set of emails. Optionally, converted PDF versions of the emails can be used to add the PDF page count of each email to the log.

## Install

Clone the repository and install dependencies:
```
git clone https://github.com/john-bieren/email-logger.git
cd email-logger
& ./update.ps1
```
**Note**: The tags in this repository do not correspond to releases, they simply indicate breaking changes.

For ease of use, desktop shortcuts to `email_logger.ps1` and `update.ps1` can be created. `email_logger.ps1` allows the script to be run in a virtual environment without using the terminal, and includes a pause statement so that the output can be read before the window closes. `update.ps1` automates the process of setting up the virtual environment and keeping the repository and its dependencies up to date. Make sure that the shortcuts start in the project directory.

## Usage

To create an email log:
1. Run `main.py` or `email_logger.ps1`.
2. Enter the full path to the folder that contains the .eml files for the emails you wish to log.
3. Optionally, enter the full path to the folder that contains the .pdf versions of the emails. The corresponding .pdf and .eml versions of an email must have the same file name (apart from the extension) for the page count to be properly included in the log.
4. Enter the full path to the folder where you want the log to be saved.

The given paths do not have to be different, the script will work as expected even if some or all of the paths are the same.
