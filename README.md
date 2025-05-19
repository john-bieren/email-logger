# Email Logger

A Python script that uses .eml files to create a log (in Excel .xlsx format) of sender, recipients (including CCs), subject line, and time sent for a given set of emails. Optionally, converted PDF versions of the emails can be used to add the PDF page count of each email to the log.

## Install

Clone the repository:
```
git clone https://github.com/john-bieren/email-logger.git
```
For ease of use, desktop shortcuts to `email_logger.ps1` and `update.ps1` can be created. `email_logger.ps1` allows the script to be run without using the terminal, and includes a pause statement so that the output can be read before closing the window. `update.ps1` automates the process of keeping the repository and its dependencies up to date. For these to work, make sure that the shortcuts start in the project directory.

## Usage

To create an email log:
1. Run `main.py` or `email_logger.ps1`.
2. Enter the full path to the folder that contains the .eml files for the emails you wish to log.
3. Optionally, enter the full path to the folder that contains the .pdf versions of the emails. The corresponding .pdf and .eml versions of an email must have the same file name (apart from the extension) for the page count to be properly included in the log.
4. Enter the full path to the folder where you want the log to be saved.

The given paths do not have to be different, the script will work as expected even if some or all of the paths are the same.
