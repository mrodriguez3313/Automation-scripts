# README

This script is used for automating the process of disputing invoices. 

See the [flowchart](flowchart.pptx) to understand the logic the script takes to complete 99.99% of claims provided in an excel sheet.

To learn how to run, skip down to the [how_to_run](#how-to-run) section

## Results:
* Average: 15 seconds per claim.
* Fastest: 12 seconds per claim.
* Slowest: Dependent on how fast/slow the site loads. But roughly 17-23 seconds per claim.

Can complete a month of claims (~11K claims) in a week (7 days).

## What to know
1. You can not have the same workbook open, at the time the program finishes. YOU MAY LOSE ALL YOUR DATA.
2. The browser is opened in 'undetectable' (UC) mode, so theoretically the receiving servers think the script is just a fast human. __*Running this script multiple times under the same login, may raise some flags.*__
3. Back up selectors have not been implemented as of yet. Changes to the site may break the script. View logs to traceback the issues.
4. Currently, there may be bugs that are caused by not accounting for specific website behaviors/errors which may cause the script to skip a claim.
5. The website the script is run on, has many errors. An unrecoverable one is "Authentication Error", which forces you to log out.

## Before you run
1. The Excel Workbook name is arbitrary. But, the Excel sheetname MUST start with "Check_" followed by 3-10 digits. EX: Check_00749264.
2. A user has 2 minutes to log in at the start of the script and the login info is not saved by the program or the browser.
3. Even if an error occurs, the script should recover all on its own.

# How to run
## Requirements
* Python 3.11+
* SeleniumBase 4.22+
* Selenium 4.x+
* easygui 0.98+
* pandas 2.1+

# Run
Execute script with python

### Windows:

If python is in the Path:

python3 <path/to/script.py>

`python3 APDP-File-disputes.py`

Otherwise:

Path/to/python3.exe Path/to/script.py

`/Downloads/python3.exe APDP-File-disputes.py`

### MAC/Linux

`python3 APDP-File-disputes.py`
