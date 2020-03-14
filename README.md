# timesheet_update.py
Python project created to update a timesheet, check if Microsoft outlook is open, and then open it if currently closed. This was used while consulting to easily create, update, and send timesheet to parent company.

Program asks for relevant field updates (e.g. bank holiday, vacation day, etc), will then add email contents including attaching the timesheet file, adding TO and CC fields, and then add prewritten text to the email body, and finally attaches the excel spreadsheet while saving a new copy in a designated folder.

The code may be useful to others if they need a template which includes relevant code to update a certain cell in a spreadsheet as well as attach the information and add relevant fields to an email in outlook.

Only tested on windows machine (as only needed on windows machine).