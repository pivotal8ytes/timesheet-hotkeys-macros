# timesheet-hotkeys-macros

This is a little set of 3 macro VB scripts I created in 2018 and have used daily since. It allows me to keep a daily timesheet of billable work organized via excel, including:

Setup: create a spreadsheet in Excel (xlsb file) and add the vbs script into the macro editor. It will include the following 3 functions:

1) CTRL + D: Add dash to beginning of current cell, with just CTRL + D. Works when multiple cells are selected also, prepends dashes to all highlighted cells.
2) CTRL + Q: Archive the currently highlighted cell to a new worksheet (it will get added at the top of that sheet, and keep all existing archived time entries there as well)
3) The third function deletes the time entry from the current worksheet after it's been archived via a CTRL + Q action
