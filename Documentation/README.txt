ADMX Settings Spreadsheet Template
This is the spreadsheet used to generate the final document. This contains a formula for 
determining what settings are new from the previous version of windows.
It is a workbook with three tabs. win8_ADMXSettingsData, Blue_ADMXSettingsData,
Result.

Win8 contains the generated spreadsheet (See: get-ADMXcsv.ps1) for the previous version of Windows.
Blue contains the generated spreadsheet for the current version of Windows
Results contains the generated spreadsheet for the current version of Windows, plus column H which compares values based on Registry Keys.
If a registry key is present in Win8_ and Blue_, the value in H will be "FALSE"
If a registry key is present in Blue_ but NOT Win8_, the value in H will be True

When the next version of Windows is released:
the new settings should be added in a new tab
the new settings should replace the settings in the results tab (but not column H)
references to win8_ will need to be replaced with references to blue_.
References to blue_ will be changed to whatever you call the new tab.


