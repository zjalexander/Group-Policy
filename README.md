This contains 2 Powershell Scripts:

**compare-admx.ps1**

Use: 
``` .\compare-admx.ps1
compare-admx "C:\Program Files (x86)\Microsoft Group Policy\Windows Server 2012\PolicyDefinitions" "C:\Program Files (x86)\Microsoft Group Policy\Windows 8.1-Windows Server 2012 R2\PolicyDefinitions" | out-file comparison.txt
```
If you have installed the two ADMX packages in ..\ADMX Packages
this generates a file called comparison.txt. This file lists all parts of the ADMX which were only present in the old version of Windows.



**get-admxCSV.ps1**

Use:

` .\get-admxCSV.ps1 `

This expects two things:
1) ADMX files in the current directory
2) ADML files in the directory .\en-us

This will scan through all ADMX files and document the settings contained within them. 
It may throw up errors and warnings. These are due to miscreated ADMX files and bugs should be filed on the owner of the ADMX files.

This script is semi-portable. Many definitions are included in i.e. Windows.ADMX, so while this script can be used to reproduce an error, it
needs to include the offending ADMX, the corresponding ADML, Windows.ADMX/ADML, as well as any other ADMX that define namespaces used by the ADML
 (like WindowsServer.ADMX\L, or errorreporting.admx\l)

Open the CSV, use Excel to translate text into columns, delimited by ^
