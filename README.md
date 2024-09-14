# Script-for-Google-Workspace-audit

This script collects users, groups and Shared Drives on a excel file for audit and review purposes. Note that it's prepared to run on GAM configured for multiple projects, change accordly if needed.

Set variables if different of defined:
```
$gamsettings = "$env:USERPROFILE\.gam"
$GAMpath = "C:\GAM7"
$destinationpath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
```

`$gamsettings` defines the settings folder of GAM
`$GAMpath` defines of GAM application folder
`$destinationpath` defines the location were script result is saved

Run file on PowerShell (right-click on file > Run with PowerShell)
Collect .zip file on `$destinationpath` that includes the report file, a screenshot and the script itself

Requirements:
* Windows 10+ or Windows Server 2019+
* PowerShell
* Module ImportExcel on PowerShell
