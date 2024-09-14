# Script-for-Google-Workspace-audit

This script collects users, groups and Shared Drives on a excel file for audit and review purposes. Note that it's prepared to run on GAM configured for multiple projects, change accordly if needed.

Set variables if different of defined:
```
$GAMpath = "C:\GAM7"
$gamsettings = "$env:USERPROFILE\.gam"
$destinationpath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
```

`$GAMpath` defines of GAM application folder

`$gamsettings` defines the settings folder of GAM

`$destinationpath` defines the location were script result is saved

Run file on PowerShell (right-click on file > Run with PowerShell)
Collect .zip file on `$destinationpath` that includes the report file, a screenshot and the script itself

Requirements:
* Windows 10+ or Windows Server 2019+
* [GAM v7+](https://github.com/GAM-team/GAM/)
* PowerShell
* Module ImportExcel on PowerShell
