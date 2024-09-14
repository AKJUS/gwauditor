# Google Workspace auditor script

This script collects users, groups and Shared Drives of a [Google Workspace](https://workspace.google.com/) environment on .xlsx file for audit and review purposes, the file is archived in a .zip file including a screenshot with hash MD5 of the .xlsx file and the script executed. Note that it's prepared to run on [GAM](https://github.com/GAM-team/GAM/) configured for multiple projects, change accordly if needed.

Set variables if different of defined:
```
$GAMpath = "C:\GAM7"
$gamsettings = "$env:USERPROFILE\.gam"
$destinationpath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
```

`$GAMpath` defines the GAM application folder

`$gamsettings` defines the settings folder of GAM

`$destinationpath` defines the location were script result is saved

Check `testing-guideline.md` file as suggestion for testing guideline

## Instructions

* Save `audit-GAM.ps1` file locally and update variables if needed
* Run `audit-GAM.ps1` file on PowerShell (right-click on file > Run with PowerShell)
* Collect .zip file on `$destinationpath`

## Screenshots
*parts ommited on screenshots are related to project name

![image](https://github.com/user-attachments/assets/489b37e0-c042-4df2-9ac9-4f5871a8d95f)
*Script startup*

![image](https://github.com/user-attachments/assets/1fe8f7c8-b506-43a3-8d81-8e82717f0a45)
*Script completed*

![image](https://github.com/user-attachments/assets/6d642c0c-dfd8-4810-b674-6280b81857ce)
*.zip file content*

## Requirements

* Windows 10+ or Windows Server 2019+
* [GAM v7+](https://github.com/GAM-team/GAM/)
* PowerShell
* Module `ImportExcel` on PowerShell
