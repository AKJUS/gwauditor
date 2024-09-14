# Google Workspace auditor script

This script collects users, groups and Shared Drives of a [Google Workspace](https://workspace.google.com/) environment on .xlsx file for audit and review purposes. Note that it's prepared to run on [GAM](https://github.com/GAM-team/GAM/) configured for multiple projects, change accordly if needed.

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

Run `audit-GAM.ps1` file on PowerShell (right-click on file > Run with PowerShell)
Collect .zip file on `$destinationpath` that includes the report file, a screenshot and the script itself

## Screenshots
*parts ommited on screenshots are related to project name

![image](https://github.com/user-attachments/assets/90b07e8e-5f68-4533-9439-24c9a58835fc)
*Script startup*

![image](https://github.com/user-attachments/assets/ca49633a-6822-46f0-9374-4358959fbc14)
*Script completed*

![image](https://github.com/user-attachments/assets/29d74797-acb8-4083-b87c-f331d5ef358e)
*.zip file content*

## Requirements

* Windows 10+ or Windows Server 2019+
* [GAM v7+](https://github.com/GAM-team/GAM/)
* PowerShell
* Module `ImportExcel` on PowerShell
