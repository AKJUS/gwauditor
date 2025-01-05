# Google Workspace Auditor script

<a target="_blank" href="https://github.com/ivancarlosti/gwauditor"><img src="https://img.shields.io/github/stars/ivancarlosti/gwauditor?style=flat" /></a>
<a target="_blank" href="https://github.com/ivancarlosti/gwauditor"><img src="https://img.shields.io/github/last-commit/ivancarlosti/gwauditor" /></a>
[![GitHub Sponsors](https://img.shields.io/github/sponsors/ivancarlosti?label=GitHub%20Sponsors)](https://github.com/sponsors/ivancarlosti)

This script collects users, groups, mailboxes delegation, Shared Drives, YouTube accounts, Analytics accounts, policies of a [Google Workspace](https://workspace.google.com/) environment on .xlsx file for audit and review purposes, the file is archived in a .zip file including a screenshot with hash MD5 of the .xlsx file and the script executed. Note that it's prepared to run on [GAM](https://github.com/GAM-team/GAM/) configured for multiple projects, change accordly if needed. This project also offer extra features:
- Archive mailbox messages to group
- List, add or remove mailbox delegation

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

You can find scripts related to mailbox delegation and mailbox archive to group in `Other scripts` folder

## Instructions

* Save all .ps1 files locally and update variables if needed
* Change variables of `mainscript.ps1` if needed and run it on PowerShell (right-click on file > Run with PowerShell)
* Follow instructions selecting project name, option 1 to generate audit report and collect .zip file on `$destinationpath`

## Screenshots
*parts ommited on screenshots are related to project/profile name

![image](https://github.com/user-attachments/assets/489b37e0-c042-4df2-9ac9-4f5871a8d95f)
*Script startup*

![image](https://github.com/user-attachments/assets/08cb9aab-cb7a-4444-bf1e-f32a518ba190)
*Script completed*

![image](https://github.com/user-attachments/assets/6d642c0c-dfd8-4810-b674-6280b81857ce)
*.zip file content*

## Requirements

* Windows 10+ or Windows Server 2019+
* [GAM v7+](https://github.com/GAM-team/GAM/) using multiproject setup 
* PowerShell
* Module `ImportExcel` on PowerShell (not required to run extra features)
