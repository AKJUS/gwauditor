# set variables
$GAMpath = "C:\GAM7"
$gamsettings = "$env:USERPROFILE\.gam"
$destinationpath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path

# collect project folders on $gamsettings
$directories = Get-ChildItem -Path $gamsettings -Directory -Exclude "gamcache" | Select-Object -ExpandProperty Name
$datetime = get-date -f yyyy-MM-dd-HH-mm

function Show-Menu {
    cls
    Write-Host "GAM project selected: $clientName"
    Write-Host ""
    Write-Host "Please choose an option:`n1. Generate audit report`n2. Archive mailbox messages to group`n3. List, add or remove mailbox delegation`n4. Change GAM project`n5. Exit script"
    return (Read-Host -Prompt "Enter your choice")
}

function Select-GAMProject {
    cls
    Write-Host "Projects available:" $directories
    Write-Host

    $selectedProject = $null
    while (($Null -eq $selectedProject) -or ($selectedProject -eq '') -or ((& "$GAMpath\gam.exe" select $selectedProject 2>&1) -match "ERROR")) {
        $selectedProject = Read-Host -Prompt "Please enter project shortname"
        if ((& "$GAMpath\gam.exe" select $selectedProject 2>&1) -match "ERROR") {
            Write-Host "Invalid project shortname. Please try again."
        }
    }

    Write-Host "GAM project selected: $selectedProject"
    & "$GAMpath\gam.exe" select $selectedProject save
    $global:clientName = $selectedProject  # Ensure it's set globally
    Write-Host "DEBUG: clientName is set to $clientName"
}

while ($true) {
    if (-not $clientName) {
        Select-GAMProject
        Write-Host "DEBUG: After Select-GAMProject, clientName is $clientName"
    }

    $option = Show-Menu

    try {
        switch ($option) {
            '1' {
                # Call Google Workspace Auditor script to generate audit report
                & ".\_script_GenerateAuditReport.ps1"
            }
            '2' {
                # Call script to archive mailbox messages to group
                & ".\_script_ArchiveMailboxMessages.ps1"
            }
            '3' {
                # Call script to list, add, or remove mailbox delegation
                & ".\_script_MailboxDelegation.ps1"
            }
            '4' {
                # Change GAM project
                Select-GAMProject
                Write-Host "DEBUG: After Select-GAMProject, clientName is $clientName"
            }
            '5' {
                Write-Output "Exiting script."
                break
            }
            default {
                Write-Output "Invalid option selected."
            }
        }
    }
    catch {
        Write-Host "An error occurred: $_"
    }
}

# Prevent the script from closing immediately if the exit option is chosen
if ($option -ne '5') {
    Read-Host -Prompt "Press Enter to exit"
}
