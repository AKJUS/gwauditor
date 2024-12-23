gam user fred@domain.com add delegate bob@domain.com

gam fred@domain.com del delegate bob@domain.com


[console]::OutputEncoding = [System.Text.Encoding]::UTF8
cls
Write-Host "### SCRIPT TO COPY GOOGLE WORKSPACE MAILBOX TO A GROUP, PLEASE FOLLOW INSTRUCTIONS ###"
Write-Host

function pause{ $null = Read-Host 'Press ENTER key to close the window' }

if (Get-Module -ListAvailable -Name ImportExcel) {
    Write-Host "Module ImportExcel found, no additional installation required"
	Write-Host
} 
else {
    Write-Host "Module ImportExcel do not exist, please run 'Install-Module -Name ImportExcel' as administrator"
	pause
	exit
}

# set variables
$GAMpath = "C:\GAM7"
$gamsettings = "$env:USERPROFILE\.gam"
$destinationpath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path

# collect project folders on $gamsettings
$directories = Get-ChildItem -Path $gamsettings -Directory -Exclude "gamcache" | Select-Object -ExpandProperty Name
$datetime = get-date -f yyyy-MM-dd-HH-mm

# user should choose the project available on GAM 
Write-Host "Projects available:" $directories
Write-Host

While ( ($Null -eq $clientName) -or ($clientName -eq '') ) {
    $clientName = Read-Host -Prompt "Please enter project shortname"
}

# user should send the mailbox address 
Write-Host "Projects available:" $directories
Write-Host

While ( ($Null -eq $clientName) -or ($clientName -eq '') ) {
    $clientName = Read-Host -Prompt "Please enter the source mailbox address"
}

# user should send the group address 
Write-Host "Projects available:" $directories
Write-Host

While ( ($Null -eq $clientName) -or ($clientName -eq '') ) {
    $clientName = Read-Host -Prompt "Please enter the target group address"
}

cls

cd $GAMpath

gam select $clientName save

Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser

gam redirect csv ./"users-report-$datetime.csv" print users fields primaryEmail creationTime id isAdmin isDelegatedAdmin isEnforcedIn2Sv isEnrolledIn2Sv lastLoginTime name suspended


cls
Write-Host "### SCRIPT TO COLLECT GOOGLE WORKSPACE DATA COMPLETED ###"

# show info after collect report
Write-Host
Write-Host Project used by GAM: $clientName
Write-Host Actual date and time: $currentdate
Write-Host MD5 hash of [audit-$clientName-$datetime.xlsx] file: $hash

pause
exit
