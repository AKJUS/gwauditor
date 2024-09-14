Write-Host "### SCRIPT TO COLLECT GOOGLE WORKSPACE DATA , PLEASE FOLLOW INSTRUCTIONS ###"
Write-Host

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
$gamsettings = "$env:USERPROFILE\.gam"
$GAMpath = "C:\GAM7"
$directories = Get-ChildItem -Path $gamsettings -Directory -Exclude "gamcache" | Select-Object -ExpandProperty Name

# user should choose the project available on GAM 
Write-Host "Please maximize window to 1:1 (1/4 if scaled 200%), a screenshot will be generated on end of it"
Write-Host "Projects available:" $directories
Write-Host
$clientName = Read-Host "Please enter project shortname"
cls

# delete files used on this project on $GAMpath
del $GAMpath\*.csv
del $GAMpath\*.xlsx
del $GAMpath\*.bmp
del $GAMpath\*.ps1
del $GAMpath\*.zip

# copy script to $GAMpath
Copy-Item $MyInvocation.MyCommand.Name $GAMpath

cd $GAMpath

gam select $clientName save

Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
Import-Module -Name ImportExcel

gam redirect csv ./"users-report-$(get-date -f yyyy-MM-dd).csv" print users fields primaryEmail creationTime id isAdmin isDelegatedAdmin isEnforcedIn2Sv isEnrolledIn2Sv lastLoginTime name suspended
gam redirect csv ./"groups-report-$(get-date -f yyyy-MM-dd).csv" print groups fields email id name adminCreated members manager owners
gam redirect csv ./"teamdriveacls-report-$(get-date -f yyyy-MM-dd).csv" print teamdriveacls oneitemperrow

Import-Csv .\users-report-$(get-date -f yyyy-MM-dd).csv -Delimiter ',' | Export-Excel -Path .\audit-$clientName-$(get-date -f yyyy-MM-dd).xlsx -WorksheetName users
Import-Csv .\groups-report-$(get-date -f yyyy-MM-dd).csv -Delimiter ',' | Export-Excel -Path .\audit-$clientName-$(get-date -f yyyy-MM-dd).xlsx -WorksheetName groups
Import-Csv .\teamdriveacls-report-$(get-date -f yyyy-MM-dd).csv -Delimiter ',' | Export-Excel -Path .\audit-$clientName-$(get-date -f yyyy-MM-dd).xlsx -WorksheetName teamdriveacls

Write-Host
certutil -hashfile "audit-$clientName-$(get-date -f yyyy-MM-dd).xlsx" MD5; get-date
Write-Host

# print screen program
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$Screen = [System.Windows.Forms.SystemInformation]::WorkingArea
$Width  = $Screen.Width
$Height = $Screen.Height
$Left   = $Screen.Left
$Top    = $Screen.Top
$bitmap  = New-Object System.Drawing.Bitmap $Width, $Height
$graphic = [System.Drawing.Graphics]::FromImage($bitmap)
$graphic.CopyFromScreen($Left, $Top, 0, 0, $bitmap.Size)

# save print screen on $GAMpath
$bitmap.Save("$GAMpath\audit-$clientName-$(get-date -f yyyy-MM-dd).bmp")

# add files to .zip file on $GAMpath
Compress-Archive "*.xlsx" -DestinationPath audit-$clientName-$(get-date -f yyyy-MM-dd).zip
Compress-Archive -Path "*.bmp" -Update -DestinationPath "audit-$clientName-$(get-date -f yyyy-MM-dd).zip"
Compress-Archive -Path "*.ps1" -Update -DestinationPath "audit-$clientName-$(get-date -f yyyy-MM-dd).zip"

Move-Item audit-$clientName-$(get-date -f yyyy-MM-dd).zip (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
Write-Host "Audit file"audit-$clientName-$(get-date -f yyyy-MM-dd).zip" saved on"(New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
Write-Host

del $GAMpath\*.csv
del $GAMpath\*.xlsx
del $GAMpath\*.bmp
del $GAMpath\*.ps1
del $GAMpath\*.zip

pause
