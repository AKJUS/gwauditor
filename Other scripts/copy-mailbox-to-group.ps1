[console]::OutputEncoding = [System.Text.Encoding]::UTF8
cls
Write-Host "### SCRIPT TO ARCHIVE GOOGLE WORKSPACE MAILBOX TO A GROUP, PLEASE FOLLOW INSTRUCTIONS ###"
Write-Host
function pause{ $null = Read-Host 'Press ENTER key to close the window' }

# set variables
$GAMpath = "C:\GAM7"
$gamsettings = "$env:USERPROFILE\.gam"

# collect project folders on $gamsettings
$directories = Get-ChildItem -Path $gamsettings -Directory -Exclude "gamcache" | Select-Object -ExpandProperty Name
$datetime = get-date -f yyyy-MM-dd-HH-mm

# user should choose the project available on GAM 
Write-Host "Projects available:" $directories
Write-Host

While ( ($Null -eq $clientName) -or ($clientName -eq '') ) {
    $clientName = Read-Host -Prompt "Please enter project shortname"
}

cls

cd $GAMpath

gam select $clientName save
Write-Host

function Check-AdminAddress {
    param (
        [string]$adminAddress
    )

    # Run GAM command to check if the admin address exists
    $output = gam info user $adminAddress 2>&1

    # Check the output for errors
    if ($output -match "Does not exist" -or $output -match "Show Info Failed" -or $output -match "ERROR" -or $output -match "Super Admin: False") {
        return $false
    } else {
        return $true
    }
}

while ($true) {
    # Prompt for the admin account
    $adminAddress = Read-Host "Please enter the admin account"

    # Check if the input is empty
    if ([string]::IsNullOrWhiteSpace($adminAddress)) {
        continue
    }

    # Check if the admin account exists
    if (Check-AdminAddress -adminAddress $adminAddress) {
        break
    } else {
        Write-Host "The admin account $adminAddress does not exist, or we have an ERROR. Please check credentials and try again."
    }
}


function Check-AdminAuth {
    param (
        [string]$adminAddress
    )

    # Run GAM command to check if the admin account have auth
    $output = gam user $adminAddress check serviceaccount 2>&1

    # Check the output for errors
    if ($output -match "Some scopes failed") {
        return $false
    } else {
        return $true
    }
}

while ($true) {
    # Check if the admin address exists
    if (Check-AdminAuth -adminAddress $adminAddress) {
        break
    } else {
        Write-Host "The admin account $adminAddress do not have proper authorization, we will run again the command to let you authorize it:"
		gam user $adminAddress check serviceaccount
    }
}

# Function to check if a mailbox address exists
function Check-EmailAddress {
    param (
        [string]$sourceAddress
    )

    # Run GAM command to check if the mailbox address exists
    $output = gam info user $sourceAddress 2>&1

    # Check the output for errors
    if ($output -match "Does not exist" -or $output -match "Show Info Failed" -or $output -match "ERROR") {
        return $false
    } else {
        return $true
    }
}

while ($true) {
    # Prompt for the mailbox address
    $sourceAddress = Read-Host "Please enter the mailbox address"

    # Check if the input is empty
    if ([string]::IsNullOrWhiteSpace($sourceAddress)) {
        continue
    }

    # Check if the mailbox address exists
    if (Check-EmailAddress -sourceAddress $sourceAddress) {
        break
    } else {
        Write-Host "The mailbox $sourceAddress does not exist, it's a group, or we have an ERROR. Please check credentials and try again."
    }
}

# Function to check if a group exists
function Check-GroupAddress {
    param (
        [string]$targetAddress
    )

    # Run GAM command to check if the group address exists
    $output = gam info group $targetAddress 2>&1

    # Check the output for errors
    if ($output -match "Does not exist" -or $output -match "Show Info Failed" -or $output -match "ERROR") {
        return $false
    } else {
        return $true
    }
}

while ($true) {
    # Prompt for the group address
    $targetAddress = Read-Host "Please enter the group address"

    # Check if the input is empty
    if ([string]::IsNullOrWhiteSpace($targetAddress)) {
        continue
    }

    # Check if the group address exists
    if (Check-GroupAddress -targetAddress $targetAddress) {
        break
    } else {
        Write-Host "The group $targetAddress does not exist, it's a user mailbox, or we have an ERROR. Please check credentials and try again."
    }
}

cd $GAMpath

Write-Host
Write-Host Archiving mailbox to group using command: "gam user $sourceAddress archive messages $targetAddress max_to_archive 0 doit"
gam user $sourceAddress archive messages $targetAddress max_to_archive 0 doit

Write-Host
Write-Host "### SCRIPT TO ARCHIVE GOOGLE WORKSPACE MAILBOX TO A GROUP COMPLETED ###"

$currentdate = Get-Date
$culture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")
$currentdate = $currentdate.ToString("dddd, dd MMMM yyyy HH:mm:ss", $culture)

# show info after running script
Write-Host
Write-Host Project used by GAM: $clientName
Write-Host Actual date and time: $currentdate
Write-Host

pause
exit
