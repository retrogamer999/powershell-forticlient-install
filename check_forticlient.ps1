################################################################################################
################################################################################################
################################################################################################
### Created by:     Usman Hussain
### Date:           14/01/2023
### Last Changed:   15/11/2024
###
### Change Description:
### Modified for East End Homes
################################################################################################
################################################################################################
################################################################################################

# Variables
$sharePath = "\\EEH-AFS-1\IT\Software\FortiClient VPN 7.0.3"
$localInstallPath = "C:\Temp\FortiClient7-0-13"
$msiFile = "FortiClient_7.0.13_x64.msi"
$mstFile = "FortiClient.mst"

# Function to install FortiClient
function Install-FortiClient {
    Write-Output "Preparing to install FortiClient version 7.0.13..."

    # Create local folder if it doesn't exist
    if (-not (Test-Path -Path $localInstallPath)) {
        New-Item -ItemType Directory -Path $localInstallPath -Force | Out-Null
    }

    # Copy files from the network share
    Write-Output "Copying installation files from $sharePath to $localInstallPath..."
    Copy-Item -Path "$sharePath\*" -Destination $localInstallPath -Recurse -Force

    # Run the installer
    Write-Output "Running the MSI installer..."
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$localInstallPath\$msiFile`" TRANSFORMS=`"$localInstallPath\$mstFile`" /quiet /norestart" -Wait

    Write-Output "FortiClient version 7.0.13 has been installed."
}

# Function to uninstall FortiClient
function Uninstall-FortiClient {
    param (
        [string]$uninstallString
    )

    Write-Host "Uninstalling FortiClient..."  # Added this line for visual feedback during uninstallation
    if ($uninstallString) {
        # Adjust the UninstallString if necessary
        $uninstallString = $uninstallString -replace "/I", "/X"

        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString /quiet /norestart" -Wait -NoNewWindow
        Write-Output "FortiClient has been uninstalled."
    } else {
        Write-Output "UninstallString not found. Unable to uninstall FortiClient."
    }
}

# Check if FortiClient is installed
$fortiClient = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" |
               Where-Object { $_.DisplayName -like "*FortiClient*" }

if ($fortiClient) {
    $installedVersion = $fortiClient.DisplayVersion
    Write-Output "FortiClient version installed: $installedVersion"

    # Check if the version does not contain 7.0.13
    if ($installedVersion -notcontains "7.0.13") {
        Write-Output "FortiClient version does not match 7.0.13. Proceeding with uninstallation and installation of 7.0.13."

        # Uninstall the existing version and reboot
        Uninstall-FortiClient -uninstallString $fortiClient.UninstallString
        
        # Trigger reboot to finalize uninstall before proceeding with installation
        Write-Output "Rebooting the system to complete uninstall..."
        Restart-Computer -Force -Wait
    } else {
        Write-Output "FortiClient is already on version 7.0.13. No action required."
    }
} else {
    Write-Output "FortiClient is not installed. Proceeding with installation..."
}

# Install the correct version after reboot
Install-FortiClient
