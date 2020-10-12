Function Get-RandomAlphanumericString {
	[CmdletBinding()]
	Param (
        [int] $length = 8
	)

	Begin{ }

	Process{
        Write-Output ( -join ((0x30..0x39) + ( 0x41..0x5A) | Get-Random -Count $length  | % {[char]$_}) )
	}	
}

$host.ui.RawUI.WindowTitle = "PCN OEM Finalizing script by Kaalus"

# Import Windows update module
Import-Module PSWindowsUpdate

# Get and install any Windows update / driver
Get-WindowsUpdate
Install-WindowsUpdate -AcceptAll -IgnoreReboot

# Change computername
$pcname = 'PCN-' + (Get-RandomAlphanumericString -length 10)
Rename-Computer -NewName $pcname

# Set PC model in system settings
$computerSystem = Get-CimInstance CIM_ComputerSystem
$computerModel = $computerSystem.Model
New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name "Model" -Value $computerModel -PropertyType String -Force | Out-Null

# Activate Windows
# This script activates Windows either by using the key in the BIOS / SLIC 3.0 or from a manual entered key.
# Regular expression to validate the keys
$keyRegExp = '^([A-Z1-9]{5})-([A-Z1-9]{5})-([A-Z1-9]{5})-([A-Z1-9]{5})-([A-Z1-9]{5})$'
# Obtain the SLIC 3.0 key through WMI
$key = (Get-WmiObject -Class SoftwareLicensingService).OA3xOriginalProductKey

if(!$key) {
    Write-Host 'No OEM key from BIOS detected, please enter it manually...'
    $key = Read-Host 'Enter Windows License key'
    Write-Host $tempkey
} else {
    Write-Host 'Found OEM license key'
}
# Re-Check if there is a license key
if ($key -match $keyRegExp) {
    Write-Host 'Using key' $key 'to activate Windows'
    Invoke-Expression "cscript /b C:\windows\system32\slmgr.vbs /upk"
    Invoke-Expression "cscript /b C:\windows\system32\slmgr.vbs /ipk $key"
    Invoke-Expression "cscript /b C:\windows\system32\slmgr.vbs /ato"
} else {
    Write-Warning 'No or invalid license key detected, skipping activation'
}

Restart-Computer