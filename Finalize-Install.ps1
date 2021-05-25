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

# Function borrowed from: https://www.reddit.com/r/PowerShell/comments/8l9fgf/changing_wallpaper_with_powershell/dzdspnh
function Set-Wallpaper {
    param (
        [string]$Path,
        [ValidateSet('Tile', 'Center', 'Stretch', 'Fill', 'Fit', 'Span')]
        [string]$Style = 'Fill'
    )

    begin {
        try {
            Add-Type @"
                using System;
                using System.Runtime.InteropServices;
                using Microsoft.Win32;
                namespace Wallpaper
                {
                    public enum Style : int
                    {
	                    Tile, Center, Stretch, Fill, Fit, Span, NoChange
                    }

                    public class Setter
                    {
	                    public const int SetDesktopWallpaper = 20;
	                    public const int UpdateIniFile = 0x01;
	                    public const int SendWinIniChange = 0x02;
	                    [DllImport( "user32.dll", SetLastError = true, CharSet = CharSet.Auto )]
	                    private static extern int SystemParametersInfo ( int uAction, int uParam, string lpvParam, int fuWinIni );
	                    public static void SetWallpaper ( string path, Wallpaper.Style style )
                        {
		                    SystemParametersInfo( SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange );
		                    RegistryKey key = Registry.CurrentUser.OpenSubKey( "Control Panel\\Desktop", true );
		                    switch( style )
		                    {
			                    case Style.Tile :
			                    key.SetValue( @"WallpaperStyle", "0" ) ;
			                    key.SetValue( @"TileWallpaper", "1" ) ;
			                    break;
			                    case Style.Center :
			                    key.SetValue( @"WallpaperStyle", "0" ) ;
			                    key.SetValue( @"TileWallpaper", "0" ) ;
			                    break;
			                    case Style.Stretch :
			                    key.SetValue( @"WallpaperStyle", "2" ) ;
			                    key.SetValue( @"TileWallpaper", "0" ) ;
			                    break;
			                    case Style.Fill :
			                    key.SetValue( @"WallpaperStyle", "10" ) ;
			                    key.SetValue( @"TileWallpaper", "0" ) ;
			                    break;
			                    case Style.Fit :
			                    key.SetValue( @"WallpaperStyle", "6" ) ;
			                    key.SetValue( @"TileWallpaper", "0" ) ;
			                    break;
			                    case Style.Span :
			                    key.SetValue( @"WallpaperStyle", "22" ) ;
			                    key.SetValue( @"TileWallpaper", "0" ) ;
			                    break;
			                    case Style.NoChange :
			                    break;
		                    }
		                    key.Close();
	                    }
                    }
                }
"@
        } catch {}

        $StyleNum = @{
            Tile = 0
            Center = 1
            Stretch = 2
            Fill = 3
            Fit = 4
            Span = 5
        }
    }

    process {
        [Wallpaper.Setter]::SetWallpaper($Path, $StyleNum[$Style])

        # sometimes the wallpaper only changes after the second run, so I'll run it twice!
        Start-Sleep -Milliseconds 200
        [Wallpaper.Setter]::SetWallpaper($Path, $StyleNum[$Style])
    }
}

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
Push-Location $ScriptDir

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

Set-WallPaper -Path 'C:\PCN-OEM\pcn-wallpaper.jpg' -Style Fill

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

Write-Host "Installation finished, rebooting computer"

Pop-Location

Restart-Computer