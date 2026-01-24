function Get-InstalledAppDetails {
    [CmdletBinding()]
    param(
        [string]$NameLike  # Optional filter (partial name match)
    )

    $results = @()

    # -----------------------------
    # WIN32 APPS (Registry-based)
    # -----------------------------
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",            # 64-bit system
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",# 32-bit on 64-bit
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"              # Per-user
    )

    foreach ($path in $regPaths) {
        Get-ItemProperty $path -ErrorAction SilentlyContinue | ForEach-Object {

            if ([string]::IsNullOrWhiteSpace($_.DisplayName)) { return }

            if ($NameLike -and ($_.DisplayName -notlike "*$NameLike*")) { return }

            $scope = if ($path -like "HKCU*") { "User" } else { "System" }
            $arch  = if ($path -like "*Wow6432Node*") { "x86" } else { "x64/Native" }

            $results += [PSCustomObject]@{
                Name                = $_.DisplayName
                Version             = $_.DisplayVersion
                Publisher           = $_.Publisher
                InstallLocation     = $_.InstallLocation
                InstallDate         = $_.InstallDate
                UninstallString     = $_.UninstallString
                QuietUninstallString= $_.QuietUninstallString
                EstimatedSizeKB     = $_.EstimatedSize
                SystemComponent     = $_.SystemComponent
                ReleaseType         = $_.ReleaseType
                Architecture        = $arch
                Scope               = $scope
                AppType             = "Win32"
                RegistryKey         = $_.PSPath
            }
        }
    }

    # -----------------------------
    # UWP / MSIX / STORE APPS
    # -----------------------------
    $appxPkgs = Get-AppxPackage -AllUsers

    foreach ($pkg in $appxPkgs) {

        if ($NameLike -and ($pkg.Name -notlike "*$NameLike*")) { continue }

        $results += [PSCustomObject]@{
            Name                = $pkg.Name
            Version             = $pkg.Version.ToString()
            Publisher           = $pkg.Publisher
            InstallLocation     = $pkg.InstallLocation
            InstallDate         = $null
            UninstallString     = "Remove-AppxPackage -Package `"$($pkg.PackageFullName)`""
            QuietUninstallString= $null
            EstimatedSizeKB     = $null
            SystemComponent     = $null
            ReleaseType         = $null
            Architecture        = $pkg.Architecture
            Scope               = if ($pkg.IsFramework) { "Framework" } else { "User/System (Appx)" }
            AppType             = "UWP/MSIX"
            RegistryKey         = $null
        }
    }

    return $results | Sort-Object Name
}

<#
.SYNOPSIS
    Retrieves installed application details from a Windows device.

.DESCRIPTION
    The Get-InstalledAppDetails function lists installed applications from all major
    Windows installation types, including:

    - Win32 programs (EXE/MSI)
    - 32-bit and 64-bit apps
    - Per-user installs
    - System-wide installs
    - Microsoft Store / UWP / MSIX apps

    It combines registry-based discovery with Appx package enumeration.

.USAGE

    1. Load the script into the current PowerShell session:

        . "C:\Scripts\Get-InstalledAppDetails.ps1"

       NOTE: The dot + space is required.

    2. Show all installed applications:

        Get-InstalledAppDetails

    3. Search for a specific application:

        Get-InstalledAppDetails -NameLike "chrome"

    4. Export results:

        Get-InstalledAppDetails | Export-Csv "InstalledApps.csv" -NoTypeInformation

.NOTES
    - Run PowerShell as Administrator to see system-wide applications.
    - Run as the target user to see per-user installs.
    - Some applications may not include uninstall strings or install paths
      depending on how the vendor registered the software.
#>
