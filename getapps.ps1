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
