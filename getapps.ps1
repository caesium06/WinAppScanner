<#
.SYNOPSIS
Windows App Inventory and Management Tool (PowerShell + GUI)

.DESCRIPTION
This script provides a complete tool to list, search, and manage installed applications on Windows. 
It supports both **Win32** and **UWP apps**. You can use it in two ways:

1. **PowerShell functions** directly in the console.
2. **GUI mode** via a Windows Forms interface.

FEATURES:
- Scan installed apps (Win32 & UWP)
- Search apps dynamically
- Dynamic column selection (Name, Version, Publisher, UninstallString, QuietUninstallString, Scope, etc.)
- Export app list to CSV
- Right-click context menu to:
    - Uninstall
    - Uninstall Quietly (if available)
- Fully functional via PowerShell console or GUI

---

.USAGE

# --- GUI Mode ---
# Simply run the script in PowerShell. The GUI will launch automatically.
# Use the "Scan Installed Apps" button to populate the app list.
# Use the search box to filter apps.
# Use checkboxes to select which columns to display.
# Right-click on an app row to uninstall normally or quietly.
# Export to CSV with the Export button.

# --- PowerShell Mode ---
# You can also call the internal function Get-InstalledAppDetails() in PowerShell to get app details.
# Example:
$apps = Get-InstalledAppDetails
$apps | Where-Object { $_.Name -like "*chrome*" }  # search via console
$apps | Export-Csv "$env:USERPROFILE\Desktop\InstalledApps.csv" -NoTypeInformation

---

.NOTES
- Admin privileges may be required for some uninstall operations.
- Quiet uninstall only works if the app provides a QuietUninstallString.
- The GUI will refresh the app list automatically after uninstall.
- Columns can be dynamically shown or hidden without affecting functionality.

#>

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



Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows App Inventory"
$form.Size = New-Object System.Drawing.Size(1200,700)
$form.StartPosition = "CenterScreen"

# --- Search ---
$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Text = "Search:"
$labelSearch.Location = New-Object System.Drawing.Point(10,15)
$labelSearch.AutoSize = $true

$textSearch = New-Object System.Windows.Forms.TextBox
$textSearch.Location = New-Object System.Drawing.Point(70,10)
$textSearch.Size = New-Object System.Drawing.Size(250,25)

# --- Buttons ---
$buttonScan = New-Object System.Windows.Forms.Button
$buttonScan.Text = "Scan Installed Apps"
$buttonScan.Size = New-Object System.Drawing.Size(180,30)
$buttonScan.Location = New-Object System.Drawing.Point(340,8)

$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Text = "Export to CSV"
$buttonExport.Size = New-Object System.Drawing.Size(150,30)
$buttonExport.Location = New-Object System.Drawing.Point(530,8)

# --- DataGridView ---
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(10,100)
$dataGrid.Size = New-Object System.Drawing.Size(1160,550)
$dataGrid.AutoSizeColumnsMode = "Fill"
$dataGrid.ReadOnly = $true
$dataGrid.AllowUserToAddRows = $false
$dataGrid.SelectionMode = "FullRowSelect"
$dataGrid.MultiSelect = $false
$dataGrid.ColumnHeadersHeightSizeMode = "AutoSize"

# --- FlowLayoutPanel for Columns ---
$flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$flowPanel.Location = New-Object System.Drawing.Point(700,5)
$flowPanel.Size = New-Object System.Drawing.Size(470,90)
$flowPanel.AutoScroll = $true
$flowPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$flowPanel.WrapContents = $true

$form.Controls.AddRange(@($labelSearch,$textSearch,$buttonScan,$buttonExport,$dataGrid,$flowPanel))

# --- App storage ---
$script:AllApps = @()

# --- Columns ---
$allColumns = @(
    "Name",
    "Version",
    "Publisher",
    "InstallLocation",
    "UninstallString",
    "QuietUninstallString",  # for quiet uninstall
    "Scope",
    "AppId",
    "InstallDate",
    "RegistryKey"
)
$checkedColumns = @("Name","Version","Publisher") # default visible

# --- Create CheckBoxes dynamically ---
$checkBoxes = @{}
foreach ($colName in $allColumns) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $colName
    $cb.Checked = $checkedColumns -contains $colName
    $cb.AutoSize = $true
    $flowPanel.Controls.Add($cb)
    $checkBoxes[$colName] = $cb
    $cb.Add_CheckedChanged({ Rebuild-Grid })
}

# --- Helper: Build Grid ---
function Build-Grid {
    param($collection)
    $dataGrid.Rows.Clear()
    $dataGrid.Columns.Clear()

    if ($collection.Count -eq 0) { return }

    $displayColumns = @()
    foreach ($col in $allColumns) { if ($checkBoxes[$col].Checked) { $displayColumns += $col } }

    foreach ($colName in $displayColumns) {
        if ($collection[0].ContainsKey($colName)) {
            $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $col.Name = $colName
            $col.HeaderText = $colName
            $dataGrid.Columns.Add($col) | Out-Null
        }
    }

    foreach ($app in $collection) {
        $values = $displayColumns | ForEach-Object { if ($app.ContainsKey($_)) { "$($app[$_])" } else { "" } }
        $dataGrid.Rows.Add($values) | Out-Null
    }
}

# --- Rebuild grid with search + selected columns ---
function Rebuild-Grid {
    if (-not $script:AllApps) { return }

    $filter = $textSearch.Text.Trim().ToLower()

    if ([string]::IsNullOrWhiteSpace($filter)) {
        Build-Grid $script:AllApps
    } else {
        $filtered = @()
        foreach ($app in $script:AllApps) {
            foreach ($value in $app.Values) {
                if ($value -and $value.ToLower().Contains($filter)) {
                    $filtered += $app
                    break
                }
            }
        }
        Build-Grid $filtered
    }
}

# --- Scan Button ---
$buttonScan.Add_Click({
    $form.Cursor = "WaitCursor"
    $script:AllApps = @()
    $apps = Get-InstalledAppDetails
    foreach ($app in $apps) {
        $ht = @{}
        $app.PSObject.Properties | ForEach-Object { $ht[$_.Name] = if ($_.Value) { "$($_.Value)" } else { "" } }
        $script:AllApps += $ht
    }
    Rebuild-Grid
    $form.Cursor = "Default"
})

# --- Search Box ---
$textSearch.Add_TextChanged({ Rebuild-Grid })

# --- Export Button ---
$buttonExport.Add_Click({
    if ($script:AllApps.Count -eq 0) { return }
    $path = "$env:USERPROFILE\Desktop\InstalledApps.csv"
    $script:AllApps | ForEach-Object { $_ } | Export-Csv $path -NoTypeInformation
    [System.Windows.Forms.MessageBox]::Show("Exported to $path","Export Complete")
})

# --- Context Menu for DataGridView ---
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuUninstall = New-Object System.Windows.Forms.ToolStripMenuItem
$menuUninstall.Text = "Uninstall"
$menuUninstallQuiet = New-Object System.Windows.Forms.ToolStripMenuItem
$menuUninstallQuiet.Text = "Uninstall Quietly"
$contextMenu.Items.AddRange(@($menuUninstall,$menuUninstallQuiet))
$dataGrid.ContextMenuStrip = $contextMenu

# --- Uninstall Normal ---
$menuUninstall.Add_Click({
    if ($dataGrid.SelectedRows.Count -eq 0) { return }
    $row = $dataGrid.SelectedRows[0]
    $appName = $row.Cells["Name"].Value

    # Find the full app info from AllApps
    $app = $script:AllApps | Where-Object { $_.Name -eq $appName }
    if (-not $app) { return }

    $uninstallStr = $app.UninstallString
    $appId = $app.AppId

    $confirm = [System.Windows.Forms.MessageBox]::Show("Uninstall $appName?","Confirm",[System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        if ($uninstallStr) {
            Start-Process "cmd.exe" "/c `"$uninstallStr`"" -Verb RunAs
        } elseif ($appId) {
            Remove-AppxPackage -Package $appId -ErrorAction Stop
        } else {
            [System.Windows.Forms.MessageBox]::Show("No uninstall information for $appName.","Error")
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to uninstall $appName.`n$_","Error")
    }

    $buttonScan.PerformClick()
})

# --- Uninstall Quietly ---
$menuUninstallQuiet.Add_Click({
    if ($dataGrid.SelectedRows.Count -eq 0) { return }
    $row = $dataGrid.SelectedRows[0]
    $appName = $row.Cells["Name"].Value

    # Find full app info
    $app = $script:AllApps | Where-Object { $_.Name -eq $appName }
    if (-not $app) { return }

    $quietStr = $app.QuietUninstallString
    $appId = $app.AppId

    if (-not $quietStr -and -not $appId) {
        [System.Windows.Forms.MessageBox]::Show("No quiet uninstall available for $appName.","Error")
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show("Quietly uninstall $appName?","Confirm",[System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        if ($quietStr) {
            Start-Process "cmd.exe" "/c `"$quietStr`"" -Verb RunAs
        } elseif ($appId) {
            Remove-AppxPackage -Package $appId -ErrorAction Stop
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to uninstall $appName quietly.`n$_","Error")
    }

    $buttonScan.PerformClick()
})


# --- Show Form ---
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

