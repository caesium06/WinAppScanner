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
        [string]$NameLike
    )

    $results = @()

    # -----------------------------
    # WIN32 APPS (Registry-based)
    # -----------------------------
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $regPaths) {
        Get-ItemProperty $path -ErrorAction SilentlyContinue | ForEach-Object {
            if ([string]::IsNullOrWhiteSpace($_.DisplayName)) { return }
            if ($NameLike -and ($_.DisplayName -notlike "*$NameLike*")) { return }

            $scope = if ($path -like "HKCU*") { "User" } else { "System" }
            $arch  = if ($path -like "*Wow6432Node*") { "x86" } else { "x64/Native" }
            
            # Format InstallDate (YYYYMMDD to YYYY-MM-DD)
            $rawDate = "$($_.InstallDate)"
            $formattedDate = if ($rawDate -match "^\d{8}$") { 
                $rawDate.Insert(4, "-").Insert(7, "-") 
            } else { $rawDate }

            # Size in MB
            $sizeMB = if ($_.EstimatedSize) { [math]::Round($_.EstimatedSize / 1024, 2) } else { 0 }

            $results += [PSCustomObject]@{
                Name                = [string]$_.DisplayName
                Version             = [string]$_.DisplayVersion
                Publisher           = [string]$_.Publisher
                InstallLocation     = [string]$_.InstallLocation
                InstallDate         = [string]$formattedDate
                UninstallString     = [string]$_.UninstallString
                QuietUninstallString= [string]$_.QuietUninstallString
                SizeMB              = [double]$sizeMB
                Architecture        = [string]$arch
                Scope               = [string]$scope
                AppType             = "Win32"
                RegistryKey         = [string]$_.PSPath
            }
        }
    }

    # -----------------------------
    # UWP / MSIX / STORE APPS
    # -----------------------------
    try {
        $appxPkgs = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    } catch {
        $appxPkgs = Get-AppxPackage -ErrorAction SilentlyContinue
    }

    foreach ($pkg in $appxPkgs) {
        if ($NameLike -and ($pkg.Name -notlike "*$NameLike*")) { continue }

        $results += [PSCustomObject]@{
            Name                = [string]$pkg.Name
            Version             = [string]$pkg.Version.ToString()
            Publisher           = [string]$pkg.Publisher
            InstallLocation     = [string]$pkg.InstallLocation
            InstallDate         = ""
            UninstallString     = "Remove-AppxPackage -Package `"$($pkg.PackageFullName)`""
            QuietUninstallString= ""
            SizeMB              = [double]0
            Architecture        = [string]$pkg.Architecture
            Scope               = if ($pkg.IsFramework) { "Framework" } else { "User/System (Appx)" }
            AppType             = "UWP/MSIX"
            RegistryKey         = ""
        }
    }

    return $results | Sort-Object Name
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows App Scanner"
$form.Size = New-Object System.Drawing.Size(1250, 750)
$form.StartPosition = "CenterScreen"

# --- Search ---
$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Text = "Search:"
$labelSearch.Location = New-Object System.Drawing.Point(10, 15)
$labelSearch.AutoSize = $true

$textSearch = New-Object System.Windows.Forms.TextBox
$textSearch.Location = New-Object System.Drawing.Point(70, 12)
$textSearch.Size = New-Object System.Drawing.Size(250, 25)

$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Text = "X"
$buttonClear.Size = New-Object System.Drawing.Size(25, 22)
$buttonClear.Location = New-Object System.Drawing.Point(322, 11)

# --- Buttons ---
$buttonScan = New-Object System.Windows.Forms.Button
$buttonScan.Text = "Load App List"
$buttonScan.Size = New-Object System.Drawing.Size(150, 30)
$buttonScan.Location = New-Object System.Drawing.Point(360, 8)

$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Text = "Export to CSV"
$buttonExport.Size = New-Object System.Drawing.Size(150, 30)
$buttonExport.Location = New-Object System.Drawing.Point(520, 8)

# --- Status Bar ---
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusBar.Items.Add($statusLabel) | Out-Null

# --- DataGridView ---
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(10, 100)
$dataGrid.Size = New-Object System.Drawing.Size(1210, 580)
$dataGrid.Anchor = "Top, Bottom, Left, Right"
$dataGrid.AutoSizeColumnsMode = "Fill"
$dataGrid.ReadOnly = $true
$dataGrid.AllowUserToAddRows = $false
$dataGrid.SelectionMode = "FullRowSelect"
$dataGrid.MultiSelect = $false
$dataGrid.ColumnHeadersHeightSizeMode = "AutoSize"

# --- Binding Source and DataTable ---
$bindingSource = New-Object System.Windows.Forms.BindingSource
$dataTable = New-Object System.Data.DataTable
$dataGrid.DataSource = $bindingSource

# --- FlowLayoutPanel for Columns ---
$flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$flowPanel.Location = New-Object System.Drawing.Point(700, 5)
$flowPanel.Size = New-Object System.Drawing.Size(520, 90)
$flowPanel.Anchor = "Top, Right"
$flowPanel.AutoScroll = $true
$flowPanel.FlowDirection = "TopDown"
$flowPanel.WrapContents = $true

$form.Controls.AddRange(@($labelSearch, $textSearch, $buttonClear, $buttonScan, $buttonExport, $dataGrid, $flowPanel, $statusBar))

# --- Columns ---
$allColumns = @(
    "Name",
    "Version",
    "Publisher",
    "InstallLocation",
    "InstallDate",
    "SizeMB",
    "Architecture",
    "Scope",
    "AppType",
    "UninstallString",
    "QuietUninstallString",
    "RegistryKey"
)
$checkedColumns = @("Name", "Version", "Publisher", "SizeMB", "AppType")

# --- Create CheckBoxes dynamically ---
$checkBoxes = @{}
foreach ($colName in $allColumns) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $colName
    $cb.Checked = $checkedColumns -contains $colName
    $cb.AutoSize = $true
    $flowPanel.Controls.Add($cb)
    $checkBoxes[$colName] = $cb
    $cb.Add_CheckedChanged({ Update-ColumnVisibility })
}

function Update-ColumnVisibility {
    foreach ($colName in $allColumns) {
        if ($dataGrid.Columns.Contains($colName)) {
            $dataGrid.Columns[$colName].Visible = $checkBoxes[$colName].Checked
        }
    }
}

# --- Rebuild grid with search ---
function Rebuild-Grid {
    $filter = $textSearch.Text.Trim().Replace("'", "''")
    if ([string]::IsNullOrWhiteSpace($filter)) {
        $bindingSource.Filter = ""
    } else {
        $filterStr = ""
        foreach ($colName in $allColumns) {
            if (-not $dataTable.Columns.Contains($colName)) { continue }
            
            if ($filterStr -ne "") { $filterStr += " OR " }
            
            # For numeric columns, convert to string for the LIKE operator
            if ($dataTable.Columns[$colName].DataType -ne [string]) {
                $filterStr += "CONVERT([$colName], 'System.String') LIKE '%$filter%'"
            } else {
                $filterStr += "[$colName] LIKE '%$filter%'"
            }
        }
        $bindingSource.Filter = $filterStr
    }
    $statusLabel.Text = "Total: $($dataTable.Rows.Count) | Filtered: $($bindingSource.Count)"
}

# --- Scan Button ---
$buttonScan.Add_Click({
    $form.Cursor = "WaitCursor"
    $statusLabel.Text = "Scanning for apps..."
    $form.Refresh()
    
    $apps = Get-InstalledAppDetails
    
    $dataTable.Rows.Clear()
    $dataTable.Columns.Clear()
    
    if ($apps.Count -gt 0) {
        foreach ($prop in $apps[0].PSObject.Properties.Name) {
            $type = if ($prop -eq "SizeMB") { [double] } else { [string] }
            $dataTable.Columns.Add($prop, $type) | Out-Null
        }
        
        foreach ($app in $apps) {
            $row = $dataTable.NewRow()
            foreach ($prop in $app.PSObject.Properties) {
                $row[$prop.Name] = if ($prop.Value -ne $null -and $prop.Value -ne "") { $prop.Value } else { [DBNull]::Value }
            }
            $dataTable.Rows.Add($row)
        }
    }
    
    $bindingSource.DataSource = $dataTable
    Update-ColumnVisibility
    Rebuild-Grid
    
    $form.Cursor = "Default"
})

# --- Double Click Action ---
$dataGrid.Add_CellDoubleClick({
    $menuOpenLocation.PerformClick()
})

# --- Search Box ---
$textSearch.Add_TextChanged({ Rebuild-Grid })
$buttonClear.Add_Click({ $textSearch.Text = "" })

# --- Export Button ---
$buttonExport.Add_Click({
    if ($dataTable.Rows.Count -eq 0) { return }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveDialog.Title = "Export App List"
    $saveDialog.FileName = "InstalledApps.csv"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $results = @()
        # Export only filtered data
        for ($i = 0; $i -lt $bindingSource.Count; $i++) {
            $results += $bindingSource[$i].Row
        }
        $results | Export-Csv $saveDialog.FileName -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show("Exported to $($saveDialog.FileName)", "Export Complete")
    }
})

# --- Context Menu for DataGridView ---
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuUninstall = New-Object System.Windows.Forms.ToolStripMenuItem
$menuUninstall.Text = "Uninstall"
$menuUninstallQuiet = New-Object System.Windows.Forms.ToolStripMenuItem
$menuUninstallQuiet.Text = "Uninstall Quietly"
$menuOpenLocation = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOpenLocation.Text = "Open Install Location"
$menuOpenRegistry = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOpenRegistry.Text = "Open in Registry"

$contextMenu.Items.AddRange(@($menuUninstall, $menuUninstallQuiet, (New-Object System.Windows.Forms.ToolStripSeparator), $menuOpenLocation, $menuOpenRegistry))
$dataGrid.ContextMenuStrip = $contextMenu

# --- Helper to get selected app row ---
function Get-SelectedApp {
    if ($dataGrid.SelectedRows.Count -eq 0) { return $null }
    return $dataGrid.SelectedRows[0].DataBoundItem.Row
}

# --- Uninstall Normal ---
$menuUninstall.Add_Click({
    $app = Get-SelectedApp
    if (-not $app) { return }

    $appName = $app.Name
    $uninstallStr = $app.UninstallString

    $confirm = [System.Windows.Forms.MessageBox]::Show("Uninstall $appName?", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        if ($uninstallStr -match "^Remove-AppxPackage") {
             # It's a UWP app command string
             $pkgName = ($uninstallStr -split "Package ")[1].Replace('"', '')
             Remove-AppxPackage -Package $pkgName -ErrorAction Stop
        } elseif ($uninstallStr) {
            Start-Process "cmd.exe" "/c `"$uninstallStr`"" -Verb RunAs
        } else {
            [System.Windows.Forms.MessageBox]::Show("No uninstall information for $appName.", "Error")
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to uninstall $appName.`n$_", "Error")
    }
    $buttonScan.PerformClick()
})

# --- Uninstall Quietly ---
$menuUninstallQuiet.Add_Click({
    $app = Get-SelectedApp
    if (-not $app) { return }

    $appName = $app.Name
    $quietStr = $app.QuietUninstallString

    if (-not $quietStr) {
        [System.Windows.Forms.MessageBox]::Show("No quiet uninstall available for $appName.", "Error")
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show("Quietly uninstall $appName?", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        Start-Process "cmd.exe" "/c `"$quietStr`"" -Verb RunAs
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to uninstall $appName quietly.`n$_", "Error")
    }
    $buttonScan.PerformClick()
})

# --- Open Location ---
$menuOpenLocation.Add_Click({
    $app = Get-SelectedApp
    if ($app -and $app.InstallLocation) {
        if (Test-Path $app.InstallLocation) {
            Invoke-Item $app.InstallLocation
        } else {
            [System.Windows.Forms.MessageBox]::Show("Location does not exist: $($app.InstallLocation)", "Error")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No install location found.", "Error")
    }
})

# --- Open Registry ---
$menuOpenRegistry.Add_Click({
    $app = Get-SelectedApp
    if ($app -and $app.RegistryKey) {
        # Convert PSPath to standard Registry path
        $regPath = $app.RegistryKey -replace '^Microsoft.PowerShell.Core\\Registry::', ''
        $regPath = $regPath -replace '^HKEY_LOCAL_MACHINE', 'HKLM'
        $regPath = $regPath -replace '^HKEY_CURRENT_USER', 'HKCU'
        
        # Open Registry Editor at path
        $regKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit"
        Set-ItemProperty -Path $regKey -Name "LastKey" -Value "Computer\$regPath"
        Start-Process "regedit.exe"
    } else {
        [System.Windows.Forms.MessageBox]::Show("No registry key found for this app.", "Error")
    }
})

# --- Show Form ---
$form.Add_Shown({ 
    $form.Activate()
    $buttonScan.PerformClick()
})
[void]$form.ShowDialog()
