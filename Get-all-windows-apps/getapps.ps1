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

# --- Process DPI Awareness via Win32 API ---
try {
    $shcore = Add-Type -MemberDefinition '[DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int value);' -Name "DpiUtil" -Namespace "Win32" -PassThru
    $shcore::SetProcessDpiAwareness(1) | Out-Null
} catch {}

function Get-InstalledAppDetails {
    [CmdletBinding()]
    param(
        [string]$NameLike
    )

    # Use a List[Object] for performance and compatibility with PS2EXE compilation
    $results = [System.Collections.Generic.List[Object]]::new()

    # -----------------------------
    # WIN32 APPS (Registry-based)
    # -----------------------------
    $regViews = @(
        [PSCustomObject]@{ Hive = [Microsoft.Win32.RegistryHive]::LocalMachine; View = [Microsoft.Win32.RegistryView]::Registry64; Scope = "System"; Arch = "x64/Native" },
        [PSCustomObject]@{ Hive = [Microsoft.Win32.RegistryHive]::LocalMachine; View = [Microsoft.Win32.RegistryView]::Registry32; Scope = "System"; Arch = "x86" },
        [PSCustomObject]@{ Hive = [Microsoft.Win32.RegistryHive]::CurrentUser;  View = [Microsoft.Win32.RegistryView]::Default;    Scope = "User";   Arch = "x64/Native" }
    )

    $uninstallKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

    foreach ($rv in $regViews) {
        try {
            $regKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($rv.Hive, $rv.View)
            $subKey = $regKey.OpenSubKey($uninstallKeyPath)
            if (-not $subKey) { continue }

            foreach ($subKeyName in $subKey.GetSubKeyNames()) {
                try {
                    $appKey = $subKey.OpenSubKey($subKeyName)
                    if (-not $appKey) { continue }

                    $displayName = $appKey.GetValue("DisplayName")
                    if ([string]::IsNullOrWhiteSpace($displayName)) {
                        $appKey.Close()
                        continue
                    }

                    if ($NameLike -and ($displayName -notlike "*$NameLike*")) {
                        $appKey.Close()
                        continue
                    }

                    $displayVersion       = $appKey.GetValue("DisplayVersion")
                    $publisher            = $appKey.GetValue("Publisher")
                    $installLocation      = $appKey.GetValue("InstallLocation")
                    $installDate          = $appKey.GetValue("InstallDate")
                    $uninstallString      = $appKey.GetValue("UninstallString")
                    $quietUninstallString = $appKey.GetValue("QuietUninstallString")
                    $estimatedSize        = $appKey.GetValue("EstimatedSize")

                    # Format InstallDate (YYYYMMDD to YYYY-MM-DD)
                    $rawDate = "$installDate"
                    $formattedDate = if ($rawDate -match "^\d{8}$") {
                        $rawDate.Insert(4, "-").Insert(7, "-")
                    } else { $rawDate }

                    # Size in MB
                    $sizeMB = 0
                    if ($estimatedSize -ne $null) {
                        $sizeMB = [math]::Round([double]$estimatedSize / 1024, 2)
                    }

                    # Reconstruct registry path compatible with standard PowerShell registry path representation
                    $hiveName = if ($rv.Hive -eq [Microsoft.Win32.RegistryHive]::LocalMachine) { "HKEY_LOCAL_MACHINE" } else { "HKEY_CURRENT_USER" }
                    $registryPath = "Microsoft.PowerShell.Core\Registry::$hiveName\$uninstallKeyPath\$subKeyName"

                    $results.Add([PSCustomObject]@{
                        Name                = [string]$displayName
                        Version             = [string]$displayVersion
                        Publisher           = [string]$publisher
                        InstallLocation     = [string]$installLocation
                        InstallDate         = [string]$formattedDate
                        UninstallString     = [string]$uninstallString
                        QuietUninstallString= [string]$quietUninstallString
                        SizeMB              = [double]$sizeMB
                        Architecture        = [string]$rv.Arch
                        Scope               = [string]$rv.Scope
                        AppType             = "Win32"
                        RegistryKey         = [string]$registryPath
                    })

                    $appKey.Close()
                } catch {
                    # Skip problematic individual keys (e.g., access denied on particular entries)
                }
            }
            $subKey.Close()
            $regKey.Close()
        } catch {
            # Skip failures opening the entire hive
        }
    }

    # -----------------------------
    # UWP / MSIX / STORE APPS
    # -----------------------------
    $appxPkgs = $null
    try {
        # -ErrorAction Stop forces catch block to run if user lacks rights for system-wide query
        $appxPkgs = Get-AppxPackage -AllUsers -ErrorAction Stop
    } catch {
        $appxPkgs = Get-AppxPackage -ErrorAction SilentlyContinue
    }

    if ($appxPkgs) {
        foreach ($pkg in $appxPkgs) {
            try {
                if ($NameLike -and ($pkg.Name -notlike "*$NameLike*")) { continue }

                # Safely access properties; disposed/transient uninstalled packages will throw an exception
                $name = [string]$pkg.Name
                $version = [string]$pkg.Version
                $publisher = [string]$pkg.Publisher
                $installLocation = [string]$pkg.InstallLocation
                $fullName = [string]$pkg.PackageFullName
                $arch = [string]$pkg.Architecture
                $isFramework = $pkg.IsFramework

                $results.Add([PSCustomObject]@{
                    Name                = $name
                    Version             = $version
                    Publisher           = $publisher
                    InstallLocation     = $installLocation
                    InstallDate         = ""
                    UninstallString     = "Remove-AppxPackage -Package `"$fullName`""
                    QuietUninstallString= ""
                    SizeMB              = [double]0
                    Architecture        = $arch
                    Scope               = if ($isFramework) { "Framework" } else { "User/System (Appx)" }
                    AppType             = "UWP/MSIX"
                    RegistryKey         = ""
                })
            } catch {
                # Skip package if it is in a transient uninstalled or disposed state
            }
        }
    }

    return $results | Sort-Object Name
}

# --- Log Exception Helper ---
function Log-Exception {
    param(
        [Parameter(Mandatory=$true)]
        $Exception,
        [string]$Context
    )
    $logPath = "c:\Users\Caesium\Downloads\useful scripts\Get-all-windows-apps\crash_log.txt"
    $errText = "$Context caught at $(Get-Date):`r`n" + $Exception.ToString()
    if ($Error.Count -gt 0 -and $Error[0].ScriptStackTrace) {
        $errText += "`r`n`r`nScript Stack Trace:`r`n" + $Error[0].ScriptStackTrace
    }
    try {
        [System.IO.File]::AppendAllText($logPath, "`r`n`r`n" + $errText)
    } catch {}
}

# --- Load Windows Forms Assembly ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Enable Visual Styles ---
try {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
} catch {}

# --- Global Unhandled Exception Handlers ---
try {
    [System.Windows.Forms.Application]::add_ThreadException({
        param($sender, $e)
        $logPath = "c:\Users\Caesium\Downloads\useful scripts\Get-all-windows-apps\crash_log.txt"
        $errText = "ThreadException caught at $(Get-Date):`r`n" + $e.Exception.ToString() + "`r`n`r`nStack Trace:`r`n" + $e.Exception.StackTrace
        [System.IO.File]::WriteAllText($logPath, $errText)
        [System.Windows.Forms.MessageBox]::Show($e.Exception.Message, "Unhandled Exception (Logged)")
    })

    [AppDomain]::CurrentDomain.add_UnhandledException({
        param($sender, $e)
        $logPath = "c:\Users\Caesium\Downloads\useful scripts\Get-all-windows-apps\crash_log.txt"
        $errText = "UnhandledException caught at $(Get-Date):`r`n" + $e.ExceptionObject.ToString()
        [System.IO.File]::WriteAllText($logPath, $errText)
        [System.Windows.Forms.MessageBox]::Show($e.ExceptionObject.Message, "Unhandled Exception (Logged)")
    })
} catch {}

# --- Color Palette (Dark Theme) ---
$colorBg = [System.Drawing.Color]::FromArgb(30, 30, 36)         # Dark Slate (#1E1E24)
$colorPanel = [System.Drawing.Color]::FromArgb(42, 42, 53)      # Slate Gray (#2A2A35)
$colorAccent = [System.Drawing.Color]::FromArgb(91, 95, 151)    # Indigo/Violet (#5B5F97)
$colorText = [System.Drawing.Color]::FromArgb(245, 245, 247)    # Off-White (#F5F5F7)
$colorMuted = [System.Drawing.Color]::FromArgb(170, 170, 185)   # Light Gray (#AAAAAA)
$colorInputBg = [System.Drawing.Color]::FromArgb(55, 55, 70)    # TextBox backgrounds
$colorGridLines = [System.Drawing.Color]::FromArgb(50, 50, 60)  # DataGrid gridlines

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows App Scanner"
$form.Size = New-Object System.Drawing.Size(1250, 750)
$form.StartPosition = "CenterScreen"
$form.BackColor = $colorBg

# --- Search ---
$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Text = "Search:"
$labelSearch.Location = New-Object System.Drawing.Point(15, 18)
$labelSearch.AutoSize = $true
$labelSearch.ForeColor = $colorText
$labelSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

$textSearch = New-Object System.Windows.Forms.TextBox
$textSearch.Location = New-Object System.Drawing.Point(75, 15)
$textSearch.Size = New-Object System.Drawing.Size(250, 25)
$textSearch.BackColor = $colorInputBg
$textSearch.ForeColor = [System.Drawing.Color]::White
$textSearch.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$textSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Text = "✕"
$buttonClear.Size = New-Object System.Drawing.Size(24, 23)
$buttonClear.Location = New-Object System.Drawing.Point(330, 14)
$buttonClear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonClear.FlatAppearance.BorderSize = 0
$buttonClear.BackColor = [System.Drawing.Color]::FromArgb(70, 70, 85)
$buttonClear.ForeColor = [System.Drawing.Color]::White
$buttonClear.Font = New-Object System.Drawing.Font("Segoe UI", 8)

# --- Buttons ---
$buttonScan = New-Object System.Windows.Forms.Button
$buttonScan.Text = "Load App List"
$buttonScan.Size = New-Object System.Drawing.Size(140, 30)
$buttonScan.Location = New-Object System.Drawing.Point(370, 11)
$buttonScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonScan.FlatAppearance.BorderSize = 0
$buttonScan.BackColor = $colorAccent
$buttonScan.ForeColor = [System.Drawing.Color]::White
$buttonScan.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Text = "Export to CSV"
$buttonExport.Size = New-Object System.Drawing.Size(140, 30)
$buttonExport.Location = New-Object System.Drawing.Point(520, 11)
$buttonExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonExport.FlatAppearance.BorderSize = 1
$buttonExport.FlatAppearance.BorderColor = $colorAccent
$buttonExport.BackColor = $colorPanel
$buttonExport.ForeColor = $colorText
$buttonExport.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

# --- Status Bar ---
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = $colorPanel
$statusBar.ForeColor = $colorText
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusBar.Items.Add($statusLabel) | Out-Null

# --- DataGridView ---
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(15, 100)
$dataGrid.Size = New-Object System.Drawing.Size(1205, 580)
$dataGrid.Anchor = "Top, Bottom, Left, Right"
$dataGrid.AutoSizeColumnsMode = "Fill"
$dataGrid.ReadOnly = $true
$dataGrid.AllowUserToAddRows = $false
$dataGrid.SelectionMode = "FullRowSelect"
$dataGrid.MultiSelect = $false
$dataGrid.ColumnHeadersHeightSizeMode = "AutoSize"

# Styling DataGridView for Dark Theme
$dataGrid.BackgroundColor = $colorBg
$dataGrid.ForeColor = $colorText
$dataGrid.GridColor = $colorGridLines
$dataGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$dataGrid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$dataGrid.RowHeadersVisible = $false

# Headers Style
$dataGrid.EnableHeadersVisualStyles = $false
$dataGrid.ColumnHeadersDefaultCellStyle.BackColor = $colorPanel
$dataGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$dataGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$dataGrid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $colorPanel
$dataGrid.ColumnHeadersDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White

# Rows Style
$dataGrid.DefaultCellStyle.BackColor = $colorBg
$dataGrid.DefaultCellStyle.ForeColor = $colorText
$dataGrid.DefaultCellStyle.SelectionBackColor = $colorAccent
$dataGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$dataGrid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Enable Double Buffering to avoid rendering flickering
try {
    $type = $dataGrid.GetType()
    $property = $type.GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
    $property.SetValue($dataGrid, $true, $null)
} catch {}

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
$checkedColumns = @("Name", "Version", "AppType", "UninstallString")

# --- Binding Source and DataTable ---
$bindingSource = New-Object System.Windows.Forms.BindingSource
$dataTable = New-Object System.Data.DataTable

# Populate DataTable schema once at startup
foreach ($colName in $allColumns) {
    $type = if ($colName -eq "SizeMB") { [double] } else { [string] }
    $dataTable.Columns.Add($colName, $type) | Out-Null
}

$bindingSource.DataSource = $dataTable
$dataGrid.DataSource = $bindingSource

# --- FlowLayoutPanel for Columns ---
$flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$flowPanel.Location = New-Object System.Drawing.Point(680, 5)
$flowPanel.Size = New-Object System.Drawing.Size(540, 90)
$flowPanel.Anchor = "Top, Right"
$flowPanel.AutoScroll = $true
$flowPanel.FlowDirection = "TopDown"
$flowPanel.WrapContents = $true
$flowPanel.BackColor = $colorBg

$form.Controls.AddRange(@($labelSearch, $textSearch, $buttonClear, $buttonScan, $buttonExport, $dataGrid, $flowPanel, $statusBar))

# --- Create CheckBoxes dynamically ---
$checkBoxes = @{}
foreach ($colName in $allColumns) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $colName
    $cb.Checked = $checkedColumns -contains $colName
    $cb.AutoSize = $true
    $cb.ForeColor = $colorText
    $cb.Font = New-Object System.Drawing.Font("Segoe UI", 9)
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
    $filter = $textSearch.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($filter)) {
        $bindingSource.Filter = ""
    } else {
        # Escape single quotes and bracket characters for ADO.NET RowFilter syntax
        $escapedFilter = $filter.Replace("'", "''").Replace("[", "[[]").Replace("]", "[]]").Replace("*", "[*]").Replace("%", "[%]")
        $filterStr = ""
        foreach ($colName in $allColumns) {
            if (-not $dataTable.Columns.Contains($colName)) { continue }
            
            if ($filterStr -ne "") { $filterStr += " OR " }
            
            # For numeric columns, convert to string for the LIKE operator
            if ($dataTable.Columns[$colName].DataType -ne [string]) {
                $filterStr += "CONVERT([$colName], 'System.String') LIKE '%$escapedFilter%'"
            } else {
                $filterStr += "[$colName] LIKE '%$escapedFilter%'"
            }
        }
        $bindingSource.Filter = $filterStr
    }
    $statusLabel.Text = "Total: $($dataTable.Rows.Count) | Filtered: $($bindingSource.Count)"
}

# --- Scan Button ---
$buttonScan.Add_Click({
    try {
        $form.Cursor = "WaitCursor"
        $statusLabel.Text = "Scanning for apps..."
        $form.Refresh()
        
        $apps = Get-InstalledAppDetails
        
        # Only clear rows; columns are statically defined at startup.
        # This completely prevents ADO.NET index corruption and preserves sorting!
        $dataTable.Rows.Clear()
        
        if ($apps.Count -gt 0) {
            foreach ($app in $apps) {
                $row = $dataTable.NewRow()
                foreach ($prop in $app.PSObject.Properties) {
                    $row[$prop.Name] = if ($prop.Value -ne $null -and $prop.Value -ne "") { $prop.Value } else { [DBNull]::Value }
                }
                $dataTable.Rows.Add($row)
            }
        }
        
        Update-ColumnVisibility
        Rebuild-Grid
        
        $form.Cursor = "Default"
    } catch {
        Log-Exception $_.Exception "Scan Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message + "`n`nStack trace logged.", "Scan Error")
        $form.Cursor = "Default"
    }
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
$contextMenu.BackColor = $colorPanel
$contextMenu.ForeColor = $colorText

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
             # Elevated launch for UWP packages if running in non-admin context
             $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
             if ($isAdmin) {
                 Remove-AppxPackage -Package (($uninstallStr -split "Package ")[1].Replace('"', '')) -ErrorAction Stop
             } else {
                 Start-Process "powershell.exe" -ArgumentList "-NoProfile -WindowStyle Hidden -Command `"$uninstallStr`"" -Verb RunAs
             }
        } elseif ($uninstallStr) {
            # Use cmd /s /c to strip outer quotes and safely handle nested command parsing
            Start-Process "cmd.exe" -ArgumentList "/s /c `"$uninstallStr`"" -Verb RunAs
        } else {
            [System.Windows.Forms.MessageBox]::Show("No uninstall information for $appName.", "Error")
        }
    } catch {
        Log-Exception $_.Exception "Uninstall Error"
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
        Start-Process "cmd.exe" -ArgumentList "/s /c `"$quietStr`"" -Verb RunAs
    } catch {
        Log-Exception $_.Exception "Quiet Uninstall Error"
        [System.Windows.Forms.MessageBox]::Show("Failed to uninstall $appName quietly.`n$_", "Error")
    }
    $buttonScan.PerformClick()
})

# --- Open Location ---
$menuOpenLocation.Add_Click({
    $app = Get-SelectedApp
    if ($app -and $app.InstallLocation) {
        try {
            if (Test-Path $app.InstallLocation) {
                Invoke-Item $app.InstallLocation
            } else {
                [System.Windows.Forms.MessageBox]::Show("Location does not exist: $($app.InstallLocation)", "Error")
            }
        } catch {
            Log-Exception $_.Exception "Open Location Error"
            [System.Windows.Forms.MessageBox]::Show("Failed to open location: $($app.InstallLocation)`n$_", "Error")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No install location found.", "Error")
    }
})

# --- Open Registry ---
$menuOpenRegistry.Add_Click({
    $app = Get-SelectedApp
    if ($app -and $app.RegistryKey) {
        try {
            # Convert PSPath back to standard Registry path
            $regPath = $app.RegistryKey -replace '^Microsoft.PowerShell.Core\\Registry::', ''
            $regPath = $regPath -replace '^HKEY_LOCAL_MACHINE', 'HKLM'
            $regPath = $regPath -replace '^HKEY_CURRENT_USER', 'HKCU'
            
            # Open Registry Editor at path
            $regKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit"
            Set-ItemProperty -Path $regKey -Name "LastKey" -Value "Computer\$regPath"
            Start-Process "regedit.exe"
        } catch {
            Log-Exception $_.Exception "Open Registry Error"
            [System.Windows.Forms.MessageBox]::Show("Failed to open registry key: $regPath`n$_", "Error")
        }
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
