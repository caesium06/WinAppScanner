.EXE CREATION

# You can convert this PowerShell script to a standalone EXE for easier distribution 
# and to hide the source code.

# Steps:

# 1. Install PS2EXE (if not already installed):
Install-Module PS2EXE -Scope CurrentUser

# 2. Open PowerShell in the folder where your script is located:
cd "C:\Path\To\Your\Script"

# 3. Convert to EXE using a custom icon (recommended 512x512 PNG converted to multi-size .ico):
Invoke-PS2EXE "YourScript.ps1" "YourApp.exe" -noConsole -icon "app.ico"

# 4. Run YourApp.exe by double-clicking — it will auto-launch the GUI with admin prompt if needed.

# 5. Optional: If no icon is available, omit the -icon parameter:
Invoke-PS2EXE "YourScript.ps1" "YourApp.exe" -noConsole
