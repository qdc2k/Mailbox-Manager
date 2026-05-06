<#
.SYNOPSIS
    Installs the Mailbox Manager application by copying files and creating a desktop shortcut.
.DESCRIPTION
    This script is intended to be run as part of an installer (e.g., Inno Setup).
    It copies the Mailbox-Manager.ps1 script and Mailbox_Manager.ico to the
    user's %APPDATA%\Mailbox-Manager directory and creates a shortcut on the
    Public Desktop.
.NOTES
    This script should be run with administrative privileges to create the
    shortcut on the Public Desktop.
#>

# Define source paths (assuming this script is in the same directory as Mailbox-Manager.ps1 and Mailbox_Manager.ico)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$sourcePs1 = Join-Path $scriptDir "Mailbox-Manager.ps1"
$sourceIco = Join-Path $scriptDir "Mailbox_Manager.ico"

# Define destination directory in %APPDATA%
$appDataMailboxManagerDir = Join-Path $env:APPDATA "Mailbox-Manager"

# 1. Copy files to %appdata%\Mailbox-Manager
Write-Host "Creating destination directory: $appDataMailboxManagerDir"
if (-not (Test-Path $appDataMailboxManagerDir)) {
    New-Item -ItemType Directory -Path $appDataMailboxManagerDir -Force | Out-Null
}

Write-Host "Copying Mailbox-Manager.ps1 to $appDataMailboxManagerDir"
Copy-Item -Path $sourcePs1 -Destination $appDataMailboxManagerDir -Force -ErrorAction Stop

Write-Host "Copying Mailbox_Manager.ico to $appDataMailboxManagerDir"
Copy-Item -Path $sourceIco -Destination $appDataMailboxManagerDir -Force -ErrorAction Stop

# 2. Create shortcut on Public Desktop
$shortcutPath = Join-Path ([Environment]::GetFolderPath("CommonDesktopDirectory")) "Mailbox Manager.lnk"
$targetPath = "powershell.exe"
$arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$appDataMailboxManagerDir\Mailbox-Manager.ps1`""
$iconLocation = Join-Path $appDataMailboxManagerDir "Mailbox_Manager.ico"

Write-Host "Creating shortcut: $shortcutPath"
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $targetPath
$Shortcut.Arguments = $arguments
$Shortcut.IconLocation = $iconLocation
$Shortcut.Description = "Exchange Online Mailbox Manager"
$Shortcut.Save()
Write-Host "Shortcut created successfully." -ForegroundColor Green

Write-Host "Installation script finished."