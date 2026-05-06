; Inno Setup Script for Mailbox Manager

[Setup]
AppId={{A7B2C3D4-E5F6-4A5B-9C8D-E1F2A3B4C5D6}}
AppName=Mailbox Manager
AppVersion=1.0
ArchitecturesInstallIn64BitMode=x64compatible
DefaultDirName={commonpf}\Mailbox-Manager
DisableDirPage=yes
DefaultGroupName=Mailbox Manager
DisableProgramGroupPage=yes
OutputDir=.
OutputBaseFilename=MailboxManagerInstaller
SetupIconFile="c:\Users\qdc\Desktop\Github\Mailbox-Manager\Mailbox_Manager.ico"
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; Required to create a shortcut on the Public (Common) Desktop
PrivilegesRequired=admin
UninstallDisplayName=Mailbox Manager
; Sets the icon for the 'Installed Apps' list in Windows Settings
UninstallDisplayIcon={app}\Mailbox_Manager.ico

[Files]
Source: "c:\Users\qdc\Desktop\Github\Mailbox-Manager\Mailbox-Manager.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "c:\Users\qdc\Desktop\Github\Mailbox-Manager\Mailbox_Manager.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Create a shortcut on the Public Desktop for all users
Name: "{commondesktop}\Mailbox Manager"; \
    Filename: "powershell.exe"; \
    Parameters: "-WindowStyle Hidden -ExecutionPolicy Bypass -File ""{app}\Mailbox-Manager.ps1"""; \
    IconFilename: "{app}\Mailbox_Manager.ico"; \
    WorkingDir: "{app}"; \
    Comment: "Exchange Online Mailbox Manager"

[Run]
; Install required PowerShell modules silently during the installation process while we have Admin rights
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -Command ""& {{ try {{ [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Write-Host 'Step 1: NuGet...'; Get-PackageProvider -Name 'NuGet' -ForceBootstrap | Out-Null; Write-Host 'Step 2: PSGallery...'; Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue; Write-Host 'Step 3: Exchange Online Module...'; Install-Module -Name ExchangeOnlineManagement -Scope AllUsers -Force -AllowClobber -Confirm:$false -WarningAction SilentlyContinue; Write-Host 'Success!' -ForegroundColor Green; Start-Sleep -Seconds 2 } catch {{ Write-Host $_ -ForegroundColor Red; Read-Host 'Press Enter to continue...' } }"""; \
    StatusMsg: "Installing required Exchange Online PowerShell modules (this may take a minute)..."