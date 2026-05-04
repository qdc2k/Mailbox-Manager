<#
.SYNOPSIS
    Exchange Online Mailbox and Calendar Permission Manager
.DESCRIPTION
    A modern, dark-themed, asynchronous WPF application for managing mailbox and calendar permissions.
#>

# ==============================================================================
# 1. PREREQUISITE CHECKS & ADMIN ELEVATION
# ==============================================================================
$appDataPath = "$env:APPDATA\Mailbox-Manager"
$markerFile = "$appDataPath\ExchangeModuleOk.txt"

if (-not (Test-Path $markerFile)) {
    Write-Host "Checking prerequisites..."
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement
    
    if (-not $module) {
        # Check if running as Admin
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if (-not $isAdmin) {
            Write-Host "Exchange module not found. Restarting as Administrator to install..." -ForegroundColor Yellow
            Start-Process powershell.exe -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
            exit
        }
        
        Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Cyan
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -AcceptLicense
    }
    
    # Create marker file so we don't check heavily next time
    if (-not (Test-Path $appDataPath)) { New-Item -ItemType Directory -Path $appDataPath -Force | Out-Null }
    Set-Content -Path $markerFile -Value "Installed and Verified"
}

# Load the module into the main session
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

# ==============================================================================
# 2. XAML GUI DEFINITION (Modern Dark Theme)
# ==============================================================================
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exchange Mailbox Manager" Height="800" Width="1450" 
        Background="#1E1E1E" Foreground="#E0E0E0" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <!-- Tooltip for Calendar Roles -->
        <ToolTip x:Key="CalendarRolesTooltip" Background="#2D2D30" Foreground="#E0E0E0" BorderBrush="#3F3F46">
            <StackPanel Margin="5">
                <TextBlock FontWeight="Bold" Margin="0,0,0,5">Calendar Access Rights:</TextBlock>
                <TextBlock>• Author: CreateItems, DeleteOwned, EditOwned, Read, Visible</TextBlock>
                <TextBlock>• Contributor: CreateItems, Visible</TextBlock>
                <TextBlock>• Editor: CreateItems, DeleteAll, EditAll, Read, Visible</TextBlock>
                <TextBlock>• NonEditingAuthor: CreateItems, DeleteOwned, Read, Visible</TextBlock>
                <TextBlock>• Owner: Full Control (Create, Delete, Edit, Subfolders, Contact)</TextBlock>
                <TextBlock>• PublishingAuthor: Author + CreateSubfolders</TextBlock>
                <TextBlock>• PublishingEditor: Editor + CreateSubfolders</TextBlock>
                <TextBlock>• Reviewer: ReadItems, FolderVisible</TextBlock>
            </StackPanel>
        </ToolTip>

        <!-- Modern Business Styles -->
        <Style TargetType="ListViewItem">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="Padding" Value="2"/>
        </Style>

        <Style TargetType="{x:Type GridViewColumnHeader}">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Top Connection Bar -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <Button Name="BtnConnect" Content="Connect to Exchange" Height="30" Width="180" Background="#007ACC" Foreground="White" BorderThickness="0" Cursor="Hand" FontWeight="Bold" Margin="0,0,20,0"/>
            
            <CheckBox Name="ChkDelegated" Content="Use Delegated Authentication:" VerticalAlignment="Center" Foreground="#E0E0E0" Margin="0,0,5,0"/>
            <TextBox Name="TxtDelegatedOrg" Width="100" Height="25" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" VerticalContentAlignment="Center" Padding="5,0" IsEnabled="False" Margin="0,0,5,0"/>
            <TextBlock Text="(z.B. zkj.ch)" VerticalAlignment="Center" Foreground="#CCCCCC" FontStyle="Italic"/>
        </StackPanel>

        <!-- Search Bar for Mailboxes -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,15">
            <TextBlock Text="Search Mailboxes:" VerticalAlignment="Center" Foreground="#E0E0E0" Margin="0,0,10,0"/>
            <TextBox Name="TxtSearchMailbox" Width="300" Height="25" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" VerticalContentAlignment="Center" Padding="5,0" ToolTip="Search by Address or Type"/>
            <Button Name="BtnClearSearch" Content="Clear" Height="25" Padding="10,0" Background="#555555" Foreground="White" BorderThickness="0" Margin="10,0,0,0"/>
        </StackPanel>

        <!-- Main Content Area -->
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="335"/>
                <ColumnDefinition Width="15"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Left: Mailbox List -->
            <Border Grid.Column="0" BorderBrush="#3F3F46" BorderThickness="1" CornerRadius="3">
                <Grid Background="#252526">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Text="Mailboxes" FontWeight="Bold" Margin="10" Foreground="#007ACC" FontSize="14"/>
                    <ListView Name="ListMailboxes" Grid.Row="1" Background="Transparent" BorderThickness="0" Margin="5" ScrollViewer.HorizontalScrollBarVisibility="Disabled" DisplayMemberPath="Address">
                        <ListView.GroupStyle>
                            <GroupStyle>
                                <GroupStyle.ContainerStyle>
                                    <Style TargetType="{x:Type GroupItem}">
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="{x:Type GroupItem}">
                                                    <Expander IsExpanded="True" Background="#333337" BorderBrush="#3F3F46" BorderThickness="0,0,0,1" Margin="0,0,0,5">
                                                        <Expander.Header>
                                                            <TextBlock Text="{Binding Name}" FontWeight="Bold" Foreground="#007ACC" Margin="5,2" FontSize="13"/>
                                                        </Expander.Header>
                                                        <ItemsPresenter />
                                                    </Expander>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </GroupStyle.ContainerStyle>
                            </GroupStyle>
                        </ListView.GroupStyle>
                    </ListView>
                </Grid>
            </Border>

            <!-- Right: Permissions & Actions -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Mailbox Permissions -->
                <Border Grid.Row="0" BorderBrush="#3F3F46" BorderThickness="1" CornerRadius="3" Margin="0,0,0,10">
                    <Grid Background="#252526">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <Grid Grid.Row="0" Margin="10,10,10,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="Mailbox Permissions" FontWeight="Bold" Foreground="#007ACC" FontSize="14"/>
                                <TextBlock Name="StatusMbx" Text="" Foreground="#AAAAAA" FontStyle="Italic" Margin="10,0,0,0" VerticalAlignment="Center"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button Name="BtnAddMbx" Content="Add User" Height="22" Padding="8,0" Margin="0,0,5,0" Background="#3E3E42" Foreground="White" BorderThickness="0" FontSize="11"/>
                                <Button Name="BtnEditMbx" Content="Edit Access" Height="22" Padding="8,0" Background="#3E3E42" Foreground="White" BorderThickness="0" FontSize="11"/>
                                <Button Name="BtnRemoveMbx" Content="Remove User" Height="22" Padding="8,0" Margin="5,0,0,0" Background="#FF6B68" Foreground="White" BorderThickness="0" FontSize="11"/>
                            </StackPanel>
                        </Grid>
                        <ListView Name="GridMbxPerms" Grid.Row="1" Background="Transparent" Foreground="#E0E0E0" BorderThickness="0" Margin="5,0,5,5">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="User (UPN)" Width="250" DisplayMemberBinding="{Binding User}"/>
                                    <GridViewColumn Header="Access Rights" Width="200" DisplayMemberBinding="{Binding AccessRights}"/>
                                </GridView>
                            </ListView.View>
                        </ListView>
                    </Grid>
                </Border>

                <!-- Calendar Permissions -->
                <Border Grid.Row="1" BorderBrush="#3F3F46" BorderThickness="1" CornerRadius="3" Margin="0,0,0,10">
                    <Grid Background="#252526">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <Grid Grid.Row="0" Margin="10,10,10,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="Calendar Permissions" FontWeight="Bold" Foreground="#007ACC" FontSize="14"/>
                                <TextBlock Name="StatusCal" Text="" Foreground="#AAAAAA" FontStyle="Italic" Margin="10,0,0,0" VerticalAlignment="Center"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button Name="BtnAddCal" Content="Add User" Height="22" Padding="8,0" Margin="0,0,5,0" Background="#3E3E42" Foreground="White" BorderThickness="0" FontSize="11"/>
                                <Button Name="BtnEditCal" Content="Edit Access" Height="22" Padding="8,0" Background="#3E3E42" Foreground="White" BorderThickness="0" FontSize="11"/>
                                <Button Name="BtnRemoveCal" Content="Remove User" Height="22" Padding="8,0" Margin="5,0,0,0" Background="#FF6B68" Foreground="White" BorderThickness="0" FontSize="11"/>
                            </StackPanel>
                        </Grid>
                        <ListView Name="GridCalPerms" Grid.Row="1" Background="Transparent" Foreground="#E0E0E0" BorderThickness="0" Margin="5,0,5,5">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="User (UPN)" Width="250" DisplayMemberBinding="{Binding User}"/>
                                    <GridViewColumn Header="Calendar Rights" Width="200" DisplayMemberBinding="{Binding AccessRights}"/>
                                </GridView>
                            </ListView.View>
                        </ListView>
                    </Grid>
                </Border>
            </Grid>
        </Grid>
    </Grid>
</Window>
"@

# ==============================================================================
# 3. GUI INITIALIZATION & LOGIC
# ==============================================================================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Map XAML Elements to Variables
$BtnConnect = $Window.FindName("BtnConnect")
$ChkDelegated = $Window.FindName("ChkDelegated")
$TxtDelegatedOrg = $Window.FindName("TxtDelegatedOrg")
$ListMailboxes = $Window.FindName("ListMailboxes")
$TxtSearchMailbox = $Window.FindName("TxtSearchMailbox")
$BtnClearSearch = $Window.FindName("BtnClearSearch")
$GridMbxPerms = $Window.FindName("GridMbxPerms")
$GridCalPerms = $Window.FindName("GridCalPerms")
$BtnAddMbx = $Window.FindName("BtnAddMbx")
$BtnEditMbx = $Window.FindName("BtnEditMbx")
$BtnAddCal = $Window.FindName("BtnAddCal")
$BtnEditCal = $Window.FindName("BtnEditCal")
$BtnRemoveMbx = $Window.FindName("BtnRemoveMbx")
$BtnRemoveCal = $Window.FindName("BtnRemoveCal")
$StatusMbx = $Window.FindName("StatusMbx")
$StatusCal = $Window.FindName("StatusCal")

# Create a synchronized hashtable to share data between GUI and Runspaces
$SyncHash = [hashtable]::Synchronized(@{})
$SyncHash.Window = $Window
$SyncHash.ListMailboxes = $ListMailboxes
$SyncHash.BtnConnect = $BtnConnect
$SyncHash.TxtSearchMailbox = $TxtSearchMailbox
$SyncHash.BtnClearSearch = $BtnClearSearch
$SyncHash.StatusMbx = $StatusMbx
$SyncHash.StatusCal = $StatusCal
$SyncHash.AllMailboxes = [System.Collections.ObjectModel.ObservableCollection[object]]::new()

# Initialize Grouping and Sorting on the UI Thread (Setup once)
$view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($SyncHash.AllMailboxes)
$view.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription("Type")))
$view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription("SortOrder", [System.ComponentModel.ListSortDirection]::Ascending)))
$view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription("Address", [System.ComponentModel.ListSortDirection]::Ascending)))
$ListMailboxes.ItemsSource = $view

# Erstelle einen RunspacePool mit InitialSessionState (um SyncHash global verfügbar zu machen)
$ISS = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$ISS.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry('SyncHash', $SyncHash, $null)))
$Pool = [runspacefactory]::CreateRunspacePool(1, 5, $ISS, $Host)
$Pool.ApartmentState = "STA"
$Pool.Open()

# ==============================================================================
# 4. PERMISSION DIALOG FUNCTION
# ==============================================================================
function Show-PermissionDialog {
    param(
        [string]$Title,
        [string]$User = "",
        [string[]]$Options,
        [string]$CurrentOption = "",
        [bool]$UserEditable = $true
    )

    [xml]$DialogXaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="$Title" Height="300" Width="270" Background="#1E1E1E" Foreground="White" WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
        <Window.Resources>
            <Style TargetType="ComboBox">
                <Setter Property="Background" Value="#2D2D30"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="BorderBrush" Value="#3F3F46"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Setter Property="Padding" Value="5,0"/>
                <Setter Property="VerticalContentAlignment" Value="Center"/>
                <Setter Property="SnapsToDevicePixels" Value="True"/>
            </Style>
            <Style TargetType="ComboBoxItem">
                <Setter Property="Background" Value="#2D2D30"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="BorderThickness" Value="0"/>
            </Style>
        </Window.Resources>
        <Grid Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Text="User Principal Name (UPN):" FontWeight="Bold" Margin="0,0,0,5"/>
            <TextBox Name="TxtUser" Grid.Row="1" Height="25" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" VerticalContentAlignment="Center" Margin="0,0,0,15"/>
            
            <TextBlock Text="Access Rights / Role:" FontWeight="Bold" Grid.Row="2" Margin="0,0,0,5"/>
            <ComboBox Name="CmbRights" Grid.Row="3" Height="25" Margin="0,0,0,20"/>
            
            <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="BtnCancel" Content="Cancel" Width="80" Height="25" Margin="0,0,10,0" Background="#555555" Foreground="White" BorderThickness="0"/>
                <Button Name="BtnSave" Content="Save Changes" Width="100" Height="25" Background="#007ACC" Foreground="White" BorderThickness="0" FontWeight="Bold"/>
            </StackPanel>
        </Grid>
    </Window>
"@
    try {
        $reader = (New-Object System.Xml.XmlNodeReader $DialogXaml)
        $diag = [Windows.Markup.XamlReader]::Load($reader)
        $diag.Owner = $SyncHash.Window
    }
    catch {
        [System.Windows.MessageBox]::Show("Error loading permission dialog XAML: $($_.Exception.Message)", "XAML Error", "OK", "Error")
        return $null
    }

    $tUser = $diag.FindName("TxtUser")
    $cRights = $diag.FindName("CmbRights")
    $bSave = $diag.FindName("BtnSave")
    $bCancel = $diag.FindName("BtnCancel")

    $tUser.Text = $User
    $tUser.IsEnabled = $UserEditable
    $Options | ForEach-Object { $null = $cRights.Items.Add($_) }
    $cRights.SelectedItem = $CurrentOption

    $result = $null
    $bSave.Add_Click({
            $upn = $tUser.Text.Trim()
            # Validation: Valid Email format OR Standard OR Anonymous
            if ($upn -match '^([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}|Standard|Anonymous)$') {
                if ($cRights.SelectedItem) {
                    $script:DiagResult = @{ User = $upn; Rights = $cRights.SelectedItem }
                    $diag.Close()
                }
                else { [System.Windows.MessageBox]::Show("Please select access rights.") }
            }
            else { [System.Windows.MessageBox]::Show("Please enter a valid UPN (e.g. user@domain.com) or use 'Standard' / 'Anonymous'.") }
        })
    $bCancel.Add_Click({ $diag.Close() })

    $diag.ShowDialog() | Out-Null
    return $script:DiagResult
}

# ==============================================================================
# 4.5. CONFIRMATION DIALOG FUNCTION
# ==============================================================================
function Show-ConfirmDialog {
    param(
        [string]$Title,
        [string]$Message
    )

    [xml]$ConfirmXaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            Title="$Title" Height="170" Width="400" Background="#1E1E1E" Foreground="White" 
            WindowStartupLocation="CenterOwner" ResizeMode="NoResize" ShowInTaskbar="False">
        <Grid Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Text="$Message" TextWrapping="Wrap" VerticalAlignment="Center" FontSize="13" Foreground="#E0E0E0" FontFamily="Segoe UI"/>
            <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
                <Button Name="BtnCancel" Content="Cancel" Width="85" Height="25" Margin="0,0,10,0" Background="#3E3E42" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                <Button Name="BtnConfirm" Content="Remove" Width="85" Height="25" Background="#FF6B68" Foreground="White" BorderThickness="0" FontWeight="Bold" Cursor="Hand"/>
            </StackPanel>
        </Grid>
    </Window>
"@
    $reader = (New-Object System.Xml.XmlNodeReader $ConfirmXaml)
    $diag = [Windows.Markup.XamlReader]::Load($reader)
    $diag.Owner = $SyncHash.Window

    $bConfirm = $diag.FindName("BtnConfirm")
    $bCancel = $diag.FindName("BtnCancel")

    $script:ConfirmResult = $false
    $bConfirm.Add_Click({ $script:ConfirmResult = $true; $diag.Close() })
    $bCancel.Add_Click({ $diag.Close() })

    $diag.ShowDialog() | Out-Null
    return $script:ConfirmResult
}

# ==============================================================================
# 5. PERMISSION UPDATE LOGIC
# ==============================================================================
function Update-PermissionAsync {
    param([string]$Type, [string]$Action, [hashtable]$Data, [string]$OldRights = "")
    
    $mbx = $SyncHash.SelectedMbx
    if (-not $mbx) { return }

    # Update UI to show we are working
    if ($Type -eq "Mailbox") { $SyncHash.StatusMbx.Text = "(Saving...)" } else { $SyncHash.StatusCal.Text = "(Saving...)" }

    $PowerShell = [powershell]::Create().AddScript({
            param($mbx, $type, $action, $data, $SyncHash, $oldRights)
            Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
            try {
                if ($type -eq "Mailbox") {
                    if ($action -eq "Add") {
                        Add-MailboxPermission -Identity $mbx -User $data.User -AccessRights $data.Rights -InheritanceType All -ErrorAction Stop
                    }
                    else {
                        # Action is "Edit"
                        # For Mailbox permissions, we remove the old set and add the new one
                        Remove-MailboxPermission -Identity $mbx -User $data.User -AccessRights ($oldRights -split ',' | ForEach-Object { $_.Trim() }) -InheritanceType All -Confirm:$false -ErrorAction Stop
                        Add-MailboxPermission -Identity $mbx -User $data.User -AccessRights $data.Rights -InheritanceType All -ErrorAction Stop
                    }
                }
                else {
                    # Calendar
                    $raw = Get-EXOMailboxFolderStatistics -Identity $mbx -FolderScope Calendar
                    $folder = $raw | Where-Object { $_.FolderType -match "Calendar" } | Select-Object -First 1
                    $path = "$($mbx):\$($folder.Name)"
                    if ($action -eq "Add") {
                        Add-MailboxFolderPermission -Identity $path -User $data.User -AccessRights $data.Rights -ErrorAction Stop
                    }
                    else {
                        Set-MailboxFolderPermission -Identity $path -User $data.User -AccessRights $data.Rights -ErrorAction Stop
                    }
                }

                # Success: Trigger UI refresh on the main thread
                $SyncHash.Window.Dispatcher.Invoke({
                        $SyncHash.ListMailboxes.RaiseEvent((New-Object System.Windows.Controls.SelectionChangedEventArgs ([System.Windows.Controls.Primitives.Selector]::SelectionChangedEvent), @(), @()))
                        $SyncHash.StatusMbx.Text = ""
                        $SyncHash.StatusCal.Text = ""
                    })
            }
            catch {
                $err = $_.Exception.Message
                $SyncHash.Window.Dispatcher.Invoke({
                        [System.Windows.MessageBox]::Show("Error updating permissions:`n$err")
                        $SyncHash.StatusMbx.Text = ""
                        $SyncHash.StatusCal.Text = ""
                    })
            }
        }).AddArgument($mbx).AddArgument($Type).AddArgument($Action).AddArgument($Data).AddArgument($SyncHash).AddArgument($OldRights)

    $PowerShell.RunspacePool = $Pool
    $PowerShell.BeginInvoke() | Out-Null # No callback needed, UI updates are handled internally
}

# ==============================================================================
# 6. PERMISSION REMOVAL LOGIC
# ==============================================================================
function Remove-PermissionAsync {
    param([string]$Type, [string]$User, [string]$AccessRights)
    
    $mbx = $SyncHash.SelectedMbx
    if (-not $mbx) { return }

    # Update UI to show we are working
    if ($Type -eq "Mailbox") { $SyncHash.StatusMbx.Text = "(Removing...)" } else { $SyncHash.StatusCal.Text = "(Removing...)" }

    $PowerShell = [powershell]::Create().AddScript({
            param($mbx, $type, $user, $accessRights, $SyncHash)
            Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
            try {
                if ($type -eq "Mailbox") {
                    # Remove-MailboxPermission requires specific access rights to be removed.
                    # If the displayed AccessRights is a comma-separated string, pass it as an array.
                    $rightsToRemove = $accessRights -split ',' | ForEach-Object { $_.Trim() }
                    Remove-MailboxPermission -Identity $mbx -User $user -AccessRights $rightsToRemove -InheritanceType All -Confirm:$false -ErrorAction Stop
                }
                else {
                    # Calendar
                    $raw = Get-EXOMailboxFolderStatistics -Identity $mbx -FolderScope Calendar
                    $folder = $raw | Where-Object { $_.FolderType -match "Calendar" } | Select-Object -First 1
                    $path = "$($mbx):\$($folder.Name)"
                    Remove-MailboxFolderPermission -Identity $path -User $user -AccessRights $accessRights -Confirm:$false -ErrorAction Stop
                }

                # Success: Trigger UI refresh on the main thread
                $SyncHash.Window.Dispatcher.Invoke({
                        $SyncHash.ListMailboxes.RaiseEvent((New-Object System.Windows.Controls.SelectionChangedEventArgs ([System.Windows.Controls.Primitives.Selector]::SelectionChangedEvent), @(), @()))
                        $SyncHash.StatusMbx.Text = ""
                        $SyncHash.StatusCal.Text = ""
                    })
            }
            catch {
                $err = $_.Exception.Message
                $SyncHash.Window.Dispatcher.Invoke({
                        [System.Windows.MessageBox]::Show("Error removing permissions:`n$err")
                        $SyncHash.StatusMbx.Text = ""
                        $SyncHash.StatusCal.Text = ""
                    })
            }
        }).AddArgument($mbx).AddArgument($Type).AddArgument($User).AddArgument($AccessRights).AddArgument($SyncHash)
    
    $PowerShell.RunspacePool = $Pool
    $PowerShell.BeginInvoke() | Out-Null
}

# --- Permission Button Events ---
$mailboxRights = @("ChangeOwner", "ChangePermission", "DeleteItem", "ExternalAccount", "FullAccess", "ReadPermission")
$calendarRoles = @("Author", "Contributor", "Editor", "NonEditingAuthor", "Owner", "PublishingAuthor", "PublishingEditor", "Reviewer")

$BtnAddMbx.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnAddMbx clicked." -ForegroundColor Magenta
        $res = Show-PermissionDialog -Title "Add Mailbox Permission" -Options $mailboxRights
        if ($res) { Update-PermissionAsync -Type "Mailbox" -Action "Add" -Data $res } else { Write-Host "[$(Get-Date -f HH:mm:ss)] Show-PermissionDialog for AddMbx returned null." -ForegroundColor Yellow }
    })

$BtnEditMbx.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnEditMbx clicked." -ForegroundColor Magenta
        $sel = $GridMbxPerms.SelectedItem
        if (-not $sel) { [System.Windows.MessageBox]::Show("Please select a user from the list."); return }
        $res = Show-PermissionDialog -Title "Edit Mailbox Permission" -User $sel.User -Options $mailboxRights -CurrentOption $sel.AccessRights -UserEditable $false
        if ($res) { Update-PermissionAsync -Type "Mailbox" -Action "Edit" -Data $res -OldRights $sel.AccessRights } else { Write-Host "[$(Get-Date -f HH:mm:ss)] Show-PermissionDialog for EditMbx returned null." -ForegroundColor Yellow }
    })

$BtnRemoveMbx.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnRemoveMbx clicked." -ForegroundColor Magenta
        $sel = $GridMbxPerms.SelectedItem
        if (-not $sel) { [System.Windows.MessageBox]::Show("Please select a user from the list."); return }
        if (Show-ConfirmDialog -Title "Confirm Removal" -Message "Are you sure you want to remove the selected mailbox permission for $($sel.User) from $($SyncHash.SelectedMbx)?") {
            Remove-PermissionAsync -Type "Mailbox" -User $sel.User -AccessRights $sel.AccessRights
        }
    })

$BtnAddCal.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnAddCal clicked." -ForegroundColor Magenta
        $res = Show-PermissionDialog -Title "Add Calendar Permission" -Options $calendarRoles
        if ($res) { Update-PermissionAsync -Type "Calendar" -Action "Add" -Data $res } else { Write-Host "[$(Get-Date -f HH:mm:ss)] Show-PermissionDialog for AddCal returned null." -ForegroundColor Yellow }
    })

$BtnEditCal.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnEditCal clicked." -ForegroundColor Magenta
        $sel = $GridCalPerms.SelectedItem
        if (-not $sel) { [System.Windows.MessageBox]::Show("Please select a user from the list."); return }
        $currentRole = $sel.AccessRights -split "," | Select-Object -First 1 | ForEach-Object { $_.Trim() }
        $res = Show-PermissionDialog -Title "Edit Calendar Permission" -User $sel.User -Options $calendarRoles -CurrentOption $currentRole -UserEditable $false
        if ($res) { Update-PermissionAsync -Type "Calendar" -Action "Edit" -Data $res } else { Write-Host "[$(Get-Date -f HH:mm:ss)] Show-PermissionDialog for EditCal returned null." -ForegroundColor Yellow }
    })

$BtnRemoveCal.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnRemoveCal clicked." -ForegroundColor Magenta
        $sel = $GridCalPerms.SelectedItem
        if (-not $sel) { [System.Windows.MessageBox]::Show("Please select a user from the list."); return }
        if (Show-ConfirmDialog -Title "Confirm Removal" -Message "Are you sure you want to remove the selected calendar permission for $($sel.User) from $($SyncHash.SelectedMbx)?") {
            $currentRole = $sel.AccessRights -split "," | Select-Object -First 1 | ForEach-Object { $_.Trim() }
            Remove-PermissionAsync -Type "Calendar" -User $sel.User -AccessRights $currentRole
        }
    })

# --- Event: Checkbox Toggle ---
$ChkDelegated.Add_Checked({
        $TxtDelegatedOrg.IsEnabled = $true
    })
$ChkDelegated.Add_Unchecked({ $TxtDelegatedOrg.IsEnabled = $false })

# --- Event: Connect & Fetch Mailboxes (ASYNC) ---
$BtnConnect.Add_Click({
        $BtnConnect.IsHitTestVisible = $false
        $BtnConnect.Content = "Connecting..."
        $BtnConnect.Background = "#FFD700" # Gold
        $BtnConnect.Foreground = "Black"   # Dark text for light background

        $Delegated = $ChkDelegated.IsChecked
        $Org = $TxtDelegatedOrg.Text
        Write-Host "[$(Get-Date -f HH:mm:ss)] Attempting to connect to Exchange (Delegated: $Delegated)..." -ForegroundColor Cyan

        $PowerShell = [powershell]::Create().AddScript({
                $Delegated = $args[0]
                $Org = $args[1]
                Import-Module ExchangeOnlineManagement
        
                try {
                    if ($Delegated -and $Org) {
                        Connect-ExchangeOnline -DelegatedOrganization $Org -ShowProgress $false
                    }
                    else {
                        Connect-ExchangeOnline -ShowProgress $false
                    }

                    $SyncHash.Window.Dispatcher.Invoke({ Write-Host "[$(Get-Date -f HH:mm:ss)] Connected. Fetching all mailboxes..." -ForegroundColor Green })

                    # Update UI to Connected
                    $SyncHash.Window.Dispatcher.Invoke({
                            $SyncHash.BtnConnect.Content = "Loading Mailboxes..."
                            $SyncHash.BtnConnect.Background = "#FFD700"
                            $SyncHash.BtnConnect.Foreground = "Black"
                        })

                    # Fetch all Mailboxes
                    $allMailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object PrimarySmtpAddress, RecipientTypeDetails

                    # Filter and categorize mailboxes
                    $filteredMailboxes = [System.Collections.ArrayList]::new()
                    foreach ($mbx in $allMailboxes) {
                        # Map RecipientTypeDetails to user-friendly group names and assign a sort order
                        $groupName = switch ($mbx.RecipientTypeDetails) {
                            "UserMailbox" { "User" }
                            "SharedMailbox" { "Shared" }
                            "RoomMailbox" { "Room & Equipment" }
                            "EquipmentMailbox" { "Room & Equipment" }
                            Default { "Other" }
                        }
                        
                        $sortOrder = switch ($groupName) {
                            "User" { 1 }
                            "Shared" { 2 }
                            "Room & Equipment" { 3 }
                            "Other" { 4 }
                            Default { 99 } # Fallback for any unhandled group names, though "Other" should catch all
                        }

                        $null = $filteredMailboxes.Add([PSCustomObject]@{
                                Address      = $mbx.PrimarySmtpAddress;
                                Type         = $groupName; # Use the mapped group name for sorting/grouping
                                OriginalType = $mbx.RecipientTypeDetails; # Keep original for potential future use
                                SortOrder    = $sortOrder
                            }
                        )
                    } # End of foreach ($mbx in $allMailboxes)

                    # Sort mailboxes alphabetically by Type and then by Address
                    $sortedMailboxes = $filteredMailboxes | Sort-Object SortOrder, Address

                    $SyncHash.Window.Dispatcher.Invoke({
                            Write-Host "[$(Get-Date -f HH:mm:ss)] Loaded $($sortedMailboxes.Count) mailboxes into sections." -ForegroundColor Gray

                            # Update ObservableCollection (UI will update automatically via binding)
                            $SyncHash.AllMailboxes.Clear()
                            foreach ($item in $sortedMailboxes) { $SyncHash.AllMailboxes.Add($item) }

                            # Refresh the view to apply current filter to new items
                            $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($SyncHash.AllMailboxes)
                            $view.Refresh()

                            # Update UI to Connected
                            $SyncHash.BtnConnect.Content = "Connected"
                            $SyncHash.BtnConnect.Background = "#28A745" # Success Green
                            $SyncHash.BtnConnect.Foreground = "White"   # White text back for green
                        })
                }
                catch {
                    $err = $_.Exception.Message
                    $SyncHash.Window.Dispatcher.Invoke({
                            Write-Host "[$(Get-Date -f HH:mm:ss)] CONNECTION ERROR: $err" -ForegroundColor Red
                            $SyncHash.BtnConnect.Content = "Connection Failed"
                            $SyncHash.BtnConnect.Background = "#FF6B68" # Red
                            $SyncHash.BtnConnect.IsHitTestVisible = $true
                            $SyncHash.BtnConnect.Foreground = "White"
                        })
                }
            }).AddArgument($Delegated).AddArgument($Org)

        $PowerShell.RunspacePool = $Pool
        $AsyncResult = $PowerShell.BeginInvoke()
    })

# --- Search Debounce Timer ---
$SearchTimer = New-Object System.Windows.Threading.DispatcherTimer
$SearchTimer.Interval = [TimeSpan]::FromMilliseconds(200)

$SearchTimer.Add_Tick({
        $SearchTimer.Stop()
        $searchText = $TxtSearchMailbox.Text.ToLower().Trim()
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($SyncHash.AllMailboxes)
    
        if ([string]::IsNullOrEmpty($searchText)) {
            $view.Filter = $null
        }
        else {
            $view.Filter = {
                param($item)
                # Smart Search: matches string in Address or Type (Section name)
                $item.Address.ToLower().Contains($searchText) -or $item.Type.ToLower().Contains($searchText)
            }
        }
    })

# --- Event: Search Mailbox ---
$TxtSearchMailbox.Add_TextChanged({
        $SearchTimer.Stop()
        $SearchTimer.Start()
    })

# --- Event: Clear Search ---
$BtnClearSearch.Add_Click({
        $SyncHash.TxtSearchMailbox.Text = ""
        $SearchTimer.Stop()
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($SyncHash.AllMailboxes)
        $view.Filter = $null
    })

# --- Event: Select Mailbox (ASYNC) ---
$ListMailboxes.Add_SelectionChanged({
        $selectedItem = $ListMailboxes.SelectedItem
        if ($null -eq $selectedItem) { return }

        $mbxAddress = $selectedItem.Address
        Write-Host "[$(Get-Date -f HH:mm:ss)] Mailbox selected: $mbxAddress. Fetching permissions..." -ForegroundColor Yellow

        # UI sofort leeren (auf dem Main Thread)
        $GridMbxPerms.Items.Clear()
        $GridCalPerms.Items.Clear()
        
        $StatusMbx.Text = "(Fetching...)"
        $StatusCal.Text = "(Fetching...)"

        # Daten für den Hintergrund-Task vorbereiten
        $SyncHash.SelectedMbx = $mbxAddress
        $SyncHash.GridMbxPerms = $GridMbxPerms
        $SyncHash.GridCalPerms = $GridCalPerms

        $PsPerms = [powershell]::Create().AddScript({
                Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
                $mbx = $SyncHash.SelectedMbx
        
                try {
                    Write-Host "[$(Get-Date -f HH:mm:ss)] DEBUG: Fetching mailbox permissions for $mbx..." -ForegroundColor Gray
                    # 1. Mailbox Permissions
                    $mbxPerms = Get-EXOMailboxPermission -Identity $mbx | Where-Object { ($_.User -eq "NT AUTHORITY\SELF" -or $_.User -notlike "NT AUTHORITY\*") -and ($_.IsInherited -eq $false) }
                    $mCount = @($mbxPerms).Count
                    
                    $SyncHash.Window.Dispatcher.Invoke([Action[string, object, int]] {
                            param($targetMbx, $perms, $count)
                            Write-Host "[$(Get-Date -f HH:mm:ss)] Loaded $count mailbox permissions." -ForegroundColor Gray
                            foreach ($p in $perms) {
                                $userDisp = if ($p.User -eq "NT AUTHORITY\SELF") { $targetMbx } else { $p.User }
                                $SyncHash.GridMbxPerms.Items.Add([PSCustomObject]@{User = $userDisp; AccessRights = ($p.AccessRights -join ', ') }) | Out-Null
                            }
                            $SyncHash.StatusMbx.Text = ""
                        }, $mbx, $mbxPerms, $mCount)

                    Write-Host "[$(Get-Date -f HH:mm:ss)] DEBUG: Locating calendar folder for $mbx..." -ForegroundColor Gray
                    # 2. Kalender-Berechtigungen (Robuste Erkennung via FolderType)
                    $rawCalFolders = Get-EXOMailboxFolderStatistics -Identity $mbx -FolderScope Calendar -ErrorAction Stop
                    
                    if ($null -eq $rawCalFolders -or $rawCalFolders.Count -eq 0) {
                        Write-Host "[$(Get-Date -f HH:mm:ss)] DEBUG: Get-EXOMailboxFolderStatistics returned no folders for $mbx." -ForegroundColor Gray
                        $SyncHash.Window.Dispatcher.Invoke({ 
                                Write-Host "[$(Get-Date -f HH:mm:ss)] WARNING: No calendar folder found for $mbx" -ForegroundColor Yellow 
                                $SyncHash.StatusCal.Text = "(Not Found)"
                            })
                    }
                    else {
                        Write-Host "[$(Get-Date -f HH:mm:ss)] DEBUG: Get-EXOMailboxFolderStatistics returned $($rawCalFolders.Count) items for $mbx." -ForegroundColor Gray
                        $calFolder = $rawCalFolders | Where-Object { $_.FolderType -match "Calendar" } | Select-Object -First 1
                        
                        if ($calFolder) {
                            Write-Host "[$(Get-Date -f HH:mm:ss)] DEBUG: Found calendar folder '$($calFolder.Name)' for $mbx. Fetching permissions..." -ForegroundColor Gray
                            $calPath = "$($mbx):\$($calFolder.Name)"
                            $calPerms = Get-EXOMailboxFolderPermission -Identity $calPath -ErrorAction Stop
                            $cCount = @($calPerms).Count
                    
                            $SyncHash.Window.Dispatcher.Invoke([Action[string, object, int]] {
                                    param($fName, $perms, $count)
                                    Write-Host "[$(Get-Date -f HH:mm:ss)] Found calendar folder: $fName. Loaded $count permissions." -ForegroundColor Gray
                                    foreach ($p in $perms) {
                                        $userDisp = if ($p.User.UserPrincipalName) { $p.User.UserPrincipalName } else { $p.User.ToString() -split ":" | Select-Object -Last 1 }
                                        $SyncHash.GridCalPerms.Items.Add([PSCustomObject]@{User = $userDisp; AccessRights = ($p.AccessRights -join ', ') }) | Out-Null
                                    }
                                    $SyncHash.StatusCal.Text = ""
                                }, $calFolder.Name, $calPerms, $cCount)
                        }
                        else {
                            # This case means rawCalFolders had items, but none matched FolderType -match "Calendar"
                            $SyncHash.Window.Dispatcher.Invoke({ 
                                    Write-Host "[$(Get-Date -f HH:mm:ss)] WARNING: No 'Calendar' type folder found among returned folders for $mbx" -ForegroundColor Yellow 
                                    $SyncHash.StatusCal.Text = "(Not Found)"
                                })
                        }
                    }
                }
                catch {
                    $err = $_.Exception.Message
                    $SyncHash.Window.Dispatcher.Invoke({ 
                            Write-Host "[$(Get-Date -f HH:mm:ss)] ERROR fetching perms for $mbx : $err" -ForegroundColor Red 
                            $SyncHash.StatusMbx.Text = "(Error)"
                            $SyncHash.StatusCal.Text = "(Error)"
                        })
                }
            })
        $PsPerms.RunspacePool = $Pool
        $PsPerms.BeginInvoke() | Out-Null
    })

# Show the GUI
$Window.ShowDialog() | Out-Null