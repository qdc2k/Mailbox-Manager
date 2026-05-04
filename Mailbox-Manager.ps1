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
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Mailbox Permissions -->
                <Border Grid.Row="0" BorderBrush="#3F3F46" BorderThickness="1" CornerRadius="3" Margin="0,0,0,10">
                    <Grid Background="#252526">
                        <StackPanel Orientation="Horizontal" Margin="10,10,0,0">
                            <TextBlock Text="Mailbox Permissions" FontWeight="Bold" Foreground="#007ACC" FontSize="14"/>
                            <TextBlock Name="StatusMbx" Text="" Foreground="#AAAAAA" FontStyle="Italic" Margin="10,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                        <ListView Name="GridMbxPerms" Background="Transparent" Foreground="#E0E0E0" BorderThickness="0" Margin="5,35,5,5">
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
                        <StackPanel Orientation="Horizontal" Margin="10,10,0,0">
                            <TextBlock Text="Calendar Permissions" FontWeight="Bold" Foreground="#007ACC" FontSize="14"/>
                            <TextBlock Name="StatusCal" Text="" Foreground="#AAAAAA" FontStyle="Italic" Margin="10,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                        <ListView Name="GridCalPerms" Background="Transparent" Foreground="#E0E0E0" BorderThickness="0" Margin="5,35,5,5">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="User (UPN)" Width="250" DisplayMemberBinding="{Binding User}"/>
                                    <GridViewColumn Header="Calendar Rights" Width="200" DisplayMemberBinding="{Binding AccessRights}"/>
                                </GridView>
                            </ListView.View>
                        </ListView>
                    </Grid>
                </Border>

                <!-- Action Panel -->
                <Border Grid.Row="2" BorderBrush="#3F3F46" BorderThickness="1" CornerRadius="3" Background="#252526" Padding="10">
                    <StackPanel>
                        <TextBlock Text="Manage Access (Selected Context)" FontWeight="Bold" Foreground="#007ACC" Margin="0,0,0,10" FontSize="14"/>
                        <StackPanel Orientation="Horizontal">
                            <TextBox Name="TxtUserUpn" Width="200" Height="25" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" Margin="0,0,10,0" ToolTip="Enter UPN to Add/Change"/>
                            <ComboBox Name="ComboRights" Width="150" Height="25" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" Margin="0,0,10,0"/>
                            <Button Name="BtnApply" Content="Add / Update Permission" Height="25" Padding="12,0" Background="#007ACC" Foreground="White" BorderThickness="0" FontWeight="Bold"/>
                        </StackPanel>
                    </StackPanel>
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
$BtnApply = $Window.FindName("BtnApply")
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