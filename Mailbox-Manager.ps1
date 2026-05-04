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
        Title="Exchange Mailbox Manager" Height="800" Width="1000" 
        Background="#1E1E1E" Foreground="#E0E0E0" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <!-- Tooltip for Calendar Roles -->
        <ToolTip x:Key="CalendarRolesTooltip" Background="#2D2D30" Foreground="#E0E0E0" BorderBrush="#3F3F46">
            <StackPanel Margin="5">
                <TextBlock FontWeight="Bold" Margin="0,0,0,5">Calendar Access Roles:</TextBlock>
                <TextBlock>• AvailabilityOnly: Free/Busy time</TextBlock>
                <TextBlock>• LimitedDetails: Free/Busy time, subject, location</TextBlock>
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
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="Padding" Value="2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListViewItem">
                        <Border Name="Border" Padding="{TemplateBinding Padding}" Background="Transparent" SnapsToDevicePixels="True">
                            <GridViewRowPresenter VerticalAlignment="{TemplateBinding VerticalContentAlignment}" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#3E3E42"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#007ACC"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="{x:Type GridViewColumnHeader}">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#3F3F46"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type GridViewColumnHeader}">
                        <Grid>
                            <Border Name="HeaderBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                            </Border>
                            <Thumb x:Name="PART_HeaderGripper" HorizontalAlignment="Right" Margin="0,0,-5,0" Width="10" Background="Transparent" Cursor="SizeWE">
                                <Thumb.Template>
                                    <ControlTemplate TargetType="{x:Type Thumb}">
                                        <Border Background="Transparent" />
                                    </ControlTemplate>
                                </Thumb.Template>
                            </Thumb>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="HeaderBorder" Property="Background" Value="#3E3E42"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Custom ScrollViewer Style for Dark Theme -->
        <Style TargetType="{x:Type ScrollViewer}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ScrollViewer}">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <ScrollContentPresenter Grid.Column="0" Grid.Row="0" CanContentScroll="{TemplateBinding CanContentScroll}" CanHorizontallyScroll="False" CanVerticallyScroll="False"/>
                            <ScrollBar Name="PART_VerticalScrollBar" Grid.Column="1" Grid.Row="0"
                                       Value="{TemplateBinding VerticalOffset}"
                                       Maximum="{TemplateBinding ScrollableHeight}"
                                       ViewportSize="{TemplateBinding ViewportHeight}"
                                       Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}"
                                       Background="#2D2D30">
                                <ScrollBar.Template>
                                    <ControlTemplate TargetType="{x:Type ScrollBar}">
                                        <Grid>
                                            <Grid.RowDefinitions>
                                                <RowDefinition Height="Auto"/>
                                                <RowDefinition Height="*"/>
                                                <RowDefinition Height="Auto"/>
                                            </Grid.RowDefinitions>
                                            <RepeatButton Grid.Row="0" Command="{x:Static ScrollBar.LineUpCommand}" Background="#3E3E42" BorderThickness="0" Width="17" Height="17">
                                                <Path Fill="White" Data="M 0 4 L 8 4 L 4 0 Z"/>
                                            </RepeatButton>
                                            <Track Name="PART_Track" Grid.Row="1" IsDirectionReversed="true" Background="#2D2D30">
                                                <Track.DecreaseRepeatButton>
                                                    <RepeatButton Command="{x:Static ScrollBar.PageUpCommand}" Opacity="0" />
                                                </Track.DecreaseRepeatButton>
                                                <Track.Thumb>
                                                    <Thumb Background="#555555" BorderBrush="#3F3F46" BorderThickness="0" Width="17">
                                                        <Thumb.Template>
                                                            <ControlTemplate TargetType="{x:Type Thumb}">
                                                                <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="2"/>
                                                                <ControlTemplate.Triggers>
                                                                    <Trigger Property="IsMouseOver" Value="True">
                                                                        <Setter Property="Background" Value="#777777"/>
                                                                    </Trigger>
                                                                    <Trigger Property="IsDragging" Value="True">
                                                                        <Setter Property="Background" Value="#007ACC"/>
                                                                    </Trigger>
                                                                </ControlTemplate.Triggers>
                                                            </ControlTemplate>
                                                        </Thumb.Template>
                                                    </Thumb>
                                                </Track.Thumb>
                                                <Track.IncreaseRepeatButton>
                                                    <RepeatButton Command="{x:Static ScrollBar.PageDownCommand}" Opacity="0" />
                                                </Track.IncreaseRepeatButton>
                                            </Track>
                                            <RepeatButton Grid.Row="2" Command="{x:Static ScrollBar.LineDownCommand}" Background="#3E3E42" BorderThickness="0" Width="17" Height="17">
                                                <Path Fill="White" Data="M 0 0 L 4 4 L 8 0 Z"/>
                                            </RepeatButton>
                                        </Grid>
                                    </ControlTemplate>
                                </ScrollBar.Template>
                            </ScrollBar>
                            <ScrollBar Name="PART_HorizontalScrollBar" Grid.Column="0" Grid.Row="1"
                                       Orientation="Horizontal"
                                       Value="{TemplateBinding HorizontalOffset}"
                                       Maximum="{TemplateBinding ScrollableWidth}"
                                       ViewportSize="{TemplateBinding ViewportWidth}"
                                       Visibility="{TemplateBinding ComputedHorizontalScrollBarVisibility}"
                                       Background="#2D2D30">
                                <ScrollBar.Template>
                                    <ControlTemplate TargetType="{x:Type ScrollBar}">
                                        <Grid>
                                            <!-- Horizontal scrollbar template (similar to vertical, but rotated) -->
                                            <!-- For brevity, this is a placeholder. A full implementation would mirror the vertical scrollbar's structure. -->
                                            <Border Background="#2D2D30" />
                                            <Thumb Background="#555555" />
                                        </Grid>
                                    </ControlTemplate>
                                </ScrollBar.Template>
                            </ScrollBar>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
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
            <Button Name="BtnConnect" Content="Connect to Exchange" Height="45" Width="150" Background="#007ACC" Foreground="White" BorderThickness="0" Cursor="Hand" FontWeight="Bold" Margin="0,0,20,0"/>
            
            <CheckBox Name="ChkDelegated" Content="Use Delegated Authentication:" VerticalAlignment="Center" Foreground="#E0E0E0" Margin="0,0,5,0"/>
            <TextBox Name="TxtDelegatedOrg" Width="100" Height="25" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" VerticalContentAlignment="Center" Padding="5,0" IsEnabled="False" Margin="0,0,5,0"/>
            <TextBlock Text="(z.B. zkj.ch)" VerticalAlignment="Center" Foreground="#CCCCCC" FontStyle="Italic"/>
        </StackPanel>

        <!-- Row 1 left empty as requested -->
        <Grid Grid.Row="1" Height="25" Margin="0,0,0,15"/>

        <!-- Main Content Area -->
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="285" MinWidth="150"/>
                <ColumnDefinition Width="5"/>
                <ColumnDefinition Width="*" MinWidth="400"/>
            </Grid.ColumnDefinitions>

            <!-- Left: Mailbox List -->
            <Border Grid.Column="0" BorderBrush="#3F3F46" BorderThickness="1" CornerRadius="3">
                <Grid Background="#252526">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Grid Grid.Row="0" Margin="10,10,10,5">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Mailboxes" FontWeight="Bold" Foreground="#007ACC" FontSize="15"/>
                            <TextBlock Name="StatusMailboxes" Text="" Foreground="#AAAAAA" FontStyle="Italic" Margin="10,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Grid>

                    <!-- Search Bar (new position) -->
                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="10,0,10,10">
                        <TextBlock Text="Search:" VerticalAlignment="Center" Foreground="#E0E0E0" Margin="0,0,5,0"/>
                        <TextBox Name="TxtSearchMailbox" Width="160" Height="25" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" VerticalContentAlignment="Center" Padding="5,0" ToolTip="Search by Address or Type"/>
                        <Button Name="BtnClearSearch" Content="Clear" Height="25" Width="50" Background="#555555" Foreground="White" BorderThickness="0" Margin="5,0,0,0"/>
                    </StackPanel>

                    <ListView Name="ListMailboxes" Grid.Row="2" Background="Transparent" BorderThickness="0" Margin="5" ScrollViewer.HorizontalScrollBarVisibility="Disabled" SelectionMode="Single">
                        <ListView.View>
                            <GridView>
                                <GridView.ColumnHeaderContainerStyle>
                                    <Style TargetType="GridViewColumnHeader">
                                        <Setter Property="Visibility" Value="Collapsed" />
                                    </Style>
                                </GridView.ColumnHeaderContainerStyle>
                                <GridViewColumn DisplayMemberBinding="{Binding Address}" Width="255"/>
                            </GridView>
                        </ListView.View>
                        <ListView.GroupStyle>
                            <GroupStyle>
                                <GroupStyle.ContainerStyle>
                                    <Style TargetType="{x:Type GroupItem}">
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="{x:Type GroupItem}">
                                                    <Expander IsExpanded="True" Background="#333337" BorderBrush="#3F3F46" BorderThickness="0,0,0,1" Margin="0,0,0,5">
                                                        <Expander.Header>
                                                            <TextBlock Text="{Binding Name}" FontWeight="Bold" Foreground="#007ACC" Margin="5,2" FontSize="14"/>
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

            <!-- Resizable Splitter -->
            <GridSplitter Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Stretch" Background="Transparent" Width="5" Cursor="SizeWE"/>

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
                                <TextBlock Text="Mailbox Permissions" FontWeight="Bold" Foreground="#007ACC" FontSize="15"/>
                                <TextBlock Name="StatusMbx" Text="" Foreground="#AAAAAA" FontStyle="Italic" Margin="10,0,0,0" VerticalAlignment="Center"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button Name="BtnAddMbx" Content="Add" Height="22" Width="75" Margin="0,0,5,0" Background="#45B36A" Foreground="White" BorderThickness="0" FontSize="11" Cursor="Hand" FontWeight="Bold"/>
                                <Button Name="BtnEditMbx" Content="Edit" Height="22" Width="75" Background="#FFB13B" Foreground="White" BorderThickness="0" FontSize="11" Cursor="Hand" FontWeight="Bold"/>
                                <Button Name="BtnRemoveMbx" Content="Remove" Height="22" Width="75" Margin="5,0,0,0" Background="#FF6B68" Foreground="White" BorderThickness="0" FontSize="11" Cursor="Hand" FontWeight="Bold"/>
                            </StackPanel>
                        </Grid>
                        <ListView Name="GridMbxPerms" Grid.Row="1" Background="Transparent" Foreground="#E0E0E0" BorderThickness="0" Margin="5,0,5,5" SelectionMode="Single">
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
                                <TextBlock Text="Calendar Permissions" FontWeight="Bold" Foreground="#007ACC" FontSize="15"/>
                                <TextBlock Name="StatusCal" Text="" Foreground="#AAAAAA" FontStyle="Italic" Margin="10,0,0,0" VerticalAlignment="Center"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button Name="BtnAddCal" Content="Add" Height="22" Width="75" Margin="0,0,5,0" Background="#45B36A" Foreground="White" BorderThickness="0" FontSize="11" Cursor="Hand" FontWeight="Bold"/>
                                <Button Name="BtnEditCal" Content="Edit" Height="22" Width="75" Background="#FFB13B" Foreground="White" BorderThickness="0" FontSize="11" Cursor="Hand" FontWeight="Bold"/>
                                <Button Name="BtnRemoveCal" Content="Remove" Height="22" Width="75" Margin="5,0,0,0" Background="#FF6B68" Foreground="White" BorderThickness="0" FontSize="11" Cursor="Hand" FontWeight="Bold"/>
                            </StackPanel>
                        </Grid>
                        <ListView Name="GridCalPerms" Grid.Row="1" Background="Transparent" Foreground="#E0E0E0" BorderThickness="0" Margin="5,0,5,5" SelectionMode="Single">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="User (UPN)" Width="250" DisplayMemberBinding="{Binding User}"/>
                                    <GridViewColumn Header="Calendar Roles" Width="200" DisplayMemberBinding="{Binding AccessRights}"/>
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
$StatusMailboxes = $Window.FindName("StatusMailboxes")

# Create a synchronized hashtable to share data between GUI and Runspaces
$SyncHash = [hashtable]::Synchronized(@{})
$SyncHash.Window = $Window
$SyncHash.ListMailboxes = $ListMailboxes
$SyncHash.BtnConnect = $BtnConnect
$SyncHash.TxtSearchMailbox = $TxtSearchMailbox
$SyncHash.BtnClearSearch = $BtnClearSearch
$SyncHash.StatusMbx = $StatusMbx
$SyncHash.StatusCal = $StatusCal
$SyncHash.StatusMailboxes = $StatusMailboxes
$SyncHash.GridMbxPerms = $GridMbxPerms
$SyncHash.GridCalPerms = $GridCalPerms
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
$Pool = [runspacefactory]::CreateRunspacePool(1, 1, $ISS, $Host)
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
        [string]$CurrentAutomapping = "", # New parameter for Automapping
        [bool]$ShowAutomapping = $true,
        [bool]$UserEditable = $true,
        [string]$Label = "Access Rights / Role:" # Label for the main access dropdown
    )

    $winHeight = if ($ShowAutomapping) { 310 } else { 235 }
    $autoVisibility = if ($ShowAutomapping) { "Visible" } else { "Collapsed" }
    $rightsMargin = if ($ShowAutomapping) { "0,0,0,10" } else { "0,0,0,20" }

    [xml]$DialogXaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="$Title" Height="$winHeight" Width="270" Background="#1E1E1E" Foreground="White" WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
        <Window.Resources>
            <Style TargetType="ComboBox">
                <Setter Property="Background" Value="#2D2D30"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="BorderBrush" Value="#3F3F46"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Setter Property="Padding" Value="5,0"/>
                <Setter Property="VerticalContentAlignment" Value="Center"/>
                <Setter Property="SnapsToDevicePixels" Value="True"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="ComboBox">
                            <Grid>
                                <ToggleButton Name="ToggleButton" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Focusable="false" IsChecked="{Binding Path=IsDropDownOpen,Mode=TwoWay,RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press">
                                    <ToggleButton.Template>
                                        <ControlTemplate TargetType="ToggleButton">
                                            <Border Name="Border" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="0">
                                                <Grid HorizontalAlignment="Stretch">
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition />
                                                        <ColumnDefinition Width="20" />
                                                    </Grid.ColumnDefinitions>
                                                    <Path Name="Arrow" Grid.Column="1" Fill="White" HorizontalAlignment="Center" VerticalAlignment="Center" Data="M 0 0 L 4 4 L 8 0 Z"/>
                                                </Grid>
                                            </Border>
                                        </ControlTemplate>
                                    </ToggleButton.Template>
                                </ToggleButton>
                                <ContentPresenter Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" Margin="8,0,25,0" VerticalAlignment="Center" HorizontalAlignment="Left" />
                                <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                    <Grid Name="DropDown" SnapsToDevicePixels="True" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                        <Border Name="DropDownBorder" Background="#2D2D30" BorderThickness="1" BorderBrush="#3F3F46"/>
                                        <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                            <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                        </ScrollViewer>
                                    </Grid>
                                </Popup>
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            <Style TargetType="ComboBoxItem">
                <Setter Property="Background" Value="#2D2D30"/>
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="ComboBoxItem">
                            <Border Name="Gd" Background="{TemplateBinding Background}" Padding="5,2">
                                <ContentPresenter />
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter TargetName="Gd" Property="Background" Value="#007ACC"/>
                                </Trigger>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="Gd" Property="Background" Value="#3E3E42"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </Window.Resources>
        <Grid Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Text="User Principal Name (UPN):" FontWeight="Bold" Margin="0,0,0,5"/>
            <TextBox Name="TxtUser" Grid.Row="1" Height="25" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" VerticalContentAlignment="Center" Margin="0,0,0,15"/>

            <TextBlock Text="$Label" FontWeight="Bold" Grid.Row="2" Margin="0,0,0,5"/>
            <ComboBox Name="CmbRights" Grid.Row="3" Height="25" Margin="$rightsMargin"/>
            
            <TextBlock Visibility="$autoVisibility" Text="Automapping Options:" FontWeight="Bold" Grid.Row="4" Margin="0,0,0,5"/>
            <StackPanel Visibility="$autoVisibility" Grid.Row="5" Orientation="Horizontal" Margin="0,0,0,20">
                <CheckBox Name="ChkEnableAuto" Content="Enable" Foreground="#E0E0E0" Margin="0,0,25,0" Cursor="Hand"/>
                <CheckBox Name="ChkDisableAuto" Content="Disable" Foreground="#E0E0E0" Cursor="Hand"/>
            </StackPanel>

            <Grid Grid.Row="6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="10"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button Name="BtnCancel" Grid.Column="0" Content="Cancel" Height="25" Background="#555555" Foreground="White" BorderThickness="0"/>
                <Button Name="BtnSave" Grid.Column="2" Content="Save Changes" Height="25" Background="#007ACC" Foreground="White" BorderThickness="0" FontWeight="Bold"/>
            </Grid>
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
    $chkEnable = $diag.FindName("ChkEnableAuto")
    $chkDisable = $diag.FindName("ChkDisableAuto")
    $bCancel = $diag.FindName("BtnCancel")

    $tUser.Text = $User
    $tUser.IsEnabled = $UserEditable
    $Options | ForEach-Object { $null = $cRights.Items.Add($_) }
    $cRights.SelectedItem = $CurrentOption

    # Mutual exclusivity for checkboxes
    $chkEnable.Add_Checked({ $chkDisable.IsChecked = $false })
    $chkDisable.Add_Checked({ $chkEnable.IsChecked = $false })

    # Set initial state if provided (otherwise nothing is selected)
    if ($CurrentAutomapping -eq "True") { $chkEnable.IsChecked = $true }
    elseif ($CurrentAutomapping -eq "False") { $chkDisable.IsChecked = $true }

    $result = $null
    $bSave.Add_Click({
            $upn = $tUser.Text.Trim()
            # Validation: Allow Default, Anonymous, or valid email format
            if ($upn -eq "Default" -or $upn -eq "Anonymous" -or $upn -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                if ($cRights.SelectedItem) {
                    if ($ShowAutomapping -and -not ($chkEnable.IsChecked -or $chkDisable.IsChecked)) {
                        [System.Windows.MessageBox]::Show("Please select an Automapping option (Enable or Disable).")
                        return
                    }
                    $script:DiagResult = @{ 
                        User        = $upn; 
                        Rights      = $cRights.SelectedItem; 
                        Automapping = if ($ShowAutomapping) { [bool]$chkEnable.IsChecked } else { $true } 
                    }
                    $diag.Close()
                }
                else { [System.Windows.MessageBox]::Show("Please select access rights.") }
            }
            else { [System.Windows.MessageBox]::Show("Please enter a valid UPN (e.g. user@domain.com) or use 'Default' / 'Anonymous'.") }
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
$SyncHash.GetPermissionsAsync = {
    param(
        [string]$Mailbox,
        [bool]$FetchMbx = $true,
        [bool]$FetchCal = $true
    )
    
    # UI Preparation on the main thread
    $SyncHash.CurrentTargetMbx = $Mailbox
    $SyncHash.SelectedMbx = $Mailbox

    if ($FetchMbx) {
        $SyncHash.GridMbxPerms.Items.Clear()
        $SyncHash.StatusMbx.Text = "(Fetching...)"
    }
    if ($FetchCal) {
        $SyncHash.GridCalPerms.Items.Clear()
        $SyncHash.StatusCal.Text = "(Fetching...)"
    }

    $PsPerms = [powershell]::Create().AddScript({
            param($mbx, $SyncHash, $doMbx, $doCal)
            Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
            
            try {
                if ($doMbx) {
                    Write-Host "[$(Get-Date -f HH:mm:ss)] DEBUG: Fetching mailbox permissions for $mbx..." -ForegroundColor Gray
                    $mbxPerms = Get-EXOMailboxPermission -Identity $mbx | Where-Object { ($_.User -match "SELF" -or $_.User -notlike "NT AUTHORITY\*") -and ($_.IsInherited -eq $false) }
                    $mCount = @($mbxPerms).Count
                    
                    $SyncHash.Window.Dispatcher.Invoke([Action[string, object, int]] {
                            param($targetMbx, $perms, $count)
                            if ($SyncHash.CurrentTargetMbx -ne $targetMbx) { return }
                            Write-Host "[$(Get-Date -f HH:mm:ss)] Loaded $count mailbox permissions." -ForegroundColor Gray
                            foreach ($p in $perms) {
                                $uStr = $p.User.ToString()
                                $userDisp = if ($uStr -match "SELF") { $targetMbx } else { $uStr.Split(':')[-1] }

                                $SyncHash.GridMbxPerms.Items.Add([PSCustomObject]@{User = $userDisp; AccessRights = ($p.AccessRights -join ', ') }) | Out-Null
                            }
                            $SyncHash.StatusMbx.Text = ""
                        }, $mbx, $mbxPerms, $mCount)
                }

                if ($doCal) {
                    Write-Host "[$(Get-Date -f HH:mm:ss)] DEBUG: Locating calendar folder for $mbx..." -ForegroundColor Gray
                    $rawCalFolders = Get-EXOMailboxFolderStatistics -Identity $mbx -FolderScope Calendar -ErrorAction Stop
                    
                    if ($null -eq $rawCalFolders -or $rawCalFolders.Count -eq 0) {
                        $SyncHash.Window.Dispatcher.Invoke({ 
                                Write-Host "[$(Get-Date -f HH:mm:ss)] WARNING: No calendar folder found for $mbx" -ForegroundColor Yellow 
                                $SyncHash.StatusCal.Text = "(Not Found)"
                            })
                    }
                    else {
                        $calFolder = $rawCalFolders | Where-Object { $_.FolderType -match "Calendar" } | Select-Object -First 1
                        if ($calFolder) {
                            Write-Host "[$(Get-Date -f HH:mm:ss)] DEBUG: Found calendar folder '$($calFolder.Name)' for $mbx. Fetching permissions..." -ForegroundColor Gray
                            $calPath = "$($mbx):\$($calFolder.Name)"
                            $calPerms = Get-EXOMailboxFolderPermission -Identity $calPath -ErrorAction Stop | Where-Object { $_.User -notlike "NT AUTHORITY\*" }
                            $cCount = @($calPerms).Count
                    
                            $SyncHash.Window.Dispatcher.Invoke([Action[string, object, int, string]] {
                                    param($fName, $perms, $count, $targetMbx)
                                    if ($SyncHash.CurrentTargetMbx -ne $targetMbx) { return }
                                    Write-Host "[$(Get-Date -f HH:mm:ss)] Found calendar folder: $fName. Loaded $count permissions." -ForegroundColor Gray
                                    foreach ($p in $perms) {
                                        $userDisp = if ($p.User.UserPrincipalName) { $p.User.UserPrincipalName }
                                        elseif ($p.User.ToString() -match "Default") { "Default" }
                                        elseif ($p.User.ToString() -match "Anonymous") { "Anonymous" }
                                        else { $p.User.ToString() -split ":" | Select-Object -Last 1 }
                                        $SyncHash.GridCalPerms.Items.Add([PSCustomObject]@{User = $userDisp; AccessRights = ($p.AccessRights -join ', ') }) | Out-Null
                                    }
                                    $SyncHash.StatusCal.Text = ""
                                }, $calFolder.Name, $calPerms, $cCount, $mbx)
                        }
                        else {
                            $SyncHash.Window.Dispatcher.Invoke({ 
                                    Write-Host "[$(Get-Date -f HH:mm:ss)] WARNING: No 'Calendar' type folder found for $mbx" -ForegroundColor Yellow 
                                    $SyncHash.StatusCal.Text = "(Not Found)"
                                })
                        }
                    }
                }
            }
            catch {
                $err = $_.Exception.Message
                if ($err -match "is closed|broken pipe|network connection") {
                    $SyncHash.Window.Dispatcher.Invoke({
                            [System.Windows.MessageBox]::Show("Exchange Online connection lost (Closed/Broken). Please reconnect.")
                            $SyncHash.StatusMbx.Text = "(Error)"
                            $SyncHash.StatusCal.Text = "(Error)"
                        })
                }
                else {
                    $SyncHash.Window.Dispatcher.Invoke({ 
                            Write-Host "[$(Get-Date -f HH:mm:ss)] ERROR fetching perms for $mbx : $err" -ForegroundColor Red 
                            if ($doMbx) { $SyncHash.StatusMbx.Text = "(Error)" }
                            if ($doCal) { $SyncHash.StatusCal.Text = "(Error)" }
                        })
                }
            }
        }).AddArgument($Mailbox).AddArgument($SyncHash).AddArgument($FetchMbx).AddArgument($FetchCal)
    
    $PsPerms.RunspacePool = $Pool
    $PsPerms.BeginInvoke() | Out-Null
}

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
                        Add-MailboxPermission -Identity $mbx -User $data.User -AccessRights @($data.Rights) -InheritanceType All -AutoMapping $data.Automapping -ErrorAction Stop
                    }
                    else {
                        # Action is "Edit"
                        # For Mailbox permissions, we remove the old set and add the new one
                        Remove-MailboxPermission -Identity $mbx -User $data.User -AccessRights ($oldRights -split ',' | ForEach-Object { $_.Trim() }) -InheritanceType All -Confirm:$false -ErrorAction Stop
                        Add-MailboxPermission -Identity $mbx -User $data.User -AccessRights @($data.Rights) -InheritanceType All -AutoMapping $data.Automapping -ErrorAction Stop
                    }
                }
                else {
                    # Calendar
                    $raw = Get-EXOMailboxFolderStatistics -Identity $mbx -FolderScope Calendar
                    $folder = $raw | Where-Object { $_.FolderType -eq "Calendar" -or $_.Name -eq "Calendar" } | Select-Object -First 1
                    if (-not $folder) { throw "Calendar folder not found for $mbx" }

                    $path = "$($mbx):\$($folder.Name)"
                    
                    if ($action -eq "Add") {
                        Add-MailboxFolderPermission -Identity $path -User $data.User -AccessRights $data.Rights -ErrorAction Stop
                    }
                    else {
                        Set-MailboxFolderPermission -Identity $path -User $data.User -AccessRights $data.Rights -ErrorAction Stop
                    }
                }

                # Show waiting status
                $SyncHash.Window.Dispatcher.Invoke({
                        if ($type -eq "Mailbox") { $SyncHash.StatusMbx.Text = "(Waiting for Sync...)" }
                        else { $SyncHash.StatusCal.Text = "(Waiting for Sync...)" }
                    })

                # Wait for Exchange propagation
                Start-Sleep -Seconds 5

                # Success: Trigger UI refresh on the main thread
                $SyncHash.Window.Dispatcher.Invoke([Action[string, string]] {
                        param($m, $t)
                        # If Mailbox changed, refresh both. If Calendar changed, only refresh Calendar.
                        $refreshMbx = ($t -eq "Mailbox")
                        & $SyncHash.GetPermissionsAsync -Mailbox $m -FetchMbx $refreshMbx -FetchCal $true
                    }, $mbx, $type)
            }
            catch {
                $err = $_.Exception.Message
                # Check if this is a connection-related error
                # We remove broad matches to see the actual error in the popup instead of guessing
                if ($err -match "is closed|broken pipe|network connection") {
                    $SyncHash.Window.Dispatcher.Invoke({
                            [System.Windows.MessageBox]::Show("Exchange Online connection lost (Closed/Broken). Please reconnect.")
                            $SyncHash.StatusMbx.Text = ""
                            $SyncHash.StatusCal.Text = ""
                        })
                }
                else {
                    $SyncHash.Window.Dispatcher.Invoke({
                            Write-Host "[$(Get-Date -f HH:mm:ss)] Update Error: $err" -ForegroundColor Red
                            [System.Windows.MessageBox]::Show("Error updating permissions:`n$err")
                            $SyncHash.StatusMbx.Text = ""
                            $SyncHash.StatusCal.Text = ""
                        })
                }
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
                    $folder = $raw | Where-Object { $_.FolderType -eq "Calendar" -or $_.Name -eq "Calendar" } | Select-Object -First 1
                    if (-not $folder) { throw "Calendar folder not found for $mbx" }

                    $path = "$($mbx):\$($folder.Name)"
                    Remove-MailboxFolderPermission -Identity $path -User $user -Confirm:$false -ErrorAction Stop
                }

                # Show waiting status
                $SyncHash.Window.Dispatcher.Invoke({
                        if ($type -eq "Mailbox") { $SyncHash.StatusMbx.Text = "(Waiting for Sync...)" }
                        else { $SyncHash.StatusCal.Text = "(Waiting for Sync...)" }
                    })

                # Wait for Exchange propagation
                Start-Sleep -Seconds 5

                # Success: Trigger UI refresh on the main thread
                $SyncHash.Window.Dispatcher.Invoke([Action[string, string]] {
                        param($m, $t)
                        # If Mailbox changed, refresh both. If Calendar changed, only refresh Calendar.
                        $refreshMbx = ($t -eq "Mailbox")
                        & $SyncHash.GetPermissionsAsync -Mailbox $m -FetchMbx $refreshMbx -FetchCal $true
                    }, $mbx, $type)
            }
            catch {
                $err = $_.Exception.Message
                # Check if this is a connection-related error
                # We remove broad matches to see the actual error in the popup instead of guessing
                if ($err -match "is closed|broken pipe|network connection") {
                    $SyncHash.Window.Dispatcher.Invoke({
                            [System.Windows.MessageBox]::Show("Exchange Online connection lost (Closed/Broken). Please reconnect.")
                            $SyncHash.StatusMbx.Text = ""
                            $SyncHash.StatusCal.Text = ""
                        })
                }
                else {
                    $SyncHash.Window.Dispatcher.Invoke({
                            Write-Host "[$(Get-Date -f HH:mm:ss)] Remove Error: $err" -ForegroundColor Red
                            [System.Windows.MessageBox]::Show("Error removing permissions:`n$err")
                            $SyncHash.StatusMbx.Text = ""
                            $SyncHash.StatusCal.Text = ""
                        })
                }
            }
        }).AddArgument($mbx).AddArgument($Type).AddArgument($User).AddArgument($AccessRights).AddArgument($SyncHash)
    
    $PowerShell.RunspacePool = $Pool
    $PowerShell.BeginInvoke() | Out-Null
}

# --- Permission Button Events ---
$mailboxRights = @("ChangeOwner", "ChangePermission", "DeleteItem", "ExternalAccount", "FullAccess", "ReadPermission") | Sort-Object
$calendarRoles = @("None", "AvailabilityOnly", "LimitedDetails", "Author", "Contributor", "Editor", "NonEditingAuthor", "Owner", "PublishingAuthor", "PublishingEditor", "Reviewer") | Sort-Object

$BtnAddMbx.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnAddMbx clicked." -ForegroundColor Magenta
        $res = Show-PermissionDialog -Title "Add Mailbox Permission" -Options $mailboxRights -Label "Access Rights:"
        if ($res) { Update-PermissionAsync -Type "Mailbox" -Action "Add" -Data $res } else { Write-Host "[$(Get-Date -f HH:mm:ss)] Show-PermissionDialog for AddMbx returned null." -ForegroundColor Yellow }
    })

$BtnEditMbx.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnEditMbx clicked." -ForegroundColor Magenta
        $sel = $GridMbxPerms.SelectedItem
        if (-not $sel) { [System.Windows.MessageBox]::Show("Please select a user from the list."); return }
        # Extract the first access right for display in the dropdown
        $currentRight = ($sel.AccessRights -split "," | Select-Object -First 1 | ForEach-Object { $_.Trim() })
        $res = Show-PermissionDialog -Title "Edit Mailbox Permission" -User $sel.User -Options $mailboxRights -CurrentOption $currentRight -CurrentAutomapping "" -UserEditable $false -Label "Access Rights:"
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
        $res = Show-PermissionDialog -Title "Add Calendar Permission" -Options $calendarRoles -Label "Access Roles:" -ShowAutomapping $false
        if ($res) { Update-PermissionAsync -Type "Calendar" -Action "Add" -Data $res } else { Write-Host "[$(Get-Date -f HH:mm:ss)] Show-PermissionDialog for AddCal returned null." -ForegroundColor Yellow }
    })

$BtnEditCal.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnEditCal clicked." -ForegroundColor Magenta
        $sel = $GridCalPerms.SelectedItem
        if (-not $sel) { [System.Windows.MessageBox]::Show("Please select a user from the list."); return }
        $currentRole = $sel.AccessRights -split "," | Select-Object -First 1 | ForEach-Object { $_.Trim() }
        $res = Show-PermissionDialog -Title "Edit Calendar Permission" -User $sel.User -Options $calendarRoles -CurrentOption $currentRole -UserEditable $false -Label "Access Roles:" -ShowAutomapping $false
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
        # Handle Disconnection if already connected
        if ($BtnConnect.Background.ToString() -eq "#FF28A745") {
            # Success Green
            Disconnect-ExchangeOnline -Confirm:$false
            $BtnConnect.Content = "Connect to Exchange"
            $BtnConnect.Background = "#007ACC" # Blue
            $SyncHash.AllMailboxes.Clear()
            $SyncHash.GridMbxPerms.Items.Clear()
            $SyncHash.GridCalPerms.Items.Clear()
            $SyncHash.StatusMailboxes.Text = ""
            Write-Host "[$(Get-Date -f HH:mm:ss)] Disconnected from Exchange Online." -ForegroundColor Yellow
            return
        }

        $BtnConnect.IsHitTestVisible = $false
        $BtnConnect.Content = "Connecting..."
        $BtnConnect.Background = "#FFD700" # Gold
        $BtnConnect.Foreground = "Black"   # Dark text for light background

        $Delegated = $ChkDelegated.IsChecked
        $Org = $TxtDelegatedOrg.Text
        Write-Host "[$(Get-Date -f HH:mm:ss)] Attempting to connect to Exchange (Delegated: $Delegated)..." -ForegroundColor Cyan

        # Store connection info in SyncHash for background runspaces to reuse
        $SyncHash.ConnectDelegated = $Delegated
        $SyncHash.ConnectOrg = $Org

        $PowerShell = [powershell]::Create().AddScript({
                $Delegated = $args[0]
                $Org = $args[1]
                Import-Module ExchangeOnlineManagement
        
                try {
                    if ($Delegated -and $Org) {
                        Connect-ExchangeOnline -DelegatedOrganization $Org -ShowProgress $false -ErrorAction Stop
                    }
                    else {
                        Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
                    }

                    $SyncHash.Window.Dispatcher.Invoke({ Write-Host "[$(Get-Date -f HH:mm:ss)] Connected. Fetching all mailboxes..." -ForegroundColor Green })

                    # Capture the actual UPN used to login. This allows other runspaces to sync silently.
                    $info = Get-ConnectionInformation | Select-Object -First 1
                    $SyncHash.ConnectedUser = $info.UserPrincipalName
                    Write-Host "[$(Get-Date -f HH:mm:ss)] Session Identity: $($SyncHash.ConnectedUser)" -ForegroundColor Gray

                    # Update UI to Connected
                    $SyncHash.Window.Dispatcher.Invoke({
                            # Create multi-line content for the button
                            $sp = New-Object System.Windows.Controls.StackPanel
                            $txt1 = New-Object System.Windows.Controls.TextBlock -Property @{
                                Text = "Connected"; FontWeight = "Bold"; HorizontalAlignment = "Center"
                            }
                            $txt2 = New-Object System.Windows.Controls.TextBlock -Property @{
                                Text = "(click to disconnect)"; FontSize = 11; HorizontalAlignment = "Center"
                            }
                            $null = $sp.Children.Add($txt1)
                            $null = $sp.Children.Add($txt2)
                            
                            $SyncHash.BtnConnect.Content = $sp
                            $SyncHash.BtnConnect.Background = "#28A745" # Success Green
                            $SyncHash.BtnConnect.Foreground = "White"
                            $SyncHash.BtnConnect.IsHitTestVisible = $true
                            $SyncHash.StatusMailboxes.Text = "(Fetching...)"
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

                            $SyncHash.StatusMailboxes.Text = ""
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
                            $SyncHash.StatusMailboxes.Text = "(Error)"
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
        # Selection always fetches both
        & $SyncHash.GetPermissionsAsync -Mailbox $mbxAddress -FetchMbx $true -FetchCal $true
    })

# ==============================================================================
# 7. START THE APPLICATION
# ==============================================================================
$Window.ShowDialog() | Out-Null
$Pool.Dispose()