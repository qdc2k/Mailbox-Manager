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
$settingsFile = "$appDataPath\settings.json"
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
        Title="Exchange Mailbox Manager" Height="775" Width="1188" MinWidth="1100"
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
            <Style.Triggers>
                <Trigger Property="Content" Value="{x:Null}">
                    <Setter Property="Visibility" Value="Collapsed"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Modern Minimalist ScrollBar & ScrollViewer Styles -->
        <Style x:Key="ModernScrollBarThumb" TargetType="{x:Type Thumb}">
            <Setter Property="OverridesDefaultStyle" Value="true"/>
            <Setter Property="IsTabStop" Value="false"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Thumb}">
                        <Border x:Name="rectangle" Background="#4F4F4F" CornerRadius="4" SnapsToDevicePixels="True" Margin="2"/>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="true">
                                <Setter TargetName="rectangle" Property="Background" Value="#6F6F6F"/>
                            </Trigger>
                            <Trigger Property="IsDragging" Value="true">
                                <Setter TargetName="rectangle" Property="Background" Value="#007ACC"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="{x:Type ScrollBar}">
            <Setter Property="Stylus.IsPressAndHoldEnabled" Value="false"/>
            <Setter Property="Stylus.IsFlicksEnabled" Value="false"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Width" Value="10"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ScrollBar}">
                        <Grid x:Name="Bg" Background="{TemplateBinding Background}" SnapsToDevicePixels="true">
                            <Track x:Name="PART_Track" IsDirectionReversed="true" IsEnabled="{TemplateBinding IsEnabled}">
                                <Track.Thumb>
                                    <Thumb Style="{StaticResource ModernScrollBarThumb}" />
                                </Track.Thumb>
                            </Track>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="Orientation" Value="Horizontal">
                                <Setter TargetName="Bg" Property="Height" Value="10"/>
                                <Setter Property="Width" Value="Auto"/>
                                <Setter Property="Height" Value="10"/>
                                <Setter TargetName="PART_Track" Property="IsDirectionReversed" Value="false"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="{x:Type ScrollViewer}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ScrollViewer}">
                        <Grid Background="{TemplateBinding Background}">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <ScrollContentPresenter x:Name="PART_ScrollContentPresenter" CanContentScroll="{TemplateBinding CanContentScroll}" />
                            <ScrollBar x:Name="PART_VerticalScrollBar" Grid.Column="1" Value="{TemplateBinding VerticalOffset}" Maximum="{TemplateBinding ScrollableHeight}" ViewportSize="{TemplateBinding ViewportHeight}" Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}"/>
                            <ScrollBar x:Name="PART_HorizontalScrollBar" Grid.Row="1" Orientation="Horizontal" Value="{TemplateBinding HorizontalOffset}" Maximum="{TemplateBinding ScrollableWidth}" ViewportSize="{TemplateBinding ViewportWidth}" Visibility="{TemplateBinding ComputedHorizontalScrollBarVisibility}"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Top Connection Bar -->
        <Grid Grid.Row="0" Margin="0,0,0,15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <Button Name="BtnConnect" Grid.Column="0" Content="Connect" Width="140" Height="35" Background="#007ACC" Foreground="White" BorderThickness="0" Cursor="Hand" FontWeight="Bold" Margin="0,0,15,0" VerticalAlignment="Top"/>
            <CheckBox Name="ChkFetchAllMailboxes" Grid.Column="1" Content="Always get all mailboxes" VerticalAlignment="Center" Foreground="#E0E0E0" Margin="0,0,20,0" ToolTip="If unchecked, you can fetch specific mailboxes by UPN"/>

            <Border Grid.Column="2" BorderBrush="#3F3F46" BorderThickness="1" CornerRadius="3" Padding="12,2">
                <!-- Optional Connection Settings Group -->
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Optional connection settings:" Foreground="#007ACC" FontWeight="Bold" FontSize="11" VerticalAlignment="Center" Margin="0,0,15,0"/>
                    <CheckBox Name="ChkDelegated" Content="Delegated Org:" VerticalAlignment="Center" Foreground="#E0E0E0" Margin="0,0,5,0"/>
                    <TextBox Name="TxtDelegatedOrg" Width="130" Height="24" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" VerticalContentAlignment="Center" Padding="5,0" IsEnabled="False" Margin="0,0,15,0" ToolTip="e.g. contoso.onmicrosoft.com"/>
                    <StackPanel Name="AdminUserPanel" Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="Admin-User:" VerticalAlignment="Center" Foreground="#E0E0E0" Margin="0,0,5,0"/>
                        <TextBox Name="TxtUserUPN" Width="180" Height="24" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" VerticalContentAlignment="Center" Padding="5,0" Margin="0,0,5,0" ToolTip="Your admin email address"/>
                    </StackPanel>
                    <!-- Explanation Text -->
                    <TextBlock Name="DelegatedWarning" Text="(Using Delegated Org doesn't allow pre-defined admin credentials)" Foreground="#AAAAAA" FontStyle="Italic" FontSize="10" VerticalAlignment="Center" Margin="10,0,0,0" Visibility="Collapsed"/>
                </StackPanel>
            </Border>
        </Grid>

        <!-- Main Content Area -->
        <Grid Grid.Row="1">
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
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Mailboxes" FontWeight="Bold" Foreground="#007ACC" FontSize="15" VerticalAlignment="Center"/>
                        <ProgressBar Name="ProgressMailboxesHeader" Grid.Column="1" Height="8" Background="#3F3F46" Foreground="#007ACC" Value="0" Maximum="100" IsIndeterminate="False" Visibility="Collapsed" Margin="10,0,10,0" VerticalAlignment="Center"/>
                        <TextBlock Name="ProgressMailboxesText" Grid.Column="2" Text="0%" Foreground="#007ACC" FontWeight="Bold" FontSize="12" VerticalAlignment="Center" Visibility="Collapsed" Width="42" TextAlignment="Right"/>
                    </Grid>

                    <!-- Search Bar / UPN Input (toggled based on checkbox) -->
                    <StackPanel Grid.Row="1" Orientation="Vertical" Margin="10,0,10,10">
                        <!-- Search Bar (shown when ChkFetchAllMailboxes is checked) -->
                        <StackPanel Name="SearchBarPanel" Orientation="Vertical">
                            <TextBlock Text="Search:" Foreground="#E0E0E0" Margin="0,0,0,5"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox Name="TxtSearchMailbox" Grid.Column="0" Height="25" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" VerticalContentAlignment="Center" Padding="5,0" ToolTip="Search by Address or Type"/>
                                <Button Name="BtnClearSearch" Grid.Column="1" Content="Clear" Height="25" Width="50" Background="#555555" Foreground="White" BorderThickness="0" Margin="5,0,0,0"/>
                            </Grid>
                        </StackPanel>
                        
                        <!-- UPN Input (shown when ChkFetchAllMailboxes is unchecked) -->
                        <StackPanel Name="UPNInputPanel" Orientation="Vertical" Visibility="Collapsed">
                            <TextBlock Text="Get Mailbox:" Foreground="#E0E0E0" Margin="0,0,0,5"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox Name="TxtFetchUPN" Grid.Column="0" Height="25" Background="#2D2D30" Foreground="White" BorderBrush="#3F3F46" VerticalContentAlignment="Center" Padding="5,0" ToolTip="Enter UPN (e.g. user@domain.com)"/>
                                <Button Name="BtnFetchUPN" Grid.Column="1" Content="Get" Height="25" Width="65" Background="#007ACC" Foreground="White" BorderThickness="0" Margin="5,0,0,0" FontWeight="Bold" Cursor="Hand" FontSize="11"/>
                            </Grid>
                        </StackPanel>
                    </StackPanel>

                    <!-- ListView with Mailbox List -->
                    <Grid Grid.Row="2" Margin="5">
                            <ListView Name="ListMailboxes" Background="Transparent" BorderThickness="0" 
                                      ScrollViewer.HorizontalScrollBarVisibility="Disabled" SelectionMode="Single"
                                      VirtualizingPanel.IsVirtualizing="True"
                                      VirtualizingPanel.IsVirtualizingWhenGrouping="True"
                                      VirtualizingPanel.ScrollUnit="Pixel"
                                      VirtualizingPanel.CacheLengthUnit="Item"
                                      VirtualizingPanel.CacheLength="20,20">
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
                        <TextBlock Name="StatusMailboxes" Text="" Foreground="#AAAAAA" FontStyle="Italic" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="14" IsHitTestVisible="False"/>
                    </Grid>
                </Grid>
            </Border>

            <!-- Resizable Splitter -->
            <GridSplitter Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Stretch" Background="Transparent" Width="5" Cursor="SizeWE"/>

            <!-- Right: Permissions & Actions -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="1.0*"/>
                    <RowDefinition Height="5"/>
                    <RowDefinition Height="1.4*"/>
                </Grid.RowDefinitions>

                <!-- Mailbox Permissions -->
                <Border Grid.Row="0" BorderBrush="#3F3F46" BorderThickness="1" CornerRadius="3" Margin="0,0,0,0">
                    <Grid Background="#252526">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <Grid Grid.Row="0" Margin="10,10,10,5">
                            <!-- Removed second column definition to extend width -->
                            <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                                <TextBlock Text="Mailbox Permissions" FontWeight="Bold" Foreground="#007ACC" FontSize="15" Width="170" VerticalAlignment="Center"/>
                                <Button Name="BtnAddMbx" Content="Add" Height="22" Width="65" Margin="0,0,5,0" Background="#45B36A" Foreground="White" BorderThickness="0" FontSize="10" Cursor="Hand" FontWeight="Bold"/>
                                <Button Name="BtnEditMbx" Content="Edit" Height="22" Width="65" Margin="0,0,5,0" Background="#FFB13B" Foreground="White" BorderThickness="0" FontSize="10" Cursor="Hand" FontWeight="Bold"/>
                                <Button Name="BtnRemoveMbx" Content="Remove" Height="22" Width="65" Background="#FF6B68" Foreground="White" BorderThickness="0" FontSize="10" Cursor="Hand" FontWeight="Bold"/>
                            </StackPanel>
                        </Grid>
                        <Grid Grid.Row="1" Margin="5,0,5,5">
                            <!-- Removed second column definition to extend width -->

                            <ListView Name="GridMbxPerms" Grid.Column="0" Background="Transparent" Foreground="#E0E0E0" BorderThickness="0" SelectionMode="Single" ScrollViewer.HorizontalScrollBarVisibility="Auto">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="User / Group" Width="278" DisplayMemberBinding="{Binding User}"/>
                                        <GridViewColumn Header="Access Rights" Width="375" DisplayMemberBinding="{Binding AccessRights}"/>
                                        <GridViewColumn Header="Send Permissions" Width="215" DisplayMemberBinding="{Binding SendRights}"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                            <TextBlock Name="StatusMbx" Text="" Foreground="#AAAAAA" FontStyle="Italic" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="14" IsHitTestVisible="False"/>
                        </Grid>
                    </Grid>
                </Border>

                <!-- Resizable Splitter (Top/Bottom) -->
                <GridSplitter Grid.Row="1" HorizontalAlignment="Stretch" VerticalAlignment="Center" Background="Transparent" Height="5" Cursor="SizeNS"/>

                <!-- Calendar Permissions -->
                <Border Grid.Row="2" BorderBrush="#3F3F46" BorderThickness="1" CornerRadius="3">
                    <Grid Background="#252526">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <Grid Grid.Row="0" Margin="10,10,10,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="380"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                                <TextBlock Text="Calendar Permissions" FontWeight="Bold" Foreground="#007ACC" FontSize="15" Width="170" VerticalAlignment="Center"/>
                                <Button Name="BtnAddCal" Content="Add" Height="22" Width="65" Margin="0,0,5,0" Background="#45B36A" Foreground="White" BorderThickness="0" FontSize="10" Cursor="Hand" FontWeight="Bold"/>
                                <Button Name="BtnEditCal" Content="Edit" Height="22" Width="65" Margin="0,0,5,0" Background="#FFB13B" Foreground="White" BorderThickness="0" FontSize="10" Cursor="Hand" FontWeight="Bold"/>
                                <Button Name="BtnRemoveCal" Content="Remove" Height="22" Width="65" Background="#FF6B68" Foreground="White" BorderThickness="0" FontSize="10" Cursor="Hand" FontWeight="Bold"/>
                            </StackPanel>
                            <TextBlock Grid.Column="1" Text="Definitions" FontWeight="Bold" Foreground="#007ACC" Margin="24,0,0,0" FontSize="15" VerticalAlignment="Center"/>
                        </Grid>
                        <Grid Grid.Row="1" Margin="5,0,5,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="380"/>
                            </Grid.ColumnDefinitions>
                            
                            <Grid Grid.Column="0">
                                <ListView Name="GridCalPerms" Background="Transparent" Foreground="#E0E0E0" BorderThickness="0" SelectionMode="Single">
                                    <ListView.View>
                                        <GridView>
                                            <GridViewColumn Header="User / Group" Width="278" DisplayMemberBinding="{Binding User}"/>
                                            <GridViewColumn Header="Calendar Roles" Width="190" DisplayMemberBinding="{Binding AccessRights}"/>
                                        </GridView>
                                    </ListView.View>
                                </ListView>
                                <TextBlock Name="StatusCal" Text="" Foreground="#AAAAAA" FontStyle="Italic" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="14" IsHitTestVisible="False"/>
                            </Grid>

                            <!-- Legend Column -->
                            <Border Grid.Column="1" BorderBrush="#3F3F46" BorderThickness="1,0,0,0" Margin="10,0,0,5" Padding="10,0,0,0">
                                    <StackPanel>
                                        <TextBlock Text="SPECIAL USERS" FontSize="9" FontWeight="Bold" Foreground="#007ACC" Margin="0,0,0,2"/>
                                        <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#CCCCCC" Margin="0,0,0,12">
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">Default:</Run> All internal users<LineBreak/>
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">Anonymous:</Run> All external users (everyone)
                                        </TextBlock>
                                        <TextBlock Text="ROLE DEFINITIONS" FontSize="9" FontWeight="Bold" Foreground="#007ACC" Margin="0,0,0,2"/>
                                        <TextBlock FontSize="11" TextWrapping="Wrap" Foreground="#CCCCCC">
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">Author:</Run> CreateItems, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems<LineBreak/>
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">Contributor:</Run> CreateItems, FolderVisible<LineBreak/>
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">Editor:</Run> CreateItems, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems<LineBreak/>
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">NonEditingAuthor:</Run> CreateItems, DeleteOwnedItems, FolderVisible, ReadItems<LineBreak/>
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">Owner:</Run> CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderContact, FolderOwner, FolderVisible, ReadItems<LineBreak/>
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">PublishingAuthor:</Run> CreateItems, CreateSubfolders, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems<LineBreak/>
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">PublishingEditor:</Run> CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems<LineBreak/>
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">Reviewer:</Run> FolderVisible, ReadItems<LineBreak/>
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">AvailabilityOnly:</Run> View only availability data<LineBreak/>
                                            <Run FontWeight="Bold" Foreground="#E0E0E0">LimitedDetails:</Run> View availability data with subject and location
                                        </TextBlock>
                                    </StackPanel>
                            </Border>
                        </Grid>
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
$ChkFetchAllMailboxes = $Window.FindName("ChkFetchAllMailboxes")
$TxtDelegatedOrg = $Window.FindName("TxtDelegatedOrg")
$TxtUserUPN = $Window.FindName("TxtUserUPN")
$AdminUserPanel = $Window.FindName("AdminUserPanel")
$DelegatedWarning = $Window.FindName("DelegatedWarning")
$ListMailboxes = $Window.FindName("ListMailboxes")
$TxtSearchMailbox = $Window.FindName("TxtSearchMailbox")
$TxtFetchUPN = $Window.FindName("TxtFetchUPN")
$BtnClearSearch = $Window.FindName("BtnClearSearch")
$BtnFetchUPN = $Window.FindName("BtnFetchUPN")
$SearchBarPanel = $Window.FindName("SearchBarPanel")
$UPNInputPanel = $Window.FindName("UPNInputPanel")
$ProgressMailboxesHeader = $Window.FindName("ProgressMailboxesHeader")
$ProgressMailboxesText = $Window.FindName("ProgressMailboxesText")
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

# --- Helper: Save Settings ---
function Save-ManagerSettings {
    $settings = @{
        UseDelegated = [bool]$ChkDelegated.IsChecked
        Org          = $TxtDelegatedOrg.Text
        UserUPN      = $TxtUserUPN.Text
        FetchAll     = [bool]$ChkFetchAllMailboxes.IsChecked
    }
    $settings | ConvertTo-Json | Set-Content $settingsFile -ErrorAction SilentlyContinue
}

# Create a synchronized hashtable to share data between GUI and Runspaces
$SyncHash = [hashtable]::Synchronized(@{})
$SyncHash.Window = $Window
$SyncHash.ListMailboxes = $ListMailboxes
$SyncHash.BtnConnect = $BtnConnect
$SyncHash.TxtSearchMailbox = $TxtSearchMailbox
$SyncHash.TxtFetchUPN = $TxtFetchUPN
$SyncHash.BtnClearSearch = $BtnClearSearch
$SyncHash.BtnFetchUPN = $BtnFetchUPN
$SyncHash.ProgressMailboxesHeader = $ProgressMailboxesHeader
$SyncHash.ProgressMailboxesText = $ProgressMailboxesText
$SyncHash.StatusMbx = $StatusMbx
$SyncHash.StatusCal = $StatusCal
$SyncHash.StatusMailboxes = $StatusMailboxes
$SyncHash.GridMbxPerms = $GridMbxPerms
$SyncHash.GridCalPerms = $GridCalPerms
$SyncHash.AllMailboxes = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$SyncHash.FetchedUPNs = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
$SyncHash.AddressCache = [System.Collections.Generic.List[string]]::new()

# --- Load Saved Settings ---
if (Test-Path $settingsFile) {
    try {
        $settings = Get-Content $settingsFile | ConvertFrom-Json
        $ChkDelegated.IsChecked = $settings.UseDelegated
        $TxtDelegatedOrg.Text = $settings.Org
        $TxtUserUPN.Text = $settings.UserUPN
        $ChkFetchAllMailboxes.IsChecked = if ($null -eq $settings.FetchAll) { $true } else { $settings.FetchAll }
        if ($settings.UseDelegated) { $TxtDelegatedOrg.IsEnabled = $true }
    }
    catch {
        Write-Host "Failed to load settings: $($_.Exception.Message)" -ForegroundColor Red
    }
}

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
        [bool]$ShowSendRights = $false,
        [string]$CurrentSendRights = "None",
        [bool]$UserEditable = $true,
        [string]$Label = "Access Rights / Role:", # Label for the main access dropdown
        [object[]]$UserList = @()
    )

    $winHeight = 245
    if ($ShowAutomapping) { $winHeight += 65 }
    if ($ShowSendRights) { $winHeight += 85 } # Extra height for separation
    $autoVisibility = if ($ShowAutomapping) { "Visible" } else { "Collapsed" }
    $sendVisibility = if ($ShowSendRights) { "Visible" } else { "Collapsed" }
    $rightsMargin = "0,0,0,15"

    [xml]$DialogXaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="$Title" Height="$winHeight" Width="280" Background="#1E1E1E" Foreground="White" WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
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
                                <TextBox x:Name="PART_EditableTextBox" Background="Transparent" Foreground="White" BorderThickness="0" Margin="8,0,25,0" VerticalAlignment="Center" HorizontalAlignment="Left" Focusable="True" Visibility="Collapsed" CaretBrush="White"/>
                                <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                    <Grid Name="DropDown" SnapsToDevicePixels="True" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                        <Border Name="DropDownBorder" Background="#2D2D30" BorderThickness="1" BorderBrush="#3F3F46"/>
                                        <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                            <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                        </ScrollViewer>
                                    </Grid>
                                </Popup>
                            </Grid>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsEditable" Value="True">
                                    <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                                    <Setter TargetName="ContentSite" Property="Visibility" Value="Collapsed"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
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
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Text="User Principal Name (UPN):" FontWeight="Bold" Margin="0,0,0,5"/>
            <ComboBox Name="CmbUser" Grid.Row="1" Height="25" IsEditable="True" IsTextSearchEnabled="False" Margin="0,0,0,15"/>

            <TextBlock Text="$Label" FontWeight="Bold" Grid.Row="2" Margin="0,0,0,5"/>
            <ComboBox Name="CmbRights" Grid.Row="3" Height="25" Margin="$rightsMargin"/>

            <TextBlock Name="LblAuto" Visibility="$autoVisibility" Text="Automapping Options:" FontWeight="Bold" Grid.Row="4" Margin="0,0,0,5"/>
            <StackPanel Name="PanelAuto" Visibility="$autoVisibility" Grid.Row="5" Orientation="Horizontal" Margin="0,0,0,15">
                <CheckBox Name="ChkEnableAuto" Content="Enable" Foreground="#E0E0E0" Margin="0,0,25,0" Cursor="Hand"/>
                <CheckBox Name="ChkDisableAuto" Content="Disable" Foreground="#E0E0E0" Cursor="Hand"/>
            </StackPanel>
            
            <Separator Visibility="$sendVisibility" Grid.Row="6" Background="#3F3F46" Margin="0,5,0,15"/>

            <TextBlock Visibility="$sendVisibility" Text="Send Permissions:" FontWeight="Bold" Grid.Row="7" Margin="0,0,0,5"/>
            <ComboBox Name="CmbSendRights" Visibility="$sendVisibility" Grid.Row="8" Height="25" Margin="0,0,0,20"/>

            <Grid Grid.Row="9">
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
        Show-MessageDialog -Title "XAML Error" -Message "Error loading permission dialog XAML: $($_.Exception.Message)" | Out-Null
        return $null
    }

    $cUser = $diag.FindName("CmbUser")
    $cRights = $diag.FindName("CmbRights")
    $cSendRights = $diag.FindName("CmbSendRights")
    $bSave = $diag.FindName("BtnSave")
    $lblAuto = $diag.FindName("LblAuto")
    $panelAuto = $diag.FindName("PanelAuto")
    $chkEnable = $diag.FindName("ChkEnableAuto")
    $chkDisable = $diag.FindName("ChkDisableAuto")
    $bCancel = $diag.FindName("BtnCancel")

    # Populate User List and add filtering (search-as-you-type)
    $userCollection = [System.Collections.ObjectModel.ObservableCollection[string]]::new()
    if ($UserList) { $UserList | Sort-Object | ForEach-Object { $userCollection.Add($_) } }
    $cUser.ItemsSource = $userCollection

    $cUser.Add_KeyUp({
            param($s, $e)
            # Filter list based on current text (Contains logic)
            $txt = $s.Text
            $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($s.ItemsSource)
            if ([string]::IsNullOrWhiteSpace($txt)) { $view.Filter = $null }
            else { $view.Filter = { param($item) $item.ToLower().Contains($txt.ToLower()) } }
            # Keep the dropdown open while typing to show results
            $s.IsDropDownOpen = $true
        })

    $cUser.Text = $User
    $cUser.IsEnabled = $UserEditable
    $Options | ForEach-Object { $null = $cRights.Items.Add($_) }
    $cRights.SelectedItem = $CurrentOption

    $initialRights = $CurrentOption
    $initialSendRights = $CurrentSendRights

    # Dynamically enable/disable Automapping UI based on selected Access Rights
    $updateAutoUI = {
        if ($ShowAutomapping) {
            $isFullAccess = ($cRights.SelectedItem -eq "FullAccess")
            $vis = if ($isFullAccess) { "Visible" } else { "Collapsed" }
            $lblAuto.Visibility = $vis
            $panelAuto.Visibility = $vis

            if ($isFullAccess) {
                # Default to Enable (True) if no choice has been made yet
                if (-not $chkEnable.IsChecked -and -not $chkDisable.IsChecked) {
                    $chkEnable.IsChecked = $true
                }
            }
            if (-not $isFullAccess) {
                $chkEnable.IsChecked = $false
                $chkDisable.IsChecked = $false
            }
        }
    }
    $cRights.Add_SelectionChanged($updateAutoUI)
    & $updateAutoUI # Set initial state

    if ($ShowSendRights) {
        $sendOpts = @("None", "SendAs", "SendOnBehalf")
        $sendOpts | ForEach-Object { $null = $cSendRights.Items.Add($_) }
        # Ensure the current right is selected by matching the string explicitly
        foreach ($item in $cSendRights.Items) {
            if ($item -eq $CurrentSendRights) { $cSendRights.SelectedItem = $item; break }
        }
    }

    # Mutual exclusivity for checkboxes
    $chkEnable.Add_Checked({ $chkDisable.IsChecked = $false })
    $chkDisable.Add_Checked({ $chkEnable.IsChecked = $false })

    # Set initial state if provided (otherwise nothing is selected)
    if ($CurrentAutomapping -eq "True") { $chkEnable.IsChecked = $true }
    elseif ($CurrentAutomapping -eq "False") { $chkDisable.IsChecked = $true }

    $script:DiagResult = $null
    $bSave.Add_Click({
            $upn = $cUser.Text.Trim()
            # Validation: Allow Default, Anonymous, Standard (German), or valid email format
            if ($upn -eq "Default" -or $upn -eq "Anonymous" -or $upn -eq "Standard" -or $upn -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                if ($cRights.SelectedItem) {
                    $selectedRights = $cRights.SelectedItem
                    $selectedSend = if ($ShowSendRights) { $cSendRights.SelectedItem } else { "None" }
                    $isNone = ($selectedRights -eq "None")
                    $rightsChanged = ($selectedRights -ne $initialRights)
                    $autoSelected = ($chkEnable.IsChecked -or $chkDisable.IsChecked)

                    # Mandatory Automapping only if Access Rights changed from None (or initial empty) to something else,
                    # OR if they were explicitly changed during an edit.
                    if ($ShowAutomapping -and $selectedRights -eq "FullAccess") {
                        if (-not $autoSelected) {
                            Show-MessageDialog -Title "Automapping Required" -Message "FullAccess rights require automapping flag be set. Please select Enable or Disable." | Out-Null
                            return
                        }
                    }

                    $script:DiagResult = @{ 
                        User              = $upn; 
                        Rights            = $selectedRights; # The selected rights
                        RightsChanged     = $rightsChanged; # Whether the rights were changed from initial
                        SendRights        = $selectedSend;
                        SendRightsChanged = ($selectedSend -ne $initialSendRights);
                        Automapping       = [bool]$chkEnable.IsChecked;
                        HasAutoSelection  = $autoSelected
                    }
                    $diag.Close()
                }
                else { Show-MessageDialog -Title "Missing Selection" -Message "Please select access rights." | Out-Null }
            }
            else { Show-MessageDialog -Title "Invalid User" -Message "Please enter a valid UPN (e.g. user@domain.com) or use 'Default' / 'Anonymous' / 'Standard'." | Out-Null }
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
# 4.6. MESSAGE DIALOG FUNCTION (Custom styled replacement for MessageBox)
# ==============================================================================
function Show-MessageDialog {
    param(
        [string]$Title,
        [string]$Message
    )

    [xml]$MessageXaml = @"
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
                <Button Name="BtnOk" Content="OK" Width="85" Height="25" Background="#007ACC" Foreground="White" BorderThickness="0" FontWeight="Bold" Cursor="Hand"/>
            </StackPanel>
        </Grid>
    </Window>
"@
    try {
        $reader = (New-Object System.Xml.XmlNodeReader $MessageXaml)
        $diag = [Windows.Markup.XamlReader]::Load($reader)
        $diag.Owner = $SyncHash.Window
    }
    catch {
        # Fallback to default MessageBox if custom dialog fails to load
        [System.Windows.MessageBox]::Show("Error loading custom message dialog XAML: $($_.Exception.Message)`n`nOriginal Message: $Message", "XAML Error", "OK", "Error") | Out-Null
        return
    }

    $bOk = $diag.FindName("BtnOk")
    $bOk.Add_Click({ $diag.Close() })
    $diag.ShowDialog() | Out-Null
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
                    Write-Host "[$(Get-Date -f HH:mm:ss)] DEBUG: Fetching mailbox & send permissions for $mbx..." -ForegroundColor Gray
                    $mbxPerms = Get-MailboxPermission -Identity $mbx | Where-Object { ($_.User -match "SELF" -or $_.User -notlike "NT AUTHORITY\*") -and ($_.IsInherited -eq $false) }
                    $recPerms = Get-RecipientPermission -Identity $mbx | Where-Object { $_.Trustee -match "SELF" -or $_.Trustee -notlike "NT AUTHORITY\*" }
                    $mbxObj = Get-Mailbox -Identity $mbx

                    $usersHash = [ordered]@{}
                    
                    # Helper to normalize identities (resolve SamAccountName/DN/SID to UPN)
                    $NormalizeId = {
                        param($id, $targetMbx)
                        if ($null -eq $id) { return $null }
                        $s = $id.ToString()
                        if ($s -match "SELF") { return $targetMbx }
                        $clean = $s.Split(':')[-1]
                        # If it's already a UPN, return it immediately
                        if ($clean -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') { return $clean }
                        # Try to resolve non-UPN identities to PrimarySmtpAddress for consolidation
                        try {
                            $rec = Get-Recipient -Identity $clean -ErrorAction SilentlyContinue | Select-Object -First 1
                            if ($rec.PrimarySmtpAddress) { return $rec.PrimarySmtpAddress.ToString() }
                        }
                        catch {}
                        return $clean
                    }
                    
                    foreach ($p in $mbxPerms) {
                        $userDisp = &$NormalizeId $p.User $mbx
                        if (-not $usersHash.Contains($userDisp)) {
                            $usersHash[$userDisp] = @{ User = $userDisp; AccessRights = [System.Collections.Generic.List[string]]::new(); SendRights = "None" }
                        }
                        $p.AccessRights | ForEach-Object { $usersHash[$userDisp].AccessRights.Add($_) }
                    }
                    
                    foreach ($p in $recPerms) {
                        $userDisp = &$NormalizeId $p.Trustee $mbx
                        if (-not $usersHash.Contains($userDisp)) {
                            $usersHash[$userDisp] = @{ User = $userDisp; AccessRights = [System.Collections.Generic.List[string]]::new(); SendRights = "SendAs" }
                        }
                        else {
                            $usersHash[$userDisp].SendRights = "SendAs"
                        }
                    }
                    
                    foreach ($u in $mbxObj.GrantSendOnBehalfTo) {
                        $userDisp = &$NormalizeId $u $mbx
                        if (-not $usersHash.Contains($userDisp)) {
                            $usersHash[$userDisp] = @{ User = $userDisp; AccessRights = [System.Collections.Generic.List[string]]::new(); SendRights = "SendOnBehalf" }
                        }
                        else {
                            $usersHash[$userDisp].SendRights = "SendOnBehalf"
                        }
                    }

                    $mCount = $usersHash.Count
                    $SyncHash.Window.Dispatcher.Invoke({
                            if ($SyncHash.CurrentTargetMbx -ne $mbx) { return }
                            Write-Host "[$(Get-Date -f HH:mm:ss)] Loaded $mCount users with mailbox/send permissions." -ForegroundColor Gray
                            foreach ($entry in $usersHash.Values) {
                                $rightsStr = ($entry.AccessRights | Select-Object -Unique | Sort-Object) -join ", "
                                if (-not $rightsStr) { $rightsStr = "None" }
                                $SyncHash.GridMbxPerms.Items.Add([PSCustomObject]@{
                                        User = $entry.User; AccessRights = $rightsStr; SendRights = $entry.SendRights
                                    }) | Out-Null
                            }
                            $SyncHash.StatusMbx.Text = ""
                        })
                }

                if ($doCal) {
                    Write-Host "[$(Get-Date -f HH:mm:ss)] DEBUG: Locating calendar folder for $mbx..." -ForegroundColor Gray
                    $rawCalFolders = Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar -ErrorAction Stop

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
                            $calPerms = Get-MailboxFolderPermission -Identity $calPath -ErrorAction Stop | Where-Object { $_.User -notlike "NT AUTHORITY\*" }
                            $cCount = @($calPerms).Count
                    
                            $SyncHash.Window.Dispatcher.Invoke({
                                    if ($SyncHash.CurrentTargetMbx -ne $mbx) { return }
                                    Write-Host "[$(Get-Date -f HH:mm:ss)] Found calendar folder: $($calFolder.Name). Loaded $cCount permissions." -ForegroundColor Gray
                                    foreach ($p in $calPerms) {
                                        $userDisp = if ($p.User.UserPrincipalName) { $p.User.UserPrincipalName }
                                        elseif ($p.User.ToString() -match "Default") { "Default" }
                                        elseif ($p.User.ToString() -match "Anonymous") { "Anonymous" }
                                        elseif ($p.User.ToString() -match "Standard") { "Standard" }
                                        else { $p.User.ToString() -split ":" | Select-Object -Last 1 }
                                        $SyncHash.GridCalPerms.Items.Add([PSCustomObject]@{User = $userDisp; AccessRights = ($p.AccessRights -join ', ') }) | Out-Null
                                    }
                                    $SyncHash.StatusCal.Text = ""
                                })
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
    param([string]$Type, [string]$Action, [hashtable]$Data, [string]$OldRights = "", [string]$OldSendRights = "None")
    
    $mbx = $SyncHash.SelectedMbx
    if (-not $mbx) { return }

    # Update UI to show we are working
    if ($Type -eq "Mailbox") { $SyncHash.StatusMbx.Text = "(Saving...)" } else { $SyncHash.StatusCal.Text = "(Saving...)" }

    $PowerShell = [powershell]::Create().AddScript({
            param($mbx, $type, $action, $data, $SyncHash, $oldRights, $oldSendRights)
            Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
            
            try {
                if ($type -eq "Mailbox") {
                    # 1. Handle Send Rights (Only if changed or new)
                    if ($data.SendRightsChanged -or $action -eq "Add") {
                        if ($action -eq "Edit") {
                            if ($oldSendRights -eq "SendAs") { 
                                Remove-RecipientPermission -Identity $mbx -Trustee $data.User -AccessRights SendAs -Confirm:$false -ErrorAction SilentlyContinue 
                            }
                            if ($oldSendRights -eq "SendOnBehalf") { 
                                Set-Mailbox -Identity $mbx -GrantSendOnBehalfTo @{Remove = $data.User } -ErrorAction SilentlyContinue 
                            }
                        }
                        
                        if ($data.SendRights -eq "SendAs") {
                            Add-RecipientPermission -Identity $mbx -Trustee $data.User -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                        }
                        elseif ($data.SendRights -eq "SendOnBehalf") {
                            Set-Mailbox -Identity $mbx -GrantSendOnBehalfTo @{Add = $data.User } -ErrorAction Stop
                        }
                    }

                    # 2. Handle Access Rights (Only if changed or Automapping explicitly set)
                    if ($data.RightsChanged -or $data.HasAutoSelection -or $action -eq "Add") {
                        if ($action -eq "Edit" -and $oldRights -and $oldRights -ne "None") {
                            Remove-MailboxPermission -Identity $mbx -User $data.User -AccessRights ($oldRights -split ',' | ForEach-Object { $_.Trim() }) -InheritanceType All -Confirm:$false -ErrorAction Stop
                        }
                        
                        if ($data.Rights -and $data.Rights -ne "None") {
                            $mbxParams = @{
                                Identity        = $mbx
                                User            = $data.User
                                AccessRights    = @($data.Rights)
                                InheritanceType = "All"
                                ErrorAction     = "Stop"
                            }
                            # Automapping is only valid and allowed for FullAccess
                            if ($data.Rights -eq "FullAccess") {
                                $mbxParams["AutoMapping"] = if ($data.HasAutoSelection) { $data.Automapping } else { $true }
                            }
                            Add-MailboxPermission @mbxParams
                        }
                    }
                }
                else {
                    # Calendar
                    $raw = Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar
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
                            Show-MessageDialog -Title "Connection Lost" -Message "Exchange Online connection lost (Closed/Broken). Please reconnect." | Out-Null
                            $SyncHash.StatusMbx.Text = ""
                            $SyncHash.StatusCal.Text = ""
                        })
                }
                else {
                    $SyncHash.Window.Dispatcher.Invoke({
                            Write-Host "[$(Get-Date -f HH:mm:ss)] Update Error: $err" -ForegroundColor Red
                            [System.Windows.MessageBox]::Show("Error updating permissions:`n$err")
                            Show-MessageDialog -Title "Update Error" -Message "Error updating permissions:`n$err" | Out-Null
                            $SyncHash.StatusMbx.Text = "" # Clear status regardless
                            $SyncHash.StatusCal.Text = "" # Clear status regardless
                        })
                }
            }
        }).AddArgument($mbx).AddArgument($Type).AddArgument($Action).AddArgument($Data).AddArgument($SyncHash).AddArgument($OldRights).AddArgument($OldSendRights)

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
                    Remove-RecipientPermission -Identity $mbx -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction SilentlyContinue
                    Set-Mailbox -Identity $mbx -GrantSendOnBehalfTo @{Remove = $user } -ErrorAction SilentlyContinue

                    if ($accessRights -and $accessRights -ne "None") {
                        $rightsToRemove = $accessRights -split ',' | ForEach-Object { $_.Trim() }
                        Remove-MailboxPermission -Identity $mbx -User $user -AccessRights $rightsToRemove -InheritanceType All -Confirm:$false -ErrorAction Stop
                    }
                }
                else {
                    # Calendar
                    $raw = Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar
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
                            Show-MessageDialog -Title "Connection Lost" -Message "Exchange Online connection lost (Closed/Broken). Please reconnect." | Out-Null
                            $SyncHash.StatusMbx.Text = ""
                            $SyncHash.StatusCal.Text = ""
                        })
                }
                else {
                    $SyncHash.Window.Dispatcher.Invoke({
                            Write-Host "[$(Get-Date -f HH:mm:ss)] Remove Error: $err" -ForegroundColor Red
                            [System.Windows.MessageBox]::Show("Error removing permissions:`n$err")
                            Show-MessageDialog -Title "Removal Error" -Message "Error removing permissions:`n$err" | Out-Null
                            $SyncHash.StatusMbx.Text = "" # Clear status regardless
                            $SyncHash.StatusCal.Text = "" # Clear status regardless
                        })
                }
            }
        }).AddArgument($mbx).AddArgument($Type).AddArgument($User).AddArgument($AccessRights).AddArgument($SyncHash)
    
    $PowerShell.RunspacePool = $Pool
    $PowerShell.BeginInvoke() | Out-Null
}

# --- Permission Button Events ---
$mailboxRights = @("None", "ChangeOwner", "ChangePermission", "DeleteItem", "ExternalAccount", "FullAccess", "ReadPermission") | Sort-Object
$calendarRoles = @("None", "AvailabilityOnly", "LimitedDetails", "Author", "Contributor", "Editor", "NonEditingAuthor", "Owner", "PublishingAuthor", "PublishingEditor", "Reviewer") | Sort-Object

$BtnAddMbx.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnAddMbx clicked." -ForegroundColor Magenta
        $userList = $SyncHash.AddressCache
        $res = Show-PermissionDialog -Title "Add Mailbox Permission" -Options $mailboxRights -Label "Access Rights:" -ShowSendRights $true -UserList $userList -UserEditable $true
        if ($res) { Update-PermissionAsync -Type "Mailbox" -Action "Add" -Data $res } else { Write-Host "[$(Get-Date -f HH:mm:ss)] Show-PermissionDialog for AddMbx returned null." -ForegroundColor Yellow }
    })

# --- Event: Detect Re-click on already selected Mailbox to refresh ---
$ListMailboxes.Add_PreviewMouseLeftButtonDown({
        param($s, $e)
        $dep = $e.OriginalSource
        # Walk up the visual tree to find the ListViewItem container
        while ($null -ne $dep -and $dep.GetType().Name -ne "ListViewItem") {
            $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep)
        }
        if ($null -ne $dep -and $dep.IsSelected) {
            # Item is already selected, so SelectionChanged won't fire. Trigger refresh manually.
            $selectedItem = $ListMailboxes.SelectedItem
            if ($null -ne $selectedItem) {
                & $SyncHash.GetPermissionsAsync -Mailbox $selectedItem.Address -FetchMbx $true -FetchCal $true
            }
        }
    })

$BtnEditMbx.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnEditMbx clicked." -ForegroundColor Magenta
        $sel = $GridMbxPerms.SelectedItem
        if (-not $sel) { Show-MessageDialog -Title "No User Selected" -Message "Please select a user from the list." | Out-Null; return }
        # Extract the first access right for display in the dropdown
        $currentRight = ($sel.AccessRights -split "," | Select-Object -First 1 | ForEach-Object { $_.Trim() })
        $userList = $SyncHash.AddressCache
        $res = Show-PermissionDialog -Title "Edit Mailbox Permission" -User $sel.User -Options $mailboxRights -CurrentOption $currentRight -CurrentAutomapping "" -UserEditable $false -Label "Access Rights:" -ShowSendRights $true -CurrentSendRights $sel.SendRights -UserList $userList
        if ($res) { Update-PermissionAsync -Type "Mailbox" -Action "Edit" -Data $res -OldRights $sel.AccessRights -OldSendRights $sel.SendRights } else { Write-Host "[$(Get-Date -f HH:mm:ss)] Show-PermissionDialog for EditMbx returned null." -ForegroundColor Yellow }
    })

$BtnRemoveMbx.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnRemoveMbx clicked." -ForegroundColor Magenta
        $sel = $GridMbxPerms.SelectedItem
        if (-not $sel) { Show-MessageDialog -Title "No User Selected" -Message "Please select a user from the list." | Out-Null; return }
        if (Show-ConfirmDialog -Title "Confirm Removal" -Message "Are you sure you want to remove the selected mailbox permission for $($sel.User) from $($SyncHash.SelectedMbx)?") {
            Remove-PermissionAsync -Type "Mailbox" -User $sel.User -AccessRights $sel.AccessRights
        }
    })

$BtnAddCal.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnAddCal clicked." -ForegroundColor Magenta
        $userList = $SyncHash.AddressCache
        $res = Show-PermissionDialog -Title "Add Calendar Permission" -Options $calendarRoles -Label "Access Roles:" -ShowAutomapping $false -UserList $userList -UserEditable $true
        if ($res) { Update-PermissionAsync -Type "Calendar" -Action "Add" -Data $res } else { Write-Host "[$(Get-Date -f HH:mm:ss)] Show-PermissionDialog for AddCal returned null." -ForegroundColor Yellow }
    })

$BtnEditCal.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnEditCal clicked." -ForegroundColor Magenta
        $sel = $GridCalPerms.SelectedItem
        if (-not $sel) { Show-MessageDialog -Title "No User Selected" -Message "Please select a user from the list." | Out-Null; return }
        $currentRole = $sel.AccessRights -split "," | Select-Object -First 1 | ForEach-Object { $_.Trim() }
        $userList = $SyncHash.AddressCache
        $res = Show-PermissionDialog -Title "Edit Calendar Permission" -User $sel.User -Options $calendarRoles -CurrentOption $currentRole -UserEditable $false -Label "Access Roles:" -ShowAutomapping $false -UserList $userList
        if ($res) { Update-PermissionAsync -Type "Calendar" -Action "Edit" -Data $res } else { Write-Host "[$(Get-Date -f HH:mm:ss)] Show-PermissionDialog for EditCal returned null." -ForegroundColor Yellow }
    })

$BtnRemoveCal.Add_Click({
        Write-Host "[$(Get-Date -f HH:mm:ss)] BtnRemoveCal clicked." -ForegroundColor Magenta
        $sel = $GridCalPerms.SelectedItem
        if (-not $sel) { Show-MessageDialog -Title "No User Selected" -Message "Please select a user from the list." | Out-Null; return }
        if (Show-ConfirmDialog -Title "Confirm Removal" -Message "Are you sure you want to remove the selected calendar permission for $($sel.User) from $($SyncHash.SelectedMbx)?") {
            $currentRole = $sel.AccessRights -split "," | Select-Object -First 1 | ForEach-Object { $_.Trim() }
            Remove-PermissionAsync -Type "Calendar" -User $sel.User -AccessRights $currentRole
        }
    })

# --- Event: Connectivity & Persistence Hooks ---
$ChkDelegated.Add_Checked({
        $TxtDelegatedOrg.IsEnabled = $true
        $AdminUserPanel.Visibility = "Collapsed"
        $DelegatedWarning.Visibility = "Visible"
        Save-ManagerSettings
    })

$ChkDelegated.Add_Unchecked({
        $TxtDelegatedOrg.IsEnabled = $false
        $AdminUserPanel.Visibility = "Visible"
        $DelegatedWarning.Visibility = "Collapsed"
        Save-ManagerSettings
    })
$TxtDelegatedOrg.Add_TextChanged({ Save-ManagerSettings })
$TxtUserUPN.Add_TextChanged({ Save-ManagerSettings })

# --- Event: Toggle between Search and UPN Input ---
$ChkFetchAllMailboxes.Add_Checked({
        Save-ManagerSettings
        $SearchBarPanel.Visibility = "Visible"
        $UPNInputPanel.Visibility = "Collapsed"
    })

$ChkFetchAllMailboxes.Add_Unchecked({
        Save-ManagerSettings
        $SearchBarPanel.Visibility = "Collapsed"
        $UPNInputPanel.Visibility = "Visible"
    })

# Initialize visibility based on checkbox state
if ($ChkFetchAllMailboxes.IsChecked) {
    $SearchBarPanel.Visibility = "Visible"
    $UPNInputPanel.Visibility = "Collapsed"
}

# Initialize Admin-User textbox state based on ChkDelegated
if ($ChkDelegated.IsChecked) {
    $AdminUserPanel.Visibility = "Collapsed"
    $DelegatedWarning.Visibility = "Visible"
}
else {
    $AdminUserPanel.Visibility = "Visible"
    $DelegatedWarning.Visibility = "Collapsed"
}

if (-not $ChkFetchAllMailboxes.IsChecked) {
    $SearchBarPanel.Visibility = "Collapsed"
    $UPNInputPanel.Visibility = "Visible"
}

# --- Event: Connect & Fetch Mailboxes (ASYNC) ---
$BtnConnect.Add_Click({
        # Handle Disconnection if already connected
        if ($BtnConnect.Background.ToString() -eq "#FF28A745") {
            # Success Green
            Disconnect-ExchangeOnline -Confirm:$false
            $BtnConnect.Content = "Connect to Exchange"
            $BtnConnect.Background = "#007ACC" # Blue
            $SyncHash.AddressCache.Clear()
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

        $Delegated = [bool]$ChkDelegated.IsChecked
        $Org = $TxtDelegatedOrg.Text.Trim()
        $UserUPN = $TxtUserUPN.Text.Trim()
        $FetchAll = [bool]$ChkFetchAllMailboxes.IsChecked
        Write-Host "[$(Get-Date -f HH:mm:ss)] Attempting to connect to Exchange (Delegated: $Delegated, FetchAll: $FetchAll)..." -ForegroundColor Cyan

        Save-ManagerSettings

        # Store connection info in SyncHash for background runspaces to reuse
        $SyncHash.ConnectDelegated = $Delegated
        $SyncHash.ConnectOrg = $Org
        $SyncHash.TargetUPN = $UserUPN
        $SyncHash.FetchAllMailboxes = $FetchAll

        # Create and show a small authentication status window
        $Window.Dispatcher.Invoke({
                $overlayXml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Height="100" Width="320" WindowStyle="None" AllowsTransparency="True" ShowInTaskbar="False"
        Background="Transparent" WindowStartupLocation="CenterScreen">
    <Border Background="#2D2D30" CornerRadius="8" BorderBrush="#007ACC" BorderThickness="2" Cursor="SizeAll">
        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
            <TextBlock Text="Authentication in Progress" Foreground="#007ACC" FontSize="14" FontWeight="Bold" Margin="0,0,0,5" HorizontalAlignment="Center"/>
            <TextBlock Text="Please check for the login window..." Foreground="#E0E0E0" FontSize="11" HorizontalAlignment="Center"/>
        </StackPanel>
    </Border>
</Window>
"@
                $overlay = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$overlayXml)))
                $overlay.Add_MouseLeftButtonDown({ $overlay.DragMove() })
                $overlay.Show()
                $SyncHash.AuthOverlay = $overlay
            })

        # Minimize window so login prompt is visible
        $Window.Dispatcher.Invoke([Action] { $Window.WindowState = "Minimized" })

        $PowerShell = [powershell]::Create().AddScript({
                param($Delegated, $Org, $UserUPN, $SyncHash, $FetchAll)
                Import-Module ExchangeOnlineManagement
        
                try {
                    try {
                        $connParams = @{ ShowProgress = $false; ErrorAction = "Stop" }

                        if ($Delegated -and (-not [string]::IsNullOrWhiteSpace($Org))) { 
                            $connParams["DelegatedOrganization"] = $Org 
                        }
                        else {
                            if (-not [string]::IsNullOrWhiteSpace($UserUPN)) { 
                                $connParams["UserPrincipalName"] = $UserUPN 
                            }
                        }

                        Connect-ExchangeOnline @connParams
                    }
                    finally {
                        $SyncHash.Window.Dispatcher.Invoke({
                                if ($SyncHash.AuthOverlay) { $SyncHash.AuthOverlay.Close(); $SyncHash.AuthOverlay = $null }
                                $SyncHash.Window.WindowState = "Normal"
                                $SyncHash.Window.Activate() | Out-Null
                            })
                    }

                    $SyncHash.Window.Dispatcher.Invoke({
                            Write-Host "[$(Get-Date -f HH:mm:ss)] Connected." -ForegroundColor Green
                        })

                    # Capture the actual UPN used to login
                    $info = Get-ConnectionInformation | Select-Object -First 1
                    $SyncHash.ConnectedUser = $info.UserPrincipalName
                    Write-Host "[$(Get-Date -f HH:mm:ss)] Session Identity: $($SyncHash.ConnectedUser)" -ForegroundColor Gray

                    # Update UI to Connected
                    $SyncHash.Window.Dispatcher.Invoke({
                            # Create multi-line content for the button
                            $sp = New-Object System.Windows.Controls.StackPanel -Property @{ VerticalAlignment = "Center" }
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
                        })

                    # Fetch mailboxes only if FetchAll is true
                    if ($FetchAll) {
                        $SyncHash.Window.Dispatcher.Invoke({
                                $SyncHash.StatusMailboxes.Text = "(Fetching list...)"
                                $SyncHash.ProgressMailboxesHeader.Visibility = "Visible"
                                $SyncHash.ProgressMailboxesHeader.IsIndeterminate = $true
                                $SyncHash.ProgressMailboxesText.Visibility = "Collapsed"
                            })

                        Write-Host "[$(Get-Date -f HH:mm:ss)] Fetching all mailboxes..." -ForegroundColor Green
                        
                        # Fetch all Mailboxes
                        $allMailboxes = @(Get-EXOMailbox -ResultSize Unlimited | Select-Object PrimarySmtpAddress, RecipientTypeDetails)
                        
                        $SyncHash.Window.Dispatcher.Invoke({
                                $SyncHash.ProgressMailboxesHeader.IsIndeterminate = $false
                                $SyncHash.ProgressMailboxesText.Visibility = "Visible"
                                $SyncHash.StatusMailboxes.Text = "(Processing results...)"
                            })

                        $totalCount = $allMailboxes.Count
                        $currentIndex = 0

                        # Filter and categorize mailboxes
                        $filteredMailboxes = [System.Collections.ArrayList]::new()
                        $addressCache = [System.Collections.Generic.List[string]]::new()
                        foreach ($mbx in $allMailboxes) {
                            $currentIndex++
                            $progressPercent = [math]::Min(100, [math]::Round(($currentIndex / $totalCount) * 100))
                            
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
                                Default { 99 }
                            }

                            $addressCache.Add($mbx.PrimarySmtpAddress)

                            $null = $filteredMailboxes.Add([PSCustomObject]@{
                                    Address      = $mbx.PrimarySmtpAddress;
                                    Type         = $groupName;
                                    OriginalType = $mbx.RecipientTypeDetails;
                                    SortOrder    = $sortOrder
                                }
                            )
                            
                            # Update progress UI
                            if ($currentIndex % [math]::Max(1, [math]::Floor($totalCount / 20)) -eq 0 -or $currentIndex -eq $totalCount) {
                                $SyncHash.Window.Dispatcher.InvokeAsync({
                                        $SyncHash.ProgressMailboxesHeader.Value = $progressPercent
                                        $SyncHash.ProgressMailboxesText.Text = "$progressPercent%"
                                    }) | Out-Null
                            }
                        }

                        # Sort mailboxes alphabetically by Type and then by Address
                        $sortedMailboxes = $filteredMailboxes | Sort-Object SortOrder, Address

                        $SyncHash.Window.Dispatcher.Invoke({
                                Write-Host "[$(Get-Date -f HH:mm:ss)] Loaded $($sortedMailboxes.Count) mailboxes into sections." -ForegroundColor Gray

                                # Update collections and address cache for dialogs
                                $newCollection = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
                                foreach ($item in $sortedMailboxes) { $newCollection.Add($item) }
                                $SyncHash.AllMailboxes = $newCollection
                                $SyncHash.AddressCache = $addressCache

                                # Re-setup the view for the new collection instance
                                $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($SyncHash.AllMailboxes)
                                $view.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription("Type")))
                                $view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription("SortOrder", [System.ComponentModel.ListSortDirection]::Ascending)))
                                $view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription("Address", [System.ComponentModel.ListSortDirection]::Ascending)))

                                # Re-bind the ItemsSource
                                $SyncHash.ListMailboxes.ItemsSource = $view
                                $view.Refresh()

                                $SyncHash.StatusMailboxes.Text = ""
                                $SyncHash.ProgressMailboxesHeader.Visibility = "Collapsed"
                                $SyncHash.ProgressMailboxesText.Visibility = "Collapsed"
                            })
                    }
                    else {
                        $SyncHash.Window.Dispatcher.Invoke({
                                $SyncHash.StatusMailboxes.Text = ""
                                $SyncHash.ProgressMailboxesHeader.Visibility = "Collapsed"
                                $SyncHash.ProgressMailboxesText.Visibility = "Collapsed"
                            })
                    }
                }
                catch {
                    $err = $_.Exception.Message
                    $SyncHash.Window.Dispatcher.Invoke({
                            if ($SyncHash.AuthOverlay) { $SyncHash.AuthOverlay.Close(); $SyncHash.AuthOverlay = $null }
                            $SyncHash.Window.WindowState = "Normal"
                            $null = $SyncHash.Window.Activate()
                            Show-MessageDialog -Title "Connection Error" -Message "CONNECTION ERROR: $err" | Out-Null
                            $SyncHash.BtnConnect.Content = "Connection Failed"
                            $SyncHash.BtnConnect.Background = "#FF6B68" # Red
                            $SyncHash.BtnConnect.Foreground = "White"
                            $SyncHash.StatusMailboxes.Text = "(Error)"
                            $SyncHash.ProgressMailboxesHeader.Visibility = "Collapsed"
                            $SyncHash.ProgressMailboxesText.Visibility = "Collapsed"
                        })
                }
                finally {
                    $SyncHash.Window.Dispatcher.Invoke({
                            $SyncHash.BtnConnect.IsHitTestVisible = $true
                        })
                }
            }).AddArgument($Delegated).AddArgument($Org).AddArgument($UserUPN).AddArgument($SyncHash).AddArgument($FetchAll)

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

# --- Event: Fetch specific UPN ---
$BtnFetchUPN.Add_Click({
        $upn = $TxtFetchUPN.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($upn)) {
            Show-MessageDialog -Title "Missing UPN" -Message "Please enter a UPN (e.g. user@domain.com)" | Out-Null
            return
        }

        $BtnFetchUPN.IsHitTestVisible = $false
        $BtnFetchUPN.Content = "Fetching..."
        $StatusMailboxes.Text = "(Fetching...)"
        $ProgressMailboxesHeader.Visibility = "Collapsed"
        $ProgressMailboxesText.Visibility = "Collapsed"

        $PowerShell = [powershell]::Create().AddScript({
                param($upn, $SyncHash)
                Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

                try {
                    Write-Host "[$(Get-Date -f HH:mm:ss)] Fetching mailbox: $upn" -ForegroundColor Green
                    $mbx = Get-EXOMailbox -Identity $upn -ErrorAction Stop | Select-Object PrimarySmtpAddress, RecipientTypeDetails

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
                        Default { 99 }
                    }

                    $mailboxObj = [PSCustomObject]@{
                        Address      = $mbx.PrimarySmtpAddress
                        Type         = $groupName
                        OriginalType = $mbx.RecipientTypeDetails
                        SortOrder    = $sortOrder
                    }

                    $SyncHash.Window.Dispatcher.Invoke({
                            # Check if this UPN already exists
                            $existing = $SyncHash.AllMailboxes | Where-Object { $_.Address -eq $mailboxObj.Address }
                            
                            if (-not $existing) {
                                $SyncHash.AllMailboxes.Add($mailboxObj)
                                $SyncHash.AddressCache.Add($mailboxObj.Address)
                                $SyncHash.ListMailboxes.SelectedItem = $mailboxObj
                                Write-Host "[$(Get-Date -f HH:mm:ss)] Added mailbox: $($mailboxObj.Address)" -ForegroundColor Green
                            }
                            else {
                                $SyncHash.ListMailboxes.SelectedItem = $existing
                                Write-Host "[$(Get-Date -f HH:mm:ss)] Mailbox already in list: $($mailboxObj.Address)" -ForegroundColor Yellow
                            }

                            # Refresh the view
                            $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($SyncHash.AllMailboxes)
                            $view.Refresh()

                            $SyncHash.StatusMailboxes.Text = ""
                            $SyncHash.BtnFetchUPN.Content = "Get"
                            $SyncHash.BtnFetchUPN.IsHitTestVisible = $true
                            $SyncHash.TxtFetchUPN.Text = ""
                        })
                }
                catch {
                    $err = $_.Exception.Message
                    $SyncHash.Window.Dispatcher.Invoke({
                            Show-MessageDialog -Title "Fetch Error" -Message "Error fetching mailbox:`n$err" | Out-Null
                            $SyncHash.StatusMailboxes.Text = ""
                            $SyncHash.BtnFetchUPN.Content = "Get"
                            $SyncHash.BtnFetchUPN.IsHitTestVisible = $true
                        })
                }
            }).AddArgument($upn).AddArgument($SyncHash)

        $PowerShell.RunspacePool = $Pool
        $PowerShell.BeginInvoke() | Out-Null
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