[console]::WindowWidth = 40; 
[console]::WindowHeight = 20; 
[console]::BufferWidth = [console]::WindowWidth

Import-Module ActiveDirectory

$fileName = Get-Date -Format "dd.MM.yyyy"
$fileName = "$env:LOCALAPPDATA\TroubleshootingTool\Logs\" + $fileName + ".log"


$windowHeight = 600
$windowWidth = 800

$minWindowHeight = 600
$minWindowWidth = 600

# Whether the logs will be output to the powershell window aswell as the file
$LogToConsole = $True
# Whether the logs will show when the tool is opened and closed
$LogOpening = $False 

$Global:RemoveMode = $False
$Global:TargetUser

#region XAML UI
#===========================================================================
# Load XAML UI
#===========================================================================
Add-Type -AssemblyName PresentationFramework

$inputXML = @"
<Window Name="window" x:Class="WpfApp2.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp2"
        Title="Groups Transfer Tool" Background="#FFD5DBDE" MinHeight="$($minWindowHeight)" MinWidth="$($minWindowWidth)" Height="$($windowHeight)" Width="$($windowWidth)">
	<Grid Name="grid" Margin="5,5,5,5">
		<Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="140"/>
            <ColumnDefinition Width="*"/>
		</Grid.ColumnDefinitions>
		<Grid.RowDefinitions>
            <RowDefinition Height="32"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="60"/>
		</Grid.RowDefinitions>
        <TextBox Name="TargetUser" Margin="2,2,64,2" Padding="5,0,0,0" Height="28" Grid.Column="0" Grid.Row="0" VerticalAlignment="Top" VerticalContentAlignment="Center"/>
        <Button Name="SearchTargetButton" Margin="2,2,2,2" Content="Search" HorizontalAlignment="Right" Width="60" Height="28" Grid.Column="0" Grid.Row="0" VerticalAlignment="Top"/>

        <Label Name="TargetUserLabel" Content="Target User" Margin="2,2,2,0" Padding="5,0,0,0" Height="28" Grid.Column="0" Grid.Row="1" VerticalAlignment="Top" VerticalContentAlignment="Center" FontSize="15"/>

		<ListBox Background="#FFFFFF" Name="TargetUserGroups" FocusManager.IsFocusScope="True" Margin="2,30,2,2" SelectionMode="Extended" Grid.Column="0" Grid.Row="1" MinHeight="200" MinWidth="100">
        </ListBox>

        <StackPanel Grid.Column="1" Grid.Row="1" VerticalAlignment="Center">
            <Button Name="TransferGroupsButton" Content="Add Groups" Margin="2,2,2,2" Height="50" FontSize="16"/>
            <TextBlock Name="InfoText" Text="" Margin="2,2,2,2" FontSize="15" TextWrapping="WrapWithOverflow" TextAlignment="Center"/>
        </StackPanel>

        <TextBox Name="RecipientUser" Margin="2,2,64,2" Padding="5,0,0,0" Height="28" Grid.Column="2" Grid.Row="0" VerticalAlignment="Top" VerticalContentAlignment="Center"/>
        <Button Name="SearchRecipentButton" Margin="2,2,2,2" Content="Search" HorizontalAlignment="Right" Width="60" Height="28" Grid.Column="2" Grid.Row="0" VerticalAlignment="Top"/>

        <Label Name="RecipientUserLabel" Content="Users to receive groups" Margin="2,2,2,0" Padding="5,0,0,0" Height="28" Grid.Column="2" Grid.Row="1" VerticalAlignment="Top" VerticalContentAlignment="Center" FontSize="15"/>

        <ListBox Background="#FFFFFF" Name="RecipientUsers" Margin="2,30,2,2" SelectionMode="Extended" Grid.Column="2" Grid.Row="1" MinHeight="200" MinWidth="100">
        </ListBox>

        <Grid Grid.Column="0" Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Button Name="SelectAllButton" Content="Select All Groups" Margin="2,2,2,2" VerticalContentAlignment="Center" Grid.Column="0"/>
            <Button Name="ToggleAddRemoveButton" Content="Toggle Mode" Margin="2,2,2,2" VerticalContentAlignment="Center" Grid.Column="1"/>
            <Button Name="AddGroupButton" Content="Add Group To User" Margin="2,2,2,2" VerticalContentAlignment="Center" IsEnabled="False" Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2"/>
        </Grid>
        <Grid Grid.Column="2" Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Button Name="ClearUsersButton" Content="Clear All Users" Margin="2,2,2,2" VerticalContentAlignment="Center" Grid.Column="0"/>
            <Button Name="RemoveUsersButton" Content="Remove Selected Users" Margin="2,2,2,2" VerticalContentAlignment="Center" Grid.Column="1"/>
            <Button Name="AddUsersButton" Content="Add Users From File" Margin="2,2,2,2" VerticalContentAlignment="Center" Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2"/>
        </Grid>
	</Grid>
</Window>
"@

$inputXML = $inputXML -replace '^<Win.*', '<Window'

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try { $Form = [Windows.Markup.XamlReader]::Load( $reader ) }
catch { 
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed. (Or the XML is incorrectly formatted)" 
}

$xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) }

Function Get-FormVariables {
    if ($global:ReadmeDisplay -ne $true) { $global:ReadmeDisplay = $true }
}

Get-FormVariables
#endregion


Function Write-Log {
    param ( 
        $Log,
        $IsNewLine = $False 
    )

    $Log = "[$(Get-Date -Format "HH:mm:ss")] $($Log)" # Adding time stamps to the logs
    if ($IsNewLine -eq $True -and [System.IO.File]::Exists($fileName)) { $Log = " `n$($Log)" }

    Add-Content -Path $fileName -Value $Log
    if ($LogToConsole -eq $True) { Write-Host $Log }
    
}

Function Get-TargetGroups() {
    param(
        $Target = $Global:TargetUser
    )
    try {
        $User = Get-ADUser -Identity $Target -Properties DisplayName | Select-Object DisplayName, SamAccountName
        $UserID = $User.SamAccountName
        $WPFTargetUserLabel.Content = "$($User.DisplayName) ($UserID)'s Groups:"
        $WPFAddGroupButton.IsEnabled = $True
        $Global:TargetUser = $UserID
        $WPFAddGroupButton.Content = "Add Group to $($UserID)"

        $WPFTargetUserGroups.Items.Clear()
        $groups = Get-ADPrincipalGroupMembership $UserID | Select-Object name, SamAccountName
        $groups = $groups | Sort-Object -Property name # Sorting the groups by name so they appear the same as in AD

        foreach ($group in $groups) { $WPFTargetUserGroups.Items.Add($group.name) }
    }
    catch {
        $WPFTargetUserLabel.Content = "Couldnt find user: `"$($Target)`""
    }
}

Function Get-Recipients() {
    param(
        $UserIn = $WPFRecipientUser.Text
    )
    try {
        $UserID = (Get-ADUser -Identity $UserIn).SamAccountName # Get the accounts proper capitalization
        if ($WPFRecipientUsers.Items.Contains($UserID)) {
            $WPFRecipientUserLabel.Content = "$($UserID) is already added"
        }
        else {
            $WPFRecipientUsers.Items.Add($UserID)
            $WPFRecipientUserLabel.Content = "Users to receive groups"
        }
    }
    catch {
        $WPFRecipientUserLabel.Content = "Couldnt find user: `"$($userID)`""
    }
}

$WPFSearchTargetButton.Add_Click({ 
        Get-TargetGroups -Target $WPFTargetUser.Text
    })
$WPFTargetUser.Add_KeyDown({
        if ($_.Key -eq "Enter") {
            Get-TargetGroups -Target $WPFTargetUser.Text
            $WPFTargetUser.Text = ""
        }
    })

$WPFSearchRecipentButton.Add_Click({ Get-Recipients })
$WPFRecipientUser.Add_KeyDown({
        if ($_.Key -eq "Enter") {
            Get-Recipients
            $WPFRecipientUser.Text = ""
        }
    })

$WPFSelectAllButton.Add_Click({
        foreach ($item in $WPFTargetUserGroups.Items) {
            $WPFTargetUserGroups.SelectedItems.Add($item);
        }
    })

$WPFToggleAddRemoveButton.Add_Click({
        if ($Global:RemoveMode -eq $False) {
            $WPFTransferGroupsButton.Content = "Remove Groups"
            $Global:RemoveMode = $True
        }
        else {
            $WPFTransferGroupsButton.Content = "Add Groups"
            $Global:RemoveMode = $False
        }
    })

$WPFClearUsersButton.Add_Click({
        $WPFRecipientUsers.Items.Clear()
    })

$WPFRemoveUsersButton.Add_Click({
        $temp = $WPFRecipientUsers.Items | Where-Object { $WPFRecipientUsers.SelectedItems -notcontains $_ }
        $WPFRecipientUsers.Items.Clear()
        foreach ($item in $temp) {
            $WPFRecipientUsers.Items.Add($item)
        }
    })

$WPFAddGroupButton.Add_Click({
        $rawXML = @"
<Window Name="window" x:Class="WpfApp2.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp2"
        Title="Add Group" Background="#FFD5DBDE" ResizeMode="NoResize" Height="125" Width="355">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="30"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Label Name="GroupNameLabel" Margin="3,3,3,0" Content="Enter Group Name:" Grid.Row="0" Grid.ColumnSpan="4"/>
        <TextBox Name="GroupName" Margin="5,0,5,0" VerticalContentAlignment="Center" Grid.Row="1" Grid.ColumnSpan="4"/>
        <Button Name="GroupConfirmButton" Content="OK" Margin="5,5,5,5" Grid.Row="2" Grid.Column="2"/>
        <Button Name="GroupCancelButton" Content="Cancel" Margin="0,5,5,5" Grid.Row="2" Grid.Column="3"/>
    </Grid>
</Window>
"@
        $rawXML = $rawXML -replace '^<Win.*', '<Window'
        
        [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
        [xml]$GroupXAML = $rawXML

        $GroupReader = (New-Object System.Xml.XmlNodeReader $GroupXAML)
        $Window = [Windows.Markup.XamlReader]::Load( $GroupReader )
        $GroupXAML.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name "WPF$($_.Name)" -Value $Window.FindName($_.Name) }
        #endregion

        Function Find-Group() {
            try {
                $ADGroup = Get-ADGroup -Identity $WPFGroupName.Text
                Write-Log "Added $($Global:TargetUser) to $($ADGroup.Name)"

                # Turns out it takes a couple of seconds for AD to update the groups in a user, so this kinda isnt very practical
                Add-ADGroupMember -Identity $ADGroup -Members $Global:TargetUser
                Start-Sleep 5 # Sleeping to hopefully update it properly

                Get-TargetGroups # Update the groups after the new one is added
                $Window.Hide() # Closing the window after a new group is added
            }
            catch {
                $WPFGroupNameLabel.Content = "Couldn't find group: $($WPFGroupName.Text)"
            }
        }

        $WPFGroupConfirmButton.Add_Click({ Find-Group })
        $WPFGroupName.Add_KeyDown({
                if ($_.Key -eq "Enter") { Find-Group }
            })

        $WPFGroupCancelButton.Add_Click({ $Window.Hide() })

        $Window.ShowDialog() | Out-Null

    })

$WPFAddUsersButton.Add_Click({
        $Dialog = [Microsoft.Win32.OpenFileDialog]::new()
        $Dialog.FileName = "In"
        $Dialog.DefaultExt = ".csv"
        $Dialog.Filter = "CSV Files|*.csv"

        $Keywords = @("User", "Users", "Username", "Usernames", "User Name", "User Names", "Account", 
            "Accounts", "LAN", "AccountID", "AccountIDs", "SamAccountName", "SamAccountNames")

        $FoundUsers = $False
        if ($Dialog.ShowDialog()) {
            $UserFile = Import-Csv -Path $Dialog.FileName

            foreach ($Word in $Keywords) {
                if (($UserFile | Get-Member -MemberType NoteProperty).name -contains $Word) {
                    $UserFile.$Word | ForEach-Object { Get-Recipients -UserIn $_ }
                    $FoundUsers = $True
                    break
                }
            }

            if ($FoundUsers -eq $False) {
                [System.Windows.MessageBox]::Show("Could not find any columns in the csv file headed with: `n$($Keywords)", "Unable to find users")
            }
            #$temp = "Users"
            #$UserFile.$temp | ForEach-Object { Get-Recipients -UserIn $_ } 
            #           ^ Trying to find a better way to do this, this requires that the header on the csv column is "Users"
        }
    })

$WPFTransferGroupsButton.Add_Click({
        $groups = $WPFTargetUserGroups.SelectedItems # Get the selected groups
        if ($groups.Count -gt 0 -and $WPFRecipientUsers.Items.Count -gt 0) {
            foreach ($group in $groups) {
                $groupSAN = Get-ADGroup -Filter { Name -eq $group } | Select-Object SamAccountName
                if ($Global:RemoveMode -eq $True) {
                    Remove-ADGroupMember -Identity $groupSAN.SamAccountName -Members $WPFRecipientUsers.Items -Confirm:$False
                    Write-Log -Log "Removed users: $($WPFRecipientUsers.Items) from group $($group) ($($groupSAN.SamAccountName))"

                }
                else {
                    Add-ADGroupMember -Identity $groupSAN.SamAccountName -Members $WPFRecipientUsers.Items
                    Write-Log -Log "Added users: $($WPFRecipientUsers.Items) to group $($group) ($($groupSAN.SamAccountName))"
                }
            }
            # This whole following section is not necessary at all but I wrote it anyway
            # It literally just makes sure if there is 1 group that the message says group instead of groups, same with people.
            $temp = ""
            if ($WPFRecipientUsers.Items.Count -eq 1) { $temp = $temp + "$($WPFRecipientUsers.Items.Count) user " }
            else { $temp = $temp + "$($WPFRecipientUsers.Items.Count) users " }

            if ($Global:RemoveMode -eq $True) { $temp = $temp + "removed from " }
            else { $temp = $temp + "added to " }
    
            if ($groups.Count -eq 1) { $temp = $temp + "$($groups.Count) group." }
            else { $temp = $temp + "$($groups.Count) groups." }
    
            $WPFInfoText.Text = $temp
        }
        else {
            $WPFInfoText.Text = "No users/groups selected."
        }
    })

Write-Host "Using Log file at $($fileName)"
[Console]::SetCursorPosition(0, 0)
[Console]::SetCursorPosition(0, 3)

if ($LogOpening) { Write-Log -Log "Group Transfer Tool Opened" -IsNewLine $True }

$Form.ShowDialog() | Out-Null