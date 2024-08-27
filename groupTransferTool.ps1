Import-Module ActiveDirectory

[console]::WindowWidth = 40; 
[console]::WindowHeight = 20; 
[console]::BufferWidth = [console]::WindowWidth

$fileName = Get-Date -Format "dd.MM.yyyy"
$fileName = "$env:LOCALAPPDATA\TroubleshootingTool\Logs\" + $fileName + ".log"


$windowHeight = 600
$windowWidth = 800

$minWindowHeight = 600
$minWindowWidth = 600

#Whether the logs will be output to the powershell window aswell as the file
$LogToConsole = $True

$Global:RemoveMode = $False

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
            <RowDefinition Height="30"/>
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

        <Grid Grid.Column="2" Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Button Name="ClearUsersButton" Content="Clear All Users" Margin="2,2,2,2" VerticalContentAlignment="Center" Grid.Column="0"/>
            <Button Name="RemoveUsersButton" Content="Remove Selected Users" Margin="2,2,2,2" VerticalContentAlignment="Center" Grid.Column="1"/>
        </Grid>
        <Grid Grid.Column="0" Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Button Name="SelectAllButton" Content="Select All Groups" Margin="2,2,2,2" VerticalContentAlignment="Center" Grid.Column="0"/>
            <Button Name="ToggleAddRemoveButton" Content="Toggle Mode" Margin="2,2,2,2" VerticalContentAlignment="Center" Grid.Column="1"/>
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
    $userID = $WPFTargetUser.Text
    try {
        $user = Get-ADUser -Identity $userID -Properties *
        $userID = $user.SamAccountName
        $WPFTargetUserLabel.Content = $userID + "'s Groups:"
    }
    catch {
        $WPFTargetUserLabel.Content = "Couldnt find user: `"$($userID)`""
    }

    $WPFTargetUserGroups.Items.Clear()
    $groups = Get-ADPrincipalGroupMembership $userID | Select-Object name, SamAccountName
    $groups = $groups | Sort-Object -Property name # Sorting the groups by name so they appear the same as in AD

    foreach ($group in $groups) { $WPFTargetUserGroups.Items.Add($group.name) }
}

Function Get-Recipients() {
    $userID = $WPFRecipientUser.Text
    try {
        $user = Get-ADUser -Identity $userID -Properties *
        $userID = $user.SamAccountName # Get the accounts proper capitalization
        if ($WPFRecipientUsers.Items.Contains($userID)) {
            $WPFRecipientUserLabel.Content = "User is already added"
        }
        else {
            $WPFRecipientUsers.Items.Add($userID)
            $WPFRecipientUserLabel.Content = "Users to receive groups"
        }
    }
    catch {
        $WPFRecipientUserLabel.Content = "Couldnt find user: `"$($userID)`""
    }
}

$WPFSearchTargetButton.Add_Click({ Get-TargetGroups })
$WPFTargetUser.Add_KeyDown({
        if ($_.Key -eq "Enter") {
            Get-TargetGroups
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

Write-Log -Log "Group Transfer Tool Opened" -IsNewLine $True

$Form.ShowDialog() | Out-Null