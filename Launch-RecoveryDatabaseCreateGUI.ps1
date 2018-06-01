<#
.NOTES
With the help of            :   Jim Moyle @jimmoyle
How-To GUI From Jim Moyle   :   https://github.com/JimMoyle/GUIDemo

#>
# region Global Variables
$global:GUIversion = "1.0"
#Storing paths, values, server names, into variables for more flexible manipulation
#EDB and LOG Folder paths:
$global:dafaultOriginalEDBFilePath = "H:\RDB-DB17FullMay25\RDB-DB17FullMay25.edb"
$global:defaultOriginalLOGFolderPath = "H:\RDB-DB17FullMay25\Logs"
#Database name:
$global:defaultRDBName = "RDB-DB17FullMay25-New"
#Server where we want to put the Recovery Database on:
$global:defaultServer = "JU1EX001"
#Temporary EDB file and LOG folder paths – because New-MailboxDatabase –Recovery requires to have a file path, and New-MailboxDatabase –Recovery won’t let you create a Database, even unmounted, where files already exist, we must first create the Recovery Database using temporary paths. We will change these after using Move-DatabasePath –ConfigurationOnly <- cool, eh !
$global:defaultTempEDBPath = "c:\temp\r-edb.edb"
$global:defaultTempLogPath = "c:\temp\"
#Endregion


#========================================================
#region Functions definitions (NOT the WPF form events)
#========================================================
function IsEmpty($Param){
    If ($Param -eq "All" -or $Param -eq "" -or $Param -eq $Null -or $Param -eq 0) {
        Return $True
    } Else {
        Return $False
    }
}

Function IsPSV3 {
    <#
    .DESCRIPTION
    Just printing Powershell version and returning "true" if powershell version
    is Powershell v3 or more recent, and "false" if it's version 2.
    .OUTPUTS
    Returns $true or $false
    .EXAMPLE
    IsPSVersionV3
    #>
    $PowerShellMajorVersion = $PSVersionTable.PSVersion.Major
    $msgPowershellMajorVersion = "You're running Powershell v$PowerShellMajorVersion"
    Write-Host $msgPowershellMajorVersion -BackgroundColor blue -ForegroundColor yellow
    If($PowerShellMajorVersion -le 2){
        Write-Host "Sorry, PowerShell v3 or more is required. Exiting."
        Return $false
        Exit
    } Else {
        Return $true
        }
}


function update-cmd {
    # Server - custom or default
    If ($wpf.chkServer.IsChecked) {
        $Server = $wpf.txtServer.text
    } Else {
        $Server = $global:defaultServer
    }
    # RDB Name - custom or default
    If ($wpf.chkRDBName.IsChecked){
        $RDBName = $wpf.txtRDBName.text
    } Else {
        $RDBName = $global:defaultRDBName
    }
    # Original EDB File path - custom or default
    If ($wpf.chkOriginalEDBFilePath.IsChecked){
        $OriginalEDBFilePath = $wpf.txtOriginalEDBFilePath.text
    } Else {
        $OriginalEDBFilePath = $global:defaultOriginalEDBFilePath
    }
    # Original LOG file path - custom or default
    If ($wpf.chkOriginalLOGFolderPath.IsChecked){
        $OriginalLOGFolderPath  = $wpf.txtOriginalLOGFolderPath.text
    } Else {
        $OriginalLOGFolderPath  = $global:defaultOriginalLOGFolderPath
    }
    # Temporary EDB file path - custom or default
    If ($wpf.chkTempEDBFilePath.IsChecked){
        $TempEDBPath  = $wpf.txtTempEDBFilePath.text
    } Else {
        $TempEDBPath  = $global:defaultTempEDBPath
    }
    # Temporary LOG file path - custom or default
    If ($wpf.chkTempLOGFolderPath.IsChecked){
        $TempLogPath = $wpf.txtTempLOGFolderPath.text
    } Else {
        $TempLogPath = $global:defaultTempLogPath
    }

    $Comments1 = '<#******* Now BIG STEP #1 – Create the Recovery Database using the Temporary EDB file and LOG folder paths: ******* #>'
    $command1 = "New-MailboxDatabase -Recovery -Name ""$RDBName"" -Server $Server -EDBFilePath ""$TempEDBPath"" -LogFolderPath ""$TempLogPath"""
    $Comments2 = '<#******* Now BIG STEP #2 – Move the Recovery Database EDB file and LOG folder paths to the paths containing your original files ******* #>'
    $Command2 = "Move-DatabasePath -Identity ""$RDBName"" -ConfigurationOnly -EdbFilePath ""$OriginalEDBFilePath"" -LogFolderPath ""$OriginalLOGFolderPath"""
    $Comments3 = '<#******* Now BIG STEP #3 – CHECK that the paths of the new Recovery Database have moved to the original ones we wanted ! ******* #>'
    $Command3 = "Get-MailboxDatabase ""$RDBName"" -Status | ft Name,EDBFilePath,LogFolderPath,mounted -a"
    $Comments4 = '<#******* Now BIG STEP #4 – Mount your database ! And then check if "Mounted"  shows "True"  ******* #>'
    $command4 = ("Mount-Database ""$RDBName""") + "`n#Check if it's mounted:`n" + $Command3

    $wpf.txtCommandLine.text = $Comments1 + "`n" + $command1 + "`n`n" + $Comments2 + "`n" + $Command2 + "`n`n" + $Comments3 + "`n" + $Command3 + "`n`n" + $Comments4 + "`n" + $Command4 + "`n`n" + $Comments5 + "`n" + $Command5
}

#========================================================
# END of Functions definitions (NOT the WPF form events)
#endregion
#========================================================

#========================================================
#region WPF form definition and load controls
#========================================================

# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{}
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
# $inputXML = Get-Content -Path "C:\Users\Kamehameha\Documents\GitHub\PowerShell\Get-EventsFromEventLog\VisualStudio2017WPFDesign\Launch-EventsCollector-WPF\Launch-EventsCollector-WPF\MainWindow.xaml"
$inputXML = @"
<Window x:Name="NewRestoreDB" x:Class="E201020162013_CreateNewRDBUseExistingFiles.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:E201020162013_CreateNewRDBUseExistingFiles"
        mc:Ignorable="d"
        Title="New Recovery Database" Height="569.75" Width="1056" ResizeMode="NoResize">
    <Window.Background>
        <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
            <GradientStop Color="#FF23BC2A" Offset="0"/>
            <GradientStop Color="#FF2118E5" Offset="1"/>
        </LinearGradientBrush>
    </Window.Background>
    <Grid>
        <TextBox x:Name="txtOriginalEDBFilePath" HorizontalAlignment="Left" Height="23" Margin="53,36,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="302" IsEnabled="False" Text="E:\Databases\DB01\EDB\DB01.edb"/>
        <Label Content="Original EDB File Path" HorizontalAlignment="Left" Margin="28,10,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="chkOriginalEDBFilePath" Content="" HorizontalAlignment="Left" Margin="28,36,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtOriginalLOGFolderPath" HorizontalAlignment="Left" Height="23" Margin="53,98,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="302" IsEnabled="False" Text="E:\Databases\DB01\LOGS\"/>
        <Label Content="Original LOG Folder Path" HorizontalAlignment="Left" Margin="28,72,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="chkOriginalLOGFolderPath" Content="" HorizontalAlignment="Left" Margin="28,98,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtTempEDBFilePath" HorizontalAlignment="Left" Height="22" Margin="53,170,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="302" IsEnabled="False" Text="c:\temp\tempDB.edb"/>
        <Label Content="Temporary EDB file path" HorizontalAlignment="Left" Margin="28,144,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="chkTempEDBFilePath" Content="" HorizontalAlignment="Left" Margin="28,170,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtTempLOGFolderPath" HorizontalAlignment="Left" Height="22" Margin="53,250,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="302" IsEnabled="False" Text="c:\temp\"/>
        <Label Content="Temporary LOG folder Path" HorizontalAlignment="Left" Margin="28,224,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="chkTempLOGFolderPath" Content="" HorizontalAlignment="Left" Margin="28,250,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtRDBName" HorizontalAlignment="Left" Height="22" Margin="53,328,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="302" IsEnabled="False" Text="RecoveryDB01"/>
        <Label Content="Recovery Database Name" HorizontalAlignment="Left" Margin="28,302,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="chkRDBName" Content="" HorizontalAlignment="Left" Margin="28,328,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtServer" HorizontalAlignment="Left" Height="22" Margin="53,396,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="302" IsEnabled="False" Text="Server001"/>
        <Label Content="Server where to store the Recovery Database on" HorizontalAlignment="Left" Margin="28,370,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="chkServer" Content="" HorizontalAlignment="Left" Margin="28,396,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtCommandLine" HorizontalAlignment="Left" Height="382" Margin="408,36,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="618"/>
        <Button x:Name="btnRun" Content="Run" HorizontalAlignment="Left" Margin="234,470,0,0" VerticalAlignment="Top" Width="138" Height="39"/>
        <Button x:Name="btnCancel" Content="Cancel" HorizontalAlignment="Left" Margin="528,470,0,0" VerticalAlignment="Top" Width="138" Height="39"/>

    </Grid>
</Window>

"@
$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

#========================================================
# END of WPF form definition and load controls
#endregion
#========================================================

#========================================================
#region WPF EVENTS definition
#========================================================

#region Buttons
$wpf.btnRun.add_Click({
    $msg = "Running the command"
    Write-Host $msg
    #Invoke-expression $wpf.txtCommand.text
})

$wpf.btnCancel.add_Click({
    $msg = "Exiting..."
    Write-Host $msg
    $wpf.NewRestoreDB.Close()
})
# End of Buttons region
#endregion

#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
$wpf.NewRestoreDB.Add_Loaded({
})
#Things to load when the WPF form is rendered aka drawn on screen
$wpf.NewRestoreDB.Add_ContentRendered({
    Update-cmd
})
$wpf.NewRestoreDB.add_Closing({
    $msg = "bye bye !"
    write-host $msg
})
# End of load, draw and closing form events
#endregion

#region Text Changed events
$wpf.txtOriginalEDBFilePath.add_TextChanged({
    Update-cmd
})
$wpf.txtOriginalLOGFolderPath.add_TextChanged({
    Update-cmd
})
$wpf.txtTempEDBFilePath.add_TextChanged({
    Update-cmd
})
$wpf.txtTempLOGFolderPath.add_TextChanged({
    Update-cmd
})
$wpf.txtRDBName.add_TextChanged({
    Update-cmd
})
$wpf.txtServer.add_TextChanged({
    Update-cmd
})

#End of Text Changed events
#endregion

#region Check Boxes
$wpf.chkOriginalEDBFilePath.add_Checked({
    $wpf.txtOriginalEDBFilePath.Isenabled = $true
    update-cmd
})
$wpf.chkOriginalEDBFilePath.add_UnChecked({
    $wpf.txtOriginalEDBFilePath.Isenabled = $false
    update-cmd
})

$wpf.chkOriginalLOGFolderPath.add_Checked({
    $wpf.txtOriginalLOGFolderPath.Isenabled = $true
    update-cmd
})
$wpf.chkOriginalLOGFolderPath.add_UnChecked({
    $wpf.txtOriginalLOGFolderPath.Isenabled = $false
    update-cmd
})

$wpf.chkTempEDBFilePath.add_Checked({
    $wpf.txtTempEDBFilePath.Isenabled = $true
    update-cmd
})
$wpf.chkTempEDBFilePath.add_UnChecked({
    $wpf.txtTempEDBFilePath.Isenabled = $false
    update-cmd
})

$wpf.chkTempLOGFolderPath.add_Checked({
    $wpf.txtTempLOGFolderPath.Isenabled = $true
    update-cmd
})
$wpf.chkTempLOGFolderPath.add_UnChecked({
    $wpf.txtTempLOGFolderPath.Isenabled = $false
    update-cmd
})

$wpf.chkRDBName.add_Checked({
    $wpf.txtRDBName.Isenabled = $true
    update-cmd
})
$wpf.chkRDBName.add_UnChecked({
    $wpf.txtRDBName.Isenabled = $false
    update-cmd
})

$wpf.chkServer.add_Checked({
    $wpf.txtServer.Isenabled = $true
    update-cmd
})
$wpf.chkServer.add_UnChecked({
    $wpf.txtServer.Isenabled = $false
    update-cmd
})

# End of check boxes
#endregion

#region Clicked on Checkboxes events
# $wpf.chkBoxName.add_Click({
#     Update-cmd
# })

# End of Clicked on Checkboxes events
#endregion

#=======================================================
#End of Events from the WPF form
#endregion
#=======================================================

IsPSV3 | out-null
# Load the form:
$wpf.NewRestoreDB.ShowDialog() | Out-Null