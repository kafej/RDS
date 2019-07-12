#$ErrorActionPreference= 'silentlycontinue'

Add-Type -AssemblyName PresentationFramework, PresentationCore, System.Windows.Forms

[System.Windows.Forms.Application]::EnableVisualStyles()

$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$AssemblyLocation = Join-Path -Path $PSScriptRoot -ChildPath .\include
foreach ($Assembly in (Get-ChildItem $AssemblyLocation -Filter *.dll)) {
    [System.Reflection.Assembly]::LoadFrom($Assembly.fullName) | out-null
}

[xml]$xaml = @"
<Controls:MetroWindow 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    xmlns:Dialog="clr-namespace:MahApps.Metro.Controls.Dialogs;assembly=MahApps.Metro"
    Title="RDS - Sesje"
    ResizeMode="NoResize"
    Width="1500"
    Height="700"
    WindowStartupLocation="CenterScreen">

    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="$AssemblyLocation\Icons.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/Cyan.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>

    <Grid Background="white">  
		<StackPanel HorizontalAlignment="Center">
            <Grid>
            <StackPanel Margin="10,10,0,0" HorizontalAlignment="Left">
                <Label Content="Razem:" HorizontalAlignment="Left" Margin="0,0,0,0" VerticalAlignment="Top"/>
                <TextBox HorizontalAlignment="Left" x:Name="info" Margin="44,-25,0,0" IsReadOnly="True" BorderThickness="0"/>
                <Label Content="Sesji Aktywnych:" HorizontalAlignment="Left" Margin="80,-27,0,0"/>
                <TextBox HorizontalAlignment="Left" x:Name="infoa" Margin="172,-26,0,0" IsReadOnly="True" BorderThickness="0"/>
                <Label Content="Sesji Nieaktywnych:" HorizontalAlignment="Left" Margin="200,-27,0,0"/>
                <TextBox HorizontalAlignment="Left" x:Name="infoi" Margin="310,-26,0,0" IsReadOnly="True" BorderThickness="0"/>
            </StackPanel>
            </Grid>
			<StackPanel Margin="10,10,10,10" HorizontalAlignment="Center" VerticalAlignment="Top" Orientation="Vertical">
                <DataGrid x:Name="gridMain" AutoGenerateColumns="False" Margin="10,0,0,0"
                    GridLinesVisibility="All" HorizontalGridLinesBrush="#FFD4D4D4"  VerticalGridLinesBrush="#FFD4D4D4" 
                    OverridesDefaultStyle="True" CanUserAddRows="False" Height="620">
                    <DataGrid.Columns>
                        <DataGridTextColumn Binding="{Binding Login}" Header="Login" IsReadOnly="True" Width="auto"/>
                        <DataGridTextColumn Binding="{Binding NameSurname}" Header="Name and Surname" IsReadOnly="True" Width="auto"/>
                        <DataGridTextColumn Binding="{Binding State}" Header="State" IsReadOnly="True" Width="auto"/>
                        <DataGridTextColumn Binding="{Binding ProfileStatus}" Header="Profile Status" IsReadOnly="True" Width="auto"/>
                        <DataGridTextColumn Binding="{Binding Server}" Header="Server" IsReadOnly="True" Width="auto"/>
                        <DataGridTextColumn Binding="{Binding Department}" Header="Department" IsReadOnly="True" Width="auto"/>
                        <DataGridTextColumn Binding="{Binding ID}" Header="ID" IsReadOnly="True" Width="auto"/>
                        <DataGridTextColumn Binding="{Binding SID}" Header="SID" IsReadOnly="True" Width="auto"/>
                        <DataGridTextColumn Binding="{Binding SessionName}" Header="Session Name" IsReadOnly="True" Width="auto"/>

                        <DataGridTemplateColumn>
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <StackPanel Orientation="Horizontal">

                                        <Button x:Name="View" Background="#1c56df" Style="{DynamicResource MetroCircleButtonStyle}" 
                                            Height="28" Width="28" Cursor="Hand" HorizontalContentAlignment="Stretch" 
                                            VerticalContentAlignment="Stretch" HorizontalAlignment="Center" VerticalAlignment="Center" 
                                            BorderThickness="0" Margin="0,0,0,0">
                                            <Rectangle Width="12" Height="12" HorizontalAlignment="Center" VerticalAlignment="Center" Fill="white">
                                                <Rectangle.OpacityMask>
                                                    <VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_monitor_to}"/>
                                                </Rectangle.OpacityMask>
                                            </Rectangle>
                                        </Button>

                                        <Button x:Name="Control" Background="#198C19" Style="{DynamicResource MetroCircleButtonStyle}" 
                                            Height="28" Width="28" Cursor="Hand" HorizontalContentAlignment="Stretch" 
                                            VerticalContentAlignment="Stretch" HorizontalAlignment="Center" VerticalAlignment="Center" 
                                            BorderThickness="0" Margin="0,0,0,0">
                                            <Rectangle Width="12" Height="12" HorizontalAlignment="Center" VerticalAlignment="Center" Fill="white">
                                                <Rectangle.OpacityMask>
                                                    <VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_network}"/>
                                                </Rectangle.OpacityMask>
                                            </Rectangle>
                                        </Button>

                                        <Button x:Name="Close" Background="#fb0013" Style="{DynamicResource MetroCircleButtonStyle}" 
                                            Height="28" Width="28" Cursor="Hand" HorizontalContentAlignment="Stretch" 
                                            VerticalContentAlignment="Stretch" HorizontalAlignment="Center" VerticalAlignment="Center" 
                                            BorderThickness="0" Margin="0,0,0,0">
                                            <Rectangle Width="10" Height="10" HorizontalAlignment="Center" VerticalAlignment="Center" Fill="white" >
                                                <Rectangle.OpacityMask>
                                                    <VisualBrush  Stretch="Fill" Visual="{StaticResource appbar_close}"/>
                                                </Rectangle.OpacityMask>
                                            </Rectangle>
                                        </Button>
                                    </StackPanel>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                    </DataGrid.Columns>
                </DataGrid>
			</StackPanel>
		</StackPanel>
        <StatusBar Name="statusBar" Height="30" HorizontalAlignment="Stretch" VerticalAlignment="Bottom">
            <StatusBarItem HorizontalAlignment="Left">
                <Label Content="Przygotował Michał Zbyl" FontSize="11"/>
            </StatusBarItem>
            <StatusBarItem HorizontalAlignment="Right" Margin="0,0,10,0">
                <Label Content="Ver. 1.0.0"/>
            </StatusBarItem>
        </StatusBar>	
    </Grid>
</Controls:MetroWindow>
"@

$Reader = (New-Object System.Xml.XmlNodeReader $xaml)
$Form = [Windows.Markup.XamlReader]::Load($reader)

Copy-Item -Path "C:\Windows\WinSxS\amd64_microsoft-windows-t..commandlinetoolsmqq_*\qwinsta.exe" -Destination "$env:TEMP\" -Force
Copy-Item -Path "C:\Windows\WinSxS\amd64_microsoft-windows-t..es-commandlinetools_*\rwinsta.exe" -Destination "$env:TEMP\" -Force

$pcs = @()
$domain =

$Prof = @()
$Content = @()

$pathRDP = $env:temp

$i = 0

foreach ($pc in $pcs) {
    $i++
    $n = $pc+$i
    $n1 = 'n'+$n
    Start-Process powershell "Get-WmiObject win32_userprofile -ComputerName '$pc' | select-object SID,status | Out-File '$pathRDP\$n.txt'" -NoNewWindow
    Start-Process cmd "/c $pathRDP\qwinsta.exe /SERVER:$pc > $pathRDP\$n1.txt" -NoNewWindow -Wait
}

Get-ChildItem "$pathRDP\" -Filter '*RD*' | 
Foreach-Object {
    $s = Get-Content $_.FullName
    $s1 = $s -replace '\s\s+', ","
    $s1 | Out-File $_.FullName
    if ($_.FullName -like '*nRDTS0*') {
        $che = $_.Name.Split(".")
        $chef = $che[0].substring(1)
        $chef = $chef.Substring(0,$chef.Length-1)
        $check = Import-Csv $_.FullName -Encoding utf8 -Header "Session Name", "Login", "ID", "State", "Server" | Select-Object -Skip 2 | select-object -SkipLast 1
        $check | ForEach-Object {
            $_.Server = $chef
        }
        $chefg = $chef.Split("T")
        $chefg = $chefg[0]+$chefg[1]
        $check | Export-Csv $pathRDP\$chefg.csv
    }
}

Get-ChildItem "$pathRDP\" -Filter 'RD*' | 
Foreach-Object {
    if ($_.FullName -like '*RDTS0*.txt') {
        $Prof += Import-Csv $_.FullName -Encoding utf8 -Header "SID", "Profile Status" | Select-Object -Skip 2
    } elseif ($_.FullName -like '*RDS0*.csv') {
        $Content += Import-Csv $_.FullName -Encoding utf8 -Header "Session Name", "Login", "ID", "State", "Server", "Name and Surname", "Department", "Profile Status", "SID" | Select-Object -Skip 1
    }
}

$akc = 0
$akd = 0

$Content | ForEach-Object {
    $usr = $_.Login

    if ($usr -like 'p0*') {
        $objUser = New-Object System.Security.Principal.NTAccount($domain, $usr)
        $SIDF = $objUser.Translate([System.Security.Principal.SecurityIdentifier])
    
        $proff = $Prof | Where-Object {$_.SID -eq "$SIDF"}
        $Prof1 = $proff | Select-Object -ExpandProperty "Profile Status"
        if ($Prof1 -contains '1') {
            $_."Profile Status" = 'Temporary'
        } else {
            $_."Profile Status" = 'Normal'
        }
    
        $_."SID" = $SIDF
    
        $ldapnamef = Get-ADUser -Filter {SamAccountName -like $usr}
        $ldapname = $ldapnamef | Select-Object -ExpandProperty Name
        $_."Name and Surname" = $ldapname
    
        $check_userAD = $ldapnamef | Select-Object -ExpandProperty DistinguishedName
        $odd1 = $check_userAD.split(',')[1]
        $odd2 = $odd1.Split('=')[1]
    
        $_."Department" = $odd2
        
        $stan = $_."State"
    
        if ($stan -like 'Aktywna') {
            $akc++
        } elseif ($stan -like 'Active') {
            $akc++
        } else {
            $akd++
        }
    }
}

$columns = $Content | Select-Object -Property "Login", "Name and Surname", "State", "Profile Status", "Server", "Department", "ID", "SID", "Session Name"

Remove-Item $pathRDP\RDS*.txt -Force
Remove-Item $pathRDP\RDS*.csv -Force
Remove-Item $pathRDP\RDTS*.txt -Force
Remove-Item $pathRDP\*RDTS*.txt -Force

$dataGrid = $Form.FindName("gridMain")
$info = $Form.FindName("info")
$infoa = $Form.FindName("infoa")
$infoi = $Form.FindName("infoi")

$info.Text = ($akc+$akd)
$infoa.Text = ($akc)
$infoi.Text = ($akd)

$script:observableCollection = [System.Collections.ObjectModel.ObservableCollection[Object]]::new()

$columns | ForEach-Object {
    $objArray = New-Object PSObject
    $objArray | Add-Member -type NoteProperty -name Login -value $_."Login"
    $objArray | Add-Member -type NoteProperty -name NameSurname -value $_."Name and Surname"
    $objArray | Add-Member -type NoteProperty -name State -value $_."State"
    $objArray | Add-Member -type NoteProperty -name ProfileStatus -value $_."Profile Status"
    $objArray | Add-Member -type NoteProperty -name Server -value $_."Server"
    $objArray | Add-Member -type NoteProperty -name Department -value $_."Department"
    $objArray | Add-Member -type NoteProperty -name ID -value $_."ID"
    $objArray | Add-Member -type NoteProperty -name SID -value $_."SID"
    $objArray | Add-Member -type NoteProperty -name SessionName -value $_."Session Name"
    $script:observableCollection.Add($objArray) | Out-Null
}

$dataGrid.ItemsSource = $Script:observableCollection

Function view($rowObj){
    $tid = $rowObj.ID
    $tserver = $rowObj.Server

    mstsc /shadow:$tid /v:$tserver /noConsentPrompt
}

Function control($rowObj){ 
    $tid = $rowObj.ID
    $tserver = $rowObj.Server

    mstsc /shadow:$tid /control /v:$tserver /noConsentPrompt  
}

Function close($rowObj){
    $tid = $rowObj.ID
    $tserver = $rowObj.Server

    Start-Process cmd "/c $pathRDP\rwinsta.exe /SERVER:$tserver $tid" -NoNewWindow
    $script:observableCollection.Remove($rowObj)
}

[System.Windows.RoutedEventHandler]$EventonDataGrid = {

    $button =  $_.OriginalSource.Name
    $Script:resultObj = $dataGrid.CurrentItem

    If ( $button -match "View" ){
        view -rowObj $resultObj
    }
    ElseIf ($button -match "Control" ){
        control -rowObj $resultObj
    }
    ElseIf ($button -match "Close" ){
        close -rowObj $resultObj
    }

}
$dataGrid.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonDataGrid)

[void]$Form.ShowDialog()
