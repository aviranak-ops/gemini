<#
.SYNOPSIS
    SharePoint Ultimate Patcher - WPF Modern UI Edition
    *NEW: Live Installation Heartbeat (Checks remote processes every 60 sec)
    *NEW: Indeterminate Progress Bar for long operations
    *FIXED: Bulletproof XAML String parsing (Single Quotes instead of Here-String)
    *FIXED: Push-Method file copy (Resolves Double-Hop Access Denied) + Strict Error Handling
    *NEW: Optimized Downtime (Files copied BEFORE services are stopped)
    *NEW: Persistent File Logging (\\Storage\SPUpdates\Logs)
    *NEW: Folder-based Storage Scanning & Auto-Archiving (Done_YYYY-MM)
    *NEW: Full Parallel Multi-Farm Support (SP2019 & SPSE concurrently)
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ==========================================
# 0. Administrator Privilege Check
# ==========================================
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    [System.Windows.MessageBox]::Show("Please run this script as Administrator.", "Admin Rights Required", 0, 48)
    exit
}

# ==========================================
# 1. Configuration Data
# ==========================================
$Global:Environments = @{
    "Dev" = @{
        Servers   = @("shirin19", "devspsub")
        CAServers = @("shirin19", "devspsub")
    }
    "Int" = @{
        Servers   = @("SP19-INT-APP", "SPSE-INT-APP")
        CAServers = @("SP19-INT-APP", "SPSE-INT-APP")
    }
    "Prod" = @{
        Servers   = @("SP19-PRD-APP1", "SP19-PRD-WFE1", "SPSE-PRD-APP1", "SPSE-PRD-WFE1")
        CAServers = @("SP19-PRD-APP1", "SPSE-PRD-APP1")
    }
}

$Global:StorageBasePath = "\\Storage\SPUpdates"
$Global:LocalTempDir = "C:\SP_Updates_Temp"
$Global:CurrentLogFile = $null

# ==========================================
# 2. XAML (Modern UI Design)
# ==========================================
$xaml = '
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SharePoint Multi-Farm Orchestrator" Height="700" Width="750"
        Background="#1E1E1E" WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="14">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1084D9"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#444444"/>
                    <Setter Property="Foreground" Value="#888888"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="250"/>
            <ColumnDefinition Width="20"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Row="0" Grid.Column="0">
            <TextBlock Text="1. Select Environment:" Foreground="#CCCCCC" FontWeight="SemiBold" Margin="0,0,0,5"/>
            <ComboBox Name="CmbEnv" Height="30" Background="#333333" Foreground="Black" FontSize="14"/>
        </StackPanel>

        <StackPanel Grid.Row="0" Grid.Column="2" Grid.RowSpan="2">
            <TextBlock Text="2. Target Servers:" Foreground="#CCCCCC" FontWeight="SemiBold" Margin="0,0,0,5"/>
            <ListBox Name="LstServers" Background="#2D2D30" BorderThickness="1" BorderBrush="#3F3F46" Height="120" ScrollViewer.VerticalScrollBarVisibility="Auto" Padding="5"/>
        </StackPanel>

        <Button Name="BtnStart" Content="▶ START PATCHING" Grid.Row="1" Grid.Column="0" Height="45" Margin="0,20,0,0" IsEnabled="False"/>

        <StackPanel Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,25,0,0">
            <TextBlock Name="LblStatus" Text="Ready." Foreground="#00BFFF" FontWeight="SemiBold" Margin="0,0,0,5"/>
            <ProgressBar Name="ProgressBar" Height="8" Minimum="0" Maximum="100" BorderThickness="0" Background="#333333" Foreground="#0078D4"/>
        </StackPanel>

        <RichTextBox Name="LogBox" Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,15,0,0"
                     Background="#0C0C0C" Foreground="#D4D4D4" FontFamily="Consolas" FontSize="12"
                     BorderThickness="1" BorderBrush="#3F3F46" IsReadOnly="True" VerticalScrollBarVisibility="Auto">
            <FlowDocument Name="LogDocument" PagePadding="5">
            </FlowDocument>
        </RichTextBox>
    </Grid>
</Window>
'

$reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
$Form = [Windows.Markup.XamlReader]::Load($reader)

# ==========================================
# 3. UI Controls Mapping
# ==========================================
$CmbEnv = $Form.FindName("CmbEnv")
$LstServers = $Form.FindName("LstServers")
$BtnStart = $Form.FindName("BtnStart")
$LblStatus = $Form.FindName("LblStatus")
$ProgressBar = $Form.FindName("ProgressBar")
$LogBox = $Form.FindName("LogBox")
$LogDocument = $Form.FindName("LogDocument")

$Global:Environments.Keys | Sort-Object | ForEach-Object { $CmbEnv.Items.Add($_) | Out-Null }

# ==========================================
# 4. Helper Functions
# ==========================================
function DoEvents {
    $Dispatcher = [System.Windows.Threading.Dispatcher]::CurrentDispatcher
    $Frame = New-Object System.Windows.Threading.DispatcherFrame
    $Dispatcher.BeginInvoke("Background", [System.Action] { $Frame.Continue = $false }) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($Frame)
}

function Write-WpfLog($Message, $Color = "LightGray") {
    $Timestamp = (Get-Date).ToString("HH:mm:ss")
    $FormattedMsg = "[$Timestamp] $Message"

    $Run = New-Object System.Windows.Documents.Run($FormattedMsg)
    $Run.Foreground = (New-Object System.Windows.Media.BrushConverter).ConvertFromString($Color)
    $Paragraph = New-Object System.Windows.Documents.Paragraph($Run)
    $Paragraph.Margin = New-Object System.Windows.Thickness(0)
    $LogDocument.Blocks.Add($Paragraph)
    $LogBox.ScrollToEnd()
    DoEvents

    if ($Global:CurrentLogFile) {
        Add-Content -Path $Global:CurrentLogFile -Value $FormattedMsg -ErrorAction SilentlyContinue
    }
}

function Set-Progress($Percent, $StatusText, $IsIndeterminate = $false) {
    $ProgressBar.IsIndeterminate = $IsIndeterminate
    if (-not $IsIndeterminate) { $ProgressBar.Value = $Percent }
    $LblStatus.Text = $StatusText
    DoEvents
}

function Wait-JobWithWPF($Job, $TimeoutMins, $ServersToCheck = $null) {
    if (-not $Job) { return $true }
    $StartTime = Get-Date
    $LastHeartbeat = Get-Date

    # רשימת תהליכים שאנחנו מחפשים בשרתים המרוחקים כדי לוודא שההתקנה עובדת
    $ProcNames = @("msiexec", "sts*", "uber*", "spserver*", "sp2019*", "psconfig", "psconfigui", "ose")

    while ($Job.State -contains 'Running') {
        $Now = Get-Date

        if ($Now -gt $StartTime.AddMinutes($TimeoutMins)) {
            Write-WpfLog "Job Timeout Exceeded!" "#FF5555"
            $Job | Stop-Job
            return $false
        }

        # דופק כל 60 שניות - בדיקה מול השרתים
        if (($Now - $LastHeartbeat).TotalSeconds -ge 60) {
            $ElapsedMins = [math]::Round(($Now - $StartTime).TotalMinutes, 1)

            if ($ServersToCheck) {
                $CheckBlock = {
                    param($Names)
                    $procs = Get-Process -Name $Names -ErrorAction SilentlyContinue
                    # התיקון: החזרת אובייקט מפורש כדי שהסינון יעבוד
                    return [PSCustomObject]@{ IsActive = [bool]$procs }
                }

                $ProcStatus = Invoke-Command -ComputerName $ServersToCheck -ScriptBlock $CheckBlock -ArgumentList (,$ProcNames) -ErrorAction SilentlyContinue

                # התיקון: סינון לפי המאפיין IsActive
                $ActiveServers = ($ProcStatus | Where-Object { $_.IsActive -eq $true }).PSComputerName -join ", "

                if ($ActiveServers) {
                    Write-WpfLog "[Elapsed: $ElapsedMins Min] -> Install/Config actively running on: $ActiveServers" "#87CEFA"
                } else {
                    Write-WpfLog "[Elapsed: $ElapsedMins Min] -> Background tasks finalizing..." "#87CEFA"
                }
            } else {
                Write-WpfLog "[Elapsed: $ElapsedMins Min] -> Task is still running..." "#87CEFA"
            }

            $LastHeartbeat = $Now
        }

        Start-Sleep -Milliseconds 500
        DoEvents
    }
    return $true
}

# ==========================================
# 5. Events & Core Logic
# ==========================================
$CmbEnv.add_SelectionChanged({
    $Env = $CmbEnv.SelectedItem
    $AllServers = $Global:Environments[$Env].Servers
    $CAServers = $Global:Environments[$Env].CAServers

    $LstServers.Items.Clear()
    foreach ($Srv in $AllServers) {
        $Suffix = if ($Srv -in $CAServers) { " (Central Admin)" } else { " (App/WFE)" }
        $Chk = New-Object System.Windows.Controls.CheckBox
        $Chk.Content = "$Srv $Suffix"
        $Chk.Foreground = "White"
        $Chk.IsChecked = $true
        $Chk.Margin = New-Object System.Windows.Thickness(0, 2, 0, 2)
        $Chk.Tag = $Srv
        $LstServers.Items.Add($Chk) | Out-Null
    }
    $BtnStart.IsEnabled = $true
})

$BtnStart.add_Click({
    $SelectedServers = @()
    $CAServers = $Global:Environments[$CmbEnv.SelectedItem].CAServers

    foreach ($Item in $LstServers.Items) {
        if ($Item.IsChecked) { $SelectedServers += $Item.Tag }
    }

    if ($SelectedServers.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one server.", "Notice", 0, 48)
        return
    }

    $BtnStart.IsEnabled = $false
    $CmbEnv.IsEnabled = $false
    $LstServers.IsEnabled = $false

    $LogFolder = Join-Path -Path $Global:StorageBasePath -ChildPath "Logs"
    if (-not (Test-Path $LogFolder)) {
        New-Item -ItemType Directory -Path $LogFolder -Force -ErrorAction SilentlyContinue | Out-Null
    }

    $DateStamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $Global:CurrentLogFile = Join-Path -Path $LogFolder -ChildPath "PatchingLog_$($CmbEnv.SelectedItem)_$DateStamp.log"

    try {
        Write-WpfLog "--- Initiating Multi-Farm Upgrade for $($CmbEnv.SelectedItem) ---" "#00BFFF"
        Write-WpfLog "Log file established at: $($Global:CurrentLogFile)" "#D4D4D4"

        # 1. Pre-Flight
        Set-Progress 10 "Detecting & Grouping SharePoint Versions..."
        $ServerMap = @{ "SP2019" = @(); "SPSE" = @() }

        $VersionCheck = Invoke-Command -ComputerName $SelectedServers -ErrorAction Stop -ScriptBlock {
            $val = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\16.0" -Name "PrecisionVersion" -ErrorAction SilentlyContinue).PrecisionVersion
            if (-not $val) {
                $val = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office Server\16.0" -Name "BuildVersion" -ErrorAction SilentlyContinue).BuildVersion
            }

            $detectedVersion = "Unknown_$val"
            if (-not $val) {
                $detectedVersion = "Unknown_NoRegistry"
            } else {
                try {
                    $v = [version]$val
                    if ($v.Major -eq 16) {
                        if ($v.Build -ge 10000 -and $v.Build -lt 14000) { $detectedVersion = "SP2019" }
                        elseif ($v.Build -ge 14000) { $detectedVersion = "SPSE" }
                    }
                } catch { $detectedVersion = "Unknown_ParseError_$val" }
            }
            return [PSCustomObject]@{ Value = $detectedVersion }
        }

        foreach ($Res in $VersionCheck) {
            $Ver = $Res.Value
            $Srv = $Res.PSComputerName
            if ($Ver -in @("SP2019", "SPSE")) {
                $ServerMap[$Ver] += $Srv
                Write-WpfLog "Mapped $Srv to $Ver" "#D4D4D4"
            } else {
                throw "Unknown version on $Srv (Detected: $Ver)"
            }
        }

        Write-WpfLog "Scanning Storage subfolders under $($Global:StorageBasePath)..." "#FFD700"
        $UpdateFiles = @{ "SP2019" = @(); "SPSE" = @() }

        foreach ($Ver in $ServerMap.Keys) {
            if ($ServerMap[$Ver].Count -gt 0) {
                $VersionFolder = Join-Path -Path $Global:StorageBasePath -ChildPath $Ver
                if (-not (Test-Path $VersionFolder)) { New-Item -ItemType Directory -Path $VersionFolder -Force | Out-Null }

                $FoundFiles = Get-ChildItem -Path $VersionFolder -File -ErrorAction SilentlyContinue
                if (-not $FoundFiles -or $FoundFiles.Count -eq 0) { throw "No update files found in: $VersionFolder" }

                $UpdateFiles[$Ver] = @($FoundFiles.FullName)
                Write-WpfLog "Found $($UpdateFiles[$Ver].Count) file(s) for $Ver in $VersionFolder" "#32CD32"
            }
        }

        Write-WpfLog "Validating disk space on all servers..." "#FFD700"
        Invoke-Command -ComputerName $SelectedServers -ErrorAction Stop -ScriptBlock {
            $Free = (Get-Volume -DriveLetter C).SizeRemaining / 1GB
            if ($Free -lt 15) { throw "Low disk space on $env:COMPUTERNAME" }
        }
        Write-WpfLog "Pre-Flight checks passed." "#32CD32"

        # 2. Copy Files
        Set-Progress 25 "Distributing Update Files (Services running, no downtime yet)..."
        foreach ($Ver in $ServerMap.Keys) {
            if ($ServerMap[$Ver].Count -gt 0) {
                $Files = $UpdateFiles[$Ver]
                foreach ($Srv in $ServerMap[$Ver]) {
                    $TargetUNC = "\\$Srv\" + $Global:LocalTempDir.Replace(':', '$')
                    if (-not (Test-Path $TargetUNC)) {
                        Write-WpfLog "Creating temp directory on $Srv..." "#D4D4D4"
                        New-Item -ItemType Directory -Path $TargetUNC -Force -ErrorAction Stop | Out-Null
                    }
                    foreach ($F in $Files) {
                        Write-WpfLog "Copying $(Split-Path $F -Leaf) to $Srv..." "#FFD700"
                        Copy-Item -Path $F -Destination $TargetUNC -Force -ErrorAction Stop
                    }
                }
            }
        }
        Write-WpfLog "All files distributed successfully. Ready for patching." "#32CD32"

        # 3. Stop Services
        Set-Progress 40 "Stopping SharePoint Services..."
        Write-WpfLog "Stopping IIS and Timer Service globally..." "#FFD700"
        Invoke-Command -ComputerName $SelectedServers -ScriptBlock {
            iisreset /stop | Out-Null
            Stop-Service "SPTimerV4" -Force -WarningAction SilentlyContinue
        }
        Write-WpfLog "Services stopped." "#32CD32"

        # 4. Patching (With Indeterminate Progress & Heartbeat)
        Set-Progress 50 "Installing Patches (This will take time)..." $true

        $MaxFiles = 0
        foreach ($Ver in $ServerMap.Keys) {
            if ($ServerMap[$Ver].Count -gt 0 -and $UpdateFiles[$Ver].Count -gt $MaxFiles) {
                $MaxFiles = $UpdateFiles[$Ver].Count
            }
        }

        for ($i = 0; $i -lt $MaxFiles; $i++) {
            $Jobs = @()
            $TargetServersForThisRun = @()

            foreach ($Ver in $ServerMap.Keys) {
                if ($ServerMap[$Ver].Count -gt 0 -and $UpdateFiles[$Ver].Count -gt $i) {
                    $File = Split-Path $UpdateFiles[$Ver][$i] -Leaf
                    Write-WpfLog "Initiating install of $File on $Ver servers..." "#FFD700"

                    $TargetServersForThisRun += $ServerMap[$Ver]
                    $Jobs += Invoke-Command -ComputerName $ServerMap[$Ver] -AsJob -ArgumentList $File, $Global:LocalTempDir -ScriptBlock {
                        param($F, $T)
                        $Process = Start-Process -FilePath "$T\$F" -ArgumentList "/passive /norestart" -Wait -PassThru
                        return $Process.ExitCode
                    }
                }
            }

            if ($Jobs.Count -gt 0) {
                # קריאה לפונקציית ההמתנה שכוללת את ההרטביט החדש (מעבירים את רשימת השרתים לבדיקה)
                Wait-JobWithWPF -Job $Jobs -TimeoutMins 120 -ServersToCheck $TargetServersForThisRun | Out-Null
                $Results = Receive-Job -Job $Jobs
                foreach ($Res in $Results) {
                    if ($Res.Value -in @(0, 3010)) { Write-WpfLog "$($Res.PSComputerName): Patch Installed" "#32CD32" }
                    else { throw "$($Res.PSComputerName): Patch FAILED (Code: $($Res.Value))" }
                }
            }
        }
        Set-Progress 65 "Patching phase completed successfully." $false

        # 5. PSConfig (CA Servers First)
        Set-Progress 70 "Running PSConfig on Central Admin Servers..." $true
        $SelectedCAs = $SelectedServers | Where-Object { $_ -in $CAServers }

        if ($SelectedCAs.Count -gt 0) {
            Write-WpfLog "Starting PSConfig on CA servers: $($SelectedCAs -join ', ')..." "#FFD700"
            $CAJobs = Invoke-Command -ComputerName $SelectedCAs -AsJob -ScriptBlock {
                $Exe = "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\BIN\PSConfig.exe"
                $Process = Start-Process $Exe -ArgumentList "-cmd upgrade -inplace b2b -wait -cmd applicationcontent -install -cmd installfeatures -cmd secureresources -cmd services -install" -Wait -PassThru -NoNewWindow
                return $Process.ExitCode
            }
            Wait-JobWithWPF -Job $CAJobs -TimeoutMins 180 -ServersToCheck $SelectedCAs | Out-Null
            $CAResults = Receive-Job -Job $CAJobs
            foreach ($Res in $CAResults) {
                if ($Res.Value -eq 0) { Write-WpfLog "$($Res.PSComputerName): CA PSConfig Success" "#32CD32" }
                else { throw "$($Res.PSComputerName): CA PSConfig FAILED (Code: $($Res.Value))" }
            }
        }

        # 6. PSConfig (Others)
        Set-Progress 85 "Running PSConfig on App/WFE servers..." $true
        $OtherServers = $SelectedServers | Where-Object { $_ -notin $CAServers }

        if ($OtherServers.Count -gt 0) {
            Write-WpfLog "Starting PSConfig on remaining servers..." "#FFD700"
            $OtherJobs = Invoke-Command -ComputerName $OtherServers -AsJob -ScriptBlock {
                $Exe = "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\BIN\PSConfig.exe"
                $Process = Start-Process $Exe -ArgumentList "-cmd upgrade -inplace b2b -wait -cmd applicationcontent -install -cmd installfeatures -cmd secureresources -cmd services -install" -Wait -PassThru -NoNewWindow
                return $Process.ExitCode
            }
            Wait-JobWithWPF -Job $OtherJobs -TimeoutMins 180 -ServersToCheck $OtherServers | Out-Null
            $OtherRes = Receive-Job -Job $OtherJobs
            foreach ($Res in $OtherRes) {
                if ($Res.Value -eq 0) { Write-WpfLog "$($Res.PSComputerName): PSConfig Success" "#32CD32" }
                else { Write-WpfLog "$($Res.PSComputerName): PSConfig FAILED" "#FF5555" }
            }
        }
        Set-Progress 95 "PSConfig phase completed successfully." $false

        # 7. Archive Files
        Set-Progress 95 "Archiving processed updates in Storage..."
        $ArchiveFolderName = "Done_" + (Get-Date -Format "yyyy-MM")

        foreach ($Ver in $ServerMap.Keys) {
            if ($ServerMap[$Ver].Count -gt 0 -and $UpdateFiles[$Ver].Count -gt 0) {
                $VersionFolder = Join-Path -Path $Global:StorageBasePath -ChildPath $Ver
                $ArchiveFolderPath = Join-Path -Path $VersionFolder -ChildPath $ArchiveFolderName

                if (-not (Test-Path $ArchiveFolderPath)) { New-Item -ItemType Directory -Path $ArchiveFolderPath -Force | Out-Null }

                foreach ($FilePath in $UpdateFiles[$Ver]) {
                    if (Test-Path $FilePath) {
                        Write-WpfLog "Moving $(Split-Path $FilePath -Leaf) to $Ver\$ArchiveFolderName" "#D4D4D4"
                        Move-Item -Path $FilePath -Destination $ArchiveFolderPath -Force
                    }
                }
            }
        }

        # 8. Finish
        Set-Progress 100 "Starting Services and Cleanup..."
        Write-WpfLog "Restarting services globally..." "#FFD700"
        Invoke-Command -ComputerName $SelectedServers -ScriptBlock {
            Start-Service "SPTimerV4"
            iisreset /start | Out-Null
            Remove-Item "C:\SP_Updates_Temp" -Recurse -Force -ErrorAction SilentlyContinue
        }
        Write-WpfLog "--- ALL FARMS SUCCESSFULLY PATCHED & ARCHIVED! ---" "#00BFFF"

    } catch {
        Write-WpfLog "ERROR: $_" "#FF5555"
        Set-Progress 0 "Process Aborted." $false
    } finally {
        $BtnStart.IsEnabled = $true
        $CmbEnv.IsEnabled = $true
        $LstServers.IsEnabled = $true
    }
})

$Form.ShowDialog() | Out-Null
