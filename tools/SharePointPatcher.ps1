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

.NOTES
    Result contract for every remote scriptblock in this script:
    a scriptblock ALWAYS returns a [PSCustomObject] with named properties, never a bare
    [int]/[bool]. A primitive survives remoting as a primitive, so a client-side
    `$Result.SomeProperty` on it silently evaluates to $null and every comparison against
    it is false. Named properties on a PSCustomObject round-trip intact.
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
$Global:IsRunning = $false

# רק קבצים שבאמת מתקינים - כל דבר אחר בתיקייה (README, קובץ סימון, זיפ) לא ירוץ כהתקנה
$Global:InstallerExtensions = @(".exe", ".msu", ".msp")

# קודי יציאה שנחשבים הצלחה: 0 תקין, 3010/17022 דורש ריסטארט, 17025 העדכון כבר מותקן
$Global:SuccessExitCodes = @(0, 3010, 17022, 17025)
$Global:RebootExitCodes = @(3010, 17022)

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

# ה-FlowDocument יושב בתוך ה-RichTextBox; אם ה-FindName לא מוצא אותו ניקח אותו ישירות מהמאפיין
$LogDocument = $Form.FindName("LogDocument")
if (-not $LogDocument) { $LogDocument = $LogBox.Document }

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

# פותח את קובץ הלוג ומוודא שבאמת אפשר לכתוב אליו. אם השיתוף חסום נופלים לתיקייה מקומית,
# כדי שלא נגלה רק בדיעבד שאין שום תיעוד לריצה שהפילה חווה
function Initialize-LogFile($PreferredPath, $FallbackPath) {
    foreach ($Candidate in @($PreferredPath, $FallbackPath)) {
        if (-not $Candidate) { continue }
        try {
            $Folder = Split-Path -Path $Candidate -Parent
            if (-not (Test-Path $Folder)) {
                New-Item -ItemType Directory -Path $Folder -Force -ErrorAction Stop | Out-Null
            }
            Set-Content -Path $Candidate -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Log opened." -ErrorAction Stop
            return $Candidate
        } catch {
            continue
        }
    }
    return $null
}

function Wait-JobWithWPF($Job, $TimeoutMins, $ServersToCheck = $null) {
    if (-not $Job) { return $true }
    $StartTime = Get-Date
    $LastHeartbeat = Get-Date

    # רשימת תהליכים שאנחנו מחפשים בשרתים המרוחקים כדי לוודא שההתקנה עובדת
    $ProcNames = @("msiexec", "sts*", "uber*", "spserver*", "sp2019*", "psconfig", "psconfigui", "ose")

    # ג'וב שנעצר, נכשל או נחסם הוא גמור בדיוק כמו ג'וב שהסתיים. כל השאר (כולל NotStarted,
    # שהוא המצב של ג'וב טרי רגע אחרי היצירה) עדיין בתהליך ואסור לצאת עליו מהלולאה
    $TerminalStates = @("Completed", "Failed", "Stopped", "Blocked")

    # פרובה עם תקרת זמן: בלי זה WinRM מחכה בברירת מחדל שלוש דקות לשרת שעלה לריסטארט,
    # וכל הזמן הזה ה-UI קפוא ובדיקת הטיימאאוט לא מספיקה לרוץ
    $ProbeOptions = $null
    try { $ProbeOptions = New-PSSessionOption -OpenTimeout 5000 -OperationTimeout 15000 } catch { }

    while (@($Job | Where-Object { $TerminalStates -notcontains $_.State.ToString() }).Count -gt 0) {
        $Now = Get-Date

        if ($Now -gt $StartTime.AddMinutes($TimeoutMins)) {
            Write-WpfLog "Job Timeout Exceeded!" "#FF5555"
            $Job | Stop-Job -ErrorAction SilentlyContinue
            return $false
        }

        # דופק כל 60 שניות - בדיקה מול השרתים
        if (($Now - $LastHeartbeat).TotalSeconds -ge 60) {
            $ElapsedMins = [math]::Round(($Now - $StartTime).TotalMinutes, 1)

            if ($ServersToCheck) {
                # הדופק הוא תצוגה בלבד. אסור שכשל בניטור יפיל את ההרצה עצמה - ההתקנה
                # ממשיכה לרוץ על השרתים גם אם לא הצלחנו לשאול אותם מה שלומם
                try {
                    $CheckBlock = {
                        param($Names)
                        $procs = Get-Process -Name $Names -ErrorAction SilentlyContinue
                        # התיקון: החזרת אובייקט מפורש כדי שהסינון יעבוד
                        return [PSCustomObject]@{ IsActive = [bool]$procs }
                    }

                    $ProbeArgs = @{
                        ComputerName = $ServersToCheck
                        ScriptBlock  = $CheckBlock
                        ArgumentList = (, $ProcNames)
                        ErrorAction  = "SilentlyContinue"
                    }
                    if ($ProbeOptions) { $ProbeArgs["SessionOption"] = $ProbeOptions }
                    $ProcStatus = Invoke-Command @ProbeArgs

                    # התיקון: סינון לפי המאפיין IsActive
                    $ActiveServers = ($ProcStatus | Where-Object { $_.IsActive -eq $true }).PSComputerName -join ", "

                    # שרת שלא ענה הוא לא "שרת שסיים" - מפרידים בין השניים כדי לא להרגיע לשווא
                    $Answered = @($ProcStatus | ForEach-Object { $_.PSComputerName })
                    $Silent = @($ServersToCheck | Where-Object { $Answered -notcontains $_ })

                    if ($ActiveServers) {
                        Write-WpfLog "[Elapsed: $ElapsedMins Min] -> Install/Config actively running on: $ActiveServers" "#87CEFA"
                    } else {
                        Write-WpfLog "[Elapsed: $ElapsedMins Min] -> Background tasks finalizing..." "#87CEFA"
                    }

                    if ($Silent.Count -gt 0) {
                        Write-WpfLog "[Elapsed: $ElapsedMins Min] -> WARNING: no heartbeat response from: $($Silent -join ', ') (rebooting or unreachable)" "#FFD700"
                    }
                } catch {
                    Write-WpfLog "[Elapsed: $ElapsedMins Min] -> Heartbeat probe unavailable ($($_.Exception.Message)); the job itself is still running." "#FFD700"
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

# אוסף את תוצאות הג'ובים וממפה אותן לפי שרת. שרת שלא החזיר תוצאה (הג'וב נכשל, נעצר
# בטיימאאוט, או שה-WinRM נפל באמצע) הוא כישלון - לא "שרת שאין עליו מה לומר"
function Get-JobResultMap($Jobs, $ExpectedKeys, $Label, $KeyProperty = "PSComputerName") {
    $JobErrors = @()
    $Raw = Receive-Job -Job $Jobs -ErrorAction SilentlyContinue -ErrorVariable JobErrors
    Remove-Job -Job $Jobs -Force -ErrorAction SilentlyContinue

    foreach ($JobError in $JobErrors) {
        Write-WpfLog "$Label - remote error: $($JobError.Exception.Message)" "#FF5555"
    }

    $Map = @{}
    foreach ($Item in $Raw) {
        if ($null -eq $Item) { continue }
        $Key = $Item.$KeyProperty
        if ($Key) { $Map["$Key"] = $Item }
    }

    $Missing = @($ExpectedKeys | Where-Object { -not $Map.ContainsKey("$_") })
    if ($Missing.Count -gt 0) {
        throw "$Label - no result returned from: $($Missing -join ', '). Treating as failure."
    }

    return $Map
}

# מחזיר את השירותים לאוויר. נקרא גם מנתיב ההצלחה וגם מה-finally אחרי כישלון,
# כי הדבר הכי גרוע שהסקריפט יכול לעשות זה להיפול ולהשאיר חווה שלמה עם IIS מכובה
function Restore-FarmServices($Servers, $TempDir, $RemoveTemp) {
    Invoke-Command -ComputerName $Servers -ArgumentList $TempDir, $RemoveTemp -ErrorAction SilentlyContinue -ScriptBlock {
        param($T, $Cleanup)
        Start-Service "SPTimerV4" -ErrorAction SilentlyContinue
        iisreset /start | Out-Null
        if ($Cleanup -and $T) { Remove-Item $T -Recurse -Force -ErrorAction SilentlyContinue }
    }
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

# משאבת ה-DoEvents ממשיכה להעביר אירועים בזמן שההנדלר רץ, אז בלי המנעול הזה
# אפשר להיכנס לריצה שנייה במקביל לראשונה
$Form.add_Closing({
    param($EventSender, $EventArgs)
    if ($Global:IsRunning) {
        [System.Windows.MessageBox]::Show("A patch run is still in progress. Let it finish or fail before closing - closing now would leave the farm with its services stopped.", "Run In Progress", 0, 48)
        $EventArgs.Cancel = $true
    }
})

$BtnStart.add_Click({
    if ($Global:IsRunning) { return }

    $SelectedServers = @()
    $EnvName = $CmbEnv.SelectedItem
    $AllEnvServers = $Global:Environments[$EnvName].Servers
    $CAServers = $Global:Environments[$EnvName].CAServers

    foreach ($Item in $LstServers.Items) {
        if ($Item.IsChecked) { $SelectedServers += $Item.Tag }
    }

    if ($SelectedServers.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one server.", "Notice", 0, 48)
        return
    }

    $Global:IsRunning = $true
    $BtnStart.IsEnabled = $false
    $CmbEnv.IsEnabled = $false
    $LstServers.IsEnabled = $false

    $DateStamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $LogFileName = "PatchingLog_${EnvName}_$DateStamp.log"
    $Global:CurrentLogFile = Initialize-LogFile `
        (Join-Path -Path (Join-Path -Path $Global:StorageBasePath -ChildPath "Logs") -ChildPath $LogFileName) `
        (Join-Path -Path $env:TEMP -ChildPath $LogFileName)

    # דגלים לשחזור המצב: נקראים ב-finally, לכן חייבים להיווצר לפני ה-try
    $ServicesStopped = $false
    $RebootPending = @()
    $CompletedCleanly = $false

    try {
        Write-WpfLog "--- Initiating Multi-Farm Upgrade for $EnvName ---" "#00BFFF"
        if ($Global:CurrentLogFile) {
            Write-WpfLog "Log file established at: $($Global:CurrentLogFile)" "#D4D4D4"
        } else {
            Write-WpfLog "WARNING: no writable log file - this run will only be recorded on screen." "#FFD700"
        }

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

        # כל שרת שנבחר חייב להופיע במיפוי, אחרת נמשיך לתקן חווה חלקית בלי לשים לב
        $MappedServers = @($ServerMap["SP2019"] + $ServerMap["SPSE"])
        $UnmappedServers = @($SelectedServers | Where-Object { $MappedServers -notcontains $_ })
        if ($UnmappedServers.Count -gt 0) {
            throw "No version detected for: $($UnmappedServers -join ', ')"
        }

        Write-WpfLog "Scanning Storage subfolders under $($Global:StorageBasePath)..." "#FFD700"
        $UpdateFiles = @{ "SP2019" = @(); "SPSE" = @() }

        foreach ($Ver in @($ServerMap.Keys)) {
            if ($ServerMap[$Ver].Count -gt 0) {
                $VersionFolder = Join-Path -Path $Global:StorageBasePath -ChildPath $Ver
                if (-not (Test-Path $VersionFolder)) { New-Item -ItemType Directory -Path $VersionFolder -Force | Out-Null }

                # מסננים להתקנות בלבד וממיינים, כדי שסדר ההתקנה יהיה קבוע ולא תלוי בסדר שהשיתוף החזיר
                $FoundFiles = @(Get-ChildItem -Path $VersionFolder -File -ErrorAction SilentlyContinue |
                    Where-Object { $Global:InstallerExtensions -contains $_.Extension.ToLower() } |
                    Sort-Object Name)

                if ($FoundFiles.Count -eq 0) { throw "No installer files ($($Global:InstallerExtensions -join ', ')) found in: $VersionFolder" }

                $UpdateFiles[$Ver] = @($FoundFiles.FullName)
                Write-WpfLog "Found $($UpdateFiles[$Ver].Count) file(s) for $Ver in $VersionFolder" "#32CD32"
                foreach ($Name in $FoundFiles.Name) { Write-WpfLog "    $Ver <- $Name" "#D4D4D4" }
            }
        }

        Write-WpfLog "Validating disk space on all servers..." "#FFD700"
        Invoke-Command -ComputerName $SelectedServers -ErrorAction Stop -ScriptBlock {
            $Free = (Get-Volume -DriveLetter C).SizeRemaining / 1GB
            if ($Free -lt 15) { throw "Low disk space on $env:COMPUTERNAME" }
        }
        Write-WpfLog "Pre-Flight checks passed." "#32CD32"

        # 2. Copy Files
        # ההעתקה רצה כג'וב ולא בתוך הת'רד של ה-UI: העתקה של חבילת CU בגודל כמה ג'יגה
        # לכל שרת מקפיאה את החלון לדקות ארוכות, והמפעיל חושב שהכלי נתקע והורג אותו
        Set-Progress 25 "Distributing Update Files (Services running, no downtime yet)..."
        $CopyJobs = @()
        foreach ($Ver in @($ServerMap.Keys)) {
            foreach ($Srv in $ServerMap[$Ver]) {
                $Payload = [PSCustomObject]@{
                    Server  = $Srv
                    TempDir = $Global:LocalTempDir
                    Files   = @($UpdateFiles[$Ver])
                }
                $CopyJobs += Start-Job -ArgumentList $Payload -ScriptBlock {
                    param($P)
                    $Target = "\\$($P.Server)\" + $P.TempDir.Replace(':', '$')
                    try {
                        if (-not (Test-Path $Target)) {
                            New-Item -ItemType Directory -Path $Target -Force -ErrorAction Stop | Out-Null
                        }
                        foreach ($F in $P.Files) {
                            Copy-Item -Path $F -Destination $Target -Force -ErrorAction Stop
                        }
                        return [PSCustomObject]@{ Server = $P.Server; Success = $true; Detail = "$($P.Files.Count) file(s)" }
                    } catch {
                        return [PSCustomObject]@{ Server = $P.Server; Success = $false; Detail = $_.Exception.Message }
                    }
                }
            }
        }

        if (-not (Wait-JobWithWPF -Job $CopyJobs -TimeoutMins 90)) {
            throw "File distribution timed out after 90 minutes."
        }
        $CopyMap = Get-JobResultMap -Jobs $CopyJobs -ExpectedKeys $MappedServers -Label "File distribution" -KeyProperty "Server"
        foreach ($Srv in $MappedServers) {
            if ($CopyMap[$Srv].Success) {
                Write-WpfLog "Copied $($CopyMap[$Srv].Detail) to $Srv" "#32CD32"
            } else {
                throw "Copy to $Srv FAILED: $($CopyMap[$Srv].Detail)"
            }
        }
        Write-WpfLog "All files distributed successfully. Ready for patching." "#32CD32"

        # 3. Stop Services
        Set-Progress 40 "Stopping SharePoint Services..."
        Write-WpfLog "Stopping IIS and Timer Service globally..." "#FFD700"

        # מסמנים שהשירותים ירדו לפני ההרצה עצמה: אם ה-Invoke נופל באמצע, חלק מהשרתים
        # כבר כבויים וה-finally חייב לדעת שצריך להרים אותם בחזרה
        $ServicesStopped = $true
        $StopJobs = Invoke-Command -ComputerName $SelectedServers -AsJob -ScriptBlock {
            $IisOk = $true
            iisreset /stop | Out-Null
            if ($LASTEXITCODE -ne 0) { $IisOk = $false }

            $TimerOk = $true
            try {
                Stop-Service "SPTimerV4" -Force -ErrorAction Stop -WarningAction SilentlyContinue
                (Get-Service "SPTimerV4").WaitForStatus("Stopped", (New-TimeSpan -Seconds 120))
            } catch {
                $TimerOk = $false
            }
            return [PSCustomObject]@{ IisStopped = $IisOk; TimerStopped = $TimerOk }
        }

        if (-not (Wait-JobWithWPF -Job $StopJobs -TimeoutMins 15)) {
            throw "Stopping services timed out after 15 minutes."
        }
        $StopMap = Get-JobResultMap -Jobs $StopJobs -ExpectedKeys $SelectedServers -Label "Stop services"
        foreach ($Srv in $SelectedServers) {
            if (-not $StopMap[$Srv].IisStopped) { throw "$Srv : iisreset /stop failed - refusing to patch with IIS still holding SharePoint files." }
            if (-not $StopMap[$Srv].TimerStopped) { throw "$Srv : SPTimerV4 did not stop - refusing to patch with the Timer service running." }
        }
        Write-WpfLog "Services stopped and verified on all $($SelectedServers.Count) server(s)." "#32CD32"

        # 4. Patching (With Indeterminate Progress & Heartbeat)
        Set-Progress 50 "Installing Patches (This will take time)..." $true

        $MaxFiles = 0
        foreach ($Ver in @($ServerMap.Keys)) {
            if ($ServerMap[$Ver].Count -gt 0 -and $UpdateFiles[$Ver].Count -gt $MaxFiles) {
                $MaxFiles = $UpdateFiles[$Ver].Count
            }
        }

        for ($i = 0; $i -lt $MaxFiles; $i++) {
            $Jobs = @()
            $TargetServersForThisRun = @()

            foreach ($Ver in @($ServerMap.Keys)) {
                if ($ServerMap[$Ver].Count -gt 0 -and $UpdateFiles[$Ver].Count -gt $i) {
                    $File = Split-Path $UpdateFiles[$Ver][$i] -Leaf
                    Write-WpfLog "Initiating install of $File on $Ver servers..." "#FFD700"

                    $TargetServersForThisRun += $ServerMap[$Ver]
                    $Jobs += Invoke-Command -ComputerName $ServerMap[$Ver] -AsJob -ArgumentList $File, $Global:LocalTempDir -ScriptBlock {
                        param($F, $T)
                        $Path = Join-Path -Path $T -ChildPath $F
                        if (-not (Test-Path $Path)) {
                            return [PSCustomObject]@{ ExitCode = -1; File = $F; Detail = "Installer missing at $Path" }
                        }
                        $Process = Start-Process -FilePath $Path -ArgumentList "/passive /norestart" -Wait -PassThru
                        return [PSCustomObject]@{ ExitCode = [int]$Process.ExitCode; File = $F; Detail = "" }
                    }
                }
            }

            if ($Jobs.Count -gt 0) {
                # קריאה לפונקציית ההמתנה שכוללת את ההרטביט החדש (מעבירים את רשימת השרתים לבדיקה)
                if (-not (Wait-JobWithWPF -Job $Jobs -TimeoutMins 120 -ServersToCheck $TargetServersForThisRun)) {
                    throw "Patch install timed out after 120 minutes on: $($TargetServersForThisRun -join ', ')"
                }

                $PatchMap = Get-JobResultMap -Jobs $Jobs -ExpectedKeys $TargetServersForThisRun -Label "Patch install"
                foreach ($Srv in $TargetServersForThisRun) {
                    $Code = $PatchMap[$Srv].ExitCode
                    $Name = $PatchMap[$Srv].File

                    if ($Global:SuccessExitCodes -notcontains $Code) {
                        $Detail = $PatchMap[$Srv].Detail
                        if ($Detail) { throw "${Srv}: Patch FAILED (Code: $Code) - $Detail" }
                        throw "${Srv}: Patch FAILED (Code: $Code)"
                    }

                    if ($Code -eq 17025) {
                        Write-WpfLog "${Srv}: $Name already installed - skipped" "#FFD700"
                    } elseif ($Global:RebootExitCodes -contains $Code) {
                        if ($RebootPending -notcontains $Srv) { $RebootPending += $Srv }
                        Write-WpfLog "${Srv}: $Name installed (Code: $Code - reboot required)" "#FFD700"
                    } else {
                        Write-WpfLog "${Srv}: $Name installed" "#32CD32"
                    }
                }
            }
        }
        Set-Progress 65 "Patching phase completed successfully." $false

        # 5. PSConfig
        # שתי חוות רצות במקביל, אבל בתוך חווה אחת רק שרת אחד בכל פעם: PSConfig מריץ
        # שדרוג מול אותו Config DB, ושתי הרצות במקביל על אותה חווה מתנגשות זו בזו.
        # לכן בונים לכל חווה תור (Central Admin ראשון) ומתקדמים סבב-סבב.
        Set-Progress 70 "Running PSConfig (Central Admin first, one server at a time per farm)..." $true

        $PSConfigQueues = @{}
        $MaxRounds = 0
        foreach ($Ver in @($ServerMap.Keys)) {
            $FarmServers = @($ServerMap[$Ver])
            if ($FarmServers.Count -eq 0) { continue }

            $FarmCAs = @($FarmServers | Where-Object { $_ -in $CAServers })
            $FarmOthers = @($FarmServers | Where-Object { $_ -notin $CAServers })
            $PSConfigQueues[$Ver] = @($FarmCAs + $FarmOthers)

            if ($PSConfigQueues[$Ver].Count -gt $MaxRounds) { $MaxRounds = $PSConfigQueues[$Ver].Count }
        }

        for ($Round = 0; $Round -lt $MaxRounds; $Round++) {
            $RoundServers = @()
            foreach ($Ver in @($PSConfigQueues.Keys)) {
                if ($PSConfigQueues[$Ver].Count -gt $Round) { $RoundServers += $PSConfigQueues[$Ver][$Round] }
            }
            if ($RoundServers.Count -eq 0) { continue }

            $Percent = 70 + [int](20 * $Round / $MaxRounds)
            Set-Progress $Percent "PSConfig round $($Round + 1) of $MaxRounds..." $true
            Write-WpfLog "PSConfig round $($Round + 1)/$MaxRounds on: $($RoundServers -join ', ')" "#FFD700"

            $PSJobs = Invoke-Command -ComputerName $RoundServers -AsJob -ScriptBlock {
                $Exe = "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\BIN\PSConfig.exe"
                if (-not (Test-Path $Exe)) {
                    return [PSCustomObject]@{ ExitCode = -1; Detail = "PSConfig.exe not found at $Exe" }
                }
                $Process = Start-Process $Exe -ArgumentList "-cmd upgrade -inplace b2b -wait -cmd applicationcontent -install -cmd installfeatures -cmd secureresources -cmd services -install" -Wait -PassThru -NoNewWindow
                return [PSCustomObject]@{ ExitCode = [int]$Process.ExitCode; Detail = "" }
            }

            if (-not (Wait-JobWithWPF -Job $PSJobs -TimeoutMins 180 -ServersToCheck $RoundServers)) {
                throw "PSConfig timed out after 180 minutes on: $($RoundServers -join ', ')"
            }

            $PSMap = Get-JobResultMap -Jobs $PSJobs -ExpectedKeys $RoundServers -Label "PSConfig"
            foreach ($Srv in $RoundServers) {
                $Code = $PSMap[$Srv].ExitCode
                if ($Code -eq 0) {
                    Write-WpfLog "${Srv}: PSConfig Success" "#32CD32"
                } else {
                    # גם כישלון על WFE הוא כישלון: חווה עם שרת אחד לא משודרג היא חווה שבורה,
                    # ואסור להמשיך ממנה לארכוב ולהודעת ההצלחה
                    $Detail = $PSMap[$Srv].Detail
                    if ($Detail) { throw "${Srv}: PSConfig FAILED (Code: $Code) - $Detail" }
                    throw "${Srv}: PSConfig FAILED (Code: $Code)"
                }
            }
        }
        Set-Progress 90 "PSConfig phase completed successfully." $false

        # 6. Archive Files
        # מארכבים רק כשכל הסביבה עברה. בריצה חלקית השרתים שלא נבחרו עדיין צריכים
        # את הקבצים האלה, והזזתם לארכיון תשאיר אותם בלי מקור להתקנה
        Set-Progress 95 "Archiving processed updates in Storage..."
        if ($SelectedServers.Count -lt $AllEnvServers.Count) {
            Write-WpfLog "Partial run ($($SelectedServers.Count)/$($AllEnvServers.Count) servers) - keeping update files in place for the remaining servers." "#FFD700"
        } else {
            $ArchiveFolderName = "Done_" + (Get-Date -Format "yyyy-MM")

            foreach ($Ver in @($ServerMap.Keys)) {
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
        }

        # 7. Finish
        Set-Progress 100 "Starting Services and Cleanup..."
        Write-WpfLog "Restarting services globally..." "#FFD700"
        Restore-FarmServices -Servers $SelectedServers -TempDir $Global:LocalTempDir -RemoveTemp $true
        $ServicesStopped = $false
        $CompletedCleanly = $true

        Write-WpfLog "--- ALL FARMS SUCCESSFULLY PATCHED & ARCHIVED! ---" "#00BFFF"
        if ($RebootPending.Count -gt 0) {
            Write-WpfLog "REBOOT REQUIRED on: $($RebootPending -join ', ') - schedule a restart for these servers." "#FFD700"
        }

    } catch {
        Write-WpfLog "ERROR: $_" "#FF5555"
        Set-Progress 0 "Process Aborted." $false
    } finally {
        # שחזור: מרימים בחזרה כל שרת שהורדנו, גם אם נפלנו באמצע ההתקנה.
        # את התיקייה הזמנית משאירים אחרי כישלון כדי שאפשר יהיה לנסות שוב בלי להעתיק הכל מחדש
        if ($ServicesStopped) {
            try {
                Write-WpfLog "Restoring services after abort on: $($SelectedServers -join ', ')..." "#FFD700"
                Restore-FarmServices -Servers $SelectedServers -TempDir $Global:LocalTempDir -RemoveTemp $false
                Write-WpfLog "Services restored. The farm is serving again - the patch itself is INCOMPLETE." "#FFD700"
            } catch {
                Write-WpfLog "CRITICAL: could not restore services automatically ($_). Start SPTimerV4 and run 'iisreset /start' on $($SelectedServers -join ', ') MANUALLY." "#FF5555"
            }
            $ServicesStopped = $false
        }

        if (-not $CompletedCleanly) {
            Set-Progress 0 "Process Aborted - see log." $false
        }

        $Global:IsRunning = $false
        $BtnStart.IsEnabled = $true
        $CmbEnv.IsEnabled = $true
        $LstServers.IsEnabled = $true
    }
})

$Form.ShowDialog() | Out-Null
