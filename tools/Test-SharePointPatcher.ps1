<#
.SYNOPSIS
    Tests for the helper functions in SharePointPatcher.ps1.

.DESCRIPTION
    SharePointPatcher.ps1 builds its WPF window at load time, so it cannot be dot-sourced.
    This harness pulls the helper functions out of the file via the PowerShell parser,
    stubs the two UI calls they make, and exercises them against real background jobs.

    The point of T1 is to pin down the defect the rest of the fixes are built around: a
    scriptblock that returns a primitive comes back as a primitive, so reading a named
    property off it silently yields $null and every comparison against it is false.

.EXAMPLE
    pwsh -NoProfile -File tools/Test-SharePointPatcher.ps1
    powershell.exe -NoProfile -File tools\Test-SharePointPatcher.ps1
#>

$ScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "SharePointPatcher.ps1"

$Tokens = $null; $ParseErrors = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$Tokens, [ref]$ParseErrors)
if ($ParseErrors.Count -gt 0) {
    $ParseErrors | ForEach-Object { Write-Host "  parse error line $($_.Extent.StartLineNumber): $($_.Message)" }
    throw "$ScriptPath does not parse."
}
Write-Host "Parsed $ScriptPath cleanly ($($Tokens.Count) tokens)."

$Wanted = @("Wait-JobWithWPF", "Get-JobResultMap", "Initialize-LogFile")
$Funcs = $Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) |
    Where-Object { $Wanted -contains $_.Name }
foreach ($F in $Funcs) { . ([scriptblock]::Create($F.Extent.Text)) }
Write-Host "Loaded: $((($Funcs | ForEach-Object { $_.Name }) -join ', '))"

# The helpers log through the WPF surface; stub it out.
function DoEvents { }
$Global:LogLines = @()
function Write-WpfLog($Message, $Color = "LightGray") { $Global:LogLines += $Message }

$OnWindows = ($env:OS -eq "Windows_NT")
$TempRoot = [System.IO.Path]::GetTempPath()

$Pass = 0; $Fail = 0
function Check($Name, $Condition, $Detail = "") {
    if ($Condition) { $script:Pass++; Write-Host "  PASS  $Name" }
    else { $script:Fail++; Write-Host "  FAIL  $Name $Detail" }
}

Write-Host "`n=== T1: the defect itself - primitive vs PSCustomObject across the job boundary ==="
$JobPrim = Start-Job { return 3010 }
$JobObj = Start-Job { return [PSCustomObject]@{ ExitCode = 3010 } }
Wait-Job $JobPrim, $JobObj | Out-Null
$ResPrim = Receive-Job $JobPrim
$ResObj = Receive-Job $JobObj
Remove-Job $JobPrim, $JobObj -Force

Check "a returned [int] has no .Value member" ($null -eq $ResPrim.Value) "-> got '$($ResPrim.Value)'"
Check "so the old test '`$Res.Value -in @(0,3010)' is False even for exit code 3010" (-not ($ResPrim.Value -in @(0, 3010)))
Check "the exit code itself did survive the trip" ($ResPrim -eq 3010)
Check "a PSCustomObject keeps its ExitCode property" ($ResObj.ExitCode -eq 3010)
Check "so the fixed test '`$Res.ExitCode -in @(0,3010)' is True" ($ResObj.ExitCode -in @(0, 3010))

Write-Host "`n=== T2: a freshly created job must not be mistaken for a finished one ==="
$Jobs = 1..3 | ForEach-Object { Start-Job { Start-Sleep -Seconds 2; return [PSCustomObject]@{ Server = "S$using:_"; Ok = $true } } }
$NotStartedAtCall = @($Jobs | Where-Object { $_.State.ToString() -eq "NotStarted" }).Count
$OldLoopWouldExit = -not ($Jobs.State -contains "Running")
$R = Wait-JobWithWPF -Job $Jobs -TimeoutMins 2
Check "Wait-JobWithWPF returns true" ($R -eq $true)
Check "every job really is Completed by the time it returns" (@($Jobs | Where-Object { $_.State.ToString() -ne "Completed" }).Count -eq 0)
Write-Host "  (info) NotStarted at call time: $NotStartedAtCall; the old 'State -contains Running' loop would have exited immediately: $OldLoopWouldExit"
Remove-Job $Jobs -Force -ErrorAction SilentlyContinue

Write-Host "`n=== T3: the timeout path returns false and stops the jobs ==="
$Slow = @(Start-Job { Start-Sleep -Seconds 120 })
$Global:LogLines = @()
$R = Wait-JobWithWPF -Job $Slow -TimeoutMins 0.03
Check "returns `$false on timeout" ($R -eq $false)
Check "logs the timeout" (($Global:LogLines -join "|") -match "Timeout")
Start-Sleep -Milliseconds 800
Check "stops the job" ($Slow[0].State.ToString() -in @("Stopped", "Stopping"))
Remove-Job $Slow -Force -ErrorAction SilentlyContinue

Write-Host "`n=== T4: a failed job is terminal, and its missing result is caught downstream ==="
$Bad = @(Start-Job { throw "boom" })
$R = Wait-JobWithWPF -Job $Bad -TimeoutMins 1
Check "the loop exits on a Failed job instead of spinning to the timeout" ($R -eq $true)
Check "the job state is Failed" ($Bad[0].State.ToString() -eq "Failed")
$Threw = $false; $Msg = ""
try { Get-JobResultMap -Jobs $Bad -ExpectedKeys @("SRV1") -Label "Patch install" -KeyProperty "Server" | Out-Null }
catch { $Threw = $true; $Msg = "$_" }
Check "Get-JobResultMap throws for a server that produced no result" $Threw
Check "and names the server" ($Msg -match "SRV1") "-> '$Msg'"

Write-Host "`n=== T5: Get-JobResultMap happy path, keyed on a custom property ==="
$CopyJobs = @("SP19-PRD-APP1", "SPSE-PRD-WFE1") | ForEach-Object {
    Start-Job -ArgumentList $_ -ScriptBlock { param($S) [PSCustomObject]@{ Server = $S; Success = $true; Detail = "2 file(s)" } }
}
Wait-JobWithWPF -Job $CopyJobs -TimeoutMins 1 | Out-Null
$FirstId = $CopyJobs[0].Id
$Map = Get-JobResultMap -Jobs $CopyJobs -ExpectedKeys @("SP19-PRD-APP1", "SPSE-PRD-WFE1") -Label "File distribution" -KeyProperty "Server"
Check "both servers are in the map" ($Map.Count -eq 2)
Check "the mapped values carry their properties" ($Map["SP19-PRD-APP1"].Success -eq $true -and $Map["SPSE-PRD-WFE1"].Detail -eq "2 file(s)")
Check "lookup is case-insensitive, as server names from WinRM may differ in case" ($null -ne $Map["sp19-prd-app1"])
Check "the jobs are cleaned up after the receive" (@(Get-Job -Id $FirstId -ErrorAction SilentlyContinue).Count -eq 0)

Write-Host "`n=== T6: one silent server out of two is still a failure ==="
$Mixed = @(Start-Job { [PSCustomObject]@{ Server = "A"; Success = $true } })
Wait-JobWithWPF -Job $Mixed -TimeoutMins 1 | Out-Null
$Threw = $false; $Msg = ""
try { Get-JobResultMap -Jobs $Mixed -ExpectedKeys @("A", "B") -Label "Stop services" -KeyProperty "Server" | Out-Null }
catch { $Threw = $true; $Msg = "$_" }
Check "throws when only some servers answered" $Threw
Check "and names the silent one" ($Msg -match "B") "-> '$Msg'"

Write-Host "`n=== T7: PSConfig scheduling - Central Admin first, one server per farm per round ==="
# Mirrors the queue construction in section 5 of the patcher.
function Build-Rounds($ServerMap, $CAServers) {
    $Q = @{}; $Max = 0
    foreach ($Ver in @($ServerMap.Keys)) {
        $FarmServers = @($ServerMap[$Ver])
        if ($FarmServers.Count -eq 0) { continue }
        $FarmCAs = @($FarmServers | Where-Object { $_ -in $CAServers })
        $FarmOthers = @($FarmServers | Where-Object { $_ -notin $CAServers })
        $Q[$Ver] = @($FarmCAs + $FarmOthers)
        if ($Q[$Ver].Count -gt $Max) { $Max = $Q[$Ver].Count }
    }
    $Rounds = @()
    for ($r = 0; $r -lt $Max; $r++) {
        $RoundServers = @()
        foreach ($Ver in @($Q.Keys)) { if ($Q[$Ver].Count -gt $r) { $RoundServers += $Q[$Ver][$r] } }
        $Rounds += , @($RoundServers)
    }
    return , $Rounds
}

$ProdMap = @{ "SP2019" = @("SP19-PRD-APP1", "SP19-PRD-WFE1"); "SPSE" = @("SPSE-PRD-APP1", "SPSE-PRD-WFE1") }
$ProdCAs = @("SP19-PRD-APP1", "SPSE-PRD-APP1")
$Rounds = Build-Rounds $ProdMap $ProdCAs
Check "Prod takes 2 rounds" ($Rounds.Count -eq 2) "-> $($Rounds.Count)"
Check "round 1 is both Central Admin servers, one per farm" (((@($Rounds[0]) | Sort-Object) -join ",") -eq "SP19-PRD-APP1,SPSE-PRD-APP1") "-> $($Rounds[0] -join ',')"
Check "round 2 is both WFEs" (((@($Rounds[1]) | Sort-Object) -join ",") -eq "SP19-PRD-WFE1,SPSE-PRD-WFE1") "-> $($Rounds[1] -join ',')"
foreach ($Rd in $Rounds) {
    $PerFarm = @($Rd | Group-Object { if ($_ -like "SP19*") { "SP2019" } else { "SPSE" } })
    Check "no round runs two servers of the same farm" (@($PerFarm | Where-Object { $_.Count -gt 1 }).Count -eq 0)
}

$OneFarm = @{ "SPSE" = @("SPSE-A", "SPSE-B", "SPSE-C"); "SP2019" = @() }
$Rounds2 = Build-Rounds $OneFarm @("SPSE-B")
Check "a single farm of 3 servers serialises into 3 rounds" ($Rounds2.Count -eq 3) "-> $($Rounds2.Count)"
Check "its Central Admin server goes first" (((@($Rounds2[0])) -join ",") -eq "SPSE-B") "-> $($Rounds2[0] -join ',')"
Check "each of those rounds holds exactly one server" (@($Rounds2 | Where-Object { @($_).Count -ne 1 }).Count -eq 0)

Write-Host "`n=== T8: exit-code classification ==="
$SuccessExitCodes = @(0, 3010, 17022, 17025); $RebootExitCodes = @(3010, 17022)
Check "0 is success, no reboot" (($SuccessExitCodes -contains 0) -and -not ($RebootExitCodes -contains 0))
Check "3010 is success and flags a reboot" (($SuccessExitCodes -contains 3010) -and ($RebootExitCodes -contains 3010))
Check "17022 is success and flags a reboot" (($SuccessExitCodes -contains 17022) -and ($RebootExitCodes -contains 17022))
Check "17025 (already installed) is success, no reboot" (($SuccessExitCodes -contains 17025) -and -not ($RebootExitCodes -contains 17025))
Check "1603 is a failure" (-not ($SuccessExitCodes -contains 1603))
Check "-1 (installer missing on the server) is a failure" (-not ($SuccessExitCodes -contains -1))

Write-Host "`n=== T9: Initialize-LogFile falls back when the preferred path is unwritable ==="
$Unwritable = if ($OnWindows) { "Z:\no-such-share\patchlog.txt" } else { "/proc/no-such-dir/patchlog.txt" }
$Unwritable2 = if ($OnWindows) { "Y:\also-missing\patchlog.txt" } else { "/proc/also-missing/patchlog.txt" }
$Good = Join-Path -Path $TempRoot -ChildPath "sp-patcher-fallback-$PID.log"
$Chosen = Initialize-LogFile $Unwritable $Good
Check "falls back to the writable path" ($Chosen -eq $Good) "-> '$Chosen'"
Check "and the fallback file really exists" (Test-Path $Good)
Check "returns `$null when neither path is writable" ($null -eq (Initialize-LogFile $Unwritable $Unwritable2))
Remove-Item $Good -Force -ErrorAction SilentlyContinue

Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
Write-Host "`n==================================="
Write-Host "PASS: $Pass   FAIL: $Fail"
if ($Fail -gt 0) { exit 1 }
