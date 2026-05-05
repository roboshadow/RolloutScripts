#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Diagnose slow / failing Windows Update scans on managed (SCCM / WSUS) and unmanaged devices.

.DESCRIPTION
    Runs a battery of WUA (Windows Update Agent) scan scenarios with different combinations of
    ServerSelection / ServiceID / Online / IncludePotentiallySupersededUpdates / search criteria,
    each capped by a per-scenario timeout. Captures environment, registered update services,
    SCCM WMI fast-path results, and a tail of WUAHandler.log when present.

    Output: a single transcript log + a CSV summary in ProgramData. Send both back to support.

.PARAMETER ScenarioTimeoutSeconds
    Per-scenario timeout. Default 600s (10min). Tune if scenarios reliably complete faster.

.PARAMETER LogDirectory
    Where to write logs. Default: %ProgramData%\RoboShadow\WUA-Diag.

.PARAMETER OnlyScenarios
    Comma-separated list of scenario IDs to run (e.g. "A,E,H"). Default: run all.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Diagnose-WindowsUpdateScan.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Diagnose-WindowsUpdateScan.ps1 -ScenarioTimeoutSeconds 300 -OnlyScenarios "A,F,H"
#>
[CmdletBinding()]
param(
    [int]$ScenarioTimeoutSeconds = 600,
    [int]$AgentReproTimeoutSeconds = 1800,
    [string]$LogDirectory = "$env:ProgramData\RoboShadow\WUA-Diag",
    [string[]]$OnlyScenarios = @(),
    [switch]$SkipAgentRepro,
    [switch]$SkipScenarios
)

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'

# ---------------------------------------------------------------------------
# Setup logging
# ---------------------------------------------------------------------------
if (-not (Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}
$stamp      = Get-Date -Format 'yyyyMMdd-HHmmss'
$hostName   = $env:COMPUTERNAME
$logPath    = Join-Path $LogDirectory "WUA-Diag_${hostName}_${stamp}.log"
$csvPath    = Join-Path $LogDirectory "WUA-Diag_${hostName}_${stamp}_summary.csv"

Start-Transcript -Path $logPath -Force | Out-Null

function Write-Section($title) {
    Write-Host ""
    Write-Host ("=" * 80)
    Write-Host (" {0}" -f $title)
    Write-Host ("=" * 80)
}

function Write-Sub($title) {
    Write-Host ""
    Write-Host ("--- {0} ---" -f $title)
}

function Format-HResult($hr) {
    if ($null -eq $hr) { return "" }
    return ("0x{0:X8}" -f $hr)
}

Write-Section "Windows Update Scan Diagnostic"
Write-Host ("Host         : {0}" -f $hostName)
Write-Host ("User         : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
Write-Host ("Started      : {0}" -f (Get-Date -Format o))
Write-Host ("Log file     : {0}" -f $logPath)
Write-Host ("CSV summary  : {0}" -f $csvPath)
Write-Host ("Per-scenario timeout: {0}s" -f $ScenarioTimeoutSeconds)

# ---------------------------------------------------------------------------
# 1. PHASE 1 - Reproduce production agent behaviour end-to-end
# ---------------------------------------------------------------------------
# Runs FIRST, before any diagnostic queries, so the WUA agent state is as
# untouched as possible. Mirrors EXACTLY what the Rubicon.Thames agent does:
#   1. Create UpdateSession + CreateUpdateSearcher
#   2. searcher.IncludePotentiallySupersededUpdates = true
#   3. searcher.ServerSelection = ssWindowsUpdate (2)
#   4. searcher.Online = true
#   5. Search("IsInstalled=1")              <- GetInstalledUpdates
#   6. Search("IsHidden=0 AND IsInstalled=0") <- CheckForMissingUpdates
#   7. If non-English UI culture, repeat both searches with Session.UserLocale=1033
#      (the English-titles second pass in WindowsUpdatesSearcherWrapper).
# Each sub-step + total wall-clock is timed so we know exactly how slow prod is.
# ---------------------------------------------------------------------------
Write-Section "1. PHASE 1 - Production Agent Reproduction (Microsoft Update direct)"

$agentReproBlock = {
    param(
        [int]$ServerSelection = 2,    # 0=Default 1=ManagedServer 2=WindowsUpdate 3=Others
        [string]$ServiceID    = $null,
        [bool]$Online            = $true,
        [bool]$IncludeSuperseded = $true
    )

    # Faithful clone of WindowsUpdatesSearcherWrapper.Search:
    #   1. Localized Search() with the supplied searcher config
    #   2. First pass: collect localized titles (Title + Identity.UpdateID per item)
    #   3. If non-English: fresh Search() with UserLocale=1033 + collect English titles
    #   4. Final pass: iterate original collection, building UpdateEntry-equivalent
    #      (UpdateID, RevisionNumber, Title, KBArticleIDs[*], SupportUrl, IsInstalled,
    #       Categories[*].Name, IsMandatory, EnglishTitle) - this is the COM-marshalled
    #      hot loop that often dominates agent time when there are many updates.
    function Run-AgentSearch {
        param(
            [string]$Criteria,
            [int]$ServerSelection,
            [string]$ServiceID,
            [bool]$Online,
            [bool]$IncludeSuperseded
        )
        $r = [pscustomobject]@{
            Criteria        = $Criteria
            SearchMs        = $null
            TitlePassMs     = $null
            EnSearchMs      = $null
            EnTitlePassMs   = $null
            EnumPassMs      = $null
            UpdateCount     = $null
            EnglishCount    = $null
            ResultCode      = $null
            CategoryCalls   = 0
            KbCalls         = 0
            SampleTitles    = $null
        }

        # ---- Localized search ----
        $session  = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $searcher.IncludePotentiallySupersededUpdates = $IncludeSuperseded
        $searcher.ServerSelection = $ServerSelection
        if ($ServerSelection -eq 3 -and $ServiceID) { $searcher.ServiceID = $ServiceID }
        $searcher.Online          = $Online

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $res = $searcher.Search($Criteria)
        $sw.Stop()
        $r.SearchMs    = $sw.ElapsedMilliseconds
        $r.UpdateCount = $res.Updates.Count
        $r.ResultCode  = $res.ResultCode

        # ---- First pass: collect localized titles ----
        $localizedTitles = @{}
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        for ($i = 0; $i -lt $res.Updates.Count; $i++) {
            $u = $res.Updates.Item($i)
            $localizedTitles[$u.Identity.UpdateID] = $u.Title
        }
        $sw.Stop()
        $r.TitlePassMs = $sw.ElapsedMilliseconds

        # ---- English-titles second pass ----
        $lcid = [System.Globalization.CultureInfo]::CurrentUICulture.LCID
        $englishTitles = @{}
        if ($lcid -ne 1033) {
            $enSession = New-Object -ComObject Microsoft.Update.Session
            $enSession.UserLocale = 1033
            $enSearcher = $enSession.CreateUpdateSearcher()
            $enSearcher.IncludePotentiallySupersededUpdates = $IncludeSuperseded
            $enSearcher.ServerSelection = $ServerSelection
            if ($ServerSelection -eq 3 -and $ServiceID) { $enSearcher.ServiceID = $ServiceID }
            $enSearcher.Online          = $Online
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $enRes = $enSearcher.Search($Criteria)
            $sw.Stop()
            $r.EnSearchMs   = $sw.ElapsedMilliseconds
            $r.EnglishCount = $enRes.Updates.Count

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            for ($i = 0; $i -lt $enRes.Updates.Count; $i++) {
                $u = $enRes.Updates.Item($i)
                $englishTitles[$u.Identity.UpdateID] = $u.Title
            }
            $sw.Stop()
            $r.EnTitlePassMs = $sw.ElapsedMilliseconds
        } else {
            $englishTitles = $localizedTitles
        }

        # ---- Final pass: build UpdateEntry-equivalent per item (wrapper line 143-172) ----
        # This reads many COM properties per update; on managed devices with hundreds
        # of approved updates, this can dominate the wall-clock time.
        $sampleTitles = New-Object System.Collections.Generic.List[string]
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        for ($i = 0; $i -lt $res.Updates.Count; $i++) {
            $u   = $res.Updates.Item($i)
            $id  = $u.Identity.UpdateID
            $rev = $u.Identity.RevisionNumber
            $sup = $u.SupportUrl
            $inst= $u.IsInstalled
            $mand= $u.IsMandatory
            $localized = $localizedTitles[$id]
            $english   = $englishTitles[$id]

            if ($u.KBArticleIDs) {
                for ($k = 0; $k -lt $u.KBArticleIDs.Count; $k++) {
                    $null = $u.KBArticleIDs.Item($k)
                    $r.KbCalls++
                }
            }
            if ($u.Categories) {
                for ($c = 0; $c -lt $u.Categories.Count; $c++) {
                    $null = $u.Categories.Item($c).Name
                    $r.CategoryCalls++
                }
            }
            if ($i -lt 5) { $sampleTitles.Add($localized) | Out-Null }
        }
        $sw.Stop()
        $r.EnumPassMs = $sw.ElapsedMilliseconds
        $r.SampleTitles = ($sampleTitles -join ' | ')

        return $r
    }

    $result = [pscustomobject]@{
        Status          = 'Unknown'
        TotalMs         = 0
        UICulture       = $null
        UICultureLcid   = $null
        EnglishRepeated = $false
        ErrorPhase      = $null
        ErrorMessage    = $null
        HResult         = $null
        Installed       = $null
        Missing         = $null
    }
    $totalSw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $result.UICulture       = [System.Globalization.CultureInfo]::CurrentUICulture.Name
        $result.UICultureLcid   = [System.Globalization.CultureInfo]::CurrentUICulture.LCID
        $result.EnglishRepeated = ($result.UICultureLcid -ne 1033)

        $result.ErrorPhase = '0.1 GetInstalledUpdates (Search + enumerate)'
        $result.Installed  = Run-AgentSearch -Criteria "IsInstalled=1" `
            -ServerSelection $ServerSelection -ServiceID $ServiceID -Online $Online -IncludeSuperseded $IncludeSuperseded

        $result.ErrorPhase = '0.2 CheckForMissingUpdates (Search + enumerate)'
        $result.Missing    = Run-AgentSearch -Criteria "IsHidden=0 AND IsInstalled=0" `
            -ServerSelection $ServerSelection -ServiceID $ServiceID -Online $Online -IncludeSuperseded $IncludeSuperseded

        $result.ErrorPhase = $null
        $result.Status     = 'Succeeded'
    }
    catch {
        $result.Status       = 'Exception'
        $result.ErrorMessage = $_.Exception.Message
        if ($_.Exception.InnerException) {
            $result.HResult = $_.Exception.InnerException.HResult
        } else {
            $result.HResult = $_.Exception.HResult
        }
    }
    finally {
        $totalSw.Stop()
        $result.TotalMs = $totalSw.ElapsedMilliseconds
    }
    return $result
}

function Write-AgentReproDetail($label, $detail) {
    if ($null -eq $detail) { return }
    Write-Host ("  {0}" -f $label)
    Write-Host ("    Criteria             : {0}" -f $detail.Criteria)
    Write-Host ("    Search()             : {0,8:N0} ms  ({1} updates, ResultCode={2})" -f $detail.SearchMs, $detail.UpdateCount, $detail.ResultCode)
    Write-Host ("    Title pass (loc.)    : {0,8:N0} ms" -f $detail.TitlePassMs)
    if ($null -ne $detail.EnSearchMs) {
        Write-Host ("    Search() en-US       : {0,8:N0} ms  ({1} updates)" -f $detail.EnSearchMs, $detail.EnglishCount)
        Write-Host ("    Title pass (en)      : {0,8:N0} ms" -f $detail.EnTitlePassMs)
    }
    Write-Host ("    Enumerate updates    : {0,8:N0} ms  ({1} category reads, {2} KB reads)" -f $detail.EnumPassMs, $detail.CategoryCalls, $detail.KbCalls)
    if ($detail.SampleTitles) {
        Write-Host ("    Sample titles        : {0}" -f $detail.SampleTitles)
    }
}

if ($SkipAgentRepro) {
    Write-Host "Skipped (-SkipAgentRepro)."
    $agentRepro = $null
} else {
    Write-Host "Reproducing the EXACT agent flow (WindowsUpdatePatchScanner + WindowsUpdatesSearcherWrapper):"
    Write-Host "  Searcher: IncludePotentiallySupersededUpdates=TRUE, ServerSelection=ssWindowsUpdate, Online=TRUE"
    Write-Host "  For each of the two criteria below:"
    Write-Host "    1. Search()"
    Write-Host "    2. First pass: collect localized titles"
    Write-Host "    3. (non-English only) Search() with UserLocale=1033 + collect English titles"
    Write-Host "    4. Final pass: read UpdateID/Title/RevisionNumber/Categories[*]/KBArticleIDs[*]/SupportUrl/IsInstalled/IsMandatory per update"
    Write-Host "  Criteria (1) IsInstalled=1                 -> GetInstalledUpdates"
    Write-Host "  Criteria (2) IsHidden=0 AND IsInstalled=0  -> CheckForMissingUpdates"
    Write-Host ""
    Write-Host ("Timeout: {0}s ({1} min). Started: {2}" -f $AgentReproTimeoutSeconds, [int]($AgentReproTimeoutSeconds/60), (Get-Date -Format o))

    # Production config: ssWindowsUpdate (2), no ServiceID, Online=true, IncludeSuperseded=true
    $reproJob = Start-Job -ScriptBlock $agentReproBlock -ArgumentList @(2, $null, $true, $true)
    $finished = Wait-Job -Job $reproJob -Timeout $AgentReproTimeoutSeconds
    if (-not $finished) {
        Write-Host ("PHASE 1 TIMEOUT after {0}s - this confirms production-equivalent scan is hanging." -f $AgentReproTimeoutSeconds)
        Stop-Job  -Job $reproJob -ErrorAction SilentlyContinue
        Remove-Job -Job $reproJob -Force -ErrorAction SilentlyContinue
        $agentRepro = [pscustomobject]@{
            Status          = 'TIMEOUT'
            TotalMs         = $AgentReproTimeoutSeconds * 1000
            UICulture       = $null
            UICultureLcid   = $null
            EnglishRepeated = $null
            ErrorPhase      = $null
            ErrorMessage    = "Exceeded $AgentReproTimeoutSeconds s"
            HResult         = $null
            Installed       = $null
            Missing         = $null
        }
    } else {
        $agentRepro = Receive-Job -Job $reproJob
        Remove-Job -Job $reproJob -Force -ErrorAction SilentlyContinue
    }

    Write-Host ""
    Write-Host "----- AGENT REPRODUCTION RESULT -----"
    Write-Host ("Status                : {0}" -f $agentRepro.Status)
    Write-Host ("UI culture            : {0} (LCID {1})" -f $agentRepro.UICulture, $agentRepro.UICultureLcid)
    Write-Host ("English repeat needed : {0}" -f $agentRepro.EnglishRepeated)
    Write-Host ("Total wall-clock time : {0:N0} ms  ({1:N1} s)" -f $agentRepro.TotalMs, ($agentRepro.TotalMs/1000))
    Write-Host ""
    Write-AgentReproDetail "0.1 GetInstalledUpdates"      $agentRepro.Installed
    Write-Host ""
    Write-AgentReproDetail "0.2 CheckForMissingUpdates"   $agentRepro.Missing
    if ($agentRepro.HResult)      { Write-Host ""; Write-Host ("HResult               : {0}" -f (Format-HResult $agentRepro.HResult)) }
    if ($agentRepro.ErrorPhase)   { Write-Host ("Failed in phase       : {0}" -f $agentRepro.ErrorPhase) }
    if ($agentRepro.ErrorMessage) { Write-Host ("Error                 : {0}" -f $agentRepro.ErrorMessage) }
}

# ---------------------------------------------------------------------------
# 2. PHASE 2 - Same agent flow but via the internal/managed source
# ---------------------------------------------------------------------------
# Detect management mode in-process, then re-run the EXACT same agent flow
# (Search + title pass + final enumeration) against the managed source so we
# can compare apples-to-apples vs Phase 1.
#   SCCM detected      -> ssOthers + ServiceID=<discovered SUP> + Online=false
#   WSUS only          -> ssManagedServer                       + Online=false
#   Unmanaged          -> ssDefault (let WUA pick)              + Online=false
# ---------------------------------------------------------------------------
Write-Section "2. PHASE 2 - Same flow via Internal Network / Managed Source"

# In-process detection (fast, doesn't significantly disturb WUA state).
$phase2Mode      = $null
$phase2ServerSel = $null
$phase2ServiceID = $null
$phase2Note      = $null

# 1. SCCM client detection
$ph2_ccmService = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
$ph2_ccmWmi     = Get-CimInstance -Namespace 'root\ccm' -ClassName '__Namespace' -ErrorAction SilentlyContinue
$ph2_isSccm     = ($ph2_ccmService -and $ph2_ccmService.Status -eq 'Running') -or $ph2_ccmWmi

# 2. WSUS GPO detection
$ph2_useWuServer = 0
try {
    $ph2_useWuServer = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -Name UseWUServer -ErrorAction Stop).UseWUServer
} catch { }

# 3. Find the managed update service (SCCM SUP) via IUpdateServiceManager
$ph2_managedService = $null
try {
    $ph2_sm = New-Object -ComObject Microsoft.Update.ServiceManager
    foreach ($svc in $ph2_sm.Services) {
        # Skip the public default services - we want the customer's managed source.
        if ($svc.ServiceID -in @('9482F4B4-E343-43B6-B170-9A65BC822C77', '7971F918-A847-4430-9279-4A52D1EFE18D', '855E8A7C-ECB4-4CA3-B045-1DFA50104289', '8B24B027-1DEE-BABB-9A95-3517DFB9C552')) {
            continue
        }
        if ($svc.IsManaged -or $svc.IsRegisteredWithAU) {
            $ph2_managedService = $svc
            break
        }
    }
} catch { }

# Decide branch
if ($ph2_isSccm -and $ph2_managedService) {
    $phase2Mode      = "SCCM-managed (ServiceID={0} '{1}')" -f $ph2_managedService.ServiceID, $ph2_managedService.Name
    $phase2ServerSel = 3   # ssOthers
    $phase2ServiceID = $ph2_managedService.ServiceID
    $phase2Note      = "Using SCCM SUP via ssOthers + ServiceID. Online=false to consume the cache CcmExec already populated."
} elseif ($ph2_isSccm -and -not $ph2_managedService) {
    $phase2Mode      = "SCCM client present but no managed service registered"
    $phase2ServerSel = 1   # ssManagedServer (best guess)
    $phase2ServiceID = $null
    $phase2Note      = "WARNING: SCCM client detected but no managed service via IUpdateServiceManager. Trying ssManagedServer as fallback."
} elseif ($ph2_useWuServer -eq 1 -and $ph2_managedService) {
    $phase2Mode      = "WSUS-managed (ServiceID={0} '{1}')" -f $ph2_managedService.ServiceID, $ph2_managedService.Name
    $phase2ServerSel = 1   # ssManagedServer
    $phase2ServiceID = $null
    $phase2Note      = "Using WSUS via ssManagedServer + Online=false."
} elseif ($ph2_useWuServer -eq 1) {
    $phase2Mode      = "WSUS GPO set but no managed service registered"
    $phase2ServerSel = 1
    $phase2ServiceID = $null
    $phase2Note      = "Using ssManagedServer (no specific ServiceID found)."
} else {
    $phase2Mode      = "Unmanaged (no SCCM, no WSUS GPO)"
    $phase2ServerSel = 0   # ssDefault
    $phase2ServiceID = $null
    $phase2Note      = "No internal/managed source detected. Running same flow with ssDefault + Online=false (best fallback for unmanaged devices)."
}

Write-Host ("Detected mode : {0}" -f $phase2Mode)
Write-Host ("Branch        : ServerSelection={0}  ServiceID={1}  Online=false  IncludeSuperseded=false" -f $phase2ServerSel, $phase2ServiceID)
Write-Host ("Note          : {0}" -f $phase2Note)
Write-Host ""

if ($SkipAgentRepro) {
    Write-Host "Skipped (-SkipAgentRepro)."
    $agentReproManaged = $null
} else {
    Write-Host ("Timeout: {0}s ({1} min). Started: {2}" -f $AgentReproTimeoutSeconds, [int]($AgentReproTimeoutSeconds/60), (Get-Date -Format o))

    $reproJob2 = Start-Job -ScriptBlock $agentReproBlock -ArgumentList @($phase2ServerSel, $phase2ServiceID, $false, $false)
    $finished2 = Wait-Job -Job $reproJob2 -Timeout $AgentReproTimeoutSeconds
    if (-not $finished2) {
        Write-Host ("PHASE 2 TIMEOUT after {0}s." -f $AgentReproTimeoutSeconds)
        Stop-Job  -Job $reproJob2 -ErrorAction SilentlyContinue
        Remove-Job -Job $reproJob2 -Force -ErrorAction SilentlyContinue
        $agentReproManaged = [pscustomobject]@{
            Status          = 'TIMEOUT'
            TotalMs         = $AgentReproTimeoutSeconds * 1000
            UICulture       = $null
            UICultureLcid   = $null
            EnglishRepeated = $null
            ErrorPhase      = $null
            ErrorMessage    = "Exceeded $AgentReproTimeoutSeconds s"
            HResult         = $null
            Installed       = $null
            Missing         = $null
        }
    } else {
        $agentReproManaged = Receive-Job -Job $reproJob2
        Remove-Job -Job $reproJob2 -Force -ErrorAction SilentlyContinue
    }

    Write-Host ""
    Write-Host "----- MANAGED-SOURCE REPRODUCTION RESULT -----"
    Write-Host ("Status                : {0}" -f $agentReproManaged.Status)
    Write-Host ("Total wall-clock time : {0:N0} ms  ({1:N1} s)" -f $agentReproManaged.TotalMs, ($agentReproManaged.TotalMs/1000))
    Write-Host ""
    Write-AgentReproDetail "2.1 GetInstalledUpdates"      $agentReproManaged.Installed
    Write-Host ""
    Write-AgentReproDetail "2.2 CheckForMissingUpdates"   $agentReproManaged.Missing
    if ($agentReproManaged.HResult)      { Write-Host ""; Write-Host ("HResult               : {0}" -f (Format-HResult $agentReproManaged.HResult)) }
    if ($agentReproManaged.ErrorPhase)   { Write-Host ("Failed in phase       : {0}" -f $agentReproManaged.ErrorPhase) }
    if ($agentReproManaged.ErrorMessage) { Write-Host ("Error                 : {0}" -f $agentReproManaged.ErrorMessage) }

    if ($agentRepro -and $agentReproManaged -and
        $agentRepro.Status -ne 'TIMEOUT' -and $agentReproManaged.Status -ne 'TIMEOUT') {
        $delta = $agentRepro.TotalMs - $agentReproManaged.TotalMs
        $pct   = if ($agentRepro.TotalMs -gt 0) { ($delta / $agentRepro.TotalMs) * 100 } else { 0 }
        Write-Host ""
        Write-Host ("Phase1 vs Phase2      : Phase1 {0:N0} ms  vs  Phase2 {1:N0} ms" -f $agentRepro.TotalMs, $agentReproManaged.TotalMs)
        if ($delta -gt 0) {
            Write-Host ("                        Phase 2 is {0:N0} ms faster ({1:N1}% saving)" -f $delta, $pct)
        } elseif ($delta -lt 0) {
            Write-Host ("                        Phase 2 is {0:N0} ms SLOWER" -f [Math]::Abs($delta))
        } else {
            Write-Host "                        No difference."
        }
    }
}

# ---------------------------------------------------------------------------
# 3. Environment detection
# ---------------------------------------------------------------------------
Write-Section "3. Environment"

Write-Sub "OS"
try {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    Write-Host ("OS         : {0}  ({1})" -f $os.Caption, $os.Version)
    Write-Host ("Build      : {0}" -f $os.BuildNumber)
    Write-Host ("Last boot  : {0}" -f $os.LastBootUpTime)
} catch { Write-Host "Failed to query OS: $_" }

Write-Sub "WUA-relevant services"
$svcNames = 'wuauserv','BITS','UsoSvc','TrustedInstaller','CcmExec','smstsmgr','ccmsetup'
foreach ($n in $svcNames) {
    $s = Get-Service -Name $n -ErrorAction SilentlyContinue
    if ($s) {
        Write-Host ("  {0,-18} {1,-10} StartType={2}" -f $s.Name, $s.Status, $s.StartType)
    } else {
        Write-Host ("  {0,-18} <not installed>" -f $n)
    }
}

Write-Sub "WindowsUpdate policy registry"
$polPaths = @(
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
)
foreach ($p in $polPaths) {
    if (Test-Path $p) {
        Write-Host ("[{0}]" -f $p)
        try {
            Get-ItemProperty -Path $p | Select-Object -Property * -ExcludeProperty PS* |
                Format-List | Out-String | Write-Host
        } catch { Write-Host "  (failed to read: $_)" }
    } else {
        Write-Host ("[{0}] - not present" -f $p)
    }
}

Write-Sub "SCCM client presence"
$ccmPresent = $false
$ccmExec = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
if ($ccmExec) {
    $ccmPresent = $true
    Write-Host ("CcmExec service: {0}" -f $ccmExec.Status)
}
$ccmWmi = Get-CimInstance -Namespace 'root\ccm' -ClassName '__Namespace' -ErrorAction SilentlyContinue
if ($ccmWmi) {
    $ccmPresent = $true
    Write-Host "WMI namespace root\ccm is present"
}
if (-not $ccmPresent) { Write-Host "No SCCM client detected." }

# ---------------------------------------------------------------------------
# 4. Registered Windows Update services (IUpdateServiceManager.Services)
# ---------------------------------------------------------------------------
Write-Section "4. Registered Update Services (IUpdateServiceManager)"

$registeredServices = @()
try {
    $sm = New-Object -ComObject Microsoft.Update.ServiceManager
    foreach ($svc in $sm.Services) {
        $obj = [pscustomobject]@{
            Name              = $svc.Name
            ServiceID         = $svc.ServiceID
            IsManaged         = $svc.IsManaged
            IsDefaultAUService= $svc.IsDefaultAUService
            IsRegisteredWith  = $svc.IsRegisteredWithAU
            OffersWindowsUpd  = $svc.OffersWindowsUpdates
            IssueDate         = $svc.IssueDate
        }
        $registeredServices += $obj
    }
    $registeredServices | Format-Table -AutoSize | Out-String | Write-Host
} catch {
    Write-Host "Failed to enumerate services: $_"
}

# Pick a likely SCCM SUP service (managed, not the default WU)
$sccmService = $registeredServices | Where-Object {
    $_.IsManaged -and $_.ServiceID -ne '9482F4B4-E343-43B6-B170-9A65BC822C77' -and $_.ServiceID -ne '7971F918-A847-4430-9279-4A52D1EFE18D'
} | Select-Object -First 1
# Note: 9482F4B4... is Windows Update; 7971F918... is Microsoft Update.
# An SCCM SUP registers a per-environment GUID different from these.
if (-not $sccmService) {
    # fall back to any managed service that isn't the public WU
    $sccmService = $registeredServices | Where-Object { $_.IsManaged } | Select-Object -First 1
}
if ($sccmService) {
    Write-Host ("Detected managed service: {0}  ({1})" -f $sccmService.Name, $sccmService.ServiceID)
} else {
    Write-Host "No managed (SCCM/WSUS) update service found."
}

# ---------------------------------------------------------------------------
# 5. SCCM WMI fast path  (root\ccm\ClientSDK\CCM_SoftwareUpdate)
# ---------------------------------------------------------------------------
Write-Section "5. SCCM WMI fast path (CCM_SoftwareUpdate)"

if ($ccmPresent) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $missing = Get-CimInstance -Namespace 'root\ccm\ClientSDK' -ClassName 'CCM_SoftwareUpdate' -ErrorAction Stop
        $sw.Stop()
        Write-Host ("Query time      : {0} ms" -f $sw.ElapsedMilliseconds)
        Write-Host ("Updates returned: {0}" -f ($missing | Measure-Object).Count)
        $missing |
            Select-Object -First 10 ArticleID,Name,EvaluationState,ComplianceState,Deadline |
            Format-Table -AutoSize | Out-String | Write-Host
    } catch {
        $sw.Stop()
        Write-Host ("CCM_SoftwareUpdate query FAILED after {0} ms: {1}" -f $sw.ElapsedMilliseconds, $_.Exception.Message)
    }
} else {
    Write-Host "Skipped (no SCCM client)."
}

# ---------------------------------------------------------------------------
# 6. Tail SCCM WUAHandler.log if present
# ---------------------------------------------------------------------------
Write-Section "6. WUAHandler.log (last 80 lines)"
$wuaHandlerCandidates = @(
    "$env:windir\CCM\Logs\WUAHandler.log",
    "$env:windir\SysWOW64\CCM\Logs\WUAHandler.log"
)
$wuaHandler = $wuaHandlerCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($wuaHandler) {
    Write-Host ("Source: {0}" -f $wuaHandler)
    Get-Content -Path $wuaHandler -Tail 80 | Out-String | Write-Host
} else {
    Write-Host "WUAHandler.log not found."
}

# ---------------------------------------------------------------------------
# 7. Scenarios - alternative configurations to compare against Phase 1 / Phase 2
# ---------------------------------------------------------------------------
Write-Section "7. WUA Scan Scenarios (alternatives)"

$results = New-Object System.Collections.Generic.List[object]

if ($SkipScenarios) {
    Write-Host "Skipped (-SkipScenarios)."
} else {

# Each scenario is run in its own background job (fresh runspace + COM apartment)
# so that we can enforce a wall-clock timeout. COM objects cannot cross runspaces,
# so the searcher is created INSIDE the scriptblock.

$scenarioBlock = {
    param(
        [string]$Id,
        [string]$Description,
        [int]$ServerSelection,         # 0=Default, 1=ManagedServer, 2=WindowsUpdate, 3=Others
        [string]$ServiceID,            # used when ServerSelection=3
        [bool]$Online,
        [bool]$IncludeSuperseded,
        [string]$Criteria
    )

    $result = [pscustomobject]@{
        Id            = $Id
        Description   = $Description
        ServerSel     = $ServerSelection
        ServiceID     = $ServiceID
        Online        = $Online
        IncludeSup    = $IncludeSuperseded
        Criteria      = $Criteria
        DurationMs    = 0
        ResultCode    = $null
        UpdateCount   = $null
        FirstTitles   = $null
        Status        = 'Unknown'
        HResult       = $null
        ErrorMessage  = $null
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $session  = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $searcher.IncludePotentiallySupersededUpdates = $IncludeSuperseded
        $searcher.ServerSelection = $ServerSelection
        if ($ServerSelection -eq 3 -and $ServiceID) {
            $searcher.ServiceID = $ServiceID
        }
        $searcher.Online = $Online

        $r = $searcher.Search($Criteria)
        $sw.Stop()

        $result.DurationMs  = $sw.ElapsedMilliseconds
        $result.ResultCode  = $r.ResultCode    # 1=InProgress 2=Succeeded 3=SucceededWithErrors 4=Failed 5=Aborted
        $result.UpdateCount = $r.Updates.Count
        $titles = @()
        for ($i = 0; $i -lt [Math]::Min(5, $r.Updates.Count); $i++) {
            $titles += $r.Updates.Item($i).Title
        }
        $result.FirstTitles = ($titles -join ' | ')
        $result.Status      = switch ($r.ResultCode) { 2 {'Succeeded'} 3 {'SucceededWithErrors'} 4 {'Failed'} 5 {'Aborted'} default {"ResultCode=$($r.ResultCode)"} }
    }
    catch {
        $sw.Stop()
        $result.DurationMs   = $sw.ElapsedMilliseconds
        $result.Status       = 'Exception'
        $result.ErrorMessage = $_.Exception.Message
        if ($_.Exception.InnerException) {
            $result.HResult = $_.Exception.InnerException.HResult
        } else {
            $result.HResult = $_.Exception.HResult
        }
    }
    return $result
}

# Build scenario list
$scenarios = New-Object System.Collections.Generic.List[object]

# A: Reproduce CURRENT production behaviour (baseline)
$scenarios.Add(@{
    Id='A'; Description='CURRENT PROD: ssWindowsUpdate + Online=true + IncludeSuperseded=true + missing-only criteria'
    ServerSelection=2; ServiceID=$null; Online=$true; IncludeSuperseded=$true
    Criteria="IsHidden=0 AND IsInstalled=0"
}) | Out-Null

# B: Default selection, online
$scenarios.Add(@{
    Id='B'; Description='ssDefault + Online=true (let WUA pick managed source)'
    ServerSelection=0; ServiceID=$null; Online=$true; IncludeSuperseded=$false
    Criteria="IsHidden=0 AND IsInstalled=0"
}) | Out-Null

# C: Default selection, offline (cached metadata)
$scenarios.Add(@{
    Id='C'; Description='ssDefault + Online=false (use cached metadata)'
    ServerSelection=0; ServiceID=$null; Online=$false; IncludeSuperseded=$false
    Criteria="IsHidden=0 AND IsInstalled=0"
}) | Out-Null

# D: Managed server, offline
$scenarios.Add(@{
    Id='D'; Description='ssManagedServer + Online=false'
    ServerSelection=1; ServiceID=$null; Online=$false; IncludeSuperseded=$false
    Criteria="IsHidden=0 AND IsInstalled=0"
}) | Out-Null

# E: Managed server, online
$scenarios.Add(@{
    Id='E'; Description='ssManagedServer + Online=true'
    ServerSelection=1; ServiceID=$null; Online=$true; IncludeSuperseded=$false
    Criteria="IsHidden=0 AND IsInstalled=0"
}) | Out-Null

# F: ssOthers + discovered SCCM/WSUS service ID, offline
if ($sccmService) {
    $scenarios.Add(@{
        Id='F'; Description=("ssOthers + ServiceID={0} ({1}) + Online=false" -f $sccmService.ServiceID, $sccmService.Name)
        ServerSelection=3; ServiceID=$sccmService.ServiceID; Online=$false; IncludeSuperseded=$false
        Criteria="IsHidden=0 AND IsInstalled=0"
    }) | Out-Null

    # G: same ServiceID online
    $scenarios.Add(@{
        Id='G'; Description=("ssOthers + ServiceID={0} + Online=true" -f $sccmService.ServiceID)
        ServerSelection=3; ServiceID=$sccmService.ServiceID; Online=$true; IncludeSuperseded=$false
        Criteria="IsHidden=0 AND IsInstalled=0"
    }) | Out-Null

    # H: SCCM-style criteria against managed service offline
    $scenarios.Add(@{
        Id='H'; Description=("SCCM-style criteria + ServiceID={0} + Online=false" -f $sccmService.ServiceID)
        ServerSelection=3; ServiceID=$sccmService.ServiceID; Online=$false; IncludeSuperseded=$false
        Criteria="(DeploymentAction=* AND Type='Software') OR (DeploymentAction=* AND Type='Driver')"
    }) | Out-Null
}

# I: Installed-only baseline (cheap; sanity check WUA itself works)
$scenarios.Add(@{
    Id='I'; Description='ssDefault + Online=false + IsInstalled=1 (installed inventory only)'
    ServerSelection=0; ServiceID=$null; Online=$false; IncludeSuperseded=$false
    Criteria="IsInstalled=1"
}) | Out-Null

# Filter
if ($OnlyScenarios -and $OnlyScenarios.Count -gt 0) {
    $allow = $OnlyScenarios -split ',' | ForEach-Object { $_.Trim() }
    $scenarios = $scenarios | Where-Object { $allow -contains $_.Id }
}

# Run them
foreach ($s in $scenarios) {
    Write-Sub ("Scenario {0}: {1}" -f $s.Id, $s.Description)
    Write-Host ("  ServerSelection={0}  ServiceID={1}  Online={2}  IncludeSuperseded={3}" -f $s.ServerSelection, $s.ServiceID, $s.Online, $s.IncludeSuperseded)
    Write-Host ("  Criteria: {0}" -f $s.Criteria)
    Write-Host ("  Started : {0}" -f (Get-Date -Format o))

    $job = Start-Job -ScriptBlock $scenarioBlock -ArgumentList @(
        $s.Id, $s.Description, [int]$s.ServerSelection, $s.ServiceID,
        [bool]$s.Online, [bool]$s.IncludeSuperseded, $s.Criteria
    )

    $finished = Wait-Job -Job $job -Timeout $ScenarioTimeoutSeconds
    if (-not $finished) {
        Write-Host ("  TIMEOUT after {0}s - stopping job." -f $ScenarioTimeoutSeconds)
        Stop-Job  -Job $job -ErrorAction SilentlyContinue
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        $row = [pscustomobject]@{
            Id           = $s.Id
            Description  = $s.Description
            ServerSel    = $s.ServerSelection
            ServiceID    = $s.ServiceID
            Online       = $s.Online
            IncludeSup   = $s.IncludeSuperseded
            Criteria     = $s.Criteria
            DurationMs   = ($ScenarioTimeoutSeconds * 1000)
            ResultCode   = $null
            UpdateCount  = $null
            FirstTitles  = $null
            Status       = 'TIMEOUT'
            HResult      = $null
            ErrorMessage = "Exceeded $ScenarioTimeoutSeconds s"
        }
    } else {
        $row = Receive-Job -Job $job
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }

    Write-Host ("  Status  : {0}" -f $row.Status)
    Write-Host ("  Duration: {0} ms" -f $row.DurationMs)
    if ($null -ne $row.UpdateCount) { Write-Host ("  Updates : {0}" -f $row.UpdateCount) }
    if ($row.HResult)               { Write-Host ("  HResult : {0}" -f (Format-HResult $row.HResult)) }
    if ($row.ErrorMessage)          { Write-Host ("  Error   : {0}" -f $row.ErrorMessage) }
    if ($row.FirstTitles)           { Write-Host ("  Sample  : {0}" -f $row.FirstTitles) }

    $results.Add($row) | Out-Null
}

} # end if (-not $SkipScenarios)

# ---------------------------------------------------------------------------
# 8. Summary
# ---------------------------------------------------------------------------
Write-Section "8. Summary"

# Headline: Phase 1 vs Phase 2
if ($agentRepro) {
    Write-Host ""
    Write-Host "PHASE 1 - Production reproduction (Microsoft Update direct):"
    Write-Host ("  Status                : {0}" -f $agentRepro.Status)
    Write-Host ("  Total wall-clock time : {0:N0} ms  ({1:N1} s)" -f $agentRepro.TotalMs, ($agentRepro.TotalMs/1000))
    if ($agentRepro.Installed) {
        $i = $agentRepro.Installed
        Write-Host ("    GetInstalledUpdates : Search {0,7:N0} ms  +  Enumerate {1,7:N0} ms  ({2} updates)" -f $i.SearchMs, $i.EnumPassMs, $i.UpdateCount)
    }
    if ($agentRepro.Missing) {
        $m = $agentRepro.Missing
        Write-Host ("    CheckForMissing     : Search {0,7:N0} ms  +  Enumerate {1,7:N0} ms  ({2} updates)" -f $m.SearchMs, $m.EnumPassMs, $m.UpdateCount)
    }
    Write-Host ""
}
if ($agentReproManaged) {
    Write-Host "PHASE 2 - Same flow via internal/managed source:"
    Write-Host ("  Mode                  : {0}" -f $phase2Mode)
    Write-Host ("  Status                : {0}" -f $agentReproManaged.Status)
    Write-Host ("  Total wall-clock time : {0:N0} ms  ({1:N1} s)" -f $agentReproManaged.TotalMs, ($agentReproManaged.TotalMs/1000))
    if ($agentReproManaged.Installed) {
        $i = $agentReproManaged.Installed
        Write-Host ("    GetInstalledUpdates : Search {0,7:N0} ms  +  Enumerate {1,7:N0} ms  ({2} updates)" -f $i.SearchMs, $i.EnumPassMs, $i.UpdateCount)
    }
    if ($agentReproManaged.Missing) {
        $m = $agentReproManaged.Missing
        Write-Host ("    CheckForMissing     : Search {0,7:N0} ms  +  Enumerate {1,7:N0} ms  ({2} updates)" -f $m.SearchMs, $m.EnumPassMs, $m.UpdateCount)
    }
    Write-Host ""
}

Write-Host "Alternative scenarios:"

$results |
    Select-Object Id, Status,
        @{n='DurationS';e={[int]([double]$_.DurationMs / 1000)}},
        @{n='Updates';e={$_.UpdateCount}},
        @{n='HResult';e={Format-HResult $_.HResult}},
        Description |
    Format-Table -AutoSize | Out-String | Write-Host

try {
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host ("CSV summary written to: {0}" -f $csvPath)
} catch {
    Write-Host ("Failed to write CSV: {0}" -f $_.Exception.Message)
}

if ($agentRepro -or $agentReproManaged) {
    $reproCsvPath = $csvPath -replace '_summary\.csv$','_agentrepro.csv'
    try {
        $rows = @()
        if ($agentRepro)        { $rows += ($agentRepro        | Add-Member -NotePropertyName Phase -NotePropertyValue 'Phase1-Production' -PassThru) }
        if ($agentReproManaged) { $rows += ($agentReproManaged | Add-Member -NotePropertyName Phase -NotePropertyValue 'Phase2-Managed'    -PassThru) }
        $rows | Export-Csv -Path $reproCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host ("Agent repro CSV written to: {0}" -f $reproCsvPath)
    } catch {
        Write-Host ("Failed to write agent-repro CSV: {0}" -f $_.Exception.Message)
    }
}

Write-Host ""
Write-Host ("Finished : {0}" -f (Get-Date -Format o))
Write-Host ("Send to support: {0}" -f $logPath)
Write-Host ("                 {0}" -f $csvPath)
Write-Host ""
Write-Host "Also useful: %windir%\Logs\WindowsUpdate\* (run Get-WindowsUpdateLog to merge ETL traces)"

Stop-Transcript | Out-Null
