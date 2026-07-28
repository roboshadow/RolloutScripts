#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [switch]$SkipGpUpdate
)

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'

function Note([string]$Message, [string]$Colour = 'Gray') {
    Write-Host $Message -ForegroundColor $Colour
}

$backupDir = Join-Path 'C:\ProgramData\RoboShadow\CisRevertBackup' (Get-Date -Format 'yyyyMMdd-HHmmss')
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
Note "Backup folder: $backupDir" 'Cyan'

$policyValues = @'
SOFTWARE\Microsoft\Cryptography\Wintrust\Config|EnableCertPaddingCheck
SOFTWARE\Microsoft\Policies\PassportForWork\Biometrics|EnableESSwithSupportedPeripherals
SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon|AutoAdminLogon
SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon|ScreenSaverGracePeriod
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\CredUI|EnumerateAdministrators
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer|NoAutorun
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer|NoDriveTypeAutoRun
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer|NoWebServices
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer|PreXPSP2ShellProtocolBehavior
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS|ADPasswordEncryptionEnabled
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS|BackupDirectory
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS|PasswordAgeDays
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS|PasswordComplexity
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS|PasswordExpirationProtectionEnabled
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS|PasswordLength
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS|PostAuthenticationActions
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS|PostAuthenticationResetDelay
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit|ProcessCreationIncludeCmdLine_Enabled
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\CredSSP\Parameters|AllowEncryptionOracle
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System|DisableAutomaticRestartSignOn
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System|DisableBkGndGroupPolicy
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System|EnableMPR
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System|LocalAccountTokenFilterPolicy
SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System|MSAOptional
SOFTWARE\Policies\Microsoft\AppHVSI|AllowCameraMicrophoneRedirection
SOFTWARE\Policies\Microsoft\AppHVSI|AllowPersistence
SOFTWARE\Policies\Microsoft\AppHVSI|AllowWindowsDefenderApplicationGuard
SOFTWARE\Policies\Microsoft\AppHVSI|AppHVSIClipboardSettings
SOFTWARE\Policies\Microsoft\AppHVSI|AuditApplicationGuard
SOFTWARE\Policies\Microsoft\AppHVSI|SaveFilesToHost
SOFTWARE\Policies\Microsoft\Biometrics\FacialFeatures|EnhancedAntiSpoofing
SOFTWARE\Policies\Microsoft\Dsh|AllowNewsAndInterests
SOFTWARE\Policies\Microsoft\InputPersonalization|AllowInputPersonalization
SOFTWARE\Policies\Microsoft\Internet Explorer\Feeds|AllowBasicAuthInClear
SOFTWARE\Policies\Microsoft\Internet Explorer\Feeds|DisableEnclosureDownload
SOFTWARE\Policies\Microsoft\MicrosoftAccount|DisableUserAuth
SOFTWARE\Policies\Microsoft\Netlogon\Parameters|BlockNetbiosDiscovery
SOFTWARE\Policies\Microsoft\Power\PowerSettings\0e796bdb-100d-47d6-a2d5-f7d2daa51f51|ACSettingIndex
SOFTWARE\Policies\Microsoft\Power\PowerSettings\0e796bdb-100d-47d6-a2d5-f7d2daa51f51|DCSettingIndex
SOFTWARE\Policies\Microsoft\Power\PowerSettings\f15576e8-98b7-4186-b944-eafa664402d9|ACSettingIndex
SOFTWARE\Policies\Microsoft\Power\PowerSettings\f15576e8-98b7-4186-b944-eafa664402d9|DCSettingIndex
SOFTWARE\Policies\Microsoft\W32Time\TimeProviders\NtpClient|Enabled
SOFTWARE\Policies\Microsoft\W32Time\TimeProviders\NtpServer|Enabled
SOFTWARE\Policies\Microsoft\Windows Defender Security Center\App and Browser protection|DisallowExploitProtectionOverride
SOFTWARE\Policies\Microsoft\Windows Defender\Features|PassiveRemediation
SOFTWARE\Policies\Microsoft\Windows Defender\MpEngine|EnableFileHashComputation
SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection|DisableBehaviorMonitoring
SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection|DisableIOAVProtection
SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection|DisableRealtimeMonitoring
SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection|DisableScriptScanning
SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection|OobeEnableRtpAndSigUpdate
SOFTWARE\Policies\Microsoft\Windows Defender\Remediation\Behavioral Network Blocks\Brute Force Protection|BruteForceProtectionConfiguredState
SOFTWARE\Policies\Microsoft\Windows Defender\Scan|DaysUntilAggressiveCatchupQuickScan
SOFTWARE\Policies\Microsoft\Windows Defender\Scan|DisableEmailScanning
SOFTWARE\Policies\Microsoft\Windows Defender\Scan|DisablePackedExeScanning
SOFTWARE\Policies\Microsoft\Windows Defender\Scan|DisableRemovableDriveScanning
SOFTWARE\Policies\Microsoft\Windows Defender\Scan|QuickScanIncludeExclusions
SOFTWARE\Policies\Microsoft\Windows Defender\Spynet|LocalSettingOverrideSpynetReporting
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|26190899-1602-49e8-8b27-eb1d0a1ce869
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|3b576869-a4ec-4529-8536-b80a7769e899
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|56a863a9-875e-4185-98a7-b882c64b5ce5
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|5beb7efe-fd9a-4556-801d-275e5ffc04cc
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|75668c1f-73b5-4cf0-bb93-3ecf5cb7cc84
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|be9ba2d9-53ea-4cdc-84e5-9b1eeee46550
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|d3e037e1-3eb8-44c8-a917-57927947596d
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|d4f940ab-401b-4efc-aadc-ad5f3c50688a
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules|e6db77e5-3df2-4cf1-b95a-636979351e5b
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR|ExploitGuard_ASR_Rules
SOFTWARE\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\Network Protection|EnableNetworkProtection
SOFTWARE\Policies\Microsoft\Windows Defender|HideExclusionsFromLocalUsers
SOFTWARE\Policies\Microsoft\Windows Defender|PUAProtection
SOFTWARE\Policies\Microsoft\Windows NT\DNSClient|EnableMDNS
SOFTWARE\Policies\Microsoft\Windows NT\DNSClient|EnableMulticast
SOFTWARE\Policies\Microsoft\Windows NT\DNSClient|EnableNetbios
SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint|NoWarningNoElevationOnInstall
SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint|RestrictDriverInstallationToAdministrators
SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint|Restricted
SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint|UpdatePromptSettings
SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC|ForceKerberosForRpc
SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC|RpcAuthentication
SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC|RpcProtocols
SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC|RpcTcpPort
SOFTWARE\Policies\Microsoft\Windows NT\Printers\RPC|RpcUseNamedPipeProtocol
SOFTWARE\Policies\Microsoft\Windows NT\Printers|CopyFilesPolicy
SOFTWARE\Policies\Microsoft\Windows NT\Printers|DisableWebPnPDownload
SOFTWARE\Policies\Microsoft\Windows NT\Printers|RedirectionguardPolicy
SOFTWARE\Policies\Microsoft\Windows NT\Printers|RegisterSpoolerRemoteRpcEndPoint
SOFTWARE\Policies\Microsoft\Windows NT\Rpc|EnableAuthEpResolution
SOFTWARE\Policies\Microsoft\Windows NT\Rpc|RestrictRemoteClients
SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services|DeleteTempDirsOnExit
SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services|DisablePasswordSaving
SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services|fAllowToGetHelp
SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services|fAllowUnsolicited
SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services|fDisableCdm
SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services|fEncryptRPCTraffic
SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services|fPromptForPassword
SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services|MinEncryptionLevel
SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services|SecurityLayer
SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services|UserAuthentication
SOFTWARE\Policies\Microsoft\Windows\AppInstaller|EnableBypassCertificatePinningForMicrosoftStore
SOFTWARE\Policies\Microsoft\Windows\AppInstaller|EnableExperimentalFeatures
SOFTWARE\Policies\Microsoft\Windows\AppInstaller|EnableHashOverride
SOFTWARE\Policies\Microsoft\Windows\AppInstaller|EnableLocalArchiveMalwareScanOverride
SOFTWARE\Policies\Microsoft\Windows\AppInstaller|EnableMSAppInstallerProtocol
SOFTWARE\Policies\Microsoft\Windows\AppPrivacy|LetAppsActivateWithVoiceAboveLock
SOFTWARE\Policies\Microsoft\Windows\Appx|BlockNonAdminUserInstall
SOFTWARE\Policies\Microsoft\Windows\Appx|DisablePerUserUnsignedPackagesByDefault
SOFTWARE\Policies\Microsoft\Windows\Bowser|EnableMailslots
SOFTWARE\Policies\Microsoft\Windows\CloudContent|DisableConsumerAccountStateContent
SOFTWARE\Policies\Microsoft\Windows\CloudContent|DisableWindowsConsumerFeatures
SOFTWARE\Policies\Microsoft\Windows\Connect|RequirePinForPairing
SOFTWARE\Policies\Microsoft\Windows\CredentialsDelegation|AllowProtectedCreds
SOFTWARE\Policies\Microsoft\Windows\CredUI|DisablePasswordReveal
SOFTWARE\Policies\Microsoft\Windows\DataCollection|AllowTelemetry
SOFTWARE\Policies\Microsoft\Windows\DataCollection|DisableOneSettingsDownloads
SOFTWARE\Policies\Microsoft\Windows\DataCollection|DoNotShowFeedbackNotifications
SOFTWARE\Policies\Microsoft\Windows\DataCollection|EnableOneSettingsAuditing
SOFTWARE\Policies\Microsoft\Windows\DataCollection|LimitDiagnosticLogCollection
SOFTWARE\Policies\Microsoft\Windows\DataCollection|LimitDumpCollection
SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization|DODownloadMode
SOFTWARE\Policies\Microsoft\Windows\Device Metadata|PreventDeviceMetadataFromNetwork
SOFTWARE\Policies\Microsoft\Windows\DeviceGuard|ConfigureKernelShadowStacksLaunch
SOFTWARE\Policies\Microsoft\Windows\DeviceGuard|ConfigureSystemGuardLaunch
SOFTWARE\Policies\Microsoft\Windows\DeviceGuard|EnableVirtualizationBasedSecurity
SOFTWARE\Policies\Microsoft\Windows\DeviceGuard|HVCIMATRequired
SOFTWARE\Policies\Microsoft\Windows\DeviceGuard|HypervisorEnforcedCodeIntegrity
SOFTWARE\Policies\Microsoft\Windows\DeviceGuard|LsaCfgFlags
SOFTWARE\Policies\Microsoft\Windows\DeviceGuard|RequirePlatformSecurityFeatures
SOFTWARE\Policies\Microsoft\Windows\EventLog\Application|MaxSize
SOFTWARE\Policies\Microsoft\Windows\EventLog\Application|Retention
SOFTWARE\Policies\Microsoft\Windows\EventLog\Security|MaxSize
SOFTWARE\Policies\Microsoft\Windows\EventLog\Security|Retention
SOFTWARE\Policies\Microsoft\Windows\EventLog\Setup|MaxSize
SOFTWARE\Policies\Microsoft\Windows\EventLog\Setup|Retention
SOFTWARE\Policies\Microsoft\Windows\EventLog\System|MaxSize
SOFTWARE\Policies\Microsoft\Windows\EventLog\System|Retention
SOFTWARE\Policies\Microsoft\Windows\Explorer|DisableMotWOnInsecurePathCopy
SOFTWARE\Policies\Microsoft\Windows\Explorer|NoAutoplayfornonVolume
SOFTWARE\Policies\Microsoft\Windows\Explorer|NoDataExecutionPrevention
SOFTWARE\Policies\Microsoft\Windows\Explorer|NoHeapTerminationOnCorruption
SOFTWARE\Policies\Microsoft\Windows\GameDVR|AllowGameDVR
SOFTWARE\Policies\Microsoft\Windows\Group Policy\{35378EAC-683F-11D2-A89A-00C04FBBCFA2}|NoBackgroundPolicy
SOFTWARE\Policies\Microsoft\Windows\Group Policy\{35378EAC-683F-11D2-A89A-00C04FBBCFA2}|NoGPOListChanges
SOFTWARE\Policies\Microsoft\Windows\Group Policy\{827D319E-6EAC-11D2-A4EA-00C04F79F83A}|NoBackgroundPolicy
SOFTWARE\Policies\Microsoft\Windows\Group Policy\{827D319E-6EAC-11D2-A4EA-00C04F79F83A}|NoGPOListChanges
SOFTWARE\Policies\Microsoft\Windows\Installer|AlwaysInstallElevated
SOFTWARE\Policies\Microsoft\Windows\Installer|EnableUserControl
SOFTWARE\Policies\Microsoft\Windows\LanmanServer|AuditClientDoesNotSupportEncryption
SOFTWARE\Policies\Microsoft\Windows\LanmanServer|AuditClientDoesNotSupportSigning
SOFTWARE\Policies\Microsoft\Windows\LanmanServer|AuditInsecureGuestLogon
SOFTWARE\Policies\Microsoft\Windows\LanmanServer|EnableAuthRateLimiter
SOFTWARE\Policies\Microsoft\Windows\LanmanServer|InvalidAuthenticationDelayTimeInMs
SOFTWARE\Policies\Microsoft\Windows\LanmanServer|MinSmb2Dialect
SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation|AllowInsecureGuestAuth
SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation|AuditInsecureGuestLogon
SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation|AuditServerDoesNotSupportEncryption
SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation|AuditServerDoesNotSupportSigning
SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation|MinSmb2Dialect
SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation|RequireEncryption
SOFTWARE\Policies\Microsoft\Windows\Network Connections|NC_AllowNetBridge_NLA
SOFTWARE\Policies\Microsoft\Windows\Network Connections|NC_ShowSharedAccessUI
SOFTWARE\Policies\Microsoft\Windows\Network Connections|NC_StdDomainUserSetLocation
SOFTWARE\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths|\\*\NETLOGON
SOFTWARE\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths|\\*\SYSVOL
Software\Policies\Microsoft\Windows\NetworkProvider|EnableMailslots
SOFTWARE\Policies\Microsoft\Windows\OneDrive|DisableFileSyncNGSC
SOFTWARE\Policies\Microsoft\Windows\Personalization|NoLockScreenCamera
SOFTWARE\Policies\Microsoft\Windows\Personalization|NoLockScreenSlideshow
SOFTWARE\Policies\Microsoft\Windows\Sandbox|AllowClipboardRedirection
SOFTWARE\Policies\Microsoft\Windows\Sandbox|AllowNetworking
SOFTWARE\Policies\Microsoft\Windows\System|AllowCustomSSPsAPs
SOFTWARE\Policies\Microsoft\Windows\System|AllowDomainPINLogon
SOFTWARE\Policies\Microsoft\Windows\System|BlockDomainPicturePassword
SOFTWARE\Policies\Microsoft\Windows\System|BlockUserFromShowingAccountDetailsOnSignin
SOFTWARE\Policies\Microsoft\Windows\System|DisableLockScreenAppNotifications
SOFTWARE\Policies\Microsoft\Windows\System|DontDisplayNetworkSelectionUI
SOFTWARE\Policies\Microsoft\Windows\System|DontEnumerateConnectedUsers
SOFTWARE\Policies\Microsoft\Windows\System|EnableCdp
SOFTWARE\Policies\Microsoft\Windows\System|EnableSmartScreen
SOFTWARE\Policies\Microsoft\Windows\System|EnumerateLocalUsers
SOFTWARE\Policies\Microsoft\Windows\System|NoLocalPasswordResetQuestions
SOFTWARE\Policies\Microsoft\Windows\System|RunAsPPL
SOFTWARE\Policies\Microsoft\Windows\System|ShellSmartScreenLevel
SOFTWARE\Policies\Microsoft\Windows\WcmSvc\GroupPolicy|fBlockNonDomain
SOFTWARE\Policies\Microsoft\Windows\WcmSvc\GroupPolicy|fMinimizeConnections
SOFTWARE\Policies\Microsoft\Windows\Windows Search|AllowCortana
SOFTWARE\Policies\Microsoft\Windows\Windows Search|AllowCortanaAboveLock
SOFTWARE\Policies\Microsoft\Windows\Windows Search|AllowIndexingEncryptedStoresOrItems
SOFTWARE\Policies\Microsoft\Windows\Windows Search|AllowSearchToUseLocation
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU|NoAutoRebootWithLoggedOnUsers
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU|NoAutoUpdate
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU|ScheduledInstallDay
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate|AllowTemporaryEnterpriseFeatureControl
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate|DeferFeatureUpdates
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate|DeferFeatureUpdatesPeriodInDays
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate|DeferQualityUpdates
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate|DeferQualityUpdatesPeriodInDays
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate|ManagePreviewBuildsPolicyValue
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate|SetAllowOptionalContent
SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate|SetDisablePauseUXAccess
SOFTWARE\Policies\Microsoft\Windows\WinRM\Client|AllowBasic
SOFTWARE\Policies\Microsoft\Windows\WinRM\Client|AllowDigest
SOFTWARE\Policies\Microsoft\Windows\WinRM\Client|AllowUnencryptedTraffic
SOFTWARE\Policies\Microsoft\Windows\WinRM\Service|AllowBasic
SOFTWARE\Policies\Microsoft\Windows\WinRM\Service|AllowUnencryptedTraffic
SOFTWARE\Policies\Microsoft\Windows\WinRM\Service|DisableRunAs
SOFTWARE\Policies\Microsoft\Windows\WTDS\Components|CaptureThreatWindow
SOFTWARE\Policies\Microsoft\Windows\WTDS\Components|NotifyMalicious
SOFTWARE\Policies\Microsoft\Windows\WTDS\Components|NotifyPasswordReuse
SOFTWARE\Policies\Microsoft\Windows\WTDS\Components|NotifyUnsafeApp
SOFTWARE\Policies\Microsoft\Windows\WTDS\Components|ServiceEnabled
SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\Logging|LogDroppedPackets
SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\Logging|LogFilePath
SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\Logging|LogFileSize
SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\Logging|LogSuccessfulConnections
SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile|DefaultInboundAction
SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile|DisableNotifications
SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile|EnableFirewall
SOFTWARE\Policies\Microsoft\WindowsFirewall\PrivateProfile\Logging|LogDroppedPackets
SOFTWARE\Policies\Microsoft\WindowsFirewall\PrivateProfile\Logging|LogFilePath
SOFTWARE\Policies\Microsoft\WindowsFirewall\PrivateProfile\Logging|LogFileSize
SOFTWARE\Policies\Microsoft\WindowsFirewall\PrivateProfile\Logging|LogSuccessfulConnections
SOFTWARE\Policies\Microsoft\WindowsFirewall\PrivateProfile|DefaultInboundAction
SOFTWARE\Policies\Microsoft\WindowsFirewall\PrivateProfile|DisableNotifications
SOFTWARE\Policies\Microsoft\WindowsFirewall\PrivateProfile|EnableFirewall
SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile\Logging|LogDroppedPackets
SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile\Logging|LogFilePath
SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile\Logging|LogFileSize
SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile\Logging|LogSuccessfulConnections
SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile|AllowLocalIPsecPolicyMerge
SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile|AllowLocalPolicyMerge
SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile|DefaultInboundAction
SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile|DisableNotifications
SOFTWARE\Policies\Microsoft\WindowsFirewall\PublicProfile|EnableFirewall
SOFTWARE\Policies\Microsoft\WindowsInkWorkspace|AllowWindowsInkWorkspace
SOFTWARE\Policies\Microsoft\WindowsStore|AutoDownload
SOFTWARE\Policies\Microsoft\WindowsStore|DisableOSUpgrade
SOFTWARE\Wow6432Node\Microsoft\Cryptography\Wintrust\Config|EnableCertPaddingCheck
SYSTEM\CurrentControlSet\Control\Lsa|RunAsPPL
SYSTEM\CurrentControlSet\Control\Print|RpcAuthnLevelPrivacyEnabled
SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest|UseLogonCredential
SYSTEM\CurrentControlSet\Control\Session Manager\kernel|DisableExceptionChainValidation
SYSTEM\CurrentControlSet\Control\Session Manager|SafeDllSearchMode
SYSTEM\CurrentControlSet\Policies\EarlyLaunch|DriverLoadPolicy
SYSTEM\CurrentControlSet\Services\Eventlog\Security|WarningLevel
SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters|SMB1
SYSTEM\CurrentControlSet\Services\mrxsmb10|Start
SYSTEM\CurrentControlSet\Services\NetBT\Parameters|NodeType
SYSTEM\CurrentControlSet\Services\NetBT\Parameters|NoNameReleaseOnDemand
SYSTEM\CurrentControlSet\Services\Tcpip\Parameters|DisableIPSourceRouting
SYSTEM\CurrentControlSet\Services\Tcpip\Parameters|EnableICMPRedirect
SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters|DisableIPSourceRouting
'@ -split "`r?`n" | Where-Object { $_ }

$deleteSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($pair in $policyValues) { [void]$deleteSet.Add($pair) }

Note ''
Note 'Step 1/6: Locating LGPO.exe' 'Cyan'
$polPath = Join-Path $env:WINDIR 'System32\GroupPolicy\Machine\Registry.pol'
$lgpo = $null
if (Test-Path $polPath) {
    $healLgpo = 'C:\ProgramData\RoboShadow\Rubicon\Control\Download\LGPO.exe'
    if ((Test-Path $healLgpo) -and (Get-Item $healLgpo).Length -gt 100000) {
        $lgpo = $healLgpo
        Note 'Using LGPO.exe already downloaded by the RoboShadow fix jobs' 'Green'
    }
    else {
        $lgpo = Join-Path $backupDir 'LGPO.exe'
        $lgpoUrl = 'https://cdn.roboshadow.com/cloudscanner/LGPO.exe'
        try {
            $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
            $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
            $handler = New-Object System.Net.Http.HttpClientHandler
            $handler.Proxy = $proxy
            $handler.UseProxy = $true
            $client = New-Object System.Net.Http.HttpClient($handler)
            [System.IO.File]::WriteAllBytes($lgpo, $client.GetByteArrayAsync($lgpoUrl).Result)
        }
        catch {
            Invoke-WebRequest -Uri $lgpoUrl -OutFile $lgpo -UseBasicParsing
        }
        if (!(Test-Path $lgpo) -or (Get-Item $lgpo).Length -lt 100000) {
            throw 'Could not download LGPO.exe - check internet access and run the script again. Nothing has been changed.'
        }
        Note 'LGPO.exe downloaded from the RoboShadow CDN' 'Green'
    }
}
else {
    Note 'No Machine Registry.pol present - LGPO.exe not needed' 'DarkGray'
}

Note ''
Note 'Step 2/6: Removing CIS fix entries from local group policy (keeping everything else)' 'Cyan'
if (Test-Path $polPath) {
    Copy-Item $polPath (Join-Path $backupDir 'Machine-Registry.pol.bak') -Force
    $parseTxt = Join-Path $backupDir 'Registry.pol.parsed.txt'
    cmd /c "`"$lgpo`" /parse /m `"$polPath`" > `"$parseTxt`" 2>nul" | Out-Null
    if ($LASTEXITCODE -ne 0 -or !(Test-Path $parseTxt)) {
        throw "LGPO.exe could not parse $polPath (exit code $LASTEXITCODE). Nothing has been changed."
    }

    $repairedCount = 0
    $lines = Get-Content $parseTxt | ForEach-Object {
        if ($_ -cmatch '^OFTWARE\\') { $repairedCount++; 'S' + $_ } else { $_ }
    }
    if ($repairedCount -gt 0) {
        Note "Repaired $repairedCount corrupt policy entries (missing leading character)" 'Yellow'
    }

    $keptBlocks = New-Object System.Collections.Generic.List[string]
    $droppedCount = 0
    $current = New-Object System.Collections.Generic.List[string]
    $allBlocks = New-Object System.Collections.Generic.List[object]
    foreach ($line in @($lines) + '') {
        if ($line.Trim() -eq '') {
            if ($current.Count -gt 0) { $allBlocks.Add($current.ToArray()); $current.Clear() }
        }
        elseif (!$line.StartsWith(';')) {
            $current.Add($line)
        }
    }
    foreach ($block in $allBlocks) {
        if ($block.Count -ge 3 -and $deleteSet.Contains("$($block[1])|$($block[2])")) {
            $droppedCount++
        }
        else {
            $keptBlocks.Add(($block -join "`r`n"))
        }
    }
    Note "Removing $droppedCount CIS fix entries, keeping $($keptBlocks.Count) other policy entries" 'Green'

    Remove-Item $polPath -Force
    if ($keptBlocks.Count -gt 0) {
        $rebuildTxt = Join-Path $backupDir 'Registry.pol.rebuild.txt'
        [System.IO.File]::WriteAllText($rebuildTxt, (($keptBlocks -join "`r`n`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
        cmd /c "`"$lgpo`" /t `"$rebuildTxt`" 2>&1" | Out-Null
        if ($LASTEXITCODE -ne 0 -or !(Test-Path $polPath)) {
            Copy-Item (Join-Path $backupDir 'Machine-Registry.pol.bak') $polPath -Force
            throw "LGPO.exe failed to rebuild the policy file (exit code $LASTEXITCODE). The original Registry.pol has been restored."
        }
        Note 'Registry.pol rebuilt without the CIS fix entries' 'Green'
    }
    else {
        Note 'No other policy entries existed - Registry.pol not recreated' 'Green'
    }
}
else {
    Note 'No Machine Registry.pol present - nothing to clean' 'DarkGray'
}

$gptIni = Join-Path $env:WINDIR 'System32\GroupPolicy\gpt.ini'
if (Test-Path $gptIni) {
    Copy-Item $gptIni (Join-Path $backupDir 'gpt.ini.bak') -Force
    $newIni = Get-Content $gptIni | ForEach-Object {
        if ($_ -match '^Version=(\d+)$') { 'Version=' + ([int]$Matches[1] + 1) } else { $_ }
    }
    Set-Content -Path $gptIni -Value $newIni -Encoding Ascii
}

Note ''
Note 'Step 3/6: Removing policy registry values applied by the CIS fixes' 'Cyan'
$removed = 0
$backupReg = New-Object System.Collections.Generic.List[string]
foreach ($pair in $policyValues) {
    $key, $valueName = $pair -split '\|', 2
    if ($key -eq 'SYSTEM\CurrentControlSet\Services\mrxsmb10' -and $valueName -eq 'Start') { continue }
    try {
        $reg = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($key, $true)
        if ($null -ne $reg) {
            if ($reg.GetValueNames() -contains $valueName) {
                $existing = $reg.GetValue($valueName, $null, 'DoNotExpandEnvironmentNames')
                $backupReg.Add("HKLM\$key|$valueName|$($existing -join ';')")
                $reg.DeleteValue($valueName, $false)
                $removed++
            }
            $reg.Close()
        }
    }
    catch {
        Note "Could not remove HKLM\$key -> $valueName ($($_.Exception.Message))" 'Yellow'
    }
}
$backupReg | Set-Content (Join-Path $backupDir 'removed-values.txt') -Encoding UTF8
Note "Removed $removed policy registry values (previous values saved to removed-values.txt)" 'Green'

$mrx = 'HKLM:\SYSTEM\CurrentControlSet\Services\mrxsmb10'
if (Test-Path $mrx) {
    Set-ItemProperty -Path $mrx -Name 'Start' -Value 3 -Type DWord
    Note 'Restored SMB1 client driver (mrxsmb10) start type to Manual' 'Green'
}

Note ''
Note 'Step 4/6: Restoring security options changed directly in the registry' 'Cyan'
function Remove-RegValue([string]$Path, [string]$Name) {
    if (Test-Path $Path) {
        Remove-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    }
}
function Set-RegValue([string]$Path, [string]$Name, $Value, [string]$Type) {
    if (!(Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type
}

$sysPolicies = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
Remove-RegValue $sysPolicies 'DisableCAD'
Remove-RegValue $sysPolicies 'InactivityTimeoutSecs'
Set-RegValue $sysPolicies 'DontDisplayLastUserName' 0 'DWord'
Set-RegValue $sysPolicies 'ConsentPromptBehaviorUser' 3 'DWord'
Set-RegValue $sysPolicies 'LegalNoticeCaption' '' 'String'
Set-RegValue $sysPolicies 'LegalNoticeText' '' 'String'
Remove-RegValue "$sysPolicies\Kerberos\Parameters" 'SupportedEncryptionTypes'

Set-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' 'ScRemoveOption' '0' 'String'

$lsa = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'
Remove-RegValue $lsa 'SCENoApplyLegacyAuditPolicy'
Remove-RegValue $lsa 'RestrictAnonymous'
Remove-RegValue $lsa 'DisableDomainCreds'
Remove-RegValue $lsa 'RestrictRemoteSAM'
Remove-RegValue $lsa 'UseMachineId'
Remove-RegValue $lsa 'LmCompatibilityLevel'
Remove-RegValue "$lsa\MSV1_0" 'AllowNullSessionFallback'
Remove-RegValue "$lsa\MSV1_0" 'NTLMMinClientSec'
Remove-RegValue "$lsa\MSV1_0" 'NTLMMinServerSec'
Remove-RegValue "$lsa\MSV1_0" 'AuditReceivingNTLMTraffic'
Remove-RegValue "$lsa\MSV1_0" 'RestrictSendingNTLMTraffic'
Remove-RegValue "$lsa\pku2u" 'AllowOnlineID'

Remove-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\SAM' 'RelaxMinimumPasswordLengthLimits'

Set-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters' 'RequireSecuritySignature' 0 'DWord'
Set-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters' 'RequireSecuritySignature' 0 'DWord'
Set-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters' 'EnableSecuritySignature' 0 'DWord'
Remove-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters' 'SmbServerNameHardeningLevel'

$defaultAllowedPaths = @(
    'System\CurrentControlSet\Control\Print\Printers',
    'System\CurrentControlSet\Services\Eventlog',
    'Software\Microsoft\OLAP Server',
    'Software\Microsoft\Windows NT\CurrentVersion\Print',
    'Software\Microsoft\Windows NT\CurrentVersion\Windows',
    'System\CurrentControlSet\Control\ContentIndex',
    'System\CurrentControlSet\Control\Terminal Server',
    'System\CurrentControlSet\Control\Terminal Server\UserConfig',
    'System\CurrentControlSet\Control\Terminal Server\DefaultUserConfiguration',
    'Software\Microsoft\Windows NT\CurrentVersion\Perflib',
    'System\CurrentControlSet\Services\SysmonLog'
)
Set-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurePipeServers\winreg\AllowedPaths' 'Machine' $defaultAllowedPaths 'MultiString'
Note 'Security options restored to Windows defaults' 'Green'

Note ''
Note 'Step 5/6: Restoring service start types disabled by the fixes' 'Cyan'
$serviceDefaults = @{
    'RpcLocator'          = 3
    'SSDPSRV'             = 3
    'upnphost'            = 3
    'WerSvc'              = 3
    'Wecsvc'              = 3
    'WMPNetworkSvc'       = 3
    'icssvc'              = 3
    'WpnService'          = 2
    'PushToInstall'       = 3
    'WinHttpAutoProxySvc' = 3
    'XboxGipSvc'          = 3
    'XblAuthManager'      = 3
    'XblGameSave'         = 3
    'XboxNetApiSvc'       = 3
}
foreach ($svc in $serviceDefaults.GetEnumerator()) {
    $svcKey = "HKLM:\SYSTEM\CurrentControlSet\Services\$($svc.Key)"
    if (Test-Path $svcKey) {
        Set-ItemProperty -Path $svcKey -Name 'Start' -Value $svc.Value -Type DWord
        Note "  $($svc.Key) -> start type restored"
    }
}
Note 'Service start types restored' 'Green'

Note ''
Note 'Step 6/6: Refreshing group policy' 'Cyan'
if ($SkipGpUpdate) {
    Note 'Skipped (-SkipGpUpdate)' 'DarkGray'
}
else {
    $gp = Start-Process -FilePath (Join-Path $env:WINDIR 'System32\gpupdate.exe') -ArgumentList '/target:computer', '/force' -Wait -PassThru -WindowStyle Hidden
    Note "gpupdate exit code: $($gp.ExitCode)" 'Green'
}

Note ''
Note 'Done. All CIS benchmark fixes applied via LGPO/registry have been reverted.' 'Green'
Note "Backups are in $backupDir if anything needs to be restored." 'Cyan'
Note 'A reboot is recommended so all reverted settings (services, logon, SMB, TLS) fully take effect.' 'Yellow'

