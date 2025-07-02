param(
    [Parameter(Mandatory=$false)]  # Change to false
    [string]$OrganisationId
)

# Prompt for OrganisationId if not provided
if (-not $OrganisationId) {
    $OrganisationId = Read-Host "Please enter your OrganisationId"
}


function Get-RegistryValue {
    param(
        [string]$Path,
        [string]$Name
    )
    try {
        $value = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return $value.$Name
    }
    catch {
        return $null
    }
}

function Test-RegistryKey {
    param([string]$Path)
    try {
        Get-Item -Path $Path -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

Clear-Host
$separator = "=" * 70
Write-Host $separator -ForegroundColor Cyan
Write-Host "    ROBOSHADOW AGENT DIAGNOSTIC REPORT" -ForegroundColor White
Write-Host $separator -ForegroundColor Cyan
Write-Host "Date/Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host "Organisation ID: $OrganisationId" -ForegroundColor Gray
Write-Host ""

# Check 1: Service Exists
Write-Host "1. SERVICE EXISTENCE CHECK" -ForegroundColor Yellow
$service = Get-Service -Name "RoboShadowAgent" -ErrorAction SilentlyContinue
if ($service) {
    Write-Host "   ✓ RoboShadowAgent service EXISTS" -ForegroundColor Green
} else {
    Write-Host "   ✗ RoboShadowAgent service NOT FOUND" -ForegroundColor Red
}
Write-Host ""

# Check 2: Service Running
Write-Host "2. SERVICE STATUS CHECK" -ForegroundColor Yellow
if ($service) {
    if ($service.Status -eq "Running") {
        Write-Host "   ✓ RoboShadowAgent service is RUNNING" -ForegroundColor Green
    } else {
        Write-Host "   ✗ RoboShadowAgent service is $($service.Status)" -ForegroundColor Red
    }
} else {
    Write-Host "   ✗ Cannot check status - service not found" -ForegroundColor Red
}
Write-Host ""

$basePath = "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon"

# Check 3: Agent Version
Write-Host "3. AGENT VERSION REGISTRY CHECK" -ForegroundColor Yellow
$agentVersionPath = "$basePath\Agent"
if (Test-RegistryKey $agentVersionPath) {
    $agentVersion = Get-RegistryValue $agentVersionPath "Version"
    if ($agentVersion) {
        Write-Host "   ✓ Agent Version key EXISTS: $agentVersion" -ForegroundColor Green
    } else {
        Write-Host "   ⚠ Agent path exists but Version value missing" -ForegroundColor DarkYellow
    }
} else {
    Write-Host "   ✗ Agent Version registry key NOT FOUND" -ForegroundColor Red
}
Write-Host ""

# Check 4: Control Version
Write-Host "4. CONTROL VERSION REGISTRY CHECK" -ForegroundColor Yellow
$controlVersionPath = "$basePath\Control"
if (Test-RegistryKey $controlVersionPath) {
    $controlVersion = Get-RegistryValue $controlVersionPath "Version"
    if ($controlVersion) {
        Write-Host "   ✓ Control Version key EXISTS: $controlVersion" -ForegroundColor Green
    } else {
        Write-Host "   ⚠ Control path exists but Version value missing" -ForegroundColor DarkYellow
    }
} else {
    Write-Host "   ✗ Control Version registry key NOT FOUND" -ForegroundColor Red
}
Write-Host ""

# Check 5: Control Organisation ID
Write-Host "5. CONTROL ORGANISATION ID CHECK" -ForegroundColor Yellow
if (Test-RegistryKey $controlVersionPath) {
    $controlOrgId = Get-RegistryValue $controlVersionPath "OrganisationId"
    if ($controlOrgId) {
        if ($controlOrgId -eq $OrganisationId) {
            Write-Host "   ✓ Control OrganisationId MATCHES: $controlOrgId|" -ForegroundColor Green
        } else {
            Write-Host "   ✗ Control OrganisationId MISMATCH" -ForegroundColor Red
            Write-Host "     Expected: $OrganisationId|" -ForegroundColor Red
            Write-Host "     Found: $controlOrgId|" -ForegroundColor Red
        }
    } else {
        Write-Host "   ✗ Control OrganisationId value NOT FOUND" -ForegroundColor Red
    }
} else {
    Write-Host "   ✗ Control registry key not found" -ForegroundColor Red
}
Write-Host ""

# Check 6: Agent Organisation ID
Write-Host "6. AGENT ORGANISATION ID CHECK" -ForegroundColor Yellow
if (Test-RegistryKey $agentVersionPath) {
    $agentOrgId = Get-RegistryValue $agentVersionPath "OrganisationId"
    if ($agentOrgId) {
        if ($agentOrgId -eq $OrganisationId) {
            Write-Host "   ✓ Agent OrganisationId MATCHES: $agentOrgId|" -ForegroundColor Green
        } else {
            Write-Host "   ✗ Agent OrganisationId MISMATCH" -ForegroundColor Red
            Write-Host "     Expected: $OrganisationId|" -ForegroundColor Red
            Write-Host "     Found: $agentOrgId|" -ForegroundColor Red
        }
    } else {
        Write-Host "   ✗ Agent OrganisationId value NOT FOUND" -ForegroundColor Red
    }
} else {
    Write-Host "   ✗ Agent registry key not found" -ForegroundColor Red
}
Write-Host ""

# Check 7: Device ID
Write-Host "7. CONTROL DEVICE ID CHECK" -ForegroundColor Yellow
if (Test-RegistryKey $controlVersionPath) {
    $deviceId = Get-RegistryValue $controlVersionPath "DeviceId"
    if ($deviceId) {
        Write-Host "   ✓ Control DeviceId EXISTS: $deviceId" -ForegroundColor Green
    } else {
        Write-Host "   ✗ Control DeviceId value NOT FOUND" -ForegroundColor Red
    }
} else {
    Write-Host "   ✗ Control registry key not found" -ForegroundColor Red
}
Write-Host ""

# Check 8: Public Key File
Write-Host "8. PUBLIC KEY FILE CHECK" -ForegroundColor Yellow
$publicKeyPath = "C:\ProgramData\RoboShadow\Rubicon\Control\Data\PublicKey"
if (Test-Path $publicKeyPath) {
    $fileInfo = Get-Item $publicKeyPath
    Write-Host "   ✓ PublicKey file EXISTS" -ForegroundColor Green
    Write-Host "     Size: $($fileInfo.Length) bytes" -ForegroundColor Gray
    Write-Host "     Modified: $($fileInfo.LastWriteTime)" -ForegroundColor Gray
} else {
    Write-Host "   ✗ PublicKey file NOT FOUND" -ForegroundColor Red
    Write-Host "     Expected location: $publicKeyPath" -ForegroundColor Red
}

Write-Host ""
Write-Host $separator -ForegroundColor Cyan
Write-Host "    DIAGNOSTIC COMPLETE" -ForegroundColor White
Write-Host $separator -ForegroundColor Cyan
Write-Host "Please screenshot this entire output and send to support." -ForegroundColor White
Write-Host "Computer Name: $env:COMPUTERNAME" -ForegroundColor Gray
Write-Host "User: $env:USERNAME" -ForegroundColor Gray

# Get OS Information
$osInfo = Get-WmiObject -Class Win32_OperatingSystem
Write-Host "Operating System: $($osInfo.Caption)" -ForegroundColor Gray
Write-Host "OS Version: $($osInfo.Version)" -ForegroundColor Gray
Write-Host "OS Build: $($osInfo.BuildNumber)" -ForegroundColor Gray
Write-Host ""