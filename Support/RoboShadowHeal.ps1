param(
    [Parameter(Mandatory=$true)]
    [string]$OrganisationId
)

function Write-LogMessage {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $(
        switch ($Level) {
            "SUCCESS" { "Green" }
            "ERROR" { "Red" }
            "WARNING" { "Yellow" }
            "INFO" { "White" }
            default { "Gray" }
        }
    )
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

function Test-AgentInstalled {
    Write-LogMessage "Checking if RoboShadow Agent is installed..."
    
    try {
        $uninstallKey = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction Stop |
            Get-ItemProperty |
            Where-Object { $_.DisplayName -like "RoboShadow Agent" } |
            Select-Object -ExpandProperty UninstallString
        
        if ($uninstallKey) {
            Write-LogMessage "RoboShadow Agent installation found" "SUCCESS"
            return $true
        } else {
            Write-LogMessage "RoboShadow Agent not found in installed programs" "ERROR"
            return $false
        }
    }
    catch {
        Write-LogMessage "Error checking agent installation: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Test-ServiceExists {
    param([string]$ServiceName)
    
    Write-LogMessage "Checking if service '$ServiceName' exists..."
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    
    if ($service) {
        Write-LogMessage "Service '$ServiceName' exists" "SUCCESS"
        return $true
    } else {
        Write-LogMessage "Service '$ServiceName' not found" "ERROR"
        return $false
    }
}

function Test-ServiceRunning {
    param([string]$ServiceName)
    
    Write-LogMessage "Checking if service '$ServiceName' is running..."
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    
    if ($service -and $service.Status -eq "Running") {
        Write-LogMessage "Service '$ServiceName' is running" "SUCCESS"
        return $true
    } elseif ($service) {
        Write-LogMessage "Service '$ServiceName' exists but is not running (Status: $($service.Status))" "WARNING"
        return $false
    } else {
        Write-LogMessage "Service '$ServiceName' not found" "ERROR"
        return $false
    }
}

function Get-CurrentOrganisationId {
    $basePath = "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon"
    $controlPath = "$basePath\Control"
    
    Write-LogMessage "Retrieving current OrganisationId from registry..."
    
    if (Test-RegistryKey $controlPath) {
        $currentOrgId = Get-RegistryValue $controlPath "OrganisationId"
        if ($currentOrgId) {
            Write-LogMessage "Found OrganisationId: $currentOrgId"
            return $currentOrgId
        } else {
            Write-LogMessage "OrganisationId registry value not found" "WARNING"
            return $null
        }
    } else {
        Write-LogMessage "Control registry path not found" "ERROR"
        return $null
    }
}

function Test-OrganisationIdMatch {
    param(
        [string]$ExpectedOrgId,
        [string]$CurrentOrgId
    )
    
    Write-LogMessage "Validating OrganisationId match..."
    Write-LogMessage "Expected: $ExpectedOrgId"
    Write-LogMessage "Current:  $CurrentOrgId"
    
    if ($CurrentOrgId -eq $ExpectedOrgId) {
        Write-LogMessage "OrganisationId matches" "SUCCESS"
        return $true
    } else {
        Write-LogMessage "OrganisationId mismatch detected" "ERROR"
        return $false
    }
}

function Stop-RoboShadowService {
    param([string]$ServiceName = "RoboShadowAgent")
    
    Write-LogMessage "Stopping service '$ServiceName'..."
    
    try {
        Stop-Service -Name $ServiceName -Force -ErrorAction Stop
        Write-LogMessage "Service '$ServiceName' stopped successfully" "SUCCESS"
        return $true
    }
    catch {
        Write-LogMessage "Failed to stop service '$ServiceName': $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Start-RoboShadowService {
    param([string]$ServiceName = "RoboShadowAgent")
    
    Write-LogMessage "Starting service '$ServiceName'..." "INFO"
    
    try {
        Start-Service -Name $ServiceName -ErrorAction Stop
        Write-LogMessage "Service '$ServiceName' start command issued successfully" "SUCCESS"
        
        # Wait 30 seconds before checking if service is still running
        Write-LogMessage "Waiting 30 seconds to verify service stability..." "INFO"
        Start-Sleep -Seconds 30
        
        # Check if service is still running after 30 seconds
        $service = Get-Service -Name $ServiceName -ErrorAction Stop
        if ($service.Status -eq "Running") {
            Write-LogMessage "Service '$ServiceName' is still running after 30 seconds" "SUCCESS"
            return $true
        } else {
            Write-LogMessage "Service '$ServiceName' stopped running after 30 seconds (Status: $($service.Status))" "ERROR"
            return $false
        }
    }
    catch {
        Write-LogMessage "Failed to start service '$ServiceName': $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Remove-DeviceId {
    $basePath = "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon"
    $controlPath = "$basePath\Control"
    
    Write-LogMessage "Removing DeviceId from registry..."
    
    try {
        if (Test-RegistryKey $controlPath) {
            $deviceId = Get-RegistryValue $controlPath "DeviceId"
            if ($deviceId) {
                Remove-ItemProperty -Path $controlPath -Name "DeviceId" -ErrorAction Stop
                Write-LogMessage "DeviceId removed successfully" "SUCCESS"
                return $true
            } else {
                Write-LogMessage "DeviceId not found in registry" "WARNING"
                return $true
            }
        } else {
            Write-LogMessage "Control registry path not found" "ERROR"
            return $false
        }
    }
    catch {
        Write-LogMessage "Failed to remove DeviceId: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Remove-PublicKey {
    $publicKeyPath = "C:\ProgramData\RoboShadow\Rubicon\Control\Data\PublicKey"
    
    Write-LogMessage "Removing PublicKey file..."
    
    try {
        if (Test-Path $publicKeyPath) {
            Remove-Item -Path $publicKeyPath -Force -ErrorAction Stop
            Write-LogMessage "PublicKey file removed successfully" "SUCCESS"
            return $true
        } else {
            Write-LogMessage "PublicKey file not found" "WARNING"
            return $true
        }
    }
    catch {
        Write-LogMessage "Failed to remove PublicKey file: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Uninstall-OldAgentVersion {
    param(
        [string]$Version,
        [string]$MsiUrl,
        [string]$MsiFileName,
        [string]$RegistryKeyToDelete = $null
    )
    
    Write-LogMessage "Attempting to uninstall old agent version: $Version" "INFO"
    
    try {
        # Create download directory if it doesn't exist
        $downloadPath = "C:\ProgramData\RoboShadow\Rubicon\Updater\Download"
        if (-not (Test-Path $downloadPath)) {
            New-Item -Path $downloadPath -ItemType Directory -Force | Out-Null
            Write-LogMessage "Created download directory: $downloadPath" "INFO"
        }
        
        $fullMsiPath = Join-Path $downloadPath $MsiFileName
        
        # Download the MSI file
        Write-LogMessage "Downloading MSI file from: $MsiUrl" "INFO"
        Invoke-WebRequest -Uri $MsiUrl -OutFile $fullMsiPath -ErrorAction Stop
        Write-LogMessage "MSI file downloaded successfully" "SUCCESS"
        
        # Uninstall using MSI
        Write-LogMessage "Starting MSI uninstall process..." "INFO"
        $arguments = "/x `"$fullMsiPath`" /qb /norestart"
        $process = Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList $arguments -Wait -PassThru
        
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 1605) {
            Write-LogMessage "MSI uninstall completed (Exit code: $($process.ExitCode))" "SUCCESS"
        } else {
            Write-LogMessage "MSI uninstall returned exit code: $($process.ExitCode)" "WARNING"
        }
        
        # Delete service if it exists
        Write-LogMessage "Attempting to delete RoboShadowAgent service..." "INFO"
        $scProcess = Start-Process "sc.exe" -ArgumentList "delete", "RoboShadowAgent" -Wait -PassThru -WindowStyle Hidden
        if ($scProcess.ExitCode -eq 0) {
            Write-LogMessage "Service deleted successfully" "SUCCESS"
        } else {
            Write-LogMessage "Service deletion returned exit code: $($scProcess.ExitCode) (may not have existed)" "INFO"
        }
        
        # Delete registry key if specified
        if ($RegistryKeyToDelete) {
            Write-LogMessage "Attempting to delete registry key: $RegistryKeyToDelete" "INFO"
            try {
                if (Test-Path "Registry::$RegistryKeyToDelete") {
                    Remove-Item -Path "Registry::$RegistryKeyToDelete" -Recurse -Force -ErrorAction Stop
                    Write-LogMessage "Registry key deleted successfully" "SUCCESS"
                } else {
                    Write-LogMessage "Registry key not found (may not have existed)" "INFO"
                }
            }
            catch {
                Write-LogMessage "Error deleting registry key: $($_.Exception.Message)" "WARNING"
            }
        }
        
        # Clean up downloaded MSI file
        try {
            if (Test-Path $fullMsiPath) {
                Remove-Item $fullMsiPath -Force -ErrorAction Stop
                Write-LogMessage "Cleaned up downloaded MSI file" "INFO"
            }
        }
        catch {
            Write-LogMessage "Could not clean up MSI file: $($_.Exception.Message)" "WARNING"
        }
        
        Write-LogMessage "Old version uninstall process completed for version: $Version" "SUCCESS"
        return $true
        
    }
    catch {
        Write-LogMessage "Error during old version uninstall: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Remove-AllOldVersions {
    Write-LogMessage "Starting removal of all known old agent versions..." "INFO"
    
    # Define old versions to remove (can be expanded later)
    $oldVersions = @(
        @{
            Version = "4.8.1.25"
            MsiUrl = "https://cdn.roboshadow.com/GetAgent/RoboShadowAgent-x64-48125.msi"
            MsiFileName = "RoboShadowAgent-x64-48125.msi"
            RegistryKey = "HKEY_CLASSES_ROOT\Installer\Products\E70D8EB794AD22F4291C1A30A96CAE37"
        }
        # Additional versions can be added here in the future
    )
    
    $removalCount = 0
    
    foreach ($version in $oldVersions) {
        Write-LogMessage "Processing old version: $($version.Version)" "INFO"
        
        if (Uninstall-OldAgentVersion -Version $version.Version -MsiUrl $version.MsiUrl -MsiFileName $version.MsiFileName -RegistryKeyToDelete $version.RegistryKey) {
            $removalCount++
            Write-LogMessage "Successfully processed version: $($version.Version)" "SUCCESS"
        } else {
            Write-LogMessage "Failed to process version: $($version.Version)" "WARNING"
        }
    }
    
    Write-LogMessage "Completed removal process. Processed $removalCount out of $($oldVersions.Count) versions" "INFO"
    return $removalCount -gt 0
}

function Invoke-AgentInstallation {
    param([string]$OrganisationId)
    
    Write-LogMessage "Executing MSI installation..." "INFO"
    
    try {
        $msiUrl = "https://cdn.roboshadow.com/GetAgent/RoboShadowAgent-x64.msi"
        $arguments = "/i $msiUrl /qb /norestart ORGANISATION_ID=$OrganisationId"
        
        Write-LogMessage "Starting MSI installation process..." "INFO"
        $process = Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList $arguments -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-LogMessage "MSI installation completed successfully" "SUCCESS"
            return $true
        } else {
            Write-LogMessage "MSI installation failed with exit code: $($process.ExitCode)" "ERROR"
            return $false
        }
    }
    catch {
        Write-LogMessage "Error during MSI installation: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Install-Agent {
    param([string]$OrganisationId)
    
    Write-LogMessage "Installing RoboShadow Agent..." "INFO"
    Write-LogMessage "Organisation ID: $OrganisationId"
    
    # Try initial installation
    if (Invoke-AgentInstallation $OrganisationId) {
        Write-LogMessage "Agent installation completed successfully" "SUCCESS"
        return $true
    }
    
    # Initial installation failed, try removing old versions
    Write-LogMessage "Initial installation failed. Attempting to remove old versions..." "WARNING"
    
    if (Remove-AllOldVersions) {
        Write-LogMessage "Old versions removal completed. Retrying installation..." "INFO"
        
        # Retry installation after removing old versions
        if (Invoke-AgentInstallation $OrganisationId) {
            Write-LogMessage "Agent installation succeeded after removing old versions" "SUCCESS"
            return $true
        } else {
            Write-LogMessage "Agent installation failed again after removing old versions" "ERROR"
            return $false
        }
    } else {
        Write-LogMessage "Old version removal did not complete successfully" "ERROR"
        return $false
    }
}

function Test-PostInstallationComponents {
    param([string]$ServiceName = "RoboShadowAgent")
    
    Write-LogMessage "Verifying post-installation components..." "INFO"
    $allGood = $true
    
    # Check service exists
    if (-not (Test-ServiceExists $ServiceName)) {
        Write-LogMessage "Service verification failed" "ERROR"
        $allGood = $false
    }
    
    # Check DeviceId exists
    $basePath = "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon"
    $controlPath = "$basePath\Control"
    
    if (Test-RegistryKey $controlPath) {
        $deviceId = Get-RegistryValue $controlPath "DeviceId"
        if ($deviceId) {
            Write-LogMessage "DeviceId verification successful: $deviceId" "SUCCESS"
        } else {
            Write-LogMessage "DeviceId verification failed - not found" "ERROR"
            $allGood = $false
        }
    } else {
        Write-LogMessage "Control registry path verification failed" "ERROR"
        $allGood = $false
    }
    
    # Check PublicKey file exists
    $publicKeyPath = "C:\ProgramData\RoboShadow\Rubicon\Control\Data\PublicKey"
    if (Test-Path $publicKeyPath) {
        Write-LogMessage "PublicKey file verification successful" "SUCCESS"
    } else {
        Write-LogMessage "PublicKey file verification failed - not found" "ERROR"
        $allGood = $false
    }
    
    if ($allGood) {
        Write-LogMessage "All post-installation components verified successfully" "SUCCESS"
    } else {
        Write-LogMessage "One or more post-installation verification checks failed" "WARNING"
    }
    
    return $allGood
}

function Invoke-ReRegisterAgent {
    param([string]$ServiceName = "RoboShadowAgent")
    
    Write-LogMessage "Starting agent re-registration process..." "INFO"
    
    # Check if service is running and stop it if needed
    $serviceRunning = Test-ServiceRunning $ServiceName
    if ($serviceRunning) {
        if (-not (Stop-RoboShadowService $ServiceName)) {
            Write-LogMessage "Failed to stop service. Aborting re-registration." "ERROR"
            return $false
        }
    }
    
    # Remove PublicKey
    if (-not (Remove-PublicKey)) {
        Write-LogMessage "Failed to remove PublicKey. Continuing..." "WARNING"
    }
    
    # Remove DeviceId
    if (-not (Remove-DeviceId)) {
        Write-LogMessage "Failed to remove DeviceId. Continuing..." "WARNING"
    }
    
    # Start the service
    if (Start-RoboShadowService $ServiceName) {
        Write-LogMessage "Agent re-registration completed successfully" "SUCCESS"
        return $true
    } else {
        Write-LogMessage "Agent re-registration failed - could not restart service" "ERROR"
        return $false
    }
}

function Start-HealingProcess {
    param([string]$ExpectedOrgId)
    
    Write-LogMessage "Starting RoboShadow Agent healing process..."
    Write-LogMessage "Target OrganisationId: $ExpectedOrgId"
    
    $serviceName = "RoboShadowAgent"
    $healingRequired = $false
    
    # Check if agent is installed
    if (-not (Test-AgentInstalled)) {
        Write-LogMessage "RoboShadow Agent is not installed. Installing now..." "WARNING"
        
        if (Install-Agent $ExpectedOrgId) {
            Write-LogMessage "Agent installation successful. Verifying components..." "SUCCESS"
            
            if (Test-PostInstallationComponents $serviceName) {
                Write-LogMessage "All installation verification checks passed" "SUCCESS"
                return $true
            } else {
                Write-LogMessage "Installation verification failed" "ERROR"
                return $false
            }
        } else {
            Write-LogMessage "Agent installation failed. Cannot proceed with healing." "ERROR"
            return $false
        }
    }
    
    # Check if service exists
    if (-not (Test-ServiceExists $serviceName)) {
        Write-LogMessage "Service not found but agent is installed. This indicates an installation issue." "ERROR"
        Write-LogMessage "Manual reinstallation of RoboShadow Agent may be required." "WARNING"
        return $false
    }
    
    # Check if service is running
    $serviceRunning = Test-ServiceRunning $serviceName
    
    if ($serviceRunning) {
        # Service is running, check OrganisationId
        $currentOrgId = Get-CurrentOrganisationId
        
        if ($currentOrgId) {
            if (-not (Test-OrganisationIdMatch $ExpectedOrgId $currentOrgId)) {
                Write-LogMessage "OrganisationId mismatch detected. Re-registration required." "WARNING"
                $healingRequired = $true
            } else {
                Write-LogMessage "OrganisationId matches. No healing required." "SUCCESS"
                return $true
            }
        } else {
            Write-LogMessage "Could not retrieve current OrganisationId. Re-registration required." "WARNING"
            $healingRequired = $true
        }
    } else {
        Write-LogMessage "Service is not running. Attempting to start service..." "INFO"
        if (Start-RoboShadowService $serviceName) {
            Write-LogMessage "Service started successfully" "SUCCESS"
            return $true
        } else {
            Write-LogMessage "Service failed to start. Re-registration required." "WARNING"
            return Invoke-ReRegisterAgent $serviceName
        }
    }
    
    if ($healingRequired) {
        return Invoke-ReRegisterAgent $serviceName
    }
    
    return $true
}

# Main execution
Clear-Host
$separator = "=" * 70
Write-Host $separator -ForegroundColor Cyan
Write-Host "    ROBOSHADOW AGENT HEALING UTILITY" -ForegroundColor White
Write-Host $separator -ForegroundColor Cyan
Write-Host "Date/Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host "Target Organisation ID: $OrganisationId" -ForegroundColor Gray
Write-Host ""

$result = Start-HealingProcess $OrganisationId

Write-Host ""
Write-Host $separator -ForegroundColor Cyan
if ($result) {
    Write-Host "    HEALING COMPLETED SUCCESSFULLY" -ForegroundColor Green
} else {
    Write-Host "    HEALING FAILED" -ForegroundColor Red
}
Write-Host $separator -ForegroundColor Cyan
Write-Host ""