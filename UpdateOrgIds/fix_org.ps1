
$keyName = "OrganisationId"
$newValue = "6caca2b0-7df0-4f47-b4f2-fdcb86b7f820"

# Define the registry path, key name, and new value for OrganisationId
$registryPathControl = "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon\Control"
$registryPathAgent = "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon\Agent"
$registryPathUpdater = "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon\Updater"

$orgIdValueControl = $null
if (Test-Path $registryPathControl) {
    $orgIdValueControl = Get-ItemPropertyValue -Path $registryPathControl -Name $keyName -ErrorAction SilentlyContinue
}

$orgIdValueAgent = $null
if (Test-Path $registryPathAgent) {
    $orgIdValueAgent = Get-ItemPropertyValue -Path $registryPathAgent -Name $keyName -ErrorAction SilentlyContinue
}

if ($orgIdValueControl -eq $newValue -and $orgIdValueAgent -eq $newValue) {
    Write-Output "Script has already run for both Control and Agent. Exiting."
    exit
}

# Service names
$services = @('RoboShadowControl', 'RoboShadowAgent')

# Stop the services
foreach ($service in $services) {
    if (Get-Service -Name $service -ErrorAction SilentlyContinue) {
        Stop-Service -Name $service -Force
        Write-Output "Stopped service: $service"
    } else {
        Write-Error "Service $service not found."
    }
}

# Update the OrganisationId for the Agent registry path
if (Test-Path $registryPathAgent) {
    Set-ItemProperty -Path $registryPathAgent -Name $keyName -Value $newValue
    Write-Output "OrganisationId registry value for Agent updated successfully."
} else {
    Write-Error "Registry path $registryPathAgent does not exist."
}

# Update the OrganisationId for the Updater registry path
if (Test-Path $registryPathUpdater) {
    Set-ItemProperty -Path $registryPathUpdater -Name $keyName -Value $newValue
    Write-Output "OrganisationId registry value for Updater updated successfully."
} else {
    Write-Error "Registry path $registryPathUpdater does not exist."
}

# Check if the registry key exists
if (Test-Path $registryPathControl) {

    # Set the registry value for OrganisationId
    Set-ItemProperty -Path $registryPathControl -Name $keyName -Value $newValue
    Write-Output "OrganisationId registry value updated successfully."
    
    # Delete the DeviceId registry value
    try {
        $value = Get-ItemPropertyValue -Path $registryPathControl -Name "DeviceId" -ErrorAction Stop
        if ($value) {
            Remove-ItemProperty -Path $registryPathControl -Name "DeviceId"
            Write-Output "DeviceId registry value deleted successfully."
        }
    } catch {
        Write-Error "Registry value DeviceId at path $registryPathControl does not exist."
    }
} else {
    Write-Error "Registry path $registryPathControl does not exist."
}


# Delete the specified file
$fileToDelete = "c:\ProgramData\RoboShadow\Rubicon\Control\Data\PublicKey"
if (Test-Path $fileToDelete) {
    Remove-Item -Path $fileToDelete -Force
    Write-Output "Deleted file: $fileToDelete"
} else {
    Write-Error "File $fileToDelete does not exist."
}

# Start the services
foreach ($service in $services) {
    if (Get-Service -Name $service -ErrorAction SilentlyContinue) {
        Start-Service -Name $service -ErrorAction SilentlyContinue
        Write-Output "Started service: $service"
    } else {
        Write-Error "Service $service not found."
    }
}
