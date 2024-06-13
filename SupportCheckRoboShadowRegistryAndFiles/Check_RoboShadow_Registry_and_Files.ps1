# Function to check registry value
function Get-RegistryValue {
    param (
        [string]$Path,
        [string]$Name
    )
    try {
        $value = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return $value.$Name
    } catch {
        Write-Output "Error retrieving $Name from $Path"
        return $null
    }
}

# Define registry paths and value names
$agentRegistryPath = "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon\Agent"
$controlRegistryPath = "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon\Control"
$agentValues = @("OrganisationId", "Version")
$controlValues = @("DeviceId", "OrganisationId", "Version")

# Retrieve values from Agent registry location
$agentResults = @{}
foreach ($value in $agentValues) {
    $agentResults[$value] = Get-RegistryValue -Path $agentRegistryPath -Name $value
}

# Retrieve values from Control registry location
$controlResults = @{}
foreach ($value in $controlValues) {
    $controlResults[$value] = Get-RegistryValue -Path $controlRegistryPath -Name $value
}

# Get machine name
$machineName = $env:COMPUTERNAME

# Check if PublicKey file exists
$publicKeyPath = "C:\ProgramData\RoboShadow\Rubicon\Control\Data\PublicKey"
$fileExists = Test-Path -Path $publicKeyPath

# Get user's Documents folder
$documentsFolder = [environment]::GetFolderPath("MyDocuments")

# Output results to a file in the user's Documents folder
$outputFile = "$documentsFolder\${machineName}_output.txt"

"Machine Name: $machineName" | Out-File -FilePath $outputFile -Append
"`nAgent Registry Values:" | Out-File -FilePath $outputFile -Append
$agentResults.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" | Out-File -FilePath $outputFile -Append }

"`nControl Registry Values:" | Out-File -FilePath $outputFile -Append
$controlResults.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" | Out-File -FilePath $outputFile -Append }

"`nPublicKey File Exists: $fileExists" | Out-File -FilePath $outputFile -Append

Write-Output "Results have been saved to $outputFile"
