param (
    [string]$apiKey = $env:ROBOSHADOW_RMM_KEY,
    [string]$organisationName = $env:CS_PROFILE_NAME
)
function Write-Log {
    param([string]$msg)
    Write-Output "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg"
}
# If you're targeting older PS (e.g., PS 2.0/3.0), forcing TLS 1.2 avoids random IRM failures
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$serviceName = "RoboShadowAgent"

if ([string]::IsNullOrWhiteSpace($apiKey)) {
    Write-Error "API key missing. Ensure ROBOSHADOW_RMM_KEY is set as a Datto Component Input Variable / global variable."
    exit 1
}
try {
    $service = Get-Service -Name $serviceName -ErrorAction Stop
    if ($service.Status -eq 'Running') {
        Write-Log "Service '$serviceName' is already running. Skipping installation."
        exit 0
    } else {
        Write-Log "Service '$serviceName' exists but is not running."
    }
} catch {
    Write-Log "Service '$serviceName' not found. Proceeding with installation."
}
if (-not $organisationName) {
    Write-Error "Organisation name is missing. Either pass it via -organisationName or ensure CS_PROFILE_NAME is set."
    exit 1
}


if ([string]::IsNullOrWhiteSpace($organisationName)) {
    Write-Log "ERROR: CS_PROFILE_NAME is not available."
    exit 1
}

Write-Log "Starting deployment for organisation: $organisationName"

# Step 1: Call RoboShadow API to initiate organisation creation
$headers = @{ "Content-Type" = "application/json" }
$body = @{
    apiKey = $apiKey
    organisationName = $organisationName
} | ConvertTo-Json -Depth 3

try {
    $response = Invoke-RestMethod -Uri "https://api.roboshadow.com/identity/rmm/organisation" -Method Post -Headers $headers -Body $body
} catch {
    Write-Log "ERROR: Failed to initiate organisation creation. $_"
    exit 1
}

# Step 2: Poll for externalId
$orgId = $null
$maxAttempts = 15
$attempt = 0

while (-not $orgId -and $attempt -lt $maxAttempts) {
    Start-Sleep -Seconds 2
    try {
        $pollResponse = Invoke-RestMethod -Uri "https://api.roboshadow.com/identity/rmm/organisation" -Method Post -Headers $headers -Body $body
        if ($pollResponse.externalId) {
            $orgId = $pollResponse.externalId
            Write-Log "Received Organisation ID: $orgId"
        } else {
            Write-Log "Polling attempt $($attempt + 1): Organisation ID not yet available."
        }
    } catch {
        Write-Log "Polling attempt $($attempt + 1) failed: $_"
    }
    $attempt++
}

if (-not $orgId) {
    Write-Log "ERROR: Failed to retrieve Organisation ID after $maxAttempts attempts."
    exit 1
}

# Step 3: Install the RoboShadow Agent
$agentUrl = "https://cdn.roboshadow.com/GetAgent/RoboShadowAgent-x64.msi"
$msiArgs = "/i `"$agentUrl`" /qn /norestart ORGANISATION_ID=$orgId"

try {
    Write-Log "Starting RoboShadow Agent installation..."
    Start-Process "msiexec.exe" -ArgumentList $msiArgs -Wait -NoNewWindow
    Write-Log "Agent installation completed successfully."
} catch {
    Write-Log "ERROR: Agent installation failed. $_"
    exit 1
}
