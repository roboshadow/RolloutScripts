# ********************************************
# If you are deploying via intune I would recommend using our intune integration feature.
# 
# How to use:
# Replace YOUR_ORGANISATION_ID with the organisationID from the correct organisation in the portal. Note changing this after install is not trvial.
# Set $haveSetOrgId to $True
# ********************************************

$organisationId = "YOUR_ORGANISATION_ID"
$haveSetOrgId = $False

$version = (Get-ItemProperty -Path "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon\Agent" -Name "Version" -ErrorAction SilentlyContinue).$valueName

if ($haveSetOrgId -and (-not $version -or [int]($version -split '\.')[0] -lt 4)) {
  Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList "/i https://cdn.roboshadow.com/GetAgent/RoboShadowAgent-x64.msi /qb /norestart ORGANISATION_ID=$organisationId" -Wait
}