# ********************************************
# Make sure you double check this OrganisationID changing organisations is not trivial.
# 
# If you are deploying via intune I would recommend using our intune integration feature.
# ********************************************

$organisationId = "YOUR_ORGANISATION_ID"

$version = (Get-ItemProperty -Path "HKLM:\SOFTWARE\RoboShadowLtd\Rubicon\Agent" -Name "Version" -ErrorAction SilentlyContinue).$valueName

if (-not $version -or [int]($version -split '\.')[0] -lt 4) {
  Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList "/i https://cdn.roboshadow.com/GetAgent/RoboShadowAgent-x64.msi /qb /norestart ORGANISATION_ID=$organisationId" -Wait
}