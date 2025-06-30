```
Invoke-WebRequest 'https://raw.githubusercontent.com/roboshadow/RolloutScripts/refs/heads/master/Support/RoboShadowDiagnostic.ps1' -OutFile 'RoboShadowDiagnostic.ps1'; .\RoboShadowDiagnostic.ps1 -OrganisationId "YOUR_ORG_ID"
$script = Invoke-WebRequest 'https://raw.githubusercontent.com/roboshadow/RolloutScripts/refs/heads/master/Support/RoboShadowDiagnostic.ps1' -UseBasicParsing; Invoke-Expression $script.Content

```