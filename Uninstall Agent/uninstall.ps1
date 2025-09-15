function Uninstall-Service {    param (        [string]$Name    )    $uninstallKey = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" |        Get-ItemProperty |        Where-Object { $_.DisplayName -like $Name } |        Select-Object -ExpandProperty UninstallString    Write-Host $uninstallKey    if ($uninstallKey) {        $msiexecPath = Join-Path $env:SystemRoot "System32"        Write-Host $msiexecPath        $msiexecPath = Join-Path $env:SystemRoot "System32\msiexec.exe"        $arguments = ($uninstallKey -split 'MsiExec.exe', 2)[1].Trim()        $arguments = $arguments -replace '^/I', '/X'        $arguments = "$arguments /q /norestart"        $process = Start-Process -FilePath $msiexecPath -ArgumentList $arguments -Wait -PassThru         Write-Host "Exit Code: $($process.ExitCode)"    } else { Write-Host "$Name not installed"}}Uninstall-Service -Name "RoboShadow Agent"Uninstall-Service -Name "RoboShadow Update Service*"$registryPathControl = "HKLM:\SOFTWARE\RoboShadowLtd"if (Test-Path $registryPathControl) {    Remove-Item -Path $registryPathControl -Recurse -Force}$programData = "C:\ProgramData\RoboShadow"if (Test-Path $programData){    Remove-Item -Path $programData -Recurse -Force}function Uninstall-Service {
	param (
        [string]$Name
    )
    $uninstallKey = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" |
        Get-ItemProperty |
        Where-Object { $_.DisplayName -like $Name } |
        Select-Object -ExpandProperty UninstallString
    Write-Host $uninstallKey
    if ($uninstallKey) {
        $msiexecPath = Join-Path $env:SystemRoot "System32"
        Write-Host $msiexecPath
        $msiexecPath = Join-Path $env:SystemRoot "System32\msiexec.exe"
        $arguments = ($uninstallKey -split 'MsiExec.exe', 2)[1].Trim()
        $arguments = $arguments -replace '^/I', '/X'
        $arguments = "$arguments /q /norestart"
        $process = Start-Process -FilePath $msiexecPath -ArgumentList $arguments -Wait -PassThru 
        Write-Host "Exit Code: $($process.ExitCode)"
    } else { Write-Host "$Name not installed"}
}

Uninstall-Service -Name "RoboShadow Agent"
Uninstall-Service -Name "RoboShadow Update Service*"

$registryPathControl = "HKLM:\SOFTWARE\RoboShadowLtd"
if (Test-Path $registryPathControl) {
    Remove-Item -Path $registryPathControl -Recurse -Force
}

$programData = "C:\ProgramData\RoboShadow"
if (Test-Path $programData)
{
    Remove-Item -Path $programData -Recurse -Force
}