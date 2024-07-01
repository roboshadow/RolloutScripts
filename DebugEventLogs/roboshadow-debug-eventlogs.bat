@echo off
SET "outputPath=%~dp0robologs.csv"

PowerShell -Command "Get-EventLog -LogName Application | Where-Object { $_.Source -like '*RoboShadow*' } | Export-Csv -Path '%outputPath%' -NoTypeInformation"

echo Event logs exported to %outputPath%
pause
