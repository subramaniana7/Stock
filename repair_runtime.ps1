param(
 [string]$ReportFolder = $PSScriptRoot,
 [string]$DependencyRoot = (Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies')
)
$ErrorActionPreference='Stop'
$nseFolder=(Resolve-Path -LiteralPath $ReportFolder).Path
$nsePython=Join-Path $DependencyRoot 'python\pythonw.exe'
$nseNode=Join-Path $DependencyRoot 'node\bin\node.exe'
$nsePackages=Join-Path $DependencyRoot 'node\node_modules'
foreach ($nseRequired in @($nsePython,$nseNode,(Join-Path $nsePackages '@oai\artifact-tool\package.json'))) {
 if (-not (Test-Path -LiteralPath $nseRequired -PathType Leaf)) { throw "Stable runtime is unavailable: $nseRequired" }
}
$nseDaily=Get-ScheduledTask -TaskName 'NSE Daily OHLC'
if ($nseDaily.State -eq 'Running') { throw 'Daily task is active; rerun repair after it finishes' }
$nseBefore=[xml](Export-ScheduledTask -TaskName 'NSE Daily OHLC')
$nseWeeklyBefore=(Export-ScheduledTask -TaskName 'NSE Weekly OHLC').Trim()
$nseBackup=Join-Path $nseFolder ('backups\runtime_repair_'+(Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $nseBackup -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $nseFolder 'daily_config.json') -Destination $nseBackup
$nseBefore.Save((Join-Path $nseBackup 'Daily Task Scheduler.xml'))
$nseConfigPath=Join-Path $nseFolder 'daily_config.json'
$nseConfig=Get-Content -LiteralPath $nseConfigPath -Raw | ConvertFrom-Json
$nseConfig.node_executable=$nseNode
$nseConfig | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $nseConfigPath -Encoding utf8
$nseLink=Join-Path $nseFolder 'daily_runtime\node_modules'
$nseExisting=Get-Item -LiteralPath $nseLink -Force -ErrorAction SilentlyContinue
if ($nseExisting) {
 if ($nseExisting.LinkType -ne 'Junction') { throw "Expected a generated junction; preserving actual directory: $nseLink" }
 # Delete only this verified junction, with no recursive operation on its target.
 [System.IO.Directory]::Delete($nseLink)
}
New-Item -ItemType Junction -Path $nseLink -Target $nsePackages | Out-Null
$nseAction=New-ScheduledTaskAction -Execute $nsePython -Argument ('"'+(Join-Path $nseFolder 'daily_update.py')+'"') -WorkingDirectory $nseFolder
Set-ScheduledTask -TaskName 'NSE Daily OHLC' -Action $nseAction | Out-Null
$nseAfter=[xml](Export-ScheduledTask -TaskName 'NSE Daily OHLC')
foreach ($nseSection in @('Triggers','Settings','Principals')) {
 if ($nseBefore.Task.$nseSection.OuterXml -ne $nseAfter.Task.$nseSection.OuterXml) { throw "Unexpected change to daily task $nseSection" }
}
if ((Export-ScheduledTask -TaskName 'NSE Weekly OHLC').Trim() -ne $nseWeeklyBefore) { throw 'Weekly task changed unexpectedly' }
Export-ScheduledTask -TaskName 'NSE Daily OHLC' | Set-Content -LiteralPath (Join-Path $nseFolder 'Daily Task Scheduler.xml') -Encoding utf8
Write-Output 'Daily Python action, shared Node configuration and workbook-package link now use stable cache paths.'
Write-Output 'Daily triggers/settings/principal and the entire weekly task are unchanged.'
