param(
 [string]$ReportFolder = $PSScriptRoot,
 [Parameter(Mandatory=$true)][string]$PythonExecutable
)
$ErrorActionPreference = 'Stop'
$nseFolder = (Resolve-Path -LiteralPath $ReportFolder).Path
$nsePython = (Resolve-Path -LiteralPath $PythonExecutable).Path
$nseTaskName = 'NSE Weekly OHLC'
if (-not (Test-Path -LiteralPath $nsePython -PathType Leaf)) { throw 'Python runtime missing' }
if (-not (Test-Path -LiteralPath (Join-Path $nseFolder 'scheduled_update.py') -PathType Leaf)) { throw 'Weekly runner missing' }
if ((Get-TimeZone).Id -ne 'India Standard Time') { throw 'Windows timezone is not India Standard Time; schedule needs review' }
if (Get-ScheduledTask -TaskName $nseTaskName -ErrorAction SilentlyContinue) { throw 'Task already exists; no duplicate created' }
$nseUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$nseAction = New-ScheduledTaskAction -Execute $nsePython -Argument ('"' + (Join-Path $nseFolder 'scheduled_update.py') + '"') -WorkingDirectory $nseFolder
$nseConfig = Get-Content -LiteralPath (Join-Path $nseFolder 'schedule_config.json') -Raw | ConvertFrom-Json
$nseFirst = [DateTimeOffset]::Parse($nseConfig.first_run)
if ($nseFirst.DayOfWeek -ne [DayOfWeek]::Saturday -or $nseFirst.Hour -ne 9 -or $nseFirst.Minute -ne 0 -or $nseFirst.Offset -ne [TimeSpan]::FromMinutes(330)) { throw 'first_run must be Saturday 09:00 IST' }
$nseWeekly = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Saturday -WeeksInterval 1 -At $nseFirst.LocalDateTime
$nseLogon = New-ScheduledTaskTrigger -AtLogOn -User $nseUser
$nseLogon.Delay = 'PT1M'
$nsePrincipal = New-ScheduledTaskPrincipal -UserId $nseUser -LogonType Interactive -RunLevel Limited
$nseSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 15) -ExecutionTimeLimit (New-TimeSpan -Hours 2) -MultipleInstances IgnoreNew
$nseTask = New-ScheduledTask -Action $nseAction -Trigger @($nseWeekly,$nseLogon) -Principal $nsePrincipal -Settings $nseSettings -Description 'NSE weekly cash OHLC for F&O stocks. Saturday 09:00 IST; catch up after missed runs and at sign-in; local files outside OneDrive. Retries three times at 15-minute intervals. Requires signed-in user and internet.'
Register-ScheduledTask -TaskName $nseTaskName -InputObject $nseTask | Select-Object TaskName,State
Export-ScheduledTask -TaskName $nseTaskName | Set-Content -LiteralPath (Join-Path $nseFolder 'Windows Task Scheduler.xml') -Encoding utf8
Get-ScheduledTaskInfo -TaskName $nseTaskName | Select-Object NextRunTime,LastRunTime,LastTaskResult
