param(
 [string]$ReportFolder = $PSScriptRoot,
 [Parameter(Mandatory=$true)][string]$PythonExecutable
)
$ErrorActionPreference='Stop'
$nseLocal=(Resolve-Path -LiteralPath $ReportFolder).Path
$nsePython=(Resolve-Path -LiteralPath $PythonExecutable).Path
$nseDailyName='NSE Daily OHLC'
if ((Get-TimeZone).Id -ne 'India Standard Time') { throw 'Expected India Standard Time' }
if (Get-ScheduledTask -TaskName $nseDailyName -ErrorAction SilentlyContinue) { throw 'Daily task already exists; preserving it' }
if (-not (Test-Path -LiteralPath $nsePython -PathType Leaf)) { throw 'Python runtime unavailable' }
$nseUser=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$nsePoll=New-ScheduledTaskTrigger -Daily -At ([datetime]::Today.AddHours(16).AddMinutes(30))
$nseRepetition=(New-ScheduledTaskTrigger -Once -At ([datetime]::Today.AddHours(16).AddMinutes(30)) -RepetitionInterval (New-TimeSpan -Minutes 15) -RepetitionDuration (New-TimeSpan -Hours 5 -Minutes 31)).Repetition
$nsePoll.Repetition=$nseRepetition
$nseMorning=New-ScheduledTaskTrigger -Daily -At ([datetime]::Today.AddHours(9))
$nseLogon=New-ScheduledTaskTrigger -AtLogOn -User $nseUser
$nseLogon.Delay='PT1M'
$nseAction=New-ScheduledTaskAction -Execute $nsePython -Argument ('"'+(Join-Path $nseLocal 'daily_update.py')+'"') -WorkingDirectory $nseLocal
$nsePrincipal=New-ScheduledTaskPrincipal -UserId $nseUser -LogonType Interactive -RunLevel Limited
$nseSettings=New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 15) -ExecutionTimeLimit (New-TimeSpan -Hours 2) -MultipleInstances IgnoreNew
$nseTask=New-ScheduledTask -Action $nseAction -Trigger @($nsePoll,$nseMorning,$nseLogon) -Principal $nsePrincipal -Settings $nseSettings -Description 'Daily NSE cash OHLC for each day F&O stocks. Availability checks 16:30-22:00 IST every 15 minutes; 09:00 and sign-in catch-up. Separate daily-sheets workbook per week, local outputs. Independent of Saturday weekly task.'
Register-ScheduledTask -TaskName $nseDailyName -InputObject $nseTask | Out-Null
Export-ScheduledTask -TaskName $nseDailyName | Set-Content -LiteralPath (Join-Path $nseLocal 'Daily Task Scheduler.xml') -Encoding utf8
Get-ScheduledTaskInfo -TaskName $nseDailyName | Select-Object NextRunTime,LastTaskResult | ConvertTo-Json
