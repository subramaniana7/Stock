param(
 [string]$ReportFolder=$PSScriptRoot,
 [string]$PythonExecutable=(Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe')
)
$ErrorActionPreference='Stop'
$nseFolder=(Resolve-Path -LiteralPath $ReportFolder).Path
$nsePython=(Resolve-Path -LiteralPath $PythonExecutable).Path
$nseTaskName='NSE Monthly OHLC'
if ((Get-TimeZone).Id -ne 'India Standard Time') { throw 'Expected India Standard Time' }
if (Get-ScheduledTask -TaskName $nseTaskName -ErrorAction SilentlyContinue) { throw 'Monthly task already exists; preserving it' }
$nseUser=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$nseNow=Get-Date
$nseFirst=([datetime]::new($nseNow.Year,$nseNow.Month,1,9,0,0)).AddMonths(1).ToString('yyyy-MM-ddTHH:mm:ss')+'+05:30'
$nseEscUser=[System.Security.SecurityElement]::Escape($nseUser)
$nseEscPython=[System.Security.SecurityElement]::Escape($nsePython)
$nseEscFolder=[System.Security.SecurityElement]::Escape($nseFolder)
$nseScript=[System.Security.SecurityElement]::Escape('"'+(Join-Path $nseFolder 'monthly_update.py')+'"')
$nseXml=@"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
 <RegistrationInfo><Description>Monthly NSE cash OHLC for F&amp;O stocks. First of every month, 09:00 IST; catch-up after missed starts and sign-in. Separate monthly reports and email notifications.</Description></RegistrationInfo>
 <Triggers>
  <CalendarTrigger><StartBoundary>$nseFirst</StartBoundary><Enabled>true</Enabled><ScheduleByMonth><DaysOfMonth><Day>1</Day></DaysOfMonth><Months><January/><February/><March/><April/><May/><June/><July/><August/><September/><October/><November/><December/></Months></ScheduleByMonth></CalendarTrigger>
  <LogonTrigger><Enabled>true</Enabled><Delay>PT1M</Delay><UserId>$nseEscUser</UserId></LogonTrigger>
 </Triggers>
 <Principals><Principal id="Author"><UserId>$nseEscUser</UserId><LogonType>InteractiveToken</LogonType><RunLevel>LeastPrivilege</RunLevel></Principal></Principals>
 <Settings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy><DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries><StopIfGoingOnBatteries>false</StopIfGoingOnBatteries><AllowHardTerminate>true</AllowHardTerminate><StartWhenAvailable>true</StartWhenAvailable><AllowStartOnDemand>true</AllowStartOnDemand><Enabled>true</Enabled><WakeToRun>true</WakeToRun><ExecutionTimeLimit>PT2H</ExecutionTimeLimit><RestartOnFailure><Interval>PT15M</Interval><Count>3</Count></RestartOnFailure></Settings>
 <Actions Context="Author"><Exec><Command>$nseEscPython</Command><Arguments>$nseScript</Arguments><WorkingDirectory>$nseEscFolder</WorkingDirectory></Exec></Actions>
</Task>
"@
Register-ScheduledTask -TaskName $nseTaskName -Xml $nseXml | Out-Null
Export-ScheduledTask -TaskName $nseTaskName | Set-Content -LiteralPath (Join-Path $nseFolder 'Monthly Task Scheduler.xml') -Encoding utf8
Get-ScheduledTaskInfo -TaskName $nseTaskName | Select-Object NextRunTime,LastTaskResult | ConvertTo-Json
