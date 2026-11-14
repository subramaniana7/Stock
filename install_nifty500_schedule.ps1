param(
 [Parameter(Mandatory=$true)][ValidateSet('weekly','monthly')][string]$Kind,
 [string]$ReportFolder=$PSScriptRoot,
 [string]$PythonExecutable=(Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe')
)
$ErrorActionPreference='Stop'
$nseFolder=(Resolve-Path -LiteralPath $ReportFolder).Path
$nsePython=(Resolve-Path -LiteralPath $PythonExecutable).Path
if((Get-TimeZone).Id -ne 'India Standard Time'){throw 'Expected India Standard Time'}
$nseLabel=if($Kind -eq 'weekly'){'Weekly'}else{'Monthly'}
$nseName='NSE500 '+$nseLabel+' OHLC'
if(Get-ScheduledTask -TaskName $nseName -ErrorAction SilentlyContinue){throw 'Task already exists; preserving it'}
$nseUser=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$nseNow=Get-Date
if($Kind -eq 'weekly'){
 $nseFirst=$nseNow.Date.AddDays((6-[int]$nseNow.DayOfWeek+7)%7).AddHours(9)
 if($nseFirst -le $nseNow){$nseFirst=$nseFirst.AddDays(7)}
 $nseSchedule='<ScheduleByWeek><WeeksInterval>1</WeeksInterval><DaysOfWeek><Saturday/></DaysOfWeek></ScheduleByWeek>'
}else{
 $nseFirst=([datetime]::new($nseNow.Year,$nseNow.Month,1,9,0,0)).AddMonths(1)
 $nseSchedule='<ScheduleByMonth><DaysOfMonth><Day>1</Day></DaysOfMonth><Months><January/><February/><March/><April/><May/><June/><July/><August/><September/><October/><November/><December/></Months></ScheduleByMonth>'
}
$nseBoundary=$nseFirst.ToString('yyyy-MM-ddTHH:mm:ss')+'+05:30'
$nseEscUser=[System.Security.SecurityElement]::Escape($nseUser)
$nseEscPython=[System.Security.SecurityElement]::Escape($nsePython)
$nseEscFolder=[System.Security.SecurityElement]::Escape($nseFolder)
$nseArgs=[System.Security.SecurityElement]::Escape('"'+(Join-Path $nseFolder 'nifty500_update.py')+'" --kind '+$Kind)
$nseXml=@"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
 <RegistrationInfo><Description>Independent Nifty 500 $Kind cash OHLC. Current official constituents, dated snapshot, coverage and CPR. Email notification without attachments.</Description></RegistrationInfo>
 <Triggers>
  <CalendarTrigger><StartBoundary>$nseBoundary</StartBoundary><Enabled>true</Enabled>$nseSchedule</CalendarTrigger>
  <LogonTrigger><Enabled>true</Enabled><Delay>PT1M</Delay><UserId>$nseEscUser</UserId></LogonTrigger>
 </Triggers>
 <Principals><Principal id="Author"><UserId>$nseEscUser</UserId><LogonType>InteractiveToken</LogonType><RunLevel>LeastPrivilege</RunLevel></Principal></Principals>
 <Settings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy><DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries><StopIfGoingOnBatteries>false</StopIfGoingOnBatteries><AllowHardTerminate>true</AllowHardTerminate><StartWhenAvailable>true</StartWhenAvailable><AllowStartOnDemand>true</AllowStartOnDemand><Enabled>true</Enabled><WakeToRun>true</WakeToRun><ExecutionTimeLimit>PT2H</ExecutionTimeLimit><RestartOnFailure><Interval>PT15M</Interval><Count>3</Count></RestartOnFailure></Settings>
 <Actions Context="Author"><Exec><Command>$nseEscPython</Command><Arguments>$nseArgs</Arguments><WorkingDirectory>$nseEscFolder</WorkingDirectory></Exec></Actions>
</Task>
"@
Register-ScheduledTask -TaskName $nseName -Xml $nseXml | Out-Null
Export-ScheduledTask -TaskName $nseName | Set-Content -LiteralPath (Join-Path $nseFolder ($nseName+' Task Scheduler.xml')) -Encoding utf8
Get-ScheduledTaskInfo -TaskName $nseName | Select-Object NextRunTime,LastTaskResult | ConvertTo-Json
