[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$taskName = 'FinancialTrackerDailyUpdate'
$scriptPath = Join-Path $PSScriptRoot 'RunMacro.vbs'
$wscriptPath = Join-Path $env:WINDIR 'System32\wscript.exe'

if (-not (Test-Path -LiteralPath $scriptPath)) { throw "RunMacro.vbs was not found at $scriptPath" }
if (-not (Test-Path -LiteralPath $wscriptPath)) { throw "wscript.exe was not found at $wscriptPath" }

$action = New-ScheduledTaskAction -Execute $wscriptPath -Argument ('"{0}"' -f $scriptPath) -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At '16:15'
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 15) -StartWhenAvailable
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description 'Runs the portable Excel financial tracker daily update.' -Force | Out-Null
Write-Host "Installed scheduled task: $taskName"
