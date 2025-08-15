param(
  [Parameter(Mandatory=$true)][string]$PythonPath,
  [Parameter(Mandatory=$true)][string]$ProjectDir,
  [string]$TaskName = "WeeklyCsvBackup",
  [string]$StartTime = "06:00", # 24h HH:mm
  [int]$DaysInterval = 7
)

$scriptPath = Join-Path $ProjectDir "backup.py"
if (!(Test-Path $scriptPath)) {
  Write-Error "backup.py not found at $scriptPath"; exit 1
}

$action = New-ScheduledTaskAction -Execute $PythonPath -Argument "`"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -Once -At ([datetime]::ParseExact($StartTime, 'HH:mm', $null)) -RepetitionInterval (New-TimeSpan -Days $DaysInterval)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel LeastPrivilege
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

try {
  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
  Write-Host "Created/updated task '$TaskName' to run every $DaysInterval days at $StartTime"
}
catch {
  Write-Error $_.Exception.Message
  exit 1
}
