$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$launcher = Join-Path $projectDir "Run IndraOne.bat"
$taskName = "IndraOne Server"

if (-not (Test-Path -LiteralPath $launcher)) {
    throw "File launcher tidak ditemukan: $launcher"
}

$action = New-ScheduledTaskAction -Execute $launcher -WorkingDirectory $projectDir
$trigger = New-ScheduledTaskTrigger -AtLogOn
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DisallowStartIfOnBatteries:$false -ExecutionTimeLimit (New-TimeSpan -Hours 0)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null

Write-Host "Autostart IndraOne berhasil dipasang." -ForegroundColor Green
Write-Host "Aplikasi akan otomatis berjalan saat user Windows login."
Write-Host "Untuk menjalankan sekarang, double-click: $launcher"
