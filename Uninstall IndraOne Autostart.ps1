$ErrorActionPreference = "Stop"
$taskName = "IndraOne Server"

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "Autostart IndraOne berhasil dihapus." -ForegroundColor Green
} else {
    Write-Host "Task autostart IndraOne tidak ditemukan."
}
