$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$launcher = Join-Path $projectDir "Run IndraOne.bat"
$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "IndraOne Server.lnk"

if (-not (Test-Path -LiteralPath $launcher)) {
    throw "File launcher tidak ditemukan: $launcher"
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $launcher
$shortcut.WorkingDirectory = $projectDir
$shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,220"
$shortcut.Description = "Jalankan aplikasi IndraOne"
$shortcut.Save()

Write-Host "Shortcut berhasil dibuat: $shortcutPath" -ForegroundColor Green
