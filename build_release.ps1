$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$distRoot = Join-Path $projectRoot "dist"
$workRoot = Join-Path $projectRoot "build"
$releaseRoot = Join-Path $projectRoot "release"
$appName = "Gestione Certificati DPI"
$packageRoot = Join-Path $distRoot $appName
$zipPath = Join-Path $releaseRoot "Gestione_Certificati_DPI.zip"

$dataFiles = @(
    "irudek_norme.json",
    "logo con scrittura .jpg"
)

$templateFiles = Get-ChildItem -Path $projectRoot -Filter "*.docx" | Sort-Object Name

if (Test-Path $packageRoot) {
    Remove-Item -Recurse -Force $packageRoot
}
if (Test-Path $zipPath) {
    Remove-Item -Force $zipPath
}

New-Item -ItemType Directory -Force -Path $distRoot | Out-Null
New-Item -ItemType Directory -Force -Path $releaseRoot | Out-Null

$pyiArgs = @(
    "--noconfirm",
    "--clean",
    "--windowed",
    "--name", $appName,
    "--icon", (Join-Path $projectRoot "gestione_certificati_dpi.ico"),
    "--distpath", $distRoot,
    "--workpath", $workRoot,
    "--specpath", $workRoot
)

foreach ($file in $dataFiles) {
    $pyiArgs += "--add-data"
    $pyiArgs += "$(Join-Path $projectRoot $file);."
}

foreach ($file in $templateFiles) {
    $pyiArgs += "--add-data"
    $pyiArgs += "$($file.FullName);."
}

$pyiArgs += (Join-Path $projectRoot "app.py")

python -m PyInstaller @pyiArgs

Copy-Item (Join-Path $projectRoot "README.md") -Destination $packageRoot -Force

$launcherBat = @"
@echo off
cd /d "%~dp0"
start "" "%~dp0$appName.exe"
"@
Set-Content -Path (Join-Path $packageRoot "avvia_programma.bat") -Value $launcherBat -Encoding ASCII

$launcherVbs = @"
Set shell = CreateObject("WScript.Shell")
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
command = Chr(34) & scriptDir & "\$appName.exe" & Chr(34)
shell.Run command, 0, False
"@
Set-Content -Path (Join-Path $packageRoot "avvia_programma.vbs") -Value $launcherVbs -Encoding ASCII

Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipPath -Force

Write-Host "Pacchetto creato in: $packageRoot"
Write-Host "Archivio ZIP creato in: $zipPath"
