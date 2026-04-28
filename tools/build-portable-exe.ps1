$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$releaseDir = Join-Path $repoRoot 'src-tauri\target\release'
$stageRoot = Join-Path $repoRoot 'build\portable'
$payloadDir = Join-Path $stageRoot 'payload'
$payloadZip = Join-Path $stageRoot 'payload.zip'
$outputDir = Join-Path $repoRoot 'dist-portable'
$outputExe = Join-Path $outputDir 'eMerge-portable.exe'
$repoPortableDir = Join-Path $repoRoot 'release\windows'
$repoPortableExe = Join-Path $repoPortableDir 'eMerge-portable.exe'
$launcherSource = Join-Path $repoRoot 'tools\PortableLauncher.cs'
$iconPath = Join-Path $repoRoot 'src-tauri\icons\icon.ico'
$cscPath = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'

Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not (Test-Path $releaseDir)) {
  throw "未找到 release 目录：$releaseDir"
}
if (-not (Test-Path (Join-Path $releaseDir 'eMerge.exe'))) {
  throw "未找到 release 可执行文件：$(Join-Path $releaseDir 'eMerge.exe')"
}
if (-not (Test-Path (Join-Path $releaseDir 'resources'))) {
  throw "未找到 release resources 目录：$(Join-Path $releaseDir 'resources')"
}
if (-not (Test-Path $cscPath)) {
  throw "未找到 C# 编译器：$cscPath"
}

if (Test-Path $stageRoot) {
  Remove-Item -LiteralPath $stageRoot -Recurse -Force
}
if (Test-Path $outputExe) {
  Remove-Item -LiteralPath $outputExe -Force
}
if (Test-Path $repoPortableExe) {
  Remove-Item -LiteralPath $repoPortableExe -Force
}

New-Item -ItemType Directory -Path $payloadDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
New-Item -ItemType Directory -Path $repoPortableDir -Force | Out-Null

Copy-Item -LiteralPath (Join-Path $releaseDir 'eMerge.exe') -Destination (Join-Path $payloadDir 'eMerge.exe') -Force
Copy-Item -LiteralPath (Join-Path $releaseDir 'resources') -Destination (Join-Path $payloadDir 'resources') -Recurse -Force

[System.IO.Compression.ZipFile]::CreateFromDirectory(
  $payloadDir,
  $payloadZip,
  [System.IO.Compression.CompressionLevel]::Optimal,
  $false
)

$cscArgs = @(
  '/nologo'
  '/target:winexe'
  "/out:$outputExe"
  "/resource:$payloadZip,payload.zip"
  "/win32icon:$iconPath"
  '/reference:System.dll'
  '/reference:System.Core.dll'
  '/reference:System.IO.Compression.dll'
  '/reference:System.IO.Compression.FileSystem.dll'
  '/reference:System.Windows.Forms.dll'
  $launcherSource
)

& $cscPath @cscArgs
if ($LASTEXITCODE -ne 0) {
  throw "portable launcher 编译失败，退出码：$LASTEXITCODE"
}

Copy-Item -LiteralPath $outputExe -Destination $repoPortableExe -Force

Get-Item -LiteralPath $outputExe | Select-Object FullName, Length, LastWriteTime
Get-Item -LiteralPath $repoPortableExe | Select-Object FullName, Length, LastWriteTime
