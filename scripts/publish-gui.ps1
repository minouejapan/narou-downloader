[CmdletBinding()]
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\artifacts')
)

$ErrorActionPreference = 'Stop'

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$outputRoot = [System.IO.Path]::GetFullPath($OutputDirectory)
$repoPrefix = $repoRoot.TrimEnd(
    [System.IO.Path]::DirectorySeparatorChar,
    [System.IO.Path]::AltDirectorySeparatorChar
) + [System.IO.Path]::DirectorySeparatorChar

if (-not $outputRoot.StartsWith($repoPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "OutputDirectory must be inside the repository: $repoRoot"
}

$projectPath = Join-Path $repoRoot 'NarouDownloaderGui\NarouDownloaderGui.csproj'
$publishDirectory = Join-Path $outputRoot 'NarouDownloaderGui-win-x64'
$archivePath = Join-Path $outputRoot 'NarouDownloaderGui-win-x64.zip'

New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null

if (Test-Path -LiteralPath $publishDirectory) {
    Remove-Item -LiteralPath $publishDirectory -Recurse -Force
}

if (Test-Path -LiteralPath $archivePath) {
    Remove-Item -LiteralPath $archivePath -Force
}

$publishArguments = @(
    'publish'
    $projectPath
    '-c'
    'Release'
    '-p:Platform=x64'
    '-r'
    'win-x64'
    '--self-contained'
    'true'
    '-p:WindowsPackageType=None'
    '-p:WindowsAppSDKSelfContained=true'
    '-p:WindowsAppSdkUndockedRegFreeWinRTInitialize=true'
    '-p:EnableMsixTooling=true'
    '-p:PublishSingleFile=true'
    '-p:IncludeAllContentForSelfExtract=true'
    '-p:DebugType=None'
    '-p:DebugSymbols=false'
    '-o'
    $publishDirectory
)

& dotnet @publishArguments
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE."
}

$executablePath = Join-Path $publishDirectory 'NarouDownloaderGui.exe'
if (-not (Test-Path -LiteralPath $executablePath)) {
    throw "Published executable was not found: $executablePath"
}

Compress-Archive -Path (Join-Path $publishDirectory '*') -DestinationPath $archivePath -CompressionLevel Optimal

Write-Host "Portable GUI created:"
Write-Host "  $archivePath"
