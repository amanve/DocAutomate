param(
    [string]$NuGetPath
)

$ErrorActionPreference = 'Stop'
$repoRoot = $PSScriptRoot
$solutionPath = Join-Path $repoRoot 'DocAutomate.sln'
$localNuGetPath = Join-Path $repoRoot '.tools\nuget.exe'

if (-not (Test-Path -LiteralPath $solutionPath -PathType Leaf)) {
    throw "Solution not found: $solutionPath"
}

if ($NuGetPath) {
    if (-not (Test-Path -LiteralPath $NuGetPath -PathType Leaf)) {
        throw "NuGet executable not found: $NuGetPath"
    }
    $nugetExe = (Resolve-Path -LiteralPath $NuGetPath).Path
} else {
    $nugetCommand = Get-Command nuget.exe -ErrorAction SilentlyContinue
    if ($nugetCommand) {
        $nugetExe = $nugetCommand.Source
    } elseif (Test-Path -LiteralPath $localNuGetPath -PathType Leaf) {
        $nugetExe = $localNuGetPath
    } else {
        $toolDirectory = Split-Path -Parent $localNuGetPath
        New-Item -ItemType Directory -Path $toolDirectory -Force | Out-Null
        Write-Host 'NuGet CLI not found; downloading it from nuget.org...'
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -OutFile $localNuGetPath
        } catch {
            Remove-Item -LiteralPath $localNuGetPath -ErrorAction SilentlyContinue
            throw "Could not download NuGet CLI. Install nuget.exe or pass -NuGetPath. $($_.Exception.Message)"
        }
        $nugetExe = $localNuGetPath
    }
}

Write-Host "Restoring packages for $solutionPath"
& $nugetExe restore $solutionPath -PackagesDirectory (Join-Path $repoRoot 'packages') -NonInteractive
if ($LASTEXITCODE -ne 0) {
    throw "NuGet restore failed with exit code $LASTEXITCODE"
}

Write-Host 'NuGet packages restored successfully.'
