[CmdletBinding()]
param(
    [string]$OutputDirectory = '',
    [switch]$SkipTests,
    [switch]$SmokeTest
)
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path $PSScriptRoot -Parent
if (!$OutputDirectory) { $OutputDirectory = Join-Path $repositoryRoot 'artifacts/win-x64' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
Push-Location $repositoryRoot
try {
    if (!$SkipTests) {
        dotnet test native/PdfMailer.Tests/PdfMailer.Tests.csproj -c Release --nologo
        if ($LASTEXITCODE -ne 0) { throw 'C# regression tests failed.' }
    }
    dotnet publish native/PdfMailer.App/PdfMailer.App.csproj -c Release -r win-x64 -p:Platform=x64 --self-contained true -p:DebugType=None -p:DebugSymbols=false -o $OutputDirectory --nologo
    if ($LASTEXITCODE -ne 0) { throw 'WinUI publish failed.' }
    Copy-Item -LiteralPath config.example.json,使用说明.md -Destination $OutputDirectory -Force
    $examplesOutput = Join-Path $OutputDirectory 'examples'
    New-Item -ItemType Directory -Force $examplesOutput | Out-Null
    Copy-Item -LiteralPath examples/email_template.txt -Destination $examplesOutput -Force
    if ($SmokeTest) {
        $qaDirectory = Join-Path $repositoryRoot ('artifacts/smoke-' + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Force $qaDirectory | Out-Null
        $appProcess = Start-Process -FilePath (Join-Path $OutputDirectory 'PdfMailer.exe') -ArgumentList @('--smoke-test',('"' + $qaDirectory + '"')) -WindowStyle Hidden -PassThru
        if (!$appProcess.WaitForExit(60000)) { $appProcess.Kill(); throw 'WinUI smoke test timed out.' }
        if ($appProcess.ExitCode -ne 0) { throw "WinUI exited with code $($appProcess.ExitCode)." }
        $reportPath = Join-Path $qaDirectory 'smoke-report.json'
        if (!(Test-Path -LiteralPath $reportPath)) { throw 'WinUI did not write a smoke report.' }
        $report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
        if (!$report.success) { throw ('WinUI smoke failed: ' + $report.error) }
        Write-Output "Native UI smoke passed: $qaDirectory"
    }
    Write-Output "Published app: $OutputDirectory"
} finally { Pop-Location }
