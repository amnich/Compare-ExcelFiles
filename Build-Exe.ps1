#Requires -Version 5.1
<#
.SYNOPSIS
    Compiles Compare-ExcelFiles.ps1 into a standalone executable (.EXE) using the PS2EXE module.

.DESCRIPTION
    Verifies the presence of the ps2exe module (installs from PSGallery if needed) and produces
    a windowed application without console (-NoConsole) in Single-Threaded Apartment (-STA) mode.
    Since Compare-ExcelFiles utilizes built-in .NET standard libraries (System.IO.Compression & System.Xml),
    the resulting executable is completely standalone with zero external dependencies.

.PARAMETER OutputFile
    Path to the output executable (.EXE). Defaults to Compare-ExcelFiles.exe next to this script.

.PARAMETER IconFile
    Path to a .ico icon file (optional).

.EXAMPLE
    .\Build-Exe.ps1
    Compiles Compare-ExcelFiles.exe in the current folder.

.NOTES
    Author: Adam Mnich
    Encoding: UTF-8 with BOM
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputFile,

    [Parameter(Mandatory = $false)]
    [string]$IconFile = "D:\Skrypty\Mnich_Adam_Skrypty\!Helper\Logo_AM6.ico"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

$InputScript = Join-Path $ScriptDir 'Compare-ExcelFiles.ps1'

if (-not $OutputFile) {
    $OutputFile = Join-Path $ScriptDir 'Compare-ExcelFiles.exe'
}

$AppTitle = 'Compare-ExcelFiles - Excel File Comparison Tool'
$AppDesc = 'High-performance Excel and CSV file comparison tool with zero external dependencies'

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " PS2EXE COMPILATION: Compare-ExcelFiles" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# Close any running instance of the output executable to avoid file lock
$targetProcName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
$runningProc = Get-Process -Name $targetProcName -ErrorAction SilentlyContinue
if ($runningProc) {
    Write-Host "Stopping running instance of $targetProcName (PID: $($runningProc.Id))..." -ForegroundColor Yellow
    $runningProc | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 600
}

# 1. Verify PS2EXE module
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "ps2exe module not found. Installing from PSGallery (CurrentUser)..." -ForegroundColor Yellow
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
    }
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
    Write-Host "ps2exe module installed successfully." -ForegroundColor Green
}

Import-Module -Name ps2exe -Force

# 2. Validate source script
if (-not (Test-Path $InputScript)) {
    throw "Source file does not exist: $InputScript"
}

Write-Host "Source file : $InputScript" -ForegroundColor White
Write-Host "Output file : $OutputFile" -ForegroundColor White

# 3. PS2EXE compilation parameters
$CompileParams = @{
    InputFile   = $InputScript
    OutputFile  = $OutputFile
    NoConsole   = $true
    STA         = $true
    Title       = $AppTitle
    Description = $AppDesc
    Company     = 'Adam Mnich'
    Product     = 'Compare-ExcelFiles'
    Copyright   = 'Copyright (c) 2026 Adam Mnich'
    Version     = '1.1.0.0'
    NoError     = $true
    NoOutput    = $true
}

if (-not [string]::IsNullOrWhiteSpace($IconFile) -and (Test-Path $IconFile)) {
    $CompileParams['IconFile'] = $IconFile
    Write-Host "Icon file   : $IconFile" -ForegroundColor White
}

# 4. Invoke compiler
Write-Host "`nStarting compilation..." -ForegroundColor Yellow
Invoke-PS2EXE @CompileParams

if (-not (Test-Path $OutputFile)) {
    throw "Compilation completed, but output file $OutputFile was not found."
}

$item = Get-Item $OutputFile
$sizeKb = [math]::Round($item.Length / 1KB, 1)
Write-Host "`n[SUCCESS] Executable created successfully!" -ForegroundColor Green
Write-Host "Path : $($item.FullName)" -ForegroundColor Green
Write-Host "Size : $sizeKb KB" -ForegroundColor Green
