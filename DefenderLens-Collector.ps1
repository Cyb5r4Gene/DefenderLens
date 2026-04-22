#Requires -RunAsAdministrator
<#
.SYNOPSIS
    DefenderLens - MDE Performance Collector v1.0
    Dynamic endpoint analyzer - exports JSON for DefenderLens web dashboard

.DESCRIPTION
    Interactively prompts for application/process name to analyze.
    Records MDE scan activity, generates report, and exports a structured
    JSON file ready to drag-and-drop into the DefenderLens web dashboard.

.AUTHOR
    DefenderLens / KBN IT Security

.NOTES
    Requirements:
    - Run as Administrator
    - PowerShell 5.1+
    - Windows 10/11 or Server 2016+
    - MDE version 4.18.2201.10+
#>

# ---------------------------------------------
#  SETTINGS
# ---------------------------------------------
$OutputFolder   = "C:\Temp\DefenderLens"
$TopN           = 20
$TopScansExport = 1000

# ---------------------------------------------
#  HELPERS
# ---------------------------------------------
function Write-Header {
    param([string]$Text)
    Write-Host ""
    Write-Host ("  " + ("-" * 58)) -ForegroundColor DarkGray
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ("  " + ("-" * 58)) -ForegroundColor DarkGray
}

function Write-Step { param([string]$T); Write-Host "  >> $T" -ForegroundColor Yellow }
function Write-OK   { param([string]$T); Write-Host "  [OK] $T" -ForegroundColor Green }
function Write-Fail { param([string]$T); Write-Host "  [FAIL] $T" -ForegroundColor Red }
function Write-Info { param([string]$T); Write-Host "  [i] $T" -ForegroundColor Gray }

# ---------------------------------------------
#  INTRO
# ---------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "   DefenderLens - MDE Performance Collector   " -ForegroundColor Cyan
Write-Host "   v1.0  |  KBN IT Security                  " -ForegroundColor Cyan
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------
#  COLLECT SESSION INFO
# ---------------------------------------------
Write-Header "Session Information"

$Hostname  = $env:COMPUTERNAME
$Username  = $env:USERNAME
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

try {
    $OSVersion = (Get-WmiObject Win32_OperatingSystem).Caption
} catch {
    $OSVersion = "Unknown OS"
}

try {
    $MDE = (Get-MpComputerStatus).AMProductVersion
} catch {
    $MDE = "Unknown"
}

Write-Info "Hostname  : $Hostname"
Write-Info "Username  : $Username"
Write-Info "OS        : $OSVersion"
Write-Info "MDE Ver   : $MDE"
Write-Info "Timestamp : $Timestamp"

# ---------------------------------------------
#  USER INPUT - APP / PROCESS NAME
# ---------------------------------------------
Write-Header "Target Application"
Write-Host ""
Write-Host "  Enter the application or process name you want to analyze." -ForegroundColor White
Write-Host "  Examples: Outlook, Bloomberg, Excel, Teams, chrome, WINWORD" -ForegroundColor DarkGray
Write-Host ""

$AppName = Read-Host "  Application name"
if ([string]::IsNullOrWhiteSpace($AppName)) {
    $AppName = "Generic"
}

Write-Host ""
Write-Host "  Optionally enter a description or ticket number (press ENTER to skip):" -ForegroundColor DarkGray
$Description = Read-Host "  Description"
if ([string]::IsNullOrWhiteSpace($Description)) {
    $Description = "MDE performance analysis for $AppName"
}

# ---------------------------------------------
#  RECORDING MODE
# ---------------------------------------------
Write-Header "Recording Mode"
Write-Host ""
Write-Host "  How do you want to record?" -ForegroundColor White
Write-Host "  [1] Manual  - press ENTER to stop when done (recommended)" -ForegroundColor Gray
Write-Host "  [2] Timed   - auto-stop after X seconds" -ForegroundColor Gray
Write-Host ""
$ModeChoice = Read-Host "  Choice (1 or 2)"

$RecordSeconds = $null
if ($ModeChoice -eq "2") {
    $SecInput = Read-Host "  Record for how many seconds? (e.g. 120)"
    if ($SecInput -match "^\d+$") {
        $RecordSeconds = [int]$SecInput
    } else {
        Write-Info "Invalid input, switching to manual mode."
        $ModeChoice = "1"
    }
}

# ---------------------------------------------
#  PREPARE OUTPUT
# ---------------------------------------------
Write-Header "Preparing Output"

$SafeAppName = $AppName -replace "[^a-zA-Z0-9_-]", "_"
$DateStamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$SessionID   = $SafeAppName + "_" + $DateStamp

$ETLFile  = "$OutputFolder\$SessionID.etl"
$JSONFile = "$OutputFolder\$SessionID.json"

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    Write-OK "Created folder: $OutputFolder"
} else {
    Write-OK "Output folder ready: $OutputFolder"
}

# ---------------------------------------------
#  START RECORDING
# ---------------------------------------------
Write-Header "Recording - $AppName"
Write-Host ""
Write-Host "  *** Open $AppName and use it normally now ***" -ForegroundColor Magenta
Write-Host "  *** The more you use it, the richer the data ***" -ForegroundColor Magenta
Write-Host ""

try {
    if ($ModeChoice -eq "2" -and $RecordSeconds) {
        Write-Info "Auto-stop in $RecordSeconds seconds..."
        New-MpPerformanceRecording -RecordTo $ETLFile -Seconds $RecordSeconds
    } else {
        Write-Host "  Press ENTER in this window when done recording." -ForegroundColor Yellow
        Write-Host ""
        New-MpPerformanceRecording -RecordTo $ETLFile
    }
    Write-OK "Recording complete"
} catch {
    Write-Fail "Recording failed: $_"
    exit 1
}

if (-not (Test-Path $ETLFile)) {
    Write-Fail "ETL file not found. Aborting."
    exit 1
}

$ETLSizeMB = [math]::Round((Get-Item $ETLFile).Length / 1MB, 2)
Write-OK "ETL size: $ETLSizeMB MB"

# ---------------------------------------------
#  ANALYZE
# ---------------------------------------------
Write-Header "Analyzing"
Write-Step "Running MDE performance report..."

try {
    $Report = Get-MpPerformanceReport `
        -Path $ETLFile `
        -TopFiles $TopN `
        -TopExtensions $TopN `
        -TopProcesses $TopN `
        -TopScans $TopN

    $TopScansData = (Get-MpPerformanceReport -Path $ETLFile -TopScans $TopScansExport).TopScans
    Write-OK "Analysis complete"
} catch {
    Write-Fail "Analysis failed: $_"
    exit 1
}

# ---------------------------------------------
#  BUILD JSON OUTPUT
# ---------------------------------------------
Write-Header "Building JSON Export"
Write-Step "Serializing results..."

function ConvertTo-SafeList {
    param($InputData)
    if ($null -eq $InputData) { return @() }
    return @($InputData | ForEach-Object {
        $obj = $_
        $hash = @{}
        $obj.PSObject.Properties | ForEach-Object {
            $val = $_.Value
            if ($val -is [TimeSpan]) {
                $hash[$_.Name] = [math]::Round($val.TotalMilliseconds, 2)
            } elseif ($null -eq $val) {
                $hash[$_.Name] = $null
            } else {
                $hash[$_.Name] = "$val"
            }
        }
        $hash
    })
}

$JSONPayload = [ordered]@{
    meta = [ordered]@{
        tool        = "DefenderLens"
        version     = "1.0"
        sessionId   = $SessionID
        appName     = $AppName
        description = $Description
        hostname    = $Hostname
        username    = $Username
        osVersion   = $OSVersion
        mdeVersion  = $MDE
        recordedAt  = $Timestamp
        etlSizeMB   = $ETLSizeMB
        topN        = $TopN
    }
    topFiles      = ConvertTo-SafeList $Report.TopFiles
    topExtensions = ConvertTo-SafeList $Report.TopExtensions
    topProcesses  = ConvertTo-SafeList $Report.TopProcesses
    topScans      = ConvertTo-SafeList $TopScansData
}

try {
    $JSONPayload | ConvertTo-Json -Depth 6 | Out-File -FilePath $JSONFile -Encoding UTF8
    Write-OK "JSON exported: $JSONFile"
} catch {
    Write-Fail "JSON export failed: $_"
    exit 1
}

# ---------------------------------------------
#  QUICK CONSOLE SUMMARY
# ---------------------------------------------
Write-Header "Quick Summary"

Write-Host ""
Write-Host "  TOP PROCESSES" -ForegroundColor Cyan
$Report.TopProcesses | Format-Table -AutoSize | Out-Host

Write-Host "  TOP FILES (first 5)" -ForegroundColor Cyan
$Report.TopFiles | Select-Object -First 5 | Format-Table -AutoSize | Out-Host

Write-Host "  TOP EXTENSIONS" -ForegroundColor Cyan
$Report.TopExtensions | Format-Table -AutoSize | Out-Host

# ---------------------------------------------
#  DONE
# ---------------------------------------------
Write-Header "Complete"
Write-Host ""
Write-OK "Session ID : $SessionID"
Write-OK "JSON File  : $JSONFile"
Write-Host ""
Write-Host "  Next step:" -ForegroundColor White
Write-Host "  Open DefenderLens.html in your browser and drag the JSON file onto it." -ForegroundColor Gray
Write-Host ""
Start-Process explorer.exe $OutputFolder
