# ==============================================
# wiki.ps1 - Bulk export Foswiki topics to PDF
# Works with PowerShell 5+ and 7+ using curl negotiate auth
# ==============================================

[CmdletBinding()]
param(
    [string]$BaseWeb = "System",
    [string]$BaseURL = "https://ams-wiki.in.audi.vwg/wiki/bin/genpdf",
    [string]$TopicsFile = "topics.txt",
    [string]$OutputDir = "wiki_output",
    [string]$QueryString = "skin=;",
    [int]$RetryCount = 2,
    [switch]$Overwrite,
    [switch]$SkipContentValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$BasePath
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path -Path $BasePath -ChildPath $Path))
}

function Test-IsPdf {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $false
    }

    $bytes = New-Object byte[] 5
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $read = $stream.Read($bytes, 0, $bytes.Length)
    }
    finally {
        $stream.Dispose()
    }

    if ($read -lt 5) {
        return $false
    }

    return ([System.Text.Encoding]::ASCII.GetString($bytes, 0, 5) -eq "%PDF-")
}

function Test-IsGuestPlaceholderPdf {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ($SkipContentValidation) {
        return $false
    }

    $maxBytes = 1048576
    $buffer = New-Object byte[] $maxBytes
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $read = $stream.Read($buffer, 0, $buffer.Length)
    }
    finally {
        $stream.Dispose()
    }

    if ($read -le 0) {
        return $false
    }

    $snippet = [System.Text.Encoding]::GetEncoding(28591).GetString($buffer, 0, $read)
    $markers = @(
        "Topic revision: 1970-01-01",
        "WikiGuest",
        "RENDERZONE{",
        "This topic:"
    )

    $hitCount = @($markers | Where-Object { $snippet -like ("*" + $_ + "*") }).Count
    return ($hitCount -ge 2)
}

function Remove-IfExists {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

$scriptDir = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).ProviderPath } else { $PSScriptRoot }
$topicsPath = Resolve-AbsolutePath -Path $TopicsFile -BasePath $scriptDir
$outputPath = Resolve-AbsolutePath -Path $OutputDir -BasePath $scriptDir
$failedLog = Join-Path -Path $outputPath -ChildPath "wiki_failed.log"
$base = $BaseURL.TrimEnd("/")
$query = if ([string]::IsNullOrWhiteSpace($QueryString)) { "" } else { "?" + $QueryString.TrimStart("?") }

$curlCmd = Get-Command curl.exe -ErrorAction SilentlyContinue
if ($null -eq $curlCmd) {
    throw "curl.exe was not found. Please use Windows with curl.exe available."
}
$curlExe = $curlCmd.Source

if (-not (Test-Path -LiteralPath $topicsPath)) {
    throw ("Topics file not found: {0}" -f $topicsPath)
}

if (-not (Test-Path -LiteralPath $outputPath)) {
    New-Item -ItemType Directory -Path $outputPath | Out-Null
}

$topics = Get-Content -LiteralPath $topicsPath |
    ForEach-Object { $_.Trim() } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith("#") }

Write-Host "Starting bulk PDF export using Windows session authentication..."
Write-Host ("Base URL   : {0}" -f $base)
Write-Host ("Base Web   : {0}" -f $BaseWeb)
Write-Host ("Topics file: {0}" -f $topicsPath)
Write-Host ("Output dir : {0}" -f $outputPath)
Write-Host ""

$ok = 0
$skipped = 0
$failed = 0

foreach ($topic in $topics) {
    if ([string]::IsNullOrWhiteSpace($topic)) { continue }

    $topicPath = if ($topic -match "/") { $topic.Trim("/") } else { "{0}/{1}" -f $BaseWeb.Trim("/"), $topic.Trim("/") }
    $encodedPath = (($topicPath -split "/") | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join "/"
    $topicUrl = "{0}/{1}{2}" -f $base, $encodedPath, $query

    $safeFileName = $topic -replace "[\\/:*?`"<>|]", "_"
    $pdfPath = Join-Path -Path $outputPath -ChildPath ($safeFileName + ".pdf")
    $tmpPath = $pdfPath + ".download"

    if ((-not $Overwrite) -and (Test-Path -LiteralPath $pdfPath)) {
        $existingSize = (Get-Item -LiteralPath $pdfPath).Length
        if (($existingSize -gt 0) -and (Test-IsPdf -Path $pdfPath) -and (-not (Test-IsGuestPlaceholderPdf -Path $pdfPath))) {
            Write-Host ("Skipping {0} (already downloaded)" -f $topic)
            $skipped++
            continue
        }
        Remove-IfExists -Path $pdfPath
    }

    $downloaded = $false
    $lastError = ""

    for ($attempt = 1; $attempt -le ($RetryCount + 1) -and -not $downloaded; $attempt++) {
        Write-Host ("Downloading {0} ..." -f $topic)
        Remove-IfExists -Path $tmpPath

        $curlArgs = @(
            "--negotiate",
            "-u", ":",
            "--location",
            "--silent",
            "--show-error",
            "--fail",
            "--output", $tmpPath,
            $topicUrl
        )

        $curlOutput = & $curlExe @curlArgs 2>&1
        $exitCode = $LASTEXITCODE

        if ($exitCode -ne 0) {
            $curlMessage = ([string]::Join(" ", $curlOutput)).Trim()
            if ([string]::IsNullOrWhiteSpace($curlMessage)) {
                $lastError = ("curl exit code {0}" -f $exitCode)
            }
            else {
                $lastError = ("curl exit code {0}: {1}" -f $exitCode, $curlMessage)
            }
        }
        elseif (-not (Test-Path -LiteralPath $tmpPath)) {
            $lastError = "Downloaded file was not created."
        }
        elseif ((Get-Item -LiteralPath $tmpPath).Length -eq 0) {
            $lastError = "Downloaded file is empty."
        }
        elseif (-not (Test-IsPdf -Path $tmpPath)) {
            $lastError = "Downloaded file is not a PDF."
        }
        elseif (Test-IsGuestPlaceholderPdf -Path $tmpPath) {
            $lastError = "Downloaded PDF contains guest/placeholder content."
        }
        else {
            Move-Item -LiteralPath $tmpPath -Destination $pdfPath -Force
            $downloaded = $true
            $ok++
        }

        if (-not $downloaded) {
            Remove-IfExists -Path $tmpPath
            if ($attempt -lt ($RetryCount + 1)) {
                Write-Warning ("Attempt {0} failed for {1}. Retrying..." -f $attempt, $topic)
                Start-Sleep -Seconds ([Math]::Min(5, $attempt * 2))
            }
        }
    }

    if (-not $downloaded) {
        Write-Warning ("Failed to download {0}. Error: {1}" -f $topic, $lastError)
        Add-Content -LiteralPath $failedLog -Value ("{0}`t{1}`t{2}`t{3}" -f (Get-Date -Format "s"), $topic, $lastError, $topicUrl)
        $failed++
    }
}

Write-Host ""
Write-Host "Done."
Write-Host ("Successful: {0}  Skipped: {1}  Failed: {2}" -f $ok, $skipped, $failed)
if ($failed -gt 0) {
    Write-Host ("Failed topics logged to: {0}" -f $failedLog)
}
