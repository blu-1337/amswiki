#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$BaseURL = "https://ams-wiki.in.audi.vwg/wiki/bin/genpdf",
    [string]$TopicsFile = "topics.txt",      # Each line: Web/TopicName
    [string]$OutputDir = "wiki_output",
    [string]$QueryString = "skin=;",
    [int]$RetryCount = 2,
    [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-SafeFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Topic
    )

    # Replace folder separators and invalid Windows filename characters.
    return ($Topic -replace "[\\/:*?`"<>|]", "_")
}

function Convert-TopicToUrlPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Topic
    )

    $parts = $Topic.Trim("/") -split "/"
    return (($parts | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join "/")
}

function Test-IsPdfFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

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

    $header = [System.Text.Encoding]::ASCII.GetString($bytes, 0, 5)
    return $header -eq "%PDF-"
}

if (-not (Test-Path -LiteralPath $TopicsFile)) {
    throw "Topics file not found: $TopicsFile"
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$failedLog = Join-Path -Path $OutputDir -ChildPath "wiki_failed.log"
$topics = Get-Content -LiteralPath $TopicsFile |
    ForEach-Object { $_.Trim() } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith("#") }

if ($topics.Count -eq 0) {
    Write-Warning "No topics found in $TopicsFile"
    exit 0
}

$ok = 0
$skipped = 0
$failed = 0
$base = $BaseURL.TrimEnd("/")

Write-Host "Starting bulk PDF export..."
Write-Host "Base URL   : $base"
Write-Host "Topics file: $TopicsFile"
Write-Host "Output dir : $OutputDir"
Write-Host ""

foreach ($topic in $topics) {
    $safeName = Get-SafeFileName -Topic $topic
    $outFile = Join-Path -Path $OutputDir -ChildPath "$safeName.pdf"

    if ((-not $Overwrite) -and (Test-Path -LiteralPath $outFile)) {
        $existingSize = (Get-Item -LiteralPath $outFile).Length
        if ($existingSize -gt 0) {
            Write-Host "Skipping $topic (already downloaded)"
            $skipped++
            continue
        }

        Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    }

    $urlPath = Convert-TopicToUrlPath -Topic $topic
    $uri = "$base/$urlPath?$QueryString"

    Write-Host "Downloading $topic ..."

    $downloaded = $false
    for ($attempt = 1; $attempt -le ($RetryCount + 1); $attempt++) {
        try {
            # "wget" in Windows PowerShell is an alias for Invoke-WebRequest.
            Invoke-WebRequest `
                -Uri $uri `
                -OutFile $outFile `
                -UseDefaultCredentials `
                -MaximumRedirection 5 `
                -ErrorAction Stop | Out-Null

            if (-not (Test-Path -LiteralPath $outFile)) {
                throw "Output file was not created."
            }

            $length = (Get-Item -LiteralPath $outFile).Length
            if ($length -eq 0) {
                throw "Downloaded file is empty."
            }

            if (-not (Test-IsPdfFile -Path $outFile)) {
                throw "Server response is not a valid PDF (likely login/HTML page)."
            }

            $downloaded = $true
            $ok++
            break
        }
        catch {
            if (Test-Path -LiteralPath $outFile) {
                Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
            }

            if ($attempt -lt ($RetryCount + 1)) {
                Write-Warning "Attempt $attempt failed for $topic. Retrying..."
                Start-Sleep -Seconds ([Math]::Min(5, $attempt * 2))
            }
            else {
                $msg = $_.Exception.Message
                Write-Warning "Failed to download ${topic}. Error: $msg"
                Add-Content -LiteralPath $failedLog -Value ("{0}`t{1}`t{2}" -f (Get-Date -Format "s"), $topic, $msg)
                $failed++
            }
        }
    }

    if (-not $downloaded) {
        continue
    }
}

Write-Host ""
Write-Host "Done."
Write-Host ("Successful: {0}  Skipped: {1}  Failed: {2}" -f $ok, $skipped, $failed)
if ($failed -gt 0) {
    Write-Host "Failed topics logged to: $failedLog"
}
