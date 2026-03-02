# ==============================================
# wiki.ps1 - Bulk export Foswiki topics to PDF
# Uses GNU wget authentication flow on Windows
# ==============================================

[CmdletBinding()]
param(
    [string]$Username = $env:USERNAME,                     # Example: g1hdmgs or DOMAIN\g1hdmgs
    [string]$BaseWeb = "System",
    [string]$BaseURL = "https://ams-wiki.in.audi.vwg/wiki/bin/genpdf",
    [string]$TopicsFile = "topics.txt",
    [string]$OutputDir = "wiki_output",
    [string]$QueryString = "skin=;",
    [int]$RetryCount = 2,
    [switch]$Overwrite,
    [switch]$SkipPlaceholderCheck,
    [switch]$AskPasswordPerDownload
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

function Convert-SecureStringToPlainText {
    param([Parameter(Mandatory = $true)][Security.SecureString]$SecureString)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Remove-IfExists {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

function Test-PdfSignature {
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

    if ($SkipPlaceholderCheck) {
        return $false
    }

    if (-not (Test-Path -LiteralPath $Path)) {
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

    $content = [System.Text.Encoding]::GetEncoding(28591).GetString($buffer, 0, $read)
    $markers = @(
        "Topic revision: 1970-01-01",
        "WikiGuest",
        "RENDERZONE{",
        "This topic:"
    )

    $hitCount = @($markers | Where-Object { $content -like ("*" + $_ + "*") }).Count
    return ($hitCount -ge 2)
}

function Split-TopicEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Entry,
        [Parameter(Mandatory = $true)][string]$DefaultWeb
    )

    $clean = $Entry.Trim().Trim("/")
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $null
    }

    if ($clean.Contains("/")) {
        $idx = $clean.LastIndexOf("/")
        $webPath = $clean.Substring(0, $idx).Trim("/")
        $topicName = $clean.Substring($idx + 1).Trim("/")
    }
    else {
        $webPath = $DefaultWeb.Trim("/")
        $topicName = $clean
    }

    if ([string]::IsNullOrWhiteSpace($webPath) -or [string]::IsNullOrWhiteSpace($topicName)) {
        return $null
    }

    return [pscustomobject]@{
        WebPath = $webPath
        TopicName = $topicName
    }
}

function Encode-WebPath {
    param([Parameter(Mandatory = $true)][string]$WebPath)
    return (($WebPath -split "/") | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join "/"
}

$scriptDir = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).ProviderPath } else { $PSScriptRoot }
$topicsPath = Resolve-AbsolutePath -Path $TopicsFile -BasePath $scriptDir
$outputPath = Resolve-AbsolutePath -Path $OutputDir -BasePath $scriptDir
$failedLog = Join-Path -Path $outputPath -ChildPath "wiki_failed.log"
$base = $BaseURL.TrimEnd("/")
$query = if ([string]::IsNullOrWhiteSpace($QueryString)) { "" } else { "?" + $QueryString.TrimStart("?") }

if ([string]::IsNullOrWhiteSpace($Username)) {
    throw "Username is empty. Pass -Username explicitly (example: -Username g1hdmgs)."
}

$wgetExe = Join-Path -Path $scriptDir -ChildPath "wget.exe"
if (-not (Test-Path -LiteralPath $wgetExe)) {
    throw ("wget.exe was not found next to wiki.ps1. Expected location: {0}" -f $wgetExe)
}

if (-not (Test-Path -LiteralPath $topicsPath)) {
    throw ("Topics file not found: {0}" -f $topicsPath)
}

if (-not (Test-Path -LiteralPath $outputPath)) {
    New-Item -ItemType Directory -Path $outputPath | Out-Null
}

$topics = Get-Content -LiteralPath $topicsPath |
    ForEach-Object { $_.Trim() } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith("#") }

if ($topics.Count -eq 0) {
    Write-Warning ("No topics found in {0}" -f $topicsPath)
    exit 0
}

$password = $null
if (-not $AskPasswordPerDownload) {
    $securePassword = Read-Host ("Password for {0}" -f $Username) -AsSecureString
    $password = Convert-SecureStringToPlainText -SecureString $securePassword
    if ([string]::IsNullOrWhiteSpace($password)) {
        throw "Empty password entered."
    }
}

Write-Host "Starting bulk PDF export using GNU wget..."
Write-Host ("Username   : {0}" -f $Username)
Write-Host ("Base URL   : {0}" -f $base)
Write-Host ("Base Web   : {0}" -f $BaseWeb)
Write-Host ("Topics file: {0}" -f $topicsPath)
Write-Host ("Output dir : {0}" -f $outputPath)
Write-Host ("wget.exe   : {0}" -f $wgetExe)
if ($AskPasswordPerDownload) {
    Write-Host "Auth mode  : wget --ask-password for each download"
}
else {
    Write-Host "Auth mode  : single prompt, reused for all downloads"
}
Write-Host ""

$ok = 0
$skipped = 0
$failed = 0

foreach ($entry in $topics) {
    $parts = Split-TopicEntry -Entry $entry -DefaultWeb $BaseWeb
    if ($null -eq $parts) {
        Write-Warning ("Skipping invalid topics entry: {0}" -f $entry)
        continue
    }

    $safeFileName = ($entry -replace "[\\/:*?`"<>|]", "_")
    $pdfPath = Join-Path -Path $outputPath -ChildPath ($safeFileName + ".pdf")
    $tmpPath = $pdfPath + ".download"

    if ((-not $Overwrite) -and (Test-Path -LiteralPath $pdfPath)) {
        $size = (Get-Item -LiteralPath $pdfPath).Length
        if (($size -gt 0) -and (Test-PdfSignature -Path $pdfPath) -and (-not (Test-IsGuestPlaceholderPdf -Path $pdfPath))) {
            Write-Host ("Skipping {0} (already downloaded)" -f $entry)
            $skipped++
            continue
        }

        Remove-IfExists -Path $pdfPath
    }

    $encodedWeb = Encode-WebPath -WebPath $parts.WebPath
    $encodedTopic = [System.Uri]::EscapeDataString($parts.TopicName)
    $topicUrl = "{0}/{1}/{2}{3}" -f $base, $encodedWeb, $encodedTopic, $query

    $downloaded = $false
    $lastError = ""

    for ($attempt = 1; $attempt -le ($RetryCount + 1) -and -not $downloaded; $attempt++) {
        Write-Host ("Downloading {0} ..." -f $entry)
        Remove-IfExists -Path $tmpPath

        $wgetArgs = @(
            "--user=$Username",
            "--content-disposition",
            "--trust-server-names",
            "--max-redirect=10",
            "--server-response",
            "--output-document=$tmpPath"
        )

        if ($AskPasswordPerDownload) {
            $wgetArgs += "--ask-password"
        }
        else {
            $wgetArgs += "--password=$password"
        }

        $wgetArgs += $topicUrl

        $wgetOutput = & $wgetExe @wgetArgs 2>&1
        $exitCode = $LASTEXITCODE

        if ($exitCode -ne 0) {
            $wgetMessage = ([string]::Join(" ", $wgetOutput)).Trim()
            if ([string]::IsNullOrWhiteSpace($wgetMessage)) {
                $lastError = ("wget exit code {0}" -f $exitCode)
            }
            else {
                $lastError = ("wget exit code {0}: {1}" -f $exitCode, $wgetMessage)
            }
        }
        elseif (-not (Test-Path -LiteralPath $tmpPath)) {
            $lastError = "Output file was not created."
        }
        elseif ((Get-Item -LiteralPath $tmpPath).Length -eq 0) {
            $lastError = "Downloaded file is empty."
        }
        elseif (-not (Test-PdfSignature -Path $tmpPath)) {
            $lastError = "Downloaded file is not a valid PDF."
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
                Write-Warning ("Attempt {0} failed for {1}. Retrying..." -f $attempt, $entry)
                Start-Sleep -Seconds ([Math]::Min(5, $attempt * 2))
            }
        }
    }

    if (-not $downloaded) {
        Write-Warning ("Failed to download {0}. Error: {1}" -f $entry, $lastError)
        Add-Content -LiteralPath $failedLog -Value ("{0}`t{1}`t{2}`t{3}" -f (Get-Date -Format "s"), $entry, $lastError, $topicUrl)
        $failed++
    }
}

$password = $null

Write-Host ""
Write-Host "Done."
Write-Host ("Successful: {0}  Skipped: {1}  Failed: {2}" -f $ok, $skipped, $failed)
if ($failed -gt 0) {
    Write-Host ("Failed topics logged to: {0}" -f $failedLog)
}
