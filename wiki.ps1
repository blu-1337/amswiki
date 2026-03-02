# ==============================================
# wiki.ps1 - Bulk export Foswiki topics to PDF
# Uses GNU wget authentication flow on Windows
# ==============================================

[CmdletBinding()]
param(
    [string]$Username = $env:USERNAME,                     # Example: g1hdmgs or DOMAIN\g1hdmgs
    [string]$BaseWeb = "System",
    [string[]]$FallbackWebs = @("PPService"),
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

# In PowerShell 7+, prevent native stderr (wget progress) from becoming terminating errors.
$nativeErrPrefVar = Get-Variable -Name "PSNativeCommandUseErrorActionPreference" -ErrorAction SilentlyContinue
if ($null -ne $nativeErrPrefVar) {
    $PSNativeCommandUseErrorActionPreference = $false
}

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

function Parse-TopicEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Entry
    )

    $clean = $Entry.Trim().Trim("/")
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $null
    }

    $webPath = $null
    $topicName = $clean
    if ($clean.Contains("/")) {
        $idx = $clean.LastIndexOf("/")
        $webPath = $clean.Substring(0, $idx).Trim("/")
        $topicName = $clean.Substring($idx + 1).Trim("/")
    }

    if ([string]::IsNullOrWhiteSpace($topicName)) {
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

function Get-WebCandidatesForEntry {
    param(
        [Parameter(Mandatory = $true)]$ParsedEntry,
        [Parameter(Mandatory = $true)][string]$BaseWeb,
        [string[]]$FallbackWebs
    )

    if (-not [string]::IsNullOrWhiteSpace($ParsedEntry.WebPath)) {
        return @($ParsedEntry.WebPath)
    }

    $ordered = @()
    if (-not [string]::IsNullOrWhiteSpace($BaseWeb)) {
        $ordered += $BaseWeb
    }
    if ($null -ne $FallbackWebs) {
        $ordered += $FallbackWebs
    }

    $unique = New-Object System.Collections.Generic.List[string]
    foreach ($w in $ordered) {
        if ($null -eq $w) { continue }
        $clean = $w.Trim().Trim("/")
        if ([string]::IsNullOrWhiteSpace($clean)) { continue }
        if (-not $unique.Contains($clean)) {
            [void]$unique.Add($clean)
        }
    }

    return $unique.ToArray()
}

$scriptDir = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).ProviderPath } else { $PSScriptRoot }
$topicsPath = Resolve-AbsolutePath -Path $TopicsFile -BasePath $scriptDir
$outputPath = Resolve-AbsolutePath -Path $OutputDir -BasePath $scriptDir
$failedLog = Join-Path -Path $outputPath -ChildPath "wiki_failed.log"
$cookieJar = Join-Path -Path $outputPath -ChildPath "wiki_session.cookies.txt"
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
Remove-IfExists -Path $cookieJar

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
if ($FallbackWebs.Count -gt 0) {
    Write-Host ("Fallbacks  : {0}" -f ([string]::Join(", ", $FallbackWebs)))
}
Write-Host ("Topics file: {0}" -f $topicsPath)
Write-Host ("Output dir : {0}" -f $outputPath)
Write-Host ("wget.exe   : {0}" -f $wgetExe)
Write-Host ("Cookie jar : {0}" -f $cookieJar)
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
    $parsed = Parse-TopicEntry -Entry $entry
    if ($null -eq $parsed) {
        Write-Warning ("Skipping invalid topics entry: {0}" -f $entry)
        continue
    }

    $webCandidates = Get-WebCandidatesForEntry -ParsedEntry $parsed -BaseWeb $BaseWeb -FallbackWebs $FallbackWebs
    if (($null -eq $webCandidates) -or ($webCandidates.Count -eq 0)) {
        Write-Warning ("No web candidates resolved for topic: {0}" -f $entry)
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

    $downloaded = $false
    $allErrors = New-Object System.Collections.Generic.List[string]
    $encodedTopic = [System.Uri]::EscapeDataString($parsed.TopicName)

    foreach ($webCandidate in $webCandidates) {
        if ($downloaded) { break }
        $encodedWeb = Encode-WebPath -WebPath $webCandidate
        $topicUrl = "{0}/{1}/{2}{3}" -f $base, $encodedWeb, $encodedTopic, $query
        $candidateError = ""

        for ($attempt = 1; $attempt -le ($RetryCount + 1) -and -not $downloaded; $attempt++) {
            Write-Host ("Downloading {0} [web={1}] ..." -f $entry, $webCandidate)
            Remove-IfExists -Path $tmpPath

            $wgetArgs = @(
                "--user=$Username",
                "--content-disposition",
                "--trust-server-names",
                "--max-redirect=10",
                "--server-response",
                "--auth-no-challenge",
                "--keep-session-cookies",
                "--save-cookies=$cookieJar",
                "--output-document=$tmpPath"
            )

            if (Test-Path -LiteralPath $cookieJar) {
                $wgetArgs += "--load-cookies=$cookieJar"
            }

            if ($AskPasswordPerDownload) {
                $wgetArgs += "--ask-password"
            }
            else {
                $wgetArgs += "--password=$password"
            }

            $wgetArgs += $topicUrl

            $oldErrorActionPreference = $ErrorActionPreference
            try {
                # GNU wget writes progress/status to stderr even on success.
                # Temporarily relax EAP so progress lines do not stop the script.
                $ErrorActionPreference = "Continue"
                $wgetOutput = & $wgetExe @wgetArgs 2>&1
                $exitCode = $LASTEXITCODE
            }
            finally {
                $ErrorActionPreference = $oldErrorActionPreference
            }

            if ($exitCode -ne 0) {
                $wgetMessage = ([string]::Join(" ", $wgetOutput)).Trim()
                if ([string]::IsNullOrWhiteSpace($wgetMessage)) {
                    $candidateError = ("wget exit code {0}" -f $exitCode)
                }
                else {
                    $candidateError = ("wget exit code {0}: {1}" -f $exitCode, $wgetMessage)
                }
            }
            elseif (-not (Test-Path -LiteralPath $tmpPath)) {
                $candidateError = "Output file was not created."
            }
            elseif ((Get-Item -LiteralPath $tmpPath).Length -eq 0) {
                $candidateError = "Downloaded file is empty."
            }
            elseif (-not (Test-PdfSignature -Path $tmpPath)) {
                $candidateError = "Downloaded file is not a valid PDF."
            }
            elseif (Test-IsGuestPlaceholderPdf -Path $tmpPath) {
                $candidateError = "Downloaded PDF contains guest/placeholder content."
            }
            else {
                Move-Item -LiteralPath $tmpPath -Destination $pdfPath -Force
                $downloaded = $true
                $ok++
                if ($webCandidate -ne $BaseWeb) {
                    Write-Host ("Resolved {0} via fallback web {1}" -f $entry, $webCandidate)
                }
            }

            if (-not $downloaded) {
                Remove-IfExists -Path $tmpPath
                if ($attempt -lt ($RetryCount + 1)) {
                    Write-Warning ("Attempt {0} failed for {1} on web {2}. Retrying..." -f $attempt, $entry, $webCandidate)
                    Start-Sleep -Seconds ([Math]::Min(5, $attempt * 2))
                }
            }
        }

        if (-not $downloaded) {
            if ([string]::IsNullOrWhiteSpace($candidateError)) {
                $candidateError = "Unknown error"
            }
            [void]$allErrors.Add(("{0}: {1}" -f $webCandidate, $candidateError))
        }
    }

    if (-not $downloaded) {
        $combinedError = if ($allErrors.Count -gt 0) { [string]::Join(" | ", $allErrors.ToArray()) } else { "Unknown failure" }
        Write-Warning ("Failed to download {0}. Error: {1}" -f $entry, $combinedError)
        Add-Content -LiteralPath $failedLog -Value ("{0}`t{1}`t{2}" -f (Get-Date -Format "s"), $entry, $combinedError)
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
