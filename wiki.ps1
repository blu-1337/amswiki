#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Username = "g1hdmgs",
    [string]$DefaultWeb = "PPService",
    [string]$BaseUrl = "https://ams-wiki.in.audi.vwg/wiki/bin/genpdf",
    [string]$TopicsFile = "topics.txt",
    [string]$OutputDir = "wiki_output",
    [string]$QueryString = "skin=;",
    [int]$RetryCount = 2,
    [switch]$Overwrite,
    [switch]$SkipPlaceholderCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# PowerShell 7: don't convert native stderr progress to terminating errors.
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

function Remove-IfExists {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
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

function Parse-TopicEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Entry,
        [Parameter(Mandatory = $true)][string]$DefaultWeb
    )

    $clean = $Entry.Trim().Trim("/")
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $null
    }

    $webPath = $DefaultWeb.Trim("/")
    $topicName = $clean

    if ($clean.Contains("/")) {
        $idx = $clean.LastIndexOf("/")
        $candidateWeb = $clean.Substring(0, $idx).Trim("/")
        $candidateTopic = $clean.Substring($idx + 1).Trim("/")
        if (-not [string]::IsNullOrWhiteSpace($candidateWeb) -and -not [string]::IsNullOrWhiteSpace($candidateTopic)) {
            $webPath = $candidateWeb
            $topicName = $candidateTopic
        }
    }

    if ([string]::IsNullOrWhiteSpace($webPath) -or [string]::IsNullOrWhiteSpace($topicName)) {
        return $null
    }

    return [pscustomobject]@{
        WebPath   = $webPath
        TopicName = $topicName
    }
}

function Encode-WebPath {
    param([Parameter(Mandatory = $true)][string]$WebPath)
    return (($WebPath -split "/") | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join "/"
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

function Test-IsPlaceholderPdf {
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

function Write-RunLog {
    param(
        [Parameter(Mandatory = $true)][string]$LogFile,
        [Parameter(Mandatory = $true)][string]$Level,
        [Parameter(Mandatory = $true)][string]$Topic,
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$Url = ""
    )

    $line = "{0}`t{1}`t{2}`t{3}`t{4}" -f (Get-Date -Format "s"), $Level, $Topic, $Message, $Url
    Add-Content -LiteralPath $LogFile -Value $line
}

$scriptDir = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).ProviderPath } else { $PSScriptRoot }
$topicsPath = Resolve-AbsolutePath -Path $TopicsFile -BasePath $scriptDir
$outputPath = Resolve-AbsolutePath -Path $OutputDir -BasePath $scriptDir
$downloadLog = Join-Path -Path $outputPath -ChildPath "wiki_download.log"
$failedLog = Join-Path -Path $outputPath -ChildPath "wiki_failed.log"
$base = $BaseUrl.TrimEnd("/")
$query = if ([string]::IsNullOrWhiteSpace($QueryString)) { "" } else { "?" + $QueryString.TrimStart("?") }

if ([string]::IsNullOrWhiteSpace($Username)) {
    throw "Username is empty. Pass -Username explicitly."
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

if (-not (Test-Path -LiteralPath $downloadLog)) {
    Add-Content -LiteralPath $downloadLog -Value "timestamp`tlevel`ttopic`tmessage`turl"
}

$topics = @(
    Get-Content -LiteralPath $topicsPath |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith("#") }
)

if ($topics.Count -eq 0) {
    Write-Warning ("No topics found in {0}" -f $topicsPath)
    Write-RunLog -LogFile $downloadLog -Level "WARN" -Topic "-" -Message "No topics found in topics file."
    exit 0
}

$securePassword = Read-Host ("Password for user {0}" -f $Username) -AsSecureString
$password = Convert-SecureStringToPlainText -SecureString $securePassword
if ([string]::IsNullOrWhiteSpace($password)) {
    throw "Empty password entered."
}

Write-Host "Starting bulk PDF export using GNU wget..."
Write-Host ("Username   : {0}" -f $Username)
Write-Host ("DefaultWeb : {0}" -f $DefaultWeb)
Write-Host ("Base URL   : {0}" -f $base)
Write-Host ("Topics file: {0}" -f $topicsPath)
Write-Host ("Output dir : {0}" -f $outputPath)
Write-Host ("wget.exe   : {0}" -f $wgetExe)
Write-Host ("Log file   : {0}" -f $downloadLog)
Write-Host ("Failed log : {0}" -f $failedLog)
Write-Host ""

Write-RunLog -LogFile $downloadLog -Level "INFO" -Topic "-" -Message ("RUN_START topics={0}" -f $topics.Count)

$ok = 0
$skipped = 0
$failed = 0

foreach ($entry in $topics) {
    $topicParts = Parse-TopicEntry -Entry $entry -DefaultWeb $DefaultWeb
    if ($null -eq $topicParts) {
        Write-Warning ("Skipping invalid topics entry: {0}" -f $entry)
        Write-RunLog -LogFile $downloadLog -Level "WARN" -Topic $entry -Message "Invalid topic line."
        continue
    }

    $safeFileName = ($entry -replace "[\\/:*?`"<>|]", "_")
    $pdfPath = Join-Path -Path $outputPath -ChildPath ($safeFileName + ".pdf")
    $tmpPath = $pdfPath + ".download"

    $encodedWeb = Encode-WebPath -WebPath $topicParts.WebPath
    $encodedTopic = [System.Uri]::EscapeDataString($topicParts.TopicName)
    $topicUrl = "{0}/{1}/{2}{3}" -f $base, $encodedWeb, $encodedTopic, $query

    if ((-not $Overwrite) -and (Test-Path -LiteralPath $pdfPath)) {
        $existingSize = (Get-Item -LiteralPath $pdfPath).Length
        if (($existingSize -gt 0) -and (Test-PdfSignature -Path $pdfPath) -and (-not (Test-IsPlaceholderPdf -Path $pdfPath))) {
            Write-Host ("Skipping {0} (already downloaded)" -f $entry)
            Write-RunLog -LogFile $downloadLog -Level "SKIP" -Topic $entry -Message ("Already exists ({0} bytes)." -f $existingSize) -Url $topicUrl
            $skipped++
            continue
        }

        Remove-IfExists -Path $pdfPath
    }

    $downloaded = $false
    $lastError = ""

    for ($attempt = 1; $attempt -le ($RetryCount + 1) -and -not $downloaded; $attempt++) {
        Write-Host ("Downloading {0} ..." -f $entry)
        Remove-IfExists -Path $tmpPath

        # Mirrors your working command pattern:
        # wget --user=... --password=... --content-disposition --trust-server-names --max-redirect=10 --server-response -O file URL
        $wgetArgs = @(
            "--user=$Username",
            "--password=$password",
            "--content-disposition",
            "--trust-server-names",
            "--max-redirect=10",
            "--server-response",
            "--output-document=$tmpPath",
            $topicUrl
        )

        $oldEap = $ErrorActionPreference
        try {
            $ErrorActionPreference = "Continue"
            $wgetOutput = & $wgetExe @wgetArgs 2>&1
            $exitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $oldEap
        }

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
        elseif (Test-IsPlaceholderPdf -Path $tmpPath) {
            $lastError = "Downloaded PDF contains placeholder/guest content."
        }
        else {
            Move-Item -LiteralPath $tmpPath -Destination $pdfPath -Force
            $size = (Get-Item -LiteralPath $pdfPath).Length
            Write-Host ("Saved PDF: {0}" -f $pdfPath)
            Write-RunLog -LogFile $downloadLog -Level "OK" -Topic $entry -Message ("Saved ({0} bytes)." -f $size) -Url $topicUrl
            $downloaded = $true
            $ok++
        }

        if (-not $downloaded) {
            Remove-IfExists -Path $tmpPath
            if ($attempt -lt ($RetryCount + 1)) {
                Write-Warning ("Attempt {0} failed for {1}. Retrying..." -f $attempt, $entry)
                Write-RunLog -LogFile $downloadLog -Level "WARN" -Topic $entry -Message ("Attempt {0} failed: {1}" -f $attempt, $lastError) -Url $topicUrl
                Start-Sleep -Seconds ([Math]::Min(5, $attempt * 2))
            }
        }
    }

    if (-not $downloaded) {
        Write-Warning ("Failed to download {0}. Error: {1}" -f $entry, $lastError)
        Write-RunLog -LogFile $downloadLog -Level "ERROR" -Topic $entry -Message $lastError -Url $topicUrl
        Add-Content -LiteralPath $failedLog -Value ("{0}`t{1}`t{2}`t{3}" -f (Get-Date -Format "s"), $entry, $lastError, $topicUrl)
        $failed++
    }
}

$password = $null

Write-RunLog -LogFile $downloadLog -Level "INFO" -Topic "-" -Message ("RUN_END ok={0} skipped={1} failed={2}" -f $ok, $skipped, $failed)

Write-Host ""
Write-Host "Done."
Write-Host ("Successful: {0}  Skipped: {1}  Failed: {2}" -f $ok, $skipped, $failed)
Write-Host ("Run log    : {0}" -f $downloadLog)
if ($failed -gt 0) {
    Write-Host ("Failed log : {0}" -f $failedLog)
}
