#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$BaseURL = "https://ams-wiki.in.audi.vwg/wiki/bin/genpdf",
    [string]$TopicsFile = "topics.txt",       # Each line: TopicName OR Web/TopicName
    [string]$OutputDir = "wiki_output",
    [string]$DefaultWeb = "PPService",        # Applied when line has only topic name
    [string]$QueryString = "",                # Example: "skin=genpdf,pattern"
    [int]$RetryCount = 2,
    [switch]$Overwrite,
    [switch]$KeepDebugResponses
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$BasePath
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path -Path $BasePath -ChildPath $Path))
}

function Get-SafeFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Topic
    )

    return ($Topic -replace "[\\/:*?`"<>|]", "_")
}

function Normalize-TopicPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Topic,
        [string]$DefaultWeb
    )

    $cleanTopic = $Topic.Trim("/")
    if ($cleanTopic -match "/") {
        return $cleanTopic
    }

    if (-not [string]::IsNullOrWhiteSpace($DefaultWeb)) {
        return "$($DefaultWeb.Trim('/'))/$cleanTopic"
    }

    return $cleanTopic
}

function Convert-TopicToUrlPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TopicPath
    )

    $parts = $TopicPath.Trim("/") -split "/"
    return (($parts | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join "/")
}

function Get-UriCandidates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Base,
        [Parameter(Mandatory = $true)]
        [string]$UrlPath,
        [string]$QueryString
    )

    $seen = @{}
    $list = New-Object System.Collections.Generic.List[string]

    function Add-UriCandidate {
        param([string]$Uri)
        if (-not $seen.ContainsKey($Uri)) {
            $seen[$Uri] = $true
            [void]$list.Add($Uri)
        }
    }

    $cleanQuery = $null
    if (-not [string]::IsNullOrWhiteSpace($QueryString)) {
        $cleanQuery = $QueryString.Trim().TrimStart("?")
    }

    Add-UriCandidate -Uri "$Base/$UrlPath"
    if (-not [string]::IsNullOrWhiteSpace($cleanQuery)) {
        Add-UriCandidate -Uri "$Base/$UrlPath?$cleanQuery"
    }

    # Useful fallback for many Foswiki GenPDF setups.
    if (($Base -match "/genpdf/?$") -and ([string]::IsNullOrWhiteSpace($cleanQuery) -or ($cleanQuery -notmatch "(^|&)skin="))) {
        Add-UriCandidate -Uri "$Base/$UrlPath?skin=genpdf,pattern"
    }

    return $list
}

function Test-PdfSignature {
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

    return ([System.Text.Encoding]::ASCII.GetString($bytes, 0, 5) -eq "%PDF-")
}

function Get-FileSnippet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [int]$ByteLimit = 262144
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return ""
    }

    $limit = [Math]::Max(1024, $ByteLimit)
    $buffer = New-Object byte[] $limit
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $read = $stream.Read($buffer, 0, $buffer.Length)
    }
    finally {
        $stream.Dispose()
    }

    if ($read -le 0) {
        return ""
    }

    return [System.Text.Encoding]::GetEncoding(28591).GetString($buffer, 0, $read)
}

function Get-HeaderValueFromDump {
    param(
        [string]$HeaderFile,
        [Parameter(Mandatory = $true)]
        [string]$HeaderName
    )

    if ([string]::IsNullOrWhiteSpace($HeaderFile) -or -not (Test-Path -LiteralPath $HeaderFile)) {
        return $null
    }

    $pattern = "^\s*" + [Regex]::Escape($HeaderName) + "\s*:\s*(.+)\s*$"
    $matches = @()
    foreach ($line in (Get-Content -LiteralPath $HeaderFile -ErrorAction SilentlyContinue)) {
        if ($line -imatch $pattern) {
            $matches += $Matches[1]
        }
    }

    if ($matches.Count -gt 0) {
        return $matches[-1]
    }

    return $null
}

function Get-PdfValidationError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string]$ResponseContentType
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return "Output file was not created."
    }

    $size = (Get-Item -LiteralPath $Path).Length
    if ($size -le 0) {
        return "Downloaded file is empty."
    }

    if (-not [string]::IsNullOrWhiteSpace($ResponseContentType)) {
        if ($ResponseContentType -notmatch "pdf|octet-stream") {
            return "Unexpected content type: $ResponseContentType"
        }
    }

    if (-not (Test-PdfSignature -Path $Path)) {
        return "Downloaded file is not a PDF."
    }

    # Detect common guest/fallback document symptoms.
    $snippet = Get-FileSnippet -Path $Path -ByteLimit 1048576
    $markers = @(
        "Topic revision: 1970-01-01",
        "WikiGuest",
        "RENDERZONE{",
        "This topic:"
    )
    $hits = @($markers | Where-Object { $snippet -like "*$_*" })
    if ($hits.Count -ge 2) {
        return "PDF content looks like guest/placeholder topic output."
    }

    return $null
}

function Download-WithCurl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $true)]
        [string]$OutFile,
        [ref]$HeaderDumpFile,
        [ref]$ErrorMessage
    )

    $headerFile = [System.IO.Path]::GetTempFileName()
    $HeaderDumpFile.Value = $headerFile

    $curlArgs = @(
        "--negotiate",
        "-u", ":",
        "--location",
        "--silent",
        "--show-error",
        "--fail",
        "--connect-timeout", "20",
        "--output", $OutFile,
        "--dump-header", $headerFile,
        $Uri
    )

    $output = & $script:CurlExe @curlArgs 2>&1
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        $ErrorMessage.Value = "curl exit $exitCode: $([string]::Join(' ', $output))"
        return $false
    }

    return $true
}

function Download-WithInvokeWebRequest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $true)]
        [string]$OutFile,
        [ref]$HeaderDumpFile,
        [ref]$ErrorMessage
    )

    try {
        $response = Invoke-WebRequest `
            -Uri $Uri `
            -OutFile $OutFile `
            -UseDefaultCredentials `
            -UseBasicParsing `
            -MaximumRedirection 10 `
            -Headers @{ Accept = "application/pdf,application/octet-stream;q=0.9,*/*;q=0.8" } `
            -ErrorAction Stop

        $headerFile = [System.IO.Path]::GetTempFileName()
        $lines = @()
        if ($null -ne $response.Headers) {
            foreach ($key in $response.Headers.AllKeys) {
                $lines += "$key: $($response.Headers[$key])"
            }
        }
        Set-Content -LiteralPath $headerFile -Value $lines -Encoding ASCII
        $HeaderDumpFile.Value = $headerFile
        return $true
    }
    catch {
        $ErrorMessage.Value = $_.Exception.Message
        return $false
    }
}

function Remove-FileIfExists {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

$scriptBase = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { (Get-Location).ProviderPath }
$TopicsFile = Resolve-AbsolutePath -Path $TopicsFile -BasePath $scriptBase
$OutputDir = Resolve-AbsolutePath -Path $OutputDir -BasePath $scriptBase

if (-not (Test-Path -LiteralPath $TopicsFile)) {
    throw "Topics file not found: $TopicsFile"
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$curlCommand = Get-Command curl.exe -ErrorAction SilentlyContinue
$script:CurlExe = if ($null -ne $curlCommand) { $curlCommand.Source } else { $null }
$downloadMethod = if ($null -ne $script:CurlExe) { "curl.exe --negotiate" } else { "Invoke-WebRequest (fallback)" }

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
Write-Host "Script dir       : $scriptBase"
Write-Host "Base URL         : $base"
Write-Host "Topics file      : $TopicsFile"
Write-Host "Output dir       : $OutputDir"
Write-Host "Default web      : $DefaultWeb"
Write-Host "Request method   : $downloadMethod"
Write-Host "Query string     : $QueryString"
Write-Host ""

foreach ($topic in $topics) {
    $topicPath = Normalize-TopicPath -Topic $topic -DefaultWeb $DefaultWeb
    $safeName = Get-SafeFileName -Topic $topic
    $outFile = Join-Path -Path $OutputDir -ChildPath "$safeName.pdf"
    $tmpFile = "$outFile.download"

    if ((-not $Overwrite) -and (Test-Path -LiteralPath $outFile)) {
        $existingValidation = Get-PdfValidationError -Path $outFile -ResponseContentType "application/pdf"
        if ($null -eq $existingValidation) {
            Write-Host "Skipping $topic (already downloaded)"
            $skipped++
            continue
        }

        Write-Warning "Existing file for $topic is invalid ($existingValidation). Re-downloading..."
        Remove-FileIfExists -Path $outFile
    }

    $urlPath = Convert-TopicToUrlPath -TopicPath $topicPath
    $uriCandidates = Get-UriCandidates -Base $base -UrlPath $urlPath -QueryString $QueryString

    $downloaded = $false
    $lastFailure = ""

    for ($attempt = 1; $attempt -le ($RetryCount + 1) -and -not $downloaded; $attempt++) {
        foreach ($uri in $uriCandidates) {
            Remove-FileIfExists -Path $tmpFile
            $headerDump = $null
            $requestError = $null

            Write-Host "Downloading $topic (attempt $attempt) from $uri"

            $okRequest = $false
            if ($null -ne $script:CurlExe) {
                $okRequest = Download-WithCurl -Uri $uri -OutFile $tmpFile -HeaderDumpFile ([ref]$headerDump) -ErrorMessage ([ref]$requestError)
            }
            else {
                $okRequest = Download-WithInvokeWebRequest -Uri $uri -OutFile $tmpFile -HeaderDumpFile ([ref]$headerDump) -ErrorMessage ([ref]$requestError)
            }

            if (-not $okRequest) {
                $lastFailure = "Request failed: $requestError"
                if (-not [string]::IsNullOrWhiteSpace($headerDump) -and (Test-Path -LiteralPath $headerDump)) {
                    Remove-FileIfExists -Path $headerDump
                }
                continue
            }

            $contentType = Get-HeaderValueFromDump -HeaderFile $headerDump -HeaderName "Content-Type"
            $validationError = Get-PdfValidationError -Path $tmpFile -ResponseContentType $contentType
            if ($null -eq $validationError) {
                Move-Item -LiteralPath $tmpFile -Destination $outFile -Force
                if (-not [string]::IsNullOrWhiteSpace($headerDump) -and (Test-Path -LiteralPath $headerDump)) {
                    Remove-FileIfExists -Path $headerDump
                }
                $downloaded = $true
                $ok++
                break
            }

            $lastFailure = "Validation failed: $validationError (Content-Type: $contentType)"
            if ($KeepDebugResponses) {
                $debugPath = Join-Path -Path $OutputDir -ChildPath "$safeName.debug.txt"
                $snippet = Get-FileSnippet -Path $tmpFile -ByteLimit 32768
                Set-Content -LiteralPath $debugPath -Value @(
                    "URI: $uri",
                    "Validation: $validationError",
                    "Content-Type: $contentType",
                    "",
                    "Snippet:",
                    $snippet
                ) -Encoding UTF8
            }

            Remove-FileIfExists -Path $tmpFile
            if (-not [string]::IsNullOrWhiteSpace($headerDump) -and (Test-Path -LiteralPath $headerDump)) {
                Remove-FileIfExists -Path $headerDump
            }
        }

        if (-not $downloaded -and $attempt -lt ($RetryCount + 1)) {
            Write-Warning "Attempt $attempt failed for $topic. Retrying..."
            Start-Sleep -Seconds ([Math]::Min(5, $attempt * 2))
        }
    }

    if (-not $downloaded) {
        Write-Warning "Failed to download $topic. $lastFailure"
        Add-Content -LiteralPath $failedLog -Value ("{0}`t{1}`t{2}" -f (Get-Date -Format "s"), $topic, $lastFailure)
        $failed++
    }
}

Write-Host ""
Write-Host "Done."
Write-Host ("Successful: {0}  Skipped: {1}  Failed: {2}" -f $ok, $skipped, $failed)
if ($failed -gt 0) {
    Write-Host "Failed topics logged to: $failedLog"
}
