#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Username = "g1hdmgs",
    [string]$DefaultWeb = "HPC",
    [string]$BaseViewAuthUrl = "https://hpc-wiki.in.audi.vwg/wiki/bin/viewauth",
    [string]$TopicsFile = "topics.txt",
    [string]$OutputDir = "wiki_output",
    [string]$HtmlQueryString = "",
    [int]$RetryCount = 1,
    [switch]$Overwrite,
    [switch]$SkipPlaceholderCheck,
    [bool]$UseAskPassword = $true,
    [bool]$AskPasswordOnce = $true,
    [bool]$ExperimentalInjectAskPassword = $false,
    [string]$Password = "",
    [bool]$ShowServerResponse = $true,
    [bool]$KeepWorkDir = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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
        Remove-Item -LiteralPath $Path -Force -Recurse -ErrorAction SilentlyContinue
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
        WebPath = $webPath
        TopicName = $topicName
    }
}

function Encode-WebPath {
    param([Parameter(Mandatory = $true)][string]$WebPath)
    return (($WebPath -split "/") | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join "/"
}

function Test-IsPlaceholderText {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$SkipCheck
    )

    if ($SkipCheck) {
        return $false
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        return $false
    }

    $maxBytes = 2097152
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

    $text = [System.Text.Encoding]::GetEncoding(28591).GetString($buffer, 0, $read)
    return ($text -like "*WikiGuest*" -and $text -like "*Topic revision: 1970-01-01*")
}

function Shorten-Message {
    param([string]$Text, [int]$MaxLength = 1500)
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $flat = ($Text -replace "\s+", " ").Trim()
    if ($flat.Length -le $MaxLength) {
        return $flat
    }
    return $flat.Substring(0, $MaxLength) + "...(truncated)"
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

function Invoke-WgetCommand {
    param(
        [Parameter(Mandatory = $true)][string]$Executable,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [switch]$Interactive,
        [string]$InjectPassword
    )

    $oldEap = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        if (-not [string]::IsNullOrWhiteSpace($InjectPassword)) {
            # Experimental mode: feed password to wget stdin while using --ask-password.
            $stdin = $InjectPassword + [Environment]::NewLine
            $output = $stdin | & $Executable @Arguments 2>&1
            $exitCode = $LASTEXITCODE
        }
        elseif ($Interactive) {
            # Keep wget attached to console so --ask-password prompt is visible.
            & $Executable @Arguments
            $exitCode = $LASTEXITCODE
            $output = @()
        }
        else {
            $output = & $Executable @Arguments 2>&1
            $exitCode = $LASTEXITCODE
        }
    }
    finally {
        $ErrorActionPreference = $oldEap
    }

    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = $output
    }
}

function Get-PreferredHtmlFile {
    param(
        [Parameter(Mandatory = $true)][string]$SearchRoot,
        [Parameter(Mandatory = $true)][string]$TopicName
    )

    if (-not (Test-Path -LiteralPath $SearchRoot)) {
        return $null
    }

    $htmlFiles = @(
        Get-ChildItem -LiteralPath $SearchRoot -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in @(".html", ".htm") }
    )

    if ($htmlFiles.Count -eq 0) {
        return $null
    }

    $topicRegex = [Regex]::Escape($TopicName)
    $preferred = @($htmlFiles | Where-Object { $_.Name -match $topicRegex -or $_.FullName -match $topicRegex })
    if ($preferred.Count -eq 0) {
        $preferred = $htmlFiles
    }

    return ($preferred | Sort-Object -Property Length -Descending | Select-Object -First 1).FullName
}

$scriptDir = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).ProviderPath } else { $PSScriptRoot }
$topicsPath = Resolve-AbsolutePath -Path $TopicsFile -BasePath $scriptDir
$outputPath = Resolve-AbsolutePath -Path $OutputDir -BasePath $scriptDir
$tmpRoot = Join-Path -Path $outputPath -ChildPath "_tmp"
$downloadLog = Join-Path -Path $outputPath -ChildPath "wiki_download.log"
$failedLog = Join-Path -Path $outputPath -ChildPath "wiki_failed.log"
$cookieJar = Join-Path -Path $outputPath -ChildPath "wiki_session.cookies.txt"
$base = $BaseViewAuthUrl.TrimEnd("/")
$query = if ([string]::IsNullOrWhiteSpace($HtmlQueryString)) { "" } else { "?" + $HtmlQueryString.TrimStart("?") }

if ([string]::IsNullOrWhiteSpace($Username)) {
    throw "Username is empty. Pass -Username explicitly."
}

$wgetExe = Join-Path -Path $scriptDir -ChildPath "wget.exe"
if (-not (Test-Path -LiteralPath $wgetExe)) {
    throw ("wget.exe was not found next to wiki.ps1. Expected: {0}" -f $wgetExe)
}

if (-not (Test-Path -LiteralPath $topicsPath)) {
    throw ("Topics file not found: {0}" -f $topicsPath)
}

if (-not (Test-Path -LiteralPath $outputPath)) {
    New-Item -ItemType Directory -Path $outputPath | Out-Null
}
if (-not (Test-Path -LiteralPath $tmpRoot)) {
    New-Item -ItemType Directory -Path $tmpRoot | Out-Null
}
Remove-IfExists -Path $cookieJar

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
    Write-RunLog -LogFile $downloadLog -Level "WARN" -Topic "-" -Message "No topics found."
    exit 0
}

if (-not $UseAskPassword) {
    if ($ExperimentalInjectAskPassword) {
        Write-Warning "ExperimentalInjectAskPassword is ignored when UseAskPassword is false."
    }
    if ([string]::IsNullOrWhiteSpace($Password)) {
        $securePassword = Read-Host ("Password for user {0}" -f $Username) -AsSecureString
        $Password = Convert-SecureStringToPlainText -SecureString $securePassword
    }
    if ([string]::IsNullOrWhiteSpace($Password)) {
        throw "Empty password entered."
    }
}
elseif ($ExperimentalInjectAskPassword) {
    # Prompt once and inject into each --ask-password call.
    if ([string]::IsNullOrWhiteSpace($Password)) {
        $securePassword = Read-Host ("Password for user {0} (experimental inject mode)" -f $Username) -AsSecureString
        $Password = Convert-SecureStringToPlainText -SecureString $securePassword
    }
    if ([string]::IsNullOrWhiteSpace($Password)) {
        throw "Empty password entered."
    }
}

Write-Host "Starting HTML wiki export (wget mirror-style)..."
Write-Host ("Username         : {0}" -f $Username)
Write-Host ("Default web      : {0}" -f $DefaultWeb)
Write-Host ("Base viewauth    : {0}" -f $base)
Write-Host ("Topics file      : {0}" -f $topicsPath)
Write-Host ("Output dir       : {0}" -f $outputPath)
Write-Host ("wget.exe         : {0}" -f $wgetExe)
Write-Host ("Cookie jar       : {0}" -f $cookieJar)
Write-Host ("RetryCount       : {0}" -f $RetryCount)
Write-Host ("UseAskPassword   : {0}" -f $UseAskPassword)
Write-Host ("AskPasswordOnce  : {0}" -f $AskPasswordOnce)
Write-Host ("InjectAskPass    : {0}" -f $ExperimentalInjectAskPassword)
if ($ExperimentalInjectAskPassword) {
    Write-Host "WARNING: Experimental password injection mode is enabled."
}
Write-Host ("ShowServerResp   : {0}" -f $ShowServerResponse)
Write-Host ("KeepWorkDir      : {0}" -f $KeepWorkDir)
Write-Host ("Download log     : {0}" -f $downloadLog)
Write-Host ("Failed log       : {0}" -f $failedLog)
Write-Host ""

Write-RunLog -LogFile $downloadLog -Level "INFO" -Topic "-" -Message ("RUN_START topics={0}" -f $topics.Count)

$cookieAuthReady = $false
if ($UseAskPassword -and $AskPasswordOnce) {
    # Ask once to establish an authenticated cookie session.
    $warmupTopic = $null
    foreach ($line in $topics) {
        $parsed = Parse-TopicEntry -Entry $line -DefaultWeb $DefaultWeb
        if ($null -ne $parsed) {
            $warmupTopic = $parsed
            break
        }
    }
    if ($null -eq $warmupTopic) {
        $warmupTopic = [pscustomobject]@{
            WebPath = $DefaultWeb
            TopicName = "WebHome"
        }
    }

    $warmupWeb = Encode-WebPath -WebPath $warmupTopic.WebPath
    $warmupTopicName = [System.Uri]::EscapeDataString($warmupTopic.TopicName)
    $warmupUrl = "{0}/{1}/{2}{3}" -f $base, $warmupWeb, $warmupTopicName, $query
    $warmupOut = Join-Path -Path $tmpRoot -ChildPath "_auth_warmup.html"
    Remove-IfExists -Path $warmupOut

    if ($ExperimentalInjectAskPassword) {
        Write-Host "Authentication warm-up (using injected password)..."
    }
    else {
        Write-Host "Authentication warm-up (one password prompt)..."
    }

    $warmupArgs = @(
        "--user=$Username",
        "--ask-password",
        "--content-disposition",
        "--trust-server-names",
        "--max-redirect=10",
        "--keep-session-cookies",
        "--save-cookies=$cookieJar",
        "--load-cookies=$cookieJar",
        "--output-document=$warmupOut"
    )
    if ($ShowServerResponse) {
        $warmupArgs += "--server-response"
    }
    $warmupArgs += $warmupUrl

    $warmupInjectPassword = ""
    if ($ExperimentalInjectAskPassword) {
        $warmupInjectPassword = $Password
    }
    $warmupResult = Invoke-WgetCommand -Executable $wgetExe -Arguments $warmupArgs -Interactive:(-not $ExperimentalInjectAskPassword) -InjectPassword $warmupInjectPassword
    if (($warmupResult.ExitCode -eq 0) -and (Test-Path -LiteralPath $warmupOut) -and ((Get-Item -LiteralPath $warmupOut).Length -gt 0)) {
        $cookieAuthReady = $true
        Write-Host "Authentication warm-up successful; reusing session cookies."
        Write-RunLog -LogFile $downloadLog -Level "INFO" -Topic "-" -Message "Auth warm-up succeeded; cookie reuse enabled." -Url $warmupUrl
    }
    else {
        $msg = ("Auth warm-up failed (exit {0}). Falling back to prompt per topic." -f $warmupResult.ExitCode)
        Write-Warning $msg
        Write-RunLog -LogFile $downloadLog -Level "WARN" -Topic "-" -Message $msg -Url $warmupUrl
    }
    Remove-IfExists -Path $warmupOut
}

$ok = 0
$skipped = 0
$failed = 0

foreach ($entry in $topics) {
    $topic = Parse-TopicEntry -Entry $entry -DefaultWeb $DefaultWeb
    if ($null -eq $topic) {
        Write-Warning ("Skipping invalid topic entry: {0}" -f $entry)
        Write-RunLog -LogFile $downloadLog -Level "WARN" -Topic $entry -Message "Invalid topic line."
        continue
    }

    $safeName = ($entry -replace "[\\/:*?`"<>|]", "_")
    $finalHtmlPath = Join-Path -Path $outputPath -ChildPath ($safeName + ".html")
    $workDir = Join-Path -Path $tmpRoot -ChildPath $safeName

    $encodedWeb = Encode-WebPath -WebPath $topic.WebPath
    $encodedTopic = [System.Uri]::EscapeDataString($topic.TopicName)
    $topicUrl = "{0}/{1}/{2}{3}" -f $base, $encodedWeb, $encodedTopic, $query

    if ((-not $Overwrite) -and (Test-Path -LiteralPath $finalHtmlPath)) {
        $existingSize = (Get-Item -LiteralPath $finalHtmlPath).Length
        if (($existingSize -gt 0) -and (-not (Test-IsPlaceholderText -Path $finalHtmlPath -SkipCheck:$SkipPlaceholderCheck))) {
            Write-Host ("Skipping {0} (already downloaded)" -f $entry)
            Write-RunLog -LogFile $downloadLog -Level "SKIP" -Topic $entry -Message ("Already exists ({0} bytes)." -f $existingSize) -Url $topicUrl
            $skipped++
            continue
        }
        Remove-IfExists -Path $finalHtmlPath
    }

    $done = $false
    $lastError = ""

    for ($attempt = 1; $attempt -le ($RetryCount + 1) -and -not $done; $attempt++) {
        Remove-IfExists -Path $workDir
        New-Item -ItemType Directory -Path $workDir | Out-Null
        Remove-IfExists -Path $finalHtmlPath

        Write-Host ("[{0}/{1}] Downloading HTML for {2} ..." -f $attempt, ($RetryCount + 1), $entry)

        # Mirrors the user-provided successful one-liner style, but in per-topic directories.
        $wgetArgs = @(
            "--user=$Username",
            "--content-disposition",
            "--trust-server-names",
            "--max-redirect=10",
            "--page-requisites",
            "--convert-links",
            "--adjust-extension",
            "--span-hosts",
            "--no-host-directories",
            "--directory-prefix=$workDir"
        )
        $interactivePrompt = $false
        $injectPassword = ""
        if ($UseAskPassword) {
            if ($ExperimentalInjectAskPassword) {
                $wgetArgs += "--ask-password"
                $injectPassword = $Password
            }
            elseif ($AskPasswordOnce -and $cookieAuthReady) {
                # Cookie-auth mode: no repeated password prompts.
                $wgetArgs += "--keep-session-cookies"
                $wgetArgs += "--save-cookies=$cookieJar"
                $wgetArgs += "--load-cookies=$cookieJar"
            }
            else {
                $wgetArgs += "--ask-password"
                $interactivePrompt = $true
            }
        }
        else {
            $wgetArgs += "--password=$Password"
        }
        if ($ShowServerResponse) {
            $wgetArgs += "--server-response"
        }
        $wgetArgs += $topicUrl

        $wgetResult = Invoke-WgetCommand -Executable $wgetExe -Arguments $wgetArgs -Interactive:$interactivePrompt -InjectPassword $injectPassword
        if ($wgetResult.ExitCode -ne 0) {
            $lastError = ("wget exit code {0}: {1}" -f $wgetResult.ExitCode, (Shorten-Message -Text ([string]::Join(" ", $wgetResult.Output))))
        }
        else {
            $htmlCandidate = Get-PreferredHtmlFile -SearchRoot $workDir -TopicName $topic.TopicName
            if ([string]::IsNullOrWhiteSpace($htmlCandidate)) {
                $lastError = "Could not find downloaded HTML file in work directory."
            }
            else {
                Copy-Item -LiteralPath $htmlCandidate -Destination $finalHtmlPath -Force

                if (-not (Test-Path -LiteralPath $finalHtmlPath)) {
                    $lastError = "Final HTML file was not created."
                }
                elseif ((Get-Item -LiteralPath $finalHtmlPath).Length -eq 0) {
                    $lastError = "Final HTML file is empty."
                }
                elseif (Test-IsPlaceholderText -Path $finalHtmlPath -SkipCheck:$SkipPlaceholderCheck) {
                    $lastError = "Final HTML contains guest/placeholder content."
                }
                else {
                    $size = (Get-Item -LiteralPath $finalHtmlPath).Length
                    Write-Host ("Saved HTML: {0}" -f $finalHtmlPath)
                    Write-RunLog -LogFile $downloadLog -Level "OK" -Topic $entry -Message ("Saved ({0} bytes)." -f $size) -Url $topicUrl
                    $done = $true
                    $ok++
                }
            }
        }

        if (-not $done) {
            Write-Warning ("Attempt {0} failed for {1}: {2}" -f $attempt, $entry, $lastError)
            Write-RunLog -LogFile $downloadLog -Level "WARN" -Topic $entry -Message ("Attempt {0}: {1}" -f $attempt, $lastError) -Url $topicUrl
            if ($attempt -lt ($RetryCount + 1)) {
                Start-Sleep -Seconds ([Math]::Min(5, $attempt * 2))
            }
        }
    }

    if (-not $KeepWorkDir) {
        Remove-IfExists -Path $workDir
    }

    if (-not $done) {
        Add-Content -LiteralPath $failedLog -Value ("{0}`t{1}`t{2}`t{3}" -f (Get-Date -Format "s"), $entry, $lastError, $topicUrl)
        Write-RunLog -LogFile $downloadLog -Level "ERROR" -Topic $entry -Message $lastError -Url $topicUrl
        $failed++
    }
}

$Password = ""
if ((Test-Path -LiteralPath $tmpRoot) -and (-not $KeepWorkDir)) {
    $leftovers = @(Get-ChildItem -LiteralPath $tmpRoot -Force -ErrorAction SilentlyContinue)
    if ($leftovers.Count -eq 0) {
        Remove-IfExists -Path $tmpRoot
    }
}

Write-RunLog -LogFile $downloadLog -Level "INFO" -Topic "-" -Message ("RUN_END ok={0} skipped={1} failed={2}" -f $ok, $skipped, $failed)

Write-Host ""
Write-Host "Done."
Write-Host ("Successful: {0}  Skipped: {1}  Failed: {2}" -f $ok, $skipped, $failed)
Write-Host ("Run log   : {0}" -f $downloadLog)
if ($failed -gt 0) {
    Write-Host ("Failed log: {0}" -f $failedLog)
}
