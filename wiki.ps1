#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Username = "g1hdmgs",
    [string]$DefaultWeb = "PPService",
    [string]$BaseViewAuthUrl = "https://ams-wiki.in.audi.vwg/wiki/bin/viewauth",
    [string]$TopicsFile = "topics.txt",
    [string]$OutputDir = "wiki_output",
    [string]$HtmlQueryString = "skin=plain;template=viewplain",
    [string]$WkhtmltopdfPath = "wkhtmltox/bin/wkhtmltopdf.exe",
    [int]$RetryCount = 1,
    [ValidateRange(1, 16)]
    [int]$MaxParallel = 1,
    [switch]$Overwrite,
    [switch]$KeepHtml,
    [switch]$SkipPlaceholderCheck,
    [switch]$SkipAssetMirror,
    [switch]$ShowServerResponse,
    [switch]$NoSessionWarmup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# PowerShell 7: native stderr can become non-terminating errors depending on preference.
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

function Test-IsPlaceholderText {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ($SkipPlaceholderCheck) {
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

function Test-IsPlaceholderPdf {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ($SkipPlaceholderCheck) {
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

function Get-PreferredHtmlForConversion {
    param(
        [Parameter(Mandatory = $true)][string]$SearchRoot,
        [Parameter(Mandatory = $true)][string]$TopicName,
        [Parameter(Mandatory = $true)][string]$FallbackHtmlPath
    )

    if (-not (Test-Path -LiteralPath $SearchRoot)) {
        return $FallbackHtmlPath
    }

    $htmlFiles = @(
        Get-ChildItem -LiteralPath $SearchRoot -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in @(".html", ".htm") }
    )

    if ($htmlFiles.Count -eq 0) {
        return $FallbackHtmlPath
    }

    $topicRegex = [Regex]::Escape($TopicName)
    $preferred = @(
        $htmlFiles |
            Where-Object { $_.Name -match $topicRegex -or $_.FullName -match $topicRegex }
    )

    if ($preferred.Count -eq 0) {
        $preferred = $htmlFiles
    }

    return ($preferred | Sort-Object -Property Length -Descending | Select-Object -First 1).FullName
}

function Shorten-Message {
    param(
        [string]$Text,
        [int]$MaxLength = 1500
    )

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
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    $oldEap = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $output = & $Executable @Arguments 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $oldEap
    }

    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = $output
    }
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

$wkhtmlExe = Resolve-AbsolutePath -Path $WkhtmltopdfPath -BasePath $scriptDir
if (-not (Test-Path -LiteralPath $wkhtmlExe)) {
    throw ("wkhtmltopdf.exe was not found. Expected: {0}" -f $wkhtmlExe)
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

$securePassword = Read-Host ("Password for user {0}" -f $Username) -AsSecureString
$password = Convert-SecureStringToPlainText -SecureString $securePassword
if ([string]::IsNullOrWhiteSpace($password)) {
    throw "Empty password entered."
}

Write-Host "Starting HTML -> PDF wiki export..."
Write-Host ("Username      : {0}" -f $Username)
Write-Host ("Default web   : {0}" -f $DefaultWeb)
Write-Host ("Base viewauth : {0}" -f $base)
Write-Host ("Topics file   : {0}" -f $topicsPath)
Write-Host ("Output dir    : {0}" -f $outputPath)
Write-Host ("wget.exe      : {0}" -f $wgetExe)
Write-Host ("wkhtmltopdf   : {0}" -f $wkhtmlExe)
Write-Host ("Cookie jar    : {0}" -f $cookieJar)
Write-Host ("RetryCount    : {0}" -f $RetryCount)
Write-Host ("MaxParallel   : {0}" -f $MaxParallel)
if ($SkipAssetMirror) {
    Write-Host "Asset mirror  : OFF (faster)"
}
else {
    Write-Host "Asset mirror  : ON"
}
if ($NoSessionWarmup) {
    Write-Host "Session warmup: OFF"
}
else {
    Write-Host "Session warmup: ON"
}
Write-Host ("Download log  : {0}" -f $downloadLog)
Write-Host ("Failed log    : {0}" -f $failedLog)
Write-Host ""

Write-RunLog -LogFile $downloadLog -Level "INFO" -Topic "-" -Message ("RUN_START topics={0}" -f $topics.Count)

$warmupTopicUrl = "{0}/{1}/WebHome{2}" -f $base, (Encode-WebPath -WebPath $DefaultWeb), $query

if (-not $NoSessionWarmup) {
    $warmupHtml = Join-Path -Path $tmpRoot -ChildPath "_warmup.html"
    Remove-IfExists -Path $warmupHtml

    $warmupArgs = @(
        "--user=$Username",
        "--password=$password",
        "--max-redirect=10",
        "--auth-no-challenge",
        "--keep-session-cookies",
        "--save-cookies=$cookieJar",
        "--output-document=$warmupHtml"
    )
    if (Test-Path -LiteralPath $cookieJar) {
        $warmupArgs += "--load-cookies=$cookieJar"
    }
    if ($ShowServerResponse) {
        $warmupArgs += "--server-response"
    }
    $warmupArgs += $warmupTopicUrl

    $warmupResult = Invoke-WgetCommand -Executable $wgetExe -Arguments $warmupArgs
    if ($warmupResult.ExitCode -eq 0) {
        Write-RunLog -LogFile $downloadLog -Level "INFO" -Topic "-" -Message "Session warmup succeeded." -Url $warmupTopicUrl
    }
    else {
        $warmMsg = ("Warmup exit code {0}: {1}" -f $warmupResult.ExitCode, (Shorten-Message -Text ([string]::Join(" ", $warmupResult.Output))))
        Write-Warning $warmMsg
        Write-RunLog -LogFile $downloadLog -Level "WARN" -Topic "-" -Message $warmMsg -Url $warmupTopicUrl
    }
    Remove-IfExists -Path $warmupHtml
}

$ok = 0
$skipped = 0
$failed = 0
$pendingConversions = New-Object System.Collections.Generic.List[object]

foreach ($entry in $topics) {
    $topic = Parse-TopicEntry -Entry $entry -DefaultWeb $DefaultWeb
    if ($null -eq $topic) {
        Write-Warning ("Skipping invalid topic entry: {0}" -f $entry)
        Write-RunLog -LogFile $downloadLog -Level "WARN" -Topic $entry -Message "Invalid topic line."
        continue
    }

    $safeName = ($entry -replace "[\\/:*?`"<>|]", "_")
    $pdfPath = Join-Path -Path $outputPath -ChildPath ($safeName + ".pdf")
    $workDir = Join-Path -Path $tmpRoot -ChildPath $safeName
    $htmlPath = Join-Path -Path $workDir -ChildPath "page.html"
    $assetsRoot = Join-Path -Path $workDir -ChildPath "assets"

    $encodedWeb = Encode-WebPath -WebPath $topic.WebPath
    $encodedTopic = [System.Uri]::EscapeDataString($topic.TopicName)
    $topicUrl = "{0}/{1}/{2}{3}" -f $base, $encodedWeb, $encodedTopic, $query

    if ((-not $Overwrite) -and (Test-Path -LiteralPath $pdfPath)) {
        $size = (Get-Item -LiteralPath $pdfPath).Length
        if (($size -gt 0) -and (Test-PdfSignature -Path $pdfPath) -and (-not (Test-IsPlaceholderPdf -Path $pdfPath))) {
            Write-Host ("Skipping {0} (already downloaded)" -f $entry)
            Write-RunLog -LogFile $downloadLog -Level "SKIP" -Topic $entry -Message ("Already exists ({0} bytes)." -f $size) -Url $topicUrl
            $skipped++
            continue
        }
        Remove-IfExists -Path $pdfPath
    }

    $done = $false
    $queuedForParallel = $false
    $lastError = ""

    for ($attempt = 1; $attempt -le ($RetryCount + 1) -and -not $done; $attempt++) {
        Remove-IfExists -Path $workDir
        New-Item -ItemType Directory -Path $workDir | Out-Null
        Remove-IfExists -Path $pdfPath

        Write-Host ("[{0}/{1}] Downloading HTML for {2} ..." -f $attempt, ($RetryCount + 1), $entry)

        $wgetArgs = @(
            "--user=$Username",
            "--password=$password",
            "--content-disposition",
            "--trust-server-names",
            "--max-redirect=10",
            "--auth-no-challenge",
            "--keep-session-cookies",
            "--save-cookies=$cookieJar",
            "--output-document=$htmlPath",
            $topicUrl
        )
        if (Test-Path -LiteralPath $cookieJar) {
            $wgetArgs += "--load-cookies=$cookieJar"
        }
        if ($ShowServerResponse) {
            $wgetArgs += "--server-response"
        }

        $wgetResult = Invoke-WgetCommand -Executable $wgetExe -Arguments $wgetArgs
        $wgetOutput = $wgetResult.Output
        $wgetExit = $wgetResult.ExitCode

        if ($wgetExit -ne 0) {
            $lastError = ("wget exit code {0}: {1}" -f $wgetExit, (Shorten-Message -Text ([string]::Join(" ", $wgetOutput))))
        }
        elseif (-not (Test-Path -LiteralPath $htmlPath)) {
            $lastError = "HTML file was not created."
        }
        elseif ((Get-Item -LiteralPath $htmlPath).Length -eq 0) {
            $lastError = "HTML file is empty."
        }
        elseif (Test-IsPlaceholderText -Path $htmlPath) {
            $lastError = "Downloaded HTML appears to be guest/placeholder content."
        }
        else {
            $htmlForPdf = $htmlPath
            if (-not $SkipAssetMirror) {
                # Second pass: fetch page requisites and rewrite links for reliable local rendering.
                $assetArgs = @(
                    "--user=$Username",
                    "--password=$password",
                    "--content-disposition",
                    "--trust-server-names",
                    "--max-redirect=10",
                    "--auth-no-challenge",
                    "--keep-session-cookies",
                    "--save-cookies=$cookieJar",
                    "--page-requisites",
                    "--convert-links",
                    "--adjust-extension",
                    "--no-host-directories",
                    "--directory-prefix=$assetsRoot",
                    $topicUrl
                )
                if (Test-Path -LiteralPath $cookieJar) {
                    $assetArgs += "--load-cookies=$cookieJar"
                }
                if ($ShowServerResponse) {
                    $assetArgs += "--server-response"
                }

                $assetResult = Invoke-WgetCommand -Executable $wgetExe -Arguments $assetArgs
                if ($assetResult.ExitCode -ne 0) {
                    $assetError = ("Asset fetch warning (exit {0}): {1}" -f $assetResult.ExitCode, (Shorten-Message -Text ([string]::Join(" ", $assetResult.Output))))
                    Write-RunLog -LogFile $downloadLog -Level "WARN" -Topic $entry -Message $assetError -Url $topicUrl
                }

                $htmlForPdf = Get-PreferredHtmlForConversion -SearchRoot $assetsRoot -TopicName $topic.TopicName -FallbackHtmlPath $htmlPath
            }

            if ($MaxParallel -gt 1) {
                [void]$pendingConversions.Add([pscustomobject]@{
                    Topic = $entry
                    Url = $topicUrl
                    HtmlPath = $htmlForPdf
                    PdfPath = $pdfPath
                    WorkDir = $workDir
                    SafeName = $safeName
                })
                $done = $true
                $queuedForParallel = $true
                Write-Host ("Queued conversion for {0}" -f $entry)
                Write-RunLog -LogFile $downloadLog -Level "INFO" -Topic $entry -Message "Queued for parallel PDF conversion." -Url $topicUrl
            }
            else {
                Write-Host ("[{0}/{1}] Converting HTML to PDF for {2} ..." -f $attempt, ($RetryCount + 1), $entry)

                $wkArgs = @(
                    "--enable-local-file-access",
                    "--load-error-handling", "ignore",
                    "--load-media-error-handling", "ignore",
                    $htmlForPdf,
                    $pdfPath
                )

                $oldEap2 = $ErrorActionPreference
                try {
                    $ErrorActionPreference = "Continue"
                    $wkOutput = & $wkhtmlExe @wkArgs 2>&1
                    $wkExit = $LASTEXITCODE
                }
                finally {
                    $ErrorActionPreference = $oldEap2
                }

                if ($wkExit -ne 0) {
                    $lastError = ("wkhtmltopdf exit code {0}: {1}" -f $wkExit, (Shorten-Message -Text ([string]::Join(" ", $wkOutput))))
                }
                elseif (-not (Test-Path -LiteralPath $pdfPath)) {
                    $lastError = "PDF file was not created."
                }
                elseif ((Get-Item -LiteralPath $pdfPath).Length -eq 0) {
                    $lastError = "PDF file is empty."
                }
                elseif (-not (Test-PdfSignature -Path $pdfPath)) {
                    $lastError = "Generated file is not a valid PDF."
                }
                elseif (Test-IsPlaceholderPdf -Path $pdfPath) {
                    $lastError = "Generated PDF contains guest/placeholder content."
                }
                else {
                    $done = $true
                    $size = (Get-Item -LiteralPath $pdfPath).Length
                    Write-Host ("Saved PDF: {0}" -f $pdfPath)
                    Write-RunLog -LogFile $downloadLog -Level "OK" -Topic $entry -Message ("Saved ({0} bytes)." -f $size) -Url $topicUrl
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

    if ((-not $KeepHtml) -and (-not $queuedForParallel)) {
        Remove-IfExists -Path $workDir
    }

    if (-not $done) {
        Add-Content -LiteralPath $failedLog -Value ("{0}`t{1}`t{2}`t{3}" -f (Get-Date -Format "s"), $entry, $lastError, $topicUrl)
        Write-RunLog -LogFile $downloadLog -Level "ERROR" -Topic $entry -Message $lastError -Url $topicUrl
        $failed++
    }
}

if (($MaxParallel -gt 1) -and ($pendingConversions.Count -gt 0)) {
    Write-Host ""
    Write-Host ("Starting parallel PDF conversion ({0} workers, {1} items)..." -f $MaxParallel, $pendingConversions.Count)
    Write-RunLog -LogFile $downloadLog -Level "INFO" -Topic "-" -Message ("PARALLEL_CONVERT_START workers={0} items={1}" -f $MaxParallel, $pendingConversions.Count)

    $running = @()
    $nextIndex = 0
    $total = $pendingConversions.Count

    while (($nextIndex -lt $total) -or ($running.Count -gt 0)) {
        while (($running.Count -lt $MaxParallel) -and ($nextIndex -lt $total)) {
            $task = $pendingConversions[$nextIndex]
            $taskIdx = $nextIndex + 1
            $nextIndex++

            $jobBase = "{0:D5}_{1}" -f $taskIdx, ($task.SafeName -replace "[^A-Za-z0-9_.-]", "_")
            $stdoutLog = Join-Path -Path $tmpRoot -ChildPath ("wkhtml_" + $jobBase + ".out.log")
            $stderrLog = Join-Path -Path $tmpRoot -ChildPath ("wkhtml_" + $jobBase + ".err.log")
            Remove-IfExists -Path $stdoutLog
            Remove-IfExists -Path $stderrLog

            $wkArgs = @(
                "--enable-local-file-access",
                "--load-error-handling", "ignore",
                "--load-media-error-handling", "ignore",
                $task.HtmlPath,
                $task.PdfPath
            )

            $proc = Start-Process -FilePath $wkhtmlExe -ArgumentList $wkArgs -NoNewWindow -PassThru -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog
            $running += [pscustomobject]@{
                Process = $proc
                Task = $task
                StdOut = $stdoutLog
                StdErr = $stderrLog
            }
        }

        $finished = @($running | Where-Object { $_.Process.HasExited })
        if ($finished.Count -eq 0) {
            Start-Sleep -Milliseconds 200
            continue
        }

        foreach ($job in $finished) {
            $err = ""
            $processOutput = ""
            if (Test-Path -LiteralPath $job.StdOut) {
                $processOutput += [System.IO.File]::ReadAllText($job.StdOut)
            }
            if (Test-Path -LiteralPath $job.StdErr) {
                $processOutput += " " + [System.IO.File]::ReadAllText($job.StdErr)
            }

            if ($job.Process.ExitCode -ne 0) {
                $err = ("wkhtmltopdf exit code {0}: {1}" -f $job.Process.ExitCode, (Shorten-Message -Text $processOutput))
            }
            elseif (-not (Test-Path -LiteralPath $job.Task.PdfPath)) {
                $err = "PDF file was not created."
            }
            elseif ((Get-Item -LiteralPath $job.Task.PdfPath).Length -eq 0) {
                $err = "PDF file is empty."
            }
            elseif (-not (Test-PdfSignature -Path $job.Task.PdfPath)) {
                $err = "Generated file is not a valid PDF."
            }
            elseif (Test-IsPlaceholderPdf -Path $job.Task.PdfPath) {
                $err = "Generated PDF contains guest/placeholder content."
            }

            if ([string]::IsNullOrWhiteSpace($err)) {
                $size = (Get-Item -LiteralPath $job.Task.PdfPath).Length
                Write-Host ("Saved PDF: {0}" -f $job.Task.PdfPath)
                Write-RunLog -LogFile $downloadLog -Level "OK" -Topic $job.Task.Topic -Message ("Saved ({0} bytes)." -f $size) -Url $job.Task.Url
                $ok++
            }
            else {
                Write-Warning ("Failed to convert {0}. Error: {1}" -f $job.Task.Topic, $err)
                Write-RunLog -LogFile $downloadLog -Level "ERROR" -Topic $job.Task.Topic -Message $err -Url $job.Task.Url
                Add-Content -LiteralPath $failedLog -Value ("{0}`t{1}`t{2}`t{3}" -f (Get-Date -Format "s"), $job.Task.Topic, $err, $job.Task.Url)
                $failed++
            }

            if (-not $KeepHtml) {
                Remove-IfExists -Path $job.Task.WorkDir
            }
            Remove-IfExists -Path $job.StdOut
            Remove-IfExists -Path $job.StdErr
        }

        $finishedIds = @($finished | ForEach-Object { $_.Process.Id })
        $running = @($running | Where-Object { $finishedIds -notcontains $_.Process.Id })
    }

    Write-RunLog -LogFile $downloadLog -Level "INFO" -Topic "-" -Message "PARALLEL_CONVERT_END"
}

$password = $null
if ((Test-Path -LiteralPath $tmpRoot) -and (-not $KeepHtml)) {
    # Clean temp root if empty after successful cleanup.
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
