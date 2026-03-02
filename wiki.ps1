#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$BaseURL = "https://ams-wiki.in.audi.vwg/wiki/bin/genpdf",
    [string]$TopicsFile = "topics.txt",
    [string]$OutputDir = "wiki_output",
    [string]$QueryString = "skin=;",
    [string]$DefaultWeb = "PPService",
    [int]$RetryCount = 2,
    [switch]$Overwrite
)

$extractorPath = Join-Path -Path $PSScriptRoot -ChildPath "wiki-extractor.ps1"
if (-not (Test-Path -LiteralPath $extractorPath)) {
    throw "Extractor script not found: $extractorPath"
}

& $extractorPath @PSBoundParameters
