<#
.SYNOPSIS
    Dispatches one worker task to the DeepSeek API and archives the response.

.DESCRIPTION
    Implements the "Opus -> Worker" leg of the multi-agent loop described in
    Claude_Code_Development_Instructions.md. The orchestrator writes a prompt file, this
    script posts it to DeepSeek, and the raw + extracted responses are archived under
    tools/worker-output/ for the validation phase.

    Endpoint note: DeepSeek exposes an Anthropic-compatible surface at
    https://api.deepseek.com/anthropic/v1/messages (this is what Start-DeepSeek.bat uses via
    ANTHROPIC_BASE_URL) as well as an OpenAI-compatible one at /chat/completions. We use the
    Anthropic surface so the model names match the rest of the toolchain.

    PowerShell 5.1 gotcha handled below: Get-Content -Raw returns a *decorated* string carrying
    PSPath/PSDrive note-properties. Passing it straight to ConvertTo-Json serialises those
    properties as an object, and the API rejects it with
    "content should be a string or a list". Cast to [string] first.

.EXAMPLE
    .\dispatch-worker.ps1 -Worker A -PromptFile .\prompts\worker-A-iter1.md
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateSet('A', 'B', 'C')][string]$Worker,
    [Parameter(Mandatory = $true)][string]$PromptFile,
    [string]$Iteration = '1',
    [string]$Model = 'deepseek-v4-pro[1m]',
    [int]$MaxTokens = 16000,
    [double]$Temperature = 0.2,
    [string]$BaseUrl = 'https://api.deepseek.com/anthropic',
    [string]$KeyFile = 'C:\Users\nguyenhoang88502\Desktop\API key\deepseek api key.txt',
    [string]$OutputDir = ''
)

$ErrorActionPreference = 'Stop'

# $PSScriptRoot is empty when the script is dot-sourced or launched from some OneDrive paths;
# fall back the same way launcher.ps1 does.
$ScriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptRoot) -and $MyInvocation.MyCommand.Path) {
    $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) { $ScriptRoot = (Get-Location).ProviderPath }

if ([string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = Join-Path $ScriptRoot 'worker-output' }
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $PromptFile)) { throw "Prompt file not found: $PromptFile" }
if (-not (Test-Path -LiteralPath $KeyFile)) { throw "DeepSeek key file not found: $KeyFile" }

$apiKey = ([string](Get-Content -LiteralPath $KeyFile -Raw)).Trim()
if ([string]::IsNullOrWhiteSpace($apiKey)) { throw 'DeepSeek key file is empty.' }

# -Encoding UTF8 so em-dashes and other non-ASCII survive; [string] cast strips the
# PSPath/PSDrive note-properties that would otherwise poison ConvertTo-Json.
$prompt = [string](Get-Content -LiteralPath $PromptFile -Raw -Encoding UTF8)

# PowerShell 5.1 defaults to TLS 1.0 on some builds; the API requires TLS 1.2+.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$payload = [ordered]@{
    model       = $Model
    max_tokens  = $MaxTokens
    temperature = $Temperature
    messages    = @(@{ role = 'user'; content = $prompt })
} | ConvertTo-Json -Depth 6 -Compress

$stem = "worker-$Worker-iter$Iteration"
$rawPath = Join-Path $OutputDir "$stem.raw.json"
$textPath = Join-Path $OutputDir "$stem.md"

$headers = @{
    'x-api-key'         = $apiKey
    'Authorization'     = "Bearer $apiKey"
    'anthropic-version' = '2023-06-01'
}

Write-Host "[dispatch] Worker $Worker -> $Model (iteration $Iteration)"
$sw = [Diagnostics.Stopwatch]::StartNew()

try {
    $response = Invoke-RestMethod -Uri "$BaseUrl/v1/messages" `
        -Method POST `
        -Headers $headers `
        -ContentType 'application/json' `
        -Body ([Text.Encoding]::UTF8.GetBytes($payload)) `
        -TimeoutSec 900
}
catch [System.Net.WebException] {
    # Surface the API's own error body; Invoke-RestMethod otherwise swallows it.
    $detail = ''
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        try { $detail = $reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    throw "Worker $Worker dispatch failed: $($_.Exception.Message)`n$detail"
}

$sw.Stop()

$response | ConvertTo-Json -Depth 12 | Out-File -LiteralPath $rawPath -Encoding utf8

# Anthropic message shape: content is an array of blocks; concatenate the text ones.
$text = ($response.content | Where-Object { $_.type -eq 'text' } | ForEach-Object { $_.text }) -join "`n"
$text | Out-File -LiteralPath $textPath -Encoding utf8

Write-Host ("[dispatch] Worker {0} done in {1:N1}s | in={2} tok out={3} tok | stop={4} | {5} chars -> {6}" -f `
        $Worker, $sw.Elapsed.TotalSeconds, $response.usage.input_tokens, $response.usage.output_tokens, `
        $response.stop_reason, $text.Length, $textPath)

if ($response.stop_reason -eq 'max_tokens') {
    Write-Warning "Worker $Worker hit the max_tokens ceiling - output is truncated. Re-dispatch with a higher -MaxTokens or a narrower task."
}
