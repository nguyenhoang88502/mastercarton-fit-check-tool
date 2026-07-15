$ErrorActionPreference = 'Stop'
$tool = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifestsOnly = $args[0] -eq '--manifests-only'
$logPath = Join-Path $tool 'WVN_Master_Carton_Fit_Tool_launcher.log'

function Write-Log {
  param([string]$msg)
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $line = "$ts $msg"
  Write-Host $line
  Add-Content -LiteralPath $logPath -Value $line
}

# Clear old log
if (Test-Path -LiteralPath $logPath) { Remove-Item -LiteralPath $logPath -Force }

try {
  Write-Log "=== WVN Master Carton Fit Tool Launcher ==="
  Write-Log "TOOL_DIR = $tool"

  # Username from tool path or fallback
  $username = $null
  if ($tool -match 'C:\\Users\\([^\\]+)') { $username = $matches[1] }
  if (-not $username) { $username = [Environment]::UserName }
  $userRoot = Join-Path 'C:\Users' $username
  Write-Log "USERNAME = $username"
  Write-Log "USER_ROOT = $userRoot"

  # Dynamic OneDrive detection
  $oneDriveRoot = $null
  $odBase = Join-Path $userRoot 'OneDrive*'
  $oneDriveCandidates = @(Get-ChildItem -Path $odBase -Directory -ErrorAction SilentlyContinue | Sort-Object Name)
  Write-Log "OneDrive candidates: $($oneDriveCandidates.Count)"
  foreach ($od in $oneDriveCandidates) {
    $npiPath = Join-Path $od.FullName 'NPI Information - Documents'
    if (Test-Path -LiteralPath $npiPath) {
      $oneDriveRoot = $od.FullName
      Write-Log "OneDrive selected: $oneDriveRoot"
      break
    }
  }
  if (-not $oneDriveRoot) {
    if ($oneDriveCandidates.Count -gt 0) {
      $oneDriveRoot = $oneDriveCandidates[0].FullName
      Write-Log "WARNING: No OneDrive folder with NPI Information - Documents found. Using first OneDrive: $oneDriveRoot"
    } else {
      Write-Log "WARNING: No OneDrive folder found at all."
    }
  }

  # Build NPI paths
  $npiToolsDir = $null
  $weeklyBomDir = $null
  if ($oneDriveRoot) {
    $npiToolsDir = Join-Path $oneDriveRoot 'NPI Information - Documents\General\Project Management\5. Tools for NPI'
    $weeklyBomDir = Join-Path $npiToolsDir '1. Weekly BOM'
  }
  $localBomDir = Join-Path $tool '1. Weekly BOM'
  $parentBomDir = Join-Path (Split-Path -Parent $tool) '1. Weekly BOM'
  Write-Log "NPI_TOOLS_DIR = $npiToolsDir"
  Write-Log "WEEKLY_BOM_DIR = $weeklyBomDir"

  # Search directories for BOM (priority: OneDrive → sibling → local)
  $bomSearchDirs = @()
  if ($weeklyBomDir -and (Test-Path -LiteralPath $weeklyBomDir)) { $bomSearchDirs += $weeklyBomDir }
  if (Test-Path -LiteralPath $parentBomDir) { $bomSearchDirs += $parentBomDir }
  if (Test-Path -LiteralPath $localBomDir) { $bomSearchDirs += $localBomDir }
  Write-Log "BOM search dirs: $($bomSearchDirs -join ', ')"

  # BOM filename date parser
  function Parse-BomDate {
    param([string]$name)
    if ($name -match 'BOM\s*&?\s*FORMULA\s+([A-Za-z]+)\s+(\d{1,2})\s+D365\.xlsx') {
      $months = @{ JAN=1;FEB=2;MAR=3;APR=4;MAY=5;JUN=6;JUL=7;AUG=8;SEP=9;OCT=10;NOV=11;DEC=12 }
      $m = $matches[1].ToUpper().Substring(0,3)
      $mon = $months[$m]
      if ($mon) {
        $d = [int]$matches[2]
        if ($d -ge 1 -and $d -le 31) {
          return [DateTime]::new(2026, $mon, $d)
        }
      }
    }
    return [DateTime]::MinValue
  }

  # Generate Weekly BOM Manifest
  $bomEntries = @()
  foreach ($dir in $bomSearchDirs) {
    $prefix = ''
    if ($dir -eq $parentBomDir) { $prefix = '../1. Weekly BOM/' }
    elseif ($dir -eq $localBomDir) { $prefix = '1. Weekly BOM/' }
    $files = Get-ChildItem -LiteralPath $dir -Filter 'BOM*FORMULA*D365.xlsx' -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '*Physical Dimension*' }
    foreach ($f in $files) {
      $bomDate = Parse-BomDate $f.Name
      $dateStr = if ($bomDate -ne [DateTime]::MinValue) { $bomDate.ToString('yyyy-MM-dd') } else { 'unknown' }
      Write-Log "BOM candidate: $($f.Name) date=$dateStr modifiedUtc=$($f.LastWriteTimeUtc.ToString('o'))"
      $bomEntries += [pscustomobject]@{
        name = $f.Name
        path = $prefix + $f.Name
        modifiedUtc = $f.LastWriteTimeUtc.ToString('o')
        sourceType = 'weeklyBom'
        _bomDate = $bomDate
      }
    }
  }
  # Sort by parsed BOM date descending, then modifiedUtc descending
  $bomEntries = @($bomEntries | Sort-Object { if ($_._bomDate -eq [DateTime]::MinValue) { [DateTime]::MinValue } else { $_._bomDate } }, @{Expression={$_.modifiedUtc}; Descending=$true} -Descending)
  $bomOutput = @($bomEntries | Select-Object name, path, modifiedUtc, sourceType)
  if ($bomOutput.Count -gt 0) {
    Write-Log "Selected BOM: $($bomOutput[0].name)"
  } else {
    Write-Log "WARNING: No BOM files found"
  }
  $bomJson = ConvertTo-Json -Compress -Depth 3 -InputObject @($bomOutput)
  Set-Content -Encoding UTF8 -LiteralPath (Join-Path $tool 'Weekly BOM Manifest.json') -Value $bomJson

  # Search for dimension workbook
  $dimPatterns = @(
    'Products Physical Dimension - FG 1.Apr.25.xlsx',
    '*Products Physical Dimension - FG *.xlsx',
    '*Products Physical Dimension*.xlsx',
    '*Physical Dimension*.xlsx',
    '*Dimension*FG*.xlsx',
    'carton_tool_input.xlsx'
  )
  $dimSearchDirs = @()
  if ($weeklyBomDir -and (Test-Path -LiteralPath $weeklyBomDir)) { $dimSearchDirs += $weeklyBomDir }
  if (Test-Path -LiteralPath $localBomDir) { $dimSearchDirs += $localBomDir }
  if ($npiToolsDir -and (Test-Path -LiteralPath $npiToolsDir)) { $dimSearchDirs += $npiToolsDir }
  $dimSearchDirs += $tool
  Write-Log "Dimension search dirs: $($dimSearchDirs -join ', ')"

  $dimFile = $null
  $dimDir = $null
  foreach ($d in $dimSearchDirs) {
    if (-not (Test-Path -LiteralPath $d)) { continue }
    foreach ($pat in $dimPatterns) {
      $dimFile = Get-ChildItem -LiteralPath $d -Filter $pat -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '~$*' } |
        Sort-Object LastWriteTimeUtc -Descending |
        Select-Object -First 1
      if ($dimFile) {
        $dimDir = $d
        Write-Log "Dimension workbook found: $($dimFile.FullName)"
        break
      }
    }
    if ($dimFile) { break }
  }

  $dims = @()
  if ($dimFile) {
    $dp = ''
    if ($dimDir -eq $parentBomDir) { $dp = '../1. Weekly BOM/' }
    elseif ($dimDir -eq $localBomDir) { $dp = '1. Weekly BOM/' }
    $url = ([System.Uri]$dimFile.FullName).AbsoluteUri
    $dims += [pscustomobject]@{
      name = $dimFile.Name
      path = $dp + $dimFile.Name
      url = $url
      modifiedUtc = $dimFile.LastWriteTimeUtc.ToString('o')
      sourceType = 'dimension'
    }
    Write-Log "Dimension workbook URL: $url"
  } else {
    Write-Log "WARNING: No dimension workbook found"
  }
  $dimJson = ConvertTo-Json -Compress -Depth 3 -InputObject @($dims)
  Set-Content -Encoding UTF8 -LiteralPath (Join-Path $tool 'WVN Master Carton Workbook Manifest.json') -Value $dimJson

  if ($manifestsOnly) {
    Write-Log "Manifests written (--manifests-only). Skipping browser launch."
  } else {
    # Launch browser
    $htmlPath = Join-Path $tool 'WVN_Master_Carton_Fit_Tool.html'
    $chromePath = Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'
    $edgePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'
    $browser = $null
    if (Test-Path -LiteralPath $chromePath) { $browser = $chromePath }
    elseif (Test-Path -LiteralPath $edgePath) { $browser = $edgePath }
    if ($browser) {
      $tempProfile = Join-Path $env:TEMP "WVNMasterCartonToolChrome-$(Get-Random)"
      Write-Log "Launching browser: $browser"
      Write-Log "HTML: $htmlPath"
      Start-Process -FilePath $browser -ArgumentList "--user-data-dir=`"$tempProfile`"", '--allow-file-access-from-files', "`"$htmlPath`""
    } else {
      Write-Log "ERROR: Neither Chrome nor Edge found. Open the HTML file manually."
    }
  }
  Write-Log "Done."
} catch {
  $errMsg = "ERROR: $($_.Exception.Message)"
  Write-Host $errMsg
  try { Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $errMsg" } catch {}
}
