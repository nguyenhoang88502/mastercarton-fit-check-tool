$ErrorActionPreference = 'Stop'

# Detect script directory and application/tools directory structure
$ScriptDirectory = $PSScriptRoot
if (-not $ScriptDirectory) {
  $ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
}

if ((Split-Path $ScriptDirectory -Leaf) -ieq 'system') {
  $AppRoot = Split-Path $ScriptDirectory -Parent
} else {
  $AppRoot = $ScriptDirectory
}

$ToolsParent = Split-Path $AppRoot -Parent
$WeeklyBomDirectory = Join-Path $ToolsParent '1. Weekly BOM'

# Map back to original variable names for compatibility
$systemDir = $ScriptDirectory
$projectDir = $AppRoot

$manifestsOnly = $args[0] -eq '--manifests-only'
$forceRebuildArg = $args[0] -eq '--force-rebuild'

$logDir = Join-Path $systemDir 'logs'
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logPath = Join-Path $logDir 'WVN_Master_Carton_Fit_Tool_launcher.log'

function Write-Log {
  param([string]$msg)
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $line = "$ts $msg"
  Write-Host $line
  Add-Content -LiteralPath $logPath -Value $line
}

# Clear old log
if (Test-Path -LiteralPath $logPath) { Remove-Item -LiteralPath $logPath -Force }

Write-Log "=== WVN Master Carton Fit Tool Launcher ==="
Write-Log "System Directory: $systemDir"
Write-Log "Project Directory: $projectDir"

function Get-ExcelColumnIndex {
  param([string]$cellReference)
  $letters = ([regex]::Match($cellReference, '^[A-Za-z]+')).Value.ToUpperInvariant()
  $index = 0
  foreach ($ch in $letters.ToCharArray()) { $index = ($index * 26) + ([int]$ch - [int][char]'A' + 1) }
  return $index - 1
}

function Read-XlsxSheets {
  param([System.IO.FileInfo]$File, [string]$SheetNamePattern = '.*', [int]$SkipSheets = 0)
  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  if (-not ('WvnXlsxReader' -as [type])) {
    Add-Type -ReferencedAssemblies 'System.IO.Compression.dll','System.IO.Compression.FileSystem.dll','System.Xml.dll' -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;
public sealed class WvnSheetRows { public string Name; public List<string[]> Rows = new List<string[]>(); }
public static class WvnXlsxReader {
  static string EntryText(ZipArchive z, string name) { var e=z.GetEntry(name); if(e==null)return null; using(var s=e.Open())using(var r=new StreamReader(s))return r.ReadToEnd(); }
  static int Col(string reference) { int n=0; foreach(char c in reference){ if(c<'A'||c>'Z')break; n=n*26+(c-'A'+1); } return n-1; }
  public static List<WvnSheetRows> Read(string path, string pattern, int skip) {
    var output=new List<WvnSheetRows>();
    using(var fs=File.OpenRead(path))using(var z=new ZipArchive(fs,ZipArchiveMode.Read,false)) {
      var shared=new List<string>(); string st=EntryText(z,"xl/sharedStrings.xml");
      if(st!=null){var d=new XmlDocument();d.LoadXml(st);foreach(XmlNode n in d.SelectNodes("//*[local-name()='si']"))shared.Add(n.InnerText);}
      var wb=new XmlDocument();wb.LoadXml(EntryText(z,"xl/workbook.xml")); var rd=new XmlDocument();rd.LoadXml(EntryText(z,"xl/_rels/workbook.xml.rels"));
      var rel=new Dictionary<string,string>();foreach(XmlNode n in rd.SelectNodes("//*[local-name()='Relationship']"))rel[n.Attributes["Id"].Value]=n.Attributes["Target"].Value;
      int number=0; foreach(XmlNode sn in wb.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
        if(number++<skip)continue; string name=sn.Attributes["name"].Value;if(!Regex.IsMatch(name,pattern))continue;
        string rid=sn.Attributes["r:id"].Value,target;if(!rel.TryGetValue(rid,out target))continue; target=target.TrimStart('/');if(!target.StartsWith("xl/"))target="xl/"+target.TrimStart('.','/');
        var entry=z.GetEntry(target);if(entry==null)continue;var sheet=new WvnSheetRows();sheet.Name=name;
        using(var stream=entry.Open())using(var xr=XmlReader.Create(stream,new XmlReaderSettings{IgnoreComments=true,IgnoreWhitespace=true})) {
          while(xr.Read())if(xr.NodeType==XmlNodeType.Element&&xr.LocalName=="row") {
            var cells=new Dictionary<int,string>();int max=-1;
            using(var rr=xr.ReadSubtree())while(rr.Read())if(rr.NodeType==XmlNodeType.Element&&rr.LocalName=="c") {
              string cr=rr.GetAttribute("r")??"A",type=rr.GetAttribute("t")??"",value="";int ci=Col(cr.ToUpperInvariant());
              using(var cc=rr.ReadSubtree())while(cc.Read())if(cc.NodeType==XmlNodeType.Element&&(cc.LocalName=="v"||cc.LocalName=="t"))value=cc.ReadElementContentAsString();
              int si;if(type=="s"&&Int32.TryParse(value,out si)&&si>=0&&si<shared.Count)value=shared[si];cells[ci]=(value??"").Trim();if(ci>max)max=ci;
            }
            var row=new string[Math.Max(0,max+1)];foreach(var kv in cells)row[kv.Key]=kv.Value;for(int i=0;i<row.Length;i++)if(row[i]==null)row[i]="";sheet.Rows.Add(row);
          }
        }
        output.Add(sheet);
      }
    }
    return output;
  }
}
'@
  }
  return [WvnXlsxReader]::Read($File.FullName, $SheetNamePattern, $SkipSheets)
}

function Normalize-Header { param($value) return ([string]$value).Trim().ToLowerInvariant() -replace '[^a-z0-9]+', ' ' -replace '^\s+|\s+$', '' }
function Get-RowCell { param($row, [int]$index) if ($index -lt 0 -or $null -eq $row -or $index -ge $row.Count) { return '' }; return ([string]$row[$index]).Trim() }

function Get-TargetQuantity {
  param($perSeries, $norm, $quantity)
  $number = 0.0
  if ([double]::TryParse([string]$perSeries, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and $number -gt 1) { return [int][math]::Floor($number) }
  foreach ($raw in @($norm, $quantity)) {
    $number = 0.0
    if ([double]::TryParse([string]$raw, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and $number -gt 0) {
      if ($number -lt 1) { return [int][math]::Round(1 / $number) }
      return [int][math]::Floor($number)
    }
  }
  return $null
}

function Get-QuantitySource {
  param($perSeries, $norm, $quantity)
  $number = 0.0
  if ([double]::TryParse([string]$perSeries, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and $number -gt 1) { return 'weeklyBomPerSeries' }
  if ([double]::TryParse([string]$norm, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and $number -gt 0) { return 'weeklyBomNorm' }
  if ([double]::TryParse([string]$quantity, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and $number -gt 0) { return 'weeklyBomQuantity' }
  return 'None'
}

function Normalize-ItemId {
  param($Value)
  if ($null -eq $Value) { return '' }
  $text = ([string]$Value).Trim()
  if ($text -match '^\d+\.0+$') { return $text -replace '\.0+$', '' }
  return $text
}

function Convert-ToFiniteNumber {
  param($Value)
  $number = 0.0
  if ([double]::TryParse(([string]$Value), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
  return [double]::NaN
}

function Get-HeaderInfo {
  param($Sheet, [object[]]$RequiredAliasGroups)
  for ($rowIndex = 0; $rowIndex -lt [math]::Min(40, $Sheet.Rows.Count); $rowIndex++) {
    $map = @{}
    for ($columnIndex = 0; $columnIndex -lt $Sheet.Rows[$rowIndex].Count; $columnIndex++) {
      $key = Normalize-Header $Sheet.Rows[$rowIndex][$columnIndex]
      if ($key -and -not $map.ContainsKey($key)) { $map[$key] = $columnIndex }
    }
    $complete = $true
    foreach ($aliases in $RequiredAliasGroups) {
      $found = $false
      foreach ($alias in @($aliases)) { if ($map.ContainsKey((Normalize-Header $alias))) { $found = $true; break } }
      if (-not $found) { $complete = $false; break }
    }
    if ($complete) { return [pscustomobject]@{ Row = $rowIndex; Map = $map } }
  }
  return $null
}

function Find-HeaderColumn {
  param($HeaderInfo, [string[]]$Aliases)
  if (-not $HeaderInfo) { return -1 }
  foreach ($alias in $Aliases) { $key = Normalize-Header $alias; if ($HeaderInfo.Map.ContainsKey($key)) { return [int]$HeaderInfo.Map[$key] } }
  return -1
}

# INI file parser helper
function Read-IniFile {
  param([string]$FilePath)
  $ini = @{}
  if (-not (Test-Path -LiteralPath $FilePath)) { return $ini }
  $section = ""
  foreach ($line in Get-Content -LiteralPath $FilePath) {
    $line = $line.Trim()
    if ($line.StartsWith(';') -or $line.StartsWith('#') -or -not $line) { continue }
    if ($line.StartsWith('[') -and $line.EndsWith(']')) {
      $section = $line.Substring(1, $line.Length - 2).Trim()
      $ini[$section] = @{}
    } elseif ($line -match '^([^=]+)=(.*)$') {
      $key = $matches[1].Trim()
      $val = $matches[2].Trim()
      if ($section) {
        $ini[$section][$key] = $val
      }
    }
  }
  return $ini
}

# Validation checks for Excel structure
function Test-IsValidDimensionWorkbook {
  param([System.IO.FileInfo]$File)
  try {
    $sheets = @(Read-XlsxSheets -File $File)
    $hasDimensionSheet = $false
    foreach ($sheet in $sheets) {
      $headerInfo = Get-HeaderInfo $sheet @(@('item number','item no','item'), @('unit','uom'), @('width'), @('depth','length'), @('height'))
      if ($headerInfo) { $hasDimensionSheet = $true; break }
    }
    return $hasDimensionSheet
  } catch {
    return $false
  }
}

function Test-IsValidBomWorkbook {
  param([System.IO.FileInfo]$File)
  try {
    $sheets = @(Read-XlsxSheets -File $File -SkipSheets 1)
    $hasBomSheet = $false
    foreach ($sheet in $sheets) {
      for ($hr = 0; $hr -lt [math]::Min(60, $sheet.Rows.Count); $hr++) {
        $hh = @($sheet.Rows[$hr] | ForEach-Object { Normalize-Header $_ })
        $fgIdx = -1; $itemIdx = -1; $productIdx = -1
        for ($hc = 0; $hc -lt $hh.Count; $hc++) {
          if ($hh[$hc] -in @('fg','fg id','fgid','finished good','finishedgood','top item id')) { $fgIdx=$hc }
          if ($hh[$hc] -in @('item number','item no','itemnumber','component','component item','child item id')) { $itemIdx=$hc }
          if ($hh[$hc] -in @('product name','product','productname','description','item description','child item name')) { $productIdx=$hc }
        }
        if ($fgIdx -ge 0 -and $itemIdx -ge 0 -and $productIdx -ge 0) { $hasBomSheet = $true; break }
      }
      if ($hasBomSheet) { break }
    }
    return $hasBomSheet
  } catch {
    return $false
  }
}

function Get-BomBusinessDateFromFilename {
  param(
    [Parameter(Mandatory = $true)]
    [System.IO.FileInfo]$File
  )
  
  $res = [pscustomobject]@{
    Success = $false
    BusinessDate = [DateTime]::MinValue
    MonthToken = ""
    Day = 0
    InferredYear = 0
    Error = ""
  }
  
  $MonthNumbers = @{
    JAN = 1; FEB = 2; MAR = 3; APR = 4; MAY = 5; JUN = 6;
    JUL = 7; AUG = 8; SEP = 9; OCT = 10; NOV = 11; DEC = 12
  }
  
  $name = $File.Name
  # Case-insensitive match
  if ($name -match '^BOM\s*&\s*FORMULA\s+([A-Za-z]{3})\s+(\d{1,2})\s+D365\.(xlsx|xlsm)$') {
    $monthToken = $Matches[1].ToUpperInvariant()
    $dayVal = [int]$Matches[2]
    
    if (-not $MonthNumbers.ContainsKey($monthToken)) {
      $res.Error = "Unrecognized month token '$monthToken'"
      return $res
    }
    
    $monthVal = $MonthNumbers[$monthToken]
    $mYear = $File.LastWriteTimeUtc.Year
    
    $candidateDate = $null
    try {
      $candidateDate = [DateTime]::new([int]$mYear, [int]$monthVal, [int]$dayVal, 0, 0, 0, [System.DateTimeKind]::Utc)
    } catch {
      # Try previous year (leap year Feb 29 fallback)
      try {
        $candidateDate = [DateTime]::new(([int]$mYear - 1), [int]$monthVal, [int]$dayVal, 0, 0, 0, [System.DateTimeKind]::Utc)
        $mYear = $mYear - 1
      } catch {
        $res.Error = "Invalid date '$mYear-$monthVal-$dayVal' ($monthToken $dayVal). Error: $_"
        return $res
      }
    }
    
    $modDate = $File.LastWriteTimeUtc.Date
    # Compare parsed date with modification date
    # If the candidate date is implausibly far in the future relative to modification date (e.g. > 45 days)
    # subtract one year.
    $diffDays = ($candidateDate - $modDate).TotalDays
    if ($diffDays -gt 45) {
      $mYear = $mYear - 1
      try {
        $candidateDate = [DateTime]::new([int]$mYear, [int]$monthVal, [int]$dayVal, 0, 0, 0, [System.DateTimeKind]::Utc)
      } catch {
        $res.Error = "Invalid date after year adjustment: $mYear-$monthVal-$dayVal. Error: $_"
        return $res
      }
    }
    
    $res.Success = $true
    $res.BusinessDate = $candidateDate
    $res.MonthToken = $monthToken
    $res.Day = $dayVal
    $res.InferredYear = $mYear
    Write-Log "Inferred business date for '$name': $($candidateDate.ToString('yyyy-MM-dd')) (Year inferred as $mYear from LastWriteTimeUtc)"
  } else {
    $res.Error = "Filename does not match expected pattern BOM & FORMULA [MONTH] [DAY] D365"
  }
  return $res
}

# Read configuration
$configPath = Join-Path $systemDir 'config.ini'
$config = Read-IniFile -FilePath $configPath

# Resolve search directory paths
$configuredWeeklyBomFolder = $null
if ($config.ContainsKey('Search')) {
  # Priority configured entry: WeeklyBomFolder
  if ($config['Search'].ContainsKey('WeeklyBomFolder') -and $config['Search']['WeeklyBomFolder']) {
    $val = $config['Search']['WeeklyBomFolder']
    $resolvedPath = $val
    if (-not [System.IO.Path]::IsPathRooted($val)) {
      $resolvedPath = [System.IO.Path]::GetFullPath((Join-Path $AppRoot $val))
    }
    if (Test-Path -LiteralPath $resolvedPath) {
      $configuredWeeklyBomFolder = $resolvedPath
      Write-Log "Configured WeeklyBomFolder: $configuredWeeklyBomFolder"
    } else {
      Write-Log "WARNING: Configured WeeklyBomFolder path '$val' (resolved: '$resolvedPath') does not exist."
    }
  }

  # Backward compatibility: Root1..Root10 if WeeklyBomFolder wasn't set/valid
  if (-not $configuredWeeklyBomFolder) {
    for ($i = 1; $i -le 10; $i++) {
      $key = "Root$i"
      if ($config['Search'].ContainsKey($key) -and $config['Search'][$key]) {
        $val = $config['Search'][$key]
        $resolvedPath = $val
        if (-not [System.IO.Path]::IsPathRooted($val)) {
          $resolvedPath = [System.IO.Path]::GetFullPath((Join-Path $AppRoot $val))
        }
        if (Test-Path -LiteralPath $resolvedPath) {
          $configuredWeeklyBomFolder = $resolvedPath
          Write-Log "Configured Root$i override: $configuredWeeklyBomFolder"
          break
        } else {
          Write-Log "WARNING: Configured Root$i path '$val' (resolved: '$resolvedPath') does not exist or is invalid."
        }
      }
    }
  }
}

# Determine target Weekly BOM directory using priorities
$resolvedWeeklyBomDir = $null
$siblingWeeklyBomDir = Join-Path $ToolsParent '1. Weekly BOM'
$siblingDiscoverySucceeded = $false

if ($configuredWeeklyBomFolder) {
  $resolvedWeeklyBomDir = $configuredWeeklyBomFolder
  Write-Log "Using configured path override: $resolvedWeeklyBomDir"
} else {
  # Priority 1: Exact sibling directory
  if (Test-Path -LiteralPath $siblingWeeklyBomDir) {
    $resolvedWeeklyBomDir = (Get-Item -LiteralPath $siblingWeeklyBomDir).FullName
    $siblingDiscoverySucceeded = $true
    Write-Log "Sibling Weekly BOM directory found and selected: $resolvedWeeklyBomDir"
  } else {
    Write-Log "Sibling directory '1. Weekly BOM' not found at '$siblingWeeklyBomDir'."
    
    # Priority 4: Search ToolsParent for a directory named '1. Weekly BOM' (non-recursively / depth 1)
    Write-Log "Searching parent directory '$ToolsParent' for a folder named '1. Weekly BOM'..."
    $foundDirs = Get-ChildItem -LiteralPath $ToolsParent -Directory -Filter '1. Weekly BOM' -ErrorAction SilentlyContinue
    if ($foundDirs.Count -gt 0) {
      $resolvedWeeklyBomDir = $foundDirs[0].FullName
      $siblingDiscoverySucceeded = $true
      Write-Log "Discovered '1. Weekly BOM' under parent: $resolvedWeeklyBomDir"
    } else {
      # Priority 3: Fall back to application root
      if (Test-Path -LiteralPath $AppRoot) {
        $resolvedWeeklyBomDir = (Get-Item -LiteralPath $AppRoot).FullName
        Write-Log "Using compatibility fallback (application root): $resolvedWeeklyBomDir"
      }
    }
  }
}

# If resolved directory is still null or does not exist, produce a clear actionable error message
if (-not $resolvedWeeklyBomDir -or -not (Test-Path -LiteralPath $resolvedWeeklyBomDir)) {
  $errMsg = @"
The folder "1. Weekly BOM" was not found beside the Master Carton tool.

Expected location:
$siblingWeeklyBomDir

Move or sync the Weekly BOM folder to that location, update config.ini, or open the workbooks manually from the application.
"@
  Write-Log "ERROR: $errMsg"
  if (-not $manifestsOnly) {
    try {
      Add-Type -AssemblyName System.Windows.Forms
      [void][System.Windows.Forms.MessageBox]::Show($errMsg, "Master Carton Fit Tool - Directory Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } catch {}
  }
}

# Log search location information as requested
Write-Log "Script directory: $ScriptDirectory"
Write-Log "Resolved application root: $AppRoot"
Write-Log "Resolved parent tools directory: $ToolsParent"
Write-Log "Expected Weekly BOM directory: $siblingWeeklyBomDir"
if ($configuredWeeklyBomFolder) {
  Write-Log "Configured override: $configuredWeeklyBomFolder"
} else {
  Write-Log "Configured override: [None]"
}
Write-Log "Whether sibling discovery succeeded: $siblingDiscoverySucceeded"
Write-Log "Search scope: Sibling/configured local Weekly BOM directory only"
Write-Log "Broader search (OneDrive root, UserProfile, SharePoint, mounted drives) was intentionally skipped to prevent slow scanning and hydration issues."

# 1. Locate Dimension Workbook
Write-Log "Checking source data..."
Write-Log "Preparing carton data..."
Write-Log "Phase: Locating Dimension Workbook"

$dimensionFile = $null
$dimCandidatesList = [System.Collections.ArrayList]::new()

if ($resolvedWeeklyBomDir -and (Test-Path -LiteralPath $resolvedWeeklyBomDir)) {
  # 1a. Try canonical dimension filename: Products Physical Dimension - FG 1.Apr.25.xlsx
  $canonicalDimPath = Join-Path $resolvedWeeklyBomDir 'Products Physical Dimension - FG 1.Apr.25.xlsx'
  $canonicalDimFileInfo = New-Object System.IO.FileInfo -ArgumentList $canonicalDimPath

  if (Test-Path -LiteralPath $canonicalDimPath) {
    [void]$dimCandidatesList.Add($canonicalDimFileInfo)
    Write-Log "Found canonical dimension file: $($canonicalDimFileInfo.FullName). Validating..."
    try {
      $stream = [System.IO.File]::OpenRead($canonicalDimFileInfo.FullName)
      $stream.Close()
      
      if (Test-IsValidDimensionWorkbook -File $canonicalDimFileInfo) {
        $dimensionFile = $canonicalDimFileInfo
        Write-Log "Selected canonical dimension file: $($dimensionFile.FullName)"
      } else {
        Write-Log "Rejected canonical dimension file: Failed structural validation."
      }
    } catch {
      Write-Log "Rejected canonical dimension file: File read/availability failure (could be cloud-only). Error: $_"
    }
  }

  # 1b. If canonical was not found or valid, find matching dimension workbooks inside the local folder
  if (-not $dimensionFile) {
    $dimCandidates = Get-ChildItem -LiteralPath $resolvedWeeklyBomDir -File -ErrorAction SilentlyContinue |
      Where-Object {
        $_.Name -notlike '~$*' -and
        $_.Length -gt 0 -and
        (
          $_.Name -like 'Products Physical Dimension*.xlsx' -or
          $_.Name -like 'Products Physical Dimension*.xlsm' -or
          $_.Name -like $dimPattern -or
          $_.Name -like ($dimPattern -replace '\.xlsx$', '.xlsm')
        ) -and
        $_.FullName -ne $canonicalDimPath
      } |
      Sort-Object @{ Expression = 'LastWriteTimeUtc'; Descending = $true }, @{ Expression = 'Name'; Descending = $true }

    foreach ($candidate in $dimCandidates) {
      [void]$dimCandidatesList.Add($candidate)
    }

    Write-Log "Found $($dimCandidates.Count) alternative local dimension candidates. Evaluating in order..."
    foreach ($candidate in $dimCandidates) {
      try {
        $stream = [System.IO.File]::OpenRead($candidate.FullName)
        $stream.Close()
        
        if (Test-IsValidDimensionWorkbook -File $candidate) {
          $dimensionFile = $candidate
          Write-Log "Selected alternative dimension file: $($dimensionFile.FullName) (Last Modified: $($dimensionFile.LastWriteTimeUtc))"
          break
        } else {
          Write-Log "Rejected dimension candidate '$($candidate.Name)': Failed structural validation."
        }
      } catch {
        Write-Log "Rejected dimension candidate '$($candidate.Name)': File read/availability failure (could be cloud-only). Error: $_"
      }
    }
  }
}

# Log findings
if ($dimCandidatesList.Count -gt 0) {
  $names = @($dimCandidatesList | ForEach-Object { $_.Name })
  Write-Log "Dimension candidates found locally: $($names -join ', ')"
} else {
  Write-Log "Dimension candidates found locally: [None]"
}

if ($dimensionFile) {
  Write-Log "Selected dimension workbook: $($dimensionFile.FullName)"
} else {
  Write-Log "WARNING: Dimension Workbook not found or could not be read."
  $errMsg = @"
No valid Dimension Workbook was found or could be read in the local search directory.

Expected canonical filename:
Products Physical Dimension - FG 1.Apr.25.xlsx

Search Location:
$resolvedWeeklyBomDir

Please ensure a valid Excel file exists and is available locally (fully downloaded from OneDrive), or open the workbooks manually from the application.
"@
  if (-not $manifestsOnly) {
    try {
      Add-Type -AssemblyName System.Windows.Forms
      [void][System.Windows.Forms.MessageBox]::Show($errMsg, "Master Carton Fit Tool - Dimension Workbook Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } catch {}
  }
}


# 2. Locate latest BOM Workbook
Write-Log "Phase: Locating latest BOM Workbook"
$bomFile = $null
$bomCandidatesList = [System.Collections.ArrayList]::new()
$decoratedCandidates = [System.Collections.ArrayList]::new()

if ($resolvedWeeklyBomDir -and (Test-Path -LiteralPath $resolvedWeeklyBomDir)) {
  # Enumerate local candidates matching the BOM patterns
  $bomCandidates = Get-ChildItem -LiteralPath $resolvedWeeklyBomDir -File -ErrorAction SilentlyContinue |
    Where-Object {
      $_.Name -notlike '~$*' -and
      $_.Length -gt 0 -and
      (
        $_.Name -like 'BOM & FORMULA*D365.xlsx' -or
        $_.Name -like 'BOM & FORMULA*D365.xlsm' -or
        $_.Name -like $bomPattern -or
        $_.Name -like ($bomPattern -replace '\.xlsx$', '.xlsm')
      )
    }

  # Parse dates and decorate candidates first
  foreach ($candidate in $bomCandidates) {
    [void]$bomCandidatesList.Add($candidate)
    
    $parsed = Get-BomBusinessDateFromFilename -File $candidate
    if (-not $parsed.Success) {
      Write-Log "Filename date could not be parsed for '$($candidate.Name)'; candidate eligible only for modification-time fallback. Reason: $($parsed.Error)"
    }
    
    [void]$decoratedCandidates.Add([pscustomobject]@{
      File                = $candidate
      Name                = $candidate.Name
      ParsedBusinessDate  = $parsed.BusinessDate
      FilenameDateValid   = $parsed.Success
      LastWriteTimeUtc    = $candidate.LastWriteTimeUtc
      WorkbookValid       = $null
      ValidationError     = ""
    })
  }
}

# Sort decorated candidates by Priority:
# 1. FilenameDateValid descending
# 2. ParsedBusinessDate descending
# 3. LastWriteTimeUtc descending
# 4. Name descending (stable final tie-breaker)
$sortedCandidates = @($decoratedCandidates | Sort-Object `
  @{ Expression = 'FilenameDateValid'; Descending = $true },
  @{ Expression = 'ParsedBusinessDate'; Descending = $true },
  @{ Expression = 'LastWriteTimeUtc'; Descending = $true },
  @{ Expression = 'Name'; Descending = $true })

# Validate all candidates and log their details
foreach ($candidate in $sortedCandidates) {
  try {
    # Check readability (OneDrive cloud-only check)
    $stream = [System.IO.File]::OpenRead($candidate.File.FullName)
    $stream.Close()

    if (Test-IsValidBomWorkbook -File $candidate.File) {
      $candidate.WorkbookValid = $true
    } else {
      $candidate.WorkbookValid = $false
      $candidate.ValidationError = "Failed structural validation."
    }
  } catch {
    $candidate.WorkbookValid = $false
    $candidate.ValidationError = "File read/availability failure. Error: $_"
  }
}

# Log findings
foreach ($candidate in $sortedCandidates) {
  Write-Log "BOM candidate:"
  Write-Log "  Name: $($candidate.Name)"
  if ($candidate.FilenameDateValid) {
    Write-Log "  Parsed business date: $($candidate.ParsedBusinessDate.ToString('yyyy-MM-dd'))"
  } else {
    Write-Log "  Parsed business date: [Unparseable]"
  }
  # Convert UTC to standard log format
  Write-Log "  Last modified UTC: $($candidate.LastWriteTimeUtc.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))"
  if ($candidate.WorkbookValid) {
    Write-Log "  Workbook validation: valid"
  } else {
    Write-Log "  Workbook validation: invalid ($($candidate.ValidationError))"
  }
}

# Select the first candidate that is structurally valid
$selectedReason = ""
foreach ($candidate in $sortedCandidates) {
  if ($candidate.WorkbookValid) {
    $bomFile = $candidate.File
    if ($candidate.FilenameDateValid) {
      $selectedReason = "Latest valid business date parsed from filename."
    } else {
      $selectedReason = "Modification-time fallback (filename date could not be parsed)."
    }
    break
  }
}

if ($bomFile) {
  Write-Log "Selected BOM:"
  Write-Log "  $($bomFile.Name)"
  Write-Log "Selection reason:"
  Write-Log "  $selectedReason"
} else {
  Write-Log "WARNING: BOM Workbook not found or could not be read."
  $errMsg = @"
No valid BOM Workbook was found or could be read in the local search directory.

Expected pattern:
BOM & FORMULA*D365.xlsx or .xlsm

Search Location:
$resolvedWeeklyBomDir

Please ensure a valid Excel file exists and is available locally (fully downloaded from OneDrive), or open the workbooks manually from the application.
"@
  if (-not $manifestsOnly) {
    try {
      Add-Type -AssemblyName System.Windows.Forms
      [void][System.Windows.Forms.MessageBox]::Show($errMsg, "Master Carton Fit Tool - BOM Workbook Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } catch {}
  }
}

# Generate cache freshness checks
$generatedDir = Join-Path $systemDir 'generated'
if (-not (Test-Path -LiteralPath $generatedDir)) { New-Item -ItemType Directory -Path $generatedDir -Force | Out-Null }
$cacheStatePath = Join-Path $generatedDir 'cache_state.ini'

$dimHash = ""
$bomHash = ""
if ($dimensionFile) { $dimHash = (Get-FileHash -LiteralPath $dimensionFile.FullName -Algorithm SHA256).Hash.ToLowerInvariant() }
if ($bomFile) { $bomHash = (Get-FileHash -LiteralPath $bomFile.FullName -Algorithm SHA256).Hash.ToLowerInvariant() }

$cacheVersion = "1.0"
$rebuildNeeded = $forceRebuild -or (-not (Test-Path -LiteralPath $cacheStatePath))

if (-not $rebuildNeeded) {
  $cacheState = Read-IniFile -FilePath $cacheStatePath
  
  $dimCacheMatch = $false
  if ($cacheState.ContainsKey('DimensionWorkbook') -and $dimensionFile) {
    $c = $cacheState['DimensionWorkbook']
    if ($c['Path'] -eq $dimensionFile.FullName -and
        $c['Size'] -eq $dimensionFile.Length.ToString() -and
        $c['LastModifiedUtc'] -eq $dimensionFile.LastWriteTimeUtc.ToString('o') -and
        $c['Hash'] -eq $dimHash -and
        $c['CacheSchemaVersion'] -eq $cacheVersion) {
      $dimCacheMatch = $true
    }
  }
  
  $bomCacheMatch = $false
  if ($cacheState.ContainsKey('BomWorkbook') -and $bomFile) {
    $c = $cacheState['BomWorkbook']
    if ($c['Path'] -eq $bomFile.FullName -and
        $c['Size'] -eq $bomFile.Length.ToString() -and
        $c['LastModifiedUtc'] -eq $bomFile.LastWriteTimeUtc.ToString('o') -and
        $c['Hash'] -eq $bomHash -and
        $c['CacheSchemaVersion'] -eq $cacheVersion) {
      $bomCacheMatch = $true
    }
  }
  
  $rebuildNeeded = -not ($dimCacheMatch -and $bomCacheMatch)
}

if ($rebuildNeeded) {
  Write-Log "Caches are stale, missing, or force-rebuild requested. Rebuilding CSV data..."
} else {
  Write-Log "Cache is fresh. Reusing prepared CSV data."
}

# We initialize lists for exporting CSVs
$dataQuality = [System.Collections.ArrayList]::new()
function Add-DataQuality {
  param([string]$issue, [string]$entityType, [string]$entityId, [string]$desc, [string]$detail)
  [void]$dataQuality.Add([pscustomobject]@{
    issue_type = $issue
    entity_type = $entityType
    entity_id = $entityId
    description = $desc
    source_detail = $detail
  })
}

if ($rebuildNeeded) {
  # Initialize indexes
  $knownCartonIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  $allCartons = @{} # cartonId -> carton_dimensions.csv record
  $allFgs = @{} # fgId -> fg_dimensions.csv record
  $assignments = [System.Collections.ArrayList]::new()
  $reverseIndex = [System.Collections.ArrayList]::new()

  # Read sheets from workbooks
  $sheets = @()
  $dimensionSheet = $null
  $backupSheet = $null
  
  if ($dimensionFile) {
    Write-Log "Reading Dimension Workbook..."
    $sheets = @(Read-XlsxSheets -File $dimensionFile)
    $dimensionSheet = $sheets | Where-Object Name -eq 'Dimension' | Select-Object -First 1
    $backupSheet = $sheets | Where-Object Name -eq 'Mastercarton_Dimensions' | Select-Object -First 1
    
    # 1. Pre-index known Carton IDs from sheets that look like carton dimensions
    foreach ($sh in $sheets) {
      if ($sh.Name -match '(?i)master.*carton|carton.*dimension') {
        $idCol = -1
        for ($r = 0; $r -lt [math]::Min(40, $sh.Rows.Count); $r++) {
          $headers = @($sh.Rows[$r] | ForEach-Object { Normalize-Header $_ })
          for ($c = 0; $c -lt $headers.Count; $c++) {
            if ($headers[$c] -in @('mastercarton id','master carton id','carton id','item number')) { $idCol = $c; break }
          }
          if ($idCol -ge 0) {
            for ($i = $r + 1; $i -lt $sh.Rows.Count; $i++) {
              $id = Normalize-ItemId (Get-RowCell $sh.Rows[$i] $idCol)
              if ($id) { [void]$knownCartonIds.Add($id) }
            }
            break
          }
        }
      }
    }
    Write-Log "Indexed $($knownCartonIds.Count) known carton IDs from dimension workbook."
  }

  # 2. Collect backup master carton records
  $backups = @{}
  if ($backupSheet) {
    $backupHeader = Get-HeaderInfo $backupSheet @(@('carton id','mastercarton id'), @('carton width cm','width cm'), @('carton depth cm','depth cm'), @('carton height cm','height cm'))
    if ($backupHeader) {
      $backupId = Find-HeaderColumn $backupHeader @('carton id','mastercarton id')
      $backupName = Find-HeaderColumn $backupHeader @('carton name','mastercarton name')
      $backupWidth = Find-HeaderColumn $backupHeader @('carton width cm','width cm')
      $backupDepth = Find-HeaderColumn $backupHeader @('carton depth cm','depth cm')
      $backupHeight = Find-HeaderColumn $backupHeader @('carton height cm','height cm')
      $backupLifecycle = Find-HeaderColumn $backupHeader @('lifecycle','status')
      
      for ($r = $backupHeader.Row + 1; $r -lt $backupSheet.Rows.Count; $r++) {
        $row = $backupSheet.Rows[$r]
        $cartonId = Normalize-ItemId (Get-RowCell $row $backupId)
        if (-not $cartonId) { continue }
        $w = Convert-ToFiniteNumber (Get-RowCell $row $backupWidth)
        $d = Convert-ToFiniteNumber (Get-RowCell $row $backupDepth)
        $h = Convert-ToFiniteNumber (Get-RowCell $row $backupHeight)
        $name = if ($backupName -ge 0) { Get-RowCell $row $backupName } else { "" }
        $lifecycle = if ($backupLifecycle -ge 0) { Get-RowCell $row $backupLifecycle } else { "Active" }
        
        if ([double]::IsNaN($w) -or [double]::IsNaN($d) -or [double]::IsNaN($h) -or $w -le 0 -or $d -le 0 -or $h -le 0) {
          Add-DataQuality -issue "Invalid dimensions" -entityType "Carton" -entityId $cartonId -desc "Zero or invalid backup dimensions in row $($r+1)" -detail "L=$d, W=$w, H=$h"
          continue
        }
        $backups[$cartonId] = [pscustomobject]@{
          name = $name
          length = $d * 10
          width = $w * 10
          height = $h * 10
          lifecycle = $lifecycle
          row = $r + 1
        }
      }
    }
  }

  # 3. Parse BOM carton assignments
  $bomRowsRaw = [System.Collections.ArrayList]::new()
  if ($bomFile) {
    Write-Log "Processing BOM Workbook..."
    $bomSheets = @(Read-XlsxSheets -File $bomFile -SkipSheets 1)
    
    foreach ($sheet in $bomSheets) {
      Write-Log "BOM Scan: sheet '$($sheet.Name)' ($($sheet.Rows.Count) rows)"
      
      # Try flat sheet format
      $flat = $null
      for ($hr = 0; $hr -lt [math]::Min(60, $sheet.Rows.Count); $hr++) {
        $hh = @($sheet.Rows[$hr] | ForEach-Object { Normalize-Header $_ })
        $fgIdx = -1; $fgNameIdx = -1; $itemIdx = -1; $productIdx = -1; $normIdx = -1; $quantityIdx = -1; $perSeriesIdx = -1
        for ($hc = 0; $hc -lt $hh.Count; $hc++) {
          if ($fgIdx -lt 0 -and $hh[$hc] -in @('fg','fg id','fgid','finished good','finishedgood','top item id')) { $fgIdx=$hc }
          if ($fgNameIdx -lt 0 -and $hh[$hc] -in @('fg name','fgname','finished good name','top item name')) { $fgNameIdx=$hc }
          if ($itemIdx -lt 0 -and $hh[$hc] -in @('item number','item no','itemnumber','component','component item','child item id')) { $itemIdx=$hc }
          if ($productIdx -lt 0 -and $hh[$hc] -in @('product name','product','productname','description','item description','child item name')) { $productIdx=$hc }
          if ($normIdx -lt 0 -and $hh[$hc] -in @('norm','total','child qty per','parent qty per')) { $normIdx=$hc }
          if ($quantityIdx -lt 0 -and $hh[$hc] -in @('quantity','qty','qty per','quantity per')) { $quantityIdx=$hc }
          if ($perSeriesIdx -lt 0 -and $hh[$hc] -in @('per series','perseries','series','series qty')) { $perSeriesIdx=$hc }
        }
        if ($fgIdx -ge 0 -and $itemIdx -ge 0 -and $productIdx -ge 0) {
          $flat = @{ Row=$hr; Fg=$fgIdx; FgName=$fgNameIdx; Item=$itemIdx; Product=$productIdx; Norm=$normIdx; Quantity=$quantityIdx; PerSeries=$perSeriesIdx }
          break
        }
      }
      
      if ($flat) {
        for ($r = $flat.Row + 1; $r -lt $sheet.Rows.Count; $r++) {
          $row = $sheet.Rows[$r]
          $fg = Normalize-ItemId (Get-RowCell $row $flat.Fg)
          $fgName = if ($flat.FgName -ge 0) { Get-RowCell $row $flat.FgName } else { "" }
          $item = Normalize-ItemId (Get-RowCell $row $flat.Item)
          $product = Get-RowCell $row $flat.Product
          
          if (-not $fg -or -not $item -or -not $product) { continue }
          
          $quantity = if ($flat.Quantity -ge 0) { Get-RowCell $row $flat.Quantity } else { "" }
          $norm = if ($flat.Norm -ge 0) { Get-RowCell $row $flat.Norm } else { "" }
          $perSeries = if ($flat.PerSeries -ge 0) { Get-RowCell $row $flat.PerSeries } else { "" }
          
          $knownId = $knownCartonIds.Contains($item)
          $isLabel = $product -match '(?i)carton\s*label|lith\s*label|fire\s*label|lithium\s*label|warning\s*label|rating\s*label|label\s*for|\bUPC\b|\bSCC\b|barcode'
          $isCarton = $knownId -or ($product -match '(?i)master\s*carton|mastercarton|carton\s*box|shipping\s*carton|\bcarton\b')
          
          if (-not $isCarton -or ($isLabel -and -not $knownId)) { continue }
          
          [void]$bomRowsRaw.Add([pscustomobject]@{
            fgId = $fg
            fgName = $fgName
            cartonId = $item
            cartonName = $product
            quantity = $quantity
            norm = $norm
            perSeries = $perSeries
            sourceSheet = $sheet.Name
            sourceRow = $r + 1
          })
        }
        continue
      }
      
      # Try nested structure format
      $fg = ''; $fgName = ''; $expectFg = $false; $component = $null
      for ($r = 0; $r -lt $sheet.Rows.Count; $r++) {
        $row = $sheet.Rows[$r]
        $headers = @($row | ForEach-Object { Normalize-Header $_ })
        $first = if ($headers.Count -gt 0) { $headers[0] } else { '' }
        $second = if ($headers.Count -gt 1) { $headers[1] } else { '' }
        
        if ($first -eq 'item number' -and $second -eq 'product name' -and $headers -contains 'item group') {
          $expectFg = $true
          $component = $null
          continue
        }
        if ($expectFg) {
          $candidate = Get-RowCell $row 0
          $candidateName = Get-RowCell $row 1
          if ($candidate -and $candidateName) {
            $fg = Normalize-ItemId $candidate
            $fgName = $candidateName
            $expectFg = $false
          }
          continue
        }
        if ($first -eq 'item number' -and $second -eq 'product name' -and $headers -contains 'quantity') {
          $component = @{
            Item = 0
            Product = 1
            Quantity = [array]::IndexOf($headers, 'quantity')
            Norm = [array]::IndexOf($headers, 'quantity')
            PerSeries = [array]::IndexOf($headers, 'per series')
          }
          continue
        }
        if ($first -eq 'item number' -or $first -eq 'bom') {
          $component = $null
          continue
        }
        if (-not $fg -or -not $component) { continue }
        
        $item = Normalize-ItemId (Get-RowCell $row $component.Item)
        $product = Get-RowCell $row $component.Product
        if (-not $item -or -not $product) { continue }
        
        $quantity = Get-RowCell $row $component.Quantity
        $norm = Get-RowCell $row $component.Norm
        $perSeries = Get-RowCell $row $component.PerSeries
        
        $knownId = $knownCartonIds.Contains($item)
        $isLabel = $product -match '(?i)carton\s*label|lith\s*label|fire\s*label|lithium\s*label|warning\s*label|rating\s*label|label\s*for|\bUPC\b|\bSCC\b|barcode'
        $isCarton = $knownId -or ($product -match '(?i)master\s*carton|mastercarton|carton\s*box|shipping\s*carton|\bcarton\b')
        
        if (-not $isCarton -or ($isLabel -and -not $knownId)) { continue }
        
        [void]$bomRowsRaw.Add([pscustomobject]@{
          fgId = $fg
          fgName = $fgName
          cartonId = $item
          cartonName = $product
          quantity = $quantity
          norm = $norm
          perSeries = $perSeries
          sourceSheet = $sheet.Name
          sourceRow = $r + 1
        })
      }
    }

    # Group by fgId and cartonId to identify target quantities, conflicts and build assignments list
    $groupedBoms = $bomRowsRaw | Group-Object { "$($_.fgId)|$($_.cartonId)" }
    foreach ($grp in $groupedBoms) {
      $fgId = $grp.Group[0].fgId
      $cartonId = $grp.Group[0].cartonId
      $first = $grp.Group[0]
      
      $targets = @()
      $sources = @()
      foreach ($row in $grp.Group) {
        $t = Get-TargetQuantity -perSeries $row.perSeries -norm $row.norm -quantity $row.quantity
        $src = Get-QuantitySource -perSeries $row.perSeries -norm $row.norm -quantity $row.quantity
        if ($null -ne $t) {
          $targets += $t
          $sources += $src
        }
      }
      
      $uniqueTargets = @($targets | Select-Object -Unique)
      $targetQty = if ($uniqueTargets.Count -gt 0) { $uniqueTargets[0] } else { $null }
      $qSource = if ($uniqueTargets.Count -gt 0) { $sources[0] } else { 'None' }
      
      $conflict = "false"
      $detail = ""
      if ($uniqueTargets.Count -gt 1) {
        $conflict = "true"
        $detail = "Conflicting target quantities found in BOM: $($uniqueTargets -join ', ')"
        Add-DataQuality -issue "Conflicting BOM quantities" -entityType "BOM" -entityId "$fgId|$cartonId" -desc "Multiple BOM rows specify different target quantities" -detail "$($uniqueTargets -join ', ')"
      }
      
      if ($grp.Group.Count -gt 1) {
        Add-DataQuality -issue "Duplicate assignment" -entityType "BOM" -entityId "$fgId|$cartonId" -desc "Pairing occurs $($grp.Group.Count) times in BOM" -detail "Duplicate assignment detected"
      }

      [void]$assignments.Add([pscustomobject]@{
        fg_id = $fgId
        fg_name = $first.fgName
        carton_id = $cartonId
        carton_name = $first.cartonName
        target_quantity = if ($null -eq $targetQty) { "" } else { $targetQty.ToString() }
        quantity_source = $qSource
        source_workbook = $bomFile.Name
        source_sheet = $first.sourceSheet
        source_row = $first.sourceRow
        conflict_flag = $conflict
        conflict_detail = $detail
      })
      
      if ($null -ne $targetQty) {
        [void]$reverseIndex.Add([pscustomobject]@{
          carton_id = $cartonId
          fg_id = $fgId
          fg_name = $first.fgName
          target_quantity = $targetQty
        })
      }
    }
  }

  # 4. Build FG ID -> Carton ID Map, check for conflicts, and determine assignment_status per FG
  $fgToCartonMap = @{}      # resolved: fgId -> cartonId (string)
  $fgConflicts = @{}        # conflicting: fgId -> array of cartonIds
  $fgAssignmentStatus = @{} # fgId -> "SINGLE" | "MULTIPLE_CARTONS" | "QUANTITY_CONFLICT"
  $fgQuantityConflicts = @{} # fgId -> array of { cartonId, quantities[] }

  if ($bomFile) {
    $groupedByFg = $bomRowsRaw | Group-Object fgId
    foreach ($grp in $groupedByFg) {
      $fgId = $grp.Name
      $uniqueCartons = @($grp.Group | Select-Object -ExpandProperty cartonId -Unique)
      if ($uniqueCartons.Count -eq 1) {
        $fgToCartonMap[$fgId] = $uniqueCartons[0]
      } elseif ($uniqueCartons.Count -gt 1) {
        $fgConflicts[$fgId] = $uniqueCartons
        Add-DataQuality -issue "Conflicting BOM carton assignments" -entityType "FG" -entityId $fgId -desc "FG is assigned to multiple conflicting cartons in BOM" -detail "$($uniqueCartons -join ', ')"
      }
    }
  }

  # Determine assignment_status per FG by re-evaluating assignments
  $fgCartonCounts = @{}  # fgId -> @{ cartonIds[], quantityMap }
  foreach ($assign in $assignments) {
    $fgId = $assign.fg_id
    $cartonId = $assign.carton_id
    $qty = $assign.target_quantity
    if (-not $fgCartonCounts.ContainsKey($fgId)) {
      $fgCartonCounts[$fgId] = @{ CartonIds = [System.Collections.ArrayList]::new(); CartonQtyMap = @{} }
    }
    [void]$fgCartonCounts[$fgId].CartonIds.Add($cartonId)
    if (-not $fgCartonCounts[$fgId].CartonQtyMap.ContainsKey($cartonId)) {
      $fgCartonCounts[$fgId].CartonQtyMap[$cartonId] = [System.Collections.ArrayList]::new()
    }
    if ($qty -and $qty -ne '') {
      [void]$fgCartonCounts[$fgId].CartonQtyMap[$cartonId].Add($qty)
    }
  }

  foreach ($fgId in $fgCartonCounts.Keys) {
    $info = $fgCartonCounts[$fgId]
    $uniqueCartonList = @($info.CartonIds | Select-Object -Unique)
    if ($uniqueCartonList.Count -gt 1) {
      $fgAssignmentStatus[$fgId] = "MULTIPLE_CARTONS"
    } else {
      # One unique carton - check for quantity conflicts
      $cartonId = $uniqueCartonList[0]
      $qtyList = $info.CartonQtyMap[$cartonId]
      $uniqueQtys = @($qtyList | Select-Object -Unique)
      if ($uniqueQtys.Count -gt 1) {
        $fgAssignmentStatus[$fgId] = "QUANTITY_CONFLICT"
        $fgQuantityConflicts[$fgId] = @{ cartonId = $cartonId; quantities = $uniqueQtys }
      } else {
        $fgAssignmentStatus[$fgId] = "SINGLE"
      }
    }
  }

  # Update assignment_status on each assignment row
  foreach ($assign in $assignments) {
    $fgId = $assign.fg_id
    $cartonId = $assign.carton_id
    $status = $fgAssignmentStatus[$fgId]
    if (-not $status) { $status = "SINGLE" }
    $assign | Add-Member -MemberType NoteProperty -Name 'assignment_status' -Value $status -Force

    # For MULTIPLE_CARTONS, the conflict_flag should indicate there are alternative cartons
    if ($status -eq "MULTIPLE_CARTONS") {
      $otherCartons = @($fgConflicts[$fgId] | Where-Object { $_ -ne $cartonId })
      $assign.conflict_flag = "true"
      $assign.conflict_detail = "Alternative carton assignments in latest BOM: $($otherCartons -join ', ')"
    } elseif ($status -eq "QUANTITY_CONFLICT" -and $assign.conflict_flag -ne "true") {
      $assign.conflict_flag = "true"
      $qtys = $fgQuantityConflicts[$fgId].quantities
      $assign.conflict_detail = "Conflicting BOM quantities for carton $cartonId`: $($qtys -join ', ')"
    }
  }

  # 5. Parse Dimension workbook sheet for FG (EA) and Carton (CA) measurements
  $fgMeasurements = @{}
  $cartonObservations = @{} # resolved: cartonId -> list of [pscustomobject]@{ fgId, L, W, H, row, fgName }
  $cartonDimensionSources = [System.Collections.ArrayList]::new() # Traceability list
  
  if ($dimensionSheet) {
    $dimHeader = Get-HeaderInfo $dimensionSheet @(@('item number','item no','item'), @('unit','uom'), @('width'), @('depth','length'), @('height'))
    if ($dimHeader) {
      $itemCol = Find-HeaderColumn $dimHeader @('item number','item no','item')
      $unitCol = Find-HeaderColumn $dimHeader @('unit','uom')
      $widthCol = Find-HeaderColumn $dimHeader @('width')
      $depthCol = Find-HeaderColumn $dimHeader @('depth','length')
      $heightCol = Find-HeaderColumn $dimHeader @('height')
      $nameCol = Find-HeaderColumn $dimHeader @('name','description','item description','product name','item name')
      
      for ($r = $dimHeader.Row + 1; $r -lt $dimensionSheet.Rows.Count; $r++) {
        $row = $dimensionSheet.Rows[$r]
        $itemId = Get-RowCell $row $itemCol
        $unitVal = Get-RowCell $row $unitCol
        $name = if ($nameCol -ge 0) { Get-RowCell $row $nameCol } else { "" }
        
        # Check missing FG ID
        if ([string]::IsNullOrWhiteSpace($itemId)) {
          if ($unitVal -eq 'CA' -or $unitVal -eq 'EA') {
            Add-DataQuality -issue "CA row missing FG ID" -entityType $unitVal -entityId "" -desc "Workbook row $($r+1) is missing item/FG ID" -detail ""
          }
          continue
        }
        
        $fgId = Normalize-ItemId $itemId
        $unit = (Normalize-Header $unitVal).ToUpperInvariant()
        
        if ($unit -notin @('EA','CA')) { continue }
        
        $width = Convert-ToFiniteNumber (Get-RowCell $row $widthCol)
        $depth = Convert-ToFiniteNumber (Get-RowCell $row $depthCol)
        $height = Convert-ToFiniteNumber (Get-RowCell $row $heightCol)
        
        # Validate dimensions
        $invalidDims = [double]::IsNaN($width) -or [double]::IsNaN($depth) -or [double]::IsNaN($height) -or $width -le 0 -or $depth -le 0 -or $height -le 0
        if ($invalidDims) {
          Add-DataQuality -issue "CA row has invalid dimensions" -entityType $unit -entityId $fgId -desc "Zero, negative, or invalid dimension value in workbook row $($r+1)" -detail "L=$depth, W=$width, H=$height"
          
          if ($unit -eq 'CA') {
            [void]$cartonDimensionSources.Add([pscustomobject]@{
              carton_id = ""
              source_fg_id = $fgId
              source_fg_name = $name
              measured_length_mm = if ([double]::IsNaN($depth)) { "" } else { $depth * 10 }
              measured_width_mm = if ([double]::IsNaN($width)) { "" } else { $width * 10 }
              measured_height_mm = if ([double]::IsNaN($height)) { "" } else { $height * 10 }
              source_sheet = $dimensionSheet.Name
              source_row = $r + 1
              included_in_average = "false"
              exclusion_reason = "Invalid dimensions"
            })
          }
          continue
        }
        
        $w_mm = $width * 10
        $l_mm = $depth * 10
        $h_mm = $height * 10
        
        if ($unit -eq 'EA') {
          if (-not $fgMeasurements.ContainsKey($fgId)) { $fgMeasurements[$fgId] = @() }
          $fgMeasurements[$fgId] += [pscustomobject]@{
            name = $name
            length = $l_mm
            width = $w_mm
            height = $h_mm
            row = $r + 1
          }
        } elseif ($unit -eq 'CA') {
          # Process CA row: sort descending for canonical axis order
          $sorted = @($l_mm, $w_mm, $h_mm) | Sort-Object -Descending
          $canonL = $sorted[0]
          $canonW = $sorted[1]
          $canonH = $sorted[2]
          
          # Join to BOM assignments
          if ($fgConflicts.ContainsKey($fgId)) {
            [void]$cartonDimensionSources.Add([pscustomobject]@{
              carton_id = ""
              source_fg_id = $fgId
              source_fg_name = $name
              measured_length_mm = $canonL
              measured_width_mm = $canonW
              measured_height_mm = $canonH
              source_sheet = $dimensionSheet.Name
              source_row = $r + 1
              included_in_average = "false"
              exclusion_reason = "Conflicting BOM carton assignments: $($fgConflicts[$fgId] -join ', ')"
            })
          } elseif (-not $fgToCartonMap.ContainsKey($fgId)) {
            Add-DataQuality -issue "CA measurement has no BOM carton assignment" -entityType "FG" -entityId $fgId -desc "CA row for FG $fgId has no BOM carton assignment" -detail "Workbook row $($r+1)"
            [void]$cartonDimensionSources.Add([pscustomobject]@{
              carton_id = ""
              source_fg_id = $fgId
              source_fg_name = $name
              measured_length_mm = $canonL
              measured_width_mm = $canonW
              measured_height_mm = $canonH
              source_sheet = $dimensionSheet.Name
              source_row = $r + 1
              included_in_average = "false"
              exclusion_reason = "CA measurement has no BOM carton assignment"
            })
          } else {
            $cartonId = $fgToCartonMap[$fgId]
            if (-not $cartonObservations.ContainsKey($cartonId)) { $cartonObservations[$cartonId] = @() }
            $cartonObservations[$cartonId] += [pscustomobject]@{
              fgId = $fgId
              fgName = $name
              L = $canonL
              W = $canonW
              H = $canonH
              row = $r + 1
            }
            [void]$cartonDimensionSources.Add([pscustomobject]@{
              carton_id = $cartonId
              source_fg_id = $fgId
              source_fg_name = $name
              measured_length_mm = $canonL
              measured_width_mm = $canonW
              measured_height_mm = $canonH
              source_sheet = $dimensionSheet.Name
              source_row = $r + 1
              included_in_average = "true"
              exclusion_reason = ""
            })
          }
        }
      }
    }
  }

  # 6. Build fg_dimensions.csv records
  foreach ($fgId in $fgMeasurements.Keys) {
    $mList = $fgMeasurements[$fgId]
    $mCount = $mList.Count
    $first = $mList[0]
    
    $avgL = ($mList | Measure-Object length -Average).Average
    $avgW = ($mList | Measure-Object width -Average).Average
    $avgH = ($mList | Measure-Object height -Average).Average

    $conflict = "false"
    $detail = ""
    $distinct = @($mList | Group-Object { "$($_.length)|$($_.width)|$($_.height)" })
    if ($distinct.Count -gt 1) {
      $conflict = "true"
      $detail = "Conflicting Finished Good dimensions found across UOM=EA rows."
      Add-DataQuality -issue "Conflicting FG dimensions" -entityType "FG" -entityId $fgId -desc "Multiple EA rows have different dimensions" -detail "$mCount measurements"
    }

    $allFgs[$fgId] = [pscustomobject]@{
      fg_id = $fgId
      fg_name = $first.name
      length_mm = $avgL
      width_mm = $avgW
      height_mm = $avgH
      measurement_count = $mCount
      source_workbook = $dimensionFile.Name
      source_sheet = $dimensionSheet.Name
      source_row = $first.row
      conflict_flag = $conflict
      conflict_detail = $detail
    }
  }

  # 7. Group and average carton dimensions
  $cartonIds = @($cartonObservations.Keys + $backups.Keys | Select-Object -Unique)
  foreach ($cartonId in $cartonIds) {
    $obsList = $cartonObservations[$cartonId]
    $backup = $backups[$cartonId]
    
    $mCount = 0
    $fgCount = 0
    $dimSource = "Missing"
    $sourceWorkbook = if ($dimensionFile) { $dimensionFile.Name } else { "" }
    $cName = if ($backup) { $backup.name } else { "Derived Carton $cartonId" }
    $lifecycle = if ($backup) { $backup.lifecycle } else { "Active" }
    
    $avgL = $null; $avgW = $null; $avgH = $null
    $minL = $null; $maxL = $null; $spreadL = 0
    $minW = $null; $maxW = $null; $spreadW = 0
    $minH = $null; $maxH = $null; $spreadH = 0
    $warning = "false"
    
    if ($obsList -and $obsList.Count -gt 0) {
      $mCount = $obsList.Count
      $fgCount = @($obsList | Select-Object -ExpandProperty fgId -Unique).Count
      $dimSource = "BOM_LINKED_CA_AVERAGE"
      
      $avgL = ($obsList | Measure-Object L -Average).Average
      $avgW = ($obsList | Measure-Object W -Average).Average
      $avgH = ($obsList | Measure-Object H -Average).Average
      
      $minL = ($obsList | Measure-Object L -Minimum).Minimum
      $maxL = ($obsList | Measure-Object L -Maximum).Maximum
      $minW = ($obsList | Measure-Object W -Minimum).Minimum
      $maxW = ($obsList | Measure-Object W -Maximum).Maximum
      $minH = ($obsList | Measure-Object H -Minimum).Minimum
      $maxH = ($obsList | Measure-Object H -Maximum).Maximum
      
      $spreadL = $maxL - $minL
      $spreadW = $maxW - $minW
      $spreadH = $maxH - $minH
      
      if ($mCount -gt 1) {
        $hasRangeWarning = $spreadL -gt 10 -or $spreadW -gt 10 -or $spreadH -gt 10
        $hasPctWarning = ($spreadL / $avgL -gt 0.05) -or ($spreadW / $avgW -gt 0.05) -or ($spreadH / $avgH -gt 0.05)
        if ($hasRangeWarning -or $hasPctWarning) {
          $warning = "true"
          Add-DataQuality -issue "Large carton measurement variation" -entityType "Carton" -entityId $cartonId -desc "Dimension measurements spread varies materially (L spread=$spreadL, W spread=$spreadW, H spread=$spreadH)" -detail "$mCount observations"
        }
      }
      
      if ($mCount -eq 1) {
        Add-DataQuality -issue "Single CA measurement" -entityType "Carton" -entityId $cartonId -desc "Master Carton has only one contributing CA measurement" -detail "No averaging possible"
      }
      
      $byFg = $obsList | Group-Object fgId
      if ($byFg.Count -gt 1) {
        $fgAverages = @()
        foreach ($fgrp in $byFg) {
          $fgAverages += [pscustomobject]@{
            fgId = $fgrp.Name
            L = ($fgrp.Group | Measure-Object L -Average).Average
            W = ($fgrp.Group | Measure-Object W -Average).Average
            H = ($fgrp.Group | Measure-Object H -Average).Average
          }
        }
        $maxDiffL = 0; $maxDiffW = 0; $maxDiffH = 0
        for ($i = 0; $i -lt $fgAverages.Count; $i++) {
          for ($j = $i + 1; $j -lt $fgAverages.Count; $j++) {
            $diffL = [math]::Abs($fgAverages[$i].L - $fgAverages[$j].L)
            $diffW = [math]::Abs($fgAverages[$i].W - $fgAverages[$j].W)
            $diffH = [math]::Abs($fgAverages[$i].H - $fgAverages[$j].H)
            if ($diffL -gt $maxDiffL) { $maxDiffL = $diffL }
            if ($diffW -gt $maxDiffW) { $maxDiffW = $diffW }
            if ($diffH -gt $maxDiffH) { $maxDiffH = $diffH }
          }
        }
        if ($maxDiffL -gt 5 -or $maxDiffW -gt 5 -or $maxDiffH -gt 5) {
          Add-DataQuality -issue "Material difference between FGs" -entityType "Carton" -entityId $cartonId -desc "Averaged dimensions differ materially between FGs assigned to this carton (L diff=$maxDiffL, W diff=$maxDiffW, H diff=$maxDiffH)" -detail "FGs: $($byFg.Name -join ', ')"
        }
      }
      
      if ($backup) {
        $bSorted = @($backup.length, $backup.width, $backup.height) | Sort-Object -Descending
        $diffL = [math]::Abs($avgL - $bSorted[0])
        $diffW = [math]::Abs($avgW - $bSorted[1])
        $diffH = [math]::Abs($avgH - $bSorted[2])
        if ($diffL -gt 5 -or $diffW -gt 5 -or $diffH -gt 5) {
          Add-DataQuality -issue "Backup dimension mismatch" -entityType "Carton" -entityId $cartonId -desc "Backup dimensions from Mastercarton_Dimensions disagree with CA average (L diff=$diffL, W diff=$diffW, H diff=$diffH)" -detail "CA average: $avgL x $avgW x $avgH, Backup: $($bSorted[0]) x $($bSorted[1]) x $($bSorted[2])"
        }
      }
      
    } elseif ($backup) {
      $mCount = 1
      $fgCount = 0
      $dimSource = "Mastercarton_Dimensions backup"
      $cName = $backup.name
      $lifecycle = $backup.lifecycle
      
      $sorted = @($backup.length, $backup.width, $backup.height) | Sort-Object -Descending
      $avgL = $sorted[0]; $avgW = $sorted[1]; $avgH = $sorted[2]
      $minL = $avgL; $maxL = $avgL
      $minW = $avgW; $maxW = $avgW
      $minH = $avgH; $maxH = $avgH
    }
    
    if ($null -ne $avgL) {
      $allCartons[$cartonId] = [pscustomobject]@{
        carton_id = $cartonId
        carton_name = $cName
        outside_length_mm = $avgL
        outside_width_mm = $avgW
        outside_height_mm = $avgH
        measurement_count = $mCount
        contributing_fg_count = $fgCount
        length_min_mm = $minL
        length_max_mm = $maxL
        width_min_mm = $minW
        width_max_mm = $maxW
        height_min_mm = $minH
        height_max_mm = $maxH
        length_spread_mm = $spreadL
        width_spread_mm = $spreadW
        height_spread_mm = $spreadH
        variation_warning = $warning
        dimension_source = $dimSource
        source_type = $dimSource
        source_workbook = $sourceWorkbook
        lifecycle = $lifecycle
      }
    }
  }

  # Build Data Quality Warnings
  # Look for: Missing FG dimensions, Missing carton dimensions, Current BOM carton not found
  foreach ($assign in $assignments) {
    $fgId = $assign.fg_id
    $cartonId = $assign.carton_id
    
    # 1. Missing FG dimensions
    if (-not $allFgs.ContainsKey($fgId)) {
      Add-DataQuality -issue "Missing FG dimensions" -entityType "FG" -entityId $fgId -desc "Finished Good has no packaging dimensions in dimension workbook" -detail "Missing FG EA record"
    }
    
    # 2. Missing Carton dimensions
    if (-not $allCartons.ContainsKey($cartonId) -or $allCartons[$cartonId].source_type -eq 'Missing') {
      Add-DataQuality -issue "Missing carton dimensions" -entityType "Carton" -entityId $cartonId -desc "Master Carton has no outer or backup dimensions" -detail "Missing CA or backup record"
    }
  }

  # Output CSV files to system\generated\
  Write-Log "Writing fg_dimensions.csv..."
  $fgExport = @($allFgs.Values)
  if ($fgExport.Count -eq 0) { $fgExport = @([pscustomobject]@{ fg_id=""; fg_name=""; length_mm=""; width_mm=""; height_mm=""; measurement_count=""; source_workbook=""; source_sheet=""; source_row=""; conflict_flag=""; conflict_detail="" }) }
  $fgExport | Export-Csv -Path (Join-Path $generatedDir 'fg_dimensions.csv') -NoTypeInformation -Encoding UTF8

  Write-Log "Writing carton_dimensions.csv..."
  $ctExport = @($allCartons.Values)
  if ($ctExport.Count -eq 0) { $ctExport = @([pscustomobject]@{ carton_id=""; carton_name=""; outside_length_mm=""; outside_width_mm=""; outside_height_mm=""; measurement_count=""; contributing_fg_count=""; length_min_mm=""; length_max_mm=""; width_min_mm=""; width_max_mm=""; height_min_mm=""; height_max_mm=""; length_spread_mm=""; width_spread_mm=""; height_spread_mm=""; variation_warning=""; dimension_source=""; source_type=""; source_workbook=""; lifecycle="" }) }
  $ctExport | Export-Csv -Path (Join-Path $generatedDir 'carton_dimensions.csv') -NoTypeInformation -Encoding UTF8

  Write-Log "Writing carton_dimension_sources.csv..."
  $sourcesExport = @($cartonDimensionSources)
  if ($sourcesExport.Count -eq 0) {
    $sourcesExport = @([pscustomobject]@{
      carton_id = ""
      source_fg_id = ""
      source_fg_name = ""
      measured_length_mm = ""
      measured_width_mm = ""
      measured_height_mm = ""
      source_sheet = ""
      source_row = ""
      included_in_average = ""
      exclusion_reason = ""
    })
  }
  $sourcesExport | Export-Csv -Path (Join-Path $generatedDir 'carton_dimension_sources.csv') -NoTypeInformation -Encoding UTF8

  Write-Log "Writing bom_carton_assignments.csv..."
  $assignExport = @($assignments)
  if ($assignExport.Count -eq 0) { $assignExport = @([pscustomobject]@{ fg_id=""; fg_name=""; carton_id=""; carton_name=""; target_quantity=""; quantity_source=""; source_workbook=""; source_sheet=""; source_row=""; conflict_flag=""; conflict_detail=""; assignment_status="" }) }
  $assignExport | Export-Csv -Path (Join-Path $generatedDir 'bom_carton_assignments.csv') -NoTypeInformation -Encoding UTF8

  Write-Log "Writing carton_fg_reverse_index.csv..."
  $revExport = @($reverseIndex)
  if ($revExport.Count -eq 0) { $revExport = @([pscustomobject]@{ carton_id=""; fg_id=""; fg_name=""; target_quantity="" }) }
  $revExport | Export-Csv -Path (Join-Path $generatedDir 'carton_fg_reverse_index.csv') -NoTypeInformation -Encoding UTF8

  Write-Log "Writing data_quality.csv..."
  $dqExport = @($dataQuality)
  if ($dqExport.Count -eq 0) { $dqExport = @([pscustomobject]@{ issue_type=""; entity_type=""; entity_id=""; description=""; source_detail="" }) }
  $dqExport | Export-Csv -Path (Join-Path $generatedDir 'data_quality.csv') -NoTypeInformation -Encoding UTF8

  # Generate cache_state.ini
  Write-Log "Writing cache_state.ini..."
  $dimPathVal = if ($dimensionFile) { $dimensionFile.FullName } else { "" }
  $dimNameVal = if ($dimensionFile) { $dimensionFile.Name } else { "" }
  $dimSizeVal = if ($dimensionFile) { $dimensionFile.Length } else { "" }
  $dimModVal = if ($dimensionFile) { $dimensionFile.LastWriteTimeUtc.ToString('o') } else { "" }
  
  $bomPathVal = if ($bomFile) { $bomFile.FullName } else { "" }
  $bomNameVal = if ($bomFile) { $bomFile.Name } else { "" }
  $bomSizeVal = if ($bomFile) { $bomFile.Length } else { "" }
  $bomModVal = if ($bomFile) { $bomFile.LastWriteTimeUtc.ToString('o') } else { "" }

  $stateContent = @(
    "[DimensionWorkbook]",
    "Path=$dimPathVal",
    "Filename=$dimNameVal",
    "Size=$dimSizeVal",
    "LastModifiedUtc=$dimModVal",
    "Hash=$dimHash",
    "CacheSchemaVersion=$cacheVersion",
    "PreprocessingVersion=2.0",
    "GeneratedUtc=$((Get-Date).ToUniversalTime().ToString('o'))",
    "",
    "[BomWorkbook]",
    "Path=$bomPathVal",
    "Filename=$bomNameVal",
    "Size=$bomSizeVal",
    "LastModifiedUtc=$bomModVal",
    "Hash=$bomHash",
    "CacheSchemaVersion=$cacheVersion",
    "PreprocessingVersion=2.0",
    "GeneratedUtc=$((Get-Date).ToUniversalTime().ToString('o'))"
  )
  Set-Content -LiteralPath $cacheStatePath -Value $stateContent -Encoding UTF8
  Write-Log "Preprocessing completed successfully. Derived dimensions and indices generated."
}

# Browser launching
if ($manifestsOnly) {
  Write-Log "Manifests and data files prepared. Skipping browser launch (--manifests-only)."
} else {
  $htmlPath = Join-Path $systemDir 'WVN_Master_Carton_Fit_Tool.html'
  $chromePath = Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'
  $edgePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'
  $browser = $null
  if (Test-Path -LiteralPath $chromePath) { $browser = $chromePath }
  elseif (Test-Path -LiteralPath $edgePath) { $browser = $edgePath }
  
  if ($browser) {
    $tempProfile = Join-Path $env:TEMP "WVNMasterCartonToolChrome-$(Get-Random)"
    Write-Log "Launching browser: $browser"
    Write-Log "HTML URL: $htmlPath"
    
    # Hide console window if ShowConsole=false
    $showConsole = if ($config.ContainsKey('Launcher') -and $config['Launcher'].ContainsKey('ShowConsole')) { $config['Launcher']['ShowConsole'] -eq 'true' } else { $false }
    
    if (-not $showConsole) {
      # Use a background process to open browser so console window exits quietly
      Start-Process -FilePath $browser -ArgumentList "--user-data-dir=`"$tempProfile`"", '--allow-file-access-from-files', "`"$htmlPath`""
    } else {
      # Keep console open (this command waits, or starts and pauses)
      Start-Process -FilePath $browser -ArgumentList "--user-data-dir=`"$tempProfile`"", '--allow-file-access-from-files', "`"$htmlPath`""
    }
  } else {
    Write-Log "ERROR: Google Chrome or Microsoft Edge was not found in Program Files. Please open system/WVN_Master_Carton_Fit_Tool.html manually."
  }
}

Write-Log "Done."
