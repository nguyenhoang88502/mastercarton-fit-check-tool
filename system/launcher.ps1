$ErrorActionPreference = 'Stop'
$systemDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $systemDir
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

# Read configuration
$configPath = Join-Path $systemDir 'config.ini'
$config = Read-IniFile -FilePath $configPath

$searchRoots = @()
if ($config.ContainsKey('Search')) {
  for ($i = 1; $i -le 10; $i++) {
    $key = "Root$i"
    if ($config['Search'].ContainsKey($key) -and $config['Search'][$key]) {
      $searchRoots += $config['Search'][$key]
    }
  }
  $maxDepth = if ($config['Search'].ContainsKey('MaxDepth')) { [int]$config['Search']['MaxDepth'] } else { 6 }
} else {
  $maxDepth = 6
}

$dimPattern = if ($config.ContainsKey('Files') -and $config['Files'].ContainsKey('DimensionPattern')) { $config['Files']['DimensionPattern'] } else { '*Physical Dimension*.xlsx' }
$bomPattern = if ($config.ContainsKey('Files') -and $config['Files'].ContainsKey('BomPattern')) { $config['Files']['BomPattern'] } else { 'BOM & FORMULA*D365.xlsx' }

$forceRebuild = $forceRebuildArg
if (-not $forceRebuild -and $config.ContainsKey('Launcher') -and $config['Launcher'].ContainsKey('ForceRebuild')) {
  $forceRebuild = $config['Launcher']['ForceRebuild'] -eq 'true'
}

Write-Log "Search parameters: MaxDepth=$maxDepth, DimensionPattern=$dimPattern, BomPattern=$bomPattern, ForceRebuild=$forceRebuild"

# User profile detection
$username = [Environment]::UserName
$userRoot = Join-Path 'C:\Users' $username
Write-Log "User Profile Root: $userRoot"

# Construct discovery search directories
$discoveryDirs = [System.Collections.ArrayList]::new()
foreach ($root in $searchRoots) {
  if (Test-Path -LiteralPath $root) { [void]$discoveryDirs.Add($root) }
}

# OneDrive candidate roots
$odBase = Join-Path $userRoot 'OneDrive*'
$oneDriveCandidates = @(Get-ChildItem -Path $odBase -Directory -ErrorAction SilentlyContinue | Sort-Object Name)
foreach ($od in $oneDriveCandidates) {
  [void]$discoveryDirs.Add($od.FullName)
  $npiPath = Join-Path $od.FullName 'NPI Information - Documents'
  if (Test-Path -LiteralPath $npiPath) {
    [void]$discoveryDirs.Add($npiPath)
  }
}

# Likely SharePoint-synced company folders
Get-ChildItem -Path $userRoot -Directory -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -like '*WAHL*' -or $_.Name -like '*SharePoint*' } |
  ForEach-Object { [void]$discoveryDirs.Add($_.FullName) }

# Sibling and project directories
[void]$discoveryDirs.Add($projectDir)
[void]$discoveryDirs.Add((Join-Path $projectDir '1. Weekly BOM'))
$parentDir = Split-Path -Parent $projectDir
[void]$discoveryDirs.Add($parentDir)
[void]$discoveryDirs.Add((Join-Path $parentDir '1. Weekly BOM'))

# Deduplicate search paths
$uniqueSearchDirs = @()
foreach ($d in $discoveryDirs) {
  if ($d -and (Test-Path -LiteralPath $d)) {
    $normalized = (Get-Item -LiteralPath $d).FullName
    if ($normalized -notin $uniqueSearchDirs) { $uniqueSearchDirs += $normalized }
  }
}
Write-Log "Deduplicated Search Roots: $($uniqueSearchDirs -join ', ')"

# 1. Locate Dimension Workbook
Write-Log "Checking source data..."
Write-Log "Preparing carton data..."
Write-Log "Phase: Locating Dimension Workbook"
$dimensionFile = $null
foreach ($d in $uniqueSearchDirs) {
  $files = Get-ChildItem -Path $d -Filter $dimPattern -File -Recurse -Depth $maxDepth -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notlike '~$*' } |
    Sort-Object LastWriteTimeUtc -Descending
  foreach ($f in $files) {
    if (Test-IsValidDimensionWorkbook -File $f) {
      $dimensionFile = $f
      break
    }
  }
  if ($dimensionFile) { break }
}

if ($dimensionFile) {
  Write-Log "Found Dimension Workbook: $($dimensionFile.FullName)"
} else {
  Write-Log "WARNING: Dimension Workbook not found"
}

# 2. Locate latest BOM Workbook
Write-Log "Phase: Locating latest BOM Workbook"
$bomFile = $null
$bomCandidates = [System.Collections.ArrayList]::new()

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
        try { return [DateTime]::new(2026, $mon, $d) } catch {}
      }
    }
  }
  return [DateTime]::MinValue
}

foreach ($d in $uniqueSearchDirs) {
  $files = Get-ChildItem -Path $d -Filter $bomPattern -File -Recurse -Depth $maxDepth -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '*Physical Dimension*' }
  foreach ($f in $files) {
    if (Test-IsValidBomWorkbook -File $f) {
      $bomDate = Parse-BomDate $f.Name
      [void]$bomCandidates.Add([pscustomobject]@{
        File = $f
        Date = $bomDate
        Modified = $f.LastWriteTimeUtc
      })
    }
  }
}

if ($bomCandidates.Count -gt 0) {
  # Sort by parsed Date descending, then LastWriteTimeUtc descending
  $sortedBoms = @($bomCandidates | Sort-Object Date, Modified -Descending)
  $bomFile = $sortedBoms[0].File
  Write-Log "Found BOM Workbook: $($bomFile.FullName) (Parsed Date: $($sortedBoms[0].Date.ToString('yyyy-MM-dd')))"
} else {
  Write-Log "WARNING: BOM Workbook not found"
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

  if ($dimensionFile) {
    Write-Log "Processing Dimension Workbook..."
    $sheets = @(Read-XlsxSheets -File $dimensionFile)
    
    $dimensionSheet = $sheets | Where-Object Name -eq 'Dimension' | Select-Object -First 1
    $assignmentSheet = $sheets | Where-Object Name -eq 'FG_MasterCarton' | Select-Object -First 1
    $backupSheet = $sheets | Where-Object Name -eq 'Mastercarton_Dimensions' | Select-Object -First 1
    if (-not $assignmentSheet) {
      $assignmentSheet = $sheets | Where-Object { Get-HeaderInfo $_ @(@('finishgood id','finished good id','fg id'), @('mastercarton id','master carton id')) } | Select-Object -First 1
    }
    
    # Pre-index known Carton IDs from sheets that look like carton dimensions
    foreach ($sheet in $sheets) {
      if ($sheet.Name -match '(?i)master.*carton|carton.*dimension') {
        $idCol = -1
        for ($r = 0; $r -lt [math]::Min(40, $sheet.Rows.Count); $r++) {
          $headers = @($sheet.Rows[$r] | ForEach-Object { Normalize-Header $_ })
          for ($c = 0; $c -lt $headers.Count; $c++) {
            if ($headers[$c] -in @('mastercarton id','master carton id','carton id','item number')) { $idCol = $c; break }
          }
          if ($idCol -ge 0) {
            for ($i = $r + 1; $i -lt $sheet.Rows.Count; $i++) {
              $id = Normalize-ItemId (Get-RowCell $sheet.Rows[$i] $idCol)
              if ($id) { [void]$knownCartonIds.Add($id) }
            }
            break
          }
        }
      }
    }
    if ($backupSheet) {
      $backupHeader = Get-HeaderInfo $backupSheet @(@('carton id','mastercarton id'))
      if ($backupHeader) {
        $idCol = Find-HeaderColumn $backupHeader @('carton id','mastercarton id')
        for ($i = $backupHeader.Row + 1; $i -lt $backupSheet.Rows.Count; $i++) {
          $id = Normalize-ItemId (Get-RowCell $backupSheet.Rows[$i] $idCol)
          if ($id) { [void]$knownCartonIds.Add($id) }
        }
      }
    }
    Write-Log "Indexed $($knownCartonIds.Count) known carton IDs from dimension workbook."

    # Parse Dimension Sheet for FG (EA) and Carton (CA) dimensions
    if ($dimensionSheet) {
      $dimHeader = Get-HeaderInfo $dimensionSheet @(@('item number','item no','item'), @('unit','uom'), @('width'), @('depth','length'), @('height'))
      if ($dimHeader) {
        $itemCol = Find-HeaderColumn $dimHeader @('item number','item no','item')
        $unitCol = Find-HeaderColumn $dimHeader @('unit','uom')
        $widthCol = Find-HeaderColumn $dimHeader @('width')
        $depthCol = Find-HeaderColumn $dimHeader @('depth','length')
        $heightCol = Find-HeaderColumn $dimHeader @('height')
        $nameCol = Find-HeaderColumn $dimHeader @('name','description','item description','product name','item name')
        
        # Temp groups to collect multi-measurements
        $fgMeasurements = @{}
        $cartonObservations = @{} # carton_id -> list of raw dimensions from CA rows

        # Assign carton observations map using FG_MasterCarton assignments
        $fgToCartonMap = @{}
        if ($assignmentSheet) {
          $assignHeader = Get-HeaderInfo $assignmentSheet @(@('finishgood id','finished good id','fg id'), @('mastercarton id','master carton id'))
          if ($assignHeader) {
            $fgCol = Find-HeaderColumn $assignHeader @('finishgood id','finished good id','fg id')
            $mcCol = Find-HeaderColumn $assignHeader @('mastercarton id','master carton id')
            for ($r = $assignHeader.Row + 1; $r -lt $assignmentSheet.Rows.Count; $r++) {
              $fgId = Normalize-ItemId (Get-RowCell $assignmentSheet.Rows[$r] $fgCol)
              $mcId = Normalize-ItemId (Get-RowCell $assignmentSheet.Rows[$r] $mcCol)
              if ($fgId -and $mcId) {
                if (-not $fgToCartonMap.ContainsKey($fgId)) { $fgToCartonMap[$fgId] = @() }
                $fgToCartonMap[$fgId] += $mcId
              }
            }
          }
        }

        for ($r = $dimHeader.Row + 1; $r -lt $dimensionSheet.Rows.Count; $r++) {
          $row = $dimensionSheet.Rows[$r]
          $itemId = Normalize-ItemId (Get-RowCell $row $itemCol)
          $unit = (Get-RowCell $row $unitCol).ToUpperInvariant()
          $name = if ($nameCol -ge 0) { Get-RowCell $row $nameCol } else { "" }
          
          if (-not $itemId -or $unit -notin @('EA','CA')) { continue }
          
          $width = Convert-ToFiniteNumber (Get-RowCell $row $widthCol)
          $depth = Convert-ToFiniteNumber (Get-RowCell $row $depthCol)
          $height = Convert-ToFiniteNumber (Get-RowCell $row $heightCol)
          
          if ([double]::IsNaN($width) -or [double]::IsNaN($depth) -or [double]::IsNaN($height) -or $width -le 0 -or $depth -le 0 -or $height -le 0) {
            Add-DataQuality -issue "Invalid dimensions" -entityType $unit -entityId $itemId -desc "Zero, negative, or invalid dimension value in workbook row $($r+1)" -detail "L=$depth, W=$width, H=$height"
            continue
          }

          # Normalize to mm (* 10)
          $w_mm = $width * 10
          $l_mm = $depth * 10
          $h_mm = $height * 10

          if ($unit -eq 'EA') {
            if (-not $fgMeasurements.ContainsKey($itemId)) { $fgMeasurements[$itemId] = @() }
            $fgMeasurements[$itemId] += [pscustomobject]@{
              name = $name
              length = $l_mm
              width = $w_mm
              height = $h_mm
              row = $r + 1
            }
          } elseif ($unit -eq 'CA') {
            # This is carton dimension observation from the FG UOM = CA.
            # We map this to whatever carton ID(s) this FG is assigned to.
            if ($fgToCartonMap.ContainsKey($itemId)) {
              foreach ($mcId in $fgToCartonMap[$itemId]) {
                if (-not $cartonObservations.ContainsKey($mcId)) { $cartonObservations[$mcId] = @() }
                $cartonObservations[$mcId] += [pscustomobject]@{
                  length = $l_mm
                  width = $w_mm
                  height = $h_mm
                  source_fg = $itemId
                  row = $r + 1
                }
              }
            }
          }
        }

        # 1. Build fg_dimensions.csv records
        foreach ($fgId in $fgMeasurements.Keys) {
          $mList = $fgMeasurements[$fgId]
          $mCount = $mList.Count
          $first = $mList[0]
          
          # Compute average
          $avgL = ($mList | Measure-Object length -Average).Average
          $avgW = ($mList | Measure-Object width -Average).Average
          $avgH = ($mList | Measure-Object height -Average).Average

          # Check for conflicts
          $conflict = "false"
          $detail = ""
          $distinct = @($mList | Group-Object { "$($_.length)|$($_.width)|$($_.height)" })
          if ($distinct.Count -gt 1) {
            $conflict = "true"
            $detail = "Conflicting Finished Good dimensions found across multiple UOM=EA rows."
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

        # 3. Resolve Carton canonical dimensions
        # Combine all unique carton IDs
        $cartonIds = @($cartonObservations.Keys + $backups.Keys | Select-Object -Unique)
        foreach ($cartonId in $cartonIds) {
          $obsList = $cartonObservations[$cartonId]
          $backup = $backups[$cartonId]
          
          $mCount = 0
          $sourceType = ""
          $sourceSheet = ""
          $sourceRow = 0
          $cName = ""
          $lifecycle = "Active"
          
          # Canonical sorted measurements collector
          $alignedMeasurements = @()
          
          if ($obsList -and $obsList.Count -gt 0) {
            $mCount = $obsList.Count
            $sourceType = "Dimension CA average"
            $sourceSheet = $dimensionSheet.Name
            $sourceRow = $obsList[0].row
            $cName = if ($backup) { $backup.name } else { "Derived Carton $cartonId" }
            $lifecycle = if ($backup) { $backup.lifecycle } else { "Active" }
            
            foreach ($obs in $obsList) {
              # Align dimensions consistently by sorting them largest to smallest
              $sorted = @($obs.length, $obs.width, $obs.height) | Sort-Object -Descending
              $alignedMeasurements += [pscustomobject]@{ L = $sorted[0]; W = $sorted[1]; H = $sorted[2] }
            }
          } elseif ($backup) {
            $mCount = 1
            $sourceType = "Mastercarton_Dimensions backup"
            $sourceSheet = $backupSheet.Name
            $sourceRow = $backup.row
            $cName = $backup.name
            $lifecycle = $backup.lifecycle
            
            $sorted = @($backup.length, $backup.width, $backup.height) | Sort-Object -Descending
            $alignedMeasurements += [pscustomobject]@{ L = $sorted[0]; W = $sorted[1]; H = $sorted[2] }
          }
          
          if ($alignedMeasurements.Count -gt 0) {
            # Compute stats
            $avgL = ($alignedMeasurements | Measure-Object L -Average).Average
            $avgW = ($alignedMeasurements | Measure-Object W -Average).Average
            $avgH = ($alignedMeasurements | Measure-Object H -Average).Average
            
            $minL = ($alignedMeasurements | Measure-Object L -Minimum).Minimum
            $maxL = ($alignedMeasurements | Measure-Object L -Maximum).Maximum
            $minW = ($alignedMeasurements | Measure-Object W -Minimum).Minimum
            $maxW = ($alignedMeasurements | Measure-Object W -Maximum).Maximum
            $minH = ($alignedMeasurements | Measure-Object H -Minimum).Minimum
            $maxH = ($alignedMeasurements | Measure-Object H -Maximum).Maximum
            
            $rangeL = $maxL - $minL
            $rangeW = $maxW - $minW
            $rangeH = $maxH - $minH
            
            # Check variation
            $warning = "false"
            if ($alignedMeasurements.Count -gt 1) {
              $hasRangeWarning = $rangeL -gt 10 -or $rangeW -gt 10 -or $rangeH -gt 10
              $hasPctWarning = ($rangeL / $avgL -gt 0.05) -or ($rangeW / $avgW -gt 0.05) -or ($rangeH / $avgH -gt 0.05)
              if ($hasRangeWarning -or $hasPctWarning) {
                $warning = "true"
                Add-DataQuality -issue "Large carton measurement variation" -entityType "Carton" -entityId $cartonId -desc "Dimension measurements spread varies materially (L spread=$rangeL, W spread=$rangeW, H spread=$rangeH)" -detail "$mCount observations"
              }
            }
            
            $allCartons[$cartonId] = [pscustomobject]@{
              carton_id = $cartonId
              carton_name = $cName
              outside_length_mm = $avgL
              outside_width_mm = $avgW
              outside_height_mm = $avgH
              measurement_count = $mCount
              length_min_mm = $minL
              length_max_mm = $maxL
              width_min_mm = $minW
              width_max_mm = $maxW
              height_min_mm = $minH
              height_max_mm = $maxH
              variation_warning = $warning
              source_type = $sourceType
              lifecycle = $lifecycle
            }
          }
        }
      }
    }
  }

  # Process BOM Workbook
  if ($bomFile) {
    Write-Log "Processing BOM Workbook..."
    $sheets = @(Read-XlsxSheets -File $bomFile -SkipSheets 1)
    
    $bomRowsRaw = [System.Collections.ArrayList]::new()
    
    foreach ($sheet in $sheets) {
      Write-Log "BOM Scan: sheet '$($sheet.Name)' ($($sheet.Rows.Count) rows)"
      
      # 1. Try flat sheet format
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
          
          # Criteria checks
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
      
      # 2. Try nested structure format
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

    # Process and deduplicate BOM assignments
    # Group by fgId and cartonId to identify target quantities and conflict checks
    $groupedBoms = $bomRowsRaw | Group-Object { "$($_.fgId)|$($_.cartonId)" }
    foreach ($grp in $groupedBoms) {
      $fgId = $grp.Group[0].fgId
      $cartonId = $grp.Group[0].cartonId
      $first = $grp.Group[0]
      
      # Extract target quantities for all occurrences
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
      
      # Duplicate assignment check
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
      
      # Reverse index entry
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
  if ($ctExport.Count -eq 0) { $ctExport = @([pscustomobject]@{ carton_id=""; carton_name=""; outside_length_mm=""; outside_width_mm=""; outside_height_mm=""; measurement_count=""; length_min_mm=""; length_max_mm=""; width_min_mm=""; width_max_mm=""; height_min_mm=""; height_max_mm=""; variation_warning=""; source_type=""; lifecycle="" }) }
  $ctExport | Export-Csv -Path (Join-Path $generatedDir 'carton_dimensions.csv') -NoTypeInformation -Encoding UTF8

  Write-Log "Writing bom_carton_assignments.csv..."
  $assignExport = @($assignments)
  if ($assignExport.Count -eq 0) { $assignExport = @([pscustomobject]@{ fg_id=""; fg_name=""; carton_id=""; carton_name=""; target_quantity=""; quantity_source=""; source_workbook=""; source_sheet=""; source_row=""; conflict_flag=""; conflict_detail="" }) }
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
  $stateContent = @(
    "[DimensionWorkbook]",
    "Path=$($dimensionFile.FullName)",
    "Filename=$($dimensionFile.Name)",
    "Size=$($dimensionFile.Length)",
    "LastModifiedUtc=$($dimensionFile.LastWriteTimeUtc.ToString('o'))",
    "Hash=$dimHash",
    "CacheSchemaVersion=$cacheVersion",
    "PreprocessingVersion=1.0",
    "GeneratedUtc=$((Get-Date).ToUniversalTime().ToString('o'))",
    "",
    "[BomWorkbook]",
    "Path=$($bomFile.FullName)",
    "Filename=$($bomFile.Name)",
    "Size=$($bomFile.Length)",
    "LastModifiedUtc=$($bomFile.LastWriteTimeUtc.ToString('o'))",
    "Hash=$bomHash",
    "CacheSchemaVersion=$cacheVersion",
    "PreprocessingVersion=1.0",
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
