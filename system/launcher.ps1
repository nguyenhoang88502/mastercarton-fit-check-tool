$ErrorActionPreference = 'Stop'

# Resolve every application path from the launcher location only.
$SystemDir = $PSScriptRoot
$AppRoot = Split-Path -Parent $SystemDir
$ToolsRoot = Split-Path -Parent $AppRoot
$WeeklyBomDir = Join-Path $ToolsRoot '1. Weekly BOM'
$GeneratedDir = Join-Path $SystemDir 'generated'
$LogDir = Join-Path $SystemDir 'logs'
$HtmlPath = Join-Path $SystemDir 'WVN_Master_Carton_Fit_Tool.html'

# Preserve the original variable names used by the preprocessing code.
$ScriptDirectory = $SystemDir
$ToolsParent = $ToolsRoot
$WeeklyBomDirectory = $WeeklyBomDir
$systemDir = $SystemDir
$projectDir = $AppRoot
$generatedDir = $GeneratedDir
$htmlPath = $HtmlPath

$manifestsOnly = $args -contains '--manifests-only'
$forceRebuildArg = $args -contains '--force-rebuild'
$verboseLogging = $args -contains '--verbose'
$logPath = $null

try {
  if (-not (Test-Path -LiteralPath $LogDir -PathType Container)) {
    New-Item -ItemType Directory -Path $LogDir -Force -ErrorAction Stop | Out-Null
  }
  $runStamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
  $logPath = Join-Path $LogDir "WVN_Master_Carton_Fit_Tool_launcher_${runStamp}_$PID.log"
} catch {
  Write-Warning "Launcher log folder is unavailable. Continuing without a file log. $($_.Exception.Message)"
}

# Write-Log is the plain-English line a manager or QA reader should see: what
# happened, which files were used, whether there is a problem. Write-Detail is
# the technical trace (hashes, raw paths, per-sheet/per-candidate mechanics);
# it is written only when --verbose (or config.ini [Launcher] Verbose=true) is
# set, so the default console output and log file both stay readable.
function Write-Log {
  param([string]$msg)
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $line = "$ts $msg"
  Write-Host $line
  if ($logPath) {
    try {
      Add-Content -LiteralPath $logPath -Value $line -ErrorAction Stop
    } catch {
      Write-Warning "Could not append to launcher log '$logPath'. Continuing. $($_.Exception.Message)"
    }
  }
}

function Write-Detail {
  param([string]$msg)
  if (-not $verboseLogging) { return }
  Write-Log "    $msg"
}

$launcherMutex = New-Object System.Threading.Mutex($false, 'Local\WVN_Master_Carton_Fit_Tool_Launcher')
$launcherMutexAcquired = $false
try {
  $launcherMutexAcquired = $launcherMutex.WaitOne(0)
} catch [System.Threading.AbandonedMutexException] {
  $launcherMutexAcquired = $true
}

if (-not $launcherMutexAcquired) {
  Write-Log "Another Master Carton Fit Tool launcher is already preparing data. Please wait for it to finish."
  $launcherMutex.Dispose()
  exit 0
}

try {

Write-Log "=== WVN Master Carton Fit Tool ==="
Write-Detail "System directory: $systemDir"
Write-Detail "Project directory: $projectDir"

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

function Move-StagedFileWithRetry {
  param(
    [Parameter(Mandatory = $true)][string]$StagedPath,
    [Parameter(Mandatory = $true)][string]$DestinationPath,
    [int]$Attempts = 6,
    [int]$DelayMilliseconds = 200
  )

  $lastError = $null
  $replaceBackupPath = Join-Path ([System.IO.Path]::GetDirectoryName($DestinationPath)) (".replace-backup-$PID-$([guid]::NewGuid().ToString('N')).tmp")
  for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
    try {
      if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) {
        [System.IO.File]::Replace($StagedPath, $DestinationPath, $replaceBackupPath, $true)
        if (Test-Path -LiteralPath $replaceBackupPath -PathType Leaf) {
          Remove-Item -LiteralPath $replaceBackupPath -Force -ErrorAction SilentlyContinue
        }
      } else {
        [System.IO.File]::Move($StagedPath, $DestinationPath)
      }
      return
    } catch {
      $lastError = $_.Exception
      if ($attempt -lt $Attempts) { Start-Sleep -Milliseconds $DelayMilliseconds }
    }
  }
  if (Test-Path -LiteralPath $replaceBackupPath -PathType Leaf) {
    Remove-Item -LiteralPath $replaceBackupPath -Force -ErrorAction SilentlyContinue
  }
  throw "Could not replace '$DestinationPath' after $Attempts attempts. $($lastError.Message)"
}

function Publish-SingleGeneratedFile {
  param(
    [Parameter(Mandatory = $true)][string]$StagedPath,
    [Parameter(Mandatory = $true)][string]$DestinationPath
  )
  if (-not (Test-Path -LiteralPath $StagedPath -PathType Leaf) -or
      (Get-Item -LiteralPath $StagedPath).Length -le 0) {
    throw "Staged output is missing or empty: $StagedPath"
  }
  Move-StagedFileWithRetry -StagedPath $StagedPath -DestinationPath $DestinationPath
}

function Build-MasterCartonPictureIndex {
  param(
    [string]$PictureFolder,
    [string]$OutputCsvPath
  )

  $header = 'carton_id,image_sequence,variant_number,filename,image_uri'
  $supportedExtensions = @('.png', '.jpg', '.jpeg', '.webp')
  $filenamePattern = '^(?<cartonId>[^_]+)_image(?<sequence>\d+)(?:_(?<variant>\d+))?$'
  $records = New-Object System.Collections.Generic.List[object]
  $totalImageFiles = 0
  $ignoredFilenames = 0

  $stagedPath = Join-Path ([System.IO.Path]::GetDirectoryName($OutputCsvPath)) (".picture-index-$PID-$([guid]::NewGuid().ToString('N')).tmp")
  try {
    if (-not (Test-Path -LiteralPath $PictureFolder -PathType Container)) {
      Write-Log "WARNING: No master-carton picture folder was found. Picture galleries will be empty until pictures are added."
      Write-Detail "Picture folder not found: $PictureFolder"
      Set-Content -LiteralPath $stagedPath -Value $header -Encoding UTF8 -ErrorAction Stop
      Publish-SingleGeneratedFile -StagedPath $stagedPath -DestinationPath $OutputCsvPath
      Write-Detail "Master-carton picture index: total image files found=0; valid indexed files=0; ignored filenames=0; unique carton IDs=0"
      return
    }

    $imageFiles = @(
      Get-ChildItem -LiteralPath $PictureFolder -File -ErrorAction Stop |
        Where-Object { $supportedExtensions -contains $_.Extension.ToLowerInvariant() }
    )
    $totalImageFiles = $imageFiles.Count

    foreach ($file in $imageFiles) {
      if ($file.BaseName -notmatch $filenamePattern) {
        $ignoredFilenames++
        continue
      }

      $parsedCartonId = $Matches.cartonId
      $imageSequence = [int]$Matches.sequence
      $parsedVariant = $Matches.variant
      $cartonId = Normalize-ItemId $parsedCartonId
      if (-not $cartonId) {
        $ignoredFilenames++
        continue
      }

      $variantNumber = 0
      if ($parsedVariant) { $variantNumber = [int]$parsedVariant }
      $fileUri = (New-Object System.Uri($file.FullName)).AbsoluteUri
      $records.Add([pscustomobject]@{
        carton_id = $cartonId
        image_sequence = $imageSequence
        variant_number = $variantNumber
        filename = $file.Name
        image_uri = $fileUri
      })
    }

    $sortedRecords = @(
      $records |
        Sort-Object `
          @{ Expression = { [string]$_.carton_id } }, `
          @{ Expression = { [int]$_.image_sequence } }, `
          @{ Expression = { [int]$_.variant_number } }, `
          @{ Expression = { [string]$_.filename } }
    )

    if ($sortedRecords.Count -gt 0) {
      $sortedRecords | Export-Csv -LiteralPath $stagedPath -NoTypeInformation -Encoding UTF8
    } else {
      Set-Content -LiteralPath $stagedPath -Value $header -Encoding UTF8
    }
    Publish-SingleGeneratedFile -StagedPath $stagedPath -DestinationPath $OutputCsvPath

    $uniqueCartonCount = @($sortedRecords | Select-Object -ExpandProperty carton_id -Unique).Count
    Write-Detail "Master-carton picture folder resolved: $PictureFolder"
    Write-Log "Found $totalImageFiles carton picture(s); $($sortedRecords.Count) indexed for $uniqueCartonCount carton(s)."
    if ($ignoredFilenames -gt 0) {
      Write-Log "WARNING: $ignoredFilenames picture file name(s) did not match the expected naming pattern and were skipped."
    }
  } catch {
    Write-Log "WARNING: Could not update the carton picture list. Keeping the previous version."
    Write-Detail "Picture index error: $($_.Exception.Message)"
  } finally {
    if (Test-Path -LiteralPath $stagedPath -PathType Leaf) {
      Remove-Item -LiteralPath $stagedPath -Force -ErrorAction SilentlyContinue
    }
  }
}

function Publish-GeneratedDataset {
  param(
    [Parameter(Mandatory = $true)][hashtable]$CsvData,
    [Parameter(Mandatory = $true)][AllowEmptyString()][string[]]$CacheStateContent,
    [Parameter(Mandatory = $true)][string]$OutputDirectory
  )

  $csvNames = @(
    'fg_dimensions.csv',
    'carton_dimensions.csv',
    'bom_carton_assignments.csv',
    'carton_fg_reverse_index.csv',
    'carton_dimension_sources.csv',
    'data_quality.csv'
  )
  $publishNames = @($csvNames + 'cache_state.ini')
  $stagingDirectory = Join-Path $OutputDirectory (".dataset-staging-$PID-$([guid]::NewGuid().ToString('N'))")
  $backupDirectory = Join-Path $stagingDirectory 'previous'
  $publishedNames = New-Object System.Collections.Generic.List[string]
  $previouslyExisted = @{}

  New-Item -ItemType Directory -Path $backupDirectory -Force -ErrorAction Stop | Out-Null
  try {
    foreach ($csvName in $csvNames) {
      if (-not $CsvData.ContainsKey($csvName)) {
        throw "No staged data was supplied for $csvName."
      }
      $stagedCsvPath = Join-Path $stagingDirectory $csvName
      @($CsvData[$csvName]) | Export-Csv -LiteralPath $stagedCsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
      $firstLine = Get-Content -LiteralPath $stagedCsvPath -TotalCount 1 -ErrorAction Stop
      if (-not $firstLine -or $firstLine -notmatch ',') {
        throw "Temporary CSV validation failed for $csvName."
      }
    }

    $stagedCachePath = Join-Path $stagingDirectory 'cache_state.ini'
    Set-Content -LiteralPath $stagedCachePath -Value $CacheStateContent -Encoding UTF8 -ErrorAction Stop
    if (-not (Select-String -LiteralPath $stagedCachePath -SimpleMatch '[DimensionWorkbook]' -Quiet) -or
        -not (Select-String -LiteralPath $stagedCachePath -SimpleMatch '[BomWorkbook]' -Quiet)) {
      throw "Temporary cache_state.ini validation failed."
    }

    # Back up every destination before changing any final file. If a locked
    # destination cannot be copied, publishing stops with the old dataset intact.
    foreach ($name in $publishNames) {
      $destination = Join-Path $OutputDirectory $name
      $previouslyExisted[$name] = Test-Path -LiteralPath $destination -PathType Leaf
      if ($previouslyExisted[$name]) {
        Copy-Item -LiteralPath $destination -Destination (Join-Path $backupDirectory $name) -Force -ErrorAction Stop
      }
    }

    foreach ($name in $publishNames) {
      $stagedPath = Join-Path $stagingDirectory $name
      $destination = Join-Path $OutputDirectory $name
      if ($name -eq 'cache_state.ini') {
        Write-Detail "Writing cache_state.ini last..."
      } else {
        Write-Detail "Publishing $name..."
      }
      Move-StagedFileWithRetry -StagedPath $stagedPath -DestinationPath $destination
      $publishedNames.Add($name)
    }
  } catch {
    $publishError = $_
    Write-Log "WARNING: Could not save the new data files. Keeping the previous complete dataset instead."
    for ($i = $publishedNames.Count - 1; $i -ge 0; $i--) {
      $name = $publishedNames[$i]
      $destination = Join-Path $OutputDirectory $name
      try {
        if ($previouslyExisted[$name]) {
          $rollbackStage = Join-Path $stagingDirectory (".rollback-$name")
          Copy-Item -LiteralPath (Join-Path $backupDirectory $name) -Destination $rollbackStage -Force -ErrorAction Stop
          Move-StagedFileWithRetry -StagedPath $rollbackStage -DestinationPath $destination
        } elseif (Test-Path -LiteralPath $destination -PathType Leaf) {
          Remove-Item -LiteralPath $destination -Force -ErrorAction Stop
        }
      } catch {
        Write-Log "ERROR: Could not restore the previous version of $name. $($_.Exception.Message)"
      }
    }
    throw $publishError
  } finally {
    $resolvedOutput = [System.IO.Path]::GetFullPath($OutputDirectory).TrimEnd('\') + '\'
    $resolvedStaging = [System.IO.Path]::GetFullPath($stagingDirectory)
    if ($resolvedStaging.StartsWith($resolvedOutput, [System.StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $resolvedStaging -PathType Container)) {
      Remove-Item -LiteralPath $resolvedStaging -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
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

function Test-FileReadAvailable {
  param([Parameter(Mandatory = $true)][System.IO.FileInfo]$File)
  $stream = $null
  try {
    if (-not $File.Exists -or $File.Length -le 0) { return $false }
    $stream = [System.IO.File]::Open(
      $File.FullName,
      [System.IO.FileMode]::Open,
      [System.IO.FileAccess]::Read,
      [System.IO.FileShare]::Read
    )
    return $true
  } catch {
    return $false
  } finally {
    if ($stream) { $stream.Dispose() }
  }
}

function Get-UnlockedFileHash {
  param([Parameter(Mandatory = $true)][System.IO.FileInfo]$File)
  if (-not (Test-FileReadAvailable -File $File)) {
    Write-Log "The workbook '$($File.Name)' became unavailable (locked) while preparing data. It will be skipped."
    Write-Detail "Locked file: $($File.FullName)"
    return $null
  }
  try {
    return (Get-FileHash -LiteralPath $File.FullName -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
  } catch {
    Write-Log "The workbook '$($File.Name)' could not be read while preparing data. It will be skipped."
    Write-Detail "Locked file: $($File.FullName). $($_.Exception.Message)"
    return $null
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
    Write-Detail "Inferred business date for '$name': $($candidateDate.ToString('yyyy-MM-dd')) (Year inferred as $mYear from LastWriteTimeUtc)"
  } else {
    $res.Error = "Filename does not match expected pattern BOM & FORMULA [MONTH] [DAY] D365"
  }
  return $res
}

# Read configuration. Folder names may be overridden via [Folders]; every
# resolved path still stays an exact relative sibling under $ToolsRoot.
$configPath = Join-Path $SystemDir 'config.ini'
try {
  $config = Read-IniFile -FilePath $configPath
} catch {
  Write-Log "WARNING: The settings file (config.ini) could not be read; using default settings. $($_.Exception.Message)"
  $config = @{}
}

if (-not $verboseLogging -and $config.ContainsKey('Launcher') -and $config['Launcher']['Verbose']) {
  $verboseLogging = $config['Launcher']['Verbose'] -eq 'true'
}

$dimPattern = '*Physical Dimension*.xlsx'
$bomPattern = 'BOM & FORMULA*D365.xlsx'
if ($config.ContainsKey('Files')) {
  if ($config['Files']['DimensionPattern']) { $dimPattern = $config['Files']['DimensionPattern'] }
  if ($config['Files']['BomPattern']) { $bomPattern = $config['Files']['BomPattern'] }
}

$forceRebuild = $forceRebuildArg
if (-not $forceRebuild -and $config.ContainsKey('Launcher') -and $config['Launcher']['ForceRebuild']) {
  $forceRebuild = $config['Launcher']['ForceRebuild'] -eq 'true'
}

if ($config.ContainsKey('Search')) {
  $legacySearchKeys = @($config['Search'].Keys | Where-Object { $config['Search'][$_] })
  if ($legacySearchKeys.Count -gt 0) {
    Write-Log "WARNING: The old [Search] settings in config.ini are no longer used and were ignored: $($legacySearchKeys -join ', ')."
  }
}

# Folder names are configurable but always resolved as exact relative
# siblings under $ToolsRoot / $WeeklyBomDir. No absolute or scanned paths.
$weeklyBomFolderName = '1. Weekly BOM'
if ($config.ContainsKey('Folders')) {
  if ($config['Folders']['WeeklyBomFolderName']) {
    $weeklyBomFolderName = $config['Folders']['WeeklyBomFolderName']
  } elseif ($config['Folders']['WeeklyBomFolder']) {
    $weeklyBomFolderName = $config['Folders']['WeeklyBomFolder']
    Write-Log "WARNING: The config.ini setting 'WeeklyBomFolder' is outdated. Please rename it to 'WeeklyBomFolderName'."
  }
}
$WeeklyBomDir = Join-Path $ToolsRoot $weeklyBomFolderName
$WeeklyBomDirectory = $WeeklyBomDir

$pictureFolderName = 'master-carton-pictures'
if ($config.ContainsKey('Folders') -and $config['Folders']['PictureFolderName']) {
  $pictureFolderName = $config['Folders']['PictureFolderName']
}

$expectedAppFolderName = 'WVN_Master_Carton_Fit_Tool'
if ((Split-Path -Leaf $AppRoot) -cne $expectedAppFolderName) {
  Write-Log "WARNING: This folder should be named '$expectedAppFolderName' (it is currently '$(Split-Path -Leaf $AppRoot)'). The tool still works, but please rename it when convenient."
}

$duplicateNestedRoot = Join-Path $AppRoot $expectedAppFolderName
if ((Test-Path -LiteralPath (Join-Path $duplicateNestedRoot 'Open_WVN_Master_Carton_Fit_Tool.cmd') -PathType Leaf) -and
    (Test-Path -LiteralPath (Join-Path $duplicateNestedRoot 'system') -PathType Container)) {
  Write-Log "WARNING: This tool appears to be nested inside a duplicate copy of itself. Keep only the folder that directly contains the CMD file and the 'system' folder."
}

if (-not (Test-Path -LiteralPath $WeeklyBomDir -PathType Container)) {
  throw @"
The folder "$weeklyBomFolderName" was not found beside the Master Carton tool.
Expected location: $WeeklyBomDir
Keep "$weeklyBomFolderName" and "$expectedAppFolderName" as sibling folders.
"@
}

$resolvedWeeklyBomDir = (Get-Item -LiteralPath $WeeklyBomDir -ErrorAction Stop).FullName
$siblingWeeklyBomDir = $resolvedWeeklyBomDir
Write-Detail "Script directory: $SystemDir"
Write-Detail "Resolved application root: $AppRoot"
Write-Detail "Resolved parent tools directory: $ToolsRoot"
Write-Detail "Weekly BOM directory: $resolvedWeeklyBomDir"
Write-Detail "Search scope: exact sibling Weekly BOM directory only."
Write-Detail "Broader profile, OneDrive, SharePoint, drive, and application-root scans are disabled."

# 1. Locate Dimension Workbook
Write-Log "Looking for the dimension workbook..."

$dimensionFile = $null
$dimCandidatesList = [System.Collections.ArrayList]::new()
$lockedDimensionFiles = [System.Collections.ArrayList]::new()

if ($resolvedWeeklyBomDir -and (Test-Path -LiteralPath $resolvedWeeklyBomDir)) {
  # 1a. Try canonical dimension filename: Products Physical Dimension - FG 1.Apr.25.xlsx
  $canonicalDimPath = Join-Path $resolvedWeeklyBomDir 'Products Physical Dimension - FG 1.Apr.25.xlsx'
  $canonicalDimFileInfo = New-Object System.IO.FileInfo -ArgumentList $canonicalDimPath

  if (Test-Path -LiteralPath $canonicalDimPath) {
    [void]$dimCandidatesList.Add($canonicalDimFileInfo)
    Write-Detail "Found canonical dimension file: $($canonicalDimFileInfo.FullName). Checking strict read availability..."
    if (-not (Test-FileReadAvailable -File $canonicalDimFileInfo)) {
      [void]$lockedDimensionFiles.Add($canonicalDimFileInfo)
      Write-Log "The workbook '$($canonicalDimFileInfo.Name)' is currently open in Excel. Skipping it and checking for another copy..."
      Write-Detail "Locked file: $($canonicalDimFileInfo.FullName)"
    } else {
      try {
      if (Test-IsValidDimensionWorkbook -File $canonicalDimFileInfo) {
        $dimensionFile = $canonicalDimFileInfo
        Write-Log "Using dimension workbook: $($dimensionFile.Name)"
        Write-Detail "Dimension workbook path: $($dimensionFile.FullName)"
      } else {
        Write-Detail "Rejected canonical dimension file: Failed structural validation."
      }
      } catch {
        Write-Detail "Rejected canonical dimension file: validation read failed. $($_.Exception.Message)"
      }
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

    Write-Detail "Found $($dimCandidates.Count) alternative local dimension candidates. Evaluating in order..."
    foreach ($candidate in $dimCandidates) {
      if (-not (Test-FileReadAvailable -File $candidate)) {
        [void]$lockedDimensionFiles.Add($candidate)
        Write-Log "The workbook '$($candidate.Name)' is currently open in Excel. Skipping it and checking for another copy..."
        Write-Detail "Locked file: $($candidate.FullName)"
        continue
      }
      try {
        if (Test-IsValidDimensionWorkbook -File $candidate) {
          $dimensionFile = $candidate
          Write-Log "Using dimension workbook: $($dimensionFile.Name)"
          Write-Detail "Dimension workbook path: $($dimensionFile.FullName) (Last Modified: $($dimensionFile.LastWriteTimeUtc))"
          break
        } else {
          Write-Detail "Rejected dimension candidate '$($candidate.Name)': Failed structural validation."
        }
      } catch {
        Write-Detail "Rejected dimension candidate '$($candidate.Name)': validation read failed. $($_.Exception.Message)"
      }
    }
  }
}

# Log findings
if ($dimCandidatesList.Count -gt 0) {
  $names = @($dimCandidatesList | ForEach-Object { $_.Name })
  Write-Detail "Dimension candidates found locally: $($names -join ', ')"
} else {
  Write-Detail "Dimension candidates found locally: [None]"
}

if ($dimensionFile) {
  Write-Detail "Selected dimension workbook: $($dimensionFile.FullName)"
} else {
  Write-Log "WARNING: No usable dimension workbook was found. All copies may be locked or missing."
}


# 2. Locate latest BOM Workbook
Write-Log "Looking for the latest weekly BOM workbook..."
$bomFile = $null
$bomCandidatesList = [System.Collections.ArrayList]::new()
$decoratedCandidates = [System.Collections.ArrayList]::new()
$lockedBomFiles = [System.Collections.ArrayList]::new()

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
      Write-Detail "Filename date could not be parsed for '$($candidate.Name)'; candidate eligible only for modification-time fallback. Reason: $($parsed.Error)"
    }
    
    [void]$decoratedCandidates.Add([pscustomobject]@{
      File                = $candidate
      Name                = $candidate.Name
      ParsedBusinessDate  = $parsed.BusinessDate
      FilenameDateValid   = $parsed.Success
      LastWriteTimeUtc    = $candidate.LastWriteTimeUtc
      Locked                = $false
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
  if (-not (Test-FileReadAvailable -File $candidate.File)) {
    $candidate.Locked = $true
    $candidate.WorkbookValid = $false
    $candidate.ValidationError = "Skipped locked file."
    [void]$lockedBomFiles.Add($candidate.File)
    Write-Log "The workbook '$($candidate.File.Name)' is currently open in Excel. Skipping it and checking for another copy..."
    Write-Detail "Locked file: $($candidate.File.FullName)"
    continue
  }
  try {
    if (Test-IsValidBomWorkbook -File $candidate.File) {
      $candidate.WorkbookValid = $true
    } else {
      $candidate.WorkbookValid = $false
      $candidate.ValidationError = "Failed structural validation."
    }
  } catch {
    $candidate.WorkbookValid = $false
    $candidate.ValidationError = "Validation read failure. Error: $($_.Exception.Message)"
  }
}

# Log findings
foreach ($candidate in $sortedCandidates) {
  Write-Detail "BOM candidate:"
  Write-Detail "  Name: $($candidate.Name)"
  if ($candidate.FilenameDateValid) {
    Write-Detail "  Parsed business date: $($candidate.ParsedBusinessDate.ToString('yyyy-MM-dd'))"
  } else {
    Write-Detail "  Parsed business date: [Unparseable]"
  }
  # Convert UTC to standard log format
  Write-Detail "  Last modified UTC: $($candidate.LastWriteTimeUtc.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))"
  if ($candidate.WorkbookValid) {
    Write-Detail "  Workbook validation: valid"
  } else {
    Write-Detail "  Workbook validation: invalid ($($candidate.ValidationError))"
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
  Write-Log "Using BOM workbook: $($bomFile.Name)"
  Write-Detail "Selection reason: $selectedReason"
} else {
  Write-Log "WARNING: No usable BOM workbook was found. All copies may be locked or missing."
}

# Generate cache freshness checks
if (-not (Test-Path -LiteralPath $GeneratedDir -PathType Container)) {
  New-Item -ItemType Directory -Path $GeneratedDir -Force -ErrorAction Stop | Out-Null
}
$generatedDir = $GeneratedDir
$cacheStatePath = Join-Path $GeneratedDir 'cache_state.ini'
$CacheSchemaVersion = "2.0"
$PreprocessingVersion = "BOM_LINKED_CA_AVERAGE_V1"

$RequiredGeneratedFiles = @(
  (Join-Path $GeneratedDir 'fg_dimensions.csv'),
  (Join-Path $GeneratedDir 'carton_dimensions.csv'),
  (Join-Path $GeneratedDir 'bom_carton_assignments.csv'),
  (Join-Path $GeneratedDir 'carton_fg_reverse_index.csv'),
  (Join-Path $GeneratedDir 'carton_dimension_sources.csv'),
  (Join-Path $GeneratedDir 'data_quality.csv')
)

$dimHash = if ($dimensionFile) { Get-UnlockedFileHash -File $dimensionFile } else { $null }
if ($dimensionFile -and -not $dimHash) {
  [void]$lockedDimensionFiles.Add($dimensionFile)
  $dimensionFile = $null
}

$bomHash = if ($bomFile) { Get-UnlockedFileHash -File $bomFile } else { $null }
if ($bomFile -and -not $bomHash) {
  [void]$lockedBomFiles.Add($bomFile)
  $bomFile = $null
}

function Test-RequiredGeneratedFiles {
  foreach ($requiredFile in $RequiredGeneratedFiles) {
    if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
      Write-Detail "Cache invalid: generated file is missing: $([System.IO.Path]::GetFileName($requiredFile))"
      return $false
    }
    try {
      if ((Get-Item -LiteralPath $requiredFile -ErrorAction Stop).Length -le 0) {
        Write-Detail "Cache invalid: generated file is empty: $([System.IO.Path]::GetFileName($requiredFile))"
        return $false
      }
    } catch {
      Write-Detail "Cache invalid: generated file is unreadable: $([System.IO.Path]::GetFileName($requiredFile))"
      return $false
    }
  }
  return $true
}

function Test-CacheSourceSection {
  param(
    [hashtable]$State,
    [string]$SectionName,
    [System.IO.FileInfo]$SelectedFile,
    [string]$SelectedHash,
    [System.Collections.IEnumerable]$LockedFiles
  )

  if (-not $State.ContainsKey($SectionName)) { return $false }
  $section = $State[$SectionName]
  if ($section['CacheSchemaVersion'] -ne $CacheSchemaVersion -or
      $section['PreprocessingVersion'] -ne $PreprocessingVersion) {
    return $false
  }

  if ($SelectedFile) {
    return (
      $section['Path'] -eq $SelectedFile.FullName -and
      $section['Size'] -eq $SelectedFile.Length.ToString() -and
      $section['LastModifiedUtc'] -eq $SelectedFile.LastWriteTimeUtc.ToString('o') -and
      $section['Hash'] -and
      $section['Hash'] -eq $SelectedHash
    )
  }

  foreach ($lockedFile in @($LockedFiles)) {
    if ($section['Path'] -eq $lockedFile.FullName -and
        $section['Size'] -eq $lockedFile.Length.ToString() -and
        $section['LastModifiedUtc'] -eq $lockedFile.LastWriteTimeUtc.ToString('o') -and
        $section['Hash']) {
      return $true
    }
  }
  return $false
}

$cacheState = @{}
$cacheReadable = $false
try {
  if (Test-Path -LiteralPath $cacheStatePath -PathType Leaf) {
    $cacheState = Read-IniFile -FilePath $cacheStatePath
    $cacheReadable = $true
  }
} catch {
  Write-Detail "Cache invalid: cache_state.ini is unreadable. $($_.Exception.Message)"
}

$generatedCacheComplete = $cacheReadable -and (Test-RequiredGeneratedFiles)
$dimCacheMatch = $generatedCacheComplete -and
  (Test-CacheSourceSection -State $cacheState -SectionName 'DimensionWorkbook' -SelectedFile $dimensionFile -SelectedHash $dimHash -LockedFiles $lockedDimensionFiles)
$bomCacheMatch = $generatedCacheComplete -and
  (Test-CacheSourceSection -State $cacheState -SectionName 'BomWorkbook' -SelectedFile $bomFile -SelectedHash $bomHash -LockedFiles $lockedBomFiles)
$cacheValid = $dimCacheMatch -and $bomCacheMatch
$sourceUnavailable = -not ($dimensionFile -and $bomFile)

if ($sourceUnavailable) {
  if ($cacheValid -and (($lockedDimensionFiles.Count -gt 0) -or ($lockedBomFiles.Count -gt 0))) {
    $rebuildNeeded = $false
    $lockedSourceNames = @(
      @($lockedDimensionFiles | Select-Object -Last 1 -ExpandProperty Name -ErrorAction SilentlyContinue) +
      @($lockedBomFiles | Select-Object -Last 1 -ExpandProperty Name -ErrorAction SilentlyContinue)
    ) | Where-Object { $_ } | Select-Object -Unique
    $lockedSourceText = if ($lockedSourceNames.Count -gt 0) { $lockedSourceNames -join ', ' } else { 'the source workbook' }
    Write-Log "The workbook '$lockedSourceText' is currently open in Excel. The tool will use the previously prepared data instead."
  } else {
    throw @"
No valid unlocked workbook is available.
Close the workbook in Excel or wait for synchronization to finish.
The previous generated cache is missing, incomplete, stale, or does not match the expected source paths.
"@
  }
} elseif ($forceRebuild) {
  $rebuildNeeded = $true
  Write-Log "A full data refresh was requested. Reprocessing the workbooks..."
} elseif ($cacheValid) {
  $rebuildNeeded = $false
  Write-Log "Data already up-to-date - using the previously prepared files."
} else {
  $rebuildNeeded = $true
  Write-Log "The prepared data is out of date. Reprocessing the workbooks..."
}

$masterCartonPictureFolder = Join-Path $resolvedWeeklyBomDir $pictureFolderName
$legacyMasterCartonPictureFolder = Join-Path $resolvedWeeklyBomDir 'Master_Carton_Pics'
if (-not (Test-Path -LiteralPath $masterCartonPictureFolder -PathType Container) -and
    (Test-Path -LiteralPath $legacyMasterCartonPictureFolder -PathType Container)) {
  Write-Log "Please rename the folder 'Master_Carton_Pics' to '$pictureFolderName' so future runs can find it automatically. Using the old folder name for this run."
  $masterCartonPictureFolder = $legacyMasterCartonPictureFolder
}
$masterCartonPictureIndexPath = Join-Path $GeneratedDir 'master_carton_picture_index.csv'
Build-MasterCartonPictureIndex -PictureFolder $masterCartonPictureFolder -OutputCsvPath $masterCartonPictureIndexPath

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
    Write-Log "Reading the dimension workbook..."
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
    Write-Log "Found $($knownCartonIds.Count) carton records in the workbook."
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
  $totalBomRowsProcessed = 0
  if ($bomFile) {
    Write-Log "Reading the BOM workbook..."
    $bomSheets = @(Read-XlsxSheets -File $bomFile -SkipSheets 1)

    foreach ($sheet in $bomSheets) {
      $totalBomRowsProcessed += $sheet.Rows.Count
      Write-Detail "BOM Scan: sheet '$($sheet.Name)' ($($sheet.Rows.Count) rows)"

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

  # Build every output in memory before publishing.
  $fgExport = @($allFgs.Values)
  if ($fgExport.Count -eq 0) { $fgExport = @([pscustomobject]@{ fg_id=""; fg_name=""; length_mm=""; width_mm=""; height_mm=""; measurement_count=""; source_workbook=""; source_sheet=""; source_row=""; conflict_flag=""; conflict_detail="" }) }
  $ctExport = @($allCartons.Values)
  if ($ctExport.Count -eq 0) { $ctExport = @([pscustomobject]@{ carton_id=""; carton_name=""; outside_length_mm=""; outside_width_mm=""; outside_height_mm=""; measurement_count=""; contributing_fg_count=""; length_min_mm=""; length_max_mm=""; width_min_mm=""; width_max_mm=""; height_min_mm=""; height_max_mm=""; length_spread_mm=""; width_spread_mm=""; height_spread_mm=""; variation_warning=""; dimension_source=""; source_type=""; source_workbook=""; lifecycle="" }) }
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
  $assignExport = @($assignments)
  if ($assignExport.Count -eq 0) { $assignExport = @([pscustomobject]@{ fg_id=""; fg_name=""; carton_id=""; carton_name=""; target_quantity=""; quantity_source=""; source_workbook=""; source_sheet=""; source_row=""; conflict_flag=""; conflict_detail=""; assignment_status="" }) }
  $revExport = @($reverseIndex)
  if ($revExport.Count -eq 0) { $revExport = @([pscustomobject]@{ carton_id=""; fg_id=""; fg_name=""; target_quantity="" }) }
  $dqExport = @($dataQuality)
  if ($dqExport.Count -eq 0) { $dqExport = @([pscustomobject]@{ issue_type=""; entity_type=""; entity_id=""; description=""; source_detail="" }) }

  # Generate cache_state.ini — write only after confirming generated outputs exist
  $csvData = @{
    'fg_dimensions.csv' = $fgExport
    'carton_dimensions.csv' = $ctExport
    'bom_carton_assignments.csv' = $assignExport
    'carton_fg_reverse_index.csv' = $revExport
    'carton_dimension_sources.csv' = $sourcesExport
    'data_quality.csv' = $dqExport
  }

  $genUtc = (Get-Date).ToUniversalTime().ToString('o')
  $stateContent = @(
    "[DimensionWorkbook]",
    "Path=$($dimensionFile.FullName)",
    "Filename=$($dimensionFile.Name)",
    "Size=$($dimensionFile.Length)",
    "LastModifiedUtc=$($dimensionFile.LastWriteTimeUtc.ToString('o'))",
    "Hash=$dimHash",
    "CacheSchemaVersion=$CacheSchemaVersion",
    "PreprocessingVersion=$PreprocessingVersion",
    "GeneratedUtc=$genUtc",
    "",
    "[BomWorkbook]",
    "Path=$($bomFile.FullName)",
    "Filename=$($bomFile.Name)",
    "Size=$($bomFile.Length)",
    "LastModifiedUtc=$($bomFile.LastWriteTimeUtc.ToString('o'))",
    "Hash=$bomHash",
    "CacheSchemaVersion=$CacheSchemaVersion",
    "PreprocessingVersion=$PreprocessingVersion",
    "GeneratedUtc=$genUtc"
  )

  Publish-GeneratedDataset -CsvData $csvData -CacheStateContent $stateContent -OutputDirectory $GeneratedDir
  Write-Detail "Cache state written successfully after all generated CSVs."
  Write-Log "Finished preparing the data."
  $summaryFgCount = $allFgs.Count
  $summaryCartonCount = $allCartons.Count
  $summaryAssignmentCount = $assignments.Count
  $summaryBomRowsText = "{0:N0}" -f $totalBomRowsProcessed
}

# ---- Stakeholder-friendly summary --------------------------------------------
# Answers the four questions a manager, QA, or process owner actually asks:
# did it work, which files were used, how much was processed, where did it go.
function Get-GeneratedCsvRecordCount {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
  try {
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 1) {
      $allBlank = $true
      foreach ($prop in $rows[0].PSObject.Properties) { if ($prop.Value) { $allBlank = $false; break } }
      if ($allBlank) { return 0 }
    }
    return $rows.Count
  } catch {
    return $null
  }
}

if (-not $rebuildNeeded) {
  $summaryFgCount = Get-GeneratedCsvRecordCount (Join-Path $GeneratedDir 'fg_dimensions.csv')
  $summaryCartonCount = Get-GeneratedCsvRecordCount (Join-Path $GeneratedDir 'carton_dimensions.csv')
  $summaryAssignmentCount = Get-GeneratedCsvRecordCount (Join-Path $GeneratedDir 'bom_carton_assignments.csv')
  $summaryBomRowsText = "not reprocessed this run (using previously prepared data)"
}

Write-Log ""
Write-Log "==================== SUMMARY ===================="
Write-Log "Result: Success"
Write-Log "Dimension workbook used: $(if ($dimensionFile) { $dimensionFile.Name } else { '(none used this run - previously prepared data was reused)' })"
Write-Log "BOM workbook used: $(if ($bomFile) { $bomFile.Name } else { '(none used this run - previously prepared data was reused)' })"
Write-Log "Finished Goods (FGs): $summaryFgCount"
Write-Log "Master Cartons: $summaryCartonCount"
Write-Log "FG-to-carton assignments: $summaryAssignmentCount"
Write-Log "BOM rows processed: $summaryBomRowsText"
Write-Log "Generated data files: $GeneratedDir"
Write-Log "==================================================="
Write-Log ""
if (-not $verboseLogging) { Write-Log "(For a detailed technical log, run with --verbose.)" }

# Browser launching
if ($manifestsOnly) {
  Write-Log "Data files are ready. Skipping the browser launch (--manifests-only)."
} else {
  $htmlPath = Join-Path $systemDir 'WVN_Master_Carton_Fit_Tool.html'
  $chromePath = Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'
  $edgePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'
  $browser = $null
  if (Test-Path -LiteralPath $chromePath) { $browser = $chromePath }
  elseif (Test-Path -LiteralPath $edgePath) { $browser = $edgePath }

  if ($browser) {
    $tempProfile = Join-Path $env:TEMP "WVNMasterCartonToolChrome-$(Get-Random)"
    Write-Log "Opening the tool in your browser..."
    Write-Detail "Launching browser: $browser"
    Write-Detail "HTML URL: $htmlPath"

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
} finally {
  if ($launcherMutexAcquired) {
    try { $launcherMutex.ReleaseMutex() } catch {}
  }
  if ($launcherMutex) { $launcherMutex.Dispose() }
}
