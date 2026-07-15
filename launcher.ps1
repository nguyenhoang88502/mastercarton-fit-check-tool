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

function Get-KnownCartonIds {
  param([System.IO.FileInfo]$DimensionFile)
  $known = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  if (-not $DimensionFile) { return $known }
  Write-Log 'Preprocessing: reading known Master Carton IDs from dimension workbook'
  foreach ($sheet in @(Read-XlsxSheets -File $DimensionFile -SheetNamePattern '(?i)master.*carton|carton.*dimension')) {
    $idColumn = -1
    for ($r = 0; $r -lt [math]::Min(40, $sheet.Rows.Count); $r++) {
      $headers = @($sheet.Rows[$r] | ForEach-Object { Normalize-Header $_ })
      for ($c = 0; $c -lt $headers.Count; $c++) {
        if ($headers[$c] -in @('mastercarton id','master carton id','carton id','item number')) { $idColumn = $c; break }
      }
      if ($idColumn -ge 0) {
        for ($i = $r + 1; $i -lt $sheet.Rows.Count; $i++) { $id = Get-RowCell $sheet.Rows[$i] $idColumn; if ($id) { [void]$known.Add($id) } }
        break
      }
    }
  }
  Write-Log "Preprocessing: $($known.Count) known Master Carton IDs indexed"
  return $known
}

function Write-BomCartonIndex {
  param([System.IO.FileInfo]$BomFile, [System.IO.FileInfo]$DimensionFile, [string]$OutputPath)
  $version = 1
  if (Test-Path -LiteralPath $OutputPath) {
    try {
      $cached = Get-Content -Raw -LiteralPath $OutputPath | ConvertFrom-Json
      if ($cached.bomCartonIndexVersion -eq $version -and $cached.sourceFile -eq $BomFile.Name -and $cached.sourceModifiedUtc -eq $BomFile.LastWriteTimeUtc.ToString('o')) {
        Write-Log "Preprocessing: compact BOM cache is fresh ($($cached.assignments.Count) assignments)"
        return
      }
    } catch { Write-Log 'Preprocessing: existing compact BOM cache is invalid; rebuilding' }
  }
  Write-Log "Preprocessing: reading BOM workbook $($BomFile.Name)"
  $known = Get-KnownCartonIds -DimensionFile $DimensionFile
  Write-Log 'Preprocessing phase: Filtering carton rows'
  $all = [System.Collections.ArrayList]::new()
  foreach ($sheet in @(Read-XlsxSheets -File $BomFile -SkipSheets 1)) {
    Write-Log "Preprocessing: scanning sheet '$($sheet.Name)' ($($sheet.Rows.Count) rows)"
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
      if ($fgIdx -ge 0 -and $itemIdx -ge 0 -and $productIdx -ge 0) { $flat=@{ Row=$hr; Fg=$fgIdx; FgName=$fgNameIdx; Item=$itemIdx; Product=$productIdx; Norm=$normIdx; Quantity=$quantityIdx; PerSeries=$perSeriesIdx }; break }
    }
    if ($flat) {
      for ($r = $flat.Row + 1; $r -lt $sheet.Rows.Count; $r++) {
        $row=$sheet.Rows[$r]; $fg=Get-RowCell $row $flat.Fg; $fgName=Get-RowCell $row $flat.FgName; $item=Get-RowCell $row $flat.Item; $product=Get-RowCell $row $flat.Product
        if (-not $fg -or -not $item -or -not $product) { continue }
        $quantity=Get-RowCell $row $flat.Quantity; $norm=Get-RowCell $row $flat.Norm; $perSeries=Get-RowCell $row $flat.PerSeries
        $knownId=$known.Contains($item); $isLabel=$product -match '(?i)carton\s*label|lith\s*label|fire\s*label|lithium\s*label|warning\s*label|rating\s*label|label\s*for|\bUPC\b|\bSCC\b|barcode'
        $isCarton=$knownId -or ($product -match '(?i)master\s*carton|mastercarton|carton\s*box|shipping\s*carton|\bcarton\b')
        if (-not $isCarton -or ($isLabel -and -not $knownId)) { continue }
        $target=Get-TargetQuantity $perSeries $norm $quantity
        [void]$all.Add([pscustomobject]@{ fgId=$fg; fgName=$fgName; cartonId=$item; cartonName=$product; quantity=$quantity; norm=$norm; perSeries=$perSeries; targetQuantity=$target; sourceSheet=$sheet.Name; sourceRow=($r+1) })
      }
      continue
    }
    $fg = ''; $fgName = ''; $expectFg = $false; $component = $null
    for ($r = 0; $r -lt $sheet.Rows.Count; $r++) {
      $row = $sheet.Rows[$r]
      $headers = @($row | ForEach-Object { Normalize-Header $_ })
      $first = if ($headers.Count -gt 0) { $headers[0] } else { '' }
      $second = if ($headers.Count -gt 1) { $headers[1] } else { '' }
      if ($first -eq 'item number' -and $second -eq 'product name' -and $headers -contains 'item group') { $expectFg = $true; $component = $null; continue }
      if ($expectFg) { $candidate = Get-RowCell $row 0; $candidateName = Get-RowCell $row 1; if ($candidate -and $candidateName) { $fg=$candidate; $fgName=$candidateName; $expectFg=$false }; continue }
      if ($first -eq 'item number' -and $second -eq 'product name' -and $headers -contains 'quantity') {
        $component = @{ Item=0; Product=1; Quantity=[array]::IndexOf($headers,'quantity'); Norm=[array]::IndexOf($headers,'quantity'); PerSeries=[array]::IndexOf($headers,'per series') }
        continue
      }
      if ($first -eq 'item number' -or $first -eq 'bom') { $component = $null; continue }
      if (-not $fg -or -not $component) { continue }
      $item = Get-RowCell $row $component.Item; $product = Get-RowCell $row $component.Product
      if (-not $item -or -not $product) { continue }
      $quantity = Get-RowCell $row $component.Quantity; $norm = Get-RowCell $row $component.Norm; $perSeries = Get-RowCell $row $component.PerSeries
      $knownId = $known.Contains($item); $isLabel = $product -match '(?i)carton\s*label|lith\s*label|fire\s*label|lithium\s*label|warning\s*label|rating\s*label|label\s*for|\bUPC\b|\bSCC\b|barcode'
      $isCarton = $knownId -or ($product -match '(?i)master\s*carton|mastercarton|carton\s*box|shipping\s*carton|\bcarton\b')
      if (-not $isCarton -or ($isLabel -and -not $knownId)) { continue }
      $target = Get-TargetQuantity $perSeries $norm $quantity
      [void]$all.Add([pscustomobject]@{ fgId=$fg; fgName=$fgName; cartonId=$item; cartonName=$product; quantity=$quantity; norm=$norm; perSeries=$perSeries; targetQuantity=$target; sourceSheet=$sheet.Name; sourceRow=($r+1) })
    }
  }
  $assignments = @($all | Group-Object { "$($_.fgId)|$($_.cartonId)|$($_.targetQuantity)" } | ForEach-Object { $_.Group[0] })
  Write-Log 'Preprocessing phase: Building FG-to-carton index'
  $reverse = [ordered]@{}; $fgIndex = [ordered]@{}
  foreach ($a in $assignments) {
    if (-not $fgIndex.Contains($a.fgId)) { $fgIndex[$a.fgId] = @() }; $fgIndex[$a.fgId] += $a.cartonId
    if (-not $reverse.Contains($a.cartonId)) { $reverse[$a.cartonId] = @() }; $reverse[$a.cartonId] += $a.fgId
  }
  foreach ($key in @($fgIndex.Keys)) { $fgIndex[$key] = @($fgIndex[$key] | Select-Object -Unique) }
  Write-Log 'Preprocessing phase: Building carton-to-FG reverse index'
  foreach ($key in @($reverse.Keys)) { $reverse[$key] = @($reverse[$key] | Select-Object -Unique) }
  $conflicts = @($assignments | Group-Object { "$($_.fgId)|$($_.cartonId)" } | Where-Object { @($_.Group.targetQuantity | Where-Object { $null -ne $_ } | Select-Object -Unique).Count -gt 1 } | ForEach-Object { [pscustomobject]@{ fgId=$_.Group[0].fgId; cartonId=$_.Group[0].cartonId; targetQuantities=@($_.Group.targetQuantity | Select-Object -Unique) } })
  $payload = [ordered]@{ bomCartonIndexVersion=$version; sourceFile=$BomFile.Name; sourceUrl=([uri]$BomFile.FullName).AbsoluteUri; sourceModifiedUtc=$BomFile.LastWriteTimeUtc.ToString('o'); generatedUtc=(Get-Date).ToUniversalTime().ToString('o'); rowCount=$assignments.Count; assignments=$assignments; fgIndex=$fgIndex; reverseIndex=$reverse; conflicts=$conflicts }
  Write-Log 'Preprocessing phase: Writing compact JSON'
  $payload | ConvertTo-Json -Depth 7 -Compress | Set-Content -Encoding UTF8 -LiteralPath $OutputPath
  Write-Log "Preprocessing complete: $($assignments.Count) carton assignments, $($fgIndex.Count) FGs, $($reverse.Count) cartons, $($conflicts.Count) conflicts"
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
  Write-Log 'Phase: Locating latest BOM'
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
        _fullPath = $f.FullName
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

  if ($bomEntries.Count -gt 0) {
    $selectedBomFile = Get-Item -LiteralPath $bomEntries[0]._fullPath
    $compactPath = Join-Path $tool 'Weekly BOM Carton Index.json'
    try {
      Write-BomCartonIndex -BomFile $selectedBomFile -DimensionFile $dimFile -OutputPath $compactPath
    } catch {
      $preprocessError = $_.Exception.Message
      Write-Log "WARNING: Compact BOM preprocessing failed: $preprocessError"
      [ordered]@{ bomCartonIndexVersion=1; generatedUtc=(Get-Date).ToUniversalTime().ToString('o'); sourceFile=$selectedBomFile.Name; sourceModifiedUtc=$selectedBomFile.LastWriteTimeUtc.ToString('o'); assignments=@(); error=$preprocessError } | ConvertTo-Json -Depth 4 -Compress | Set-Content -Encoding UTF8 -LiteralPath $compactPath
      Write-Log 'Continuing with browser-side Weekly BOM fallback.'
    }
  } else {
    Write-Log 'WARNING: Compact BOM index was not generated because no Weekly BOM workbook was found.'
  }

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
      Write-Log 'Phase: Launching application'
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
