<#
.SYNOPSIS
  Build dist/SEC_XBRL_Addin.xlam from modules/*.bas + dependencies/JsonConverter.bas
  + customUI/customUI14.xml using Excel COM automation.

.DESCRIPTION
  Replaces the unreliable Linux-only vba_purge_final.py pipeline.
  Produces a clean, normally-formed xlam that any version of Excel will load.

  REQUIREMENTS
    - Windows + Microsoft Excel installed (any version 2010+)
    - "Trust access to the VBA project object model" enabled in Excel:
        File > Options > Trust Center > Trust Center Settings >
        Macro Settings > "Trust access to the VBA project object model"
      (This script auto-detects and prints a clear error if not enabled.)

  USAGE
    powershell -ExecutionPolicy Bypass -File scripts/build_xlam.ps1
#>
param(
  [switch]$Verbose
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$modulesDir = Join-Path $repoRoot 'modules'
$depsDir = Join-Path $repoRoot 'dependencies'
$customUIPath = Join-Path $repoRoot 'customUI/customUI14.xml'
$distDir = Join-Path $repoRoot 'dist'
$outXlam = Join-Path $distDir 'SEC_XBRL_Addin.xlam'

# Module import order is significant only insofar as ThisWorkbook must come first
# (it's a document module). All standard modules can be imported in any order.
$standardModules = @(
  'modConfig.bas',
  'modProgress.bas',
  'modHTTP.bas',
  'modTickerLookup.bas',
  'modJSONParser.bas',
  'modClassifier.bas',
  'modExcelWriter.bas',
  'modRibbon.bas',
  'modMain.bas'
)
$thisWorkbookCls = 'ThisWorkbook.cls'
$jsonConverter = 'JsonConverter.bas'

Write-Host '=== SEC EDGAR XBRL Add-in - Build ===' -ForegroundColor Cyan
Write-Host "Repo root:    $repoRoot"
Write-Host "Output:       $outXlam"
Write-Host ''

# ---------------------------------------------------------------------------
# 1. Verify all source files exist
# ---------------------------------------------------------------------------
$missing = @()
foreach ($name in $standardModules + @($thisWorkbookCls)) {
  $p = Join-Path $modulesDir $name
  if (-not (Test-Path $p)) { $missing += $p }
}
$jcPath = Join-Path $depsDir $jsonConverter
if (-not (Test-Path $jcPath)) { $missing += $jcPath }
if (-not (Test-Path $customUIPath)) { $missing += $customUIPath }
if ($missing.Count -gt 0) {
  Write-Host 'ERROR: missing source files:' -ForegroundColor Red
  $missing | ForEach-Object { Write-Host "  $_" }
  exit 1
}

# Reject any non-ASCII bytes in module sources. VBA's import treats .bas/.cls
# files as ANSI/Windows-1252; UTF-8 multi-byte chars (em-dashes, smart quotes,
# etc.) get re-interpreted as garbage and can corrupt comments or, worse,
# string/identifier tokens. Forcing ASCII keeps the import deterministic.
$asciiCheckPaths = @()
foreach ($n in $standardModules + @($thisWorkbookCls)) { $asciiCheckPaths += Join-Path $modulesDir $n }
$asciiCheckPaths += $jcPath

$asciiBad = @()
foreach ($p in $asciiCheckPaths) {
  if (-not (Test-Path $p)) { continue }
  $bytes = [System.IO.File]::ReadAllBytes($p)
  for ($i=0; $i -lt $bytes.Length; $i++) {
    if ($bytes[$i] -gt 127) {
      $lineNum = 1
      for ($j=0; $j -lt $i; $j++) { if ($bytes[$j] -eq 10) { $lineNum++ } }
      $asciiBad += ("  {0} line {1} : non-ASCII byte 0x{2:x2}" -f $p, $lineNum, $bytes[$i])
      break  # one finding per file is enough
    }
  }
}
if ($asciiBad.Count -gt 0) {
  Write-Host 'ERROR: source files contain non-ASCII bytes (will corrupt VBA import):' -ForegroundColor Red
  $asciiBad | ForEach-Object { Write-Host $_ }
  Write-Host '  Replace em-dashes / smart-quotes / other Unicode with plain ASCII.'
  exit 1
}
Write-Host '  ASCII-only source check: OK' -ForegroundColor Green

# ---------------------------------------------------------------------------
# 2. Verify VBOM trust setting
# ---------------------------------------------------------------------------
$accessVBOM = $null
foreach ($ver in @('16.0','15.0','14.0')) {
  $key = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
  if (Test-Path $key) {
    $v = (Get-ItemProperty -Path $key -Name 'AccessVBOM' -ErrorAction SilentlyContinue).AccessVBOM
    if ($null -ne $v) { $accessVBOM = $v; break }
  }
}
if ($accessVBOM -ne 1) {
  Write-Host 'ERROR: "Trust access to the VBA project object model" must be enabled.' -ForegroundColor Red
  Write-Host '  In Excel: File > Options > Trust Center > Trust Center Settings >'
  Write-Host '            Macro Settings > check "Trust access to the VBA project object model"'
  exit 1
}
Write-Host '  AccessVBOM trusted: OK' -ForegroundColor Green

# ---------------------------------------------------------------------------
# 3. Read customUI XML (we will inject it into the xlam ZIP at the end)
# ---------------------------------------------------------------------------
$customUIBytes = [System.IO.File]::ReadAllBytes($customUIPath)
Write-Host "  customUI14.xml read: $($customUIBytes.Length) bytes" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 4. Launch Excel and build the workbook
# ---------------------------------------------------------------------------
$tempXlam = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "sec_xbrl_build_$([System.Guid]::NewGuid().ToString('N')).xlam")

$excel = $null
$wb = $null
try {
  Write-Host ''
  Write-Host '  Starting Excel...'
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable - no macros run during build

  Write-Host '  Creating new workbook...'
  $wb = $excel.Workbooks.Add()

  # Remove default Sheet2/3/etc, leave only one sheet (Sheet1)
  while ($wb.Sheets.Count -gt 1) {
    $wb.Sheets.Item($wb.Sheets.Count).Delete()
  }

  # Pre-set the title so the xlam name is sensible
  $wb.Title = 'SEC EDGAR XBRL Add-in'

  # ----- 4a. Replace ThisWorkbook code (cannot import a document module by file) -----
  Write-Host '  Importing ThisWorkbook code...'
  $thisWbCode = Get-Content -Path (Join-Path $modulesDir $thisWorkbookCls) -Raw
  # Strip the .cls header (VERSION ... Attribute lines) so we just have the module body.
  # The .cls file looks like:
  #   VERSION 1.0 CLASS
  #   BEGIN
  #     MultiUse = -1  'True
  #   END
  #   Attribute VB_Name = "ThisWorkbook"
  #   Attribute VB_GlobalNameSpace = False
  #   Attribute VB_Creatable = False
  #   Attribute VB_PredeclaredId = True
  #   Attribute VB_Exposed = True
  #   Option Explicit
  #   ...code...
  $lines = $thisWbCode -split "`r?`n"
  $bodyStart = 0
  for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -notmatch '^\s*(VERSION|BEGIN|END|MultiUse|Attribute)\b' -and
        $lines[$i].Trim() -ne '') {
      $bodyStart = $i; break
    }
  }
  $body = ($lines[$bodyStart..($lines.Count-1)]) -join "`r`n"

  $vbProj = $wb.VBProject
  $thisWbComp = $vbProj.VBComponents.Item('ThisWorkbook')
  # Clear existing code
  if ($thisWbComp.CodeModule.CountOfLines -gt 0) {
    $thisWbComp.CodeModule.DeleteLines(1, $thisWbComp.CodeModule.CountOfLines)
  }
  $thisWbComp.CodeModule.AddFromString($body)

  # ----- 4b. Import standard modules -----
  Write-Host '  Importing standard modules...'
  foreach ($name in $standardModules) {
    $p = Join-Path $modulesDir $name
    $vbProj.VBComponents.Import($p) | Out-Null
    Write-Host "    + $name"
  }

  # ----- 4c. Import JsonConverter -----
  Write-Host '  Importing JsonConverter dependency...'
  $vbProj.VBComponents.Import($jcPath) | Out-Null
  Write-Host "    + $jsonConverter"

  # ----- 4d. Save as xlam to temp path -----
  # xlFileFormat 55 = xlOpenXMLAddIn (.xlam, the macro-enabled add-in format)
  Write-Host ''
  Write-Host '  Saving xlam...'
  $wb.IsAddin = $true
  $wb.SaveAs($tempXlam, 55)
  $wb.Close($false)
  $wb = $null
} finally {
  if ($wb -ne $null) {
    try { $wb.Close($false) } catch {}
  }
  if ($excel -ne $null) {
    try { $excel.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

if (-not (Test-Path $tempXlam)) {
  Write-Host 'ERROR: Excel did not produce the expected xlam file.' -ForegroundColor Red
  exit 1
}
$tempSize = (Get-Item $tempXlam).Length
Write-Host "  Built temp xlam: $tempSize bytes" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 5. Inject customUI14.xml + a top-level _rels/.rels with the relationship
# ---------------------------------------------------------------------------
Write-Host ''
Write-Host '  Injecting Ribbon XML and root _rels/.rels...'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Excel writes a workbook-level rels (xl/_rels/workbook.xml.rels) but NOT a
# top-level _rels/.rels when saving via SaveAs. The Custom UI relationship must
# live at the *package* level (root _rels/.rels), so we must construct that file
# ourselves. We rebuild the ZIP into a fresh archive to make the change.

$tmp2 = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "sec_xbrl_build2_$([System.Guid]::NewGuid().ToString('N')).xlam")
$zipIn = [System.IO.Compression.ZipFile]::OpenRead($tempXlam)
$zipOut = [System.IO.Compression.ZipFile]::Open($tmp2, 'Create')
try {
  # Names we generate ourselves; skip if Excel happened to include them.
  $skip = @{ '_rels/.rels' = $true; 'customUI/customUI14.xml' = $true }

  # Locate the package-level Relationships in the existing rels (if any).
  # Excel never writes a top-level _rels/.rels when SaveAs(.xlam) so we always
  # construct one. The standard package rels include:
  #   - rId1 -> xl/workbook.xml (officeDocument relationship)
  #   - rId2 -> docProps/core.xml (core-properties)
  #   - rId3 -> docProps/app.xml (extended-properties)
  #   - rId4 -> customUI/customUI14.xml (Custom UI extensibility)
  # We synthesize all four. Office tolerates additional rels but must have the
  # officeDocument rel.
  $rels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
</Relationships>
"@

  foreach ($e in $zipIn.Entries) {
    if ($skip.ContainsKey($e.FullName)) { continue }
    $newE = $zipOut.CreateEntry($e.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
    $sIn = $e.Open(); $sOut = $newE.Open()
    try { $sIn.CopyTo($sOut) } finally { $sIn.Dispose(); $sOut.Dispose() }
  }

  # Write _rels/.rels
  $relsBytes = [System.Text.Encoding]::UTF8.GetBytes($rels)
  $newRels = $zipOut.CreateEntry('_rels/.rels', [System.IO.Compression.CompressionLevel]::Optimal)
  $sOut = $newRels.Open()
  try { $sOut.Write($relsBytes, 0, $relsBytes.Length) } finally { $sOut.Dispose() }

  # Write customUI/customUI14.xml
  $newCu = $zipOut.CreateEntry('customUI/customUI14.xml', [System.IO.Compression.CompressionLevel]::Optimal)
  $sOut = $newCu.Open()
  try { $sOut.Write($customUIBytes, 0, $customUIBytes.Length) } finally { $sOut.Dispose() }

} finally {
  $zipIn.Dispose()
  $zipOut.Dispose()
}

# Move final into dist/
if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }
Move-Item -Force -Path $tmp2 -Destination $outXlam
Remove-Item -Force $tempXlam -ErrorAction SilentlyContinue

$finalSize = (Get-Item $outXlam).Length
Write-Host "  Final xlam written: $outXlam ($finalSize bytes)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 6. Verify by re-opening with Excel and listing modules
# ---------------------------------------------------------------------------
Write-Host ''
Write-Host '  Verifying built xlam...'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3
try {
  $vwb = $excel.Workbooks.Open($outXlam, 0, $true) # ReadOnly
  $names = @()
  foreach ($c in $vwb.VBProject.VBComponents) {
    $lc = $c.CodeModule.CountOfLines
    $names += "    {0,-22} (lines={1})" -f $c.Name, $lc
  }
  $vwb.Close($false)
  Write-Host '  Modules in built xlam:' -ForegroundColor Green
  $names | ForEach-Object { Write-Host $_ }
} catch {
  Write-Host "  Verification FAILED: $($_.Exception.Message)" -ForegroundColor Red
  throw
} finally {
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host ''
Write-Host '=== Build OK ===' -ForegroundColor Cyan
