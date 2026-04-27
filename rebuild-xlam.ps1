<#
.SYNOPSIS
    Rebuilds dist/SEC_XBRL_Addin.xlam from fixed source files.

.DESCRIPTION
    Strategy:
      1. Open Excel COM automation
      2. Create a brand-new blank workbook
      3. Import all fixed .bas module source files
      4. Replace ThisWorkbook event code
      5. Save the workbook as .xlam (Excel Open XML Add-In format)
      6. Close Excel
      7. Use .NET ZipFile to inject the ribbon customUI XML from the source tree
      8. Copy the rebuilt XLAM to dist/SEC_XBRL_Addin.xlam

    IMPORTANT: Close all Excel windows before running this script.

.NOTES
    Requires Excel to be installed. Run from the repository root, or pass -RepoRoot explicitly.
#>

param(
    [string]$RepoRoot = $PSScriptRoot
)

$ErrorActionPreference = "Stop"

# Paths
$modulesPath  = Join-Path $RepoRoot "modules"
$depsPath     = Join-Path $RepoRoot "dependencies"
$customUIPath = Join-Path $RepoRoot "customUI\customUI14.xml"
$distPath     = Join-Path $RepoRoot "dist\SEC_XBRL_Addin.xlam"
$tempXlam     = "C:\Temp\SEC_rebuilt.xlam"

foreach ($p in @($modulesPath, $depsPath, $customUIPath)) {
    if (-not (Test-Path $p)) { throw "Required path not found: $p" }
}
New-Item -Path "C:\Temp" -ItemType Directory -Force | Out-Null

# Kill any running Excel
$runningExcel = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($runningExcel) {
    Write-Host "Stopping Excel processes..."
    $runningExcel | Stop-Process -Force
    Start-Sleep -Seconds 3
}

# Enable VBA project access and lower macro security
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
$prevAccess = $null; try { $prevAccess = (Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction Stop).AccessVBOM } catch {}
$prevLevel  = $null; try { $prevLevel  = (Get-ItemProperty -Path $regPath -Name "Level"      -ErrorAction Stop).Level      } catch {}
Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force
Set-ItemProperty -Path $regPath -Name "Level"      -Value 1 -Type DWord -Force
Write-Host "Registry: VBA access enabled, macro security set to Low"

# Module list (name -> source path)
$modules = [ordered]@{
    "modConfig"       = "$modulesPath\modConfig.bas"
    "modHTTP"         = "$modulesPath\modHTTP.bas"
    "modProgress"     = "$modulesPath\modProgress.bas"
    "modClassifier"   = "$modulesPath\modClassifier.bas"
    "modTickerLookup" = "$modulesPath\modTickerLookup.bas"
    "modJSONParser"   = "$modulesPath\modJSONParser.bas"
    "modExcelWriter"  = "$modulesPath\modExcelWriter.bas"
    "modRibbon"       = "$modulesPath\modRibbon.bas"
    "modMain"         = "$modulesPath\modMain.bas"
    "JsonConverter"   = "$depsPath\JsonConverter.bas"
}

function Remove-VBC([object]$proj, [string]$name) {
    try { $c = $proj.VBComponents.Item($name); $proj.VBComponents.Remove($c); Write-Host "  removed: $name" } catch {}
}

# === Phase 1: Excel COM - build VBA workbook ===
$xl = $null
Write-Host ""
Write-Host "=== Phase 1: Building VBA workbook via Excel COM ==="
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible       = $false
    $xl.DisplayAlerts = $false
    Write-Host "Excel $($xl.Version) started"
    Start-Sleep -Seconds 2

    $wb  = $xl.Workbooks.Add()
    $vbp = $wb.VBProject
    Write-Host "New workbook created: $($wb.Name)"

    # Remove any auto-created standard modules
    $toRemove = @()
    for ($i = 1; $i -le $vbp.VBComponents.Count; $i++) {
        $c = $vbp.VBComponents.Item($i)
        if ($c.Type -eq 1) { $toRemove += $c.Name }
    }
    foreach ($n in $toRemove) { Remove-VBC $vbp $n }

    # Import all fixed modules
    foreach ($entry in $modules.GetEnumerator()) {
        $name = $entry.Key
        $path = $entry.Value
        if (-not (Test-Path $path)) { Write-Warning "  NOT FOUND: $path"; continue }
        $comp = $vbp.VBComponents.Import($path)
        Write-Host "  imported: $($comp.Name)"
    }

    # Update ThisWorkbook event code
    Write-Host "  updating ThisWorkbook..."
    $clsSource = Get-Content "$modulesPath\ThisWorkbook.cls" -Raw
    $startIdx  = $clsSource.IndexOf("Option Explicit")
    if ($startIdx -lt 0) { $startIdx = $clsSource.IndexOf("Private Sub") }
    $codeBody  = $clsSource.Substring($startIdx)
    $twComp    = $vbp.VBComponents.Item("ThisWorkbook")
    $cm        = $twComp.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.InsertLines(1, $codeBody)
    Write-Host "  ThisWorkbook: $($cm.CountOfLines) lines"

    # Save as XLAM (55 = xlOpenXMLAddIn)
    if (Test-Path $tempXlam) { Remove-Item $tempXlam -Force }
    $wb.SaveAs($tempXlam, 55)
    Write-Host "  Saved: $tempXlam ($([Math]::Round((Get-Item $tempXlam).Length/1KB,1)) KB)"
    $wb.Close($false)
    $xl.Quit()
    Write-Host "Excel closed"

} finally {
    if ($xl) { try { $xl.Quit() } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($null -eq $prevAccess) { try { Remove-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction Stop } catch {} } else { Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value $prevAccess -Type DWord -Force }
    if ($null -eq $prevLevel)  { try { Remove-ItemProperty -Path $regPath -Name "Level"      -ErrorAction Stop } catch {} } else { Set-ItemProperty -Path $regPath -Name "Level"      -Value $prevLevel  -Type DWord -Force }
    Write-Host "Registry: security settings restored"
}

if (-not (Test-Path $tempXlam)) { throw "Phase 1 failed: XLAM not created at $tempXlam" }

# === Phase 2: Inject ribbon customUI XML ===
Write-Host ""
Write-Host "=== Phase 2: Injecting customUI ribbon XML ==="
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$zipPath = $tempXlam -replace '\.xlam$', '.zip'
[System.IO.File]::Copy($tempXlam, $zipPath, $true)

$zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    # Add customUI/customUI14.xml
    $xmlContent = [System.IO.File]::ReadAllText($customUIPath, [System.Text.Encoding]::UTF8)
    $xmlBytes   = [System.Text.Encoding]::UTF8.GetBytes($xmlContent)
    $existing   = $zip.GetEntry("customUI/customUI14.xml")
    if ($existing) { $existing.Delete() }
    $cuEntry = $zip.CreateEntry("customUI/customUI14.xml")
    $stream  = $cuEntry.Open()
    $stream.Write($xmlBytes, 0, $xmlBytes.Length)
    $stream.Close()
    Write-Host "  added: customUI/customUI14.xml"

    # Update _rels/.rels
    $relsEntry = $zip.GetEntry("_rels/.rels")
    $sr        = New-Object System.IO.StreamReader($relsEntry.Open())
    $relsXml   = $sr.ReadToEnd()
    $sr.Close()
    $relsXml = [regex]::Replace($relsXml, '<Relationship[^>]+Type="http://schemas\.microsoft\.com/office/200[679]/relationships/ui/extensibility"[^>]*/>', '')
    $newRel  = '<Relationship Id="rIdUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
    $relsXml = $relsXml -replace '</Relationships>', "$newRel</Relationships>"
    $relsEntry.Delete()
    $newRE = $zip.CreateEntry("_rels/.rels")
    $rb    = [System.Text.Encoding]::UTF8.GetBytes($relsXml)
    $ws    = $newRE.Open(); $ws.Write($rb, 0, $rb.Length); $ws.Close()
    Write-Host "  updated: _rels/.rels"

    # Update [Content_Types].xml
    $ctEntry = $zip.GetEntry("[Content_Types].xml")
    $sr2     = New-Object System.IO.StreamReader($ctEntry.Open())
    $ctXml   = $sr2.ReadToEnd()
    $sr2.Close()
    if ($ctXml -notmatch 'customUI14') {
        $ctXml = $ctXml -replace '</Types>', '<Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/></Types>'
        $ctEntry.Delete()
        $newCT   = $zip.CreateEntry("[Content_Types].xml")
        $ctBytes = [System.Text.Encoding]::UTF8.GetBytes($ctXml)
        $ws2     = $newCT.Open(); $ws2.Write($ctBytes, 0, $ctBytes.Length); $ws2.Close()
        Write-Host "  updated: [Content_Types].xml"
    }
} finally {
    $zip.Dispose()
}

[System.IO.File]::Copy($zipPath, $tempXlam, $true)
Remove-Item $zipPath -Force
Write-Host "  Ribbon XML injected OK"

# === Phase 3: Deploy to dist/ ===
Write-Host ""
Write-Host "=== Phase 3: Deploying to dist/ ==="
[System.IO.File]::Copy($tempXlam, $distPath, $true)
$sz = [Math]::Round((Get-Item $distPath).Length / 1KB, 1)
Write-Host "  Deployed: $distPath ($sz KB)"
Write-Host ""
Write-Host "SUCCESS: dist/SEC_XBRL_Addin.xlam rebuilt from fixed source ($sz KB)"
Write-Host "Modified: $((Get-Item $distPath).LastWriteTime)"
