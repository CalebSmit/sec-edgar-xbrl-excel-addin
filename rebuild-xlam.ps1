<#
.SYNOPSIS
    Updates dist/SEC_XBRL_Addin.xlam ribbon with latest customUI14.xml.

.NOTES
    Does NOT require Excel. Uses only .NET System.IO.Compression (built-in).
    Extracts XLAM (ZIP), updates 3 files, re-zips, and replaces the original.
    No registry manipulation, no Excel COM, no VBA import complexity.
#>

param([string]$RepoRoot = $PSScriptRoot)

$ErrorActionPreference = "Stop"

# Paths
$customUIPath = Join-Path $RepoRoot "customUI\customUI14.xml"
$distPath     = Join-Path $RepoRoot "dist\SEC_XBRL_Addin.xlam"
$tempDir      = "C:\Temp\xlam_update_$(Get-Random)"

if (-not (Test-Path $customUIPath)) { throw "customUI/customUI14.xml not found: $customUIPath" }
if (-not (Test-Path $distPath))     { throw "dist/SEC_XBRL_Addin.xlam not found: $distPath" }

Write-Host "Updating ribbon in dist/SEC_XBRL_Addin.xlam..."

# Extract XLAM (which is a ZIP file)
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item $tempDir -ItemType Directory -Force | Out-Null
[System.IO.Compression.ZipFile]::ExtractToDirectory($distPath, $tempDir)
Write-Host "  extracted XLAM"

# Step 1: Update customUI/customUI14.xml
$cuiDir = Join-Path $tempDir "customUI"
New-Item $cuiDir -ItemType Directory -Force | Out-Null
Copy-Item $customUIPath "$cuiDir\customUI14.xml" -Force
Write-Host "  customUI/customUI14.xml: updated"

# Step 2: Update _rels/.rels (add UI relationship if missing)
$relsPath = Join-Path $tempDir "_rels\.rels"
[xml]$relsXml = Get-Content $relsPath
$nsm = New-Object Xml.XmlNamespaceManager($relsXml.NameTable)
$nsm.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")
$uiRel = $relsXml.DocumentElement.SelectSingleNode("//r:Relationship[@Type='http://schemas.microsoft.com/office/2007/relationships/ui/extensibility']", $nsm)

if (-not $uiRel) {
    $rel = $relsXml.CreateElement("Relationship")
    $rel.SetAttribute("Id", "rIdUI")
    $rel.SetAttribute("Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility")
    $rel.SetAttribute("Target", "customUI/customUI14.xml")
    $relsXml.DocumentElement.AppendChild($rel) | Out-Null
    $relsXml.Save($relsPath)
    Write-Host "  _rels/.rels: added UI relationship"
} else {
    Write-Host "  _rels/.rels: UI relationship already present"
}

# Step 3: Update [Content_Types].xml (add customUI14 override if missing)
$ctPath = Join-Path $tempDir "[Content_Types].xml"
[xml]$ctXml = Get-Content -LiteralPath $ctPath
$override = $ctXml.DocumentElement.SelectSingleNode("//Override[@PartName='/customUI/customUI14.xml']")

if (-not $override) {
    $over = $ctXml.CreateElement("Override")
    $over.SetAttribute("PartName", "/customUI/customUI14.xml")
    $over.SetAttribute("ContentType", "application/xml")
    $ctXml.DocumentElement.AppendChild($over) | Out-Null
    $ctXml.Save($ctPath)
    Write-Host "  [Content_Types].xml: added customUI14 override"
} else {
    Write-Host "  [Content_Types].xml: customUI14 override already present"
}

# Re-zip and replace
$tempZip = Join-Path "C:\Temp" "rebuilt_$(Get-Random).zip"
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $tempZip)
Remove-Item $distPath -Force
Move-Item $tempZip $distPath
Remove-Item $tempDir -Recurse -Force

$sz = [Math]::Round((Get-Item $distPath).Length / 1KB, 1)
Write-Host ""
Write-Host "SUCCESS: dist/SEC_XBRL_Addin.xlam updated ($sz KB)"
