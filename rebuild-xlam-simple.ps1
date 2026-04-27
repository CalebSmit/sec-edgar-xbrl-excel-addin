
# Simple XLAM rebuild - imports updated modules
# CRITICAL: Excel and the add-in must be completely closed first!

param([string]$XlamPath = "dist\SEC_XBRL_Addin.xlam")

$XlamFullPath = Join-Path $PSScriptRoot $XlamPath

Write-Host "Rebuilding XLAM with updated modules..."
Write-Host "Opening Excel..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening XLAM..."
    $workbook = $excel.Workbooks.Open($XlamFullPath, $false, $false)
    $vbProject = $workbook.VBProject
    
    Write-Host "Removing old modConfig..."
    foreach ($comp in $vbProject.VBComponents) {
        if ($comp.Name -eq "modConfig") {
            $vbProject.VBComponents.Remove($comp)
            break
        }
    }
    
    Start-Sleep -Milliseconds 500
    
    Write-Host "Importing new modConfig.bas..."
    $moduleFile = Join-Path $PSScriptRoot "modules\modConfig.bas"
    $vbProject.VBComponents.Import($moduleFile) | Out-Null
    
    Write-Host "Importing new modHTTP.bas..."
    $httpFile = Join-Path $PSScriptRoot "modules\modHTTP.bas"
    foreach ($comp in $vbProject.VBComponents) {
        if ($comp.Name -eq "modHTTP") {
            $vbProject.VBComponents.Remove($comp)
            break
        }
    }
    Start-Sleep -Milliseconds 300
    $vbProject.VBComponents.Import($httpFile) | Out-Null
    
    Start-Sleep -Milliseconds 1000
    
    Write-Host "Saving XLAM..."
    $workbook.Save()
    $workbook.Close($false)
    
    Write-Host "SUCCESS - XLAM rebuilt with updated code!"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
