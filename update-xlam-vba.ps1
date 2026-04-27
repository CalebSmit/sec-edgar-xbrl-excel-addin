<#
.SYNOPSIS
    Updates the VBA code in the XLAM file by importing updated modules from source.
    This script uses Excel COM to import the modified modConfig.bas into the XLAM.

.NOTES
    - Requires Excel to be installed
    - Closes any open instances of the XLAM
    - Updates vbaProject.bin with new module content
#>

param(
    [string]$XlamPath = "dist\SEC_XBRL_Addin.xlam",
    [string]$ModulePath = "modules\modConfig.bas"
)

$XlamFullPath = Join-Path $PSScriptRoot $XlamPath
$ModuleFullPath = Join-Path $PSScriptRoot $ModulePath

if (-not (Test-Path $XlamFullPath)) {
    Write-Error "XLAM not found: $XlamFullPath"
    exit 1
}

if (-not (Test-Path $ModuleFullPath)) {
    Write-Error "Module not found: $ModuleFullPath"
    exit 1
}

Write-Host "Opening Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening XLAM: $XlamFullPath"
    $workbook = $excel.Workbooks.Open($XlamFullPath, $false, $false)
    
    Write-Host "Accessing VBE..."
    $vbProject = $workbook.VBProject
    
    if ($vbProject -eq $null) {
        Write-Error "Cannot access VBE. Ensure the add-in is not already loaded in Excel."
        exit 1
    }
    
    # Find and remove old modConfig module
    $moduleExists = $false
    foreach ($component in $vbProject.VBComponents) {
        if ($component.Name -eq "modConfig") {
            Write-Host "Removing old modConfig module..."
            $vbProject.VBComponents.Remove($component)
            $moduleExists = $true
            break
        }
    }
    
    # Import the updated module
    Write-Host "Importing updated modConfig.bas..."
    $vbProject.VBComponents.Import($ModuleFullPath) | Out-Null
    
    Write-Host "Saving XLAM..."
    $workbook.Save()
    
    Write-Host "Closing workbook..."
    $workbook.Close()
    
    Write-Host "Successfully updated XLAM!"
}
catch {
    Write-Error "Error: $_"
    exit 1
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}
