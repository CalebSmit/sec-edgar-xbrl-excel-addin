<#
.SYNOPSIS
    Clean XLAM rebuild by directly manipulating the ZIP file.
    DOES NOT require Excel to be running.
    CRITICAL: User must close Excel and uninstall the add-in FIRST.

.INSTRUCTIONS
    BEFORE running this script:
    1. Close Excel completely
    2. Excel → File → Options → Add-ins → Manage: Excel Add-ins → Go...
    3. Select "SEC XBRL Excel Add-in" → REMOVE
    4. Close Excel
    5. Run this script
    6. Re-download from GitHub and install fresh

.NOTES
    This script rebuilds vbaProject.bin by using Excel with the add-in CLOSED.
#>

param(
    [string]$SourceDir = "modules",
    [string]$XlamPath = "dist\SEC_XBRL_Addin.xlam"
)

Write-Host @"
╔═══════════════════════════════════════════════════════════════════════════╗
║  XLAM CLEAN REBUILD                                                       ║
║  ───────────────────────────────────────────────────────────────────────  ║
║  WARNING: Ensure Excel is COMPLETELY CLOSED and add-in is UNINSTALLED    ║
║           before proceeding!                                              ║
╚═══════════════════════════════════════════════════════════════════════════╝
"@

$XlamFullPath = Join-Path $PSScriptRoot $XlamPath

# Check if XLAM is locked by Excel
if (Test-Path $XlamFullPath) {
    try {
        $file = [System.IO.File]::Open($XlamFullPath, 'Open', 'Read', 'None')
        $file.Close()
        Write-Host "✓ XLAM file is not locked by Excel"
    }
    catch {
        Write-Error @"
XLAM file is LOCKED by Excel. 

You must:
1. Close Excel completely (all windows)
2. Close the Visual Basic editor
3. Uninstall the add-in from Excel Add-ins dialog
4. Then run this script again

The error was: $_
"@
        exit 1
    }
}

Write-Host "`nStep 1: Opening Excel (hidden) to access VBE..."
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
}
catch {
    Write-Error "Failed to create Excel COM object. Ensure Excel is installed. Error: $_"
    exit 1
}

try {
    Write-Host "Step 2: Opening XLAM workbook..."
    # Use ReadOnly flag to open workbook without triggering add-in loading
    $workbook = $excel.Workbooks.Open($XlamFullPath, $false, $false)
    
    Write-Host "Step 3: Accessing VBA project..."
    $vbProject = $workbook.VBProject
    
    if ($null -eq $vbProject) {
        Write-Error "Cannot access VBA project. The XLAM may be corrupted or require unblocking."
        $workbook.Close($false)
        exit 1
    }
    
    Write-Host "Step 4: Removing old modConfig module..."
    $removed = $false
    foreach ($component in $vbProject.VBComponents) {
        if ($component.Name -eq "modConfig") {
            $vbProject.VBComponents.Remove($component)
            $removed = $true
            Write-Host "  ✓ Removed old modConfig"
            break
        }
    }
    
    if (-not $removed) {
        Write-Warning "  ⚠ modConfig module not found (may have been already removed)"
    }
    
    # Wait a moment for removal to complete
    Start-Sleep -Milliseconds 500
    
    Write-Host "Step 5: Importing updated modConfig.bas..."
    $moduleFile = Join-Path $PSScriptRoot "modules\modConfig.bas"
    if (-not (Test-Path $moduleFile)) {
        Write-Error "modules\modConfig.bas not found!"
        $workbook.Close($false)
        exit 1
    }
    
    try {
        $vbProject.VBComponents.Import($moduleFile) | Out-Null
        Write-Host "  ✓ Successfully imported modConfig.bas"
    }
    catch {
        Write-Error "Failed to import modConfig.bas: $_"
        $workbook.Close($false)
        exit 1
    }
    
    # Wait for import to complete
    Start-Sleep -Milliseconds 1000
    
    Write-Host "Step 6: Saving XLAM workbook..."
    try {
        $workbook.Save()
        Write-Host "  ✓ Saved successfully"
    }
    catch {
        Write-Error "Failed to save workbook: $_"
        exit 1
    }
    
    Write-Host "Step 7: Closing workbook..."
    $workbook.Close($false)
    Write-Host "  ✓ Closed"
    
}
finally {
    Write-Host "Step 8: Closing Excel..."
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Host "  ✓ Excel closed"
}

Write-Host @"

╔═══════════════════════════════════════════════════════════════════════════╗
║  REBUILD COMPLETE                                                         ║
║  ───────────────────────────────────────────────────────────────────────  ║
║  Next steps:                                                              ║
║  1. Download fresh XLAM from GitHub:                                     ║
║     https://github.com/CalebSmit/sec-edgar-xbrl-excel-addin/raw/master/  ║
║     dist/SEC_XBRL_Addin.xlam                                            ║
║  2. Save to: C:\Users\YourName\Documents\SEC_XBRL_Addin.xlam             ║
║  3. Right-click → Properties → check "Unblock" → OK                      ║
║  4. Excel: File → Options → Add-ins → Trust Center → Add Documents       ║
║  5. Excel: File → Options → Add-ins → Manage: Excel Add-ins → Go        ║
║  6. Browse and select the XLAM file                                       ║
║  7. Test with ticker "AAPL"                                              ║
╚═══════════════════════════════════════════════════════════════════════════╝
"@
