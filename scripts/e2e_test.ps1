<#
.SYNOPSIS
  End-to-end test of dist/SEC_XBRL_Addin.xlam: load it as an add-in, then call
  PullSECFinancials directly via Application.Run with a fixed ticker (AAPL),
  verify the three output sheets were written.

  This test bypasses the InputBox by calling an internal helper sub. We add
  one if not present.

  USAGE
    powershell -ExecutionPolicy Bypass -File scripts/e2e_test.ps1
#>
param(
  [string]$Ticker = 'AAPL'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$xlamPath = Join-Path $repoRoot 'dist/SEC_XBRL_Addin.xlam'

if (-not (Test-Path $xlamPath)) {
  Write-Host "ERROR: $xlamPath not found - run scripts/build_xlam.ps1 first" -ForegroundColor Red
  exit 1
}

Write-Host '=== SEC EDGAR XBRL Add-in - E2E Test ===' -ForegroundColor Cyan
Write-Host "Add-in:  $xlamPath"
Write-Host "Ticker:  $Ticker"
Write-Host ''

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
# AutomationSecurity = msoAutomationSecurityLow (1) so macros can run
$excel.AutomationSecurity = 1
$excel.EnableEvents = $true

$tempOutXlsx = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "sec_xbrl_e2e_$([System.Guid]::NewGuid().ToString('N')).xlsx")

$exitCode = 1
$addin = $null
try {
  Write-Host '  Creating output workbook...'
  $outWb = $excel.Workbooks.Add()

  Write-Host "  Installing add-in via AddIns.Add..."
  $addin = $excel.AddIns.Add($xlamPath, $false)
  $addin.Installed = $true

  Write-Host "  Calling PullSECFinancialsForTicker('$Ticker')..."
  # The add-in is installed, so its public macros are in the global macro
  # namespace. PullSECFinancialsForTicker is pre-compiled (added in this
  # build) and takes the ticker as an argument so we don't need an InputBox.
  $excel.Run('PullSECFinancialsForTicker', $Ticker, $true)  # silent=True

  Write-Host "  Status: $($excel.StatusBar)"

  # Verify the three output sheets exist and have rows
  $sheets = @{}
  foreach ($ws in $outWb.Sheets) {
    $sheets[$ws.Name] = $ws.UsedRange.Rows.Count
  }
  Write-Host '  Sheets in output workbook:'
  $sheets.GetEnumerator() | Sort-Object Name | ForEach-Object {
    Write-Host ('    {0,-22} rows={1}' -f $_.Key, $_.Value)
  }

  $required = @('Income Statement','Balance Sheet','Cash Flow')
  $missing = $required | Where-Object { -not $sheets.ContainsKey($_) }
  if ($missing.Count -gt 0) {
    Write-Host "  FAIL - missing sheets: $($missing -join ', ')" -ForegroundColor Red
    $exitCode = 1
  } else {
    $minRows = ($required | ForEach-Object { $sheets[$_] } | Measure-Object -Minimum).Minimum
    if ($minRows -lt 10) {
      Write-Host "  FAIL - sheets exist but too few rows ($minRows)" -ForegroundColor Red
      $exitCode = 1
    } else {
      Write-Host '  PASS - all three sheets present and populated' -ForegroundColor Green
      $exitCode = 0
    }
  }

  # Save output for visual inspection
  $outWb.SaveAs($tempOutXlsx, 51)
  $outWb.Close($false)
  Write-Host "  Output saved at: $tempOutXlsx"

  $addin.Close($false)

} catch {
  Write-Host "  EXCEPTION: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "  ScriptStackTrace: $($_.ScriptStackTrace)" -ForegroundColor DarkGray
  $exitCode = 1
} finally {
  if ($null -ne $addin) {
    try { $addin.Installed = $false } catch {}
  }
  try { $excel.Quit() } catch {}
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

if ($exitCode -eq 0) {
  Write-Host ''
  Write-Host '=== E2E PASS ===' -ForegroundColor Green
} else {
  Write-Host ''
  Write-Host '=== E2E FAIL ===' -ForegroundColor Red
}
exit $exitCode
