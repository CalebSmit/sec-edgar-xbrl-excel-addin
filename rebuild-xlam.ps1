<#
.SYNOPSIS
    DEPRECATED: Do not run this script. The dist/SEC_XBRL_Addin.xlam is a
    checked-in binary artifact. It should NOT be modified by automated scripts.

.NOTES
    This script is kept for historical reference only.
    To update the ribbon, use the manual Office RibbonX Editor approach instead.
#>

Write-Host @"
╔════════════════════════════════════════════════════════════════╗
║ DEPRECATED: rebuild-xlam.ps1 should not be run.               ║
║                                                                ║
║ The dist/SEC_XBRL_Addin.xlam is a checked-in binary artifact. ║
║ Automated scripts risk corruption (ZipFile.Open API issues).  ║
║                                                                ║
║ Manual Ribbon Update (for developers only):                   ║
║  1. Download Office RibbonX Editor                            ║
║     https://github.com/OfficeDev/office-ribbonx-editor        ║
║  2. File → Open → dist/SEC_XBRL_Addin.xlam                   ║
║  3. Insert → Office 2010+ Custom UI Part                      ║
║  4. Paste contents of customUI/customUI14.xml                 ║
║  5. Validate (Ctrl+Enter) → Save (Ctrl+S)                     ║
║  6. Close editor, test in Excel                               ║
║  7. git add dist/SEC_XBRL_Addin.xlam && git commit            ║
║                                                                ║
╚════════════════════════════════════════════════════════════════╝
"@
