$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$out = "C:\Users\user\Downloads\LLC\_work\cur"
if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out | Out-Null }
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$proj = $acc.VBE.VBProjects.Item(1)
foreach ($comp in $proj.VBComponents) {
    $cm = $comp.CodeModule
    $n = $cm.CountOfLines
    $txt = ""
    if ($n -gt 0) { $txt = $cm.Lines(1, $n) }
    $safe = $comp.Name -replace '[\\/:*?"<>|]', '_'
    [System.IO.File]::WriteAllText((Join-Path $out ($safe + ".txt")), $txt, [System.Text.Encoding]::UTF8)
}
Write-Host ("exported " + $proj.VBComponents.Count + " components to " + $out)
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
