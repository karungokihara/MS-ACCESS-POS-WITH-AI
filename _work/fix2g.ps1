$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$acc.DoCmd.OpenForm("F_UnitList", 1)
Start-Sleep -Milliseconds 500
$f = $acc.Forms.Item("F_UnitList")

function Find-ByName($form, $name) {
    foreach ($c in $form.Controls) { if ($c.Name -eq $name) { return $c } }
    return $null
}

foreach ($n in @("UnitCode_Label","UnitName_Label","Description_Label","ItemCount_Label","RecordStatus_Label")) {
    try { $acc.DeleteControl("F_UnitList", $n); Write-Host ("removed " + $n) } catch { Write-Host ("remove failed " + $n) }
}

$f.DefaultView = 1
$f.ScrollBars = 3
$f.Section(1).Height = 700
$f.Section(0).Height = 420

$cols = @(
    @("hdrUnitCode", "Unit Code", 360, 1200),
    @("hdrUnitName", "Unit Name", 1800, 3000),
    @("hdrDescription", "Description", 5040, 3600),
    @("hdrItemCount", "Items", 8880, 900),
    @("hdrRecordStatus", "Status", 10020, 1400)
)
foreach ($col in $cols) {
    $lbl = Find-ByName $f $col[0]
    if ($null -eq $lbl) {
        $lbl = $acc.CreateControl("F_UnitList", 100, 1, "", "", $col[2], 380, $col[3], 260)
        $lbl.Name = $col[0]
    }
    $lbl.Caption = $col[1]
    $lbl.Left = $col[2]
    $lbl.Top = 380
    $lbl.Width = $col[3]
    $lbl.FontName = "Segoe UI Semibold"
    $lbl.FontSize = 10
    Write-Host ("header " + $col[0])
}

$texts = @(
    @("txtUnitCode", 360, 1200),
    @("txtUnitName", 1800, 3000),
    @("txtDescription", 5040, 3600),
    @("txtItemCount", 8880, 900),
    @("txtRecordStatus", 10020, 1400)
)
foreach ($t in $texts) {
    $txt = Find-ByName $f $t[0]
    if ($null -ne $txt) {
        $txt.Left = $t[1]
        $txt.Top = 60
        $txt.Width = $t[2]
        $txt.Height = 330
    }
}

foreach ($bn in @("cmdNewUnit","cmdEditUnit","cmdCloseUnit")) {
    $b = Find-ByName $f $bn
    if ($null -ne $b) { $b.Top = 60; $b.Height = 340 }
}

$acc.DoCmd.Save(2, "F_UnitList")
$acc.DoCmd.Close(2, "F_UnitList", 1)
Write-Host "saved"

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
