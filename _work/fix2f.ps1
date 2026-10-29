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

$title = Find-ByName $f "Label12"
if ($null -ne $title) {
    $title.Caption = "Units"
    $title.Left = 2200
    Write-Host "title updated"
}

$cols = @(
    @("txtUnitCode", "UnitCode_Label", "Unit Code"),
    @("txtUnitName", "UnitName_Label", "Unit Name"),
    @("txtDescription", "Description_Label", "Description"),
    @("txtItemCount", "ItemCount_Label", "Items"),
    @("txtRecordStatus", "RecordStatus_Label", "Status")
)
foreach ($col in $cols) {
    $txt = Find-ByName $f $col[0]
    if ($null -eq $txt) { Write-Host ("missing " + $col[0]); continue }
    $lbl = Find-ByName $f $col[1]
    if ($null -eq $lbl) {
        $lbl = $acc.CreateControl("F_UnitList", 100, 0, "", "", $txt.Left, 0, $txt.Width, 260)
        $lbl.Name = $col[1]
        Write-Host ("created " + $col[1])
    }
    $lbl.Caption = $col[2]
    $lbl.Left = $txt.Left
    $lbl.Top = 0
    $lbl.Width = $txt.Width
    $lbl.FontName = "Segoe UI Semibold"
    $lbl.FontSize = 10
    $lbl.LabelAlign = 0
    $txt.LabelName = $lbl.Name
    Write-Host ("linked " + $col[1])
}

$f.Section(0).Height = 420
$acc.DoCmd.Save(2, "F_UnitList")
$acc.DoCmd.Close(2, "F_UnitList", 1)
Write-Host "saved"

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
