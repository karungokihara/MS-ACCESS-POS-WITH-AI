$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)

$exists = $false
foreach ($af in $acc.CurrentProject.AllForms) { if ($af.Name -eq "F_UnitList") { $exists = $true } }
if ($exists) {
    Write-Host "F_UnitList already exists; removing for rebuild"
    $acc.DoCmd.DeleteObject(2, "F_UnitList")
}

$acc.DoCmd.CopyObject("", "F_UnitList", 2, "F_CategoriesList")
Write-Host "copied F_CategoriesList -> F_UnitList"

$acc.DoCmd.OpenForm("F_UnitList", 1)
Start-Sleep -Milliseconds 500
$f = $acc.Forms.Item("F_UnitList")
Write-Host ("opened design; controls=" + $f.Controls.Count)

$drop = @(
    "CategoryID","CategoryID_Label",
    "ItemTypeID","ItemTypeID_Label",
    "CategoryName","CategoryName_Label",
    "Description","Description_Label",
    "RequiresSerial","RequiresSerial_Label",
    "RequiresExpiry","RequiresExpiry_Label"
)
foreach ($n in $drop) {
    try { $acc.DeleteControl("F_UnitList", $n); Write-Host ("  removed " + $n) } catch { Write-Host ("  remove failed " + $n) }
}

$f.RecordSource = "Q_UnitList"
$f.Caption = "Units"
$f.DefaultView = 2
$f.FilterOn = $false
$f.OrderByOn = $false

function Find-ByName($form, $name) {
    foreach ($c in $form.Controls) { if ($c.Name -eq $name) { return $c } }
    return $null
}
function Add-Text($form, $name, $source, $caption, $left, $width) {
    $c = Find-ByName $form $name
    if ($null -eq $c) {
        $c = $acc.CreateControl($form.Name, 109, 0, "", "", $left, 60, $width, 300)
        $c.Name = $name
    }
    $c.ControlSource = $source
    try { $c.DatasheetCaption = $caption } catch {}
    $c.FontName = "Segoe UI"
    $c.FontSize = 10
    try { $c.ColumnWidth = -1 } catch {}
    Write-Host ("  column " + $name + " -> " + $source)
}
function Add-Button($form, $name, $caption, $left, $onClick) {
    $c = Find-ByName $form $name
    if ($null -eq $c) {
        $c = $acc.CreateControl($form.Name, 104, 1, "", "", $left, 300, 1800, 340)
        $c.Name = $name
    }
    $c.Caption = $caption
    $c.OnClick = $onClick
    $c.FontName = "Segoe UI Semibold"
    $c.FontSize = 10
    Write-Host ("  button " + $name)
}

Add-Text $f "txtUnitCode"     "UnitCode"     "Unit Code"   360   1200
Add-Text $f "txtUnitName"     "UnitName"     "Unit Name"   1800  3000
Add-Text $f "txtDescription"  "Description"  "Description" 5040  3600
Add-Text $f "txtItemCount"    "ItemCount"    "Items"       8880  900
Add-Text $f "txtRecordStatus" "RecordStatus" "Status"      10020 1400

Add-Button $f "cmdNewUnit"   "New Unit"  9060  "=UX_OpenNewUnit()"
Add-Button $f "cmdEditUnit"  "Edit Unit" 11060 "=UX_EditCurrentUnit()"
Add-Button $f "cmdCloseUnit" "Close"     13060 "=UX_CloseActiveForm()"

$f.Section(0).Height = 420

$rc = Find-ByName $f "txtRecordStatus"
try {
    $fc = $rc.FormatConditions.Add(1, 0, "RecordStatus='Incomplete'")
    $fc.ForeColor = 255
    $fc.FontBold = $true
    Write-Host "conditional format added"
} catch {
    Write-Host ("format condition failed: " + $_.Exception.Message)
}

$acc.DoCmd.Save(2, "F_UnitList")
$acc.DoCmd.Close(2, "F_UnitList", 1)
Write-Host "F_UnitList saved"

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
