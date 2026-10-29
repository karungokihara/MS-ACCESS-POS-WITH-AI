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
    Write-Host "F_UnitList already exists"
    $acc.CloseCurrentDatabase()
    $acc.Quit()
    Write-Host "DONE"
    exit
}

function Find-ByName($form, $name) {
    foreach ($c in $form.Controls) { if ($c.Name -eq $name) { return $c } }
    return $null
}
function Make-Label($form, $name, $caption, $left, $top, $width) {
    $c = Find-ByName $form $name
    if ($null -eq $c) {
        $c = $acc.CreateControl($form.Name, 100, 1, "", "", $left, $top, $width, 260)
        $c.Name = $name
    }
    $c.Caption = $caption
    $c.FontName = "Segoe UI Semibold"
    $c.FontSize = 10
}
function Make-Text($form, $name, $source, $left, $width) {
    $c = Find-ByName $form $name
    if ($null -eq $c) {
        $c = $acc.CreateControl($form.Name, 109, 0, "", "", $left, 60, $width, 300)
        $c.Name = $name
    }
    $c.ControlSource = $source
    $c.FontName = "Segoe UI"
    $c.FontSize = 10
    $c.BorderStyle = 1
}
function Make-Button($form, $name, $caption, $left, $onClick) {
    $c = Find-ByName $form $name
    if ($null -eq $c) {
        $c = $acc.CreateControl($form.Name, 104, 1, "", "", $left, 300, 1600, 320)
        $c.Name = $name
    }
    $c.Caption = $caption
    $c.OnClick = $onClick
    $c.FontName = "Segoe UI Semibold"
    $c.FontSize = 10
}

$nf = $acc.CreateForm()
$nf.Name = "F_UnitList"
$nf.Caption = "Units"
$nf.RecordSource = "Q_UnitList"
$nf.DefaultView = 1
$nf.ScrollBars = 3
$nf.NavigationButtons = $false
$nf.RecordSelectors = $false
$nf.AutoCenter = $true
$nf.AllowAdditions = $false
$nf.AllowEdits = $false
$nf.AllowDeletions = $false
$nf.Width = 11520
$nf.Section(1).Height = 640
$nf.Section(0).Height = 420

Make-Label $nf "lblUnitCode"     "Unit Code"   360   60 1200
Make-Label $nf "lblUnitName"     "Unit Name"   1800  60 3000
Make-Label $nf "lblDescription"  "Description" 5040  60 3600
Make-Label $nf "lblItemCount"    "Items"       8880  60 900
Make-Label $nf "lblRecordStatus" "Status"      10020 60 1400

Make-Button $nf "cmdNewUnit"   "New Unit"  360  "=UX_OpenNewUnit()"
Make-Button $nf "cmdEditUnit"  "Edit Unit" 2100 "=UX_EditCurrentUnit()"
Make-Button $nf "cmdCloseUnit" "Close"     3840 "=UX_CloseActiveForm()"

Make-Text $nf "txtUnitCode"     "UnitCode"     360   1200
Make-Text $nf "txtUnitName"     "UnitName"     1800  3000
Make-Text $nf "txtDescription"  "Description"  5040  3600
Make-Text $nf "txtItemCount"    "ItemCount"    8880  900
Make-Text $nf "txtRecordStatus" "RecordStatus" 10020 1400

$rc = Find-ByName $nf "txtRecordStatus"
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
