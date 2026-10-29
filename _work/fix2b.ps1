$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$db = $acc.CurrentDb()

function Find-ByCaption($form, $type, $caption) {
    foreach ($c in $form.Controls) {
        if ($c.ControlType -eq $type -and [string]$c.Caption -eq $caption) { return $c }
    }
    return $null
}
function Find-BySource($form, $type, $source) {
    foreach ($c in $form.Controls) {
        if ($c.ControlType -eq $type -and [string]$c.ControlSource -eq $source) { return $c }
    }
    return $null
}
function Find-ByName($form, $name) {
    foreach ($c in $form.Controls) { if ($c.Name -eq $name) { return $c } }
    return $null
}

Write-Host "== F_AssemblyTemplate control names"
$acc.DoCmd.OpenForm("F_AssemblyTemplate", 1)
Start-Sleep -Milliseconds 400
$f = $acc.Forms.Item("F_AssemblyTemplate")

$renames = @(
    @("ManufacturerPartNumber", 109, "txtManufacturerPartNumber"),
    @("VendorPartNumber", 109, "txtVendorPartNumber"),
    @("InternalPartNumber", 109, "txtInternalPartNumber"),
    @("AutoGeneratePartNumber", 106, "chkAutoGeneratePartNumber")
)
foreach ($r in $renames) {
    $c = Find-BySource $f $r[1] $r[0]
    if ($null -ne $c) {
        if ($c.Name -ne $r[2]) { $c.Name = $r[2]; Write-Host ("  renamed " + $r[2]) }
        else { Write-Host ("  ok " + $r[2]) }
    } else {
        Write-Host ("  NOT FOUND for source " + $r[0])
    }
}
$labels = @(
    @("Part Numbers", "lblPartNumbersTitle"),
    @("Manufacturer Part No.", "lblManufacturerPartNumber"),
    @("Vendor Part No.", "lblVendorPartNumber"),
    @("Internal Part No.", "lblInternalPartNumber")
)
foreach ($r in $labels) {
    $c = Find-ByCaption $f 100 $r[0]
    if ($null -ne $c -and $c.Name -ne $r[1]) { $c.Name = $r[1]; Write-Host ("  renamed " + $r[1]) }
}

$chk = Find-ByName $f "chkAutoGeneratePartNumber"
if ($null -ne $chk) {
    $chk.Width = 300
    $chk.ControlSource = "AutoGeneratePartNumber"
    $chk.TripleState = $false
    Write-Host "  checkbox sized"
}

$lbl = Find-ByName $f "lblAutoGeneratePartNumber"
if ($null -eq $lbl) {
    $lbl = $acc.CreateControl("F_AssemblyTemplate", 100, 0, "", "", 7360, 13040, 4600, 300)
    $lbl.Name = "lblAutoGeneratePartNumber"
    Write-Host "  label created"
}
$lbl.Caption = "Auto generate internal part number"
$lbl.FontName = "Segoe UI"
$lbl.FontSize = 10
$acc.DoCmd.Close(2, "F_AssemblyTemplate", 1)

Write-Host "== F_SystemSettings button names"
$acc.DoCmd.OpenForm("F_SystemSettings", 1)
Start-Sleep -Milliseconds 400
$f = $acc.Forms.Item("F_SystemSettings")
$b = Find-ByCaption $f 104 "Backup All Data"
if ($null -ne $b -and $b.Name -ne "cmdBackupData") { $b.Name = "cmdBackupData"; Write-Host "  renamed cmdBackupData" }
$b = Find-ByCaption $f 104 "Restore All Data"
if ($null -ne $b -and $b.Name -ne "cmdRestoreData") { $b.Name = "cmdRestoreData"; Write-Host "  renamed cmdRestoreData" }
$acc.DoCmd.Close(2, "F_SystemSettings", 1)

Write-Host "== template statuses"
$tcount = $acc.DCount("*", "T_AssemblyTemplates")
Write-Host ("  rows = " + $tcount)
if ($tcount -gt 0) {
    $db.Execute("UPDATE T_AssemblyTemplates SET Status='Complete' WHERE Status='Active'") | Out-Null
    $db.Execute("UPDATE T_AssemblyTemplates SET Status='Pending Approval' WHERE Status='On Hold'") | Out-Null
    $db.Execute("UPDATE T_AssemblyTemplates SET Status='Complete' WHERE Status='Retired'") | Out-Null
    $db.Execute("UPDATE T_AssemblyTemplates SET Status='Draft' WHERE Status IS NULL OR Len(Trim(Status))=0") | Out-Null
    Write-Host "  migrated"
}

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
