$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)

$acc.DoCmd.OpenForm("F_CustomerList")
Start-Sleep -Milliseconds 1000
$cl = $acc.Forms.Item("F_CustomerList")
Write-Host ("controls=" + $cl.Controls.Count)
$sub = $null
foreach ($c in $cl.Controls) {
    if ($c.ControlType -eq 112) { $sub = $c; Write-Host ("sub=" + $c.Name + " src=" + [string]$c.SourceObject) }
}
if ($null -ne $sub) {
    Write-Host ("subform recs=" + [string]$sub.Form.Recordset.RecordCount)
}
$filterText = [string]$acc.Eval('UX_AreaDrillFilter("customer")')
Write-Host ("filter=[" + $filterText + "]")
if ($null -ne $sub) {
    $sub.Form.Filter = $filterText
    $sub.Form.FilterOn = $true
    Write-Host ("applied filter=[" + [string]$sub.Form.Filter + "] on=" + [string]$sub.Form.FilterOn)
    Write-Host ("filtered recs=" + [string]$sub.Form.Recordset.RecordCount)
}
$acc.DoCmd.Close(2, "F_CustomerList", 2)
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
