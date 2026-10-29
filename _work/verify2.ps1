$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)

foreach ($t in @("T_Customers","T_Suppliers","T_Categories","T_Units","T_Items","T_AssemblyTemplates","T_Warehouses","T_Locations")) {
    Write-Host ($t + " = " + $acc.DCount("*", $t))
}
try {
    $rs = $acc.CurrentDb().OpenRecordset("SELECT CustomerType, Count(*) AS N FROM T_Customers GROUP BY CustomerType")
    while (-not $rs.EOF) { Write-Host ("  type [" + [string]$rs.Fields(0).Value + "] = " + [string]$rs.Fields(1).Value); $rs.MoveNext() }
    $rs.Close()
} catch { Write-Host ("type query err " + $_.Exception.Message) }

Write-Host "== unit count (MoveLast)"
$acc.DoCmd.OpenForm("F_UnitList")
Start-Sleep -Milliseconds 600
$ul = $acc.Forms.Item("F_UnitList")
$rs = $ul.Recordset
$rs.MoveLast()
Write-Host ("  records = " + $rs.RecordCount)
$acc.DoCmd.Close(2, "F_UnitList", 2)

Write-Host "== customer drill"
$acc.Eval('UX_OpenAreaList("customer",True)') | Out-Null
Start-Sleep -Milliseconds 800
$cl = $acc.Forms.Item("F_CustomerList")
Write-Host ("  loaded = " + $acc.CurrentProject.AllForms("F_CustomerList").IsLoaded)
foreach ($c in $cl.Controls) {
    if ($c.ControlType -eq 112) {
        Write-Host ("  subform " + $c.Name + " source=" + [string]$c.SourceObject)
        Write-Host ("    filter=[" + [string]$c.Form.Filter + "] on=" + [string]$c.Form.FilterOn + " recs=" + [string]$c.Form.Recordset.RecordCount)
    }
}
$acc.DoCmd.Close(2, "F_CustomerList", 2)

Write-Host "== warehouse identity"
$wid = $acc.DLookup("WarehouseID", "T_Warehouses")
Write-Host ("  first id = " + [string]$wid)
if (-not ($null -eq $wid)) {
    $missing = [System.Reflection.Missing]::Value
    $where = "WarehouseID=" + [string]$wid
    $acc.DoCmd.OpenForm("F_Warehouses", 0, $missing, $where, $missing, $missing, $missing)
    Start-Sleep -Milliseconds 500
    $w = $acc.Forms.Item("F_Warehouses")
    $actual = ""
    try { $actual = [string]$w.Controls.Item("WarehouseID").Value } catch { $actual = "n/a" }
    Write-Host ("  expected=" + [string]$wid + " actual=" + $actual + " filter=[" + [string]$w.Filter + "] on=" + [string]$w.FilterOn)
    $acc.DoCmd.Close(2, "F_Warehouses", 2)
}

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
