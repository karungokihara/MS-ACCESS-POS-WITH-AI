$ErrorActionPreference = "Stop"
$root = "C:\Users\user\Downloads\LLC"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.OpenCurrentDatabase("$root\LLC.accdb")
$acc.DoCmd.SetWarnings($false)
try {
    $acc.LoadFromText(5, "M_UX", "$root\_inspect\M_UX.txt")
    Write-Host "M_UX loaded"
    $sql = @"
SELECT W.WarehouseID, W.WarehouseCode, W.WarehouseName, W.Address, W.ContactPerson, (SELECT Count(*) FROM T_Locations AS L WHERE L.WarehouseID = W.WarehouseID) AS LocationCount, UX_WarehouseLocationNames(W.WarehouseID) AS LocationName FROM T_Warehouses AS W ORDER BY W.WarehouseName, W.WarehouseCode;
"@
    $db = $acc.CurrentDb()
    $found = $false
    foreach ($q in $db.QueryDefs) {
        if ($q.Name -eq "Q_WarehouseList") { $q.SQL = $sql; $found = $true }
    }
    if (-not $found) { [void]$db.CreateQueryDef("Q_WarehouseList", $sql) }
    $rs = $db.OpenRecordset("SELECT WarehouseName, LocationName FROM Q_WarehouseList")
    while (-not $rs.EOF) {
        Write-Host ("ROW " + $rs.Fields.Item("WarehouseName").Value + " | " + $rs.Fields.Item("LocationName").Value)
        $rs.MoveNext()
    }
    $rs.Close()

    while ($acc.CurrentProject.AllForms.Item("SF_WarehouseList").IsLoaded) {
        $acc.DoCmd.Close(2, "SF_WarehouseList", 2)
    }
    $acc.DoCmd.OpenForm("SF_WarehouseList", 1)
    $frm = $acc.Forms.Item("SF_WarehouseList")
    $exists = $false
    try { $null = $frm.Controls.Item("LocationName"); $exists = $true } catch {}
    if (-not $exists) {
        $txt = $acc.CreateControl("SF_WarehouseList", 109, 0, "", "LocationName", 300, 4300, 3600, 300)
        $txt.Name = "LocationName"
        $txt.ControlSource = "LocationName"
        try { $txt.TabIndex = 6 } catch {}
        try { $txt.ColumnWidth = 3200 } catch {}
        Write-Host "LocationName column added"
    } else {
        $frm.Controls.Item("LocationName").ControlSource = "LocationName"
        Write-Host "LocationName column already present"
    }
    $acc.DoCmd.Close(2, "SF_WarehouseList", 1)
    $acc.DoCmd.RunCommand(126)
    Write-Host "Compiled"
    $acc.CloseCurrentDatabase()
    $acc.Quit()
} catch {
    Write-Host "FATAL $($_.Exception.Message)"
    try { $acc.Quit() } catch {}
    throw
}
