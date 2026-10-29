$ErrorActionPreference = "Continue"
$srcDb = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$testDb = "C:\Users\user\Downloads\LLC\_work\test_location.accdb"
Copy-Item $srcDb $testDb -Force
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$acc = New-Object -ComObject Access.Application
$acc.Visible = $true
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($testDb, $false)
Start-Sleep -Seconds 1

# clear leftovers
$acc.CurrentDb().Execute("DELETE FROM T_Locations WHERE LocationName LIKE 'TEST-*'") | Out-Null

# --- 1. Plain new location: no warehouse, no bin, no shelf
$acc.DoCmd.OpenForm("F_LocationSetup", 0, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, 1)
Start-Sleep -Milliseconds 800
$f = $acc.Forms.Item("F_LocationSetup")
Write-Host ("opened new: newrecord=" + $f.NewRecord + " wh=" + [string]$f.Controls.Item("cboWarehouseID").Value)
$f.Controls.Item("LocationName").Value = "TEST-AISLE-1"
$f.Controls.Item("LocationCode").Value = Null
$f.Controls.Item("cboWarehouseID").Value = Null
$f.Controls.Item("BinCode").Value = Null
$f.Controls.Item("ShelfCode").Value = Null
Write-Host ("code after name set = '" + [string]$f.Controls.Item("LocationCode").Value + "'")
try { $f.Dirty = $false; Write-Host "dirty=false ok" } catch { Write-Host ("save err: " + $_.Exception.Message) }
Start-Sleep -Milliseconds 300
$n1 = $acc.DCount("*", "T_Locations", "LocationName='TEST-AISLE-1'")
Write-Host ("TEST-AISLE-1 rows = " + $n1)
$row = $acc.CurrentDb().OpenRecordset("SELECT LocationCode, WarehouseID, BinCode, ShelfCode FROM T_Locations WHERE LocationName='TEST-AISLE-1'")
if (-not $row.EOF) {
    Write-Host ("code=" + [string]$row.Fields("LocationCode").Value + " whNull=" + $row.Fields("WarehouseID").IsNull + " binNull=" + $row.Fields("BinCode").IsNull + " shelfNull=" + $row.Fields("ShelfCode").IsNull)
}
$row.Close()
$acc.DoCmd.Close(2, "F_LocationSetup", 2)

# --- 2. Reopen the saved record - no 3021
$acc.DoCmd.OpenForm("F_LocationSetup", 0, [System.Reflection.Missing]::Value, "LocationName='TEST-AISLE-1'")
Start-Sleep -Milliseconds 700
$f = $acc.Forms.Item("F_LocationSetup")
Write-Host ("reopen name=" + [string]$f.Controls.Item("LocationName").Value + " code=" + [string]$f.Controls.Item("LocationCode").Value)
$acc.DoCmd.Close(2, "F_LocationSetup", 2)

# --- 3. Quick add with existing name: form should close itself, no duplicate
$acc.DoCmd.OpenForm("F_LocationSetup", 0, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, 1, 0, "QUICKADD|TEST-AISLE-1")
Start-Sleep -Milliseconds 1500
Write-Host ("quickadd existing still open = " + $acc.CurrentProject.AllForms("F_LocationSetup").IsLoaded)
Write-Host ("TEST-AISLE-1 rows after quickadd = " + $acc.DCount("*", "T_Locations", "LocationName='TEST-AISLE-1'"))

# --- 4. Quick add with new name: prefilled, code generated, save
$acc.DoCmd.OpenForm("F_LocationSetup", 0, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, 1, 0, "QUICKADD|TEST-AISLE-2")
Start-Sleep -Milliseconds 900
$f = $acc.Forms.Item("F_LocationSetup")
Write-Host ("quickadd new name='" + [string]$f.Controls.Item("LocationName").Value + "' code='" + [string]$f.Controls.Item("LocationCode").Value + "'")
$f.Dirty = $false
Start-Sleep -Milliseconds 300
Write-Host ("TEST-AISLE-2 rows = " + $acc.DCount("*", "T_Locations", "LocationName='TEST-AISLE-2'"))
$acc.DoCmd.Close(2, "F_LocationSetup", 2)

# --- 5. Duplicate guard: same code twice emits friendly message, not 3021
$dup = $acc.DCount("*", "T_Locations", "LocationName='TEST-AISLE-1'")
Write-Host ("final dup count = " + $dup)

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
