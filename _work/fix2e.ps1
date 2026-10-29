$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$work = "C:\Users\user\Downloads\LLC\_work"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$proj = $acc.VBE.VBProjects.Item(1)

function Import-Module-Bas([string]$name, [string]$file) {
    $existing = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq $name) { $existing = $c } }
    if ($null -ne $existing) { $proj.VBComponents.Remove($existing) }
    $proj.VBComponents.Import($file) | Out-Null
    Write-Host ("imported " + $name)
}
Import-Module-Bas "M_Fixes"  (Join-Path $work "M_Fixes.bas")
Import-Module-Bas "M_Backup" (Join-Path $work "M_Backup.bas")

Write-Host "== compile"
try { $acc.DoCmd.RunCommand(126) | Out-Null } catch { Write-Host ("compile cmd error: " + $_.Exception.Message) }
Write-Host ("IsCompiled = " + $acc.IsCompiled)

Write-Host "== eval checks"
$checks = @(
    'UX_AreaTotal("unit")',
    'UX_AreaIncomplete("unit")',
    'UX_AreaReady("unit")',
    'UX_AreaTotal("customer")',
    'UX_AreaIncomplete("customer")',
    'UX_AreaIncomplete("template")',
    'UX_AreaDrafts("template")',
    'UX_GeneratePartNumber("ABC","2")',
    'UX_AreaDrillFilter("unit")',
    'UX_AreaCriteria("template")'
)
foreach ($expr in $checks) {
    try { Write-Host ("  " + $expr + " = " + $acc.Eval($expr)) } catch { Write-Host ("  " + $expr + " ERR: " + $_.Exception.Message) }
}

Write-Host "== open F_UnitList normally"
try {
    $acc.DoCmd.OpenForm("F_UnitList")
    Start-Sleep -Milliseconds 800
    $ul = $acc.Forms.Item("F_UnitList")
    Write-Host ("  records = " + $ul.Recordset.RecordCount)
    $names = @()
    foreach ($fd in $ul.Recordset.Fields) { $names += $fd.Name }
    Write-Host ("  fields = " + ($names -join ","))
    $ctlNames = @()
    foreach ($c in $ul.Controls) { $ctlNames += $c.Name }
    Write-Host ("  controls = " + ($ctlNames -join ","))
    $acc.DoCmd.Close(2, "F_UnitList", 2)
} catch { Write-Host ("  F_UnitList failed: " + $_.Exception.Message) }

Write-Host "== drill-down unit"
try {
    $acc.Eval('UX_OpenAreaList("unit",False)')
    Start-Sleep -Milliseconds 600
    if ($acc.CurrentProject.AllForms("F_UnitList").IsLoaded) {
        Write-Host ("  drilled records = " + $acc.Forms.Item("F_UnitList").Recordset.RecordCount)
        $acc.DoCmd.Close(2, "F_UnitList", 2)
    } else {
        Write-Host "  F_UnitList not loaded"
    }
} catch { Write-Host ("  drill failed: " + $_.Exception.Message) }

Write-Host "== drill-down incomplete customers"
try {
    $acc.Eval('UX_OpenAreaList("customer",True)')
    Start-Sleep -Milliseconds 600
    if ($acc.CurrentProject.AllForms("F_CustomerList").IsLoaded) {
        $cl = $acc.Forms.Item("F_CustomerList")
        $filterText = ""
        foreach ($c in $cl.Controls) {
            if ($c.ControlType -eq 112) {
                try { $filterText = [string]$c.Form.Filter } catch {}
            }
        }
        Write-Host ("  subform filter = " + $filterText)
        $acc.DoCmd.Close(2, "F_CustomerList", 2)
    } else {
        Write-Host "  F_CustomerList not loaded"
    }
} catch { Write-Host ("  drill failed: " + $_.Exception.Message) }

Write-Host "== warehouse record identity"
try {
    $firstID = $acc.DLookup("WarehouseID", "T_Warehouses")
    if (-not ($null -eq $firstID)) {
        $missing = [System.Reflection.Missing]::Value
        $where = "WarehouseID=" + [string]$firstID
        $acc.DoCmd.OpenForm("F_Warehouses", 0, $missing, $where, $missing, $missing, $missing)
        Start-Sleep -Milliseconds 500
        $w = $acc.Forms.Item("F_Warehouses")
        Write-Host ("  expected=" + [string]$firstID + " actual=" + [string]$w.Controls.Item("WarehouseID").Value)
        $acc.DoCmd.Close(2, "F_Warehouses", 2)
    } else {
        Write-Host "  no warehouse rows"
    }
} catch { Write-Host ("  warehouse test failed: " + $_.Exception.Message) }

Write-Host "== backup"
try {
    $backupPath = [string]$acc.Eval('UX_BackupAllData(True)')
    Write-Host ("  backup = " + $backupPath)
    Write-Host ("  exists = " + (Test-Path -LiteralPath $backupPath))
} catch { Write-Host ("  backup ERR: " + $_.Exception.Message) }

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
