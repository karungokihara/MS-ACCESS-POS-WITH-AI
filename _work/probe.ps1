$ErrorActionPreference = "Continue"
$src = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$scratch = Join-Path $env:TEMP "llc_probe.accdb"
Copy-Item -LiteralPath $src -Destination $scratch -Force

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($scratch, $false)

foreach ($formName in @("F_AssemblyTemplate", "F_SystemSettings")) {
    Write-Host ("=== " + $formName)
    $acc.DoCmd.OpenForm($formName, 1)
    Start-Sleep -Milliseconds 500
    $f = $acc.Forms.Item($formName)
    Write-Host ("Form InsideWidth=" + $f.InsideWidth + " InsideHeight=" + $f.InsideHeight + " Width=" + $f.Width + " Section0H=" + $f.Section(0).Height + " Section1H=" + $f.Section(1).Height)
    foreach ($c in $f.Controls) {
        Write-Host ("{0}`t{1}`tL={2}`tT={3}`tW={4}`tH={5}`tVis={6}" -f $c.Name, $c.ControlType, $c.Left, $c.Top, $c.Width, $c.Height, $c.Visible)
    }
    $acc.DoCmd.Close(2, $formName, 2)
}

foreach ($formName in @("F_CustomerList", "F_ItemSetupList", "F_WarehouseList", "F_ProjectList", "F_TemplateList", "F_SupplierList", "F_InventoryList")) {
    Write-Host ("=== " + $formName)
    $acc.DoCmd.OpenForm($formName, 1)
    Start-Sleep -Milliseconds 400
    $f = $acc.Forms.Item($formName)
    Write-Host ("Form ScrollBars=" + $f.ScrollBars + " Width=" + $f.Width + " DefaultView=" + $f.DefaultView)
    foreach ($c in $f.Controls) {
        if ($c.ControlType -eq 112) {
            Write-Host ("SUBFORM {0} L={1} T={2} W={3} H={4} Src={5}" -f $c.Name, $c.Left, $c.Top, $c.Width, $c.Height, $c.SourceObject)
        }
    }
    $acc.DoCmd.Close(2, $formName, 2)
}

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
