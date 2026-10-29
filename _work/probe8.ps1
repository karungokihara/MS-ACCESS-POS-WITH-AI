$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)

$forms = @("F_CustomerList","F_SupplierList","F_ItemSetupList","F_TemplateList","F_WarehouseList","F_ProjectList")
foreach ($name in $forms) {
    $acc.DoCmd.OpenForm($name, 1)
    Start-Sleep -Milliseconds 350
    $f = $acc.Forms.Item($name)
    Write-Host ("=== " + $name)
    foreach ($c in $f.Controls) {
        if ($c.ControlType -eq 112) {
            Write-Host ("  sub " + $c.Name + " source=[" + [string]$c.SourceObject + "] linkChild=[" + [string]$c.LinkChildFields + "] linkMaster=[" + [string]$c.LinkMasterFields + "]")
        }
    }
    $acc.DoCmd.Close(2, $name, 2)
}
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
