$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)

foreach ($name in @("F_CustomerSetup","F_AssemblyTemplate","F_SystemSettings")) {
    Write-Host ("=== " + $name)
    $acc.DoCmd.OpenForm($name, 1)
    Start-Sleep -Milliseconds 400
    $f = $acc.Forms.Item($name)
    foreach ($c in $f.Controls) {
        Write-Host ("{0} type={1} src={2} cap={3} L={4} T={5}" -f $c.Name, $c.ControlType, $c.ControlSource, $c.Caption, $c.Left, $c.Top)
    }
    $acc.DoCmd.Close(2, $name, 2)
}
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
