$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)

foreach ($name in @("F_CategoriesList","F_TemplateList","F_UnitSetup")) {
    Write-Host ("=== " + $name)
    $acc.DoCmd.OpenForm($name, 1)
    Start-Sleep -Milliseconds 400
    $f = $acc.Forms.Item($name)
    Write-Host ("RecordSource=" + $f.RecordSource + " DefaultView=" + $f.DefaultView)
    foreach ($i in 0..4) {
        try { Write-Host ("Section " + $i + " height=" + $f.Section($i).Height) } catch { Write-Host ("Section " + $i + " none") }
    }
    foreach ($c in $f.Controls) {
        Write-Host ("{0} type={1} src={2} cap={3} L={4} T={5} W={6} H={7}" -f $c.Name, $c.ControlType, $c.ControlSource, $c.Caption, $c.Left, $c.Top, $c.Width, $c.Height)
    }
    $acc.DoCmd.Close(2, $name, 2)
}
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
