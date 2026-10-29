$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$acc.DoCmd.OpenForm("F_CustomerList")
Start-Sleep -Milliseconds 1500
Write-Host ("forms loaded = " + $acc.Forms.Count)
for ($i = 0; $i -lt $acc.Forms.Count; $i++) {
    $ff = $acc.Forms.Item($i)
    Write-Host ("  " + $i + " " + $ff.Name + " controls=" + $ff.Controls.Count + " visible=" + $ff.Visible + " hasmodule=" + $ff.HasModule)
}
$cl = $acc.Forms.Item("F_CustomerList")
Write-Host ("direct: name=" + $cl.Name + " controls=" + $cl.Controls.Count + " recordcount=" + [string]$cl.Recordset.RecordCount)
Write-Host ("IsLoaded=" + $acc.CurrentProject.AllForms("F_CustomerList").IsLoaded)
Write-Host ("Section0 height=" + $cl.Section(0).Height)
$acc.DoCmd.Close(2, "F_CustomerList", 2)
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
