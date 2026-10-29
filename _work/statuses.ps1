$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
Write-Host ("templates=" + $acc.DCount("*", "T_AssemblyTemplates"))
$rs = $acc.CurrentDb().OpenRecordset("SELECT AssemblyTemplateID, TemplateCode, Status FROM T_AssemblyTemplates", 4)
Write-Host ("count=" + $rs.RecordCount)
while (-not $rs.EOF) {
    $id = $rs.Fields.Item(0).Value
    $code = $rs.Fields.Item(1).Value
    $st = $rs.Fields.Item(2).Value
    $isNull = $rs.Fields.Item(2).Value -eq $null
    Write-Host ("id=" + [string]$id + " code=" + [string]$code + " status=" + [string]$st + " isnull=" + [string]$isNull)
    $rs.MoveNext()
}
$rs.Close()
$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
