$ErrorActionPreference = "Continue"

$root = "C:\Users\user\Downloads\LLC"
$dbPath = Join-Path $root "LLC_v3_2026-10-29.accdb"
$work = Join-Path $root "_work"

Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

if (-not (Test-Path (Join-Path $root "LLC_v3_before_fixes2.accdb"))) {
    Copy-Item -LiteralPath $dbPath -Destination (Join-Path $root "LLC_v3_before_fixes2.accdb") -Force
    Write-Host "safety copy created"
}

function Normalize([string]$s) {
    if ($null -eq $s) { return "" }
    return ($s -replace "`r`n", "`n")
}
function Get-Code($compName) {
    $cm = $acc.VBE.VBProjects.Item(1).VBComponents.Item($compName).CodeModule
    return $cm.Lines(1, $cm.CountOfLines)
}
function Set-Code($compName, [string]$text) {
    $cm = $acc.VBE.VBProjects.Item(1).VBComponents.Item($compName).CodeModule
    $n = $cm.CountOfLines
    if ($n -gt 0) { $cm.DeleteLines(1, $n) }
    $text = (Normalize $text) -replace "`n", "`r`n"
    $cm.AddFromString($text)
}
function Patch-Code($compName, $pairs) {
    $text = Normalize (Get-Code $compName)
    $changed = $false
    foreach ($p in $pairs) {
        $old = Normalize $p.Old
        $new = Normalize $p.New
        if ($text.Contains($new)) {
            Write-Host ("  skip (already applied): " + $compName + " :: " + $p.Label)
            continue
        }
        if ($text.Contains($old)) {
            $text = $text.Replace($old, $new)
            $changed = $true
            Write-Host ("  patched: " + $compName + " :: " + $p.Label)
        } else {
            Write-Host ("  MISS   : " + $compName + " :: " + $p.Label)
        }
    }
    if ($changed) { Set-Code $compName $text }
}

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)
$db = $acc.CurrentDb()
$proj = $acc.VBE.VBProjects.Item(1)

Write-Host "== modules"
function Import-Module-Bas([string]$name, [string]$file) {
    $existing = $null
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq $name) { $existing = $c } }
    if ($null -ne $existing) {
        $proj.VBComponents.Remove($existing)
        Write-Host ("  removed old " + $name)
    }
    $proj.VBComponents.Import($file) | Out-Null
    Write-Host ("  imported " + $name)
}
Import-Module-Bas "M_Fixes"  (Join-Path $work "M_Fixes.bas")
Import-Module-Bas "M_Backup" (Join-Path $work "M_Backup.bas")

Write-Host "== patch F_Warehouses"
Patch-Code "Form_F_Warehouses" @(
    @{ Label = "restore incoming filter"; Old = @'
    UX_MarkRequired Me, "WarehouseName", "WarehouseCode"
End Sub
'@; New = @'
    UX_MarkRequired Me, "WarehouseName", "WarehouseCode"
    If Len(keepFilter) > 0 Then
        Me.Filter = keepFilter
        Me.FilterOn = keepFilterOn
    End If
End Sub
'@ }
)

Write-Host "== patch F_ItemSetup"
Patch-Code "Form_F_ItemSetup" @(
    @{ Label = "declare max ids"; Old = @'
    Dim enteredText As String
    Dim savedLocationID As Variant

    CreateTypedLocation = False
'@; New = @'
    Dim enteredText As String
    Dim savedLocationID As Variant
    Dim oldMaximumID As Long
    Dim newMaximumID As Long

    CreateTypedLocation = False
'@ }
    @{ Label = "remember max id"; Old = @'
    If Not CloseLocationSetupIfLoaded Then
        Exit Function
    End If

    DoCmd.OpenForm _
        formName:="F_LocationSetup", _
        View:=acNormal, _
        DataMode:=acFormAdd, _
        WindowMode:=acDialog, _
        OpenArgs:="QUICKADD|" & enteredText
'@; New = @'
    If Not CloseLocationSetupIfLoaded Then
        Exit Function
    End If

    oldMaximumID = Nz(DMax("LocationID", "T_Locations"), 0)

    DoCmd.OpenForm _
        formName:="F_LocationSetup", _
        View:=acNormal, _
        DataMode:=acFormAdd, _
        WindowMode:=acDialog, _
        OpenArgs:="QUICKADD|" & enteredText
'@ }
    @{ Label = "resolve saved id"; Old = @'
    savedLocationID = DLookup( _
        "LocationID", _
        "T_Locations", _
        "Nz([LocationName],'')=" & SqlText(enteredText))

    If IsNull(savedLocationID) Then

        MsgBox _
            "The location was not found after Location Setup " & _
            "closed." & vbCrLf & vbCrLf & _
            "Save the location using the same name that you typed.", _
            vbInformation, _
            "Location Not Saved"

        Exit Function

    End If

    CreateTypedLocation = True
'@; New = @'
    Me.cboDefaultLocationID.Requery

    savedLocationID = DLookup( _
        "LocationID", _
        "T_Locations", _
        "StrComp(Trim(Nz([LocationName],''))," & _
            SqlText(enteredText) & ",1)=0")

    If IsNull(savedLocationID) Then

        newMaximumID = Nz(DMax("LocationID", "T_Locations"), 0)

        If newMaximumID > oldMaximumID Then
            savedLocationID = newMaximumID
        End If

    End If

    If IsNull(savedLocationID) Then

        MsgBox _
            "The location was not saved." & vbCrLf & vbCrLf & _
            "Enter the location name in Location Setup and press Save.", _
            vbInformation, _
            "Location Not Saved"

        Exit Function

    End If

    Me.cboDefaultLocationID.Value = CLng(savedLocationID)
    CreateTypedLocation = True
'@ }
)

Write-Host "== patch F_AssemblyTemplate"
Patch-Code "Form_F_AssemblyTemplate" @(
    @{ Label = "lock on current record"; Old = @'
    ApplyTemplateDefaults
    ConfigureContextControls
    SetSubformAvailability

    On Error GoTo 0
'@; New = @'
    ApplyTemplateDefaults
    ConfigureContextControls
    SetSubformAvailability
    ApplyPartNumberLock

    On Error GoTo 0
'@ }
)

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "PATCHES DONE"
