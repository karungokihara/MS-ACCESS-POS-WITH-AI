$ErrorActionPreference = "Continue"
$dbPath = "C:\Users\user\Downloads\LLC\LLC_v3_2026-10-29.accdb"
Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

function Normalize([string]$s) {
    if ($null -eq $s) { return "" }
    return ($s -replace "`r`n", "`n")
}

$acc = New-Object -ComObject Access.Application
$acc.Visible = $false
$acc.AutomationSecurity = 1
$acc.OpenCurrentDatabase($dbPath, $false)

$cm = $acc.VBE.VBProjects.Item(1).VBComponents.Item("Form_F_ItemSetup").CodeModule
$text = Normalize ($cm.Lines(1, $cm.CountOfLines))

$old = Normalize @'
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
'@

$new = Normalize @'
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
'@

if ($text.Contains($new)) {
    Write-Host "already applied"
} elseif ($text.Contains($old)) {
    $text = $text.Replace($old, $new)
    $n = $cm.CountOfLines
    if ($n -gt 0) { $cm.DeleteLines(1, $n) }
    $cm.AddFromString(($text -replace "`n", "`r`n"))
    Write-Host "patched"
} else {
    Write-Host "MISS"
}

$acc.CloseCurrentDatabase()
$acc.Quit()
Write-Host "DONE"
