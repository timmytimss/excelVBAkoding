param(
    [string]$Path,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# Kolonnevelger — alt-i-en installer/oppdaterer.
# Bygger de fire VBA-komponentene fra kildekoden innebygd nedenfor, i en
# midlertidig arbeidsbok, og installerer/oppdaterer dem deretter i mal-filen.
# Ingen andre filer trengs eller etterlates.
# ============================================================================

$oSlash = [char]0x00F8   # o with stroke, brukt for aa unngaa kodings-problemer i norske tegn i denne fila

$capLoadoutsHint = "Loadouts (venstreklikk = bruk, h" + $oSlash + "yreklikk = rediger):"
$capDeleteSlot = "Slett (t" + $oSlash + "m denne plassen)"

# ---- Innebygd VBA-kildekode ----

$moduleCode = @'
Option Explicit

Public Const COLUMN_PICKER_VERSION As String = "1.0.1"

Private Const LOADOUT_SHEET As String = "ColumnLoadouts_Hidden"

' Sett denne makroen som "Assign Macro" på en Form Control-knapp i hvert ark.
' Knappen forutsetter at arket har nøyaktig én tabell (ListObject).
Public Sub ShowColumnPicker()
    Dim ws As Worksheet
    Dim tbl As ListObject

    If TypeName(Application.Caller) = "String" Then
        Set ws = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Worksheet
    Else
        Set ws = ActiveSheet
    End If

    If ws.ListObjects.Count = 0 Then
        MsgBox "Fant ingen tabell i arket """ & ws.Name & """.", vbExclamation
        Exit Sub
    ElseIf ws.ListObjects.Count > 1 Then
        MsgBox "Fant flere tabeller i arket """ & ws.Name & """. Denne knappen forutsetter kun én tabell per ark.", vbExclamation
        Exit Sub
    End If

    Set tbl = ws.ListObjects(1)

    Dim frm As New frmColumnPicker
    Set frm.TargetTable = tbl
    frm.LoadColumns
    frm.Show
End Sub

' ---- Loadout-lagring (8 faste plasser per ark+tabell, i skjult ark i denne arbeidsboken) ----

Private Function GetLoadoutSheet() As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(LOADOUT_SHEET)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sh.Name = LOADOUT_SHEET
        sh.Visible = xlSheetVeryHidden
        sh.Range("A1:F1").Value = Array("Ark", "Tabell", "Slot", "Navn", "Farge", "Kolonner")
    End If
    Set GetLoadoutSheet = sh
End Function

Private Function FindLoadoutRow(ByVal sh As Worksheet, ByVal sheetName As String, ByVal tableName As String, ByVal slotIndex As Long) As Long
    Dim lastRow As Long, r As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If sh.Cells(r, 1).Value = sheetName And sh.Cells(r, 2).Value = tableName And sh.Cells(r, 3).Value = slotIndex Then
            FindLoadoutRow = r
            Exit Function
        End If
    Next r
    FindLoadoutRow = 0
End Function

Public Sub SaveLoadoutSlot(ByVal sheetName As String, ByVal tableName As String, ByVal slotIndex As Long, ByVal loadoutName As String, ByVal colorValue As Long, ByVal visibleColumnsCsv As String)
    Dim sh As Worksheet, r As Long, lastRow As Long
    Set sh = GetLoadoutSheet()
    r = FindLoadoutRow(sh, sheetName, tableName, slotIndex)
    If r = 0 Then
        lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
        r = lastRow + 1
        sh.Cells(r, 1).Value = sheetName
        sh.Cells(r, 2).Value = tableName
        sh.Cells(r, 3).Value = slotIndex
    End If
    sh.Cells(r, 4).Value = loadoutName
    sh.Cells(r, 5).Value = colorValue
    sh.Cells(r, 6).Value = visibleColumnsCsv
End Sub

Public Sub ClearLoadoutSlot(ByVal sheetName As String, ByVal tableName As String, ByVal slotIndex As Long)
    Dim sh As Worksheet, r As Long
    Set sh = GetLoadoutSheet()
    r = FindLoadoutRow(sh, sheetName, tableName, slotIndex)
    If r > 0 Then sh.Rows(r).Delete
End Sub

Public Function GetLoadoutSlotName(ByVal sheetName As String, ByVal tableName As String, ByVal slotIndex As Long) As String
    Dim sh As Worksheet, r As Long
    Set sh = GetLoadoutSheet()
    r = FindLoadoutRow(sh, sheetName, tableName, slotIndex)
    If r = 0 Then
        GetLoadoutSlotName = ""
    Else
        GetLoadoutSlotName = CStr(sh.Cells(r, 4).Value)
    End If
End Function

Public Function GetLoadoutSlotColor(ByVal sheetName As String, ByVal tableName As String, ByVal slotIndex As Long) As Long
    Dim sh As Worksheet, r As Long
    Set sh = GetLoadoutSheet()
    r = FindLoadoutRow(sh, sheetName, tableName, slotIndex)
    If r = 0 Then
        GetLoadoutSlotColor = 0
    Else
        GetLoadoutSlotColor = CLng(sh.Cells(r, 5).Value)
    End If
End Function

Public Function GetLoadoutSlotColumnsCsv(ByVal sheetName As String, ByVal tableName As String, ByVal slotIndex As Long) As String
    Dim sh As Worksheet, r As Long
    Set sh = GetLoadoutSheet()
    r = FindLoadoutRow(sh, sheetName, tableName, slotIndex)
    If r = 0 Then
        GetLoadoutSlotColumnsCsv = ""
    Else
        GetLoadoutSlotColumnsCsv = CStr(sh.Cells(r, 6).Value)
    End If
End Function

' ---- Vis/skjul kolonner ----

Public Sub ApplyVisibleColumns(ByVal tbl As ListObject, ByVal visibleNames As Collection)
    Dim col As ListColumn
    Dim nm As Variant
    Dim isVisible As Boolean

    For Each col In tbl.ListColumns
        isVisible = False
        For Each nm In visibleNames
            If nm = col.Name Then
                isVisible = True
                Exit For
            End If
        Next nm
        col.Range.EntireColumn.Hidden = Not isVisible
    Next col
End Sub

Public Sub ResetAllColumns(ByVal tbl As ListObject)
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        col.Range.EntireColumn.Hidden = False
    Next col
End Sub
'@

$classCode = @'
Option Explicit

Public WithEvents Chk As MSForms.CheckBox
Public ColumnName As String
Public TargetTable As ListObject

Private Sub Chk_Click()
    If TargetTable Is Nothing Then Exit Sub
    TargetTable.ListColumns(ColumnName).Range.EntireColumn.Hidden = Not Chk.Value
End Sub
'@

$formCode = @'
Option Explicit

Public TargetTable As ListObject
Private mHandlers As Collection

Private Const MAX_FRAME_HEIGHT As Single = 200

Public Sub LoadColumns()
    Dim col As ListColumn
    Dim chk As MSForms.CheckBox
    Dim h As clsColumnCheckHandler
    Dim idx As Long
    Dim colsPerRow As Long
    Dim colWidth As Single
    Dim rowHeight As Single
    Dim r As Long, c As Long
    Dim displayName As String
    Dim totalRows As Long
    Dim neededFrameHeight As Single

    Set mHandlers = New Collection
    Me.Caption = "Velg kolonner"

    cmdReset.Font.Bold = True
    cmdReset.ForeColor = vbRed

    If Not TargetTable Is Nothing Then
        lblTitle.Caption = "Ark: " & TargetTable.Parent.Name & "   Tabell: " & TargetTable.Name
    End If

    colsPerRow = 3
    rowHeight = 20
    colWidth = (frameColumns.InsideWidth - 12) / colsPerRow

    idx = 0
    For Each col In TargetTable.ListColumns
        idx = idx + 1
        r = (idx - 1) \ colsPerRow
        c = (idx - 1) Mod colsPerRow

        displayName = Replace(Replace(col.Name, vbLf, " "), vbCr, "")

        Set chk = frameColumns.Controls.Add("Forms.CheckBox.1", "chkCol" & idx, True)
        With chk
            .Caption = displayName
            .WordWrap = False
            .Left = 6 + c * colWidth
            .Top = 6 + r * rowHeight
            .Width = colWidth - 6
            .Height = rowHeight - 2
            .Value = Not col.Range.EntireColumn.Hidden
            .Tag = col.Name
        End With

        Set h = New clsColumnCheckHandler
        Set h.Chk = chk
        h.ColumnName = col.Name
        Set h.TargetTable = TargetTable
        mHandlers.Add h
    Next col

    totalRows = ((idx - 1) \ colsPerRow) + 1
    neededFrameHeight = 6 + totalRows * rowHeight + 16
    If neededFrameHeight > MAX_FRAME_HEIGHT Then neededFrameHeight = MAX_FRAME_HEIGHT

    frameColumns.Height = neededFrameHeight
    If (6 + totalRows * rowHeight + 10) > neededFrameHeight Then
        frameColumns.ScrollBars = fmScrollBarsVertical
        frameColumns.ScrollHeight = 6 + totalRows * rowHeight + 10
    Else
        frameColumns.ScrollBars = fmScrollBarsNone
        frameColumns.ScrollHeight = 0
    End If

    PositionLowerControls

    RefreshLoadoutButtons
End Sub

Private Sub PositionLowerControls()
    Const BTN_W As Single = 200
    Const BTN_H As Single = 30
    Const BTN_GAP As Single = 6
    Const CHROME_BUFFER As Single = 56

    Dim gridLefts(0 To 2) As Single
    Dim slotCtrls(0 To 8) As MSForms.CommandButton
    Dim topCursor As Single
    Dim gi As Long, gr As Long, gc As Long
    Dim k As Long

    gridLefts(0) = 10: gridLefts(1) = 220: gridLefts(2) = 430

    topCursor = frameColumns.Top + frameColumns.Height + 10

    lblLoadouts.Left = 10
    lblLoadouts.Top = topCursor
    lblLoadouts.Width = 620
    lblLoadouts.Height = 18
    lblLoadouts.Caption = "Loadouts (venstreklikk = bruk, høyreklikk = rediger):"

    topCursor = topCursor + 22

    Set slotCtrls(0) = cmdReset
    For k = 1 To 8
        Set slotCtrls(k) = Me.Controls("btnLoadout" & k)
    Next k

    For gi = 0 To 8
        gr = gi \ 3
        gc = gi Mod 3
        With slotCtrls(gi)
            .Left = gridLefts(gc)
            .Top = topCursor + gr * (BTN_H + BTN_GAP)
            .Width = BTN_W
            .Height = BTN_H
        End With
    Next gi

    Dim gridBottom As Single
    gridBottom = topCursor + 2 * (BTN_H + BTN_GAP) + BTN_H

    Me.Height = gridBottom + CHROME_BUFFER
End Sub

Public Sub RefreshLoadoutButtons()
    Dim i As Long
    Dim nm As String
    Dim clr As Long
    Dim btn As MSForms.CommandButton

    For i = 1 To 8
        Set btn = Me.Controls("btnLoadout" & i)
        nm = GetLoadoutSlotName(TargetTable.Parent.Name, TargetTable.Name, i)
        If nm = "" Then
            btn.Caption = "Loadout " & i
            btn.BackColor = &H8000000F
        Else
            btn.Caption = nm
            clr = GetLoadoutSlotColor(TargetTable.Parent.Name, TargetTable.Name, i)
            If clr <= 0 Then
                btn.BackColor = &H8000000F
            Else
                btn.BackColor = clr
            End If
        End If
    Next i
End Sub

Public Function GetCheckedColumnNamesCsv() As String
    Dim h As clsColumnCheckHandler
    Dim csv As String
    csv = ""
    For Each h In mHandlers
        If h.Chk.Value Then
            If csv <> "" Then csv = csv & "|"
            csv = csv & h.ColumnName
        End If
    Next h
    GetCheckedColumnNamesCsv = csv
End Function

Private Sub ApplyLoadoutSlot(ByVal slotIndex As Long)
    Dim csv As String
    Dim nm As String
    Dim parts() As String
    Dim h As clsColumnCheckHandler
    Dim i As Long
    Dim found As Boolean

    nm = GetLoadoutSlotName(TargetTable.Parent.Name, TargetTable.Name, slotIndex)
    If nm = "" Then
        MsgBox "Denne loadout-plassen er tom." & vbCrLf & "Høyreklikk knappen for å lagre gjeldende kolonnevalg her.", vbInformation
        Exit Sub
    End If

    csv = GetLoadoutSlotColumnsCsv(TargetTable.Parent.Name, TargetTable.Name, slotIndex)
    parts = Split(csv, "|")

    For Each h In mHandlers
        found = False
        For i = LBound(parts) To UBound(parts)
            If parts(i) = h.ColumnName Then
                found = True
                Exit For
            End If
        Next i
        h.Chk.Value = found
    Next h
End Sub

Private Sub EditLoadoutSlot(ByVal slotIndex As Long)
    Dim f As New frmLoadoutEdit
    f.SheetName = TargetTable.Parent.Name
    f.TableName = TargetTable.Name
    f.SlotIndex = slotIndex
    Set f.ParentPicker = Me
    f.InitEdit
    f.Show vbModal
End Sub

Private Sub cmdReset_Click()
    Dim h As clsColumnCheckHandler
    For Each h In mHandlers
        h.Chk.Value = True
    Next h
    ResetAllColumns TargetTable
End Sub

Private Sub btnLoadout1_Click(): ApplyLoadoutSlot 1: End Sub
Private Sub btnLoadout2_Click(): ApplyLoadoutSlot 2: End Sub
Private Sub btnLoadout3_Click(): ApplyLoadoutSlot 3: End Sub
Private Sub btnLoadout4_Click(): ApplyLoadoutSlot 4: End Sub
Private Sub btnLoadout5_Click(): ApplyLoadoutSlot 5: End Sub
Private Sub btnLoadout6_Click(): ApplyLoadoutSlot 6: End Sub
Private Sub btnLoadout7_Click(): ApplyLoadoutSlot 7: End Sub
Private Sub btnLoadout8_Click(): ApplyLoadoutSlot 8: End Sub

Private Sub btnLoadout1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then EditLoadoutSlot 1
End Sub
Private Sub btnLoadout2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then EditLoadoutSlot 2
End Sub
Private Sub btnLoadout3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then EditLoadoutSlot 3
End Sub
Private Sub btnLoadout4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then EditLoadoutSlot 4
End Sub
Private Sub btnLoadout5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then EditLoadoutSlot 5
End Sub
Private Sub btnLoadout6_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then EditLoadoutSlot 6
End Sub
Private Sub btnLoadout7_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then EditLoadoutSlot 7
End Sub
Private Sub btnLoadout8_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then EditLoadoutSlot 8
End Sub
'@

$editFormCode = @'
Option Explicit

Public SheetName As String
Public TableName As String
Public SlotIndex As Long
Public ParentPicker As frmColumnPicker

Private mSelectedColor As Long

Public Sub InitEdit()
    Dim nm As String
    Me.Caption = "Rediger loadout " & SlotIndex
    lblSlotInfo.Caption = "Loadout " & SlotIndex

    nm = GetLoadoutSlotName(SheetName, TableName, SlotIndex)
    If nm = "" Then
        txtName.Text = "Loadout " & SlotIndex
        mSelectedColor = 0
    Else
        txtName.Text = nm
        mSelectedColor = GetLoadoutSlotColor(SheetName, TableName, SlotIndex)
    End If
End Sub

Private Sub SelectColor(ByVal clr As Long)
    mSelectedColor = clr
End Sub

Private Sub btnColorNone_Click(): SelectColor 0: End Sub
Private Sub btnColor1_Click(): SelectColor RGB(255, 199, 206): End Sub
Private Sub btnColor2_Click(): SelectColor RGB(255, 235, 156): End Sub
Private Sub btnColor3_Click(): SelectColor RGB(198, 239, 206): End Sub
Private Sub btnColor4_Click(): SelectColor RGB(189, 215, 238): End Sub
Private Sub btnColor5_Click(): SelectColor RGB(255, 192, 0): End Sub
Private Sub btnColor6_Click(): SelectColor RGB(146, 208, 80): End Sub
Private Sub btnColor7_Click(): SelectColor RGB(0, 176, 240): End Sub
Private Sub btnColor8_Click(): SelectColor RGB(204, 153, 255): End Sub
Private Sub btnColor9_Click(): SelectColor RGB(217, 217, 217): End Sub

Private Sub cmdSaveSlot_Click()
    Dim nm As String
    Dim csv As String

    nm = Trim(txtName.Text)
    If nm = "" Then
        MsgBox "Skriv inn et navn.", vbExclamation
        Exit Sub
    End If

    csv = ParentPicker.GetCheckedColumnNamesCsv()
    SaveLoadoutSlot SheetName, TableName, SlotIndex, nm, mSelectedColor, csv

    ParentPicker.RefreshLoadoutButtons
    Unload Me
End Sub

Private Sub cmdDeleteSlot_Click()
    If MsgBox("Tømme loadout-plass " & SlotIndex & "?", vbQuestion + vbYesNo) = vbYes Then
        ClearLoadoutSlot SheetName, TableName, SlotIndex
        ParentPicker.RefreshLoadoutButtons
        Unload Me
    End If
End Sub

Private Sub cmdCancelEdit_Click()
    Unload Me
End Sub
'@

function Get-PickerVersion {
    param([string]$Code)
    if ($Code -match 'COLUMN_PICKER_VERSION\s*As\s*String\s*=\s*"([^"]+)"') {
        return $Matches[1]
    }
    return $null
}

$sourceVersion = Get-PickerVersion $moduleCode

# ---- 1. Finn maalfilen ----

if (-not $Path) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel-filer (*.xlsm)|*.xlsm|Alle filer (*.*)|*.*"
    $dlg.Title = "Velg Excel-fil som skal faa Kolonnevelger"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Output "Avbrutt."
        exit 0
    }
    $Path = $dlg.FileName
}

if (-not (Test-Path $Path)) {
    Write-Output "FEIL: Fant ikke filen: $Path"
    exit 1
}
$Path = (Resolve-Path $Path).Path

# ---- 2. Finn en aapen Excel-instans som har filen, ellers aapne den selv ----

$excel = $null
$wb = $null

try {
    $running = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    foreach ($w in $running.Workbooks) {
        if ($w.FullName -eq $Path) { $wb = $w; $excel = $running; break }
    }
    if ($null -eq $excel) { $excel = $running }
} catch {
    $excel = $null
}

if ($null -eq $excel) {
    $excel = New-Object -ComObject Excel.Application
}
$excel.Visible = $true
$excel.DisplayAlerts = $false

if ($null -eq $wb) {
    $wb = $excel.Workbooks.Open($Path)
}
$wb.Activate()

try {
    $vbproj = $wb.VBProject
    $null = $vbproj.VBComponents.Count
} catch {
    Write-Output "FEIL: Faar ikke tilgang til VBA-prosjektet."
    Write-Output "Sjekk at 'Klarer tilgang til VBA-prosjektobjektmodellen' er huket av i Excel sitt Klareringssenter (Filer > Alternativer > Klareringssenter > Innstillinger for klareringssenteret > Makroinnstillinger)."
    exit 1
}

if ($Uninstall) {
    # Avinstallerer - fjerner KUN de noyaktige, kjente komponentnavnene og
    # knappen med denne makroens EGET OnAction-navn, aldri et generisk
    # filter (se lesson om Kolonnevelger-knapp-hendelsen i minnet - et ark
    # kan ha flere uavhengig installerte makroers knapper side om side).
    # Den skjulte ColumnLoadouts_Hidden-fana (loadouts) la staa urort, med
    # vilje - en senere reinstallasjon gjenbruker da gamle loadouts.
    Write-Output "Fjerner Kolonnevelger fra $($wb.Name) ..."
    foreach ($navn in @("modColumnPicker", "clsColumnCheckHandler", "frmColumnPicker", "frmLoadoutEdit")) {
        try {
            $eksisterende = $vbproj.VBComponents.Item($navn)
            $vbproj.VBComponents.Remove($eksisterende)
            Write-Output "  Fjernet $navn"
        } catch {}
    }
    try {
        foreach ($ws in $wb.Worksheets) {
            foreach ($btn in @($ws.Buttons())) {
                if ($btn.OnAction -eq "ShowColumnPicker" -or $btn.OnAction -like "*!ShowColumnPicker") { $btn.Delete() }
            }
        }
    } catch {}
    try { $wb.Save() } catch {
        Write-Output "FEIL ved lagring: $($_.Exception.Message)"
        exit 1
    }
    Write-Output "Ferdig - Kolonnevelger er fjernet fra $($wb.Name)."
    exit 0
}

# ---- 3. Sjekk hva som allerede er installert ----

$existingModule = $null
foreach ($comp in $vbproj.VBComponents) {
    if ($comp.Name -eq "modColumnPicker") { $existingModule = $comp }
}

$installedVersion = $null
if ($existingModule) {
    $existingCode = $existingModule.CodeModule.Lines(1, $existingModule.CodeModule.CountOfLines)
    $installedVersion = Get-PickerVersion $existingCode
}

$isFreshInstall = ($null -eq $existingModule)

if (-not $isFreshInstall -and $installedVersion -eq $sourceVersion) {
    Write-Output "Kolonnevelger er allerede installert og oppdatert (versjon $installedVersion) i $($wb.Name)."
    exit 0
}

if ($isFreshInstall) {
    Write-Output "Installerer Kolonnevelger (versjon $sourceVersion) i $($wb.Name) ..."
} else {
    Write-Output "Oppdaterer Kolonnevelger: versjon $installedVersion -> $sourceVersion i $($wb.Name) ..."
}

# ---- 4. Bygg de fire komponentene i en midlertidig arbeidsbok, eksporter til en temp-mappe ----

$tempDir = Join-Path $env:TEMP ("Kolonnevelger_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir | Out-Null

try {
    $buildWb = $excel.Workbooks.Add()
    $buildProj = $buildWb.VBProject

    $mod = $buildProj.VBComponents.Add(1)
    $mod.Name = "modColumnPicker"
    $mod.CodeModule.AddFromString($moduleCode)

    $cls = $buildProj.VBComponents.Add(2)
    $cls.Name = "clsColumnCheckHandler"
    $cls.CodeModule.AddFromString($classCode)

    $frmComp = $buildProj.VBComponents.Add(3)
    $frmComp.Name = "frmColumnPicker"
    $frmComp.Properties("Width").Value = 650
    $frmComp.Properties("Height").Value = 390
    $frmComp.Properties("Caption").Value = "Velg kolonner"
    $designer = $frmComp.Designer

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Left = 10; $lblTitle.Top = 10; $lblTitle.Width = 620; $lblTitle.Height = 18

    $frameColumns = $designer.Controls.Add("Forms.Frame.1", "frameColumns", $true)
    $frameColumns.Left = 10; $frameColumns.Top = 34; $frameColumns.Width = 630; $frameColumns.Height = 200
    $frameColumns.Caption = "Kolonner"
    $frameColumns.ScrollBars = 2

    $lblLoadouts = $designer.Controls.Add("Forms.Label.1", "lblLoadouts", $true)
    $lblLoadouts.Left = 10; $lblLoadouts.Top = 244; $lblLoadouts.Width = 150; $lblLoadouts.Height = 18
    $lblLoadouts.Caption = $capLoadoutsHint

    $gridLefts = @(10, 220, 430)
    $gridTops = @(264, 300, 336)

    $cmdReset = $designer.Controls.Add("Forms.CommandButton.1", "cmdReset", $true)
    $cmdReset.Left = $gridLefts[0]; $cmdReset.Top = $gridTops[0]; $cmdReset.Width = 200; $cmdReset.Height = 30
    $cmdReset.Caption = "RESET (VIS ALLE)"

    $cells = @(@(0,1), @(0,2), @(1,0), @(1,1), @(1,2), @(2,0), @(2,1), @(2,2))
    for ($i = 0; $i -lt 8; $i++) {
        $rr = $cells[$i][0]; $cc = $cells[$i][1]
        $slot = $i + 1
        $btn = $designer.Controls.Add("Forms.CommandButton.1", "btnLoadout$slot", $true)
        $btn.Left = $gridLefts[$cc]; $btn.Top = $gridTops[$rr]; $btn.Width = 200; $btn.Height = 30
        $btn.Caption = "Loadout $slot"
    }

    $frmComp.CodeModule.AddFromString($formCode)

    $editComp = $buildProj.VBComponents.Add(3)
    $editComp.Name = "frmLoadoutEdit"
    $editComp.Properties("Width").Value = 340
    $editComp.Properties("Height").Value = 260
    $editComp.Properties("Caption").Value = "Rediger loadout"
    $editDesigner = $editComp.Designer

    $lblSlotInfo = $editDesigner.Controls.Add("Forms.Label.1", "lblSlotInfo", $true)
    $lblSlotInfo.Left = 10; $lblSlotInfo.Top = 8; $lblSlotInfo.Width = 300; $lblSlotInfo.Height = 18

    $lblNameCap = $editDesigner.Controls.Add("Forms.Label.1", "lblNameCap", $true)
    $lblNameCap.Left = 10; $lblNameCap.Top = 36; $lblNameCap.Width = 50; $lblNameCap.Height = 18
    $lblNameCap.Caption = "Navn:"

    $txtName = $editDesigner.Controls.Add("Forms.TextBox.1", "txtName", $true)
    $txtName.Left = 64; $txtName.Top = 34; $txtName.Width = 266; $txtName.Height = 20

    $lblColorCap = $editDesigner.Controls.Add("Forms.Label.1", "lblColorCap", $true)
    $lblColorCap.Left = 10; $lblColorCap.Top = 64; $lblColorCap.Width = 60; $lblColorCap.Height = 18
    $lblColorCap.Caption = "Farge:"

    $colorLefts = @(10, 54, 98, 142, 186)
    $colorNames = @("btnColorNone", "btnColor1", "btnColor2", "btnColor3", "btnColor4")
    $colorCaptions = @("Ingen", "", "", "", "")
    $colorRGB1 = @($null, @(255,199,206), @(255,235,156), @(198,239,206), @(189,215,238))
    for ($i = 0; $i -lt 5; $i++) {
        $b = $editDesigner.Controls.Add("Forms.CommandButton.1", $colorNames[$i], $true)
        $b.Left = $colorLefts[$i]; $b.Top = 84; $b.Width = 40; $b.Height = 24
        $b.Caption = $colorCaptions[$i]
        if ($colorRGB1[$i]) {
            $rgb = $colorRGB1[$i]
            $b.BackColor = [System.Convert]::ToInt64($rgb[2]) * 65536 + [System.Convert]::ToInt64($rgb[1]) * 256 + [System.Convert]::ToInt64($rgb[0])
        }
    }

    $colorNames2 = @("btnColor5", "btnColor6", "btnColor7", "btnColor8", "btnColor9")
    $colorRGB2 = @(@(255,192,0), @(146,208,80), @(0,176,240), @(204,153,255), @(217,217,217))
    for ($i = 0; $i -lt 5; $i++) {
        $b = $editDesigner.Controls.Add("Forms.CommandButton.1", $colorNames2[$i], $true)
        $b.Left = $colorLefts[$i]; $b.Top = 112; $b.Width = 40; $b.Height = 24
        $rgb = $colorRGB2[$i]
        $b.BackColor = [System.Convert]::ToInt64($rgb[2]) * 65536 + [System.Convert]::ToInt64($rgb[1]) * 256 + [System.Convert]::ToInt64($rgb[0])
    }

    $cmdSaveSlot = $editDesigner.Controls.Add("Forms.CommandButton.1", "cmdSaveSlot", $true)
    $cmdSaveSlot.Left = 10; $cmdSaveSlot.Top = 144; $cmdSaveSlot.Width = 320; $cmdSaveSlot.Height = 28
    $cmdSaveSlot.Caption = "Lagre gjeldende kolonnevalg her"

    $cmdDeleteSlot = $editDesigner.Controls.Add("Forms.CommandButton.1", "cmdDeleteSlot", $true)
    $cmdDeleteSlot.Left = 10; $cmdDeleteSlot.Top = 178; $cmdDeleteSlot.Width = 320; $cmdDeleteSlot.Height = 26
    $cmdDeleteSlot.Caption = $capDeleteSlot

    $cmdCancelEdit = $editDesigner.Controls.Add("Forms.CommandButton.1", "cmdCancelEdit", $true)
    $cmdCancelEdit.Left = 10; $cmdCancelEdit.Top = 210; $cmdCancelEdit.Width = 320; $cmdCancelEdit.Height = 24
    $cmdCancelEdit.Caption = "Avbryt"

    $editComp.CodeModule.AddFromString($editFormCode)

    $mod.Export((Join-Path $tempDir "modColumnPicker.bas"))
    $cls.Export((Join-Path $tempDir "clsColumnCheckHandler.cls"))
    $frmComp.Export((Join-Path $tempDir "frmColumnPicker.frm"))
    $editComp.Export((Join-Path $tempDir "frmLoadoutEdit.frm"))

    $buildWb.Close($false)

    # ---- 5. Bytt ut / legg til de fire komponentene i maal-arbeidsboken ----

    $componentsToReplace = @(
        @{ Name = "modColumnPicker"; File = "modColumnPicker.bas" },
        @{ Name = "clsColumnCheckHandler"; File = "clsColumnCheckHandler.cls" },
        @{ Name = "frmColumnPicker"; File = "frmColumnPicker.frm" },
        @{ Name = "frmLoadoutEdit"; File = "frmLoadoutEdit.frm" }
    )

    foreach ($item in $componentsToReplace) {
        $filePath = Join-Path $tempDir $item.File

        $existing = $null
        foreach ($comp in $vbproj.VBComponents) {
            if ($comp.Name -eq $item.Name) { $existing = $comp }
        }
        if ($existing) {
            $vbproj.VBComponents.Remove($existing)
        }
        $null = $vbproj.VBComponents.Import($filePath)
        Write-Output "  $($item.Name) installert"
    }

    # ---- 6. Ved fersk installasjon: legg til en knapp paa det aktive arket ----

    if ($isFreshInstall) {
        # Reaktiver maalarbeidsboken eksplisitt, og kvalifiser OnAction med dens
        # eget navn - hvis flere arbeidsboker er aapne i samme Excel-instans kan
        # en ukvalifisert .OnAction-tildeling bli feilkvalifisert mot en LIKNENDE
        # navngitt makro i en helt ANNEN, samtidig aapen fil (sett i praksis).
        $wb.Activate()
        $ws = $wb.ActiveSheet
        $tableCount = $ws.ListObjects.Count

        if ($tableCount -ne 1) {
            Write-Output "  ADVARSEL: Aktivt ark ('$($ws.Name)') har $tableCount tabeller (forventet 1). Knappen legges til likevel - flytt den til riktig ark om noedvendig."
        }

        $btn = $ws.Buttons().Add(10, 5, 140, 32)
        $btn.OnAction = "'" + $wb.Name + "'!ShowColumnPicker"
        $btn.Caption = "Kolonnevelger"
        $btn.Placement = 3   # xlFreeFloating: ikke flytt/endre storrelse med celler
        Write-Output "  Knapp lagt til i arket '$($ws.Name)' (kopier den til andre ark ved behov)"
    }

    Write-Output ""
    Write-Output "Ferdig. Husk aa lagre filen selv (Ctrl+S) naar du er klar."
} finally {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
