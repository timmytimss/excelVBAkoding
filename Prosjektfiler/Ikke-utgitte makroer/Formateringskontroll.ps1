param(
    [string]$Path,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# Formateringskontroll — alt-i-en installer/oppdaterer.
# Bygger de to VBA-komponentene fra kildekoden innebygd nedenfor, i en
# midlertidig arbeidsbok, og installerer/oppdaterer dem deretter i mal-filen.
# Ingen andre filer trengs eller etterlates.
#
# v1-omfang (bevisst enkelt utenom selve fleksibiliteten i regel-oppsettet):
# en UserForm der man velger en "aktiv kolonne" i tabellen og enten
# (a) legger til en IF-fargeregel med Excels egne native FormatConditions —
#     med rikt utvalg av betingelser og fri valg av hvilke kolonne(r)/hele
#     raden som skal fargelegges, eller
# (b) setter/fjerner en dropdown-liste (Data Validation) på kolonnen.
# Ingen egen regel-database, Audit Log, Health Check eller Repair i v1 —
# reglene lever kun som Excels egne CF/Validation-objekter, akkurat som om
# de var satt opp manuelt.
# ============================================================================

# ---- Innebygd VBA-kildekode ----

$moduleCode = @'
Option Explicit

Public Const FORMATKONTROLL_VERSION As String = "1.1.1"

' Sett denne makroen som "Assign Macro" på en Form Control-knapp i hvert ark.
' Knappen forutsetter at arket har nøyaktig én tabell (ListObject).
Public Sub ShowFormateringskontroll()
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

    Dim frm As New frmFormateringskontroll
    Set frm.TargetTable = tbl
    frm.LoadColumns
    frm.Show
End Sub

' ---- Bygg Excel-formel fra en betingelse ----
' value2Text brukes kun av BETWEEN ("Mellom").

Public Function BuildConditionFormula(ByVal columnName As String, ByVal conditionKey As String, ByVal valueText As String, Optional ByVal value2Text As String = "") As String
    Dim colRef As String
    Dim val As String
    Dim val2 As String

    colRef = "[@" & columnName & "]"
    val = Replace(valueText, """", """""")
    val2 = Replace(value2Text, """", """""")

    Select Case conditionKey
        Case "EQUALS"
            BuildConditionFormula = "=" & colRef & "=""" & val & """"
        Case "NOT_EQUALS"
            BuildConditionFormula = "=" & colRef & "<>""" & val & """"
        Case "GREATER"
            BuildConditionFormula = "=" & colRef & ">" & val
        Case "GREATER_EQUAL"
            BuildConditionFormula = "=" & colRef & ">=" & val
        Case "LESS"
            BuildConditionFormula = "=" & colRef & "<" & val
        Case "LESS_EQUAL"
            BuildConditionFormula = "=" & colRef & "<=" & val
        Case "BETWEEN"
            BuildConditionFormula = "=AND(" & colRef & ">=" & val & "," & colRef & "<=" & val2 & ")"
        Case "CONTAINS"
            BuildConditionFormula = "=ISNUMBER(SEARCH(""" & val & """," & colRef & "))"
        Case "NOT_CONTAINS"
            BuildConditionFormula = "=NOT(ISNUMBER(SEARCH(""" & val & """," & colRef & ")))"
        Case "STARTS_WITH"
            BuildConditionFormula = "=LEFT(" & colRef & "," & Len(valueText) & ")=""" & val & """"
        Case "ENDS_WITH"
            BuildConditionFormula = "=RIGHT(" & colRef & "," & Len(valueText) & ")=""" & val & """"
        Case "EMPTY"
            BuildConditionFormula = "=" & colRef & "="""""
        Case "NOT_EMPTY"
            BuildConditionFormula = "=" & colRef & "<>"""""
        Case Else
            BuildConditionFormula = ""
    End Select
End Function

' ---- Betinget formatering (IF-fargelegging) ----
'
' VIKTIG: FormatConditions.Add feiler med "Invalid procedure call or
' argument" hvis Range-objektet har flere ikke-sammenhengende områder
' (f.eks. en Union av to kolonner som ikke ligger inntil hverandre).
' Derfor legges regelen til PER KOLONNE i en løkke når "flere kolonner"
' er valgt, i stedet for på én Union-range. "Hele raden" er trygt fordi
' tbl.DataBodyRange alltid er ett sammenhengende område.

Public Sub AddColorRule(ByVal tbl As ListObject, ByVal wholeRow As Boolean, ByVal targetColumnNames As Collection, ByVal formula As String, ByVal colorKey As String, ByVal boldOn As Boolean, ByVal italicOn As Boolean)
    Dim rng As Range
    Dim nm As Variant

    On Error GoTo ErrorHandler

    If Len(formula) = 0 Then Exit Sub
    If Left$(formula, 1) <> "=" Then Exit Sub

    If wholeRow Then
        ApplyColorToRange tbl.DataBodyRange, formula, colorKey, boldOn, italicOn
    Else
        For Each nm In targetColumnNames
            Set rng = Nothing
            On Error Resume Next
            Set rng = tbl.ListColumns(CStr(nm)).DataBodyRange
            On Error GoTo ErrorHandler
            If Not rng Is Nothing Then
                ApplyColorToRange rng, formula, colorKey, boldOn, italicOn
            End If
        Next nm
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Kunne ikke legge til fargeregel:" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub ApplyColorToRange(ByVal rng As Range, ByVal formula As String, ByVal colorKey As String, ByVal boldOn As Boolean, ByVal italicOn As Boolean)
    Dim fc As FormatCondition

    If rng Is Nothing Then Exit Sub

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)

    Select Case colorKey
        Case "RED"
            fc.Interior.Color = RGB(255, 199, 206)
            fc.Font.Color = RGB(156, 0, 6)
        Case "GREEN"
            fc.Interior.Color = RGB(198, 239, 206)
            fc.Font.Color = RGB(0, 97, 0)
        Case "YELLOW"
            fc.Interior.Color = RGB(255, 235, 156)
            fc.Font.Color = RGB(156, 101, 0)
        Case "BLUE"
            fc.Interior.Color = RGB(189, 215, 238)
            fc.Font.Color = RGB(31, 73, 125)
        Case "ORANGE"
            fc.Interior.Color = RGB(255, 224, 178)
            fc.Font.Color = RGB(204, 102, 0)
        Case "PURPLE"
            fc.Interior.Color = RGB(230, 208, 240)
            fc.Font.Color = RGB(112, 48, 160)
        Case "TEAL"
            fc.Interior.Color = RGB(198, 235, 232)
            fc.Font.Color = RGB(0, 105, 92)
        Case "GRAY"
            fc.Interior.Color = RGB(217, 217, 217)
            fc.Font.Color = RGB(64, 64, 64)
    End Select

    If boldOn Then fc.Font.Bold = True
    If italicOn Then fc.Font.Italic = True
End Sub

Public Sub ClearTargetColorRules(ByVal tbl As ListObject, ByVal wholeRow As Boolean, ByVal targetColumnNames As Collection)
    Dim rng As Range
    Dim nm As Variant

    On Error Resume Next

    If wholeRow Then
        tbl.DataBodyRange.FormatConditions.Delete
    Else
        For Each nm In targetColumnNames
            Set rng = tbl.ListColumns(CStr(nm)).DataBodyRange
            If Not rng Is Nothing Then rng.FormatConditions.Delete
        Next nm
    End If

    On Error GoTo 0
End Sub

Public Sub ClearAllTableColorRules(ByVal tbl As ListObject)
    On Error Resume Next
    tbl.DataBodyRange.FormatConditions.Delete
    On Error GoTo 0
End Sub

' ---- Data-dropdown (Data Validation) ----

Public Sub SetDropdown(ByVal tbl As ListObject, ByVal columnName As String, ByVal valuesFormula As String, ByVal allowBlank As Boolean)
    Dim rng As Range

    On Error GoTo ErrorHandler

    Set rng = tbl.ListColumns(columnName).DataBodyRange
    If rng Is Nothing Then Exit Sub

    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=valuesFormula
        .IgnoreBlank = allowBlank
        .InCellDropdown = True
    End With

    Exit Sub

ErrorHandler:
    MsgBox "Kunne ikke sette dropdown:" & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub RemoveDropdown(ByVal tbl As ListObject, ByVal columnName As String)
    On Error Resume Next
    tbl.ListColumns(columnName).DataBodyRange.Validation.Delete
    On Error GoTo 0
End Sub
'@

$formCode = @'
Option Explicit

Public TargetTable As ListObject

Private Sub UserForm_Initialize()
    Me.Caption = "Formateringskontroll"

    cmbCondition.Clear
    cmbCondition.AddItem "Lik"
    cmbCondition.AddItem "Ikke lik"
    cmbCondition.AddItem "Større enn"
    cmbCondition.AddItem "Større enn eller lik"
    cmbCondition.AddItem "Mindre enn"
    cmbCondition.AddItem "Mindre enn eller lik"
    cmbCondition.AddItem "Mellom"
    cmbCondition.AddItem "Inneholder"
    cmbCondition.AddItem "Inneholder ikke"
    cmbCondition.AddItem "Starter med"
    cmbCondition.AddItem "Slutter med"
    cmbCondition.AddItem "Er tom"
    cmbCondition.AddItem "Er ikke tom"
    cmbCondition.AddItem "Egendefinert formel"
    cmbCondition.ListIndex = 0

    chkAllowBlank.Value = True

    UpdateValueControlsForCondition
End Sub

Public Sub LoadColumns()
    Dim col As ListColumn
    Dim chk As MSForms.CheckBox
    Dim idx As Long, colsPerRow As Long
    Dim colWidth As Single, rowHeight As Single
    Dim r As Long, c As Long, totalRows As Long, neededHeight As Single
    Dim displayName As String
    Const MAX_GRID_HEIGHT As Single = 100

    If TargetTable Is Nothing Then Exit Sub

    lblTitle.Caption = "Ark: " & TargetTable.Parent.Name & "   Tabell: " & TargetTable.Name

    cmbColumn.Clear
    For Each col In TargetTable.ListColumns
        cmbColumn.AddItem col.Name
    Next col
    If cmbColumn.ListCount > 0 Then cmbColumn.ListIndex = 0

    colsPerRow = 3
    rowHeight = 18
    colWidth = (frameTargetColumns.InsideWidth - 12) / colsPerRow

    idx = 0
    For Each col In TargetTable.ListColumns
        idx = idx + 1
        r = (idx - 1) \ colsPerRow
        c = (idx - 1) Mod colsPerRow

        displayName = Replace(Replace(col.Name, vbLf, " "), vbCr, "")

        Set chk = frameTargetColumns.Controls.Add("Forms.CheckBox.1", "chkTgt" & idx, True)
        With chk
            .Caption = displayName
            .WordWrap = False
            .Left = 6 + c * colWidth
            .Top = 6 + r * rowHeight
            .Width = colWidth - 6
            .Height = rowHeight - 2
            .Tag = col.Name
            .Value = False
        End With
    Next col

    totalRows = ((idx - 1) \ colsPerRow) + 1
    neededHeight = 6 + totalRows * rowHeight + 10

    If neededHeight > MAX_GRID_HEIGHT Then
        frameTargetColumns.ScrollBars = fmScrollBarsVertical
        frameTargetColumns.ScrollHeight = neededHeight
    Else
        frameTargetColumns.ScrollBars = fmScrollBarsNone
    End If

    SyncTargetColumnsToActiveColumn
End Sub

Private Function SelectedColumnName() As String
    If cmbColumn.ListIndex < 0 Then
        SelectedColumnName = ""
    Else
        SelectedColumnName = cmbColumn.List(cmbColumn.ListIndex)
    End If
End Function

Private Function SelectedConditionKey() As String
    Select Case cmbCondition.ListIndex
        Case 0: SelectedConditionKey = "EQUALS"
        Case 1: SelectedConditionKey = "NOT_EQUALS"
        Case 2: SelectedConditionKey = "GREATER"
        Case 3: SelectedConditionKey = "GREATER_EQUAL"
        Case 4: SelectedConditionKey = "LESS"
        Case 5: SelectedConditionKey = "LESS_EQUAL"
        Case 6: SelectedConditionKey = "BETWEEN"
        Case 7: SelectedConditionKey = "CONTAINS"
        Case 8: SelectedConditionKey = "NOT_CONTAINS"
        Case 9: SelectedConditionKey = "STARTS_WITH"
        Case 10: SelectedConditionKey = "ENDS_WITH"
        Case 11: SelectedConditionKey = "EMPTY"
        Case 12: SelectedConditionKey = "NOT_EMPTY"
        Case 13: SelectedConditionKey = "FORMULA"
        Case Else: SelectedConditionKey = ""
    End Select
End Function

Private Sub UpdateValueControlsForCondition()
    Select Case SelectedConditionKey()
        Case "EMPTY", "NOT_EMPTY"
            txtValue.Enabled = False
            txtValue.Text = ""
            lblValue.Caption = "Verdi:"
            lblValue2.Visible = False
            txtValue2.Visible = False
            txtValue2.Text = ""
        Case "BETWEEN"
            txtValue.Enabled = True
            lblValue.Caption = "Fra:"
            lblValue2.Caption = "Til:"
            lblValue2.Visible = True
            txtValue2.Visible = True
        Case "FORMULA"
            txtValue.Enabled = True
            lblValue.Caption = "Formel (må starte med =):"
            lblValue2.Visible = False
            txtValue2.Visible = False
            txtValue2.Text = ""
        Case Else
            txtValue.Enabled = True
            lblValue.Caption = "Verdi:"
            lblValue2.Visible = False
            txtValue2.Visible = False
            txtValue2.Text = ""
    End Select
End Sub

Private Sub cmbCondition_Change()
    UpdateValueControlsForCondition
End Sub

Private Sub SyncTargetColumnsToActiveColumn()
    Dim ctl As MSForms.Control
    Dim activeName As String
    activeName = SelectedColumnName()
    For Each ctl In frameTargetColumns.Controls
        If TypeName(ctl) = "CheckBox" Then
            ctl.Value = (ctl.Tag = activeName)
        End If
    Next ctl
End Sub

Private Sub cmbColumn_Change()
    If Not CBool(chkWholeRow.Value) Then SyncTargetColumnsToActiveColumn
End Sub

Private Sub chkWholeRow_Click()
    Dim isEnabled As Boolean
    isEnabled = Not CBool(chkWholeRow.Value)
    frameTargetColumns.Enabled = isEnabled
    btnSelectAllCols.Enabled = isEnabled
    btnSelectNoCols.Enabled = isEnabled
    If isEnabled Then SyncTargetColumnsToActiveColumn
End Sub

Private Sub btnSelectAllCols_Click()
    Dim ctl As MSForms.Control
    For Each ctl In frameTargetColumns.Controls
        If TypeName(ctl) = "CheckBox" Then ctl.Value = True
    Next ctl
End Sub

Private Sub btnSelectNoCols_Click()
    Dim ctl As MSForms.Control
    For Each ctl In frameTargetColumns.Controls
        If TypeName(ctl) = "CheckBox" Then ctl.Value = False
    Next ctl
End Sub

Private Function GetCheckedTargetColumns() As Collection
    Dim result As New Collection
    Dim ctl As MSForms.Control
    For Each ctl In frameTargetColumns.Controls
        If TypeName(ctl) = "CheckBox" Then
            If ctl.Value Then result.Add ctl.Tag
        End If
    Next ctl
    Set GetCheckedTargetColumns = result
End Function

Private Function BuildFormulaFromForm(ByVal colName As String, ByVal condKey As String) As String
    Dim v1 As String, v2 As String

    Select Case condKey
        Case "FORMULA"
            v1 = Trim(txtValue.Text)
            If Left$(v1, 1) <> "=" Then
                MsgBox "Egendefinert formel må starte med ""="".", vbExclamation
                Exit Function
            End If
            BuildFormulaFromForm = v1

        Case "EMPTY", "NOT_EMPTY"
            BuildFormulaFromForm = modFormateringskontroll.BuildConditionFormula(colName, condKey, "")

        Case "BETWEEN"
            v1 = Trim(txtValue.Text)
            v2 = Trim(txtValue2.Text)
            If v1 = "" Or v2 = "" Then
                MsgBox "Skriv inn begge verdiene for ""Mellom"".", vbExclamation
                Exit Function
            End If
            BuildFormulaFromForm = modFormateringskontroll.BuildConditionFormula(colName, condKey, v1, v2)

        Case Else
            v1 = Trim(txtValue.Text)
            If v1 = "" Then
                MsgBox "Skriv inn en verdi.", vbExclamation
                Exit Function
            End If
            BuildFormulaFromForm = modFormateringskontroll.BuildConditionFormula(colName, condKey, v1)
    End Select
End Function

Private Sub HandleAddColor(ByVal colorKey As String)
    Dim colName As String
    Dim condKey As String
    Dim formula As String
    Dim targetNames As Collection
    Dim wholeRow As Boolean

    colName = SelectedColumnName()
    If colName = "" Then
        MsgBox "Velg en kolonne først.", vbExclamation
        Exit Sub
    End If

    condKey = SelectedConditionKey()
    formula = BuildFormulaFromForm(colName, condKey)
    If formula = "" Then Exit Sub

    wholeRow = CBool(chkWholeRow.Value)
    Set targetNames = GetCheckedTargetColumns()

    If Not wholeRow And targetNames.Count = 0 Then
        MsgBox "Velg minst én kolonne å fargelegge, eller kryss av for hele raden.", vbExclamation
        Exit Sub
    End If

    modFormateringskontroll.AddColorRule TargetTable, wholeRow, targetNames, formula, colorKey, CBool(chkBold.Value), CBool(chkItalic.Value)

    If wholeRow Then
        lblStatus.Caption = "Fargeregel lagt til (hele raden)."
    Else
        lblStatus.Caption = "Fargeregel lagt til (" & targetNames.Count & " kolonne(r))."
    End If
End Sub

Private Sub btnColorRed_Click(): HandleAddColor "RED": End Sub
Private Sub btnColorGreen_Click(): HandleAddColor "GREEN": End Sub
Private Sub btnColorYellow_Click(): HandleAddColor "YELLOW": End Sub
Private Sub btnColorBlue_Click(): HandleAddColor "BLUE": End Sub
Private Sub btnColorOrange_Click(): HandleAddColor "ORANGE": End Sub
Private Sub btnColorPurple_Click(): HandleAddColor "PURPLE": End Sub
Private Sub btnColorTeal_Click(): HandleAddColor "TEAL": End Sub
Private Sub btnColorGray_Click(): HandleAddColor "GRAY": End Sub

Private Sub cmdClearTargetRules_Click()
    Dim targetNames As Collection
    Dim wholeRow As Boolean

    wholeRow = CBool(chkWholeRow.Value)
    Set targetNames = GetCheckedTargetColumns()

    If Not wholeRow And targetNames.Count = 0 Then
        MsgBox "Velg minst én kolonne, eller kryss av for hele raden.", vbExclamation
        Exit Sub
    End If

    modFormateringskontroll.ClearTargetColorRules TargetTable, wholeRow, targetNames
    lblStatus.Caption = "Fargeregler fjernet for valgt mål."
End Sub

Private Sub cmdResetTableRules_Click()
    If MsgBox("Fjerne ALLE fargeregler i hele tabellen """ & TargetTable.Name & """?", vbQuestion + vbYesNo) = vbYes Then
        modFormateringskontroll.ClearAllTableColorRules TargetTable
        lblStatus.Caption = "Alle fargeregler i tabellen er fjernet."
    End If
End Sub

Private Sub cmdSetDropdown_Click()
    Dim colName As String
    Dim valuesText As String

    colName = SelectedColumnName()
    If colName = "" Then
        MsgBox "Velg en kolonne først.", vbExclamation
        Exit Sub
    End If

    valuesText = Trim(txtDropdownValues.Text)
    If valuesText = "" Then
        MsgBox "Skriv inn minst én verdi.", vbExclamation
        Exit Sub
    End If

    modFormateringskontroll.SetDropdown TargetTable, colName, valuesText, CBool(chkAllowBlank.Value)
    lblStatus.Caption = "Dropdown satt på """ & colName & """."
End Sub

Private Sub cmdRemoveDropdown_Click()
    Dim colName As String
    colName = SelectedColumnName()
    If colName = "" Then Exit Sub

    modFormateringskontroll.RemoveDropdown TargetTable, colName
    lblStatus.Caption = "Dropdown fjernet fra """ & colName & """."
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
'@

function Get-FormatkontrollVersion {
    param([string]$Code)
    if ($Code -match 'FORMATKONTROLL_VERSION\s*As\s*String\s*=\s*"([^"]+)"') {
        return $Matches[1]
    }
    return $null
}

$sourceVersion = Get-FormatkontrollVersion $moduleCode

# ---- 1. Finn maalfilen ----

if (-not $Path) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel-filer (*.xlsm)|*.xlsm|Alle filer (*.*)|*.*"
    $dlg.Title = "Velg Excel-fil som skal faa Formateringskontroll"
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
    # filter (et ark kan ha flere uavhengig installerte makroers knapper
    # side om side).
    Write-Output "Fjerner Formateringskontroll fra $($wb.Name) ..."
    foreach ($navn in @("modFormateringskontroll", "frmFormateringskontroll")) {
        try {
            $eksisterende = $vbproj.VBComponents.Item($navn)
            $vbproj.VBComponents.Remove($eksisterende)
            Write-Output "  Fjernet $navn"
        } catch {}
    }
    try {
        foreach ($ws in $wb.Worksheets) {
            foreach ($btn in @($ws.Buttons())) {
                if ($btn.OnAction -eq "ShowFormateringskontroll") { $btn.Delete() }
            }
        }
    } catch {}
    try { $wb.Save() } catch {
        Write-Output "FEIL ved lagring: $($_.Exception.Message)"
        exit 1
    }
    Write-Output "Ferdig - Formateringskontroll er fjernet fra $($wb.Name)."
    exit 0
}

# ---- 3. Sjekk hva som allerede er installert ----

$existingModule = $null
foreach ($comp in $vbproj.VBComponents) {
    if ($comp.Name -eq "modFormateringskontroll") { $existingModule = $comp }
}

$installedVersion = $null
if ($existingModule) {
    $existingCode = $existingModule.CodeModule.Lines(1, $existingModule.CodeModule.CountOfLines)
    $installedVersion = Get-FormatkontrollVersion $existingCode
}

$isFreshInstall = ($null -eq $existingModule)

if (-not $isFreshInstall -and $installedVersion -eq $sourceVersion) {
    Write-Output "Formateringskontroll er allerede installert og oppdatert (versjon $installedVersion) i $($wb.Name)."
    exit 0
}

if ($isFreshInstall) {
    Write-Output "Installerer Formateringskontroll (versjon $sourceVersion) i $($wb.Name) ..."
} else {
    Write-Output "Oppdaterer Formateringskontroll: versjon $installedVersion -> $sourceVersion i $($wb.Name) ..."
}

# ---- 4. Bygg de to komponentene i en midlertidig arbeidsbok, eksporter til en temp-mappe ----

$tempDir = Join-Path $env:TEMP ("Formateringskontroll_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir | Out-Null

try {
    $buildWb = $excel.Workbooks.Add()
    $buildProj = $buildWb.VBProject

    $mod = $buildProj.VBComponents.Add(1)
    $mod.Name = "modFormateringskontroll"
    $mod.CodeModule.AddFromString($moduleCode)

    $frmComp = $buildProj.VBComponents.Add(3)
    $frmComp.Name = "frmFormateringskontroll"
    $frmComp.Properties("Width").Value = 500
    $frmComp.Properties("Height").Value = 560
    $frmComp.Properties("Caption").Value = "Formateringskontroll"
    $designer = $frmComp.Designer

    # ---- Fast topptekst ----

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Left = 10; $lblTitle.Top = 10; $lblTitle.Width = 470; $lblTitle.Height = 18

    # ---- Scrollbart innholdsområde ----

    $frameScroll = $designer.Controls.Add("Forms.Frame.1", "frameScroll", $true)
    $frameScroll.Left = 10; $frameScroll.Top = 32; $frameScroll.Width = 470; $frameScroll.Height = 460
    $frameScroll.Caption = ""
    $frameScroll.ScrollBars = 2   # fmScrollBarsVertical

    # -- Betingelse --

    $frameCondition = $frameScroll.Controls.Add("Forms.Frame.1", "frameCondition", $true)
    $frameCondition.Left = 0; $frameCondition.Top = 0; $frameCondition.Width = 450; $frameCondition.Height = 100
    $frameCondition.Caption = "Betingelse"

    $lblColumnCap = $frameCondition.Controls.Add("Forms.Label.1", "lblColumnCap", $true)
    $lblColumnCap.Left = 10; $lblColumnCap.Top = 10; $lblColumnCap.Width = 90; $lblColumnCap.Height = 18
    $lblColumnCap.Caption = "Aktiv kolonne:"

    $cmbColumn = $frameCondition.Controls.Add("Forms.ComboBox.1", "cmbColumn", $true)
    $cmbColumn.Left = 104; $cmbColumn.Top = 8; $cmbColumn.Width = 336; $cmbColumn.Height = 20
    $cmbColumn.Style = 2

    $lblCondition = $frameCondition.Controls.Add("Forms.Label.1", "lblCondition", $true)
    $lblCondition.Left = 10; $lblCondition.Top = 36; $lblCondition.Width = 90; $lblCondition.Height = 18
    $lblCondition.Caption = "Betingelse:"

    $cmbCondition = $frameCondition.Controls.Add("Forms.ComboBox.1", "cmbCondition", $true)
    $cmbCondition.Left = 104; $cmbCondition.Top = 34; $cmbCondition.Width = 336; $cmbCondition.Height = 20
    $cmbCondition.Style = 2

    $lblValue = $frameCondition.Controls.Add("Forms.Label.1", "lblValue", $true)
    $lblValue.Left = 10; $lblValue.Top = 62; $lblValue.Width = 90; $lblValue.Height = 18
    $lblValue.Caption = "Verdi:"

    $txtValue = $frameCondition.Controls.Add("Forms.TextBox.1", "txtValue", $true)
    $txtValue.Left = 104; $txtValue.Top = 60; $txtValue.Width = 150; $txtValue.Height = 20

    $lblValue2 = $frameCondition.Controls.Add("Forms.Label.1", "lblValue2", $true)
    $lblValue2.Left = 264; $lblValue2.Top = 62; $lblValue2.Width = 30; $lblValue2.Height = 18
    $lblValue2.Caption = "Til:"

    $txtValue2 = $frameCondition.Controls.Add("Forms.TextBox.1", "txtValue2", $true)
    $txtValue2.Left = 294; $txtValue2.Top = 60; $txtValue2.Width = 146; $txtValue2.Height = 20

    # -- Hvor skal fargen brukes --

    $frameTarget = $frameScroll.Controls.Add("Forms.Frame.1", "frameTarget", $true)
    $frameTarget.Left = 0; $frameTarget.Top = 110; $frameTarget.Width = 450; $frameTarget.Height = 176
    $frameTarget.Caption = "Hvor skal fargen brukes?"

    $chkWholeRow = $frameTarget.Controls.Add("Forms.CheckBox.1", "chkWholeRow", $true)
    $chkWholeRow.Left = 10; $chkWholeRow.Top = 10; $chkWholeRow.Width = 250; $chkWholeRow.Height = 18
    $chkWholeRow.Caption = "Hele raden (alle kolonner)"

    $btnSelectAllCols = $frameTarget.Controls.Add("Forms.CommandButton.1", "btnSelectAllCols", $true)
    $btnSelectAllCols.Left = 270; $btnSelectAllCols.Top = 8; $btnSelectAllCols.Width = 80; $btnSelectAllCols.Height = 22
    $btnSelectAllCols.Caption = "Velg alle"

    $btnSelectNoCols = $frameTarget.Controls.Add("Forms.CommandButton.1", "btnSelectNoCols", $true)
    $btnSelectNoCols.Left = 356; $btnSelectNoCols.Top = 8; $btnSelectNoCols.Width = 84; $btnSelectNoCols.Height = 22
    $btnSelectNoCols.Caption = "Velg ingen"

    $lblTargetHint = $frameTarget.Controls.Add("Forms.Label.1", "lblTargetHint", $true)
    $lblTargetHint.Left = 10; $lblTargetHint.Top = 36; $lblTargetHint.Width = 350; $lblTargetHint.Height = 18
    $lblTargetHint.Caption = "Eller kryss av enkeltkolonner (flere valgt = flere kolonner farges):"

    $frameTargetColumns = $frameTarget.Controls.Add("Forms.Frame.1", "frameTargetColumns", $true)
    $frameTargetColumns.Left = 10; $frameTargetColumns.Top = 56; $frameTargetColumns.Width = 430; $frameTargetColumns.Height = 100
    $frameTargetColumns.Caption = ""

    # -- Farge og stil --

    $frameStyle = $frameScroll.Controls.Add("Forms.Frame.1", "frameStyle", $true)
    $frameStyle.Left = 0; $frameStyle.Top = 296; $frameStyle.Width = 450; $frameStyle.Height = 130
    $frameStyle.Caption = "Farge og stil"

    $lblFarge = $frameStyle.Controls.Add("Forms.Label.1", "lblFarge", $true)
    $lblFarge.Left = 10; $lblFarge.Top = 10; $lblFarge.Width = 400; $lblFarge.Height = 18
    $lblFarge.Caption = "Farge (klikk for å legge til regelen med gjeldende betingelse og mål):"

    $colorButtons = @(
        @{ Name = "btnColorRed";    Caption = "Rød";   Left = 10;  Top = 34; RGB = @(255,199,206) },
        @{ Name = "btnColorGreen";  Caption = "Grønn"; Left = 120; Top = 34; RGB = @(198,239,206) },
        @{ Name = "btnColorYellow"; Caption = "Gul";    Left = 230; Top = 34; RGB = @(255,235,156) },
        @{ Name = "btnColorBlue";   Caption = "Blå";    Left = 340; Top = 34; RGB = @(189,215,238) },
        @{ Name = "btnColorOrange"; Caption = "Oransje"; Left = 10;  Top = 66; RGB = @(255,224,178) },
        @{ Name = "btnColorPurple"; Caption = "Lilla";  Left = 120; Top = 66; RGB = @(230,208,240) },
        @{ Name = "btnColorTeal";   Caption = "Turkis"; Left = 230; Top = 66; RGB = @(198,235,232) },
        @{ Name = "btnColorGray";   Caption = "Grå";    Left = 340; Top = 66; RGB = @(217,217,217) }
    )
    foreach ($cb in $colorButtons) {
        $btn = $frameStyle.Controls.Add("Forms.CommandButton.1", $cb.Name, $true)
        $btn.Left = $cb.Left; $btn.Top = $cb.Top; $btn.Width = 100; $btn.Height = 26
        $btn.Caption = $cb.Caption
        $rgb = $cb.RGB
        $btn.BackColor = [System.Convert]::ToInt64($rgb[2]) * 65536 + [System.Convert]::ToInt64($rgb[1]) * 256 + [System.Convert]::ToInt64($rgb[0])
    }

    $chkBold = $frameStyle.Controls.Add("Forms.CheckBox.1", "chkBold", $true)
    $chkBold.Left = 10; $chkBold.Top = 98; $chkBold.Width = 110; $chkBold.Height = 18
    $chkBold.Caption = "Fet skrift"

    $chkItalic = $frameStyle.Controls.Add("Forms.CheckBox.1", "chkItalic", $true)
    $chkItalic.Left = 130; $chkItalic.Top = 98; $chkItalic.Width = 110; $chkItalic.Height = 18
    $chkItalic.Caption = "Kursiv"

    # -- Verktøy-knapper --

    $cmdClearTargetRules = $frameScroll.Controls.Add("Forms.CommandButton.1", "cmdClearTargetRules", $true)
    $cmdClearTargetRules.Left = 0; $cmdClearTargetRules.Top = 436; $cmdClearTargetRules.Width = 220; $cmdClearTargetRules.Height = 26
    $cmdClearTargetRules.Caption = "Fjern regler for valgt mål"

    $cmdResetTableRules = $frameScroll.Controls.Add("Forms.CommandButton.1", "cmdResetTableRules", $true)
    $cmdResetTableRules.Left = 230; $cmdResetTableRules.Top = 436; $cmdResetTableRules.Width = 220; $cmdResetTableRules.Height = 26
    $cmdResetTableRules.Caption = "Nullstill ALT i tabellen"

    # -- Dropdown --

    $frameDropdown = $frameScroll.Controls.Add("Forms.Frame.1", "frameDropdown", $true)
    $frameDropdown.Left = 0; $frameDropdown.Top = 472; $frameDropdown.Width = 450; $frameDropdown.Height = 124
    $frameDropdown.Caption = "Dropdown (nedtrekksliste)"

    $lblValues = $frameDropdown.Controls.Add("Forms.Label.1", "lblValues", $true)
    $lblValues.Left = 10; $lblValues.Top = 10; $lblValues.Width = 300; $lblValues.Height = 18
    $lblValues.Caption = "Verdier (kommaseparert):"

    $txtDropdownValues = $frameDropdown.Controls.Add("Forms.TextBox.1", "txtDropdownValues", $true)
    $txtDropdownValues.Left = 10; $txtDropdownValues.Top = 30; $txtDropdownValues.Width = 430; $txtDropdownValues.Height = 20

    $chkAllowBlank = $frameDropdown.Controls.Add("Forms.CheckBox.1", "chkAllowBlank", $true)
    $chkAllowBlank.Left = 10; $chkAllowBlank.Top = 56; $chkAllowBlank.Width = 150; $chkAllowBlank.Height = 18
    $chkAllowBlank.Caption = "Tillat tom celle"

    $cmdSetDropdown = $frameDropdown.Controls.Add("Forms.CommandButton.1", "cmdSetDropdown", $true)
    $cmdSetDropdown.Left = 10; $cmdSetDropdown.Top = 82; $cmdSetDropdown.Width = 205; $cmdSetDropdown.Height = 26
    $cmdSetDropdown.Caption = "Sett dropdown"

    $cmdRemoveDropdown = $frameDropdown.Controls.Add("Forms.CommandButton.1", "cmdRemoveDropdown", $true)
    $cmdRemoveDropdown.Left = 225; $cmdRemoveDropdown.Top = 82; $cmdRemoveDropdown.Width = 215; $cmdRemoveDropdown.Height = 26
    $cmdRemoveDropdown.Caption = "Fjern dropdown"

    $frameScroll.ScrollHeight = 600
    $frameScroll.ScrollWidth = 452

    # ---- Fast bunntekst ----

    $lblStatus = $designer.Controls.Add("Forms.Label.1", "lblStatus", $true)
    $lblStatus.Left = 10; $lblStatus.Top = 500; $lblStatus.Width = 400; $lblStatus.Height = 30

    $cmdClose = $designer.Controls.Add("Forms.CommandButton.1", "cmdClose", $true)
    $cmdClose.Left = 420; $cmdClose.Top = 500; $cmdClose.Width = 60; $cmdClose.Height = 24
    $cmdClose.Caption = "Lukk"

    $frmComp.CodeModule.AddFromString($formCode)

    $mod.Export((Join-Path $tempDir "modFormateringskontroll.bas"))
    $frmComp.Export((Join-Path $tempDir "frmFormateringskontroll.frm"))

    $buildWb.Close($false)

    # ---- 5. Bytt ut / legg til de to komponentene i maal-arbeidsboken ----

    $componentsToReplace = @(
        @{ Name = "modFormateringskontroll"; File = "modFormateringskontroll.bas" },
        @{ Name = "frmFormateringskontroll"; File = "frmFormateringskontroll.frm" }
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
        $ws = $wb.ActiveSheet
        $tableCount = $ws.ListObjects.Count

        if ($tableCount -ne 1) {
            Write-Output "  ADVARSEL: Aktivt ark ('$($ws.Name)') har $tableCount tabeller (forventet 1). Knappen legges til likevel - flytt den til riktig ark om noedvendig."
        }

        $btn = $ws.Buttons().Add(10, 5, 160, 32)
        $btn.OnAction = "ShowFormateringskontroll"
        $btn.Caption = "Formateringskontroll"
        $btn.Placement = 3   # xlFreeFloating: ikke flytt/endre storrelse med celler
        Write-Output "  Knapp lagt til i arket '$($ws.Name)' (kopier den til andre ark ved behov)"
    }

    Write-Output ""
    Write-Output "Ferdig. Husk aa lagre filen selv (Ctrl+S) naar du er klar."
} finally {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
