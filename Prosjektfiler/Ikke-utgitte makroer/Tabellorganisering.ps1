param(
    [string]$Path,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# Tabellorganisering - alt-i-en installer/oppdaterer.
# Bygger VBA-modulen "modTabellorganisering" fra kildekoden innebygd nedenfor,
# i en midlertidig arbeidsbok, og installerer/oppdaterer den deretter i
# maalfilen. Ved forstegangsinstallasjon legges det ogsaa til en knapp paa
# aktivt ark, samt et konfigurasjonsark "_Oppsett" forhaandsutfylt fra
# tabellen som allerede finnes der. Ingen andre filer trengs eller etterlates.
# ============================================================================

# ---- Innebygd VBA-kildekode ----

$moduleCode = @'
Option Explicit

Public Const TABELLORGANISERING_VERSION As String = "1.0.0"

' Sett denne makroen som "Assign Macro" paa en Form Control-knapp i hvert
' datark. Knappen forutsetter at arket har noyaktig en Excel-tabell
' (ListObject). Layout og innhold styres av tblOppsett i arket "_Oppsett" -
' VBA-koden er generell og kjenner ikke til konkrete kolonnenavn.
Public Sub OrganiserTabell()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim oppsett As ListObject
    Dim data As Variant
    Dim kategoriFarger As Object

    On Error GoTo Feil

    If TypeName(Application.Caller) = "String" Then
        Set ws = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Worksheet
    Else
        Set ws = ActiveSheet
    End If

    If ws.ListObjects.Count = 0 Then
        MsgBox "Fant ingen Excel-tabell i arket """ & ws.Name & """.", vbExclamation, "Organiser tabell"
        Exit Sub
    ElseIf ws.ListObjects.Count > 1 Then
        MsgBox "Fant flere tabeller i arket """ & ws.Name & """. Denne knappen forutsetter kun en tabell per ark.", vbExclamation, "Organiser tabell"
        Exit Sub
    End If

    Set tbl = ws.ListObjects(1)

    Set oppsett = FinnOppsettTabell()
    If oppsett Is Nothing Then
        MsgBox "Fant ikke tabellen 'tblOppsett' i arket '_Oppsett'." & vbCrLf & _
               "Kjor installasjonen paa nytt, eller opprett arket manuelt.", vbExclamation, "Organiser tabell"
        Exit Sub
    End If

    ValiderOppsett tbl, oppsett

    data = HentSortertOppsett(oppsett)
    Set kategoriFarger = BuildKategoriFarger(data)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    OrganiserKolonner tbl, data
    EndreKolonnenavn tbl, data
    FormaterSpesialkolonner tbl, data
    LagKategoriRad tbl, data, kategoriFarger
    FormaterTabell tbl
    FormaterKategoriRad tbl

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Tabellen er organisert og formatert.", vbInformation, "Organiser tabell"

    Exit Sub

Feil:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Det oppstod en feil:" & vbCrLf & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, "Organiser tabell"

End Sub

' ---- Finn oppsett-/kategoritabell ----

Private Function FinnOppsettTabell() As ListObject
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_Oppsett")
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set FinnOppsettTabell = ws.ListObjects("tblOppsett")
    On Error GoTo 0
End Function

Private Function FinnKategoriTabell() As ListObject
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_Oppsett")
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set FinnKategoriTabell = ws.ListObjects("tblKategorier")
    On Error GoTo 0
End Function

Private Function FinnListColumn(tbl As ListObject, kolonnenavn As String) As ListColumn
    Dim lc As ListColumn

    For Each lc In tbl.ListColumns
        If StrComp(Trim(CStr(lc.Name)), Trim(kolonnenavn), vbTextCompare) = 0 Then
            Set FinnListColumn = lc
            Exit Function
        End If
    Next lc

    Set FinnListColumn = Nothing
End Function

Private Function OppsettKolonneIndeks(tbl As ListObject, navn As String) As Long
    Dim lc As ListColumn
    Set lc = FinnListColumn(tbl, navn)
    If lc Is Nothing Then
        OppsettKolonneIndeks = 0
    Else
        OppsettKolonneIndeks = lc.Index
    End If
End Function

' ---- Validering ----

Private Sub ValiderOppsett(tbl As ListObject, oppsett As ListObject)

    Dim r As ListRow

    Dim iOriginal As Long, iNytt As Long, iKategori As Long, iRekkefolge As Long

    Dim originalNavn As String
    Dim nyttNavn As String
    Dim kategori As String
    Dim rekkefolge As Variant

    Dim duplikat As String

    If oppsett.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1001, , "_Oppsett inneholder ingen konfigurasjonsrader."
    End If

    iOriginal = OppsettKolonneIndeks(oppsett, "Original kolonne")
    iNytt = OppsettKolonneIndeks(oppsett, "Nytt navn")
    iKategori = OppsettKolonneIndeks(oppsett, "Kategori")
    iRekkefolge = OppsettKolonneIndeks(oppsett, "Rekkefølge")

    If iOriginal = 0 Or iNytt = 0 Or iKategori = 0 Or iRekkefolge = 0 Then
        Err.Raise vbObjectError + 1002, , _
            "tblOppsett mangler en eller flere paakrevde kolonner: " & _
            "'Original kolonne', 'Nytt navn', 'Kategori', 'Rekkefølge'."
    End If

    duplikat = FinnDuplikatNyttNavn(oppsett, iNytt)
    If duplikat <> "" Then
        Err.Raise vbObjectError + 1003, , "Nytt kolonnenavn finnes flere ganger: '" & duplikat & "'."
    End If

    duplikat = FinnDuplikatRekkefolge(oppsett, iRekkefolge)
    If duplikat <> "" Then
        Err.Raise vbObjectError + 1004, , "Rekkefølge finnes flere ganger: '" & duplikat & "'."
    End If

    For Each r In oppsett.ListRows

        originalNavn = Trim(CStr(r.Range.Cells(1, iOriginal).Value))
        nyttNavn = Trim(CStr(r.Range.Cells(1, iNytt).Value))
        kategori = Trim(CStr(r.Range.Cells(1, iKategori).Value))
        rekkefolge = r.Range.Cells(1, iRekkefolge).Value

        If originalNavn = "" Then
            Err.Raise vbObjectError + 1005, , "En rad i _Oppsett mangler 'Original kolonne'."
        End If

        If nyttNavn = "" Then
            Err.Raise vbObjectError + 1006, , "Kolonnen '" & originalNavn & "' mangler 'Nytt navn'."
        End If

        If kategori = "" Then
            Err.Raise vbObjectError + 1007, , "Kolonnen '" & originalNavn & "' mangler kategori."
        End If

        If Not IsNumeric(rekkefolge) Then
            Err.Raise vbObjectError + 1008, , "Kolonnen '" & originalNavn & "' har ugyldig rekkefolge."
        End If

        If FinnListColumn(tbl, originalNavn) Is Nothing Then
            Err.Raise vbObjectError + 1009, , "Fant ikke kolonnen '" & originalNavn & "' i tabellen '" & tbl.Name & "'."
        End If

    Next r

End Sub

Private Function FinnDuplikatNyttNavn(oppsett As ListObject, iNytt As Long) As String
    Dim i As Long, j As Long
    Dim navn1 As String, navn2 As String

    For i = 1 To oppsett.ListRows.Count
        navn1 = Trim(CStr(oppsett.DataBodyRange.Cells(i, iNytt).Value))
        For j = i + 1 To oppsett.ListRows.Count
            navn2 = Trim(CStr(oppsett.DataBodyRange.Cells(j, iNytt).Value))
            If navn1 <> "" And StrComp(navn1, navn2, vbTextCompare) = 0 Then
                FinnDuplikatNyttNavn = navn1
                Exit Function
            End If
        Next j
    Next i

    FinnDuplikatNyttNavn = ""
End Function

Private Function FinnDuplikatRekkefolge(oppsett As ListObject, iRekkefolge As Long) As String
    Dim i As Long, j As Long
    Dim verdi1 As Variant, verdi2 As Variant

    For i = 1 To oppsett.ListRows.Count
        verdi1 = oppsett.DataBodyRange.Cells(i, iRekkefolge).Value
        For j = i + 1 To oppsett.ListRows.Count
            verdi2 = oppsett.DataBodyRange.Cells(j, iRekkefolge).Value
            If IsNumeric(verdi1) And IsNumeric(verdi2) Then
                If CDbl(verdi1) = CDbl(verdi2) Then
                    FinnDuplikatRekkefolge = CStr(verdi1)
                    Exit Function
                End If
            End If
        Next j
    Next i

    FinnDuplikatRekkefolge = ""
End Function

' ---- Hent sortert konfigurasjon ----
' Kolonner i returnert array: 1 Original, 2 Nytt navn, 3 Kategori,
' 4 Rekkefolge, 5 Format (valgfri), 6 Justering (valgfri). Sortert
' stigende etter Rekkefolge.

Private Function HentSortertOppsett(oppsett As ListObject) As Variant
    Dim arr() As Variant
    Dim i As Long, j As Long, k As Long
    Dim temp As Variant
    Dim antall As Long

    Dim iOriginal As Long, iNytt As Long, iKategori As Long, iRekkefolge As Long
    Dim iFormat As Long, iJustering As Long

    iOriginal = OppsettKolonneIndeks(oppsett, "Original kolonne")
    iNytt = OppsettKolonneIndeks(oppsett, "Nytt navn")
    iKategori = OppsettKolonneIndeks(oppsett, "Kategori")
    iRekkefolge = OppsettKolonneIndeks(oppsett, "Rekkefølge")
    iFormat = OppsettKolonneIndeks(oppsett, "Format")
    iJustering = OppsettKolonneIndeks(oppsett, "Justering")

    antall = oppsett.ListRows.Count
    ReDim arr(1 To antall, 1 To 6)

    For i = 1 To antall
        arr(i, 1) = oppsett.DataBodyRange.Cells(i, iOriginal).Value
        arr(i, 2) = oppsett.DataBodyRange.Cells(i, iNytt).Value
        arr(i, 3) = oppsett.DataBodyRange.Cells(i, iKategori).Value
        arr(i, 4) = oppsett.DataBodyRange.Cells(i, iRekkefolge).Value

        If iFormat > 0 Then
            arr(i, 5) = CStr(oppsett.DataBodyRange.Cells(i, iFormat).Value)
        Else
            arr(i, 5) = ""
        End If

        If iJustering > 0 Then
            arr(i, 6) = CStr(oppsett.DataBodyRange.Cells(i, iJustering).Value)
        Else
            arr(i, 6) = ""
        End If
    Next i

    For i = 1 To antall - 1
        For j = i + 1 To antall
            If CDbl(arr(j, 4)) < CDbl(arr(i, 4)) Then
                For k = 1 To 6
                    temp = arr(i, k)
                    arr(i, k) = arr(j, k)
                    arr(j, k) = temp
                Next k
            End If
        Next j
    Next i

    HentSortertOppsett = arr
End Function

' ---- Organiser kolonner (rekkefolge) ----

Private Sub OrganiserKolonner(tbl As ListObject, data As Variant)
    Dim i As Long
    Dim originalNavn As String
    Dim onsketPosisjon As Long

    For i = LBound(data, 1) To UBound(data, 1)
        originalNavn = CStr(data(i, 1))
        onsketPosisjon = CLng(data(i, 4))
        FlyttKolonne tbl, originalNavn, onsketPosisjon
    Next i
End Sub

Private Sub FlyttKolonne(tbl As ListObject, kolonnenavn As String, onsketPosisjon As Long)
    Dim lc As ListColumn
    Dim maalKolonne As ListColumn

    Set lc = FinnListColumn(tbl, kolonnenavn)
    If lc Is Nothing Then Exit Sub
    If lc.Index = onsketPosisjon Then Exit Sub

    If onsketPosisjon > tbl.ListColumns.Count Then onsketPosisjon = tbl.ListColumns.Count
    If onsketPosisjon < 1 Then onsketPosisjon = 1

    lc.Range.Cut
    Set maalKolonne = tbl.ListColumns(onsketPosisjon)
    maalKolonne.Range.Insert Shift:=xlToRight
    Application.CutCopyMode = False
End Sub

' ---- Endre kolonnenavn ----

Private Sub EndreKolonnenavn(tbl As ListObject, data As Variant)
    Dim i As Long
    Dim originalNavn As String, nyttNavn As String
    Dim lc As ListColumn

    For i = LBound(data, 1) To UBound(data, 1)
        originalNavn = CStr(data(i, 1))
        nyttNavn = Trim(CStr(data(i, 2)))
        Set lc = FinnListColumn(tbl, originalNavn)
        If Not lc Is Nothing Then
            If StrComp(lc.Name, nyttNavn, vbTextCompare) <> 0 Then
                lc.Name = nyttNavn
            End If
        End If
    Next i
End Sub

' ---- Valgfri formatering/justering per kolonne (styrt fra tblOppsett) ----

Private Sub FormaterSpesialkolonner(tbl As ListObject, data As Variant)
    Dim i As Long
    Dim nyttNavn As String
    Dim formatStreng As String
    Dim justering As String
    Dim lc As ListColumn

    For i = LBound(data, 1) To UBound(data, 1)
        nyttNavn = Trim(CStr(data(i, 2)))
        formatStreng = Trim(CStr(data(i, 5)))
        justering = Trim(CStr(data(i, 6)))

        If formatStreng = "" And justering = "" Then GoTo NesteRad

        Set lc = FinnListColumn(tbl, nyttNavn)
        If lc Is Nothing Then GoTo NesteRad
        If lc.DataBodyRange Is Nothing Then GoTo NesteRad

        If formatStreng <> "" Then
            On Error Resume Next
            lc.DataBodyRange.NumberFormat = formatStreng
            On Error GoTo 0
        End If

        Select Case LCase(justering)
            Case "venstre"
                lc.DataBodyRange.HorizontalAlignment = xlLeft
            Case "hoyre", "høyre"
                lc.DataBodyRange.HorizontalAlignment = xlRight
            Case "sentrert", "midtstilt"
                lc.DataBodyRange.HorizontalAlignment = xlCenter
        End Select

NesteRad:
    Next i
End Sub

' ---- Kategorirad ----

Private Sub LagKategoriRad(tbl As ListObject, data As Variant, kategoriFarger As Object)
    Dim ws As Worksheet
    Dim rad As Long
    Dim forsteKolonne As Long, sisteKolonne As Long

    Set ws = tbl.Parent
    rad = tbl.Range.Row - 1

    If rad < 1 Then
        Err.Raise vbObjectError + 1100, , "Det finnes ingen ledig rad over tabellen for kategorirad. Sorg for at tabellen ikke starter paa rad 1."
    End If

    forsteKolonne = tbl.Range.Column
    sisteKolonne = tbl.Range.Column + tbl.ListColumns.Count - 1

    With ws.Range(ws.Cells(rad, forsteKolonne), ws.Cells(rad, sisteKolonne))
        .UnMerge
        .ClearContents
        .ClearFormats
    End With

    ByggKategorier ws, tbl, rad, data, kategoriFarger
End Sub

Private Function FinnKategoriForKolonne(data As Variant, kolonnenavn As String) As String
    Dim i As Long

    For i = LBound(data, 1) To UBound(data, 1)
        If StrComp(Trim(CStr(data(i, 2))), Trim(kolonnenavn), vbTextCompare) = 0 Then
            FinnKategoriForKolonne = Trim(CStr(data(i, 3)))
            Exit Function
        End If
    Next i

    FinnKategoriForKolonne = "ANNET"
End Function

Private Sub ByggKategorier(ws As Worksheet, tbl As ListObject, rad As Long, data As Variant, kategoriFarger As Object)
    Dim i As Long
    Dim kategori As String, forrigeKategori As String
    Dim startKol As Long, sluttKol As Long

    If tbl.ListColumns.Count = 0 Then Exit Sub

    startKol = tbl.Range.Column

    For i = 1 To tbl.ListColumns.Count
        kategori = FinnKategoriForKolonne(data, tbl.ListColumns(i).Name)

        If i = 1 Then
            forrigeKategori = kategori
        ElseIf kategori <> forrigeKategori Then
            sluttKol = tbl.Range.Column + i - 2
            LagKategori ws, rad, startKol, sluttKol, forrigeKategori, kategoriFarger
            startKol = tbl.Range.Column + i - 1
            forrigeKategori = kategori
        End If
    Next i

    sluttKol = tbl.Range.Column + tbl.ListColumns.Count - 1
    LagKategori ws, rad, startKol, sluttKol, forrigeKategori, kategoriFarger
End Sub

Private Sub LagKategori(ws As Worksheet, rad As Long, forsteKol As Long, sisteKol As Long, tekst As String, kategoriFarger As Object)
    If forsteKol > sisteKol Then Exit Sub

    With ws.Range(ws.Cells(rad, forsteKol), ws.Cells(rad, sisteKol))
        .Merge
        .Value = tekst
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = HentKategoriFarge(kategoriFarger, tekst)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

' ---- Kategorifarger (fra tblKategorier hvis satt, ellers auto-palett) ----

Private Function HentKategoriFarge(kategoriFarger As Object, kategori As String) As Long
    Dim nokkel As String
    nokkel = UCase(Trim(kategori))

    If kategoriFarger.Exists(nokkel) Then
        HentKategoriFarge = kategoriFarger(nokkel)
    Else
        HentKategoriFarge = RGB(90, 90, 90)
    End If
End Function

Private Function BuildKategoriFarger(data As Variant) As Object
    Dim farger As Object
    Dim kategoriTabell As ListObject
    Dim palett As Variant
    Dim i As Long, idx As Long
    Dim kategori As String
    Dim farge As Long

    Set farger = CreateObject("Scripting.Dictionary")
    Set kategoriTabell = FinnKategoriTabell()

    palett = Array(RGB(31, 78, 121), RGB(70, 130, 90), RGB(190, 110, 45), _
                    RGB(100, 75, 140), RGB(60, 120, 150), RGB(85, 105, 125), _
                    RGB(150, 70, 70), RGB(120, 110, 50))

    farger.Add "ANNET", RGB(90, 90, 90)
    idx = 0

    For i = LBound(data, 1) To UBound(data, 1)
        kategori = UCase(Trim(CStr(data(i, 3))))
        If kategori <> "" And Not farger.Exists(kategori) Then
            farge = -1
            If Not kategoriTabell Is Nothing Then
                farge = FinnFargeITabell(kategoriTabell, kategori)
            End If
            If farge = -1 Then
                farge = palett(idx Mod (UBound(palett) + 1))
                idx = idx + 1
            End If
            farger.Add kategori, farge
        End If
    Next i

    Set BuildKategoriFarger = farger
End Function

Private Function FinnFargeITabell(kategoriTabell As ListObject, kategori As String) As Long
    Dim iKategori As Long, iFarge As Long
    Dim r As ListRow
    Dim navn As String
    Dim hexStreng As String

    FinnFargeITabell = -1

    iKategori = OppsettKolonneIndeks(kategoriTabell, "Kategori")
    iFarge = OppsettKolonneIndeks(kategoriTabell, "Farge")

    If iKategori = 0 Or iFarge = 0 Then Exit Function
    If kategoriTabell.DataBodyRange Is Nothing Then Exit Function

    For Each r In kategoriTabell.ListRows
        navn = UCase(Trim(CStr(r.Range.Cells(1, iKategori).Value)))
        If navn = UCase(Trim(kategori)) Then
            hexStreng = Trim(CStr(r.Range.Cells(1, iFarge).Value))
            If hexStreng <> "" Then
                FinnFargeITabell = HexTilFarge(hexStreng)
            End If
            Exit Function
        End If
    Next r
End Function

Private Function HexTilFarge(hexStreng As String) As Long
    Dim s As String

    s = Trim(hexStreng)
    If Left(s, 1) = "#" Then s = Mid(s, 2)

    If Len(s) <> 6 Then
        HexTilFarge = -1
        Exit Function
    End If

    On Error GoTo Ugyldig

    HexTilFarge = RGB(CLng("&H" & Mid(s, 1, 2)), CLng("&H" & Mid(s, 3, 2)), CLng("&H" & Mid(s, 5, 2)))
    Exit Function

Ugyldig:
    HexTilFarge = -1
End Function

' ---- Generell tabellformatering ----

Private Sub FormaterTabell(tbl As ListObject)
    With tbl.Range
        .Font.Name = "Aptos"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With tbl.HeaderRowRange
        .Font.Bold = True
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(55, 55, 55)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    tbl.HeaderRowRange.RowHeight = 26

    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Rows.RowHeight = 21
    End If

    tbl.Range.Columns.AutoFit
End Sub

Private Sub FormaterKategoriRad(tbl As ListObject)
    Dim ws As Worksheet
    Dim rad As Long

    Set ws = tbl.Parent
    rad = tbl.Range.Row - 1

    ws.Rows(rad).RowHeight = 25
End Sub
'@

function Get-OrganiserVersion {
    param([string]$Code)
    if ($Code -match 'TABELLORGANISERING_VERSION\s*As\s*String\s*=\s*"([^"]+)"') {
        return $Matches[1]
    }
    return $null
}

$sourceVersion = Get-OrganiserVersion $moduleCode

# ---- 1. Finn maalfilen ----

if (-not $Path) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel-filer (*.xlsm)|*.xlsm|Alle filer (*.*)|*.*"
    $dlg.Title = "Velg Excel-fil som skal faa Tabellorganisering"
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
    # Avinstallerer - fjerner KUN den noyaktige, kjente komponenten
    # "modTabellorganisering" og knapper med denne makroens EGET
    # OnAction-navn ("OrganiserTabell"), aldri et generisk filter - et ark
    # kan ha flere uavhengig installerte makroers knapper side om side.
    # Arket "_Oppsett" med tblOppsett/tblKategorier lar vi staa urort med
    # vilje, slik at en senere reinstallasjon gjenbruker gammel konfigurasjon.
    Write-Output "Fjerner Tabellorganisering fra $($wb.Name) ..."
    try {
        $eksisterende = $vbproj.VBComponents.Item("modTabellorganisering")
        $vbproj.VBComponents.Remove($eksisterende)
        Write-Output "  Fjernet modTabellorganisering"
    } catch {}
    try {
        foreach ($ws in $wb.Worksheets) {
            foreach ($btn in @($ws.Buttons())) {
                if ($btn.OnAction -eq "OrganiserTabell") { $btn.Delete() }
            }
        }
    } catch {}
    try { $wb.Save() } catch {
        Write-Output "FEIL ved lagring: $($_.Exception.Message)"
        exit 1
    }
    Write-Output "Ferdig - Tabellorganisering er fjernet fra $($wb.Name)."
    exit 0
}

# ---- 3. Sjekk hva som allerede er installert ----

$existingModule = $null
foreach ($comp in $vbproj.VBComponents) {
    if ($comp.Name -eq "modTabellorganisering") { $existingModule = $comp }
}

$installedVersion = $null
if ($existingModule) {
    $existingCode = $existingModule.CodeModule.Lines(1, $existingModule.CodeModule.CountOfLines)
    $installedVersion = Get-OrganiserVersion $existingCode
}

$isFreshInstall = ($null -eq $existingModule)

if (-not $isFreshInstall -and $installedVersion -eq $sourceVersion) {
    Write-Output "Tabellorganisering er allerede installert og oppdatert (versjon $installedVersion) i $($wb.Name)."
    exit 0
}

if ($isFreshInstall) {
    Write-Output "Installerer Tabellorganisering (versjon $sourceVersion) i $($wb.Name) ..."
} else {
    Write-Output "Oppdaterer Tabellorganisering: versjon $installedVersion -> $sourceVersion i $($wb.Name) ..."
}

# ---- 4. Bygg modulen i en midlertidig arbeidsbok, eksporter til en temp-mappe ----

$tempDir = Join-Path $env:TEMP ("Tabellorganisering_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir | Out-Null

try {
    $buildWb = $excel.Workbooks.Add()
    $buildProj = $buildWb.VBProject

    $mod = $buildProj.VBComponents.Add(1)
    $mod.Name = "modTabellorganisering"
    $mod.CodeModule.AddFromString($moduleCode)

    $mod.Export((Join-Path $tempDir "modTabellorganisering.bas"))

    $buildWb.Close($false)

    # ---- 5. Bytt ut / legg til modulen i maal-arbeidsboken ----

    $filePath = Join-Path $tempDir "modTabellorganisering.bas"

    $existing = $null
    foreach ($comp in $vbproj.VBComponents) {
        if ($comp.Name -eq "modTabellorganisering") { $existing = $comp }
    }
    if ($existing) {
        $vbproj.VBComponents.Remove($existing)
    }
    $null = $vbproj.VBComponents.Import($filePath)
    Write-Output "  modTabellorganisering installert"

    # ---- 6. Ved fersk installasjon: legg til knapp + _Oppsett-ark ----

    if ($isFreshInstall) {
        $ws = $wb.ActiveSheet
        $tableCount = $ws.ListObjects.Count

        if ($tableCount -ne 1) {
            Write-Output "  ADVARSEL: Aktivt ark ('$($ws.Name)') har $tableCount tabeller (forventet 1). Knappen legges til likevel - flytt den til riktig ark om noedvendig."
        }

        $btn = $ws.Buttons().Add(10, 5, 140, 32)
        $btn.OnAction = "OrganiserTabell"
        $btn.Caption = "ORGANISER TABELL"
        $btn.Placement = 3   # xlFreeFloating: ikke flytt/endre storrelse med celler
        Write-Output "  Knapp lagt til i arket '$($ws.Name)' (kopier den til andre ark ved behov)"

        $dataTbl = $null
        if ($tableCount -ge 1) { $dataTbl = $ws.ListObjects.Item(1) }

        $oppsettWs = $null
        foreach ($sheet in $wb.Worksheets) {
            if ($sheet.Name -eq "_Oppsett") { $oppsettWs = $sheet }
        }

        if ($null -eq $oppsettWs) {
            $oppsettWs = $wb.Worksheets.Add()
            $oppsettWs.Move($wb.Worksheets.Item($wb.Worksheets.Count))
            $oppsettWs.Name = "_Oppsett"

            $oppsettWs.Range("A1").Value = "TABELLKONFIGURASJON"
            $oppsettWs.Range("A1").Font.Bold = $true
            $oppsettWs.Range("A1").Font.Size = 12

            $oppsettWs.Range("A2").Value = "Rediger radene under for aa endre rekkefolge, navn og kategori paa kolonnene i tabellen. Kjor deretter ORGANISER TABELL-knappen paa nytt."
            $oppsettWs.Range("A2").Font.Italic = $true

            $headerRow = 4
            $oppsettWs.Range("A4:F4").Value = @("Original kolonne", "Nytt navn", "Kategori", "Rekkefølge", "Format", "Justering")

            # Skriver hele radblokken som ett 2D-array til et Range i ett kall,
            # via .Value (IKKE .Value2). Bruk konsekvent .Value for ALLE
            # celletilordninger i dette scriptet - PowerShell 5.1 sin
            # COM-binding for .Value2 far et feil type-cache naar bade
            # streng- og array-tilordninger skjer i samme skript, og kaster da
            # tilfeldige "Unable to cast..."-feil. .Value har ikke dette
            # problemet. (Gjelder kun PowerShell -> COM; paavirker ikke
            # .Value2-bruk inne i selve VBA-koden over, som kjorer i Excel sin
            # egen VBA-motor.)
            $lastRow = $headerRow
            if ($null -ne $dataTbl -and $dataTbl.ListColumns.Count -gt 0) {
                $antallKol = $dataTbl.ListColumns.Count
                $arr = New-Object 'object[,]' $antallKol, 6
                $i = 0
                foreach ($lc in $dataTbl.ListColumns) {
                    $arr[$i, 0] = $lc.Name
                    $arr[$i, 1] = $lc.Name
                    $arr[$i, 2] = "ANNET"
                    $arr[$i, 3] = [double]($i + 1)
                    $arr[$i, 4] = ""
                    $arr[$i, 5] = ""
                    $i++
                }
                $lastRow = $headerRow + $antallKol
            } else {
                # Ingen datatabell funnet - legg til en tom eksempelrad
                $arr = New-Object 'object[,]' 1, 6
                $arr[0, 0] = "Kolonnenavn"
                $arr[0, 1] = "Kolonnenavn"
                $arr[0, 2] = "ANNET"
                $arr[0, 3] = [double]1
                $arr[0, 4] = ""
                $arr[0, 5] = ""
                $lastRow = $headerRow + 1
            }
            $dataSkriveRange = $oppsettWs.Range($oppsettWs.Cells.Item($headerRow + 1, 1), $oppsettWs.Cells.Item($lastRow, 6))
            $dataSkriveRange.Value = $arr

            $tblRange = $oppsettWs.Range($oppsettWs.Cells.Item($headerRow, 1), $oppsettWs.Cells.Item($lastRow, 6))
            $oppsettTbl = $oppsettWs.ListObjects.Add(1, $tblRange, [System.Type]::Missing, 1)
            $oppsettTbl.Name = "tblOppsett"
            $oppsettTbl.TableStyle = "TableStyleMedium2"

            $catNoteRow = $lastRow + 3
            $oppsettWs.Cells.Item($catNoteRow, 1).Value = "Valgfritt: overstyr kategorifarger her (Farge = HEX-kode, f.eks. 1F4E79). Ellers velges farge automatisk."
            $oppsettWs.Cells.Item($catNoteRow, 1).Font.Italic = $true

            $catHeaderRow = $catNoteRow + 1
            $catRange = $oppsettWs.Range($oppsettWs.Cells.Item($catHeaderRow, 1), $oppsettWs.Cells.Item($catHeaderRow, 2))
            $catRange.Value = @("Kategori", "Farge")
            $catTbl = $oppsettWs.ListObjects.Add(1, $catRange, [System.Type]::Missing, 1)
            $catTbl.Name = "tblKategorier"
            $catTbl.TableStyle = "TableStyleMedium2"

            $oppsettWs.Columns.Item("A:F").AutoFit() | Out-Null

            Write-Output "  Opprettet arket '_Oppsett' med tblOppsett (forhaandsutfylt fra gjeldende tabell) og tblKategorier"
        } else {
            Write-Output "  Arket '_Oppsett' fantes allerede - lar det vaere urort"
        }
    }

    Write-Output ""
    Write-Output "Ferdig. Husk aa lagre filen selv (Ctrl+S) naar du er klar."
} finally {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
