param(
    [string]$Path,
    [switch]$Uninstall,
    # Tvinger reinstallasjon selv om versjonen er lik (brukes av EMI "Reinstaller alle").
    [switch]$Force,
    # Kun for testing: start en EGEN Excel-prosess i stedet for a koble til den
    # som allerede kjorer. Samme monster som de andre makroene.
    [switch]$NyExcelInstans
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# Katalogmekker - alt-i-en installer/oppdaterer.
#
# Lager en kontaktfil (.vcf, vCard 3.0) fra en tabell, til import i Google
# Kontakter (contacts.google.com > Importer) eller direkte pa en Android-
# telefon. Telefonnummeret er nokkelen: hvert nummer renskrives til norsk
# standard (8 siffer; 0047 / +47 / 47 foran fjernes), og rader med tomt
# nummer, "..." / ellipse eller et nummer som ikke kan tolkes hoppes over og
# listes opp. Samme nummer to ganger i tabellen gir bare EN kontakt (den
# forste raden vinner, de andre listes som hoppet over).
#
# Navnet bygges fra en eller flere valgte kolonner, i den rekkefolgen de
# hukes av, med valgfritt "kun forste ord" per kolonne og et valgfritt
# skilletegn mellom kolonnene. En felles kategori (CATEGORIES - blir etikett
# i Google Kontakter) kan skrives inn; flere kategorier skilles med komma.
# Forhandsvisningen viser en eksempelkontakt fra tabellen og den ferdige
# vCard-teksten. Valgene huskes i et skjult ark i arbeidsboken.
#
# Oppdatering av eksisterende kontakter: en import legger alltid TIL
# kontakter. Fordi nummeret er nokkelen, finner Google Kontakter sin
# "Slå sammen og rett" duplikatene etter en ny import (se kunnskapsbasen,
# seksjon mac-katalogmekker).
#
# v0.1.0 (2026-10-12): prototype, skrevet i en sky-okt etter Hakons
# spesifikasjon. Ikke testet i Excel ennå - ma testes i en lokal okt mot
# Prosjektfiler\Testfiler (se testfila Katalogmekker-testkontakter.xlsx).
#
# VBA-komponenter (alle prefikset "Katalogmekker" for a unnga kollisjon med
# andre makroers komponentnavn):
#   modKatalogmekker       - innstillinger, nummerrensing, vCard-bygging,
#                            fillagring, inngangspunkt (AapneKatalogmekker)
#   clsKatalogmekkerNavn   - WithEvents-handler for de dynamiske
#                            avkrysningsboksene i navnelisten
#   frmKatalogmekker       - hovedvinduet (alle innstillinger + forhandsvisning)
# ============================================================================

$moduleCodeKatalogmekker = @'
Option Explicit

' Katalogmekker - lager en kontaktfil (.vcf / vCard 3.0) fra en tabell.
' Telefonnummeret er nokkelen (renskrives til 8 siffer, norsk standard).
' Valgene lagres i et skjult ark (KM_OPPSETT_ARK) som nokkel i kolonne A og
' verdi i kolonne B - samme monster som de andre makroene, slik at
' Makromeny sitt "Installert info"-vindu kan vise dem.

Public Const KATALOGMEKKER_VERSION As String = "0.1.0"
Public Const KM_OPPSETT_ARK As String = "KatalogmekkerOppsett_Hidden"

Public Sub AapneKatalogmekker()
    frmKatalogmekker.Show
End Sub

' ---- Innstillinger (skjult ark, nokkel i A / verdi i B) ----

Private Function HentOppsettArk(ByVal Opprett As Boolean) As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(KM_OPPSETT_ARK)
    On Error GoTo 0
    If sh Is Nothing Then
        If Not Opprett Then Exit Function
        Dim aktivt As Object
        On Error Resume Next
        Set aktivt = ActiveSheet
        On Error GoTo 0
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sh.Name = KM_OPPSETT_ARK
        sh.Range("A1:B1").Value = Array("Nokkel", "Verdi")
        sh.Visible = 2 ' xlSheetVeryHidden
        If Not aktivt Is Nothing Then
            On Error Resume Next
            aktivt.Activate
            On Error GoTo 0
        End If
    End If
    Set HentOppsettArk = sh
End Function

Private Function FinnInnstillingRad(ByVal sh As Worksheet, ByVal Nokkel As String) As Long
    Dim sisteRad As Long, r As Long
    sisteRad = sh.Cells(sh.Rows.Count, 1).End(-4162).Row ' xlUp
    For r = 2 To sisteRad
        If CStr(sh.Cells(r, 1).Value) = Nokkel Then
            FinnInnstillingRad = r
            Exit Function
        End If
    Next r
End Function

Public Function LesInnstilling(ByVal Nokkel As String, ByVal Standard As String) As String
    LesInnstilling = Standard
    Dim sh As Worksheet
    Set sh = HentOppsettArk(False)
    If sh Is Nothing Then Exit Function
    Dim r As Long
    r = FinnInnstillingRad(sh, Nokkel)
    If r > 0 Then LesInnstilling = CStr(sh.Cells(r, 2).Value)
End Function

' Verdien skrives med apostrof foran, slik at Excel lagrer den som tekst
' uendret (ledende mellomrom i skilletegnet, "=" eller tall som "0047"
' blir ellers tolket - se g-tekst-blir-tall-i-celle).
Public Sub SkrivInnstilling(ByVal Nokkel As String, ByVal Verdi As String)
    Dim sh As Worksheet
    Set sh = HentOppsettArk(True)
    Dim r As Long
    r = FinnInnstillingRad(sh, Nokkel)
    If r = 0 Then
        r = sh.Cells(sh.Rows.Count, 1).End(-4162).Row + 1
        If r < 2 Then r = 2
        sh.Cells(r, 1).Value = "'" & Nokkel
    End If
    sh.Cells(r, 2).Value = "'" & Verdi
End Sub

' ---- Tabeller ----

' Navnene pa alle SYNLIGE ark som har minst en tabell (ListObject).
Public Function SynligeArkMedTabeller() As Collection
    Dim res As Collection
    Set res = New Collection
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = -1 Then ' xlSheetVisible
            If ws.ListObjects.Count > 0 Then res.Add ws.Name
        End If
    Next ws
    Set SynligeArkMedTabeller = res
End Function

Public Function HentTabell(ByVal ArkNavn As String, ByVal TabellNavn As String) As ListObject
    On Error Resume Next
    Set HentTabell = ThisWorkbook.Worksheets(ArkNavn).ListObjects(TabellNavn)
    On Error GoTo 0
End Function

' Verdiene i en kolonne som 2D-array (1..n, 1..1), hentet i ETT COM-kall.
' Empty hvis kolonnen ikke finnes (f.eks. omdopt siden sist).
Private Function KolonneVerdier(ByVal tbl As ListObject, ByVal Navn As String, ByVal nRader As Long) As Variant
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = tbl.ListColumns(Navn)
    On Error GoTo 0
    If lc Is Nothing Then Exit Function
    Dim v As Variant
    If nRader = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = lc.DataBodyRange.Value2
    Else
        v = lc.DataBodyRange.Value2
    End If
    KolonneVerdier = v
End Function

' ---- Tekst og telefonnummer ----

' Celleverdi som trimmet tekst. Tall skrives uten desimaler og uten
' E-notasjon (et nummer lagret som tall, f.eks. 4798765432, skal bli
' "4798765432"). Linjeskift og hardt mellomrom blir vanlige mellomrom.
Public Function CelleTekst(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbDecimal Then
        If v = Int(v) And Abs(v) < 1E+15 Then
            CelleTekst = Format$(v, "0")
            Exit Function
        End If
    End If
    Dim s As String
    s = CStr(v)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW(160), " ")
    CelleTekst = Trim$(s)
End Function

Public Function ForsteOrd(ByVal s As String) As String
    s = Trim$(s)
    Dim p As Long
    p = InStr(s, " ")
    If p > 0 Then
        ForsteOrd = Left$(s, p - 1)
    Else
        ForsteOrd = s
    End If
End Function

' True hvis teksten inneholder "..." eller ellipsetegnet (U+2026).
Public Function HarEllipse(ByVal s As String) As Boolean
    HarEllipse = (InStr(s, "...") > 0) Or (InStr(s, ChrW(8230)) > 0)
End Function

' Renskriver et telefonnummer til norsk standard: 8 siffer, uten landskode.
' Alt som ikke er siffer fjernes (mellomrom, bindestrek, punktum, parentes,
' tekst som "Mob:"); et "+" for forste siffer betyr at landskoden folger.
'   +47 / + 47 / 0047 foran  -> fjernes
'   47 foran, og nummeret er lenger enn 8 siffer -> 47 fjernes
'   +46..., 0046... (annet land) -> avvist (ikke norsk nummer)
' Returnerer "" hvis nummeret ikke kan tolkes; Grunn far da en kort forklaring.
Public Function NormaliserNummer(ByVal Raa As String, ByRef Grunn As String) As String
    Dim i As Long, ch As String, t As String
    Dim harPluss As Boolean, harSiffer As Boolean
    Grunn = ""
    For i = 1 To Len(Raa)
        ch = Mid$(Raa, i, 1)
        If ch >= "0" And ch <= "9" Then
            t = t & ch
            harSiffer = True
        ElseIf ch = "+" And Not harSiffer Then
            harPluss = True
        End If
    Next i

    If t = "" Then
        Grunn = "ingen siffer"
        Exit Function
    End If

    If harPluss Then
        If Left$(t, 2) = "47" Then
            t = Mid$(t, 3)
        Else
            Grunn = "utenlandsk nummer"
            Exit Function
        End If
    ElseIf Left$(t, 4) = "0047" Then
        t = Mid$(t, 5)
    ElseIf Left$(t, 2) = "00" Then
        Grunn = "utenlandsk nummer"
        Exit Function
    ElseIf Len(t) > 8 And Left$(t, 2) = "47" Then
        t = Mid$(t, 3)
    End If

    If Len(t) <> 8 Then
        Grunn = "ikke 8 siffer"
        Exit Function
    End If
    NormaliserNummer = t
End Function

Public Function FormaterNummer(ByVal Nummer As String, ByVal MedLandskode As Boolean) As String
    If MedLandskode Then
        FormaterNummer = "+47" & Nummer
    Else
        FormaterNummer = Nummer
    End If
End Function

' ---- Kontaktliste ----

' Leser tabellen og bygger kontaktlisten.
'   NavnKolonner: Collection av Array(kolonnenavn, kunForsteOrd) i
'                 navnets rekkefolge.
'   Kontakter:    Collection av Array(radNr, navn, nummer) - nummer er
'                 renskrevet (8 siffer), navn kan vaere "" (uten navn).
'   Hoppet:       Collection av Array(radNr, grunn).
'   AntallTomme:  rader der bade nummeret og alle navnekolonnene er tomme
'                 (telles, men listes ikke enkeltvis).
' radNr er radnummeret i ARKET (det brukeren ser i Excel), ikke i tabellen.
Public Sub ByggKontakter(ByVal tbl As ListObject, ByVal NummerKolonne As String, ByVal NavnKolonner As Collection, _
        ByVal Separator As String, ByVal KunSynlige As Boolean, ByRef Kontakter As Collection, _
        ByRef Hoppet As Collection, ByRef AntallTomme As Long, ByRef AntallUtenNavn As Long)

    Set Kontakter = New Collection
    Set Hoppet = New Collection
    AntallTomme = 0
    AntallUtenNavn = 0

    If tbl Is Nothing Then Exit Sub
    ' En tom tabell har ingen DataBodyRange (Nothing, eller feil i eldre Excel).
    Dim kropp As Range
    On Error Resume Next
    Set kropp = tbl.DataBodyRange
    On Error GoTo 0
    If kropp Is Nothing Then Exit Sub

    Dim nRader As Long
    nRader = kropp.Rows.Count

    Dim nrVals As Variant
    nrVals = KolonneVerdier(tbl, NummerKolonne, nRader)
    If Not IsArray(nrVals) Then Exit Sub

    Dim nNavn As Long, k As Long
    nNavn = NavnKolonner.Count
    Dim navnVals() As Variant, forste() As Boolean
    If nNavn > 0 Then
        ReDim navnVals(1 To nNavn)
        ReDim forste(1 To nNavn)
        Dim par As Variant
        For k = 1 To nNavn
            par = NavnKolonner(k)
            navnVals(k) = KolonneVerdier(tbl, CStr(par(0)), nRader)
            forste(k) = CBool(par(1))
        Next k
    End If

    Dim sett As Object ' nummer -> radNr for forste forekomst
    Set sett = CreateObject("Scripting.Dictionary")

    Dim forsteRad As Long
    forsteRad = kropp.Row

    Dim r As Long, radNr As Long
    Dim raa As String, navn As String, del As String, grunn As String, nummer As String
    Dim ellipse As Boolean, tom As Boolean
    For r = 1 To nRader
        radNr = forsteRad + r - 1

        If KunSynlige Then
            If kropp.Rows(r).EntireRow.Hidden Then GoTo NesteRad
        End If

        raa = CelleTekst(nrVals(r, 1))
        ellipse = HarEllipse(raa)
        tom = (raa = "")
        navn = ""
        For k = 1 To nNavn
            If IsArray(navnVals(k)) Then
                del = CelleTekst(navnVals(k)(r, 1))
                If del <> "" Then tom = False
                If HarEllipse(del) Then ellipse = True
                If forste(k) Then del = ForsteOrd(del)
                If del <> "" Then
                    If navn <> "" Then navn = navn & Separator
                    navn = navn & del
                End If
            End If
        Next k

        If tom Then
            AntallTomme = AntallTomme + 1
        ElseIf ellipse Then
            Hoppet.Add Array(radNr, "inneholder ""...""")
        ElseIf raa = "" Then
            Hoppet.Add Array(radNr, "mangler nummer")
        Else
            nummer = NormaliserNummer(raa, grunn)
            If nummer = "" Then
                Hoppet.Add Array(radNr, grunn & " (" & raa & ")")
            ElseIf sett.Exists(nummer) Then
                Hoppet.Add Array(radNr, "samme nummer som rad " & sett(nummer))
            Else
                sett.Add nummer, radNr
                If navn = "" Then AntallUtenNavn = AntallUtenNavn + 1
                Kontakter.Add Array(radNr, navn, nummer)
            End If
        End If
NesteRad:
    Next r
End Sub

' ---- vCard ----

' Tegn med spesiell betydning i vCard-verdier (RFC 2426) escapes.
Public Function VcardEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, ",", "\,")
    s = Replace(s, ";", "\;")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbCr, "\n")
    VcardEscape = s
End Function

' Kategoriteksten slik den skrives i CATEGORIES: flere kategorier skilles
' med komma (Google Kontakter gjor hver til en egen etikett).
Public Function KategoriListe(ByVal Kategori As String) As String
    Dim deler() As String, i As Long, d As String, res As String
    deler = Split(Kategori, ",")
    For i = LBound(deler) To UBound(deler)
        d = Trim$(deler(i))
        If d <> "" Then
            d = Replace(d, "\", "\\")
            d = Replace(d, ";", "\;")
            If res <> "" Then res = res & ","
            res = res & d
        End If
    Next i
    KategoriListe = res
End Function

' En kontakt som vCard 3.0-tekst (CRLF-linjeskift). Hele navnet legges i
' fornavnfeltet i N (navnet er satt sammen fra valgfrie kolonner og kan ikke
' deles sikkert i for- og etternavn). Uten navn brukes nummeret som navn.
' Lange linjer foldes ikke (RFC-en anbefaler det over 75 byte, men Google og
' Android leser ufoldede linjer fint).
Public Function ByggVcard(ByVal Navn As String, ByVal Nummer As String, ByVal Kategori As String, ByVal MedLandskode As Boolean) As String
    Dim visNavn As String, tlf As String, kat As String, s As String
    tlf = FormaterNummer(Nummer, MedLandskode)
    visNavn = Navn
    If visNavn = "" Then visNavn = tlf
    s = "BEGIN:VCARD" & vbCrLf
    s = s & "VERSION:3.0" & vbCrLf
    s = s & "N:;" & VcardEscape(visNavn) & ";;;" & vbCrLf
    s = s & "FN:" & VcardEscape(visNavn) & vbCrLf
    s = s & "TEL;TYPE=CELL:" & tlf & vbCrLf
    kat = KategoriListe(Kategori)
    If kat <> "" Then s = s & "CATEGORIES:" & kat & vbCrLf
    s = s & "END:VCARD" & vbCrLf
    ByggVcard = s
End Function

' ---- Fil ----

' Lagringsdialog. Returnerer "" hvis brukeren avbryter. Foreslar samme mappe
' som arbeidsboken, eller Nedlastinger hvis arbeidsboken ligger i skyen
' (ThisWorkbook.Path er da en URL - se g-thisworkbook-path-url).
Public Function VelgLagringssti(ByVal ForslagNavn As String) As String
    Dim mappe As String
    mappe = ThisWorkbook.Path
    If mappe = "" Or InStr(1, mappe, "://") > 0 Then mappe = Environ$("USERPROFILE") & "\Downloads"

    Dim v As Variant
    v = Application.GetSaveAsFilename(InitialFileName:=mappe & "\" & ForslagNavn, _
        FileFilter:="vCard-kontakter (*.vcf), *.vcf", Title:="Lagre kontaktfil")
    If VarType(v) = vbBoolean Then Exit Function

    Dim sti As String
    sti = CStr(v)
    If LCase$(Right$(sti, 4)) <> ".vcf" Then sti = sti & ".vcf"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sti) Then
        If MsgBox("Fila finnes allerede:" & vbCrLf & sti & vbCrLf & vbCrLf & "Vil du erstatte den?", _
                vbQuestion + vbYesNo, "Katalogmekker") <> vbYes Then Exit Function
    End If
    VelgLagringssti = sti
End Function

' Skriver teksten som UTF-8 UTEN BOM (aeoeaa blir riktige i Google og pa
' Android; noen vCard-lesere tar BOM-en med i "BEGIN:VCARD").
Public Function SkrivUtf8Fil(ByVal Sti As String, ByVal Tekst As String) As Boolean
    On Error GoTo Feil
    Dim st As Object, ut As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText Tekst
    st.Position = 3 ' hopp over BOM-en ADODB skriver forst
    Set ut = CreateObject("ADODB.Stream")
    ut.Type = 1 ' adTypeBinary
    ut.Open
    st.CopyTo ut
    ut.SaveToFile Sti, 2 ' adSaveCreateOverWrite
    ut.Close
    st.Close
    SkrivUtf8Fil = True
    Exit Function
Feil:
    MsgBox "Kunne ikke lagre fila:" & vbCrLf & Sti & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Katalogmekker"
End Function

' ---- Utseende: samme "farget knapp"-teknikk som Statistikkern/Kontakt-
' sentralen, egen kopi (Km-prefiks) slik at makroen er selvstendig. ----

Public Sub KmStilKnapp(ByVal Knapp As Object, ByVal Fargenavn As String)
    Dim farge As Long
    Select Case Fargenavn
        Case "gronn": farge = RGB(77, 124, 15)
        Case "blagra": farge = RGB(91, 107, 138)
        Case Else: farge = RGB(138, 138, 132)
    End Select
    On Error Resume Next
    Knapp.Style = 1 ' fmButtonStyleGraphical - kreves for at BackColor skal vises
    Knapp.BackColor = farge
    Knapp.ForeColor = RGB(255, 255, 255)
    Knapp.Font.Bold = True
End Sub
'@

$classCodeKatalogmekkerNavn = @'
Option Explicit

' En rad i navnelisten i frmKatalogmekker: avkrysning for kolonnen + "kun
' forste ord". Kontrollene lages dynamisk (antall kolonner varierer), og
' dynamiske kontroller far bare hendelser gjennom en WithEvents-klasse som
' holdes i live i en Collection pa skjemaet (g-dynamiske-kontroller-withevents).

Public WithEvents Chk As MSForms.CheckBox
Public WithEvents ChkForste As MSForms.CheckBox
Public Indeks As Long
Public Forelder As Object

Private Sub Chk_Click()
    If Not Forelder Is Nothing Then Forelder.NavnValgEndret Indeks
End Sub

Private Sub ChkForste_Click()
    If Not Forelder Is Nothing Then Forelder.ForsteOrdEndret
End Sub
'@

$formCodeKatalogmekker = @'
Option Explicit

Private mLaster As Boolean           ' True mens skjemaet fylles fra kode - demper hendelsene
Private mNavnRader As Collection     ' clsKatalogmekkerNavn per kolonne (holder handlerne i live)
Private mRekkefolge As Collection    ' kolonneindekser i den rekkefolgen de ble huket av
Private mKontakter As Collection     ' Array(radNr, navn, nummer) fra siste beregning
Private mHoppet As Collection        ' Array(radNr, grunn)
Private mAntallTomme As Long
Private mAntallUtenNavn As Long
Private mEksempel As Long            ' indeks i mKontakter som vises i forhandsvisningen

Private Sub UserForm_Initialize()
    mLaster = True
    Me.Caption = "Katalogmekker"
    Me.BackColor = RGB(246, 247, 249)
    Me.fraNavn.BackColor = RGB(255, 255, 255)
    Me.fraForhand.BackColor = RGB(255, 255, 255)

    Dim overskrift As Variant
    For Each overskrift In Array("lblKilde", "lblNummerTittel", "lblNavnTittel", "lblKategoriTittel", "lblForhandTittel")
        Me.Controls(CStr(overskrift)).Font.Bold = True
        Me.Controls(CStr(overskrift)).Font.Size = 10
        Me.Controls(CStr(overskrift)).ForeColor = RGB(63, 98, 18)
    Next overskrift
    For Each overskrift In Array("lblNummerHjelp", "lblNavnHjelp", "lblKategoriHjelp", "lblEksempel", "lblVcardTittel")
        Me.Controls(CStr(overskrift)).ForeColor = RGB(110, 116, 128)
    Next overskrift

    Me.lblPrevNavn.Font.Bold = True
    Me.lblPrevNavn.Font.Size = 12
    Me.txtVcard.Font.Name = "Consolas"
    Me.txtVcard.Font.Size = 8

    modKatalogmekker.KmStilKnapp Me.cmdLagVcf, "gronn"
    modKatalogmekker.KmStilKnapp Me.cmdLukk, "gra"
    modKatalogmekker.KmStilKnapp Me.cmdHoppet, "blagra"
    modKatalogmekker.KmStilKnapp Me.cmdForrige, "blagra"
    modKatalogmekker.KmStilKnapp Me.cmdNeste, "blagra"

    Set mRekkefolge = New Collection
    Set mNavnRader = New Collection
    mEksempel = 1

    Me.txtSeparator.Value = modKatalogmekker.LesInnstilling("Separator", " ")
    Me.txtKategori.Value = modKatalogmekker.LesInnstilling("Kategori", "")
    Me.chkLandskode.Value = (modKatalogmekker.LesInnstilling("Landskode", "0") = "1")
    Me.chkKunSynlige.Value = (modKatalogmekker.LesInnstilling("KunSynlige", "0") = "1")

    FyllArkliste
    FyllTabelliste
    mLaster = False
    Beregn
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    LagreInnstillinger
    SlippHandlere
End Sub

' Klasseinstansene peker tilbake pa skjemaet (Forelder) - slipp dem ved
' lukking sa skjemaet ikke holdes i live av en sirkulaer referanse.
Private Sub SlippHandlere()
    Dim h As Variant
    If Not mNavnRader Is Nothing Then
        For Each h In mNavnRader
            Set h.Forelder = Nothing
        Next h
    End If
    Set mNavnRader = Nothing
End Sub

' ---- Fylling av listene ----

Private Function FinnIListe(ByVal cbo As MSForms.ComboBox, ByVal Tekst As String) As Long
    FinnIListe = -1
    If Tekst = "" Then Exit Function
    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If CStr(cbo.List(i)) = Tekst Then
            FinnIListe = i
            Exit Function
        End If
    Next i
End Function

Private Sub FyllArkliste()
    Dim ark As Collection
    Set ark = modKatalogmekker.SynligeArkMedTabeller()
    Me.cboArk.Clear
    Dim i As Long
    For i = 1 To ark.Count
        Me.cboArk.AddItem CStr(ark(i))
    Next i
    If ark.Count = 0 Then Exit Sub

    Dim idx As Long, aktivtNavn As String
    idx = FinnIListe(Me.cboArk, modKatalogmekker.LesInnstilling("Ark", ""))
    If idx < 0 Then
        On Error Resume Next
        aktivtNavn = ActiveSheet.Name
        On Error GoTo 0
        idx = FinnIListe(Me.cboArk, aktivtNavn)
    End If
    If idx < 0 Then idx = 0
    Me.cboArk.ListIndex = idx
End Sub

Private Sub FyllTabelliste()
    Dim gammel As Boolean
    gammel = mLaster
    mLaster = True
    Me.cboTabell.Clear

    Dim ws As Worksheet
    If Me.cboArk.ListIndex >= 0 Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(Me.cboArk.Value))
        On Error GoTo 0
    End If
    If Not ws Is Nothing Then
        Dim lo As ListObject
        For Each lo In ws.ListObjects
            Me.cboTabell.AddItem lo.Name
        Next lo
        If Me.cboTabell.ListCount > 0 Then
            Dim idx As Long
            idx = FinnIListe(Me.cboTabell, modKatalogmekker.LesInnstilling("Tabell", ""))
            If idx < 0 Then idx = 0
            Me.cboTabell.ListIndex = idx
        End If
    End If

    ByggKolonner
    mLaster = gammel
    If Not mLaster Then Beregn
End Sub

Private Function ValgtTabell() As ListObject
    If Me.cboArk.ListIndex < 0 Or Me.cboTabell.ListIndex < 0 Then Exit Function
    Set ValgtTabell = modKatalogmekker.HentTabell(CStr(Me.cboArk.Value), CStr(Me.cboTabell.Value))
End Function

' Forste kolonne med et navn som ligner pa et telefonnummer, ellers -1.
Private Function GjettNummerkolonne() As Long
    GjettNummerkolonne = -1
    Dim i As Long, n As String, ord As Variant
    For i = 0 To Me.cboNummer.ListCount - 1
        n = LCase$(CStr(Me.cboNummer.List(i)))
        For Each ord In Array("mobil", "telefon", "tlf", "nummer", "phone")
            If InStr(n, CStr(ord)) > 0 Then
                GjettNummerkolonne = i
                Exit Function
            End If
        Next ord
    Next i
End Function

' Lagret navnevalg: en linje per kolonne, "1|Kolonnenavn" (1 = kun forste
' ord, 0 = hele cellen), i navnets rekkefolge.
Private Function LesNavnevalg() As Collection
    Dim res As Collection
    Set res = New Collection
    Dim linjer() As String, i As Long, p As Long
    linjer = Split(modKatalogmekker.LesInnstilling("Navnekolonner", ""), vbLf)
    For i = LBound(linjer) To UBound(linjer)
        p = InStr(linjer(i), "|")
        If p > 1 Then res.Add Array(Mid$(linjer(i), p + 1), Left$(linjer(i), p - 1) = "1")
    Next i
    Set LesNavnevalg = res
End Function

Private Function KolonneIndeks(ByVal tbl As ListObject, ByVal Navn As String) As Long
    Dim c As Long
    For c = 1 To tbl.ListColumns.Count
        If tbl.ListColumns(c).Name = Navn Then
            KolonneIndeks = c
            Exit Function
        End If
    Next c
End Function

Private Function IRekkefolge(ByVal Indeks As Long) As Boolean
    Dim v As Variant
    For Each v In mRekkefolge
        If CLng(v) = Indeks Then
            IRekkefolge = True
            Exit Function
        End If
    Next v
End Function

Private Sub FjernFraRekkefolge(ByVal Indeks As Long)
    Dim i As Long
    For i = mRekkefolge.Count To 1 Step -1
        If CLng(mRekkefolge(i)) = Indeks Then mRekkefolge.Remove i
    Next i
End Sub

' Bygger nummerlisten og navnelisten (en rad per kolonne i tabellen) pa
' nytt, og gjenoppretter lagrede valg for kolonner som fortsatt finnes.
Private Sub ByggKolonner()
    Dim gammel As Boolean
    gammel = mLaster
    mLaster = True

    SlippHandlere
    Dim i As Long
    For i = Me.fraNavn.Controls.Count - 1 To 0 Step -1
        Me.fraNavn.Controls.Remove Me.fraNavn.Controls(i).Name
    Next i
    Set mNavnRader = New Collection
    Set mRekkefolge = New Collection
    Me.cboNummer.Clear

    Dim tbl As ListObject
    Set tbl = ValgtTabell()
    If tbl Is Nothing Then
        mLaster = gammel
        Exit Sub
    End If

    Dim c As Long, n As Long
    n = tbl.ListColumns.Count
    For c = 1 To n
        Me.cboNummer.AddItem tbl.ListColumns(c).Name
    Next c
    Dim idx As Long
    idx = FinnIListe(Me.cboNummer, modKatalogmekker.LesInnstilling("Nummerkolonne", ""))
    If idx < 0 Then idx = GjettNummerkolonne()
    If idx >= 0 Then Me.cboNummer.ListIndex = idx

    Dim y As Single
    y = 4
    For c = 1 To n
        Dim lbl As MSForms.Label, chk As MSForms.CheckBox, chkF As MSForms.CheckBox
        Set lbl = Me.fraNavn.Controls.Add("Forms.Label.1", "lblNr" & c, True)
        lbl.Left = 4: lbl.Top = y + 3: lbl.Width = 20: lbl.Height = 14
        lbl.Caption = ""
        lbl.TextAlign = fmTextAlignRight
        lbl.Font.Bold = True
        lbl.ForeColor = RGB(77, 124, 15)
        lbl.BackStyle = fmBackStyleTransparent

        Set chk = Me.fraNavn.Controls.Add("Forms.CheckBox.1", "chkNavn" & c, True)
        chk.Left = 28: chk.Top = y: chk.Width = 190: chk.Height = 18
        chk.Caption = tbl.ListColumns(c).Name
        chk.BackStyle = fmBackStyleTransparent

        Set chkF = Me.fraNavn.Controls.Add("Forms.CheckBox.1", "chkForste" & c, True)
        chkF.Left = 222: chkF.Top = y: chkF.Width = 112: chkF.Height = 18
        chkF.Caption = "Kun første ord"
        chkF.BackStyle = fmBackStyleTransparent

        Dim h As clsKatalogmekkerNavn
        Set h = New clsKatalogmekkerNavn
        Set h.Chk = chk
        Set h.ChkForste = chkF
        h.Indeks = c
        Set h.Forelder = Me
        mNavnRader.Add h
        y = y + 22
    Next c
    Me.fraNavn.ScrollHeight = y + 4
    Me.fraNavn.ScrollTop = 0

    Dim lagrede As Collection, k As Long, par As Variant, ci As Long
    Set lagrede = LesNavnevalg()
    For k = 1 To lagrede.Count
        par = lagrede(k)
        ci = KolonneIndeks(tbl, CStr(par(0)))
        If ci > 0 Then
            If Not IRekkefolge(ci) Then
                Me.fraNavn.Controls("chkNavn" & ci).Value = True
                Me.fraNavn.Controls("chkForste" & ci).Value = CBool(par(1))
                mRekkefolge.Add ci
            End If
        End If
    Next k
    OppdaterNummerering

    mLaster = gammel
    If Not mLaster Then Beregn
End Sub

' Viser 1., 2., 3. ... foran de avkryssede kolonnene (navnets rekkefolge).
Private Sub OppdaterNummerering()
    Dim i As Long
    For i = 1 To mNavnRader.Count
        Me.fraNavn.Controls("lblNr" & i).Caption = ""
    Next i
    For i = 1 To mRekkefolge.Count
        Me.fraNavn.Controls("lblNr" & CLng(mRekkefolge(i))).Caption = i & "."
    Next i
End Sub

' ---- Kalles fra clsKatalogmekkerNavn ----

Public Sub NavnValgEndret(ByVal Indeks As Long)
    If mLaster Then Exit Sub
    If CBool(Me.fraNavn.Controls("chkNavn" & Indeks).Value) Then
        If Not IRekkefolge(Indeks) Then mRekkefolge.Add Indeks
    Else
        FjernFraRekkefolge Indeks
    End If
    OppdaterNummerering
    Beregn
End Sub

Public Sub ForsteOrdEndret()
    If mLaster Then Exit Sub
    Beregn
End Sub

' ---- Beregning og forhandsvisning ----

Private Function NavneValg() As Collection
    Dim res As Collection
    Set res = New Collection
    Dim v As Variant
    For Each v In mRekkefolge
        res.Add Array(CStr(Me.fraNavn.Controls("chkNavn" & CLng(v)).Caption), _
                      CBool(Me.fraNavn.Controls("chkForste" & CLng(v)).Value))
    Next v
    Set NavneValg = res
End Function

Private Sub Beregn()
    Set mKontakter = New Collection
    Set mHoppet = New Collection
    mAntallTomme = 0
    mAntallUtenNavn = 0

    Dim tbl As ListObject
    Set tbl = ValgtTabell()
    If tbl Is Nothing Then
        If Me.cboArk.ListCount = 0 Then
            Me.lblStatus.Caption = "Fant ingen tabeller i denne arbeidsboken. Gjør området om til en tabell først (Sett inn > Tabell)."
        Else
            Me.lblStatus.Caption = "Velg ark og tabell."
        End If
    ElseIf Me.cboNummer.ListIndex < 0 Then
        Me.lblStatus.Caption = "Velg kolonnen med telefonnummer."
    Else
        modKatalogmekker.ByggKontakter tbl, CStr(Me.cboNummer.Value), NavneValg(), CStr(Me.txtSeparator.Value), _
            CBool(Me.chkKunSynlige.Value), mKontakter, mHoppet, mAntallTomme, mAntallUtenNavn
        VisOppsummering
    End If

    If mEksempel > mKontakter.Count Then mEksempel = mKontakter.Count
    If mEksempel < 1 Then mEksempel = 1
    VisForhand
    Me.cmdLagVcf.Enabled = (mKontakter.Count > 0)
    Me.cmdHoppet.Enabled = (mHoppet.Count > 0)
End Sub

Private Sub VisOppsummering()
    Dim s As String
    s = mKontakter.Count & " kontakt(er) blir med i fila."
    If mHoppet.Count > 0 Then s = s & vbLf & mHoppet.Count & " rad(er) hoppes over - trykk «Vis hoppet over»."
    If mAntallTomme > 0 Then s = s & vbLf & mAntallTomme & " tomme rad(er) ignoreres."
    If mAntallUtenNavn > 0 Then s = s & vbLf & mAntallUtenNavn & " kontakt(er) uten navn får nummeret som navn."
    If mRekkefolge.Count = 0 Then s = s & vbLf & "Tips: huk av minst én kolonne under «3. Navn»."
    Me.lblStatus.Caption = s
End Sub

Private Sub VisForhand()
    Dim landskode As Boolean
    landskode = CBool(Me.chkLandskode.Value)
    If mKontakter.Count = 0 Then
        Me.lblPrevNavn.Caption = "(ingen kontakter å vise)"
        Me.lblPrevNummer.Caption = ""
        Me.lblPrevKategori.Caption = ""
        Me.lblEksempel.Caption = ""
        Me.txtVcard.Value = ""
        Me.cmdForrige.Enabled = False
        Me.cmdNeste.Enabled = False
        Exit Sub
    End If

    Dim k As Variant
    k = mKontakter(mEksempel)
    If CStr(k(1)) = "" Then
        Me.lblPrevNavn.Caption = "(uten navn - nummeret brukes som navn)"
    Else
        Me.lblPrevNavn.Caption = CStr(k(1))
    End If
    Me.lblPrevNummer.Caption = "Mobil:   " & modKatalogmekker.FormaterNummer(CStr(k(2)), landskode)

    Dim kat As String
    kat = Replace(modKatalogmekker.KategoriListe(CStr(Me.txtKategori.Value)), ",", ", ")
    If kat = "" Then kat = "(ingen)"
    Me.lblPrevKategori.Caption = "Etikett:   " & kat
    Me.lblEksempel.Caption = "Eksempel " & mEksempel & " av " & mKontakter.Count & "  ·  rad " & k(0) & " i arket"
    Me.txtVcard.Value = modKatalogmekker.ByggVcard(CStr(k(1)), CStr(k(2)), CStr(Me.txtKategori.Value), landskode)
    Me.cmdForrige.Enabled = (mEksempel > 1)
    Me.cmdNeste.Enabled = (mEksempel < mKontakter.Count)
End Sub

' Liste over hoppede rader (inntil Maks linjer) til MsgBox.
Private Function HoppetTekst(ByVal Maks As Long) As String
    If mHoppet.Count = 0 Then Exit Function
    Dim s As String, i As Long, h As Variant
    s = mHoppet.Count & " rad(er) ble hoppet over:" & vbCrLf
    For i = 1 To mHoppet.Count
        If i > Maks Then
            s = s & "  ... og " & (mHoppet.Count - Maks) & " til." & vbCrLf
            Exit For
        End If
        h = mHoppet(i)
        s = s & "  Rad " & h(0) & ": " & h(1) & vbCrLf
    Next i
    HoppetTekst = s
End Function

' Lagres ved lukking og nar fila lages. Uten tabeller i arbeidsboken er det
' ingenting a huske (og da lages heller ikke det skjulte arket).
Private Sub LagreInnstillinger()
    If Me.cboArk.ListCount = 0 Then Exit Sub
    On Error Resume Next
    If Me.cboArk.ListIndex >= 0 Then modKatalogmekker.SkrivInnstilling "Ark", CStr(Me.cboArk.Value)
    If Me.cboTabell.ListIndex >= 0 Then modKatalogmekker.SkrivInnstilling "Tabell", CStr(Me.cboTabell.Value)
    If Me.cboNummer.ListIndex >= 0 Then modKatalogmekker.SkrivInnstilling "Nummerkolonne", CStr(Me.cboNummer.Value)
    If Me.cboTabell.ListIndex >= 0 Then
        Dim s As String, v As Variant, par As Variant
        For Each v In NavneValg()
            par = v
            If s <> "" Then s = s & vbLf
            s = s & IIf(CBool(par(1)), "1", "0") & "|" & CStr(par(0))
        Next v
        modKatalogmekker.SkrivInnstilling "Navnekolonner", s
    End If
    modKatalogmekker.SkrivInnstilling "Separator", CStr(Me.txtSeparator.Value)
    modKatalogmekker.SkrivInnstilling "Kategori", CStr(Me.txtKategori.Value)
    modKatalogmekker.SkrivInnstilling "Landskode", IIf(CBool(Me.chkLandskode.Value), "1", "0")
    modKatalogmekker.SkrivInnstilling "KunSynlige", IIf(CBool(Me.chkKunSynlige.Value), "1", "0")
End Sub

' ---- Hendelser ----

Private Sub cboArk_Change()
    If mLaster Then Exit Sub
    FyllTabelliste
End Sub

Private Sub cboTabell_Change()
    If mLaster Then Exit Sub
    ByggKolonner
End Sub

Private Sub cboNummer_Change()
    If mLaster Then Exit Sub
    Beregn
End Sub

Private Sub txtSeparator_Change()
    If mLaster Then Exit Sub
    Beregn
End Sub

Private Sub txtKategori_Change()
    If mLaster Then Exit Sub
    VisForhand
End Sub

Private Sub chkLandskode_Click()
    If mLaster Then Exit Sub
    VisForhand
End Sub

Private Sub chkKunSynlige_Click()
    If mLaster Then Exit Sub
    Beregn
End Sub

Private Sub cmdForrige_Click()
    If mEksempel > 1 Then mEksempel = mEksempel - 1
    VisForhand
End Sub

Private Sub cmdNeste_Click()
    If mEksempel < mKontakter.Count Then mEksempel = mEksempel + 1
    VisForhand
End Sub

Private Sub cmdHoppet_Click()
    If mHoppet.Count = 0 Then Exit Sub
    MsgBox HoppetTekst(40), vbInformation, "Katalogmekker - hoppet over"
End Sub

Private Sub cmdLagVcf_Click()
    Beregn
    If mKontakter.Count = 0 Then
        MsgBox "Ingen kontakter å lagre med disse innstillingene.", vbExclamation, "Katalogmekker"
        Exit Sub
    End If
    LagreInnstillinger

    Dim sti As String
    sti = modKatalogmekker.VelgLagringssti("Kontakter - " & CStr(Me.cboTabell.Value) & ".vcf")
    If sti = "" Then Exit Sub

    Dim landskode As Boolean, kat As String
    landskode = CBool(Me.chkLandskode.Value)
    kat = CStr(Me.txtKategori.Value)

    Dim deler() As String, i As Long, k As Variant
    ReDim deler(1 To mKontakter.Count)
    For i = 1 To mKontakter.Count
        k = mKontakter(i)
        deler(i) = modKatalogmekker.ByggVcard(CStr(k(1)), CStr(k(2)), kat, landskode)
    Next i

    If modKatalogmekker.SkrivUtf8Fil(sti, Join(deler, "")) Then
        Dim melding As String
        melding = "Lagret " & mKontakter.Count & " kontakt(er) i:" & vbCrLf & sti & vbCrLf & vbCrLf
        If mHoppet.Count > 0 Then melding = melding & HoppetTekst(15) & vbCrLf
        melding = melding & "Neste steg: åpne contacts.google.com > Importer, og velg fila. " & _
            "Har noen av kontaktene vært importert før, finner «Slå sammen og rett» duplikatene."
        MsgBox melding, vbInformation, "Katalogmekker"
    End If
End Sub

' Lagring og opprydding skjer i UserForm_QueryClose (gjelder ogsa X-knappen).
Private Sub cmdLukk_Click()
    Unload Me
End Sub
'@

function Get-KatalogmekkerVersion {
    param([string]$Code)
    if ($Code -match 'KATALOGMEKKER_VERSION\s*As\s*String\s*=\s*"([^"]+)"') {
        return $Matches[1]
    }
    return $null
}

$sourceVersion = Get-KatalogmekkerVersion $moduleCodeKatalogmekker

$allComponents = @(
    @{ Name = "modKatalogmekker";     Code = $moduleCodeKatalogmekker;     File = "modKatalogmekker.bas" },
    @{ Name = "clsKatalogmekkerNavn"; Code = $classCodeKatalogmekkerNavn;  File = "clsKatalogmekkerNavn.cls" },
    @{ Name = "frmKatalogmekker";     Code = $formCodeKatalogmekker;       File = "frmKatalogmekker.frm" }
)

$buttonOnActions = @("AapneKatalogmekker")

# ---- 1. Finn målfilen ----

if (-not $Path) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel-filer (*.xlsm)|*.xlsm|Alle filer (*.*)|*.*"
    $dlg.Title = "Velg Excel-fil som skal få Katalogmekker (opprett en tom .xlsm først om nødvendig)"
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

# ---- 2. Finn en åpen Excel-instans som har filen, ellers åpne den selv ----

$excel = $null
$wb = $null

try {
    if ($NyExcelInstans) {
        $running = $null
    } else {
        $running = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    }
    if ($running) {
        foreach ($w in $running.Workbooks) {
            if ($w.FullName -eq $Path) { $wb = $w; $excel = $running; break }
        }
        if ($null -eq $excel) { $excel = $running }
    }
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
    Write-Output "FEIL: Får ikke tilgang til VBA-prosjektet."
    Write-Output "Sjekk at 'Klarer tilgang til VBA-prosjektobjektmodellen' er huket av i Excel sitt Klareringssenter (Filer > Alternativer > Klareringssenter > Innstillinger for klareringssenteret > Makroinnstillinger)."
    exit 1
}

if ($Uninstall) {
    # Fjerner KUN denne makroens egne, eksakt navngitte VBA-komponenter og
    # knapper med denne makroens EGNE OnAction-navn. Innstillingsarket
    # (skjult) og brukerens tabeller røres IKKE - kun kode fjernes.
    Write-Output "Fjerner Katalogmekker fra $($wb.Name) ..."
    foreach ($item in $allComponents) {
        try {
            $eksisterende = $vbproj.VBComponents.Item($item.Name)
            $vbproj.VBComponents.Remove($eksisterende)
            Write-Output "  Fjernet $($item.Name)"
        } catch {}
    }
    try {
        foreach ($ws in $wb.Worksheets) {
            foreach ($btn in @($ws.Buttons())) {
                $onAction = [string]$btn.OnAction
                $kort = ($onAction -split '!')[-1]
                if ($buttonOnActions -contains $kort) { $btn.Delete() }
            }
        }
    } catch {}
    try { $wb.Save() } catch {
        Write-Output "FEIL ved lagring: $($_.Exception.Message)"
        exit 1
    }
    Write-Output "Ferdig - Katalogmekker er fjernet fra $($wb.Name). Innstillingene er beholdt."
    exit 0
}

# ---- 3. Sjekk hva som allerede er installert ----

$existingModule = $null
foreach ($comp in $vbproj.VBComponents) {
    if ($comp.Name -eq "modKatalogmekker") { $existingModule = $comp }
}

$installedVersion = $null
if ($existingModule) {
    if ($existingModule.CodeModule.CountOfLines -gt 0) {
        $existingCode = $existingModule.CodeModule.Lines(1, $existingModule.CodeModule.CountOfLines)
        $installedVersion = Get-KatalogmekkerVersion $existingCode
    } else {
        Write-Output "ADVARSEL: fant modKatalogmekker, men den er tom (0 linjer) - sannsynligvis fra en tidligere avbrutt installasjon. Installerer på nytt."
        $existingModule = $null
    }
}

$isFreshInstall = ($null -eq $existingModule)

if (-not $isFreshInstall -and $installedVersion -eq $sourceVersion -and -not $Force) {
    Write-Output "Katalogmekker er allerede installert og oppdatert (versjon $installedVersion) i $($wb.Name)."
    exit 0
}

if ($isFreshInstall) {
    Write-Output "Installerer Katalogmekker (versjon $sourceVersion) i $($wb.Name) ..."
} else {
    Write-Output "Oppdaterer Katalogmekker: versjon $installedVersion -> $sourceVersion i $($wb.Name) ..."
}

# ---- 4. Bygg komponentene i en midlertidig arbeidsbok, eksporter til en temp-mappe ----

$tempDir = Join-Path $env:TEMP ("Katalogmekker_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir | Out-Null

try {
    $buildWb = $excel.Workbooks.Add()
    $buildProj = $buildWb.VBProject

    # --- frmKatalogmekker (bygges forst, sa MSForms-referansen finnes i
    # byggeprosjektet for klassen som bruker MSForms.CheckBox) ---
    # Alle skrifter/farger settes i VBA (UserForm_Initialize), ikke her -
    # .Font.* fra PowerShell-designeren er upalitelig (g-designer-font-bold-nullref).
    # Venstre kolonne: valg (1-4). Hoyre kolonne: forhandsvisning + knapper.
    $comp = $buildProj.VBComponents.Add(3)
    $comp.Name = "frmKatalogmekker"
    $comp.Properties("Width").Value = 768
    $comp.Properties("Height").Value = 572
    $comp.Properties("Caption").Value = "Katalogmekker"
    $d = $comp.Designer

    # Setter posisjon/storrelse/tekst pa en kontroll som allerede er lagt til.
    # (Returnerer bevisst ingenting - en funksjon som returnerer et COM-objekt
    # kan fa det "pakket ut" av PowerShell.)
    function Set-Ctl($c, $left, $top, $width, $height, $caption) {
        $c.Left = $left; $c.Top = $top; $c.Width = $width; $c.Height = $height
        if ($null -ne $caption) { $c.Caption = $caption }
    }

    # 1. Kilde
    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblKilde", $true)) 16 12 360 16 "1. Hvor ligger kontaktene?"
    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblArk", $true)) 16 35 44 16 "Ark:"
    $cboArk = $d.Controls.Add("Forms.ComboBox.1", "cboArk", $true); Set-Ctl $cboArk 62 32 314 18 $null
    $cboArk.Style = 2
    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblTabell", $true)) 16 59 44 16 "Tabell:"
    $cboTabell = $d.Controls.Add("Forms.ComboBox.1", "cboTabell", $true); Set-Ctl $cboTabell 62 56 314 18 $null
    $cboTabell.Style = 2
    Set-Ctl ($d.Controls.Add("Forms.CheckBox.1", "chkKunSynlige", $true)) 16 80 360 18 "Bare synlige rader (respekter filter)"

    # 2. Telefonnummer
    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblNummerTittel", $true)) 16 108 360 16 "2. Telefonnummer"
    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblNummer", $true)) 16 131 44 16 "Kolonne:"
    $cboNummer = $d.Controls.Add("Forms.ComboBox.1", "cboNummer", $true); Set-Ctl $cboNummer 62 128 314 18 $null
    $cboNummer.Style = 2
    Set-Ctl ($d.Controls.Add("Forms.CheckBox.1", "chkLandskode", $true)) 16 152 360 18 "Skriv numrene med +47 foran"
    $lblNummerHjelp = $d.Controls.Add("Forms.Label.1", "lblNummerHjelp", $true); Set-Ctl $lblNummerHjelp 16 172 360 40 "Numrene renskrives til 8 siffer (norsk standard): 0047, +47 og 47 foran fjernes. Rader med tomt nummer, «...» eller et nummer som ikke er norsk hoppes over. Samme nummer to ganger gir én kontakt."
    $lblNummerHjelp.WordWrap = $true

    # 3. Navn
    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblNavnTittel", $true)) 16 218 360 16 "3. Navn"
    $lblNavnHjelp = $d.Controls.Add("Forms.Label.1", "lblNavnHjelp", $true); Set-Ctl $lblNavnHjelp 16 236 360 26 "Huk av kolonnene som skal utgjøre navnet. Rekkefølgen blir den du huker av i (vist med 1., 2., ...)."
    $lblNavnHjelp.WordWrap = $true
    $fraNavn = $d.Controls.Add("Forms.Frame.1", "fraNavn", $true); Set-Ctl $fraNavn 16 264 360 130 ""
    $fraNavn.ScrollBars = 2
    $fraNavn.KeepScrollBarsVisible = 2
    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblSeparator", $true)) 16 405 160 16 "Skilletegn mellom kolonnene:"
    Set-Ctl ($d.Controls.Add("Forms.TextBox.1", "txtSeparator", $true)) 180 402 80 18 $null

    # 4. Kategori
    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblKategoriTittel", $true)) 16 432 360 16 "4. Kategori (etikett i Google Kontakter)"
    Set-Ctl ($d.Controls.Add("Forms.TextBox.1", "txtKategori", $true)) 16 452 360 18 $null
    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblKategoriHjelp", $true)) 16 474 360 14 "Felles for alle kontaktene. Flere kategorier skilles med komma."

    # 5. Forhandsvisning (hoyre kolonne). Barn av rammen legges til via
    # rammens egen Controls.Add (g-frame-children-com).
    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblForhandTittel", $true)) 392 12 352 16 "5. Forhåndsvisning"
    $fraForhand = $d.Controls.Add("Forms.Frame.1", "fraForhand", $true); Set-Ctl $fraForhand 392 32 352 150 ""
    Set-Ctl ($fraForhand.Controls.Add("Forms.Label.1", "lblPrevNavn", $true)) 12 10 326 20 ""
    Set-Ctl ($fraForhand.Controls.Add("Forms.Label.1", "lblPrevNummer", $true)) 12 36 326 16 ""
    $lblPrevKategori = $fraForhand.Controls.Add("Forms.Label.1", "lblPrevKategori", $true); Set-Ctl $lblPrevKategori 12 56 326 30 ""
    $lblPrevKategori.WordWrap = $true
    Set-Ctl ($fraForhand.Controls.Add("Forms.Label.1", "lblEksempel", $true)) 12 94 326 14 ""
    Set-Ctl ($fraForhand.Controls.Add("Forms.CommandButton.1", "cmdForrige", $true)) 12 114 90 24 "< Forrige"
    Set-Ctl ($fraForhand.Controls.Add("Forms.CommandButton.1", "cmdNeste", $true)) 108 114 90 24 "Neste >"

    Set-Ctl ($d.Controls.Add("Forms.Label.1", "lblVcardTittel", $true)) 392 192 352 14 "Slik blir kontakten i fila:"
    $txtVcard = $d.Controls.Add("Forms.TextBox.1", "txtVcard", $true); Set-Ctl $txtVcard 392 208 352 130 $null
    $txtVcard.MultiLine = $true
    $txtVcard.ScrollBars = 3
    $txtVcard.Locked = $true
    $txtVcard.WordWrap = $false

    $lblStatus = $d.Controls.Add("Forms.Label.1", "lblStatus", $true); Set-Ctl $lblStatus 392 348 352 80 ""
    $lblStatus.WordWrap = $true

    Set-Ctl ($d.Controls.Add("Forms.CommandButton.1", "cmdHoppet", $true)) 392 496 140 30 "Vis hoppet over"
    Set-Ctl ($d.Controls.Add("Forms.CommandButton.1", "cmdLagVcf", $true)) 540 496 120 30 "Lag .vcf ..."
    Set-Ctl ($d.Controls.Add("Forms.CommandButton.1", "cmdLukk", $true)) 668 496 76 30 "Lukk"

    # Kort pause mellom designerbygging og AddFromString (g-designer-font-bold-nullref).
    Start-Sleep -Milliseconds 400
    $comp.CodeModule.AddFromString($formCodeKatalogmekker)
    $comp.Export((Join-Path $tempDir "frmKatalogmekker.frm"))

    # --- clsKatalogmekkerNavn ---
    $cls = $buildProj.VBComponents.Add(2)   # vbext_ct_ClassModule
    $cls.Name = "clsKatalogmekkerNavn"
    $cls.CodeModule.AddFromString($classCodeKatalogmekkerNavn)
    $cls.Export((Join-Path $tempDir "clsKatalogmekkerNavn.cls"))

    # --- modKatalogmekker ---
    $mod = $buildProj.VBComponents.Add(1)   # vbext_ct_StdModule
    $mod.Name = "modKatalogmekker"
    $mod.CodeModule.AddFromString($moduleCodeKatalogmekker)
    $mod.Export((Join-Path $tempDir "modKatalogmekker.bas"))

    $buildWb.Close($false)

    # ---- 5. Bytt ut / legg til komponentene i mål-arbeidsboken ----

    foreach ($item in $allComponents) {
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

    # Ingen automatisk ark-knapp (samme faste regel som de andre makroene -
    # tilgang skjer via Makromeny). Innstillingsarket (skjult) lager seg
    # selv forste gang noe lagres (se HentOppsettArk i VBA).

    Write-Output ""
    Write-Output "Ferdig. Husk å lagre filen selv (Ctrl+S) når du er klar."
    if ($isFreshInstall) {
        Write-Output ""
        Write-Output "Neste steg:"
        Write-Output "  1. Kjør ""AapneKatalogmekker"" (f.eks. via Makromeny - ingen knapp legges til i arket automatisk)."
        Write-Output "  2. Velg ark, tabell, nummerkolonne og kolonnene som skal utgjøre navnet."
        Write-Output "  3. Sjekk forhåndsvisningen, og trykk ""Lag .vcf ..."" for å lagre kontaktfila."
        Write-Output "  4. Importer fila på contacts.google.com (Importer)."
    }
} finally {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
