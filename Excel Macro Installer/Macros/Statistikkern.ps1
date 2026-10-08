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
# Statistikkern - alt-i-en installer/oppdaterer.
#
# Forer statistikk over unike verdier i valgte kolonner i en eller flere
# tabeller. Hver unik verdi i en "hovedkolonne" vises i statistikkmenyen
# (SM) som en fargelagt flis, med et tall for hvor mange ganger verdien
# forekommer. Fargen hentes fra cellens FAKTISKE viste fyllfarge
# (DisplayFormat.Interior - fanger opp betinget formatering, som
# prioriterer over manuell formatering; se g-vba-const-midt-i-modul og
# lignende gotchas i kunnskapsbasen for beslektede COM-detaljer).
#
# Valgfri "underkolonne" per hovedkolonne bryter tallene ned videre innad
# i hver hovedverdi (uten egen farge - kun hovedkolonnen har fargede
# fliser). Oppsettet (hvilke tabeller/kolonner) velges i et eget
# Oppsett-vindu og lagres i et skjult ark i arbeidsboken.
#
# v1.0.0 (2026-09-29, Nils): forste utkast. Bygget etter direkte spesifikasjon
# fra Hakon. UI-valg (avklart med Hakon for bygging startet): SM er et
# flytende vindu (UserForm), samme monster som de andre makroene - IKKE et
# eget ark med Shapes. Flisene er enkle fargede rektangler (Frame-kontroller
# med BackColor), ingen ekte avrundede hjorner (bevisst forenklet - Hakon
# ba eksplisitt om at dette skal vaere enkelt og fungere i forste vindu som
# dukker opp, samme som alle andre makroer). Nar flere tabelloppsett er
# konfigurert: alt vises i EN lang, rullende visning (ikke faner/velger).
# v1.0.1: rettet manglende UTF-8 BOM (korrumperte et spesialtegn).
# v1.3.0 (2026-10-02, Nils): tredje visning "Horisontal" - smale, like brede
# fliser (bredde malt fra den lengste teksten), alle fliser fra en tabell i
# en rad og vinduet tilpasses bredden (begrenset til Excel-vinduet, da med
# vannrett rullefelt); flere tabeller legges under hverandre. Knappen
# veksler Liste -> Rute -> Horisontal og viser gjeldende visning. Data
# beregnes na en gang (ByggData) og tegnes etterpa. Rullingen nullstilles
# ved hver ombygging (rullefeil ved bytte av visning).
# v1.2.0 (2026-09-30, Nils): ingen tittel/undertekst i SM (mer plass til
# menyen); Oppsett-knappen ligger under rullemenyen, med en ny knapp ved
# siden av som veksler mellom liste- og rutevisning (fliser side om side,
# antall kolonner etter tilgjengelig bredde). Valgt visning huskes pa rad 1
# i oppsett-arket. Layout i SM regnes na fra InsideWidth/InsideHeight.
# v1.1.0 (2026-09-29, Nils): Hakons tilbakemelding etter forste test -
# tomme underkolonne-celler telles na som en egen "Tomme"-gruppe (kursiv,
# se TOMME_VERDI); hovedkolonne-flisene sorteres alfabetisk A-AA i stedet
# for etter antall (SorterAlfabetisk); "Oppdater"-knappen er fjernet -
# statistikken regnes na alltid pa nytt automatisk ved apning. Visuell
# runde: farget primaerknapp (StatStilKnapp, samme teknikk som
# Kontaktsentralen sin KtsStilKnapp), kontrastberegnet tekstfarge inni
# flisene (StatTekstFargeMot), tynn kant pa flisene, lysere bakgrunn,
# undertekst under hver vindutittel. Layoutfiks: "Ny tabell"-vinduet hadde
# et reelt overflow-avvik (kolonnerad-innholdet stakk utenfor den rullbare
# rammen nar rullefeltet var synlig) - bredder regnet om med eksplisitt
# margin til rullefelt pa alle tre vinduer, se HTML-notatet for detaljene.
#
# VBA-komponenter (alle prefikset "Statistikkern" for a unnga kollisjon med
# andre makroers komponentnavn pa samme ark):
#   modStatistikkern         - registerhjelpere, tabelldeteksjon, statistikk-
#                               beregning, inngangspunkt (AapneStatistikkern)
#   frmStatistikkern          - hovedvinduet (SM) som apnes av AapneStatistikkern
#   frmStatistikkernOppsett   - liste over konfigurerte tabeller/kolonner
#   frmStatistikkernNyTabell  - velg tabell + huk av hovedkolonner/underkolonner
# ============================================================================

$moduleCodeStatistikkern = @'
Attribute VB_Name = "modStatistikkern"
Option Explicit

' Statistikkern - forer statistikk over unike verdier i valgte kolonner i en
' eller flere tabeller. Hver unik verdi i en hovedkolonne vises som en flis
' med farge hentet fra cellens FAKTISKE viste fyllfarge (DisplayFormat -
' fanger opp betinget formatering, som prioriterer over manuell
' formatering). En valgfri underkolonne bryter tallene ned videre innad i
' hver hovedverdi (ingen egen farge for underverdier - kun hovedkolonnen
' har fargede fliser, jf. Hakons opprinnelige spesifikasjon).
'
' Oppsettet (hvilke tabeller/kolonner) lagres i et skjult ark
' (OPPSETT_SHEET), en rad per hovedkolonne (samme tabell kan ha flere
' hovedkolonner, hver sin rad - visningsrekkefolgen i SM er radrekkefolgen
' i dette arket).

Public Const STATISTIKKERN_VERSION As String = "1.3.0"
Public Const OPPSETT_SHEET As String = "StatistikkernOppsett_Hidden"
' Brukt for tomme underkolonne-celler - se BeregnUnderkolonneStatistikk.
' Vises i kursiv av frmStatistikkern (matchet pa noyaktig denne strengen).
Public Const TOMME_VERDI As String = "Tomme"

Public Sub AapneStatistikkern()
    frmStatistikkern.Show
End Sub

' ---- Oppsett-registeret (skjult ark) ----

Public Function GetOppsettSheet() As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(OPPSETT_SHEET)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sh.Name = OPPSETT_SHEET
        sh.Range("A1:D1").Value = Array("TabellArk", "TabellNavn", "HovedKolonne", "UnderKolonne")
    End If
    sh.Visible = xlSheetVeryHidden
    Set GetOppsettSheet = sh
End Function

' Valgt visningsmodus (liste/rute) huskes pa RAD 1 i oppsett-arket (F1 =
' "Visning", G1 = "Rute" eller "Liste") - rad 1 slettes aldri av
' LagreHovedkolonneValg/FjernOppsettRader (de rorer kun rad 2 og nedover).
' Visning: 0 = Liste, 1 = Rute, 2 = Horisontal.
Public Function LesVisning() As Long
    Select Case CStr(GetOppsettSheet().Range("G1").Value)
        Case "Rute": LesVisning = 1
        Case "Horisontal": LesVisning = 2
        Case Else: LesVisning = 0
    End Select
End Function

Public Sub LagreVisning(ByVal Visning As Long)
    Dim sh As Worksheet
    Set sh = GetOppsettSheet()
    sh.Range("F1").Value = "Visning"
    Select Case Visning
        Case 1: sh.Range("G1").Value = "Rute"
        Case 2: sh.Range("G1").Value = "Horisontal"
        Case Else: sh.Range("G1").Value = "Liste"
    End Select
End Sub

' En Dictionary per rad (Rad/TabellArk/TabellNavn/HovedKolonne/UnderKolonne).
' Rekkefolgen i Collection-en er samme som radrekkefolgen i arket - dette ER
' visningsrekkefolgen i SM, ingen egen sorteringskolonne trengs.
Public Function LesOppsett() As Collection
    Dim resultat As New Collection
    Dim sh As Worksheet
    Set sh = GetOppsettSheet()
    Dim lastRow As Long, r As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(sh.Cells(r, 1).Value) <> "" Then
            Dim d As Object
            Set d = CreateObject("Scripting.Dictionary")
            d("Rad") = r
            d("TabellArk") = CStr(sh.Cells(r, 1).Value)
            d("TabellNavn") = CStr(sh.Cells(r, 2).Value)
            d("HovedKolonne") = CStr(sh.Cells(r, 3).Value)
            d("UnderKolonne") = CStr(sh.Cells(r, 4).Value)
            resultat.Add d
        End If
    Next r
    Set LesOppsett = resultat
End Function

' Synkroniserer hovedkolonne-valgene for EN tabell mot registeret:
' ValgteKolonner er en Dictionary HovedKolonneNavn -> UnderKolonneNavn
' ("" = ingen underkolonne). Rader for denne tabellen som ikke lenger er
' valgt fjernes; valgte rader legges til pa nytt (enklere og like trygt som
' a oppdatere i place, siden hele settet for tabellen sendes inn samlet).
Public Sub LagreHovedkolonneValg(ByVal TabellArk As String, ByVal TabellNavn As String, ByVal ValgteKolonner As Object)
    Dim sh As Worksheet
    Set sh = GetOppsettSheet()

    Dim lastRow As Long, r As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    For r = lastRow To 2 Step -1
        If CStr(sh.Cells(r, 1).Value) = TabellArk And CStr(sh.Cells(r, 2).Value) = TabellNavn Then
            sh.Rows(r).Delete
        End If
    Next r

    Dim nesteRad As Long
    nesteRad = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row + 1
    If nesteRad < 2 Then nesteRad = 2

    Dim k As Variant
    For Each k In ValgteKolonner.Keys
        sh.Cells(nesteRad, 1).Value = TabellArk
        sh.Cells(nesteRad, 2).Value = TabellNavn
        sh.Cells(nesteRad, 3).Value = CStr(k)
        sh.Cells(nesteRad, 4).Value = CStr(ValgteKolonner(k))
        nesteRad = nesteRad + 1
    Next k
End Sub

' Rader identifisert ved sine radnumre i selve arket (som de kom inn fra
' LesOppsett - "Rad"-nokkelen). Sorteres synkende for skanning trygt kan
' slette bakfra uten at tidligere radnumre forskyves underveis.
Public Sub FjernOppsettRader(ByVal Rader As Collection)
    Dim sh As Worksheet
    Set sh = GetOppsettSheet()
    Dim arr() As Long
    Dim i As Long, j As Long
    ReDim arr(1 To Rader.Count)
    For i = 1 To Rader.Count
        arr(i) = CLng(Rader(i))
    Next i
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) > arr(i) Then
                Dim tmp As Long
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
    For i = 1 To UBound(arr)
        sh.Rows(arr(i)).Delete
    Next i
End Sub

' ---- Tabelldeteksjon ----

' Alle ListObjects i arbeidsboken, pa tvers av alle SYNLIGE ark (utelukker
' bl.a. dette registerets eget skjulte ark og andre makroers skjulte
' registre). Hver oppforing er en Dictionary ArkNavn/TabellNavn/
' AntallKolonner - selve ListObject-et hentes pa nytt ved behov via
' HentTabell, i stedet for a holde en direkte objektreferanse over tid som
' kan bli ugyldig hvis arket slettes/endres mens SM star apen.
Public Function AlleListObjekter() As Collection
    Dim resultat As New Collection
    Dim ws As Worksheet
    Dim tbl As ListObject
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            For Each tbl In ws.ListObjects
                Dim d As Object
                Set d = CreateObject("Scripting.Dictionary")
                d("ArkNavn") = ws.Name
                d("TabellNavn") = tbl.Name
                d("AntallKolonner") = tbl.ListColumns.Count
                resultat.Add d
            Next tbl
        End If
    Next ws
    Set AlleListObjekter = resultat
End Function

Public Function HentTabell(ByVal ArkNavn As String, ByVal TabellNavn As String) As ListObject
    On Error Resume Next
    Set HentTabell = ThisWorkbook.Worksheets(ArkNavn).ListObjects(TabellNavn)
    On Error GoTo 0
End Function

' ---- Statistikkberegning ----

' VerdiRekkefolge: Collection med String-verdier, i den rekkefolgen de forst
' ble sett. AntallPerVerdi: Dictionary verdi -> antall (Long). FargePerVerdi:
' Dictionary verdi -> RGB-farge (Long) - en verdi mangler i denne Dictionary-
' en hvis INGEN forekomst av den hadde noen fyllfarge ("dersom verdien ikke
' har en farge, skal ingen farge vises").
Public Sub BeregnHovedkolonneStatistikk(ByVal tbl As ListObject, ByVal HovedKolonneNavn As String, _
        ByRef VerdiRekkefolge As Collection, ByRef AntallPerVerdi As Object, ByRef FargePerVerdi As Object)

    Set VerdiRekkefolge = New Collection
    Set AntallPerVerdi = CreateObject("Scripting.Dictionary")
    Set FargePerVerdi = CreateObject("Scripting.Dictionary")

    If tbl.ListRows.Count = 0 Then Exit Sub

    Dim lc As ListColumn
    On Error Resume Next
    Set lc = tbl.ListColumns(HovedKolonneNavn)
    On Error GoTo 0
    If lc Is Nothing Then Exit Sub

    Dim dataRange As Range
    Set dataRange = lc.DataBodyRange

    Dim fargeTeller As Object ' verdi -> Dictionary(fargeLong -> antall)
    Set fargeTeller = CreateObject("Scripting.Dictionary")

    ' Massehenting av verdiene i ETT COM-kall (Value2), i stedet for ett kall
    ' per celle - kun selve fargeoppslaget (som ma skje celle for celle, det
    ' finnes ingen tilsvarende massehenting for DisplayFormat) gjores i
    ' loopen under. For store tabeller er DET steget den reelle kostnaden.
    Dim vals As Variant
    If dataRange.Rows.Count = 1 Then
        ReDim vals(1 To 1, 1 To 1)
        vals(1, 1) = dataRange.Value2
    Else
        vals = dataRange.Value2
    End If

    Dim r As Long
    For r = 1 To UBound(vals, 1)
        Dim verdi As String
        verdi = Trim$(CStr(vals(r, 1)))
        If verdi <> "" Then
            If Not AntallPerVerdi.Exists(verdi) Then
                AntallPerVerdi.Add verdi, 0
                VerdiRekkefolge.Add verdi
                Set fargeTeller(verdi) = CreateObject("Scripting.Dictionary")
            End If
            AntallPerVerdi(verdi) = AntallPerVerdi(verdi) + 1

            Dim cell As Range
            Set cell = dataRange.Cells(r, 1)
            If cell.DisplayFormat.Interior.Pattern <> xlPatternNone Then
                Dim fargeLong As Long
                fargeLong = cell.DisplayFormat.Interior.Color
                Dim fd As Object
                Set fd = fargeTeller(verdi)
                If fd.Exists(fargeLong) Then
                    fd(fargeLong) = fd(fargeLong) + 1
                Else
                    fd.Add fargeLong, 1
                End If
            End If
        End If
    Next r

    ' Modus-farge (mest brukte forekomst) per verdi.
    Dim v As Variant
    For Each v In VerdiRekkefolge
        Dim fd2 As Object
        Set fd2 = fargeTeller(CStr(v))
        If fd2.Count > 0 Then
            Dim besteFarge As Long, besteAntall As Long, forstGang As Boolean
            forstGang = True
            Dim fk As Variant
            For Each fk In fd2.Keys
                If forstGang Or fd2(fk) > besteAntall Then
                    besteAntall = fd2(fk)
                    besteFarge = CLng(fk)
                    forstGang = False
                End If
            Next fk
            FargePerVerdi.Add CStr(v), besteFarge
        End If
    Next v
End Sub

' Samme prinsipp som over, men avgrenset til rader der HovedKolonneNavn =
' HovedVerdi, og uten fargeoppslag (underverdier vises uten egen farge, jf.
' Hakons spesifikasjon - kun hovedkolonnens fliser er fargede). Tomme
' underkolonne-celler telles OGSA, som en egen "Tomme"-gruppe (Hakons
' eksplisitte onske - motsatt av hovedkolonnen, der en tom celle fortsatt
' ikke gir noen egen flis).
Public Sub BeregnUnderkolonneStatistikk(ByVal tbl As ListObject, ByVal HovedKolonneNavn As String, ByVal HovedVerdi As String, _
        ByVal UnderKolonneNavn As String, ByRef VerdiRekkefolge As Collection, ByRef AntallPerVerdi As Object)

    Set VerdiRekkefolge = New Collection
    Set AntallPerVerdi = CreateObject("Scripting.Dictionary")

    If tbl.ListRows.Count = 0 Then Exit Sub

    Dim lcHoved As ListColumn, lcUnder As ListColumn
    On Error Resume Next
    Set lcHoved = tbl.ListColumns(HovedKolonneNavn)
    Set lcUnder = tbl.ListColumns(UnderKolonneNavn)
    On Error GoTo 0
    If lcHoved Is Nothing Or lcUnder Is Nothing Then Exit Sub

    Dim nRows As Long
    nRows = lcHoved.DataBodyRange.Rows.Count

    Dim hovedVals As Variant, underVals As Variant
    If nRows = 1 Then
        ReDim hovedVals(1 To 1, 1 To 1): hovedVals(1, 1) = lcHoved.DataBodyRange.Value2
        ReDim underVals(1 To 1, 1 To 1): underVals(1, 1) = lcUnder.DataBodyRange.Value2
    Else
        hovedVals = lcHoved.DataBodyRange.Value2
        underVals = lcUnder.DataBodyRange.Value2
    End If

    Dim r As Long
    For r = 1 To nRows
        If Trim$(CStr(hovedVals(r, 1))) = HovedVerdi Then
            Dim uv As String
            uv = Trim$(CStr(underVals(r, 1)))
            If uv = "" Then uv = TOMME_VERDI
            If Not AntallPerVerdi.Exists(uv) Then
                AntallPerVerdi.Add uv, 0
                VerdiRekkefolge.Add uv
            End If
            AntallPerVerdi(uv) = AntallPerVerdi(uv) + 1
        End If
    Next r
End Sub

' Sorterer en Collection av verdinavn synkende etter antall (brukt av bade
' SM og underkolonne-visningen, slik at storste gruppe vises forst) - enkel
' boblesortering; antall UNIKE verdier per kolonne forventes lavt (typisk
' kategoriverdier, ikke radantall), sa dette er ikke en ytelsesbekymring.
Public Function SorterEtterAntall(ByVal VerdiRekkefolge As Collection, ByVal AntallPerVerdi As Object) As Collection
    Dim n As Long
    n = VerdiRekkefolge.Count
    If n = 0 Then
        Set SorterEtterAntall = New Collection
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(1 To n)
    Dim i As Long, j As Long
    For i = 1 To n
        arr(i) = CStr(VerdiRekkefolge(i))
    Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If AntallPerVerdi(arr(j)) > AntallPerVerdi(arr(i)) Then
                Dim tmp As String
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    Dim resultat As New Collection
    For i = 1 To n
        resultat.Add arr(i)
    Next i
    Set SorterEtterAntall = resultat
End Function

' Sorterer en Collection av verdinavn alfabetisk A-AA (StrComp med
' vbTextCompare respekterer norsk lokalsortering - AE/O/AA havner riktig
' etter Z). Brukt for HOVEDkolonne-flisene (Hakons eksplisitte onske);
' underkolonne-nedbrytningen bruker fortsatt SorterEtterAntall.
Public Function SorterAlfabetisk(ByVal VerdiRekkefolge As Collection) As Collection
    Dim n As Long
    n = VerdiRekkefolge.Count
    If n = 0 Then
        Set SorterAlfabetisk = New Collection
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(1 To n)
    Dim i As Long, j As Long
    For i = 1 To n
        arr(i) = CStr(VerdiRekkefolge(i))
    Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If StrComp(arr(j), arr(i), vbTextCompare) < 0 Then
                Dim tmp As String
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    Dim resultat As New Collection
    For i = 1 To n
        resultat.Add arr(i)
    Next i
    Set SorterAlfabetisk = resultat
End Function

' ---- Utseende: samme "farget knapp"/kontrast-teknikk som Kontaktsentralen
' (KtsStilKnapp/KtsTekstFargeMot) - egen kopi her (Stat-prefiks) for a
' unnga tverr-makro-avhengighet, se excel-macro-installer-pattern i minnet
' for hvorfor hver makro skal vaere selvstendig. ----

' Farget knapp (hvit fet tekst). Krever at knappen har Style=1 (graphical)
' satt i designeren - BackColor vises ellers ikke pa en CommandButton.
' Fargenavn: rod (Statistikkerns egen aksentfarge - primaerhandling/slett),
' gronn (lagre/positiv), gra (noytral/lukk/avbryt).
Public Sub StatStilKnapp(ByVal Knapp As Object, ByVal Fargenavn As String)
    Dim farge As Long
    Select Case Fargenavn
        Case "rod": farge = RGB(220, 38, 38)
        Case "gronn": farge = RGB(46, 139, 87)
        Case "gra": farge = RGB(138, 138, 132)
        Case "blagra": farge = RGB(91, 107, 138)
        Case Else: farge = RGB(138, 138, 132)
    End Select
    On Error Resume Next
    Knapp.Style = 1 ' fmButtonStyleGraphical - kreves for at BackColor skal vises
    Knapp.BackColor = farge
    Knapp.ForeColor = RGB(255, 255, 255)
    Knapp.Font.Bold = True
End Sub

' Lesbar tekstfarge (mørkt eller hvitt) mot en gitt bakgrunnsfarge - brukt
' for tekst inni de fargede flisene i SM, slik at en mørk flisfarge ikke
' gir uleselig sort tekst.
Public Function StatTekstFargeMot(ByVal Bakgrunn As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = Bakgrunn Mod 256: g = (Bakgrunn \ 256) Mod 256: b = (Bakgrunn \ 65536) Mod 256
    If (299 * r + 587 * g + 114 * b) / 1000 < 140 Then
        StatTekstFargeMot = RGB(255, 255, 255)
    Else
        StatTekstFargeMot = RGB(30, 30, 30)
    End If
End Function
'@

$formCodeStatistikkern = @'
Attribute VB_Name = "frmStatistikkern"
Option Explicit

Private Const TILE_HOYDE_BASE As Long = 32       ' hoyde for selve hovedverdi-linjen
Private Const UNDERLINJE_HOYDE As Long = 16      ' hoyde per underverdi-linje
Private Const TILE_MELLOMROM As Long = 10        ' luft mellom fliser
Private Const SEKSJON_MELLOMROM As Long = 20     ' luft mellom seksjoner (tabell/hovedkolonne)
Private Const TILE_HOYRE_MARG As Long = 40       ' fraInnhold.Width minus dette = tilgjengelig bredde for fliser (rullefelt + luft)
Private Const RUTE_MIN_BREDDE As Long = 170      ' minste flisbredde i rutevisning - avgjor antall kolonner
Private Const VIS_LISTE As Long = 0              ' en kolonne, full bredde
Private Const VIS_RUTE As Long = 1               ' fliser side om side i rader, pakket etter tilgjengelig bredde
Private Const VIS_HORISONTAL As Long = 2         ' smale, like brede fliser; alle fliser fra en tabell i en rad, vinduet tilpasses
Private mVisning As Long
Private mBasisBredde As Single                   ' vinduets bredde slik den er satt i designeren (brukt i liste/rute)
Private mFlisNr As Long                          ' teller for unike kontrollnavn

Private Sub UserForm_Initialize()
    Me.Caption = "Statistikkern"
    Me.BackColor = RGB(246, 247, 249)
    Me.fraInnhold.BackColor = RGB(246, 247, 249)
    modStatistikkern.StatStilKnapp Me.cmdOppsett, "rod"
    modStatistikkern.StatStilKnapp Me.cmdVisning, "blagra"
    Me.StartUpPosition = 0
    mBasisBredde = Me.Width

    mVisning = modStatistikkern.LesVisning()
    OppdaterVisningKnapp
    ' Ingen "Oppdater"-knapp (Hakons onske) - statistikken regnes alltid pa
    ' nytt automatisk hver gang vinduet apnes (her) eller nar man kommer
    ' tilbake fra Oppsett-vinduet (se cmdOppsett_Click).
    GjenoppbyggInnhold
End Sub

' Layout regnes ut fra vinduets faktiske indre storrelse (ikke faste tall),
' slik at ingenting kan stikke utenfor vinduet: rullemenyen fyller det meste,
' knappene ligger under den. Kalles etter hver endring av vinduets bredde.
Private Sub PlasserKontroller()
    Const MARG As Single = 12
    Const KNAPPH As Single = 30
    Me.fraInnhold.Left = MARG
    Me.fraInnhold.Top = MARG
    Me.fraInnhold.Width = Me.InsideWidth - 2 * MARG
    Me.fraInnhold.Height = Me.InsideHeight - 3 * MARG - KNAPPH
    Me.cmdOppsett.Left = MARG
    Me.cmdOppsett.Top = Me.fraInnhold.Top + Me.fraInnhold.Height + MARG
    Me.cmdOppsett.Width = 130
    Me.cmdOppsett.Height = KNAPPH
    Me.cmdVisning.Left = MARG + 130 + 8
    Me.cmdVisning.Top = Me.cmdOppsett.Top
    Me.cmdVisning.Width = 170
    Me.cmdVisning.Height = KNAPPH
End Sub

Private Sub SentrerVindu()
    On Error Resume Next
    Dim t As Single, l As Single
    t = Application.Top + (Application.Height - Me.Height) / 2
    l = Application.Left + (Application.Width - Me.Width) / 2
    If t < Application.Top Then t = Application.Top
    If l < Application.Left Then l = Application.Left
    Me.Top = t
    Me.Left = l
End Sub

Private Sub OppdaterVisningKnapp()
    Select Case mVisning
        Case VIS_RUTE: Me.cmdVisning.Caption = "Visning: Rute"
        Case VIS_HORISONTAL: Me.cmdVisning.Caption = "Visning: Horisontal"
        Case Else: Me.cmdVisning.Caption = "Visning: Liste"
    End Select
End Sub

' Veksler Liste -> Rute -> Horisontal -> Liste.
Private Sub cmdVisning_Click()
    mVisning = (mVisning + 1) Mod 3
    modStatistikkern.LagreVisning mVisning
    OppdaterVisningKnapp
    GjenoppbyggInnhold
End Sub

Private Sub cmdOppsett_Click()
    frmStatistikkernOppsett.Show
    GjenoppbyggInnhold
End Sub

Private Sub RyddInnhold()
    Dim i As Long
    For i = Me.fraInnhold.Controls.Count - 1 To 0 Step -1
        Me.fraInnhold.Controls.Remove Me.fraInnhold.Controls(i).Name
    Next i
End Sub

' Leser oppsettet og beregner all statistikk EN gang. Resultat: en Collection
' med en Dictionary per tabell-/hovedkolonne-seksjon (Tittel, Fliser).
' Hver flis er en Dictionary (Verdi, Antall, Farge [-1 = ingen], Under =
' Collection av Array(tekst, antall)). Selve tegningen skjer i
' GjenoppbyggInnhold, slik at Horisontal-visningen kan male alle fliser
' for den vet hvor brede flisene ma vaere (maling kreves for vindusbredden).
Private Function ByggData() As Collection
    Dim res As Collection
    Set res = New Collection
    Dim oppsett As Collection
    Set oppsett = modStatistikkern.LesOppsett()

    Dim i As Long
    For i = 1 To oppsett.Count
        Dim rad As Object
        Set rad = oppsett(i)

        Dim tbl As ListObject
        Set tbl = modStatistikkern.HentTabell(CStr(rad("TabellArk")), CStr(rad("TabellNavn")))
        ' Tabellen er borte/omdopt siden sist - hopp over seksjonen i stedet
        ' for a krasje. Blir liggende i oppsettet til brukeren rydder bort.
        If Not tbl Is Nothing Then
            Dim sek As Object
            Set sek = CreateObject("Scripting.Dictionary")
            sek("Tittel") = CStr(rad("TabellNavn")) & "  ·  " & CStr(rad("HovedKolonne"))

            Dim vr As Collection, ap As Object, fp As Object
            modStatistikkern.BeregnHovedkolonneStatistikk tbl, CStr(rad("HovedKolonne")), vr, ap, fp
            ' Hovedkolonnens fliser: alfabetisk A-AA (Hakons eksplisitte onske).
            Dim sortert As Collection
            Set sortert = modStatistikkern.SorterAlfabetisk(vr)

            Dim fliser As Collection
            Set fliser = New Collection
            Dim underKol As String
            underKol = CStr(rad("UnderKolonne"))

            Dim vv As Variant
            For Each vv In sortert
                Dim f As Object
                Set f = CreateObject("Scripting.Dictionary")
                f("Verdi") = CStr(vv)
                f("Antall") = CLng(ap(CStr(vv)))
                If fp.Exists(CStr(vv)) Then f("Farge") = CLng(fp(CStr(vv))) Else f("Farge") = -1

                Dim under As Collection
                Set under = New Collection
                If underKol <> "" Then
                    ' Underkolonnen sorteres fortsatt etter antall (storste
                    ' gruppe forst) - kun hovedkolonnen skal vaere alfabetisk.
                    Dim ur As Collection, ua As Object, us As Collection
                    modStatistikkern.BeregnUnderkolonneStatistikk tbl, CStr(rad("HovedKolonne")), CStr(vv), underKol, ur, ua
                    Set us = modStatistikkern.SorterEtterAntall(ur, ua)
                    Dim uv As Variant
                    For Each uv In us
                        under.Add Array(CStr(uv), CLng(ua(CStr(uv))))
                    Next uv
                End If
                Set f("Under") = under
                fliser.Add f
            Next vv
            Set sek("Fliser") = fliser
            res.Add sek
        End If
    Next i
    Set ByggData = res
End Function

Private Function FlisHoyde(ByVal f As Object) As Long
    FlisHoyde = TILE_HOYDE_BASE + (f("Under").Count * UNDERLINJE_HOYDE) + 8
End Function

' Maler tekstbredden (i punkter) med en midlertidig AutoSize-label.
Private Function MaalBredde(ByVal tekst As String, ByVal fet As Boolean, ByVal skriftstr As Single) As Single
    Dim lm As MSForms.Label
    Set lm = Me.fraInnhold.Controls.Add("Forms.Label.1", "lblMaal", True)
    lm.WordWrap = False
    lm.AutoSize = True
    lm.Font.Bold = fet
    If skriftstr > 0 Then lm.Font.Size = skriftstr
    lm.Caption = tekst
    MaalBredde = lm.Width
    Me.fraInnhold.Controls.Remove "lblMaal"
    If MaalBredde < 5 Then MaalBredde = Len(tekst) * 6
End Function

' Horisontal-visning: en felles flisbredde som er akkurat nok til den
' bredeste teksten blant ALLE fliser (alle fliser skal vaere like brede).
' Maler kun den lengste hovedteksten og den lengste underteksten (etter
' antall tegn) - to malinger i stedet for en per flis.
Private Function HorisontalFlisBredde(ByVal data As Collection) As Single
    Dim lengsteHoved As String, lengsteUnder As String
    Dim sek As Object, f As Object, u As Variant
    Dim s As String
    For Each sek In data
        For Each f In sek("Fliser")
            s = CStr(f("Verdi")) & "   " & CStr(f("Antall"))
            If Len(s) > Len(lengsteHoved) Then lengsteHoved = s
            For Each u In f("Under")
                s = CStr(u(0)) & "   " & CStr(u(1))
                If Len(s) > Len(lengsteUnder) Then lengsteUnder = s
            Next u
        Next f
    Next sek
    Dim bHoved As Single, bUnder As Single
    If Len(lengsteHoved) > 0 Then bHoved = MaalBredde(lengsteHoved, True, 10) + 10 + 16
    If Len(lengsteUnder) > 0 Then bUnder = MaalBredde(lengsteUnder, False, 0) + 24 + 16
    HorisontalFlisBredde = bHoved
    If bUnder > HorisontalFlisBredde Then HorisontalFlisBredde = bUnder
    If HorisontalFlisBredde < 110 Then HorisontalFlisBredde = 110
End Function

' Setter vinduets bredde slik at en rad med radBredde punkter fliser far
' plass (begrenset til Excel-vinduet; ved begrensning far man vannrett
' rullefelt i stedet).
Private Sub TilpassBredde(ByVal radBredde As Single)
    Dim ny As Single
    ny = Me.Width - Me.InsideWidth + radBredde + TILE_HOYRE_MARG + 24
    On Error Resume Next
    Dim maks As Single
    maks = Application.Width - 20
    If maks > 300 And ny > maks Then ny = maks
    On Error GoTo 0
    If ny < 400 Then ny = 400
    Me.Width = ny
    PlasserKontroller
End Sub

Private Sub TegnFlis(ByVal x As Single, ByVal y As Single, ByVal b As Single, ByVal h As Single, ByVal f As Object)
    Dim tileFarge As Long, tekstFarge As Long
    If CLng(f("Farge")) >= 0 Then
        tileFarge = CLng(f("Farge"))
    Else
        tileFarge = RGB(238, 239, 242)
    End If
    tekstFarge = modStatistikkern.StatTekstFargeMot(tileFarge)

    mFlisNr = mFlisNr + 1
    Dim fra As MSForms.Frame
    Set fra = Me.fraInnhold.Controls.Add("Forms.Frame.1", "fraTile_" & mFlisNr, True)
    fra.Left = x: fra.Top = y: fra.Width = b: fra.Height = h
    fra.Caption = ""
    fra.BackColor = tileFarge
    ' Tynn kant gir flisene "kort"-preg selv uten avrundede hjorner.
    fra.BorderStyle = fmBorderStyleSingle
    On Error Resume Next
    fra.BorderColor = RGB(0, 0, 0)
    On Error GoTo 0

    Dim lblHoved As MSForms.Label
    Set lblHoved = fra.Controls.Add("Forms.Label.1", "lblHoved", True)
    lblHoved.Left = 10: lblHoved.Top = 5: lblHoved.Width = b - 24: lblHoved.Height = 18
    lblHoved.BackStyle = fmBackStyleTransparent
    lblHoved.Font.Bold = True
    lblHoved.Font.Size = 10
    lblHoved.ForeColor = tekstFarge
    lblHoved.Caption = CStr(f("Verdi")) & "   " & CStr(f("Antall"))

    Dim uy As Long
    uy = TILE_HOYDE_BASE
    Dim u As Variant
    For Each u In f("Under")
        Dim lblUnder As MSForms.Label
        Set lblUnder = fra.Controls.Add("Forms.Label.1", "lblUnder" & uy, True)
        lblUnder.Left = 24: lblUnder.Top = uy: lblUnder.Width = b - 40: lblUnder.Height = UNDERLINJE_HOYDE
        lblUnder.BackStyle = fmBackStyleTransparent
        lblUnder.ForeColor = tekstFarge
        lblUnder.Caption = CStr(u(0)) & "   " & CStr(u(1))
        ' Tomme-gruppen (se modStatistikkern.TOMME_VERDI) vises i kursiv.
        If CStr(u(0)) = modStatistikkern.TOMME_VERDI Then lblUnder.Font.Italic = True
        uy = uy + UNDERLINJE_HOYDE
    Next u
End Sub

' Leser oppsettet og bygger hele SM-innholdet pa nytt fra bunnen - kalt bade
' ved forste apning, ved bytte av visning og etter en tur innom Oppsett.
Private Sub GjenoppbyggInnhold()
    ' Nullstill rullingen (ellers kan en gammel rulleposisjon ligge igjen
    ' utenfor det nye innholdet ved bytte av visning).
    Me.fraInnhold.ScrollTop = 0
    Me.fraInnhold.ScrollLeft = 0
    RyddInnhold
    mFlisNr = 0

    Dim data As Collection
    Set data = ByggData()

    Dim flisB As Single, radB As Single, maksAntall As Long
    If mVisning = VIS_HORISONTAL And data.Count > 0 Then
        flisB = HorisontalFlisBredde(data)
        Dim sk As Object
        For Each sk In data
            If sk("Fliser").Count > maksAntall Then maksAntall = sk("Fliser").Count
        Next sk
        radB = maksAntall * flisB + (maksAntall - 1) * TILE_MELLOMROM
        If maksAntall < 1 Then radB = 0
        TilpassBredde radB
        Me.fraInnhold.ScrollBars = 3
        Me.fraInnhold.KeepScrollBarsVisible = 0
    Else
        Me.Width = mBasisBredde
        PlasserKontroller
        Me.fraInnhold.ScrollBars = 2
        Me.fraInnhold.KeepScrollBarsVisible = 3
    End If
    SentrerVindu

    If data.Count = 0 Then
        Dim lblTom As MSForms.Label
        Set lblTom = Me.fraInnhold.Controls.Add("Forms.Label.1", "lblTom", True)
        lblTom.Left = 10: lblTom.Top = 10: lblTom.Width = Me.fraInnhold.Width - 40: lblTom.Height = 40
        lblTom.WordWrap = True
        lblTom.BackStyle = fmBackStyleTransparent
        lblTom.ForeColor = RGB(110, 116, 128)
        lblTom.Caption = "Ingen oppsett ennå. Trykk ""Oppsett..."" for å velge en tabell og kolonner å føre statistikk for."
        Me.fraInnhold.ScrollHeight = 60
        Exit Sub
    End If

    Dim avail As Long
    avail = Me.fraInnhold.Width - TILE_HOYRE_MARG

    Dim y As Long
    y = 8
    Dim k As Long
    For k = 1 To data.Count
        Dim sek As Object
        Set sek = data(k)
        Dim fliser As Collection
        Set fliser = sek("Fliser")

        If k > 1 Then y = y + SEKSJON_MELLOMROM

        Dim lblSeksjon As MSForms.Label
        Set lblSeksjon = Me.fraInnhold.Controls.Add("Forms.Label.1", "lblSeksjon" & k, True)
        lblSeksjon.Left = 4: lblSeksjon.Top = y: lblSeksjon.Width = Me.fraInnhold.Width - 40: lblSeksjon.Height = 16
        lblSeksjon.Caption = CStr(sek("Tittel"))
        lblSeksjon.BackStyle = fmBackStyleTransparent
        lblSeksjon.Font.Bold = True
        lblSeksjon.ForeColor = RGB(90, 95, 105)
        y = y + 22

        If fliser.Count = 0 Then
            Dim lblIngen As MSForms.Label
            Set lblIngen = Me.fraInnhold.Controls.Add("Forms.Label.1", "lblIngen" & k, True)
            lblIngen.Left = 10: lblIngen.Top = y: lblIngen.Width = Me.fraInnhold.Width - 50: lblIngen.Height = 16
            lblIngen.BackStyle = fmBackStyleTransparent
            lblIngen.Caption = "(ingen data)"
            lblIngen.ForeColor = RGB(150, 155, 165)
            y = y + 20
        Else
            ' Antall kolonner og flisbredde per visning.
            Dim antKol As Long, flisBredde As Long, felleshoyde As Long
            felleshoyde = 0
            Select Case mVisning
                Case VIS_HORISONTAL
                    antKol = fliser.Count
                    flisBredde = CLng(flisB)
                    ' Like hoye fliser i raden (like brede er de allerede).
                    Dim ff As Object
                    For Each ff In fliser
                        If FlisHoyde(ff) > felleshoyde Then felleshoyde = FlisHoyde(ff)
                    Next ff
                Case VIS_RUTE
                    antKol = (avail + TILE_MELLOMROM) \ (RUTE_MIN_BREDDE + TILE_MELLOMROM)
                    If antKol < 1 Then antKol = 1
                    flisBredde = (avail - (antKol - 1) * TILE_MELLOMROM) \ antKol
                Case Else
                    antKol = 1
                    flisBredde = avail
            End Select

            Dim kolIdx As Long, radHoyde As Long
            kolIdx = 0
            radHoyde = 0
            Dim fl As Object
            For Each fl In fliser
                Dim h As Long
                If felleshoyde > 0 Then h = felleshoyde Else h = FlisHoyde(fl)
                TegnFlis 8 + kolIdx * (flisBredde + TILE_MELLOMROM), y, flisBredde, h, fl
                If h > radHoyde Then radHoyde = h
                kolIdx = kolIdx + 1
                If kolIdx >= antKol Then
                    y = y + radHoyde + TILE_MELLOMROM
                    kolIdx = 0
                    radHoyde = 0
                End If
            Next fl
            If kolIdx > 0 Then y = y + radHoyde + TILE_MELLOMROM
        End If
    Next k

    Me.fraInnhold.ScrollHeight = y + 10
    If mVisning = VIS_HORISONTAL Then Me.fraInnhold.ScrollWidth = radB + 24
End Sub
'@

$formCodeStatistikkernOppsett = @'
Attribute VB_Name = "frmStatistikkernOppsett"
Option Explicit

Private mRader As Collection ' cache av modStatistikkern.LesOppsett() - listeindeks i lstOppsett == indeks-1 her

Private Sub UserForm_Initialize()
    Me.Caption = "Statistikkern - oppsett"
    Me.BackColor = RGB(246, 247, 249)
    Me.lblTittel.Font.Bold = True
    Me.lblTittel.Font.Size = 12
    modStatistikkern.StatStilKnapp Me.cmdNy, "gronn"
    modStatistikkern.StatStilKnapp Me.cmdFjern, "rod"
    modStatistikkern.StatStilKnapp Me.cmdLukk, "gra"
    LastListe
End Sub

Private Sub LastListe()
    Me.lstOppsett.Clear
    Set mRader = modStatistikkern.LesOppsett()
    Dim i As Long
    For i = 1 To mRader.Count
        Dim rad As Object
        Set rad = mRader(i)
        Dim tekst As String
        tekst = CStr(rad("TabellNavn")) & "  (ark: " & CStr(rad("TabellArk")) & ")  -  Hovedkolonne: " & CStr(rad("HovedKolonne"))
        If CStr(rad("UnderKolonne")) <> "" Then
            tekst = tekst & "   [+ Under: " & CStr(rad("UnderKolonne")) & "]"
        End If
        Me.lstOppsett.AddItem tekst
    Next i
End Sub

Private Sub cmdNy_Click()
    frmStatistikkernNyTabell.Show
    LastListe
End Sub

Private Sub cmdFjern_Click()
    Dim valgte As New Collection
    Dim i As Long
    For i = 0 To Me.lstOppsett.ListCount - 1
        If Me.lstOppsett.Selected(i) Then
            valgte.Add CLng(mRader(i + 1)("Rad"))
        End If
    Next i
    If valgte.Count = 0 Then
        MsgBox "Velg minst én rad i listen først.", vbExclamation, "Fjern"
        Exit Sub
    End If
    If MsgBox("Fjerne " & valgte.Count & " oppføring(er) fra oppsettet?", vbQuestion + vbYesNo, "Fjern") <> vbYes Then Exit Sub
    modStatistikkern.FjernOppsettRader valgte
    LastListe
End Sub

Private Sub cmdLukk_Click()
    Unload Me
End Sub
'@

$formCodeStatistikkernNyTabell = @'
Attribute VB_Name = "frmStatistikkernNyTabell"
Option Explicit

Private mKolonneNavn() As String
Private mAntallKolonner As Long

Private Sub UserForm_Initialize()
    Me.Caption = "Statistikkern - ny tabell/kolonner"
    Me.BackColor = RGB(246, 247, 249)
    Me.lblTittel.Font.Bold = True
    Me.lblTittel.Font.Size = 12
    modStatistikkern.StatStilKnapp Me.cmdLagre, "gronn"
    modStatistikkern.StatStilKnapp Me.cmdAvbryt, "gra"

    Dim tabeller As Collection
    Set tabeller = modStatistikkern.AlleListObjekter()
    Me.cboTabell.Clear
    Dim i As Long
    For i = 1 To tabeller.Count
        Dim d As Object
        Set d = tabeller(i)
        Me.cboTabell.AddItem CStr(d("ArkNavn")) & ": " & CStr(d("TabellNavn"))
    Next i
    If tabeller.Count = 0 Then
        MsgBox "Fant ingen tabeller (Sett inn > Tabell) i denne arbeidsboken ennå. Opprett en tabell først.", vbExclamation, "Ingen tabeller"
    End If
End Sub

Private Sub cboTabell_Change()
    ByggKolonneListe
End Sub

Private Sub ByggKolonneListe()
    Dim i As Long
    For i = Me.fraKolonner.Controls.Count - 1 To 0 Step -1
        Me.fraKolonner.Controls.Remove Me.fraKolonner.Controls(i).Name
    Next i

    If Me.cboTabell.ListIndex < 0 Then Exit Sub

    Dim valgtTekst As String
    valgtTekst = Me.cboTabell.Value
    Dim delePos As Long
    delePos = InStr(valgtTekst, ": ")
    If delePos = 0 Then Exit Sub
    Dim arkNavn As String, tabellNavn As String
    arkNavn = Left$(valgtTekst, delePos - 1)
    tabellNavn = Mid$(valgtTekst, delePos + 2)

    Dim tbl As ListObject
    Set tbl = modStatistikkern.HentTabell(arkNavn, tabellNavn)
    If tbl Is Nothing Then Exit Sub

    ' Eksisterende oppsett for denne tabellen, sa vi kan forhandsutfylle
    ' avkrysning + underkolonne for kolonner som allerede er konfigurert
    ' (gjor dette vinduet trygt a apne pa nytt for a REDIGERE, ikke bare
    ' for forstegangsoppsett).
    Dim eksisterende As Object
    Set eksisterende = CreateObject("Scripting.Dictionary")
    Dim alleRader As Collection
    Set alleRader = modStatistikkern.LesOppsett()
    Dim r As Long
    For r = 1 To alleRader.Count
        Dim rad As Object
        Set rad = alleRader(r)
        If CStr(rad("TabellArk")) = arkNavn And CStr(rad("TabellNavn")) = tabellNavn Then
            eksisterende(CStr(rad("HovedKolonne"))) = CStr(rad("UnderKolonne"))
        End If
    Next r

    mAntallKolonner = tbl.ListColumns.Count
    ReDim mKolonneNavn(1 To mAntallKolonner)

    ' Kolonneoverskrifter over selve listen (radene under har ingen egne
    ' overskrifter) - samme mønster som Nettskjema-henters tilkoblingsliste.
    Dim lblOverHoved As MSForms.Label
    Set lblOverHoved = Me.fraKolonner.Controls.Add("Forms.Label.1", "lblOverHoved", True)
    lblOverHoved.Left = 8: lblOverHoved.Top = 2: lblOverHoved.Width = 230: lblOverHoved.Height = 14
    lblOverHoved.Caption = "Hovedkolonne"
    lblOverHoved.Font.Bold = True
    lblOverHoved.ForeColor = RGB(110, 116, 128)

    Dim lblOverUnder As MSForms.Label
    Set lblOverUnder = Me.fraKolonner.Controls.Add("Forms.Label.1", "lblOverUnder", True)
    lblOverUnder.Left = 290: lblOverUnder.Top = 2: lblOverUnder.Width = 248: lblOverUnder.Height = 14
    lblOverUnder.Caption = "Underkolonne (valgfritt)"
    lblOverUnder.Font.Bold = True
    lblOverUnder.ForeColor = RGB(110, 116, 128)

    Dim y As Long
    y = 20
    Dim c As Long
    For c = 1 To mAntallKolonner
        Dim kolNavn As String
        kolNavn = tbl.ListColumns(c).Name
        mKolonneNavn(c) = kolNavn

        Dim chk As MSForms.CheckBox
        Set chk = Me.fraKolonner.Controls.Add("Forms.CheckBox.1", "chkCol" & c, True)
        chk.Left = 8: chk.Top = y: chk.Width = 230: chk.Height = 18
        chk.Caption = kolNavn

        Dim cbo As MSForms.ComboBox
        Set cbo = Me.fraKolonner.Controls.Add("Forms.ComboBox.1", "cboUnder" & c, True)
        cbo.Left = 290: cbo.Top = y - 2: cbo.Width = 248: cbo.Height = 18
        cbo.Style = fmStyleDropDownList
        cbo.AddItem "(ingen)"
        Dim c2 As Long
        For c2 = 1 To mAntallKolonner
            If c2 <> c Then cbo.AddItem tbl.ListColumns(c2).Name
        Next c2
        cbo.ListIndex = 0

        If eksisterende.Exists(kolNavn) Then
            chk.Value = True
            Dim underVerdi As String
            underVerdi = CStr(eksisterende(kolNavn))
            If underVerdi <> "" Then
                Dim iu As Long
                For iu = 0 To cbo.ListCount - 1
                    If cbo.List(iu) = underVerdi Then
                        cbo.ListIndex = iu
                        Exit For
                    End If
                Next iu
            End If
        End If

        y = y + 24
    Next c

    Me.fraKolonner.ScrollHeight = y + 10
End Sub

Private Sub cmdLagre_Click()
    If Me.cboTabell.ListIndex < 0 Then
        MsgBox "Velg en tabell først.", vbExclamation, "Lagre"
        Exit Sub
    End If

    Dim valgtTekst As String
    valgtTekst = Me.cboTabell.Value
    Dim delePos As Long
    delePos = InStr(valgtTekst, ": ")
    Dim arkNavn As String, tabellNavn As String
    arkNavn = Left$(valgtTekst, delePos - 1)
    tabellNavn = Mid$(valgtTekst, delePos + 2)

    Dim valgteKolonner As Object
    Set valgteKolonner = CreateObject("Scripting.Dictionary")

    Dim c As Long
    For c = 1 To mAntallKolonner
        Dim chk As MSForms.CheckBox
        Set chk = Me.fraKolonner.Controls("chkCol" & c)
        If chk.Value = True Then
            Dim cbo As MSForms.ComboBox
            Set cbo = Me.fraKolonner.Controls("cboUnder" & c)
            Dim underVal As String
            underVal = ""
            If cbo.ListIndex > 0 Then underVal = cbo.Value
            valgteKolonner(mKolonneNavn(c)) = underVal
        End If
    Next c

    If valgteKolonner.Count = 0 Then
        MsgBox "Huk av minst én hovedkolonne, eller trykk Avbryt.", vbExclamation, "Lagre"
        Exit Sub
    End If

    modStatistikkern.LagreHovedkolonneValg arkNavn, tabellNavn, valgteKolonner
    Unload Me
End Sub

Private Sub cmdAvbryt_Click()
    Unload Me
End Sub
'@

function Get-StatistikkernVersion {
    param([string]$Code)
    if ($Code -match 'STATISTIKKERN_VERSION\s*As\s*String\s*=\s*"([^"]+)"') {
        return $Matches[1]
    }
    return $null
}

$sourceVersion = Get-StatistikkernVersion $moduleCodeStatistikkern

$allComponents = @(
    @{ Name = "modStatistikkern";         Code = $moduleCodeStatistikkern;        File = "modStatistikkern.bas"; Type = 1 },
    @{ Name = "frmStatistikkern";         Code = $formCodeStatistikkern;          File = "frmStatistikkern.frm"; Type = 3 },
    @{ Name = "frmStatistikkernOppsett";  Code = $formCodeStatistikkernOppsett;   File = "frmStatistikkernOppsett.frm"; Type = 4 },
    @{ Name = "frmStatistikkernNyTabell"; Code = $formCodeStatistikkernNyTabell;  File = "frmStatistikkernNyTabell.frm"; Type = 5 }
)

$buttonOnActions = @("AapneStatistikkern")

# ---- 1. Finn målfilen ----

if (-not $Path) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel-filer (*.xlsm)|*.xlsm|Alle filer (*.*)|*.*"
    $dlg.Title = "Velg Excel-fil som skal få Statistikkern (opprett en tom .xlsm først om nødvendig)"
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
    # knappene med denne makroens EGNE OnAction-navn. Oppsett-arket (skjult)
    # og selve dataene i brukerens tabeller røres IKKE - kun kode fjernes.
    Write-Output "Fjerner Statistikkern fra $($wb.Name) ..."
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
                if ($buttonOnActions -contains $btn.OnAction) { $btn.Delete() }
            }
        }
    } catch {}
    try { $wb.Save() } catch {
        Write-Output "FEIL ved lagring: $($_.Exception.Message)"
        exit 1
    }
    Write-Output "Ferdig - Statistikkern er fjernet fra $($wb.Name). Oppsett og ark er beholdt."
    exit 0
}

# ---- 3. Sjekk hva som allerede er installert ----

$existingModule = $null
foreach ($comp in $vbproj.VBComponents) {
    if ($comp.Name -eq "modStatistikkern") { $existingModule = $comp }
}

$installedVersion = $null
if ($existingModule) {
    if ($existingModule.CodeModule.CountOfLines -gt 0) {
        $existingCode = $existingModule.CodeModule.Lines(1, $existingModule.CodeModule.CountOfLines)
        $installedVersion = Get-StatistikkernVersion $existingCode
    } else {
        Write-Output "ADVARSEL: fant modStatistikkern, men den er tom (0 linjer) - sannsynligvis fra en tidligere avbrutt installasjon. Installerer på nytt."
        $existingModule = $null
    }
}

$isFreshInstall = ($null -eq $existingModule)

if (-not $isFreshInstall -and $installedVersion -eq $sourceVersion -and -not $Force) {
    Write-Output "Statistikkern er allerede installert og oppdatert (versjon $installedVersion) i $($wb.Name)."
    exit 0
}

if ($isFreshInstall) {
    Write-Output "Installerer Statistikkern (versjon $sourceVersion) i $($wb.Name) ..."
} else {
    Write-Output "Oppdaterer Statistikkern: versjon $installedVersion -> $sourceVersion i $($wb.Name) ..."
}

# ---- 4. Bygg de fire komponentene i en midlertidig arbeidsbok, eksporter til en temp-mappe ----

$tempDir = Join-Path $env:TEMP ("Statistikkern_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir | Out-Null

try {
    $buildWb = $excel.Workbooks.Add()
    $buildProj = $buildWb.VBProject

    foreach ($item in $allComponents) {
        if ($item.Type -eq 3) {
            # UserForm (Designer) - frmStatistikkern: hovedvinduet (SM). Selve
            # flisene bygges dynamisk av VBA-koden (UserForm_Initialize /
            # GjenoppbyggInnhold) - antallet er ikke kjent her.
            $comp = $buildProj.VBComponents.Add(3)
            $comp.Name = $item.Name
            $comp.Properties("Width").Value = 640
            $comp.Properties("Height").Value = 580
            $comp.Properties("Caption").Value = "Statistikkern"
            $designer = $comp.Designer

            # Ingen tittel/undertekst (Hakons onske) - mest mulig plass til
            # rullemenyen. Endelig posisjon/storrelse av alle tre kontroller
            # settes i VBA (UserForm_Initialize) ut fra vinduets InsideWidth/
            # InsideHeight, sa verdiene under er bare startverdier.
            $fraInnhold = $designer.Controls.Add("Forms.Frame.1", "fraInnhold", $true)
            $fraInnhold.Left = 12; $fraInnhold.Top = 12; $fraInnhold.Width = 600; $fraInnhold.Height = 470
            $fraInnhold.Caption = ""
            $fraInnhold.ScrollBars = 2

            $cmdOppsett = $designer.Controls.Add("Forms.CommandButton.1", "cmdOppsett", $true)
            $cmdOppsett.Left = 12; $cmdOppsett.Top = 494; $cmdOppsett.Width = 130; $cmdOppsett.Height = 30
            $cmdOppsett.Caption = "Oppsett..."

            $cmdVisning = $designer.Controls.Add("Forms.CommandButton.1", "cmdVisning", $true)
            $cmdVisning.Left = 150; $cmdVisning.Top = 494; $cmdVisning.Width = 130; $cmdVisning.Height = 30
            $cmdVisning.Caption = "Rutevisning"

            $comp.CodeModule.AddFromString($item.Code)
            $comp.Export((Join-Path $tempDir $item.File))
        } elseif ($item.Type -eq 4) {
            # UserForm (Designer) - frmStatistikkernOppsett: liste over
            # konfigurerte tabeller/hovedkolonner + Ny/Fjern/Lukk.
            $comp = $buildProj.VBComponents.Add(3)
            $comp.Name = $item.Name
            $comp.Properties("Width").Value = 500
            $comp.Properties("Height").Value = 380
            $comp.Properties("Caption").Value = "Statistikkern - oppsett"
            $designer = $comp.Designer

            $lblTittel = $designer.Controls.Add("Forms.Label.1", "lblTittel", $true)
            $lblTittel.Left = 16; $lblTittel.Top = 12; $lblTittel.Width = 400; $lblTittel.Height = 22
            $lblTittel.Caption = "Konfigurerte tabeller og kolonner"

            $lblUndertittel = $designer.Controls.Add("Forms.Label.1", "lblUndertittel", $true)
            $lblUndertittel.Left = 16; $lblUndertittel.Top = 34; $lblUndertittel.Width = 460; $lblUndertittel.Height = 14
            $lblUndertittel.Caption = "Tabeller og kolonner statistikken i hovedvinduet bygges fra."
            $lblUndertittel.ForeColor = 6579300

            $lstOppsett = $designer.Controls.Add("Forms.ListBox.1", "lstOppsett", $true)
            $lstOppsett.Left = 16; $lstOppsett.Top = 56; $lstOppsett.Width = 468; $lstOppsett.Height = 246
            $lstOppsett.MultiSelect = 2

            $cmdNy = $designer.Controls.Add("Forms.CommandButton.1", "cmdNy", $true)
            $cmdNy.Left = 16; $cmdNy.Top = 312; $cmdNy.Width = 190; $cmdNy.Height = 32
            $cmdNy.Caption = "+ Ny tabell/kolonner..."

            $cmdFjern = $designer.Controls.Add("Forms.CommandButton.1", "cmdFjern", $true)
            $cmdFjern.Left = 214; $cmdFjern.Top = 312; $cmdFjern.Width = 140; $cmdFjern.Height = 32
            $cmdFjern.Caption = "Fjern valgt"

            $cmdLukk = $designer.Controls.Add("Forms.CommandButton.1", "cmdLukk", $true)
            $cmdLukk.Left = 362; $cmdLukk.Top = 312; $cmdLukk.Width = 122; $cmdLukk.Height = 32
            $cmdLukk.Caption = "Lukk"

            $comp.CodeModule.AddFromString($item.Code)
            $comp.Export((Join-Path $tempDir $item.File))
        } elseif ($item.Type -eq 5) {
            # UserForm (Designer) - frmStatistikkernNyTabell: velg tabell, huk
            # av hovedkolonner + valgfri underkolonne per kolonne. Kolonnelisten
            # bygges dynamisk (ByggKolonneListe) - antall kolonner er ikke kjent
            # her, kun rammen den legges i.
            $comp = $buildProj.VBComponents.Add(3)
            $comp.Name = $item.Name
            $comp.Properties("Width").Value = 600
            $comp.Properties("Height").Value = 460
            $comp.Properties("Caption").Value = "Statistikkern - ny tabell/kolonner"
            $designer = $comp.Designer

            $lblTittel = $designer.Controls.Add("Forms.Label.1", "lblTittel", $true)
            $lblTittel.Left = 16; $lblTittel.Top = 12; $lblTittel.Width = 500; $lblTittel.Height = 22
            $lblTittel.Caption = "Ny tabell/kolonner"

            $lblUndertittel = $designer.Controls.Add("Forms.Label.1", "lblUndertittel", $true)
            $lblUndertittel.Left = 16; $lblUndertittel.Top = 34; $lblUndertittel.Width = 560; $lblUndertittel.Height = 14
            $lblUndertittel.Caption = "Huk av hovedkolonner, og velg valgfritt en underkolonne for hver."
            $lblUndertittel.ForeColor = 6579300

            $lblVelgTabell = $designer.Controls.Add("Forms.Label.1", "lblVelgTabell", $true)
            $lblVelgTabell.Left = 16; $lblVelgTabell.Top = 60; $lblVelgTabell.Width = 60; $lblVelgTabell.Height = 18
            $lblVelgTabell.Caption = "Tabell:"

            $cboTabell = $designer.Controls.Add("Forms.ComboBox.1", "cboTabell", $true)
            $cboTabell.Left = 80; $cboTabell.Top = 57; $cboTabell.Width = 504; $cboTabell.Height = 20
            $cboTabell.Style = 2

            $fraKolonner = $designer.Controls.Add("Forms.Frame.1", "fraKolonner", $true)
            $fraKolonner.Left = 16; $fraKolonner.Top = 86; $fraKolonner.Width = 568; $fraKolonner.Height = 310
            $fraKolonner.Caption = ""
            $fraKolonner.ScrollBars = 2

            $cmdLagre = $designer.Controls.Add("Forms.CommandButton.1", "cmdLagre", $true)
            $cmdLagre.Left = 16; $cmdLagre.Top = 406; $cmdLagre.Width = 276; $cmdLagre.Height = 34
            $cmdLagre.Caption = "Lagre"

            $cmdAvbryt = $designer.Controls.Add("Forms.CommandButton.1", "cmdAvbryt", $true)
            $cmdAvbryt.Left = 308; $cmdAvbryt.Top = 406; $cmdAvbryt.Width = 276; $cmdAvbryt.Height = 34
            $cmdAvbryt.Caption = "Avbryt"

            $comp.CodeModule.AddFromString($item.Code)
            $comp.Export((Join-Path $tempDir $item.File))
        } else {
            $comp = $buildProj.VBComponents.Add(1)   # vbext_ct_StdModule
            $comp.Name = $item.Name
            $comp.CodeModule.AddFromString($item.Code)
            $comp.Export((Join-Path $tempDir $item.File))
        }
    }

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

    # Ingen automatisk ark-knapp (samme "Fast regel" som de andre fire
    # makroene - tilgang skjer via Makromeny). Oppsett-registeret (skjult
    # ark) lager seg selv forste gang det trengs (se GetOppsettSheet i VBA).

    Write-Output ""
    Write-Output "Ferdig. Husk å lagre filen selv (Ctrl+S) når du er klar."
    if ($isFreshInstall) {
        Write-Output ""
        Write-Output "Neste steg:"
        Write-Output "  1. Kjør ""AapneStatistikkern"" (f.eks. via Makromeny - ingen knapp legges til i arket automatisk)."
        Write-Output "  2. Trykk ""Oppsett...""  ->  ""+ Ny tabell/kolonner...""."
        Write-Output "  3. Velg en tabell, huk av én eller flere hovedkolonner, og (valgfritt) en underkolonne per hovedkolonne."
        Write-Output "  4. Lagre, lukk oppsett-vinduene, og statistikken vises som fargede fliser i hovedvinduet."
    }
} finally {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
