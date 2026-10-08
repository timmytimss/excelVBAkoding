param(
    [string]$Path,
    [string]$Mappe,
    [switch]$Force,
    [switch]$Uninstall,
    # Kun for testing: start en EGEN Excel-prosess i stedet for a koble til den
    # som allerede kjorer (ellers apnes testfila inne i brukerens egen Excel).
    [switch]$NyExcelInstans
)

$ErrorActionPreference = 'Stop'

# Bump denne ved ENHVER endring i VBA-kilden under (moduleCodeKontaktsentralen/
# formCodeKontaktsentralen/formCodeKtsOppsett) - se excel-vba-installer-pattern
# i minnet for hvorfor: uten en bump gjor installeren ingenting pa neste
# kjoring selv om koden faktisk er endret.
$KtsVersion = "0.11.0"

# ============================================================
# Kontaktsentralen - alt-i-en installer/oppdaterer.
#
# Sporer hvilke rader (f.eks. skoler) i en tabell som har fatt en
# kategorisert e-post i en delt Outlook-postboks - uavhengig av
# Mail-utsender. Denne makroen har sin EGEN konfigurasjon (skjulte
# faner "KontaktsentralenKonfig" og "KontaktsentralenCache") og
# fungerer selv om Mail-utsender aldri er installert pa samme ark -
# se excel-macro-installer-pattern i minnet for hvorfor hver makro
# skal vaere sitt eget selvstendige installer-script.
#
# Alle VBA-komponentnavn er prefikset "Kts" (modKontaktsentralen,
# frmKontaktsentralen, frmKtsOppsett, frmKtsRad, frmKtsUavklart) for
# a garantere at de aldri kolliderer med Mail-utsender sine
# komponentnavn (modGenerisk, frmSendMail, frmOppsett, frmMasseSend,
# frmFeilsjekk) nar begge makroene er installert pa samme ark.
#
# STOR OMBYGGING 2026-09-18 (v0.7.0), etter at Håkon selv gjorde
# betydelige endringer i en separat ChatGPT-samtale (v0.6.2) og ba om
# stabilisering/restrukturering rundt DEN arkitekturen, ikke en
# tilbakerulling: Outlook-kategorier (MailItem.Categories) er na det
# sentrale statussignalet, ikke en fast Svart/Videresendt/Automatisk
# svar/Ikke levert-klassifisering. Oppfolging-aksen (Venter/Ma
# purres/Purret/Avklart-*) er fjernet helt - Håkon og teamet hans
# styrer oppfolging via sitt eget Outlook-kategori-system.
#
# Hoved-arkitektur:
#  - En Kontrollkolonne (obligatorisk, valgt i Innstillinger) fylles
#    automatisk med den kategorien som har hoyest prioritet (en
#    brukerordnet liste, KTS_KATEGORI_PRIORITET) blant e-postene som
#    er matchet til raden.
#  - modKontaktsentralen.KjørSkann() kobler til den valgte delte
#    Outlook-postboksen og gar gjennom kategoriserte meldinger.
#    Matching prover e-postadresse FORST (eksakt, normalisert - aldri
#    substring), gjenkjenningskolonner (emne+brodtekst, helords-
#    substring) som fallback. En PERSISTENT cache (skjult fane
#    "KontaktsentralenCache", nokkel = Outlook EntryID) gjor at en
#    allerede kjent melding ALDRI kjores gjennom matching-motoren pa
#    nytt ved senere skanninger - bare kategorien leses pa nytt.
#    Samtalekoblinger (ConversationID) er ogsa persistente via cachen.
#  - Dashbordet (frmKontaktsentralen) viser ALLE rader (aldri skjult),
#    sortert med kategoriserte rader forst. Dobbeltklikk apner
#    frmKtsRad - full meldingshistorikk for raden (med Outlook-mappe),
#    postboksens kategorier som knapper som slas av/pa for den valgte
#    e-posten (Lagre i Outlook / Angre), apne mail i Outlook, og en snarvei til a
#    feste en uavklart e-post til nettopp DENNE raden (frmKtsUavklart).
#  - Uklare/uavklarte treff (ingen eller flertydig match) havner i en
#    egen liste (frmKtsUavklart) for manuell matching i stedet for a
#    bli gjettet.
# Se Kontaktsentralen-planen i minnet for full byggehistorikk.
#
# Samme grunnmonster som Mail-utsender.ps1:
#  - Bygger VBA-komponenter i en midlertidig arbeidsbok, eksporterer
#    dem, og importerer i malfilen. Ingen andre filer trengs.
#  - Excel-fila (.xlsm) VELGES ALLTID BEVISST via filvelger-dialog -
#    aldri automatisk gjetting.
#  - Portabel sti-hakndtering trengs ikke her (ingen mal-fil-sti a
#    regne ut, i motsetning til Mail-utsender).
#  - Scriptet lagrer ALDRI brukerens fil selv ved vanlig installasjon/
#    oppdatering - kun ved -Uninstall (se der for hvorfor).
#
# Bruk:
#   .\Kontaktsentralen.ps1
# ============================================================

# ---------------------------------------------------------------
# VBA-kildekode: modKontaktsentralen (standard-modul)
# ---------------------------------------------------------------
$moduleCodeKontaktsentralen = @'
Attribute VB_Name = "modKontaktsentralen"
Option Explicit

' ============================================================
' Kontaktsentralen
' Installeres/oppdateres av Kontaktsentralen.ps1 - se den fila for
' hele installer-monsteret (samme grunnmonster som Mail-utsender.ps1,
' men helt uavhengig konfigurasjon og komponentnavn).
'
' Konfig-lesing/skriving, tabell-hjelpefunksjoner, skannemotor
' (KjørSkann), matching og dashbord-hjelpefunksjonene (statistikk,
' radliste, oppfolgingsstatus) er alle her.
' ============================================================

Public Const KTS_VERSION As String = "{{VERSION}}"

Private Const KTS_KONFIG_FANENAVN As String = "KontaktsentralenKonfig"
Private Const KTS_CACHE_FANENAVN As String = "KontaktsentralenCache"

' Siste skanning sine treff, radNr -> Dictionary(Emne/Avsender/Mottatt/
' Kategori) - kun for forhandsvisning i dashbordet i inneværende okt,
' lagres ikke noe sted og overlever ikke at Excel lukkes/apnes pa nytt.
Public KtsSisteTreff As Object
Public KtsSisteMeldinger As Object
Public KtsValgtRad As Long

' Satt av frmKtsRad rett for den apner frmKtsUavklart, sa uavklart-
' vinduet vet a returnere til RADSIDEN igjen (i stedet for hoveddash-
' bordet) nar brukeren er ferdig - se cmdTilbake_Click i frmKtsUavklart.
Public KtsUavklartApnetFraRad As Boolean

' Siste skanning sine uavklarte treff - en Collection av Dictionary
' (Emne/Avsender/Mottatt/Kategori/Arsak), 1-indeksert. Brukes av
' frmKtsUavklart til a la Håkon manuelt matche en uavklart e-post til
' riktig rad. Samme session-only begrensning som KtsSisteTreff.
Public KtsSisteUavklarte As Collection

' Satt av FinnMatchendeRadNy ved flertydig treff: hvilke arkrader som delte
' toppscoren - vises i Arsak-teksten sa Haakon ser HVORFOR det ble flertydig.
Private KtsSisteKandidatTekst As String

' --- Konfig-lesing/skriving (skjult fane "KontaktsentralenKonfig") ---
' Egen fane, adskilt fra Mail-utsender sin "MailKonfig" - denne makroen
' skal fungere selv om Mail-utsender aldri er installert pa samme ark.

Private Function KtsKonfigFane() As Worksheet
    On Error Resume Next
    Set KtsKonfigFane = ThisWorkbook.Worksheets(KTS_KONFIG_FANENAVN)
    On Error GoTo 0
    If KtsKonfigFane Is Nothing Then
        Set KtsKonfigFane = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        KtsKonfigFane.Name = KTS_KONFIG_FANENAVN
        KtsKonfigFane.Range("A1").Value = "Nokkel"
        KtsKonfigFane.Range("B1").Value = "Verdi"
        KtsKonfigFane.Visible = xlSheetVeryHidden
    End If
End Function

Public Function KtsLesKonfig() As Object
    Dim ws As Worksheet, dict As Object, r As Long
    Dim nokkel As String, verdi As String
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = KtsKonfigFane()
    r = 2
    Do While r <= ws.Rows.Count
        If IsError(ws.Cells(r, 1).Value) Or IsNull(ws.Cells(r, 1).Value) Or IsEmpty(ws.Cells(r, 1).Value) Then Exit Do
        nokkel = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(nokkel) = 0 Then Exit Do
        verdi = ""
        If Not IsError(ws.Cells(r, 2).Value) And Not IsNull(ws.Cells(r, 2).Value) And Not IsEmpty(ws.Cells(r, 2).Value) Then verdi = CStr(ws.Cells(r, 2).Value)
        dict(nokkel) = verdi
        r = r + 1
    Loop
    Set KtsLesKonfig = dict
End Function

Public Sub KtsSkrivKonfigVerdi(ByVal nokkel As String, ByVal verdi As String)
    Dim ws As Worksheet, r As Long
    Set ws = KtsKonfigFane()
    r = 2
    Do While ws.Cells(r, 1).Value <> ""
        If ws.Cells(r, 1).Value = nokkel Then
            ws.Cells(r, 2).Value = verdi
            Exit Sub
        End If
        r = r + 1
    Loop
    ws.Cells(r, 1).Value = nokkel
    ws.Cells(r, 2).Value = verdi
End Sub

' --- Tabell-hjelpefunksjoner (samme grunnidé som Mail-utsender sin
'     AktivTabell/KolonneNr i modGenerisk, men egen kopi her for a
'     holde denne makroen helt uavhengig) ---

' Tabellen Kontaktsentralen jobber mot: den som er valgt i Innstillinger
' (KTS_TABELL), ellers den forste tabellen i arbeidsboka.
Public Function KtsAktivTabell() As ListObject
    ' Ingen automatisk "forste tabell i boka"-fallback lenger (Håkons
    ' eksplisitte onske 2026-09-22: tabellvalg skal alltid vare et bevisst
    ' valg i Innstillinger, aldri stille gjettet - en hovedfil med mange
    ' ark kan ha "feil" forste tabell). Er KTS_TABELL ikke satt/finnes ikke,
    ' returneres Nothing og VisKontaktsentralen sender brukeren til Oppsett.
    Dim valgt As String
    valgt = KtsValgtTabellNavn()
    If Len(valgt) > 0 Then Set KtsAktivTabell = KtsFinnTabell(valgt)
End Function

Private Function KtsValgtTabellNavn() As String
    ' Ikke opprett konfigarket bare for a spore opp tabellen (denne kalles
    ' ogsa fra dobbeltklikk-handleren i arbeidsbokas ovrige celler).
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(KTS_KONFIG_FANENAVN)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim k As Object: Set k = KtsLesKonfig()
    If k.Exists("KTS_TABELL") Then KtsValgtTabellNavn = Trim(CStr(k("KTS_TABELL")))
End Function

' Finner en tabell (ListObject) med dette navnet i hvilket som helst ark.
Public Function KtsFinnTabell(ByVal navn As String) As ListObject
    Dim ws As Worksheet, tbl As ListObject
    If Len(navn) = 0 Then Exit Function
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> KTS_KONFIG_FANENAVN And ws.Name <> KTS_CACHE_FANENAVN Then
            For Each tbl In ws.ListObjects
                If StrComp(tbl.Name, navn, vbTextCompare) = 0 Then
                    Set KtsFinnTabell = tbl
                    Exit Function
                End If
            Next tbl
        End If
    Next ws
End Function

' Alle tabeller i arbeidsboka (utenom Kontaktsentralens egne skjulte ark).
Public Function KtsAlleTabeller() As Collection
    Dim res As New Collection, ws As Worksheet, tbl As ListObject
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> KTS_KONFIG_FANENAVN And ws.Name <> KTS_CACHE_FANENAVN Then
            For Each tbl In ws.ListObjects
                res.Add tbl
            Next tbl
        End If
    Next ws
    Set KtsAlleTabeller = res
End Function

Public Function KtsKolonneNr(ByVal tbl As ListObject, ByVal kolonnenavn As String) As Long
    Dim kol As ListColumn
    If kolonnenavn = "" Then Exit Function
    If tbl Is Nothing Then Exit Function
    For Each kol In tbl.ListColumns
        If kol.Name = kolonnenavn Then
            KtsKolonneNr = kol.Index
            Exit Function
        End If
    Next kol
End Function

' --- Radnokkel: en fast, unik verdi per rad (f.eks. skole-ID) som e-postene
'     knyttes til i stedet for radnummer. Da tåler tabellen at brukeren
'     sorterer, filtrerer eller setter inn rader mellom to skanninger - en
'     rads posisjon endrer seg, men nokkelen den har gjor ikke det. ---

' Kolonnen som brukes som radnokkel: den som er valgt i Innstillinger
' (KTS_KOL_NOKKEL), ellers forste gjenkjenningskolonne, ellers tabellens
' forste kolonne.
Public Function KtsNokkelKolonneNavn(ByVal tbl As ListObject, ByVal konfig As Object) As String
    Dim s As String, kand As Variant
    If tbl Is Nothing Then Exit Function
    If konfig.Exists("KTS_KOL_NOKKEL") Then
        s = Trim(CStr(konfig("KTS_KOL_NOKKEL")))
        If Len(s) > 0 Then
            If KtsKolonneNr(tbl, s) > 0 Then
                KtsNokkelKolonneNavn = s
                Exit Function
            End If
        End If
    End If
    If konfig.Exists("KTS_KOL_GJENKJENNING") Then
        For Each kand In Split(CStr(konfig("KTS_KOL_GJENKJENNING")), "|")
            s = Trim(CStr(kand))
            If Len(s) > 0 Then
                If KtsKolonneNr(tbl, s) > 0 Then
                    KtsNokkelKolonneNavn = s
                    Exit Function
                End If
            End If
        Next kand
    End If
    If tbl.ListColumns.Count > 0 Then KtsNokkelKolonneNavn = tbl.ListColumns(1).Name
End Function

' Verdiene i en kolonne som 1-basert String-array (trimmet), eller Empty hvis
' kolonnen/tabellen mangler. Element r = tabellens rad r.
Public Function KtsNokkelVerdierForKolonne(ByVal tbl As ListObject, ByVal kolNavn As String) As Variant
    Dim kolNr As Long: kolNr = KtsKolonneNr(tbl, kolNavn)
    If kolNr = 0 Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = tbl.ListColumns(kolNr).DataBodyRange.Value
    Dim n As Long: n = tbl.DataBodyRange.Rows.Count
    Dim res() As String, r As Long
    ReDim res(1 To n)
    For r = 1 To n
        If IsArray(v) Then
            If IsError(v(r, 1)) Then res(r) = "" Else res(r) = Trim(CStr(v(r, 1)))
        Else
            If IsError(v) Then res(r) = "" Else res(r) = Trim(CStr(v))
        End If
    Next r
    KtsNokkelVerdierForKolonne = res
End Function

' Nokkel (lowercase) -> radnummer. -1 = nokkelen finnes pa flere rader.
' antallTomme = rader uten nokkel, antallDuplikater = nokkelverdier som
' forekommer mer enn en gang.
Private Function KtsKartFraNokler(nk As Variant, ByRef antallTomme As Long, ByRef antallDuplikater As Long) As Object
    Dim kart As Object: Set kart = CreateObject("Scripting.Dictionary")
    antallTomme = 0: antallDuplikater = 0
    Set KtsKartFraNokler = kart
    If IsEmpty(nk) Then Exit Function
    Dim r As Long, k As String
    For r = 1 To UBound(nk)
        k = LCase(CStr(nk(r)))
        If Len(k) = 0 Then
            antallTomme = antallTomme + 1
        ElseIf kart.Exists(k) Then
            If CLng(kart(k)) <> -1 Then antallDuplikater = antallDuplikater + 1
            kart(k) = -1
        Else
            kart.Add k, r
        End If
    Next r
End Function

' Sjekker en kandidat til radnokkel-kolonne (brukes av Innstillinger for a
' advare for lagring).
Public Sub KtsSjekkNokkelkolonne(ByVal tbl As ListObject, ByVal kolNavn As String, ByRef antallTomme As Long, ByRef antallDuplikater As Long)
    Dim nk As Variant: nk = KtsNokkelVerdierForKolonne(tbl, kolNavn)
    Dim kart As Object: Set kart = KtsKartFraNokler(nk, antallTomme, antallDuplikater)
End Sub

' Nokkelverdien til en rad i den aktive tabellen (tom hvis ingen).
Public Function KtsNokkelForRad(ByVal rad As Long) As String
    Dim tbl As ListObject: Set tbl = KtsAktivTabell()
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If rad < 1 Or rad > tbl.DataBodyRange.Rows.Count Then Exit Function
    Dim kolNr As Long: kolNr = KtsKolonneNr(tbl, KtsNokkelKolonneNavn(tbl, KtsLesKonfig()))
    If kolNr = 0 Then Exit Function
    Dim v As Variant: v = tbl.DataBodyRange.Cells(rad, kolNr).Value
    If IsError(v) Then Exit Function
    KtsNokkelForRad = Trim(CStr(v))
End Function

' Tom streng hvis raden kan ha e-post koblet til seg, ellers en forklaring
' (raden mangler nokkel, eller nokkelen finnes pa flere rader).
Public Function KtsRadNokkelFeil(ByVal rad As Long) As String
    Dim tbl As ListObject: Set tbl = KtsAktivTabell()
    If tbl Is Nothing Then Exit Function
    Dim konfig As Object: Set konfig = KtsLesKonfig()
    Dim kolNavn As String: kolNavn = KtsNokkelKolonneNavn(tbl, konfig)
    Dim nk As Variant: nk = KtsNokkelVerdierForKolonne(tbl, kolNavn)
    If IsEmpty(nk) Then Exit Function
    If rad < 1 Or rad > UBound(nk) Then Exit Function
    Dim tomme As Long, dup As Long
    Dim kart As Object: Set kart = KtsKartFraNokler(nk, tomme, dup)
    Dim k As String: k = LCase(CStr(nk(rad)))
    If Len(k) = 0 Then
        KtsRadNokkelFeil = "Raden har ingen verdi i radnøkkel-kolonnen '" & kolNavn & "', så e-post kan ikke kobles til den. Fyll ut verdien, eller velg en annen radnøkkel-kolonne i Innstillinger."
    ElseIf CLng(kart(k)) = -1 Then
        KtsRadNokkelFeil = "Verdien '" & CStr(nk(rad)) & "' i radnøkkel-kolonnen '" & kolNavn & "' finnes på flere rader, så e-post kan ikke kobles sikkert til denne raden. Velg en radnøkkel-kolonne med unike verdier i Innstillinger."
    End If
End Function

' Gir hver cache-oppforing riktig radnummer ut fra radnokkelen (sa cachen
' folger radene ogsa etter sortering/filtrering/innsetting).
' migrering = cachen er fra for radnokkelen fantes: da er lagret radnummer
' siste kjente posisjon, og nokkelen leses fra raden som na ligger der.
Private Sub KtsLosOppRadNokler(ByVal cache As Object, ByVal migrering As Boolean)
    Dim tbl As ListObject: Set tbl = KtsAktivTabell()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim konfig As Object: Set konfig = KtsLesKonfig()
    Dim nk As Variant: nk = KtsNokkelVerdierForKolonne(tbl, KtsNokkelKolonneNavn(tbl, konfig))
    If IsEmpty(nk) Then Exit Sub
    Dim tomme As Long, dup As Long
    Dim kart As Object: Set kart = KtsKartFraNokler(nk, tomme, dup)
    Dim n As Long: n = UBound(nk)
    Dim key As Variant, d As Object, rad As Long, k As String
    For Each key In cache.Keys
        Set d = cache(key)
        rad = CLng(d("RadNr"))
        If migrering And rad > 0 Then
            If rad <= n Then d("RadNokkel") = CStr(nk(rad)) Else d("RadNokkel") = ""
        End If
        k = LCase(CStr(d("RadNokkel")))
        If Len(k) > 0 Then
            If Not kart.Exists(k) Then
                d("RadNr") = 0
                d("Arsak") = "Raden med nøkkel '" & CStr(d("RadNokkel")) & "' finnes ikke lenger i tabellen"
            ElseIf CLng(kart(k)) = -1 Then
                d("RadNr") = -1
                d("Arsak") = "Radnøkkelen '" & CStr(d("RadNokkel")) & "' finnes på flere rader"
            Else
                d("RadNr") = CLng(kart(k))
                d("Arsak") = ""
            End If
        ElseIf migrering And rad > 0 Then
            d("RadNr") = 0
            d("Arsak") = "Raden mangler radnøkkel"
        End If
    Next key
End Sub

' For lagring: setter radnokkel pa alle oppforinger med en rad. En rad som
' ikke har en (unik) nokkel kan ikke huskes trygt - da havner e-posten i
' uavklart-lista i stedet for a bli knyttet til en rad som senere kan flytte seg.
Private Sub KtsSettRadNokler(ByVal cache As Object)
    Dim tbl As ListObject: Set tbl = KtsAktivTabell()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim nk As Variant: nk = KtsNokkelVerdierForKolonne(tbl, KtsNokkelKolonneNavn(tbl, KtsLesKonfig()))
    If IsEmpty(nk) Then Exit Sub
    Dim tomme As Long, dup As Long
    Dim kart As Object: Set kart = KtsKartFraNokler(nk, tomme, dup)
    Dim key As Variant, d As Object, rad As Long, k As String
    For Each key In cache.Keys
        Set d = cache(key)
        rad = CLng(d("RadNr"))
        If rad > 0 Then
            k = ""
            If rad <= UBound(nk) Then k = CStr(nk(rad))
            If Len(k) = 0 Then
                d("RadNr") = 0
                d("RadNokkel") = ""
                d("Arsak") = "Raden mangler radnøkkel"
            ElseIf CLng(kart(LCase(k))) = -1 Then
                d("RadNr") = -1
                d("RadNokkel") = k
                d("Arsak") = "Radnøkkelen '" & k & "' finnes på flere rader"
            Else
                d("RadNokkel") = k
            End If
        End If
    Next key
End Sub

' --- Persistent cache (skjult fane KontaktsentralenCache) ---
' En rad per Outlook-melding Kontaktsentralen noensinne har matchet
' (eller forsokt a matche) til en tabellrad. Gjor gjentatte skanninger
' raske: en melding som allerede er kjent (samme EntryID) trenger ALDRI
' kjores gjennom den kostbare matching-motoren pa nytt, bare kategorien
' leses pa nytt i tilfelle den er endret. Gjor ogsa samtalekoblinger og
' manuelle uavklart-matcher persistente pa tvers av okter, ikke bare
' innenfor en enkelt skanning slik 0.6.2 gjorde.
Private Function KtsCacheFane() As Worksheet
    On Error Resume Next
    Set KtsCacheFane = ThisWorkbook.Worksheets(KTS_CACHE_FANENAVN)
    On Error GoTo 0
    If KtsCacheFane Is Nothing Then
        Set KtsCacheFane = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        KtsCacheFane.Name = KTS_CACHE_FANENAVN
        KtsCacheFane.Range("A1").Value = "EntryID"
        KtsCacheFane.Range("B1").Value = "StoreID"
        KtsCacheFane.Range("C1").Value = "RadNr"
        KtsCacheFane.Range("D1").Value = "Kategori"
        KtsCacheFane.Range("E1").Value = "Emne"
        KtsCacheFane.Range("F1").Value = "Avsender"
        KtsCacheFane.Range("G1").Value = "Mottatt"
        KtsCacheFane.Range("H1").Value = "ConversationID"
        KtsCacheFane.Range("I1").Value = "Arsak"
        KtsCacheFane.Range("J1").Value = "Mappe"
        KtsCacheFane.Range("K1").Value = "RadNokkel"
        KtsCacheFane.Visible = xlSheetVeryHidden
    End If
End Function

' Leser hele cachen til en Dictionary: EntryID -> Dictionary(RadNr/
' Kategori/Emne/Avsender/Mottatt/ConversationID/StoreID/Arsak/Mappe/RadNokkel).
' RadNr er radens posisjon I DAG - funnet via radnokkelen, sa den er riktig
' selv om tabellen er sortert eller filtrert siden cachen ble lagret.
Public Function KtsLastCache() As Object
    Dim resultat As Object: Set resultat = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = KtsCacheFane()
    Dim migrering As Boolean
    migrering = (Trim(CStr(ws.Cells(1, 11).Value)) <> "RadNokkel")
    Dim sisteRad As Long: sisteRad = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If sisteRad >= 2 Then
        Dim v As Variant: v = ws.Range(ws.Cells(2, 1), ws.Cells(sisteRad, 11)).Value
        Dim i As Long, d As Object, entryIdNokkel As String
        For i = 1 To UBound(v, 1)
            entryIdNokkel = Trim(CStr(v(i, 1)))
            If Len(entryIdNokkel) > 0 Then
                Set d = CreateObject("Scripting.Dictionary")
                d("EntryID") = entryIdNokkel
                d("StoreID") = CStr(v(i, 2))
                d("RadNr") = CLng(Val(v(i, 3)))
                d("Kategori") = CStr(v(i, 4))
                d("Emne") = CStr(v(i, 5))
                d("Avsender") = CStr(v(i, 6))
                d("Mottatt") = CStr(v(i, 7))
                d("ConversationID") = CStr(v(i, 8))
                d("Arsak") = CStr(v(i, 9))
                d("Mappe") = CStr(v(i, 10))
                d("RadNokkel") = CStr(v(i, 11))
                If Not resultat.Exists(entryIdNokkel) Then resultat.Add entryIdNokkel, d
            End If
        Next i
    End If
    KtsLosOppRadNokler resultat, migrering
    Set KtsLastCache = resultat
End Function

' Skriver hele cachen tilbake (enkel, trygg full-omskriving - cachen er
' typisk noen hundre/fa tusen rader, ikke stor nok til at delta-oppdatering
' er verdt kompleksiteten/risikoen).
Public Sub KtsLagreCache(ByVal cache As Object)
    KtsSettRadNokler cache
    Dim ws As Worksheet: Set ws = KtsCacheFane()
    Dim sisteRad As Long: sisteRad = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If sisteRad > 1 Then ws.Range(ws.Cells(2, 1), ws.Cells(sisteRad, 11)).ClearContents
    ws.Cells(1, 9).Value = "Arsak"
    ws.Cells(1, 10).Value = "Mappe"
    ws.Cells(1, 11).Value = "RadNokkel"
    ' Tekstformat, ellers gjor Excel om en nokkel som "0012" til tallet 12.
    ws.Columns(11).NumberFormat = "@"
    Dim r As Long: r = 2
    Dim key As Variant, d As Object
    For Each key In cache.Keys
        Set d = cache(key)
        ws.Cells(r, 1).Value = d("EntryID")
        ws.Cells(r, 2).Value = d("StoreID")
        ws.Cells(r, 3).Value = d("RadNr")
        ws.Cells(r, 4).Value = d("Kategori")
        ws.Cells(r, 5).Value = d("Emne")
        ws.Cells(r, 6).Value = d("Avsender")
        ws.Cells(r, 7).Value = d("Mottatt")
        ws.Cells(r, 8).Value = d("ConversationID")
        If d.Exists("Arsak") Then ws.Cells(r, 9).Value = d("Arsak") Else ws.Cells(r, 9).Value = ""
        If d.Exists("Mappe") Then ws.Cells(r, 10).Value = d("Mappe") Else ws.Cells(r, 10).Value = ""
        If d.Exists("RadNokkel") Then ws.Cells(r, 11).Value = CStr(d("RadNokkel")) Else ws.Cells(r, 11).Value = ""
        r = r + 1
    Next key
End Sub

' Bygger de session-variablene dashbordet/detaljsiden leser (KtsSisteTreff/
' KtsSisteMeldinger) fra HELE cachen - ikke bare det som ble sett i denne
' skanningen - sa detaljsiden alltid viser full historikk for en rad, pa
' tvers av Excel-restarter.
Public Sub KtsGjenoppbyggSessionFraCache(ByVal cache As Object)
    Set KtsSisteTreff = CreateObject("Scripting.Dictionary")
    Set KtsSisteMeldinger = CreateObject("Scripting.Dictionary")
    Set KtsSisteUavklarte = New Collection
    Dim key As Variant, d As Object, info As Object, rad As Long, u As Object
    For Each key In cache.Keys
        Set d = cache(key)
        rad = CLng(d("RadNr"))
        If rad <= 0 Then
            ' Ikke matchet til noen rad (0 = ingen sikker match, -1 = flertydig)
            ' - huskes i cachen slik at uavklart-lista og tallet pa knappen
            ' overlever at Excel lukkes.
            Set u = CreateObject("Scripting.Dictionary")
            u("Emne") = d("Emne"): u("Avsender") = d("Avsender"): u("Mottatt") = d("Mottatt")
            u("Kategori") = d("Kategori"): u("EntryID") = d("EntryID"): u("StoreID") = d("StoreID")
            u("ConversationID") = d("ConversationID"): u("Mappe") = d("Mappe")
            If Len(CStr(d("Arsak"))) > 0 Then u("Arsak") = d("Arsak") Else u("Arsak") = "Ingen sikker match"
            KtsSisteUavklarte.Add u
        Else
            Set info = CreateObject("Scripting.Dictionary")
            info("Emne") = d("Emne"): info("Avsender") = d("Avsender"): info("Mottatt") = d("Mottatt")
            info("Kategori") = d("Kategori"): info("EntryID") = d("EntryID"): info("StoreID") = d("StoreID")
            info("Mappe") = d("Mappe")
            If Not KtsSisteMeldinger.Exists(rad) Then
                Dim c As Collection: Set c = New Collection
                KtsSisteMeldinger.Add rad, c
            End If
            KtsSisteMeldinger(rad).Add info
            If KtsSisteTreff.Exists(rad) Then KtsSisteTreff.Remove rad
            KtsSisteTreff.Add rad, info
        End If
    Next key
End Sub

' Samler en mappe og ALLE undermapper (rekursivt, i dybden) i en
' Collection. Brukt bade av skannemotoren og kategori-oppdagelsen i
' Oppsett, sa begge finner e-post uansett hvilken undermappe en Outlook-
' regel har sortert den til.
Private Sub KtsSamleMapperRekursivt(ByVal mappe As Object, ByVal samling As Collection)
    samling.Add mappe
    Dim undermappe As Object
    For Each undermappe In mappe.Folders
        KtsSamleMapperRekursivt undermappe, samling
    Next undermappe
End Sub

' Kalt fra ThisWorkbook sin Workbook_SheetBeforeDoubleClick (satt inn av
' installeren). Dobbeltklikk pa en celle i Kontrollkolonnen i tabellen apner
' radsiden for akkurat den raden. Alle andre dobbeltklikk lates helt urort -
' ingen feilmelding, ingen endring, selv om noe her feiler i sjekkene.
Public Sub KtsHaandterDobbeltklikk(ByVal Sh As Object, ByVal Target As Range, ByRef Cancel As Boolean)
    Dim lo As ListObject, tbl As ListObject, konfig As Object
    Dim kolNr As Long, rad As Long, kolRad As Range

    On Error Resume Next
    If Target.Cells.Count <> 1 Then Exit Sub
    Set lo = Target.ListObject
    If lo Is Nothing Then Exit Sub
    Set tbl = KtsAktivTabell()
    If tbl Is Nothing Then Exit Sub
    If lo.Name <> tbl.Name Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Set konfig = KtsLesKonfig()
    If Not konfig.Exists("KTS_KOL_RESPONS") Then Exit Sub
    kolNr = KtsKolonneNr(tbl, CStr(konfig("KTS_KOL_RESPONS")))
    If kolNr = 0 Then Exit Sub
    Set kolRad = tbl.ListColumns(kolNr).DataBodyRange
    If Intersect(Target, kolRad) Is Nothing Then Exit Sub
    rad = Target.Row - tbl.DataBodyRange.Row + 1
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo Feil

    If rad < 1 Then Exit Sub
    Cancel = True
    KtsValgtRad = rad
    KtsSikreSessionFraCache
    frmKtsRad.Show
    Exit Sub
Feil:
    MsgBox "Klarte ikke å åpne radsiden: " & Err.Description, vbExclamation, "Kontaktsentralen"
End Sub

' Bygger session-variablene (meldingslister per rad) pa nytt fra den lagrede
' cachen - brukes nar dashbordet/radsiden apnes uten a skanne. Bygges ALLTID
' pa nytt (ikke bare forste gang), for radene kan ha flyttet seg (sortering,
' filter) siden sist, og radnokkelen gir da riktige radnummer.
Public Sub KtsSikreSessionFraCache()
    Dim cache As Object: Set cache = KtsLastCache()
    KtsGjenoppbyggSessionFraCache cache
End Sub

' --- Skannemotor ---
' Leser den valgte delte postboksen og gar gjennom kategoriserte
' meldinger. For hver melding: hvis EntryID allerede er kjent fra
' cachen, gjenbruk den lagrede radkoblingen (ingen ny matching-kjoring -
' bare kategorien oppdateres i tilfelle den er endret). For NYE
' meldinger: prov samtalekobling forst (arv raden til en tidligere sikkert
' matchet melding i samme Outlook-samtale), ellers kjor full matching
' (e-postadresse forst, gjenkjenningskolonner som fallback). Beregner til
' slutt Kontrollkolonne-status for alle rader med minst en kjent melding
' (per KTS_KATEGORI_PRIORITET), fra HELE cachen - ikke bare denne oktens
' funn.

Public Function KjørSkann() As Object
    On Error GoTo Feilhandtering

    Dim resultat As Object: Set resultat = CreateObject("Scripting.Dictionary")
    resultat("AntallSkannet") = 0
    resultat("AntallKategorisert") = 0
    resultat("AntallKjentFraCache") = 0
    resultat("AntallNyeMatchet") = 0
    resultat("AntallMatchetViaSamtale") = 0
    resultat("AntallOppdatertRad") = 0
    Dim konfig As Object: Set konfig = KtsLesKonfig()
    Dim postboksNavn As String
    If konfig.Exists("KTS_POSTBOKS") Then postboksNavn = konfig("KTS_POSTBOKS")
    If Len(postboksNavn) = 0 Then
        resultat("Feil") = "Ingen delt postboks er valgt. Åpne Innstillinger først."
        Set KjørSkann = resultat: Exit Function
    End If

    Dim epostKolonner() As String, gjenkjenningKolonner() As String
    epostKolonner = SplitEllerTom(IIf(konfig.Exists("KTS_KOL_EPOST"), konfig("KTS_KOL_EPOST"), ""))
    gjenkjenningKolonner = SplitEllerTom(IIf(konfig.Exists("KTS_KOL_GJENKJENNING"), konfig("KTS_KOL_GJENKJENNING"), ""))
    If ErTom(epostKolonner) And ErTom(gjenkjenningKolonner) Then
        resultat("Feil") = "Ingen matchingkolonner er valgt. Åpne Innstillinger først."
        Set KjørSkann = resultat: Exit Function
    End If

    Dim tbl As ListObject: Set tbl = KtsAktivTabell()
    If tbl Is Nothing Then resultat("Feil") = "Fant ingen Excel-tabell.": Set KjørSkann = resultat: Exit Function
    If tbl.DataBodyRange Is Nothing Then resultat("Feil") = "Tabellen har ingen datarader.": Set KjørSkann = resultat: Exit Function

    Dim kontrollKolNr As Long, kontrollKol As String
    If konfig.Exists("KTS_KOL_RESPONS") Then kontrollKol = konfig("KTS_KOL_RESPONS")
    kontrollKolNr = KtsKolonneNr(tbl, kontrollKol)
    If kontrollKolNr = 0 Then
        resultat("Feil") = "Kontrollkolonne er ikke valgt i Innstillinger."
        Set KjørSkann = resultat: Exit Function
    End If

    Dim startpunkt As Date, harStartpunkt As Boolean
    If konfig.Exists("KTS_STARTPUNKT") And Len(konfig("KTS_STARTPUNKT")) > 0 Then
        On Error Resume Next: startpunkt = CDate(konfig("KTS_STARTPUNKT")): harStartpunkt = (Err.Number = 0): Err.Clear: On Error GoTo Feilhandtering
    End If

    Dim olApp As Object, olNs As Object, deltPostboks As Object, innboks As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    Set olNs = olApp.GetNamespace("MAPI")
    Set deltPostboks = olNs.Folders.Item(postboksNavn)
    If Not deltPostboks Is Nothing Then
        Set innboks = deltPostboks.Folders.Item("Innboks")
        If innboks Is Nothing Then Set innboks = deltPostboks.Folders.Item("Inbox")
    End If
    On Error GoTo Feilhandtering
    If innboks Is Nothing Then resultat("Feil") = "Fant ikke innboksen i '" & postboksNavn & "'.": Set KjørSkann = resultat: Exit Function

    ' Lagre Outlooks kategorifarger (kosmetikk - feiler stille) sa dashbord
    ' og Kontrollkolonnen kan bruke de ekte fargene.
    KtsLagreKategoriFarger innboks, olNs

    ' --- Last inn persistent cache (kjente meldinger fra tidligere skanninger) ---
    Dim cache As Object: Set cache = KtsLastCache()

    ' --- Gjenoppbygg samtalekoblinger fra cachen, sa de er kjent pa
    '     tvers av okter - ikke bare innenfor denne ene skanningen. ---
    Dim samtaleTilRad As Object: Set samtaleTilRad = CreateObject("Scripting.Dictionary")
    Dim ck As Variant, cd As Object
    For Each ck In cache.Keys
        Set cd = cache(ck)
        If CLng(cd("RadNr")) > 0 And Len(CStr(cd("ConversationID"))) > 0 Then
            If Not samtaleTilRad.Exists(CStr(cd("ConversationID"))) Then samtaleTilRad(CStr(cd("ConversationID"))) = CLng(cd("RadNr"))
        End If
    Next ck

    ' --- Samle Innboks + ALLE undermapper rekursivt - en Outlook-regel kan
    '     ha sortert innkommende svar til undermapper (f.eks. "1. e-post
    '     til rektor" under selve Innboks), ikke bare rett i Innboks selv.
    '     Bekreftet 2026-09-19 som roten til at Håkons skanning fant 0
    '     kategoriserte rader - selve Innboks-mappa hadde nesten ingenting
    '     i seg, alt reelt innhold la i undermapper. ---
    Dim alleMapper As New Collection
    KtsSamleMapperRekursivt innboks, alleMapper

    Dim settDenneSkann As Object: Set settDenneSkann = CreateObject("Scripting.Dictionary")
    Dim mappe As Object
    For Each mappe In alleMapper
        Dim items As Object: Set items = Nothing
        If harStartpunkt Then
            On Error Resume Next
            Dim filterStr As String
            filterStr = "[ReceivedTime] >= '" & Format(startpunkt, "ddddd h:nn AMPM") & "'"
            Set items = mappe.Items.Restrict(filterStr)
            Err.Clear
            On Error GoTo Feilhandtering
        End If
        If items Is Nothing Then Set items = mappe.Items
        On Error Resume Next: items.Sort "[ReceivedTime]", False: On Error GoTo Feilhandtering

        Dim i As Long
        For i = 1 To items.Count
            Dim item As Object: Set item = Nothing
            On Error Resume Next: Set item = items.Item(i): On Error GoTo Feilhandtering
            If Not item Is Nothing Then
                resultat("AntallSkannet") = resultat("AntallSkannet") + 1

                Dim entryId As String: entryId = ""
                On Error Resume Next: entryId = CStr(item.EntryID): On Error GoTo Feilhandtering

                Dim kategorier As String: kategorier = ""
                On Error Resume Next: kategorier = Trim(CStr(item.Categories)): On Error GoTo Feilhandtering

                If Len(kategorier) > 0 Then
                    resultat("AntallKategorisert") = resultat("AntallKategorisert") + 1

                    Dim emne As String, avsender As String, mottattTekst As String, convId As String, storeId As String, mappeNavn As String
                    emne = "": mottattTekst = "": convId = "": storeId = "": mappeNavn = ""
                    On Error Resume Next
                    emne = item.Subject
                    mottattTekst = Format(item.ReceivedTime, "dd.mm.yyyy HH:mm")
                    convId = CStr(item.ConversationID)
                    storeId = mappe.StoreID
                    mappeNavn = CStr(mappe.Name)
                    On Error GoTo Feilhandtering
                    avsender = HentAvsenderEpost(item)

                    If Len(entryId) > 0 Then settDenneSkann(entryId) = True

                    Dim radNr As Long
                    Dim kjentRad As Long: kjentRad = 0
                    If Len(entryId) > 0 Then
                        If cache.Exists(entryId) Then kjentRad = CLng(cache(entryId)("RadNr"))
                    End If
                    If kjentRad > 0 Then
                        ' Kjent og matchet fra for - IKKE kjor matching pa nytt.
                        ' Gjenbruk raden fra forrige gang; bare kategorien kan
                        ' ha endret seg. (Tidligere umatchede e-poster provs
                        ' derimot pa nytt hver skann, i tilfelle tabellen er
                        ' rettet opp siden sist.)
                        resultat("AntallKjentFraCache") = resultat("AntallKjentFraCache") + 1
                        radNr = kjentRad
                    Else
                        If Len(convId) > 0 And samtaleTilRad.Exists(convId) Then
                            radNr = CLng(samtaleTilRad(convId))
                            If radNr > 0 Then resultat("AntallMatchetViaSamtale") = resultat("AntallMatchetViaSamtale") + 1
                        Else
                            radNr = FinnMatchendeRadNy(tbl, item, epostKolonner, gjenkjenningKolonner)
                            If radNr > 0 Then resultat("AntallNyeMatchet") = resultat("AntallNyeMatchet") + 1
                        End If
                        If radNr > 0 And Len(convId) > 0 And Not samtaleTilRad.Exists(convId) Then samtaleTilRad(convId) = radNr
                    End If

                    If radNr > 0 Then
                        Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
                        d("EntryID") = entryId: d("StoreID") = storeId: d("RadNr") = radNr
                        d("Kategori") = kategorier: d("Emne") = emne: d("Avsender") = avsender
                        d("Mottatt") = mottattTekst: d("ConversationID") = convId
                        d("Arsak") = "": d("Mappe") = mappeNavn
                        If cache.Exists(entryId) Then cache.Remove entryId
                        cache.Add entryId, d
                    ElseIf Len(entryId) > 0 Then
                        ' Ikke matchet - lagres i cachen med RadNr 0 (ingen sikker
                        ' match) eller -1 (flertydig), slik at uavklart-lista og
                        ' tallet pa knappen overlever at Excel lukkes.
                        Dim u As Object: Set u = CreateObject("Scripting.Dictionary")
                        u("EntryID") = entryId: u("StoreID") = storeId: u("RadNr") = radNr
                        u("Kategori") = kategorier: u("Emne") = emne: u("Avsender") = avsender
                        u("Mottatt") = mottattTekst: u("ConversationID") = convId
                        u("Mappe") = mappeNavn
                        If radNr = -1 Then u("Arsak") = "Flertydig match" & KtsSisteKandidatTekst Else u("Arsak") = "Ingen sikker match"
                        If cache.Exists(entryId) Then cache.Remove entryId
                        cache.Add entryId, u
                    End If
                End If
            End If
        Next i
    Next mappe

    ' --- Fjern uavklarte cache-oppforinger som IKKE ble sett i denne
    '     skanningen (e-posten er slettet/flyttet/utenfor startpunkt) -
    '     ellers ville de stat som uavklarte for alltid. ---
    Dim fjernListe As New Collection, fjernNokkel As Variant
    For Each ck In cache.Keys
        Set cd = cache(ck)
        If CLng(cd("RadNr")) <= 0 Then
            If Not settDenneSkann.Exists(CStr(ck)) Then fjernListe.Add CStr(ck)
        End If
    Next ck
    For Each fjernNokkel In fjernListe
        cache.Remove CStr(fjernNokkel)
    Next fjernNokkel

    ' --- Lagre oppdatert cache og gjenoppbygg de session-variablene som
    '     dashbord/detaljside leser, fra HELE cachen (inkl. uavklarte). ---
    KtsLagreCache cache
    KtsGjenoppbyggSessionFraCache cache
    resultat.Add "Uavklarte", KtsSisteUavklarte

    ' --- Beregn kontrollstatus for ALLE rader med minst en kjent melding
    '     (fra hele cachen, ikke bare denne skanningens funn). ---
    tbl.ListColumns(kontrollKolNr).DataBodyRange.ClearContents
    Dim key As Variant
    For Each key In KtsSisteMeldinger.Keys
        tbl.DataBodyRange.Cells(CLng(key), kontrollKolNr).Value = KtsKontrollstatusForRad(CLng(key), konfig)
        resultat("AntallOppdatertRad") = resultat("AntallOppdatertRad") + 1
    Next key
    KtsFormaterKontrollkolonne tbl, kontrollKolNr, KtsLesKonfig()

    Set KjørSkann = resultat
    Exit Function
Feilhandtering:
    Dim feilResultat As Object: Set feilResultat = CreateObject("Scripting.Dictionary")
    feilResultat("Feil") = "Uventet feil under synkronisering (VBA-feil " & Err.Number & ": " & Err.Description & ")."
    Set KjørSkann = feilResultat
End Function

' Returnerer radnummeret (1-basert) hvis NOYAKTIG en rad matcher, 0 hvis
' ingen rad matcher, -1 hvis flere ulike rader matcher (flertydig - skal
' IKKE gjettes, havner i uavklart-lista i stedet). E-postadresse proves
' FORST (sterkeste signal, Håkons eksplisitte prioritering 2026-09-18),
' gjenkjenningskolonner (emne+brodtekst) er fallback.
'
' E-postmatching sammenligner HELE, normaliserte adresser mot hverandre -
' bade e-post-cellen OG mailens Sender+CC-tekst splittes forst pa
' komma/semikolon (en celle KAN inneholde flere adresser). Aldri løs
' substring-matching her - det kunne tidligere feilaktig la
' "anna@x.no" treffe inni "hannah@x.no".
Private Function FinnMatchendeRadNy(ByVal tbl As ListObject, ByVal item As Object, epostKolonner() As String, gjenkjenningKolonner() As String) As Long
    ' Kombinerer signalene (endret 2026-09-24 etter at en e-post med baade
    ' kjent avsender-adresse OG skole-ID i emnet havnet i Uavklart): TIDLIGERE
    ' ga flere e-posttreff (samme kontakt/adresse paa flere rader, f.eks. ett
    ' per trinn) umiddelbart "flertydig" uten a bruke skole-ID/navn i det hele
    ' tatt, og gjenkjenningsverdiene ble bare brukt hvis e-post ga NULL treff.
    ' NA: 1) e-post gir kandidatene (unik rad = ferdig, som for). 2) Er det
    ' flere e-postkandidater, er gjenkjenningsverdiene TIE-BREAKER blant kun
    ' dem. 3) Ingen e-postkandidat: gjenkjenningsverdier over alle rader, der
    ' treff i EMNET teller tyngre enn treff bare i brodteksten. Flere rader
    ' likt paa toppen = fortsatt flertydig (aldri gjett).
    Dim r As Long, kolNavn As Variant, kolNr As Long
    Dim antRader As Long: antRader = tbl.DataBodyRange.Rows.Count
    Dim erKandidat() As Boolean: ReDim erKandidat(1 To antRader)
    Dim antKand As Long, forsteKand As Long
    KtsSisteKandidatTekst = ""

    If Not ErTom(epostKolonner) Then
        Dim mailAdresser As New Collection
        Dim a As Variant
        For Each a In KtsSplitAdresser(HentAvsenderEpost(item))
            mailAdresser.Add a
        Next a
        For Each a In KtsSplitAdresser(HentCcAdresser(item))
            mailAdresser.Add a
        Next a

        If mailAdresser.Count > 0 Then
            For Each kolNavn In epostKolonner
                kolNr = KtsKolonneNr(tbl, CStr(kolNavn))
                If kolNr > 0 Then
                    For r = 1 To antRader
                        If Not erKandidat(r) Then
                            Dim celleAdresser As Collection
                            Set celleAdresser = KtsSplitAdresser(CStr(tbl.DataBodyRange.Cells(r, kolNr).Value))
                            If KtsHarFellesAdresse(celleAdresser, mailAdresser) Then
                                erKandidat(r) = True
                                antKand = antKand + 1
                                If forsteKand = 0 Then forsteKand = r
                            End If
                        End If
                    Next r
                End If
            Next kolNavn
            If antKand = 1 Then FinnMatchendeRadNy = forsteKand: Exit Function
        End If
    End If

    ' Poengsum per rad fra gjenkjenningsverdiene. Har e-post gitt kandidater
    ' (antKand > 1), vurderes bare de; ellers alle rader.
    Dim emneTekst As String, altTekst As String
    On Error Resume Next
    emneTekst = CStr(item.Subject)
    altTekst = emneTekst & vbCrLf & CStr(item.Body)
    On Error GoTo 0

    Dim poeng() As Long: ReDim poeng(1 To antRader)
    Dim harGjenkjenning As Boolean
    If Not ErTom(gjenkjenningKolonner) Then
        For Each kolNavn In gjenkjenningKolonner
            kolNr = KtsKolonneNr(tbl, CStr(kolNavn))
            If kolNr > 0 Then
                For r = 1 To antRader
                    If antKand <= 1 Or erKandidat(r) Then
                        Dim v As String: v = Trim(CStr(tbl.DataBodyRange.Cells(r, kolNr).Value))
                        If Len(v) > 0 Then
                            If HeleOrdFinnes(emneTekst, v) Then
                                poeng(r) = poeng(r) + 10: harGjenkjenning = True
                            ElseIf HeleOrdFinnes(altTekst, v) Then
                                poeng(r) = poeng(r) + 1: harGjenkjenning = True
                            End If
                        End If
                    End If
                Next r
            End If
        Next kolNavn
    End If

    Dim best As Long, antBest As Long, bestRad As Long
    For r = 1 To antRader
        If antKand > 1 And Not erKandidat(r) Then GoTo NesteRad
        Dim p As Long: p = poeng(r)
        If antKand > 1 Then p = p + 1000
        If p > best Then
            best = p: antBest = 1: bestRad = r
        ElseIf p = best And p > 0 Then
            antBest = antBest + 1
        End If
NesteRad:
    Next r

    If best = 0 Then
        FinnMatchendeRadNy = 0
    ElseIf antBest = 1 Then
        FinnMatchendeRadNy = bestRad
    Else
        FinnMatchendeRadNy = -1
        Dim liste As String, tell As Long
        For r = 1 To antRader
            If antKand > 1 And Not erKandidat(r) Then GoTo NesteKand
            If poeng(r) + IIf(antKand > 1, 1000, 0) = best Then
                tell = tell + 1
                If tell <= 6 Then liste = liste & IIf(Len(liste) > 0, ", ", "") & tbl.DataBodyRange.Rows(r).Row
            End If
NesteKand:
        Next r
        If tell > 6 Then liste = liste & " ... (" & tell & " rader)"
        KtsSisteKandidatTekst = " (arkrader: " & liste & ")"
    End If
End Function

Private Function HentCcAdresser(ByVal item As Object) As String
    Dim s As String, rec As Object
    On Error Resume Next
    For Each rec In item.Recipients
        If rec.Type = 2 Then
            Dim a As String: a = rec.Address
            If rec.AddressEntry.Type = "EX" Then
                Dim exu As Object: Set exu = rec.AddressEntry.GetExchangeUser()
                If Not exu Is Nothing Then a = exu.PrimarySmtpAddress
            End If
            If Len(a) > 0 Then s = s & ";" & a
        End If
    Next rec
    On Error GoTo 0
    HentCcAdresser = s
End Function

Public Function KtsKontrollstatusForRad(ByVal rad As Long, ByVal konfig As Object) As String
    If KtsSisteMeldinger Is Nothing Then Exit Function
    If Not KtsSisteMeldinger.Exists(rad) Then Exit Function
    Dim prior As String: If konfig.Exists("KTS_KATEGORI_PRIORITET") Then prior = konfig("KTS_KATEGORI_PRIORITET")
    Dim best As String, bestP As Long: bestP = 999999
    Dim m As Variant, cat As Variant, p As Long
    For Each m In KtsSisteMeldinger(rad)
        For Each cat In KtsSplitKategorier(CStr(m("Kategori")))
            p = KtsKategoriPrioritet(CStr(cat), prior)
            If p < bestP Then bestP = p: best = CStr(cat)
        Next cat
    Next m
    KtsKontrollstatusForRad = best
End Function

Private Function KtsKategoriPrioritet(ByVal kategori As String, ByVal prior As String) As Long
    Dim arr As Variant, i As Long
    If Len(prior) > 0 Then
        arr = Split(prior, "|")
        For i = LBound(arr) To UBound(arr)
            If StrComp(Trim(CStr(arr(i))), kategori, vbTextCompare) = 0 Then KtsKategoriPrioritet = i: Exit Function
        Next i
    End If
    KtsKategoriPrioritet = 10000
End Function

' Outlooks 25 kategorifarger (OlCategoryColor 1-25) som omtrentlige RGB-
' verdier (samme fargetoner som i Outlook). 0/ukjent = ingen farge (-1).
Public Function KtsOutlookFargeRGB(ByVal idx As Long) As Long
    KtsOutlookFargeRGB = -1
    Select Case idx
        Case 1: KtsOutlookFargeRGB = RGB(231, 161, 162)
        Case 2: KtsOutlookFargeRGB = RGB(249, 186, 137)
        Case 3: KtsOutlookFargeRGB = RGB(247, 221, 143)
        Case 4: KtsOutlookFargeRGB = RGB(252, 250, 144)
        Case 5: KtsOutlookFargeRGB = RGB(120, 209, 104)
        Case 6: KtsOutlookFargeRGB = RGB(159, 220, 201)
        Case 7: KtsOutlookFargeRGB = RGB(198, 210, 176)
        Case 8: KtsOutlookFargeRGB = RGB(157, 183, 232)
        Case 9: KtsOutlookFargeRGB = RGB(181, 161, 226)
        Case 10: KtsOutlookFargeRGB = RGB(218, 174, 194)
        Case 11: KtsOutlookFargeRGB = RGB(218, 217, 220)
        Case 12: KtsOutlookFargeRGB = RGB(107, 121, 148)
        Case 13: KtsOutlookFargeRGB = RGB(191, 191, 191)
        Case 14: KtsOutlookFargeRGB = RGB(111, 111, 111)
        Case 15: KtsOutlookFargeRGB = RGB(79, 79, 79)
        Case 16: KtsOutlookFargeRGB = RGB(193, 26, 37)
        Case 17: KtsOutlookFargeRGB = RGB(226, 98, 13)
        Case 18: KtsOutlookFargeRGB = RGB(199, 153, 48)
        Case 19: KtsOutlookFargeRGB = RGB(185, 179, 0)
        Case 20: KtsOutlookFargeRGB = RGB(54, 143, 43)
        Case 21: KtsOutlookFargeRGB = RGB(50, 155, 122)
        Case 22: KtsOutlookFargeRGB = RGB(119, 139, 69)
        Case 23: KtsOutlookFargeRGB = RGB(40, 88, 165)
        Case 24: KtsOutlookFargeRGB = RGB(92, 63, 189)
        Case 25: KtsOutlookFargeRGB = RGB(147, 68, 107)
    End Select
End Function

' Leser kategori -> Outlook-fargeindeks fra konfig (KTS_KATEGORI_FARGER,
' "navn=indeks|navn=indeks"). Nokkelen er kategorinavnet i smabokstaver.
Public Function KtsLesKategoriFarger(ByVal konfig As Object) As Object
    Dim res As Object: Set res = CreateObject("Scripting.Dictionary")
    Dim s As String, del As Variant, p As Long
    If konfig.Exists("KTS_KATEGORI_FARGER") Then s = CStr(konfig("KTS_KATEGORI_FARGER"))
    If Len(s) > 0 Then
        For Each del In Split(s, "|")
            p = InStrRev(CStr(del), "=")
            If p > 1 Then res(LCase(Left(CStr(del), p - 1))) = CLng(Val(Mid(CStr(del), p + 1)))
        Next del
    End If
    Set KtsLesKategoriFarger = res
End Function

' RGB-farge for en kategori, eller -1 hvis den er ukjent/uten farge i Outlook.
Public Function KtsKategoriRGB(ByVal navn As String, ByVal farger As Object) As Long
    KtsKategoriRGB = -1
    If farger Is Nothing Then Exit Function
    If farger.Exists(LCase(navn)) Then KtsKategoriRGB = KtsOutlookFargeRGB(CLng(farger(LCase(navn))))
End Function

' Lesbar tekstfarge (morkt eller hvitt) mot en gitt bakgrunnsfarge.
Public Function KtsTekstFargeMot(ByVal bakgrunn As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = bakgrunn Mod 256: g = (bakgrunn \ 256) Mod 256: b = (bakgrunn \ 65536) Mod 256
    If (299 * r + 587 * g + 114 * b) / 1000 < 140 Then KtsTekstFargeMot = RGB(255, 255, 255) Else KtsTekstFargeMot = RGB(30, 30, 30)
End Function

' Leser fargene til kategoriene fra Outlooks egne kategorilister (postboksens
' egen forst, sa den innloggedes) og lagrer dem i konfig, slik at dashbordet
' og Kontrollkolonnen kan bruke de ekte Outlook-fargene uten a kontakte
' Outlook hver gang. Feiler stille - farger er kun kosmetikk.
Public Sub KtsLagreKategoriFarger(ByVal ib As Object, ByVal ns As Object)
    Dim s As String, kilde As Variant, katObj As Object, katEl As Object
    Dim navn As String, farge As Long, ren As String, tilg As String, fraStore As Boolean
    Dim sett As Object: Set sett = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each kilde In Array("store", "namespace")
        Set katObj = Nothing
        If kilde = "store" Then Set katObj = ib.Store.Categories Else Set katObj = ns.Categories
        If Not katObj Is Nothing Then
            For Each katEl In katObj
                navn = "": farge = 0
                navn = Trim(CStr(katEl.Name))
                farge = CLng(katEl.Color)
                ren = Replace(Replace(navn, "|", " "), "=", " ")
                If Len(ren) > 0 Then
                    If Not sett.Exists(LCase(ren)) Then
                        sett(LCase(ren)) = True
                        If Len(s) > 0 Then s = s & "|"
                        s = s & ren & "=" & farge
                        ' Lista over kategoriene radsiden tilbyr som knapper:
                        ' postboksens EGNE kategorier, aldri den personlige lista
                        ' hvis postboksen har noen (samme regel som Innstillinger).
                        If (kilde = "store" Or Not fraStore) And InStr(navn, "|") = 0 Then
                            If Len(tilg) > 0 Then tilg = tilg & "|"
                            tilg = tilg & navn
                            If kilde = "store" Then fraStore = True
                        End If
                    End If
                End If
            Next katEl
        End If
    Next kilde
    On Error GoTo 0
    If Len(s) > 0 Then KtsSkrivKonfigVerdi "KTS_KATEGORI_FARGER", s
    If Len(tilg) > 0 Then KtsSkrivKonfigVerdi "KTS_KATEGORI_TILGJENGELIG", tilg
End Sub

' Oppdaterer fargene og lista over tilgjengelige kategorier fra en Outlook
' som allerede kjorer (GetObject - starter aldri Outlook selv). Kalles nar
' radsiden apnes. Feiler stille: da brukes det som ble lagret sist.
Public Sub KtsOppdaterKategorilisteStille()
    Dim k As Object: Set k = KtsLesKonfig()
    If Not k.Exists("KTS_POSTBOKS") Then Exit Sub
    If Len(CStr(k("KTS_POSTBOKS"))) = 0 Then Exit Sub
    Dim olApp As Object, ns As Object, root As Object, ib As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Exit Sub
    Set ns = olApp.GetNamespace("MAPI")
    Set root = ns.Folders.Item(CStr(k("KTS_POSTBOKS")))
    If root Is Nothing Then Exit Sub
    Set ib = root.Folders.Item("Innboks")
    If ib Is Nothing Then Set ib = root.Folders.Item("Inbox")
    If ib Is Nothing Then Exit Sub
    KtsLagreKategoriFarger ib, ns
End Sub

' Kategoriene radsiden viser som knapper: postboksens egne (sist lagret fra
' Outlook) i den rekkefolgen brukeren har satt i Innstillinger, sa resten -
' pluss alle kategorier som allerede sitter pa en e-post for raden, slik at
' ogsa en kategori utenfor postboksens liste kan fjernes igjen. Uten lagret
' Outlook-liste brukes prioritetslista og det som er kjent fra cachen.
Public Function KtsKategoriValgForRad(ByVal rad As Long) As Collection
    Dim res As New Collection, sett As Object: Set sett = CreateObject("Scripting.Dictionary")
    Dim k As Object: Set k = KtsLesKonfig()
    Dim liste As New Collection, del As Variant, kat As Variant, m As Variant

    If k.Exists("KTS_KATEGORI_TILGJENGELIG") Then
        For Each del In Split(CStr(k("KTS_KATEGORI_TILGJENGELIG")), "|")
            If Len(Trim(CStr(del))) > 0 Then liste.Add Trim(CStr(del))
        Next del
    End If
    If liste.Count = 0 Then
        If k.Exists("KTS_KATEGORI_PRIORITET") Then
            For Each del In Split(CStr(k("KTS_KATEGORI_PRIORITET")), "|")
                If Len(Trim(CStr(del))) > 0 Then liste.Add Trim(CStr(del))
            Next del
        End If
        For Each kat In KtsHentKjenteKategorier()
            liste.Add CStr(kat)
        Next kat
    Else
        ' Innstillingenes rekkefolge forst (prioritetslista), sa resten.
        If k.Exists("KTS_KATEGORI_PRIORITET") Then
            For Each del In Split(CStr(k("KTS_KATEGORI_PRIORITET")), "|")
                For Each kat In liste
                    If StrComp(CStr(kat), Trim(CStr(del)), vbTextCompare) = 0 Then
                        If Not sett.Exists(LCase(CStr(kat))) Then sett(LCase(CStr(kat))) = True: res.Add CStr(kat)
                        Exit For
                    End If
                Next kat
            Next del
        End If
    End If
    For Each kat In liste
        If Not sett.Exists(LCase(CStr(kat))) Then sett(LCase(CStr(kat))) = True: res.Add CStr(kat)
    Next kat

    If Not KtsSisteMeldinger Is Nothing Then
        If KtsSisteMeldinger.Exists(rad) Then
            For Each m In KtsSisteMeldinger(rad)
                For Each kat In KtsSplitKategorier(CStr(m("Kategori")))
                    If Not sett.Exists(LCase(CStr(kat))) Then sett(LCase(CStr(kat))) = True: res.Add CStr(kat)
                Next kat
            Next m
        End If
    End If
    Set KtsKategoriValgForRad = res
End Function

' Windows' listeskilletegn (komma eller semikolon) - Outlook tolker
' MailItem.Categories med det skillet, sa et feil skilletegn kan gi EN
' kategori med komma i navnet i stedet for flere kategorier.
Private Function KtsListeSkille() As String
    Dim s As String
    On Error Resume Next
    s = CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\International\sList")
    On Error GoTo 0
    s = Trim(s)
    If s <> "," And s <> ";" Then s = ";"
    KtsListeSkille = s & " "
End Function

' Fargelegger Kontrollkolonnen med Outlooks EKTE kategorifarger (fra
' KTS_KATEGORI_FARGER). Ukjent kategori eller ennå ikke lest fargeliste blir
' noytral gra.
Private Sub KtsFormaterKontrollkolonne(ByVal tbl As ListObject, ByVal kolNr As Long, ByVal konfig As Object)
    Dim farger As Object: Set farger = KtsLesKategoriFarger(konfig)
    Dim r As Long, c As Range, status As String, farge As Long
    For r = 1 To tbl.DataBodyRange.Rows.Count
        Set c = tbl.DataBodyRange.Cells(r, kolNr)
        status = Trim(CStr(c.Value))
        c.Interior.Pattern = xlNone
        c.Font.ColorIndex = xlAutomatic
        If Len(status) > 0 Then
            farge = KtsKategoriRGB(status, farger)
            If farge = -1 Then
                c.Interior.Color = RGB(217, 217, 217)
            Else
                c.Interior.Color = farge
                c.Font.Color = KtsTekstFargeMot(farge)
            End If
        End If
    Next r
End Sub

Public Sub KtsAapneMelding(ByVal rad As Long, ByVal indeks As Long)
    On Error GoTo EH
    If KtsSisteMeldinger Is Nothing Then Exit Sub
    If Not KtsSisteMeldinger.Exists(rad) Then Exit Sub
    Dim info As Object: Set info = KtsSisteMeldinger(rad)(indeks)
    KtsAapneMeldingMedId CStr(info("EntryID")), CStr(info("StoreID"))
    Exit Sub
EH:
    MsgBox "Klarte ikke å åpne e-posten i Outlook: " & Err.Description, vbExclamation, "Kontaktsentralen"
End Sub

' Apner en e-post i Outlook ut fra EntryID/StoreID (brukes ogsa av
' uavklart-vinduet, der e-posten ikke hører til noen rad enda).
Public Sub KtsAapneMeldingMedId(ByVal entryId As String, ByVal storeId As String)
    On Error GoTo EH
    If Len(entryId) = 0 Then
        MsgBox "Denne e-posten mangler ID og kan ikke åpnes herfra.", vbExclamation, "Kontaktsentralen"
        Exit Sub
    End If
    Dim olApp As Object, olNs As Object, itm As Object
    On Error Resume Next: Set olApp = GetObject(, "Outlook.Application"): If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    Set olNs = olApp.GetNamespace("MAPI")
    Set itm = olNs.GetItemFromID(entryId, storeId)
    On Error GoTo EH
    If itm Is Nothing Then
        MsgBox "Fant ikke e-posten i Outlook (er den flyttet eller slettet?).", vbExclamation, "Kontaktsentralen"
        Exit Sub
    End If
    itm.Display
    Exit Sub
EH:
    MsgBox "Klarte ikke å åpne e-posten i Outlook: " & Err.Description, vbExclamation, "Kontaktsentralen"
End Sub

' Skriver en ny kategoriliste til EN e-post i Outlook (kategoriene er de
' brukeren har slatt av/pa som knapper pa radsiden) og oppdaterer cachen og
' Kontrollkolonnen etterpa. Returnerer kategoriene slik Outlook faktisk
' lagret dem (lest tilbake). Gikk noe galt er feilTekst fylt ut, og da er
' ingenting endret i cachen.
Public Function KtsLagreKategorierPaaMelding(ByVal rad As Long, ByVal indeks As Long, ByVal kategorier As Collection, ByRef feilTekst As String) As String
    On Error GoTo EH
    feilTekst = ""
    Dim info As Object: Set info = KtsSisteMeldinger(rad)(indeks)
    Dim olApp As Object, olNs As Object, itm As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    Set olNs = olApp.GetNamespace("MAPI")
    Set itm = olNs.GetItemFromID(CStr(info("EntryID")), CStr(info("StoreID")))
    On Error GoTo EH
    If itm Is Nothing Then
        feilTekst = "Fant ikke e-posten i Outlook (er den flyttet eller slettet?)."
        Exit Function
    End If

    Dim ny As String, kat As Variant, skille As String: skille = KtsListeSkille()
    For Each kat In kategorier
        If Len(ny) > 0 Then ny = ny & skille
        ny = ny & CStr(kat)
    Next kat
    itm.Categories = ny
    itm.Save

    ' Les tilbake - Outlook er fasit pa hva som faktisk ble lagret.
    Dim lagret As String: lagret = Trim(CStr(itm.Categories))
    KtsOppdaterEtterKategoriendring rad, indeks, lagret
    KtsLagreKategorierPaaMelding = lagret
    Exit Function
EH:
    feilTekst = "Klarte ikke å endre kategorier: " & Err.Description
End Function

' Oppdaterer cachen, sesjonslistene og Kontrollkolonnen etter at en e-posts
' kategorier er endret. Cachen ma med - ellers forsvinner en manuell endring
' igjen ved neste Excel-restart (samme prinsipp som KtsMatchUavklartTilRad).
Public Sub KtsOppdaterEtterKategoriendring(ByVal rad As Long, ByVal indeks As Long, ByVal kategorier As String)
    Dim info As Object: Set info = KtsSisteMeldinger(rad)(indeks)
    Dim entryIdNokkel As String: entryIdNokkel = CStr(info("EntryID"))
    If Len(entryIdNokkel) > 0 Then
        Dim cache As Object: Set cache = KtsLastCache()
        If cache.Exists(entryIdNokkel) Then
            cache(entryIdNokkel)("Kategori") = kategorier
        Else
            Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
            d("EntryID") = entryIdNokkel
            d("StoreID") = CStr(info("StoreID"))
            d("RadNr") = rad
            d("Kategori") = kategorier
            d("Emne") = CStr(info("Emne"))
            d("Avsender") = CStr(info("Avsender"))
            d("Mottatt") = CStr(info("Mottatt"))
            d("ConversationID") = ""
            d("Arsak") = ""
            If info.Exists("Mappe") Then d("Mappe") = CStr(info("Mappe")) Else d("Mappe") = ""
            cache.Add entryIdNokkel, d
        End If
        KtsLagreCache cache
        KtsGjenoppbyggSessionFraCache cache
    Else
        info("Kategori") = kategorier
    End If

    Dim konfig As Object: Set konfig = KtsLesKonfig()
    Dim tbl As ListObject: Set tbl = KtsAktivTabell()
    Dim kolNr As Long
    If Not tbl Is Nothing And konfig.Exists("KTS_KOL_RESPONS") Then kolNr = KtsKolonneNr(tbl, konfig("KTS_KOL_RESPONS"))
    If kolNr > 0 Then
        tbl.DataBodyRange.Cells(rad, kolNr).Value = KtsKontrollstatusForRad(rad, konfig)
        KtsFormaterKontrollkolonne tbl, kolNr, konfig
    End If
End Sub

' Live fallback-skann av selve postboksen for a oppdage hvilke kategorier
' som faktisk er i bruk - bare nodvendig for FORSTE gangs oppsett, for
' cachen har noe a vise (se KtsHentKjenteKategorier). Respekterer
' startpunkt og bruker Restrict der det er satt, for a unnga a lese hvert
' eneste element i en postboks med lang historikk unodvendig.
Public Function KtsHentKategorierFraPostboks(ByVal postboksNavn As String) As Collection
    Dim res As New Collection, sett As Object: Set sett = CreateObject("Scripting.Dictionary")
    On Error GoTo Ferdig
    Dim olApp As Object, ns As Object, root As Object, ib As Object, it As Object, c As Variant, i As Long
    On Error Resume Next: Set olApp = GetObject(, "Outlook.Application"): If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    Set ns = olApp.GetNamespace("MAPI"): Set root = ns.Folders.Item(postboksNavn): Set ib = root.Folders.Item("Innboks"): If ib Is Nothing Then Set ib = root.Folders.Item("Inbox")
    On Error GoTo Ferdig
    If ib Is Nothing Then GoTo Ferdig

    KtsLagreKategoriFarger ib, ns

    ' FORST postboksens EGEN kategoriliste (masterlista for akkurat denne
    ' delte postboksen), sa vises alle tilgjengelige TIMSS-kategorier, ikke
    ' bare de som tilfeldigvis er i bruk pa skannede e-poster. Brukerens
    ' PERSONLIGE liste brukes KUN som nodlosning hvis postboksen ikke
    ' eksponerer noen egen liste - ellers ble Håkons private kategorier
    ' blandet inn (2026-09-19). Leses defensivt - ikke alle Outlook-
    ' versjoner/lagre eksponerer Categories.
    Dim katKilde As Variant, katObj As Object, katEl As Object
    For Each katKilde In Array("store", "namespace")
        If katKilde = "namespace" And res.Count > 0 Then Exit For
        Set katObj = Nothing
        On Error Resume Next
        If katKilde = "store" Then Set katObj = ib.Store.Categories Else Set katObj = ns.Categories
        On Error GoTo Ferdig
        If Not katObj Is Nothing Then
            On Error Resume Next
            For Each katEl In katObj
                Dim katNavn As String: katNavn = ""
                katNavn = Trim(CStr(katEl.Name))
                If Len(katNavn) > 0 Then
                    If Not sett.Exists(LCase(katNavn)) Then sett(LCase(katNavn)) = True: res.Add katNavn
                End If
            Next katEl
            On Error GoTo Ferdig
        End If
    Next katKilde

    Dim konfig As Object: Set konfig = KtsLesKonfig()
    Dim startpunkt As Date, harStartpunkt As Boolean
    If konfig.Exists("KTS_STARTPUNKT") And Len(konfig("KTS_STARTPUNKT")) > 0 Then
        On Error Resume Next: startpunkt = CDate(konfig("KTS_STARTPUNKT")): harStartpunkt = (Err.Number = 0): Err.Clear: On Error GoTo Ferdig
    End If

    ' Samme rekursive undermappe-gjennomgang som selve skannemotoren (se
    ' KtsSamleMapperRekursivt) - ellers finner ikke Oppsett kategoriene
    ' hvis en Outlook-regel har sortert mail til undermapper av Innboks.
    Dim alleMapper As New Collection
    KtsSamleMapperRekursivt ib, alleMapper

    Dim mappe As Object
    For Each mappe In alleMapper
        Dim items As Object: Set items = mappe.Items
        If harStartpunkt Then
            On Error Resume Next
            Dim filterStr As String: filterStr = "[ReceivedTime] >= '" & Format(startpunkt, "ddddd h:nn AMPM") & "'"
            Dim restr As Object: Set restr = items.Restrict(filterStr)
            If Not restr Is Nothing Then Set items = restr
            Err.Clear
            On Error GoTo Ferdig
        End If

        For i = 1 To items.Count
            Set it = Nothing: On Error Resume Next: Set it = items.Item(i): On Error GoTo Ferdig
            If Not it Is Nothing Then
                Dim cs As String: On Error Resume Next: cs = CStr(it.Categories): On Error GoTo Ferdig
                If Len(Trim(cs)) > 0 Then
                    For Each c In KtsSplitKategorier(cs)
                        If Not sett.Exists(LCase(CStr(c))) Then sett(LCase(CStr(c))) = True: res.Add CStr(c)
                    Next c
                End If
            End If
        Next i
    Next mappe
Ferdig:
    Set KtsHentKategorierFraPostboks = res
End Function

' Rask, Outlook-frie kategori-oppdagelse fra den PERSISTENTE cachen (se
' KtsLastCache) - normalveien etter forste skanning, sa Oppsett ikke
' trenger a skanne hele innboksen pa nytt bare for a fylle prioritetslista.
Public Function KtsHentKjenteKategorier() As Collection
    Dim res As New Collection, sett As Object: Set sett = CreateObject("Scripting.Dictionary")
    Dim cache As Object: Set cache = KtsLastCache()
    Dim key As Variant, d As Object, kat As Variant
    For Each key In cache.Keys
        Set d = cache(key)
        For Each kat In KtsSplitKategorier(CStr(d("Kategori")))
            If Not sett.Exists(LCase(CStr(kat))) Then
                sett(LCase(CStr(kat))) = True
                res.Add CStr(kat)
            End If
        Next kat
    Next key
    Set KtsHentKjenteKategorier = res
End Function

' Sentral hjelpefunksjon for a splitte MailItem.Categories robust.
' Separatoren er IKKE palitelig alltid komma - norske Windows-oppsett
' bruker ofte semikolon som listeseparator (komma er desimaltegn i
' norsk locale). Bekreftet 2026-09-17 som roten til at 0.6.2 viste null
' kategoriserte rader for Håkon. Splitter derfor pa BADE komma og
' semikolon, uansett locale - kategorinavn inneholder normalt ingen av
' delene, sa dette er trygt.
Public Function KtsSplitKategorier(ByVal raw As String) As Collection
    Dim resultat As New Collection
    If Len(Trim(raw)) = 0 Then
        Set KtsSplitKategorier = resultat
        Exit Function
    End If
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "[,;]+"
    Dim normalisert As String
    normalisert = re.Replace(raw, Chr(1))
    Dim deler() As String
    deler = Split(normalisert, Chr(1))
    Dim i As Long, t As String
    For i = LBound(deler) To UBound(deler)
        t = Trim(deler(i))
        If Len(t) > 0 Then resultat.Add t
    Next i
    Set KtsSplitKategorier = resultat
End Function

' Splitter en fritekst med flere e-postadresser (fra en Excel-celle eller
' fra Sender+CC-tekst) robust pa komma ELLER semikolon, trimmer og
' small-caser hver adresse. Brukes for a sammenligne HELE adresser -
' aldri løs substring-matching (en tidligere versjon kunne feilaktig
' matche "anna@x.no" inni "hannah@x.no" via InStr).
Private Function KtsSplitAdresser(ByVal raw As String) As Collection
    Dim resultat As New Collection
    If Len(Trim(raw)) = 0 Then
        Set KtsSplitAdresser = resultat
        Exit Function
    End If
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "[,;]+"
    Dim normalisert As String
    normalisert = re.Replace(raw, Chr(1))
    Dim deler() As String
    deler = Split(normalisert, Chr(1))
    Dim i As Long, t As String
    For i = LBound(deler) To UBound(deler)
        t = LCase(Trim(deler(i)))
        If Len(t) > 0 Then resultat.Add t
    Next i
    Set KtsSplitAdresser = resultat
End Function

Private Function KtsHarFellesAdresse(ByVal a As Collection, ByVal b As Collection) As Boolean
    Dim x As Variant, y As Variant
    For Each x In a
        For Each y In b
            If x = y Then
                KtsHarFellesAdresse = True
                Exit Function
            End If
        Next y
    Next x
End Function

' Henter avsenders SMTP-adresse robust - for interne Exchange-avsendere
' er SenderEmailAddress ofte en X.500-katalogsti (SenderEmailType="EX"),
' ikke en vanlig e-postadresse, og ma losES via GetExchangeUser().
Private Function HentAvsenderEpost(ByVal item As Object) As String
    Dim resultat As String
    On Error Resume Next
    If item.SenderEmailType = "EX" Then
        Dim exUser As Object
        Set exUser = item.Sender.GetExchangeUser()
        If Not exUser Is Nothing Then resultat = exUser.PrimarySmtpAddress
    Else
        resultat = item.SenderEmailAddress
    End If
    On Error GoTo 0
    HentAvsenderEpost = resultat
End Function

' Helords-treff (ikke bare substring) via regex - unngar f.eks. at
' skole-ID "10" matcher inni "1083". Sen-bundet VBScript.RegExp brukt
' i stedet for en prosjekt-referanse, sa dette virker uten at brukeren
' matte aktivere noen ekstra referanse i VBA-editoren.
Private Function HeleOrdFinnes(ByVal tekst As String, ByVal ord As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "(?:^|[^A-Za-z0-9æøåÆØÅ])" & RegexEscape(ord) & "(?:$|[^A-Za-z0-9æøåÆØÅ])"
    HeleOrdFinnes = re.Test(tekst)
End Function

Private Function RegexEscape(ByVal s As String) As String
    Dim tegn As Variant, resultat As String
    resultat = s
    For Each tegn In Array(".", "\", "+", "*", "?", "[", "^", "]", "$", "(", ")", "{", "}", "=", "!", "<", ">", "|", ":", "-", "#")
        resultat = Replace(resultat, tegn, "\" & tegn)
    Next tegn
    RegexEscape = resultat
End Function

Private Function SplitEllerTom(ByVal verdi As String) As String()
    ' NB: "ReDim x(0 To -1)" ser ut som et vanlig VBA-triks for a lage et
    ' tomt array, men feiler faktisk ALLTID med "Subscript out of range"
    ' (bekreftet empirisk 2026-09-11 - se lardom i minnet). Riktig mate a
    ' representere "tomt" pa er et UDIMENSJONERT dynamisk array (aldri
    ' ReDim'et) - derfor bare returnerer vi tomtResultat urort her.
    Dim tomtResultat() As String
    If Len(verdi) = 0 Then
        SplitEllerTom = tomtResultat
    Else
        SplitEllerTom = Split(verdi, "|")
    End If
End Function

Private Function ErTom(arr() As String) As Boolean
    ' Kan ikke bruke UBound/LBound direkte - de feiler pa et udimensjonert
    ' array. On Error Resume Next + en forhandsvurdert sentinelverdi (-1)
    ' er standardmonsteret i VBA for a sjekke om et dynamisk array i det
    ' hele tatt er allokert.
    Dim n As Long
    n = -1
    On Error Resume Next
    n = UBound(arr)
    On Error GoTo 0
    ErTom = (n < 0)
End Function

' --- Dashbord-hjelpefunksjoner ---
' NB: ingen array-parametere er noensinne ByVal her - se lardom om
' "ByVal pa array-parameter kompilerer ikke i VBA" i minnet.

' Kolonnene (navn) som vises for a gjenkjenne rader i listene - valgt i
' Innstillinger (KTS_KOL_VISNING, pipe-skilt, i valgt rekkefolge, maks 4).
' Uten gyldig valg: forste gjenkjenningskolonne, ellers forste e-postkolonne,
' ellers tabellens forste kolonne.
Public Function KtsHentVisningsKolonner(ByVal tbl As ListObject, ByVal konfig As Object) As Collection
    Dim res As New Collection
    Dim s As String, del As Variant, navn As String, kand As Variant
    If konfig.Exists("KTS_KOL_VISNING") Then s = CStr(konfig("KTS_KOL_VISNING"))
    If Len(s) > 0 Then
        For Each del In Split(s, "|")
            navn = Trim(CStr(del))
            If Len(navn) > 0 And res.Count < 4 Then
                If KtsKolonneNr(tbl, navn) > 0 Then res.Add navn
            End If
        Next del
    End If
    If res.Count = 0 Then
        For Each kand In Array("KTS_KOL_GJENKJENNING", "KTS_KOL_EPOST")
            If res.Count = 0 And konfig.Exists(CStr(kand)) Then
                s = Split(CStr(konfig(CStr(kand))) & "|", "|")(0)
                If Len(s) > 0 Then
                    If KtsKolonneNr(tbl, s) > 0 Then res.Add s
                End If
            End If
        Next kand
    End If
    If res.Count = 0 Then
        If tbl.ListColumns.Count > 0 Then res.Add tbl.ListColumns(1).Name
    End If
    Set KtsHentVisningsKolonner = res
End Function

' Overskrifter til kolonnene i radlista: visningskolonnene, sa Status og
' E-poster (samme rekkefolge som KtsByggRadliste fyller dem).
Public Function KtsListeOverskrifter() As Collection
    Dim res As New Collection
    Dim tbl As ListObject: Set tbl = KtsAktivTabell()
    Dim navn As Variant
    If Not tbl Is Nothing Then
        For Each navn In KtsHentVisningsKolonner(tbl, KtsLesKonfig())
            res.Add CStr(navn)
        Next navn
    End If
    res.Add "Status"
    res.Add "E-poster"
    Set KtsListeOverskrifter = res
End Function

' Lesbart navn pa en rad = verdiene i visningskolonnene skilt med " - " -
' brukes i tittelen pa radsiden i stedet for bare et radnummer.
Public Function KtsRadNavnForVisning(ByVal rad As Long) As String
    Dim tbl As ListObject: Set tbl = KtsAktivTabell()
    If tbl Is Nothing Then KtsRadNavnForVisning = "Rad " & rad: Exit Function
    If tbl.DataBodyRange Is Nothing Then KtsRadNavnForVisning = "Rad " & rad: Exit Function
    If rad < 1 Or rad > tbl.DataBodyRange.Rows.Count Then KtsRadNavnForVisning = "Rad " & rad: Exit Function
    Dim vis As Collection: Set vis = KtsHentVisningsKolonner(tbl, KtsLesKonfig())
    Dim navn As Variant, verdi As String, s As String
    For Each navn In vis
        verdi = Trim(CStr(tbl.DataBodyRange.Cells(rad, KtsKolonneNr(tbl, CStr(navn))).Value))
        If Len(verdi) > 0 Then
            If Len(s) > 0 Then s = s & " - "
            s = s & verdi
        End If
    Next navn
    If Len(s) = 0 Then s = "Rad " & rad
    KtsRadNavnForVisning = s
End Function

' Henter ALLE rader (Håkons eksplisitte valg 2026-09-18 - "alle rader kan
' vises, for alle rader skal kunne ha en oversikt"), sortert med
' kategoriserte rader forst (etter KTS_KATEGORI_PRIORITET), rader uten
' kategori-status sist. Fyller parallelle, 1-baserte arrayer: radArr = tabellens
' radnummer, visArr(rad, kolonne) = visningskolonnene (1-4), statusArr,
' meldingerArr og flereArr = antall ANDRE kategorier raden har utover status.
' maxLen(kolonne) = lengste tekst i visningskolonnen (til breddeberegning).
' Er tabellen tom er antallRader 0 og arrayene er urorte.
' nokkelArr = radens radnokkel (for a oppdage at tabellen er sortert/endret).
Public Sub KtsHentRadData(radArr() As Long, visArr() As String, statusArr() As String, meldingerArr() As String, flereArr() As Long, maxLen() As Long, nokkelArr() As String, ByRef nVis As Long, ByRef antallRader As Long)
    nVis = 0: antallRader = 0
    Dim konfig As Object
    Set konfig = KtsLesKonfig()
    Dim tbl As ListObject
    Set tbl = KtsAktivTabell()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Dim vis As Collection: Set vis = KtsHentVisningsKolonner(tbl, konfig)
    nVis = vis.Count
    ' Kolonnene leses som hele arrayer (raskt nok til a bygge lista pa nytt for
    ' hvert tastetrykk i sokefeltet).
    Dim kolVerdier() As Variant
    ReDim kolVerdier(1 To nVis)
    ReDim maxLen(1 To nVis)
    Dim jj As Long
    For jj = 1 To nVis
        kolVerdier(jj) = KtsNokkelVerdierForKolonne(tbl, CStr(vis(jj)))
        maxLen(jj) = Len(CStr(vis(jj)))
    Next jj

    Dim statusKol As Variant, kontrollNavn As String
    If konfig.Exists("KTS_KOL_RESPONS") Then kontrollNavn = CStr(konfig("KTS_KOL_RESPONS"))
    statusKol = KtsNokkelVerdierForKolonne(tbl, kontrollNavn)
    Dim nokkelKol As Variant
    nokkelKol = KtsNokkelVerdierForKolonne(tbl, KtsNokkelKolonneNavn(tbl, konfig))

    Dim prior As String
    If konfig.Exists("KTS_KATEGORI_PRIORITET") Then prior = konfig("KTS_KATEGORI_PRIORITET")

    antallRader = tbl.DataBodyRange.Rows.Count

    Dim prioritetArr() As Long
    ReDim radArr(1 To antallRader)
    ReDim statusArr(1 To antallRader)
    ReDim meldingerArr(1 To antallRader)
    ReDim prioritetArr(1 To antallRader)
    ReDim flereArr(1 To antallRader)
    ReDim nokkelArr(1 To antallRader)
    ReDim visArr(1 To antallRader, 1 To nVis)

    Dim r As Long, status As String, tekstVerdi As String, harTekst As Boolean
    For r = 1 To antallRader
        status = ""
        If Not IsEmpty(statusKol) Then status = statusKol(r)
        If Not IsEmpty(nokkelKol) Then nokkelArr(r) = nokkelKol(r)

        harTekst = False
        For jj = 1 To nVis
            tekstVerdi = ""
            If Not IsEmpty(kolVerdier(jj)) Then tekstVerdi = kolVerdier(jj)(r)
            visArr(r, jj) = tekstVerdi
            If Len(tekstVerdi) > 0 Then harTekst = True
            If Len(tekstVerdi) > maxLen(jj) Then maxLen(jj) = Len(tekstVerdi)
        Next jj
        If Not harTekst Then visArr(r, 1) = "Rad " & r

        radArr(r) = r
        If Len(status) > 0 Then
            statusArr(r) = status
            prioritetArr(r) = KtsKategoriPrioritet(status, prior)
            flereArr(r) = KtsAntallFlereKategorier(r, status)
        Else
            statusArr(r) = "(Ingen kategori ennå)"
            prioritetArr(r) = 1000000
        End If

        meldingerArr(r) = ""
        If Not KtsSisteMeldinger Is Nothing Then
            If KtsSisteMeldinger.Exists(r) Then meldingerArr(r) = KtsSisteMeldinger(r).Count & " melding(er)"
        End If
    Next r

    ' --- Enkel boble-sortering (stabil): lavest prioritetstall (= hoyest
    '     kategori-prioritet) forst, "ingen kategori" alltid sist. ---
    Dim i As Long, j As Long, tmpL As Long, tmpS As String
    For i = 1 To antallRader - 1
        For j = 1 To antallRader - i
            If prioritetArr(j) > prioritetArr(j + 1) Then
                tmpL = prioritetArr(j): prioritetArr(j) = prioritetArr(j + 1): prioritetArr(j + 1) = tmpL
                tmpL = radArr(j): radArr(j) = radArr(j + 1): radArr(j + 1) = tmpL
                tmpL = flereArr(j): flereArr(j) = flereArr(j + 1): flereArr(j + 1) = tmpL
                tmpS = statusArr(j): statusArr(j) = statusArr(j + 1): statusArr(j + 1) = tmpS
                tmpS = meldingerArr(j): meldingerArr(j) = meldingerArr(j + 1): meldingerArr(j + 1) = tmpS
                tmpS = nokkelArr(j): nokkelArr(j) = nokkelArr(j + 1): nokkelArr(j + 1) = tmpS
                For jj = 1 To nVis
                    tmpS = visArr(j, jj): visArr(j, jj) = visArr(j + 1, jj): visArr(j + 1, jj) = tmpS
                Next jj
            End If
        Next j
    Next i
End Sub

' Antall ulike kategorier pa radens e-poster utover selve statuskategorien.
Private Function KtsAntallFlereKategorier(ByVal rad As Long, ByVal status As String) As Long
    If KtsSisteMeldinger Is Nothing Then Exit Function
    If Not KtsSisteMeldinger.Exists(rad) Then Exit Function
    Dim sett As Object: Set sett = CreateObject("Scripting.Dictionary")
    Dim m As Variant, kat As Variant
    For Each m In KtsSisteMeldinger(rad)
        For Each kat In KtsSplitKategorier(CStr(m("Kategori")))
            If StrComp(CStr(kat), status, vbTextCompare) <> 0 Then sett(LCase(CStr(kat))) = True
        Next kat
    Next m
    KtsAntallFlereKategorier = sett.Count
End Function

' Kolonnebredder ("NN pt;NN pt;...") tilpasset innholdet (ca. 5,5 pt per tegn):
' visningskolonnene, sa Status og E-poster. Skaleres ned hvis de ikke far
' plass innenfor totalBredde (punkter).
Public Function KtsKolonneBredder(maxLen() As Long, ByVal nVis As Long, ByVal totalBredde As Single) As String
    Const STATUS_BREDDE As Single = 130
    Const MELDINGER_BREDDE As Single = 85
    Dim tilgjengelig As Single: tilgjengelig = totalBredde - STATUS_BREDDE - MELDINGER_BREDDE - 24
    Dim onsket() As Single, sumOnsket As Single, faktor As Single, bredde As Single, jj As Long
    ReDim onsket(1 To nVis)
    For jj = 1 To nVis
        bredde = maxLen(jj) * 5.5 + 12
        If bredde < 45 Then bredde = 45
        If bredde > 230 Then bredde = 230
        onsket(jj) = bredde
        sumOnsket = sumOnsket + bredde
    Next jj
    faktor = 1
    If sumOnsket > tilgjengelig And sumOnsket > 0 Then faktor = tilgjengelig / sumOnsket
    Dim breddeTekst As String
    For jj = 1 To nVis
        bredde = Int(onsket(jj) * faktor)
        If bredde < 30 Then bredde = 30
        breddeTekst = breddeTekst & bredde & " pt;"
    Next jj
    KtsKolonneBredder = breddeTekst & Int(STATUS_BREDDE) & " pt;" & Int(MELDINGER_BREDDE) & " pt"
End Function

' Fyller en (vanlig) ListBox med alle rader - brukt av uavklart-vinduet.
' radNumre() fylles med de underliggende tabell-radnumrene i SAMME
' rekkefolge som lista, slik at valgt listeindeks kan oversettes til riktig
' rad senere. Kolonnene er: valgte visningskolonner (1-4), Status, E-poster.
' totalBredde er listens bredde i punkter.
Public Sub KtsByggRadliste(ByVal lst As Object, radNumre() As Long, Optional ByVal totalBredde As Single = 740)
    lst.Clear
    On Error Resume Next
    Erase radNumre
    On Error GoTo 0

    Dim radArr() As Long, flereArr() As Long, maxLen() As Long
    Dim visArr() As String, statusArr() As String, meldingerArr() As String, nokkelArr() As String
    Dim nVis As Long, antallRader As Long
    KtsHentRadData radArr, visArr, statusArr, meldingerArr, flereArr, maxLen, nokkelArr, nVis, antallRader
    If nVis = 0 Then Exit Sub

    lst.ColumnCount = nVis + 2
    lst.ColumnWidths = KtsKolonneBredder(maxLen, nVis, totalBredde)

    ' Se lardom i minnet - "ReDim x(0 To -1)" feiler, sa nar antallRader=0
    ' lar vi radNumre bli staende udimensjonert (tomt) i stedet.
    If antallRader > 0 Then
        ReDim radNumre(0 To antallRader - 1)
    End If
    Dim i As Long, jj As Long
    For i = 1 To antallRader
        lst.AddItem visArr(i, 1)
        For jj = 2 To nVis
            lst.List(lst.ListCount - 1, jj - 1) = visArr(i, jj)
        Next jj
        lst.List(lst.ListCount - 1, nVis) = statusArr(i)
        lst.List(lst.ListCount - 1, nVis + 1) = meldingerArr(i)
        radNumre(i - 1) = radArr(i)
    Next i
End Sub

' Markerer radens forste celle i tabellen og blar slik at raden er synlig -
' UTEN a ta fokus fra dashbord-vinduet (det er modeless og skal kunne
' styres videre med piltaster). Returnerer False hvis raden ikke kan vises
' (skjult av et filter, ugyldig radnummer eller ingen tabell).
Public Function KtsGaTilRadStille(ByVal rad As Long) As Boolean
    Dim tbl As ListObject
    Set tbl = KtsAktivTabell()
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If rad < 1 Or rad > tbl.DataBodyRange.Rows.Count Then Exit Function
    On Error Resume Next
    If tbl.DataBodyRange.Rows(rad).EntireRow.Hidden Then Exit Function
    tbl.Parent.Parent.Activate
    tbl.Parent.Activate
    ' .Select (ikke Application.Goto med Scroll:=True): Excel ruller da bare
    ' akkurat nok til at cellen blir synlig - og ikke i det hele tatt hvis
    ' raden allerede er i bildet - i stedet for a tvinge raden opp i ovre
    ' venstre hjorne hver gang (samme losning som Mail-utsender, GaTilFunn).
    tbl.DataBodyRange.Cells(rad, 1).Select
    KtsGaTilRadStille = (Err.Number = 0)
End Function

' Er skjemaet med dette navnet lastet (apent) akkurat na?
Public Function KtsSkjemaErApent(ByVal navn As String) As Boolean
    Dim f As Object
    For Each f In VBA.UserForms
        If f.Name = navn Then
            KtsSkjemaErApent = True
            Exit Function
        End If
    Next f
End Function

' Antall uavklarte e-poster (fra cachen / siste skann) - vises pa knappen.
Public Function KtsAntallUavklarte() As Long
    If KtsSisteUavklarte Is Nothing Then Exit Function
    KtsAntallUavklarte = KtsSisteUavklarte.Count
End Function

' Tilpasser vinduets storrelse til kontrollene i det, uavhengig av hvor
' tykk tittellinje/ramme Windows tegner (varierer med skalering og tema) -
' faste hoyder gjorde at knappene nederst ble kuttet.
Public Sub KtsTilpassVindu(ByVal frm As Object)
    Dim c As Object, b As Single, r As Single
    On Error Resume Next
    For Each c In frm.Controls
        If c.Visible Then
            If c.Top + c.Height > b Then b = c.Top + c.Height
            If c.Left + c.Width > r Then r = c.Left + c.Width
        End If
    Next c
    frm.Height = frm.Height - frm.InsideHeight + b + 14
    frm.Width = frm.Width - frm.InsideWidth + r + 14
End Sub

' Farget knapp (hvit fet tekst). Fargenavn: gronn, blagra, oransje, bla, gra.
Public Sub KtsStilKnapp(ByVal knapp As Object, ByVal fargenavn As String)
    Dim farge As Long
    Select Case fargenavn
        Case "gronn": farge = RGB(46, 139, 87)
        Case "blagra": farge = RGB(91, 107, 138)
        Case "oransje": farge = RGB(217, 130, 43)
        Case "bla": farge = RGB(58, 114, 196)
        Case Else: farge = RGB(138, 138, 132)
    End Select
    On Error Resume Next
    knapp.BackColor = farge
    knapp.ForeColor = RGB(255, 255, 255)
    knapp.Font.Bold = True
End Sub

' Matcher et uavklart treff (fra siste skanning, se KtsSisteUavklarte)
' manuelt til en valgt rad - skriver Respons-/Oppfolging-kolonnene som
' om skanningen selv hadde funnet treffet, og fjerner det deretter fra
' uavklart-lista. uavklartIndeks er 1-basert (Collection-indeksering).
' Matcher et uavklart treff manuelt til en valgt rad. Lagrer det INN I
' den persistente cachen (samme mekanisme som et automatisk treff), sa
' korreksjonen huskes pa tvers av okter og fremtidige skanninger - en
' tidligere versjon la dette bare i en session-only variabel, som gjorde
' at en manuell korreksjon forsvant igjen ved neste Excel-restart.
Public Sub KtsMatchUavklartTilRad(ByVal uavklartIndeks As Long, ByVal rad As Long)
    If KtsSisteUavklarte Is Nothing Then Exit Sub
    If uavklartIndeks < 1 Or uavklartIndeks > KtsSisteUavklarte.Count Then Exit Sub
    If rad < 1 Then Exit Sub

    ' En rad uten (unik) radnokkel kan ikke huskes trygt - avvis heller enn a
    ' la e-posten stille havne i uavklart-lista igjen.
    Dim nokkelFeil As String: nokkelFeil = KtsRadNokkelFeil(rad)
    If Len(nokkelFeil) > 0 Then
        MsgBox nokkelFeil, vbExclamation, "Kontaktsentralen"
        Exit Sub
    End If

    Dim konfig As Object: Set konfig = KtsLesKonfig()
    Dim tbl As ListObject: Set tbl = KtsAktivTabell()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Dim kontrollKolNr As Long
    If konfig.Exists("KTS_KOL_RESPONS") Then kontrollKolNr = KtsKolonneNr(tbl, konfig("KTS_KOL_RESPONS"))

    Dim uavklart As Object: Set uavklart = KtsSisteUavklarte(uavklartIndeks)

    Dim entryId As String: entryId = ""
    If uavklart.Exists("EntryID") Then entryId = CStr(uavklart("EntryID"))
    If Len(entryId) > 0 Then
        Dim cache As Object: Set cache = KtsLastCache()
        Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
        d("EntryID") = entryId
        d("StoreID") = IIf(uavklart.Exists("StoreID"), CStr(uavklart("StoreID")), "")
        d("RadNr") = rad
        d("Kategori") = uavklart("Kategori")
        d("Emne") = uavklart("Emne")
        d("Avsender") = uavklart("Avsender")
        d("Mottatt") = uavklart("Mottatt")
        If uavklart.Exists("ConversationID") Then d("ConversationID") = CStr(uavklart("ConversationID")) Else d("ConversationID") = ""
        d("Arsak") = ""
        If uavklart.Exists("Mappe") Then d("Mappe") = CStr(uavklart("Mappe")) Else d("Mappe") = ""
        If cache.Exists(entryId) Then cache.Remove entryId
        cache.Add entryId, d
        KtsLagreCache cache
        KtsGjenoppbyggSessionFraCache cache
    End If

    If kontrollKolNr > 0 Then
        tbl.DataBodyRange.Cells(rad, kontrollKolNr).Value = KtsKontrollstatusForRad(rad, konfig)
        KtsFormaterKontrollkolonne tbl, kontrollKolNr, konfig
    End If

    ' Sesjonslistene er allerede bygget pa nytt fra cachen ovenfor (den
    ' matchede e-posten er da ikke lenger uavklart) - fjern KUN manuelt
    ' hvis e-posten ikke kunne lagres i cachen (ingen EntryID).
    If Len(entryId) = 0 Then KtsSisteUavklarte.Remove uavklartIndeks
End Sub

' --- Inngangspunkt fra knappen pa arket ---
' Dashbordet apnes MODELESS (ligger apent ved siden av tabellen) sa
' markering i lista kan ta deg til raden i Excel.

Public Sub VisKontaktsentralen()
    Dim k As Object: Set k = KtsLesKonfig()
    If Not k.Exists("KTS_POSTBOKS") Or Not k.Exists("KTS_KOL_RESPONS") Or Not k.Exists("KTS_TABELL") Then
        frmKtsOppsett.Show
    ElseIf Len(k("KTS_POSTBOKS")) = 0 Or Len(k("KTS_KOL_RESPONS")) = 0 Or Len(k("KTS_TABELL")) = 0 Then
        frmKtsOppsett.Show
    ElseIf modKontaktsentralen.KtsFinnTabell(CStr(k("KTS_TABELL"))) Is Nothing Then
        frmKtsOppsett.Show
    Else
        frmKontaktsentralen.Show 0
    End If
End Sub
'@

# ---------------------------------------------------------------
# VBA-kildekode: frmKontaktsentralen (dashboard-vindu)
# ---------------------------------------------------------------
$formCodeKontaktsentralen = @'
Option Explicit
Private Const ANTALL_RADER As Long = 11
Private Const RAD_HOYDE As Single = 19
Private Const LISTE_VENSTRE As Single = 10
Private Const LISTE_TOPP As Single = 50
Private Const LISTE_BREDDE As Single = 760
Private Const RULLEFELT_BREDDE As Single = 16
Private Const STRIPE_BREDDE As Single = 5
Private Const INGEN_STATUS As String = "(Ingen kategori ennå)"

Private radNr() As Long
Private visD() As String
Private statusD() As String
Private meldD() As String
Private flereD() As Long
Private maxLenD() As Long
Private nokkelD() As String
Private bredder() As Single
Private nVis As Long
Private antallRader As Long
Private antallTotalt As Long
Private farger As Object
Private toppPos As Long
Private valgtPos As Long
Private laster As Boolean

Private Sub UserForm_Initialize()
    lblTittel.Font.Bold = True: lblTittel.Caption = "Kontaktsentralen"
    lblVersjon.Caption = "Versjon: " & modKontaktsentralen.KTS_VERSION
    modKontaktsentralen.KtsStilKnapp cmdSkann, "gronn"
    modKontaktsentralen.KtsStilKnapp cmdInnstillinger, "blagra"
    ' Ingen automatisk skann ved apning (Håkons onske 2026-09-19) - lista
    ' bygges fra Kontrollkolonnen og den lagrede cachen. Bare Skann na
    ' henter nye e-poster fra Outlook.
    modKontaktsentralen.KtsSikreSessionFraCache
    toppPos = 1
    OppdaterListe
    OppdaterUavklarteKnapp
    modKontaktsentralen.KtsTilpassVindu Me
    PlasserVindu
    ' Fokus legges aldri pa rullefeltet (det blinket nar det hadde fokus):
    ' tastaturet styres fra sokefeltet, se txtSok_KeyDown.
    On Error Resume Next
    sbRader.TabStop = False
    txtSok.SetFocus
End Sub

' Dashbordet er modeless og sentreres i Excel-vinduet, likt de andre
' makroenes hovedvinduer (Haakons onske 2026-09-24 - tidligere nederst
' til hoyre/venstre). Kan flyttes fritt.
Private Sub PlasserVindu()
    On Error Resume Next
    Me.StartUpPosition = 0
    Dim t As Double, l As Double
    t = Application.Top + (Application.Height - Me.Height) / 2
    l = Application.Left + (Application.Width - Me.Width) / 2
    If t < Application.Top Then t = Application.Top
    If l < Application.Left Then l = Application.Left
    Me.Top = t
    Me.Left = l
End Sub

Private Sub OppdaterUavklarteKnapp()
    Dim n As Long: n = modKontaktsentralen.KtsAntallUavklarte()
    If n > 0 Then
        cmdUavklarte.Caption = "Vis uavklarte (" & n & ")"
        modKontaktsentralen.KtsStilKnapp cmdUavklarte, "oransje"
    Else
        cmdUavklarte.Caption = "Vis uavklarte"
        modKontaktsentralen.KtsStilKnapp cmdUavklarte, "gra"
    End If
End Sub

Private Sub Synkroniser(ByVal visMelding As Boolean)
    Dim resultat As Object: Set resultat = modKontaktsentralen.KjørSkann()
    lblVersjon.Caption = "Versjon: " & modKontaktsentralen.KTS_VERSION
    If resultat.Exists("Feil") Then MsgBox resultat("Feil"), vbExclamation, "Kontaktsentralen": Exit Sub
    OppdaterListe
    OppdaterUavklarteKnapp
    If visMelding Then
        Dim m As String
        m = "Skannet: " & resultat("AntallSkannet") & " e-poster" & vbCrLf & _
            "Med kategori: " & resultat("AntallKategorisert") & vbCrLf & _
            "Kjent fra tidligere (hoppet over ny matching): " & resultat("AntallKjentFraCache") & vbCrLf & _
            "Nytt matchet direkte: " & resultat("AntallNyeMatchet") & vbCrLf & _
            "Matchet via samtale: " & resultat("AntallMatchetViaSamtale") & vbCrLf & _
            "Uavklart/flertydig: " & resultat("Uavklarte").Count & vbCrLf & _
            "Rader med status: " & resultat("AntallOppdatertRad")
        MsgBox m, vbInformation, "Kontaktsentralen - synkronisering ferdig"
    End If
End Sub

' Henter radene pa nytt (sortert), beholder valgt rad og rulleposisjon sa langt
' det lar seg gjore, og tegner lista om.
Private Sub OppdaterListe()
    ' Valgt rad huskes med radnokkelen, ikke posisjonen - tabellen kan ha blitt
    ' sortert eller filtrert i Excel siden sist.
    Dim forrigeNokkel As String
    If valgtPos > 0 And valgtPos <= antallRader Then forrigeNokkel = nokkelD(valgtPos)
    laster = True
    modKontaktsentralen.KtsSikreSessionFraCache
    modKontaktsentralen.KtsHentRadData radNr, visD, statusD, meldD, flereD, maxLenD, nokkelD, nVis, antallRader
    antallTotalt = antallRader
    FiltrerRader
    Set farger = modKontaktsentralen.KtsLesKategoriFarger(modKontaktsentralen.KtsLesKonfig())
    BeregnBredder
    valgtPos = 0
    If Len(forrigeNokkel) > 0 Then valgtPos = FinnPosForNokkel(forrigeNokkel)
    Dim maks As Long: maks = antallRader - ANTALL_RADER
    If maks < 0 Then maks = 0
    If toppPos > maks + 1 Then toppPos = maks + 1
    If toppPos < 1 Then toppPos = 1
    sbRader.Min = 0
    sbRader.Max = maks
    sbRader.SmallChange = 1
    sbRader.LargeChange = ANTALL_RADER - 1
    sbRader.Value = toppPos - 1
    sbRader.Enabled = (maks > 0)
    laster = False
    TegnRader
    OppdaterOverskrifter
    OppdaterTekst
End Sub

' Beholder bare radene som inneholder soketeksten (i en av visningskolonnene
' eller i statusen). Tomt sokefelt = alle rader.
Private Sub FiltrerRader()
    If antallRader = 0 Then Exit Sub
    Dim sok As String: sok = Trim(txtSok.Value)
    If Len(sok) = 0 Then Exit Sub
    Dim i As Long, j As Long, n As Long, treff As Boolean
    For i = 1 To antallRader
        treff = (InStr(1, statusD(i), sok, vbTextCompare) > 0)
        If Not treff Then
            For j = 1 To nVis
                If InStr(1, visD(i, j), sok, vbTextCompare) > 0 Then treff = True: Exit For
            Next j
        End If
        If treff Then
            n = n + 1
            If n <> i Then
                radNr(n) = radNr(i): statusD(n) = statusD(i): meldD(n) = meldD(i)
                flereD(n) = flereD(i): nokkelD(n) = nokkelD(i)
                For j = 1 To nVis
                    visD(n, j) = visD(i, j)
                Next j
            End If
        End If
    Next i
    antallRader = n
End Sub

Private Function FinnPosForNokkel(ByVal nokkel As String) As Long
    Dim i As Long
    For i = 1 To antallRader
        If nokkelD(i) = nokkel Then FinnPosForNokkel = i: Exit Function
    Next i
End Function

' Kolonnebredder i punkter for cellene: visningskolonnene, Status, E-poster.
' E-poster-cellen far restbredden, slik at raden er sammenhengende til kanten.
Private Sub BeregnBredder()
    If nVis = 0 Then Exit Sub
    Dim innerBredde As Single: innerBredde = LISTE_BREDDE - RULLEFELT_BREDDE - 2 - STRIPE_BREDDE
    Dim deler() As String
    deler = Split(modKontaktsentralen.KtsKolonneBredder(maxLenD, nVis, innerBredde), ";")
    ReDim bredder(1 To nVis + 2)
    Dim k As Long, sum As Single
    For k = 1 To nVis + 2
        bredder(k) = Val(Replace(deler(k - 1), ",", "."))
        If k <= nVis + 1 Then sum = sum + bredder(k)
    Next k
    If innerBredde - sum > bredder(nVis + 2) Then bredder(nVis + 2) = innerBredde - sum
End Sub

' Overskriftsrad over lista: en etikett per kolonne, plassert etter cellebreddene.
Private Sub OppdaterOverskrifter()
    On Error Resume Next
    Dim overskrifter As Collection: Set overskrifter = modKontaktsentralen.KtsListeOverskrifter()
    Dim x As Single: x = LISTE_VENSTRE + 1 + STRIPE_BREDDE
    Dim i As Long, lbl As Object
    For i = 1 To 6
        Set lbl = Me.Controls("lblHdr" & i)
        If nVis > 0 And i <= overskrifter.Count And i <= nVis + 2 Then
            lbl.Caption = overskrifter(i)
            lbl.Font.Bold = True
            lbl.Left = x + 2
            lbl.Width = bredder(i) - 2
            lbl.Visible = True
            x = x + bredder(i)
        Else
            lbl.Visible = False
        End If
    Next i
End Sub

Private Sub OppdaterTekst()
    Dim medStatus As Long, i As Long
    For i = 1 To antallRader
        If statusD(i) <> INGEN_STATUS Then medStatus = medStatus + 1
    Next i
    If antallRader < antallTotalt Then
        lblAntall.Caption = antallRader & " av " & antallTotalt & " rader passer med søket - " & medStatus & " med kategori-status. Marker en rad for å hoppe til den i Excel, dobbeltklikk for detaljer."
    Else
        lblAntall.Caption = antallRader & " rad(er) totalt - " & medStatus & " med kategori-status. Marker en rad for å hoppe til den i Excel, dobbeltklikk for detaljer."
    End If
End Sub

' Tegner de synlige radene (ANTALL_RADER etiketter-sett) ut fra rulleposisjonen.
Private Sub TegnRader()
    Dim i As Long, pos As Long
    For i = 1 To ANTALL_RADER
        pos = toppPos + i - 1
        If pos >= 1 And pos <= antallRader And nVis > 0 Then
            TegnRad i, pos
        Else
            SkjulRad i
        End If
    Next i
End Sub

Private Sub SkjulRad(ByVal i As Long)
    Dim f As String: f = Format(i, "00")
    Dim j As Long
    Me.Controls("lblS" & f).Visible = False
    Me.Controls("lblT" & f).Visible = False
    Me.Controls("lblM" & f).Visible = False
    For j = 1 To 4
        Me.Controls("lblC" & f & j).Visible = False
    Next j
End Sub

' En rad: fargestripe til venstre og statuscellen i kategoriens Outlook-farge
' (gra hvis fargen er ukjent, ingen farge hvis raden ikke har kategori), radens
' ovrige celler hvite - lys bla nar raden er valgt.
Private Sub TegnRad(ByVal i As Long, ByVal pos As Long)
    Dim f As String: f = Format(i, "00")
    Dim y As Single: y = LISTE_TOPP + 1 + (i - 1) * RAD_HOYDE
    Dim x As Single: x = LISTE_VENSTRE + 1
    Dim bg As Long
    If pos = valgtPos Then bg = RGB(205, 225, 250) Else bg = RGB(255, 255, 255)
    Dim st As String: st = statusD(pos)
    Dim harStatus As Boolean: harStatus = (st <> INGEN_STATUS)
    Dim farge As Long: farge = -1
    If harStatus Then
        farge = modKontaktsentralen.KtsKategoriRGB(st, farger)
        If farge = -1 Then farge = RGB(217, 217, 217)
    End If

    Dim lbl As Object, j As Long
    Set lbl = Me.Controls("lblS" & f)
    lbl.Left = x: lbl.Top = y: lbl.Width = STRIPE_BREDDE: lbl.Height = RAD_HOYDE
    If harStatus Then lbl.BackColor = farge Else lbl.BackColor = bg
    lbl.Visible = True
    x = x + STRIPE_BREDDE

    For j = 1 To 4
        Set lbl = Me.Controls("lblC" & f & j)
        If j <= nVis Then
            lbl.Left = x: lbl.Top = y: lbl.Width = bredder(j): lbl.Height = RAD_HOYDE
            lbl.BackColor = bg: lbl.ForeColor = RGB(30, 30, 30)
            lbl.BorderStyle = 1: lbl.BorderColor = bg
            lbl.Caption = visD(pos, j)
            lbl.Visible = True
            x = x + bredder(j)
        Else
            lbl.Visible = False
        End If
    Next j

    Set lbl = Me.Controls("lblT" & f)
    lbl.Left = x: lbl.Top = y: lbl.Width = bredder(nVis + 1): lbl.Height = RAD_HOYDE
    ' En tynn kant i radens bakgrunnsfarge skiller de fargede cellene fra
    ' hverandre, slik at de ser ut som separate merkelapper.
    lbl.BorderStyle = 1: lbl.BorderColor = bg
    If harStatus Then
        lbl.BackColor = farge
        lbl.ForeColor = modKontaktsentralen.KtsTekstFargeMot(farge)
        If flereD(pos) > 0 Then lbl.Caption = st & "  +" & flereD(pos) Else lbl.Caption = st
    Else
        lbl.BackColor = bg
        lbl.ForeColor = RGB(140, 140, 140)
        lbl.Caption = st
    End If
    lbl.Visible = True
    x = x + bredder(nVis + 1)

    Set lbl = Me.Controls("lblM" & f)
    lbl.Left = x: lbl.Top = y: lbl.Width = bredder(nVis + 2): lbl.Height = RAD_HOYDE
    lbl.BackColor = bg: lbl.ForeColor = RGB(30, 30, 30)
    lbl.BorderStyle = 1: lbl.BorderColor = bg
    lbl.Caption = meldD(pos)
    lbl.Visible = True
End Sub

' Marker en rad (museklikk eller piltaster) - tar deg til raden i Excel uten a
' flytte fokus bort fra dashbordet.
Private Sub RadKlikk(ByVal i As Long)
    Dim pos As Long: pos = SikrePosisjon(toppPos + i - 1)
    If pos < 1 Then Exit Sub
    valgtPos = pos
    SkrollSynlig pos
    TegnRader
    On Error Resume Next
    txtSok.SetFocus
    On Error GoTo 0
    NavigerTilValgt
End Sub

Private Sub RadDblKlikk(ByVal i As Long)
    Dim pos As Long: pos = SikrePosisjon(toppPos + i - 1)
    If pos < 1 Then Exit Sub
    valgtPos = pos
    ApneRadside
End Sub

' Er raden i lista fortsatt samme rad i tabellen? Sortering/filtrering i Excel
' kan ha flyttet den siden lista ble tegnet. Er den flyttet bygges lista pa nytt,
' og radens nye posisjon i lista returneres (0 hvis den ikke finnes lenger).
Private Function SikrePosisjon(ByVal pos As Long) As Long
    If pos < 1 Or pos > antallRader Then Exit Function
    If modKontaktsentralen.KtsNokkelForRad(radNr(pos)) = nokkelD(pos) Then
        SikrePosisjon = pos
        Exit Function
    End If
    Dim k As String: k = nokkelD(pos)
    valgtPos = 0
    OppdaterListe
    If Len(k) > 0 Then SikrePosisjon = FinnPosForNokkel(k)
End Function

' Ruller lista slik at posisjonen er synlig.
Private Sub SkrollSynlig(ByVal pos As Long)
    If pos < toppPos Then toppPos = pos
    If pos > toppPos + ANTALL_RADER - 1 Then toppPos = pos - ANTALL_RADER + 1
    Dim maks As Long: maks = antallRader - ANTALL_RADER
    If maks < 0 Then maks = 0
    If toppPos > maks + 1 Then toppPos = maks + 1
    If toppPos < 1 Then toppPos = 1
    laster = True
    sbRader.Value = toppPos - 1
    laster = False
End Sub

Private Sub ApneRadside()
    If valgtPos > 0 Then valgtPos = SikrePosisjon(valgtPos)
    If valgtPos < 1 Or valgtPos > antallRader Then Exit Sub
    modKontaktsentralen.KtsValgtRad = radNr(valgtPos)
    frmKtsRad.Show
    OppdaterListe
    OppdaterUavklarteKnapp
End Sub

Private Sub NavigerTilValgt()
    If valgtPos < 1 Or valgtPos > antallRader Then Exit Sub
    If modKontaktsentralen.KtsGaTilRadStille(radNr(valgtPos)) Then
        OppdaterTekst
    Else
        lblAntall.Caption = "Raden er skjult av et filter i Excel, eller kunne ikke vises."
    End If
End Sub

' Flytter markeringen (piltaster, PageUp/PageDown, Home/End) og ruller slik at
' den valgte raden er synlig.
Private Sub FlyttValg(ByVal flytt As Long)
    If antallRader = 0 Then Exit Sub
    If valgtPos > 0 Then valgtPos = SikrePosisjon(valgtPos)
    If antallRader = 0 Then Exit Sub
    Dim ny As Long
    If valgtPos = 0 Then ny = toppPos - 1 + flytt Else ny = valgtPos + flytt
    If ny < 1 Then ny = 1
    If ny > antallRader Then ny = antallRader
    valgtPos = ny
    SkrollSynlig ny
    TegnRader
    NavigerTilValgt
End Sub

' Sokefeltet filtrerer lista mens du skriver.
Private Sub txtSok_Change()
    If laster Then Exit Sub
    toppPos = 1
    OppdaterListe
End Sub

' Tastaturstyring av lista fra sokefeltet (der fokus ligger): pil opp/ned,
' PageUp/PageDown flytter markeringen, Enter apner radsiden. Home/End og
' venstre/hoyre pil er vanlig tekstredigering i sokefeltet.
Private Sub txtSok_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 38: FlyttValg -1
        Case 40: FlyttValg 1
        Case 33: FlyttValg -(ANTALL_RADER - 1)
        Case 34: FlyttValg ANTALL_RADER - 1
        Case 13: ApneRadside
        Case Else: Exit Sub
    End Select
    KeyCode = 0
End Sub

Private Sub sbRader_Change()
    If laster Then Exit Sub
    toppPos = sbRader.Value + 1
    TegnRader
End Sub

Private Sub sbRader_Scroll()
    If laster Then Exit Sub
    toppPos = sbRader.Value + 1
    TegnRader
End Sub

Private Sub sbRader_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 38: FlyttValg -1
        Case 40: FlyttValg 1
        Case 33: FlyttValg -(ANTALL_RADER - 1)
        Case 34: FlyttValg ANTALL_RADER - 1
        Case 36: FlyttValg -antallRader
        Case 35: FlyttValg antallRader
        Case 13: ApneRadside
        Case Else: Exit Sub
    End Select
    KeyCode = 0
End Sub

Private Sub cmdSkann_Click()
    lblVersjon.Caption = "Synkroniserer ...": Me.Repaint
    Synkroniser True
End Sub
Private Sub cmdUavklarte_Click()
    Unload Me
    frmKtsUavklart.Show
End Sub
Private Sub cmdInnstillinger_Click()
    Unload Me
    frmKtsOppsett.Show
End Sub
'@
# En Click- og en DblClick-handler per etikett i radene (lblS/lblC/lblT/lblM +
# radnummer med to sifre) - alle sender bare radnummeret videre.
$radKlikkKode = ""
foreach ($rNr in 1..11) {
    $rr = "{0:00}" -f $rNr
    $navnListe = @("lblS$rr", "lblT$rr", "lblM$rr") + (1..4 | ForEach-Object { "lblC$rr$_" })
    foreach ($nm in $navnListe) {
        $radKlikkKode += "Private Sub ${nm}_Click(): RadKlikk ${rNr}: End Sub`n"
        $radKlikkKode += "Private Sub ${nm}_DblClick(ByVal Cancel As MSForms.ReturnBoolean): RadDblKlikk ${rNr}: End Sub`n"
    }
}
$formCodeKontaktsentralen = $formCodeKontaktsentralen + "`n" + $radKlikkKode

# ---------------------------------------------------------------
# VBA-kildekode: frmKtsOppsett (innstillinger)
# ---------------------------------------------------------------
$formCodeKtsOppsett = @'
Option Explicit
Private initialiserer As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo EH
    initialiserer = True

    lblTittel.Font.Bold = True
    lblTittel.Caption = "Innstillinger - Kontaktsentralen"
    modKontaktsentralen.KtsStilKnapp cmdLagre, "gronn"
    modKontaktsentralen.KtsStilKnapp cmdTilbake, "blagra"
    modKontaktsentralen.KtsStilKnapp cmdKatOpp, "gra"
    modKontaktsentralen.KtsStilKnapp cmdKatNed, "gra"
    modKontaktsentralen.KtsStilKnapp cmdVisOpp, "gra"
    modKontaktsentralen.KtsStilKnapp cmdVisNed, "gra"
    modKontaktsentralen.KtsTilpassVindu Me

    Dim tbl As ListObject
    Set tbl = modKontaktsentralen.KtsAktivTabell()

    ' Tabellvelger: alle tabeller i arbeidsboka (navn + arknavn).
    cboTabell.Clear
    Dim tabell As Object
    For Each tabell In modKontaktsentralen.KtsAlleTabeller()
        cboTabell.AddItem CStr(tabell.Name)
        cboTabell.List(cboTabell.ListCount - 1, 1) = CStr(tabell.Parent.Name)
    Next tabell
    If Not tbl Is Nothing Then SettComboHvisFinnes cboTabell, tbl.Name

    Dim k As Object
    Set k = modKontaktsentralen.KtsLesKonfig()
    LastKolonner tbl, k

    ' Hent postbokser uten at Change-eventet starter kategoriskann under initiering.
    cboPostboks.Clear
    Dim olApp As Object, ns As Object, fld As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    If Not olApp Is Nothing Then Set ns = olApp.GetNamespace("MAPI")
    If Not ns Is Nothing Then
        For Each fld In ns.Folders
            cboPostboks.AddItem CStr(fld.Name)
        Next fld
    End If
    On Error GoTo EH

    SettComboHvisFinnes cboPostboks, TryggKonfig(k, "KTS_POSTBOKS")
    txtStartpunkt.Value = TryggKonfig(k, "KTS_STARTPUNKT")

    initialiserer = False
    LastKategorier
    Exit Sub

EH:
    initialiserer = False
    MsgBox "Klarte ikke å åpne Innstillinger." & vbCrLf & vbCrLf & _
           "Feil " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Kontaktsentralen"
End Sub

' Fyller kolonnelistene ut fra en tabell og markerer det som er lagret i
' innstillingene (kolonner som ikke finnes i tabellen hoppes over).
Private Sub LastKolonner(ByVal tbl As ListObject, ByVal k As Object)
    Dim forrige As Boolean: forrige = initialiserer
    initialiserer = True
    lstEpost.Clear
    lstGjenkjenning.Clear
    cboRespons.Clear
    cboNokkel.Clear
    cboRespons.AddItem "(Velg kontrollkolonne)"
    cboNokkel.AddItem "(Første gjenkjenningskolonne)"
    If Not tbl Is Nothing Then
        Dim kol As ListColumn
        For Each kol In tbl.ListColumns
            lstEpost.AddItem CStr(kol.Name)
            lstGjenkjenning.AddItem CStr(kol.Name)
            cboRespons.AddItem CStr(kol.Name)
            cboNokkel.AddItem CStr(kol.Name)
        Next kol
    End If
    If k.Exists("KTS_KOL_EPOST") Then MarkerValgteElementer lstEpost, TryggTekst(k("KTS_KOL_EPOST"))
    If k.Exists("KTS_KOL_GJENKJENNING") Then MarkerValgteElementer lstGjenkjenning, TryggTekst(k("KTS_KOL_GJENKJENNING"))
    SettComboHvisFinnes cboRespons, TryggKonfig(k, "KTS_KOL_RESPONS")
    SettComboHvisFinnes cboNokkel, TryggKonfig(k, "KTS_KOL_NOKKEL")
    If cboNokkel.ListIndex < 0 Then cboNokkel.ListIndex = 0
    LastVisningskolonner tbl, k
    initialiserer = forrige
End Sub

' Bytter tabell: lastes kolonnelistene pa nytt (valgene i de gamle
' kolonnene gjelder ikke lenger).
Private Sub cboTabell_Change()
    If initialiserer Then Exit Sub
    Dim t As ListObject
    Set t = modKontaktsentralen.KtsFinnTabell(TryggTekst(cboTabell.Value))
    If t Is Nothing Then Exit Sub
    LastKolonner t, modKontaktsentralen.KtsLesKonfig()
End Sub

Private Function TryggTekst(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        TryggTekst = ""
    Else
        TryggTekst = CStr(v)
    End If
End Function

Private Function TryggKonfig(ByVal k As Object, ByVal nokkel As String) As String
    On Error GoTo Ferdig
    If Not k Is Nothing Then
        If k.Exists(nokkel) Then TryggKonfig = TryggTekst(k(nokkel))
    End If
Ferdig:
End Function

Private Sub SettComboHvisFinnes(ByVal cbo As Object, ByVal verdi As String)
    If Len(verdi) = 0 Then Exit Sub
    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If StrComp(TryggTekst(cbo.List(i)), verdi, vbTextCompare) = 0 Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub MarkerValgteElementer(ByVal lst As Object, ByVal lagret As String)
    If Len(lagret) = 0 Then Exit Sub
    Dim a As Variant, i As Long, x As Variant
    a = Split(lagret, "|")
    For i = 0 To lst.ListCount - 1
        For Each x In a
            If StrComp(TryggTekst(lst.List(i)), TryggTekst(x), vbTextCompare) = 0 Then
                lst.Selected(i) = True
                Exit For
            End If
        Next x
    Next i
End Sub

Private Function HentValgteElementer(ByVal lst As Object) As String
    Dim i As Long, s As String
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            If Len(s) > 0 Then s = s & "|"
            s = s & TryggTekst(lst.List(i))
        End If
    Next i
    HentValgteElementer = s
End Function

Private Sub cboPostboks_Change()
    If initialiserer Then Exit Sub
    LastKategorier
End Sub

Private Sub LastKategorier()
    On Error GoTo EH
    If initialiserer Then Exit Sub
    If Not HasControl("lstKategorier") Then Exit Sub

    lstKategorier.Clear

    Dim postboks As String
    postboks = TryggTekst(cboPostboks.Value)
    If Len(postboks) = 0 Then Exit Sub

    Dim k As Object
    Set k = modKontaktsentralen.KtsLesKonfig()

    Dim prior As String
    prior = TryggKonfig(k, "KTS_KATEGORI_PRIORITET")

    Dim funnet As Collection
    Set funnet = modKontaktsentralen.KtsHentKategorierFraPostboks(postboks)
    If funnet Is Nothing Then Exit Sub

    Dim brukt As Object
    Set brukt = CreateObject("Scripting.Dictionary")

    Dim a As Variant, x As Variant, i As Long
    If Len(prior) > 0 Then
        a = Split(prior, "|")
        For Each x In a
            For i = 1 To funnet.Count
                If StrComp(TryggTekst(funnet(i)), TryggTekst(x), vbTextCompare) = 0 Then
                    lstKategorier.AddItem TryggTekst(x)
                    brukt(LCase$(TryggTekst(x))) = True
                    Exit For
                End If
            Next i
        Next x
    End If

    For i = 1 To funnet.Count
        If Not brukt.Exists(LCase$(TryggTekst(funnet(i)))) Then
            lstKategorier.AddItem TryggTekst(funnet(i))
        End If
    Next i
    Exit Sub

EH:
    ' Kategorier er hjelpedata. Oppsett-vinduet skal fortsatt kunne åpnes
    ' selv om Outlook ikke kan lese kategoriene akkurat nå.
End Sub

Private Function HasControl(ByVal n As String) As Boolean
    On Error Resume Next
    Dim o As Object
    Set o = Me.Controls(n)
    HasControl = Not o Is Nothing
    On Error GoTo 0
End Function

Private Sub cmdKatOpp_Click(): FlyttKategori -1: End Sub
Private Sub cmdKatNed_Click(): FlyttKategori 1: End Sub

Private Sub FlyttKategori(ByVal d As Long)
    Dim i As Long
    i = lstKategorier.ListIndex
    If i < 0 Or i + d < 0 Or i + d >= lstKategorier.ListCount Then Exit Sub
    Dim t As String
    t = TryggTekst(lstKategorier.List(i + d))
    lstKategorier.List(i + d) = TryggTekst(lstKategorier.List(i))
    lstKategorier.List(i) = t
    lstKategorier.ListIndex = i + d
End Sub

' Kolonner som vises for a gjenkjenne rader i listene: alle tabellkolonner
' med avkrysning. Lagret valg kommer forst (i lagret rekkefolge, avkrysset),
' resten under. Opp/Ned endrer rekkefolgen. Uten lagret valg forhandsvelges
' forste gjenkjenningskolonne (slik lista alltid har fungert).
Private Sub LastVisningskolonner(ByVal tbl As ListObject, ByVal k As Object)
    lstVisning.Clear
    If tbl Is Nothing Then Exit Sub
    Dim lagret As String: lagret = TryggKonfig(k, "KTS_KOL_VISNING")
    Dim brukt As Object: Set brukt = CreateObject("Scripting.Dictionary")
    Dim x As Variant, kol As ListColumn, standard As String
    If Len(lagret) > 0 Then
        For Each x In Split(lagret, "|")
            For Each kol In tbl.ListColumns
                If StrComp(CStr(kol.Name), CStr(x), vbTextCompare) = 0 Then
                    If Not brukt.Exists(LCase$(CStr(kol.Name))) Then
                        lstVisning.AddItem CStr(kol.Name)
                        lstVisning.Selected(lstVisning.ListCount - 1) = True
                        brukt(LCase$(CStr(kol.Name))) = True
                    End If
                    Exit For
                End If
            Next kol
        Next x
    Else
        standard = Split(TryggKonfig(k, "KTS_KOL_GJENKJENNING") & "|", "|")(0)
    End If
    For Each kol In tbl.ListColumns
        If Not brukt.Exists(LCase$(CStr(kol.Name))) Then
            lstVisning.AddItem CStr(kol.Name)
            If Len(standard) > 0 Then
                If StrComp(CStr(kol.Name), standard, vbTextCompare) = 0 Then lstVisning.Selected(lstVisning.ListCount - 1) = True
            End If
        End If
    Next kol
End Sub

Private Sub cmdVisOpp_Click()
    FlyttVisning -1
End Sub
Private Sub cmdVisNed_Click()
    FlyttVisning 1
End Sub

Private Sub FlyttVisning(ByVal d As Long)
    Dim i As Long
    i = lstVisning.ListIndex
    If i < 0 Or i + d < 0 Or i + d >= lstVisning.ListCount Then Exit Sub
    Dim t As String, valgtMal As Boolean, valgtKilde As Boolean
    t = TryggTekst(lstVisning.List(i + d))
    valgtMal = lstVisning.Selected(i + d)
    valgtKilde = lstVisning.Selected(i)
    lstVisning.List(i + d) = TryggTekst(lstVisning.List(i))
    lstVisning.Selected(i + d) = valgtKilde
    lstVisning.List(i) = t
    lstVisning.Selected(i) = valgtMal
    lstVisning.ListIndex = i + d
End Sub

' Avkryssede visningskolonner i listens rekkefolge, pipe-skilt.
Private Function HentVisningskolonner(ByRef antall As Long) As String
    Dim i As Long, s As String
    antall = 0
    For i = 0 To lstVisning.ListCount - 1
        If lstVisning.Selected(i) Then
            antall = antall + 1
            If Len(s) > 0 Then s = s & "|"
            s = s & TryggTekst(lstVisning.List(i))
        End If
    Next i
    HentVisningskolonner = s
End Function

Private Function KategoriPrioritet() As String
    Dim i As Long, s As String
    For i = 0 To lstKategorier.ListCount - 1
        If Len(s) > 0 Then s = s & "|"
        s = s & TryggTekst(lstKategorier.List(i))
    Next i
    KategoriPrioritet = s
End Function

Private Sub cmdLagre_Click()
    Dim postboks As String
    postboks = TryggTekst(cboPostboks.Value)
    If Len(postboks) = 0 Then MsgBox "Velg delt postboks.", vbExclamation: Exit Sub

    Dim kontrollkolonne As String
    kontrollkolonne = TryggTekst(cboRespons.Value)
    If Len(kontrollkolonne) = 0 Or kontrollkolonne = "(Velg kontrollkolonne)" Then
        MsgBox "Velg Kontrollkolonne.", vbExclamation
        Exit Sub
    End If

    Dim e As String, g As String
    e = HentValgteElementer(lstEpost)
    g = HentValgteElementer(lstGjenkjenning)
    If Len(e) = 0 And Len(g) = 0 Then MsgBox "Velg minst én matchingkolonne.", vbExclamation: Exit Sub

    Dim antVis As Long, vis As String
    vis = HentVisningskolonner(antVis)
    If antVis = 0 Then MsgBox "Velg minst én kolonne som skal vises i lista.", vbExclamation: Exit Sub
    If antVis > 4 Then MsgBox "Velg maks 4 kolonner som skal vises i lista.", vbExclamation: Exit Sub

    Dim tabellNavn As String, valgtTabell As ListObject
    tabellNavn = TryggTekst(cboTabell.Value)
    Set valgtTabell = modKontaktsentralen.KtsFinnTabell(tabellNavn)
    If valgtTabell Is Nothing Then MsgBox "Velg tabell.", vbExclamation: Exit Sub

    ' Radnokkel: valgt kolonne, ellers forste avkryssede gjenkjenningskolonne,
    ' ellers tabellens forste kolonne (samme regel som KtsNokkelKolonneNavn).
    Dim nokkel As String, effektivNokkel As String
    If cboNokkel.ListIndex > 0 Then nokkel = TryggTekst(cboNokkel.Value)
    effektivNokkel = nokkel
    If Len(effektivNokkel) = 0 Then effektivNokkel = Split(g & "|", "|")(0)
    If Len(effektivNokkel) = 0 Then effektivNokkel = CStr(valgtTabell.ListColumns(1).Name)
    Dim tomme As Long, dupl As Long
    modKontaktsentralen.KtsSjekkNokkelkolonne valgtTabell, effektivNokkel, tomme, dupl
    If tomme > 0 Or dupl > 0 Then
        If MsgBox("Radnøkkel-kolonnen '" & effektivNokkel & "' har " & tomme & " rad(er) uten verdi og " & dupl & " verdi(er) som finnes på flere rader." & vbCrLf & vbCrLf & _
                  "E-post kan ikke kobles sikkert til slike rader - de havner i 'uavklart' i stedet. Velg helst en kolonne med unike verdier (f.eks. skole-ID)." & vbCrLf & vbCrLf & _
                  "Lagre likevel?", vbExclamation + vbYesNo, "Kontaktsentralen") = vbNo Then Exit Sub
    End If

    ' Lagrede meldingskoblinger folger radnokkelen. Bytter du nokkelkolonne
    ' (i samme tabell) leses de inn med de gamle innstillingene og lagres pa
    ' nytt med de nye, sa ingenting mister raden sin.
    Dim gjeldendeTabell As ListObject, gammelCache As Object
    Set gjeldendeTabell = modKontaktsentralen.KtsAktivTabell()
    If Not gjeldendeTabell Is Nothing Then
        If StrComp(gjeldendeTabell.Name, valgtTabell.Name, vbTextCompare) = 0 Then Set gammelCache = modKontaktsentralen.KtsLastCache()
    End If

    modKontaktsentralen.KtsSkrivKonfigVerdi "KTS_TABELL", CStr(valgtTabell.Name)
    modKontaktsentralen.KtsSkrivKonfigVerdi "KTS_KOL_NOKKEL", nokkel
    modKontaktsentralen.KtsSkrivKonfigVerdi "KTS_KOL_VISNING", vis
    modKontaktsentralen.KtsSkrivKonfigVerdi "KTS_POSTBOKS", postboks
    modKontaktsentralen.KtsSkrivKonfigVerdi "KTS_KOL_EPOST", e
    modKontaktsentralen.KtsSkrivKonfigVerdi "KTS_KOL_GJENKJENNING", g
    modKontaktsentralen.KtsSkrivKonfigVerdi "KTS_KOL_RESPONS", kontrollkolonne
    modKontaktsentralen.KtsSkrivKonfigVerdi "KTS_STARTPUNKT", TryggTekst(txtStartpunkt.Value)
    modKontaktsentralen.KtsSkrivKonfigVerdi "KTS_KATEGORI_PRIORITET", KategoriPrioritet()

    If Not gammelCache Is Nothing Then modKontaktsentralen.KtsLagreCache gammelCache

    Unload Me
    frmKontaktsentralen.Show 0
End Sub

Private Sub cmdTilbake_Click()
    Unload Me
    frmKontaktsentralen.Show 0
End Sub
'@

# ---------------------------------------------------------------
# VBA-kildekode: frmKtsRad (detaljside for valgt tabellrad)
# ---------------------------------------------------------------
$formCodeKtsRad = @'
Option Explicit
Private Const MAKS_KNAPPER As Long = 30
Private Const KATEGORI_TOPP As Single = 232
Private laster As Boolean
Private valgtIndeks As Long
Private kategoriNavn As Collection
Private origRekkefolge As Collection
Private opprinnelig As Object
Private valgt As Object
Private kategoriFarger As Object

Private Sub UserForm_Initialize()
    Dim rad As Long: rad = modKontaktsentralen.KtsValgtRad
    lblTittel.Font.Bold = True
    lblTittel.Caption = "Radkontroll - " & modKontaktsentralen.KtsRadNavnForVisning(rad)
    modKontaktsentralen.KtsStilKnapp cmdAapneOutlook, "bla"
    modKontaktsentralen.KtsStilKnapp cmdMatchUavklart, "oransje"
    modKontaktsentralen.KtsStilKnapp cmdHovedmeny, "blagra"
    Set valgt = CreateObject("Scripting.Dictionary")
    Set opprinnelig = CreateObject("Scripting.Dictionary")
    Set origRekkefolge = New Collection
    valgtIndeks = -1
    ' Ferske farger og kategoriliste fra en Outlook som allerede kjorer
    ' (feiler stille, da brukes det som ble lagret sist).
    modKontaktsentralen.KtsOppdaterKategorilisteStille
    Set kategoriFarger = modKontaktsentralen.KtsLesKategoriFarger(modKontaktsentralen.KtsLesKonfig())
    PlasserOverskrifter
    LastMeldinger
    ByggKnapper rad
    OppdaterKnapper
    modKontaktsentralen.KtsTilpassVindu Me
    ' Ta Excel til raden ogsa nar radsiden apnes (f.eks. via dashbordet).
    modKontaktsentralen.KtsGaTilRadStille rad
    If lstMeldinger.ListCount = 1 Then lstMeldinger.ListIndex = 0
End Sub

' Overskriftsrad over meldingslista (samme teknikk som dashbordet).
Private Sub PlasserOverskrifter()
    On Error Resume Next
    Dim navn As Variant: navn = Array("Mottatt", "Avsender", "Emne", "Mappe", "Kategorier")
    Dim bredder() As String: bredder = Split(lstMeldinger.ColumnWidths, ";")
    Dim x As Single: x = lstMeldinger.Left + 4
    Dim i As Long, b As Single, lbl As Object
    For i = 0 To 4
        Set lbl = Me.Controls("lblHdr" & (i + 1))
        b = Val(Replace(bredder(i), ",", "."))
        lbl.Caption = navn(i)
        lbl.Font.Bold = True
        lbl.Left = x
        lbl.Width = b
        x = x + b
    Next i
End Sub

Private Sub LastMeldinger()
    Dim rad As Long: rad = modKontaktsentralen.KtsValgtRad
    laster = True
    lstMeldinger.Clear
    If Not modKontaktsentralen.KtsSisteMeldinger Is Nothing Then
        If modKontaktsentralen.KtsSisteMeldinger.Exists(rad) Then
            Dim i As Long, m As Object, n As Long
            For i = 1 To modKontaktsentralen.KtsSisteMeldinger(rad).Count
                Set m = modKontaktsentralen.KtsSisteMeldinger(rad)(i)
                lstMeldinger.AddItem m("Mottatt")
                n = lstMeldinger.ListCount - 1
                lstMeldinger.List(n, 1) = m("Avsender")
                lstMeldinger.List(n, 2) = m("Emne")
                If m.Exists("Mappe") Then lstMeldinger.List(n, 3) = m("Mappe")
                lstMeldinger.List(n, 4) = m("Kategori")
            Next i
        End If
    End If
    If valgtIndeks >= lstMeldinger.ListCount Then valgtIndeks = -1
    If valgtIndeks >= 0 Then lstMeldinger.ListIndex = valgtIndeks
    laster = False
End Sub

' Legger ut en knapp per kategori (postboksens kategorier + de som allerede
' sitter pa e-poster for raden) og flytter handlingsknappene under dem.
Private Sub ByggKnapper(ByVal rad As Long)
    Set kategoriNavn = modKontaktsentralen.KtsKategoriValgForRad(rad)
    Dim i As Long, x As Single, y As Single, b As Single, knapp As Object
    x = 10: y = KATEGORI_TOPP
    For i = 1 To MAKS_KNAPPER
        Set knapp = Me.Controls("cmdKat" & i)
        If i <= kategoriNavn.Count Then
            b = Len(kategoriNavn(i)) * 6.5 + 26
            If b < 70 Then b = 70
            If b > 220 Then b = 220
            If x + b > 730 And x > 10 Then
                x = 10
                y = y + 30
            End If
            knapp.Left = x: knapp.Top = y: knapp.Width = b: knapp.Height = 24
            knapp.Caption = kategoriNavn(i)
            knapp.Visible = True
            x = x + b + 6
        Else
            knapp.Visible = False
        End If
    Next i
    Dim bunn As Single: bunn = y + 24
    If kategoriNavn.Count > MAKS_KNAPPER Then lblKategorier.ControlTipText = "Bare de " & MAKS_KNAPPER & " forste kategoriene vises som knapper."
    Dim t As Single: t = bunn + 16
    cmdLagreKategori.Top = t: cmdAngre.Top = t: cmdAapneOutlook.Top = t: cmdMatchUavklart.Top = t: cmdHovedmeny.Top = t
End Sub

' Tegner om knappene ut fra valget: valgt kategori far Outlook-fargen (fet
' skrift), ikke valgt er lys gra. Lagre/Angre er bare aktive ved endring.
Private Sub OppdaterKnapper()
    Dim i As Long, knapp As Object, farge As Long
    For i = 1 To kategoriNavn.Count
        If i > MAKS_KNAPPER Then Exit For
        Set knapp = Me.Controls("cmdKat" & i)
        knapp.Enabled = (valgtIndeks >= 0)
        If valgt.Exists(LCase(kategoriNavn(i))) Then
            farge = modKontaktsentralen.KtsKategoriRGB(CStr(kategoriNavn(i)), kategoriFarger)
            If farge = -1 Then farge = RGB(58, 114, 196)
            knapp.BackColor = farge
            knapp.ForeColor = modKontaktsentralen.KtsTekstFargeMot(farge)
            knapp.Font.Bold = True
        Else
            knapp.BackColor = RGB(240, 240, 240)
            knapp.ForeColor = RGB(110, 110, 110)
            knapp.Font.Bold = False
        End If
    Next i
    Dim endret As Boolean: endret = ErEndret()
    cmdLagreKategori.Enabled = endret
    cmdAngre.Enabled = endret
    If endret Then
        modKontaktsentralen.KtsStilKnapp cmdLagreKategori, "gronn"
        modKontaktsentralen.KtsStilKnapp cmdAngre, "blagra"
        lblKategorier.Caption = "Kategorier på valgt e-post - ulagrede endringer. Trykk Lagre i Outlook eller Angre."
    Else
        modKontaktsentralen.KtsStilKnapp cmdLagreKategori, "gra"
        modKontaktsentralen.KtsStilKnapp cmdAngre, "gra"
        If valgtIndeks < 0 Then
            lblKategorier.Caption = "Velg en e-post i lista for å endre kategoriene."
        Else
            lblKategorier.Caption = "Kategorier på valgt e-post - klikk en kategori for å slå den av eller på."
        End If
    End If
End Sub

Private Function ErEndret() As Boolean
    If valgt Is Nothing Then Exit Function
    If valgt.Count <> opprinnelig.Count Then
        ErEndret = True
        Exit Function
    End If
    Dim k As Variant
    For Each k In valgt.Keys
        If Not opprinnelig.Exists(CStr(k)) Then
            ErEndret = True
            Exit Function
        End If
    Next k
End Function

' Leser kategoriene til den valgte e-posten inn som utgangspunkt for valget.
Private Sub LastValg()
    valgt.RemoveAll
    opprinnelig.RemoveAll
    Set origRekkefolge = New Collection
    If valgtIndeks >= 0 Then
        Dim m As Object, kat As Variant, n As String
        Set m = modKontaktsentralen.KtsSisteMeldinger(modKontaktsentralen.KtsValgtRad)(valgtIndeks + 1)
        For Each kat In modKontaktsentralen.KtsSplitKategorier(CStr(m("Kategori")))
            n = LCase(CStr(kat))
            If Not valgt.Exists(n) Then
                valgt.Add n, True
                opprinnelig.Add n, True
                origRekkefolge.Add CStr(kat)
            End If
        Next kat
    End If
    OppdaterKnapper
End Sub

Private Sub KatKlikk(ByVal i As Long)
    If valgtIndeks < 0 Then Exit Sub
    If i > kategoriNavn.Count Then Exit Sub
    Dim n As String: n = LCase(CStr(kategoriNavn(i)))
    If valgt.Exists(n) Then
        valgt.Remove n
    Else
        valgt.Add n, True
    End If
    OppdaterKnapper
End Sub

Private Sub lstMeldinger_Change()
    If laster Then Exit Sub
    If lstMeldinger.ListIndex < 0 Then Exit Sub
    If lstMeldinger.ListIndex = valgtIndeks Then Exit Sub
    If ErEndret() Then
        If MsgBox("Du har ulagrede kategori-endringer på forrige e-post. Forkaste dem?", vbQuestion + vbYesNo, "Kontaktsentralen") = vbNo Then
            laster = True
            lstMeldinger.ListIndex = valgtIndeks
            laster = False
            Exit Sub
        End If
    End If
    valgtIndeks = lstMeldinger.ListIndex
    LastValg
End Sub

Private Sub lstMeldinger_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdAapneOutlook_Click
End Sub

Private Sub cmdAapneOutlook_Click()
    If lstMeldinger.ListIndex >= 0 Then modKontaktsentralen.KtsAapneMelding modKontaktsentralen.KtsValgtRad, lstMeldinger.ListIndex + 1
End Sub

' Skriver valget til Outlook for DEN VALGTE e-posten. Eksisterende kategorier
' beholder rekkefolgen sin, nye legges til etter.
Private Function ByggLagreListe() As Collection
    Dim liste As New Collection, kat As Variant, i As Long, n As String
    For Each kat In origRekkefolge
        If valgt.Exists(LCase(CStr(kat))) Then liste.Add CStr(kat)
    Next kat
    For i = 1 To kategoriNavn.Count
        n = LCase(CStr(kategoriNavn(i)))
        If valgt.Exists(n) And Not opprinnelig.Exists(n) Then liste.Add CStr(kategoriNavn(i))
    Next i
    Set ByggLagreListe = liste
End Function

Private Sub cmdLagreKategori_Click()
    If valgtIndeks < 0 Then Exit Sub
    If Not ErEndret() Then Exit Sub
    Dim liste As Collection: Set liste = ByggLagreListe()
    Dim feil As String, lagret As String
    lagret = modKontaktsentralen.KtsLagreKategorierPaaMelding(modKontaktsentralen.KtsValgtRad, valgtIndeks + 1, liste, feil)
    If Len(feil) > 0 Then
        MsgBox feil, vbExclamation, "Kontaktsentralen"
        Exit Sub
    End If
    LastMeldinger
    LastValg
    If valgt.Count <> liste.Count Then MsgBox "Outlook lagret ikke akkurat de kategoriene du valgte. Slik ligger den nå: " & lagret, vbExclamation, "Kontaktsentralen"
End Sub

Private Sub cmdAngre_Click()
    If valgtIndeks < 0 Then Exit Sub
    LastValg
End Sub

' Apner uavklart-vinduet slik at en e-post skanningen ikke klarte a
' koble automatisk kan festes til NETTOPP denne raden - dekker Håkons
' onske 2026-09-18 om a kunne "justere parsing/matching" fra radsiden.
Private Sub cmdMatchUavklart_Click()
    modKontaktsentralen.KtsUavklartApnetFraRad = True
    frmKtsUavklart.Show
    modKontaktsentralen.KtsUavklartApnetFraRad = False
    LastMeldinger
    ByggKnapper modKontaktsentralen.KtsValgtRad
    OppdaterKnapper
    modKontaktsentralen.KtsTilpassVindu Me
End Sub

' Tilbake til hovedmenyen (dashbordet). Er dashbordet allerede apent (radsiden
' ble apnet derfra) lukker vi bare radsiden; ble den apnet med dobbeltklikk i
' en celle i Excel apnes dashbordet na.
Private Sub cmdHovedmeny_Click()
    If ErEndret() Then
        If MsgBox("Du har ulagrede kategori-endringer. Lukke uten å lagre?", vbQuestion + vbYesNo, "Kontaktsentralen") = vbNo Then Exit Sub
    End If
    Unload Me
    If Not modKontaktsentralen.KtsSkjemaErApent("frmKontaktsentralen") Then frmKontaktsentralen.Show 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 0 Then Exit Sub
    If ErEndret() Then
        If MsgBox("Du har ulagrede kategori-endringer. Lukke uten å lagre?", vbQuestion + vbYesNo, "Kontaktsentralen") = vbNo Then Cancel = True
    End If
End Sub

'@
# En Click-handler per kategoriknapp (cmdKat1..cmdKat30) - alle sender bare
# knappens nummer videre til KatKlikk.
$katKlikkKode = ""
foreach ($katNr in 1..30) { $katKlikkKode += "Private Sub cmdKat${katNr}_Click(): KatKlikk ${katNr}: End Sub`n" }
$formCodeKtsRad = $formCodeKtsRad + "`n" + $katKlikkKode

# ---------------------------------------------------------------
# VBA-kildekode: frmKtsUavklart (match uavklarte e-poster til rad)
# ---------------------------------------------------------------
$formCodeKtsUavklart = @'
Option Explicit

Private radNumre() As Long

Private Sub UserForm_Initialize()
    lblTittel.Font.Bold = True
    lblTittel.Caption = "Match uavklarte e-poster til rad"
    modKontaktsentralen.KtsStilKnapp cmdMatch, "gronn"
    modKontaktsentralen.KtsStilKnapp cmdTilbake, "blagra"
    LastUavklarte
    LastRader
    ForhaandsvelgSisteRad
    modKontaktsentralen.KtsTilpassVindu Me
End Sub

Private Sub LastUavklarte()
    lstUavklarte.Clear
    If Not modKontaktsentralen.KtsSisteUavklarte Is Nothing Then
        Dim i As Long, u As Object
        For i = 1 To modKontaktsentralen.KtsSisteUavklarte.Count
            Set u = modKontaktsentralen.KtsSisteUavklarte(i)
            lstUavklarte.AddItem u("Emne")
            lstUavklarte.List(lstUavklarte.ListCount - 1, 1) = u("Avsender")
            lstUavklarte.List(lstUavklarte.ListCount - 1, 2) = u("Kategori")
        Next i
    End If
    lblStatus.Caption = lstUavklarte.ListCount & " uavklart(e) igjen fra siste skanning"
End Sub

Private Sub LastRader()
    modKontaktsentralen.KtsByggRadliste lstRaderMatch, radNumre, lstRaderMatch.Width
    FiltrerRader
End Sub

' Beholder bare radene som inneholder soketeksten i en av kolonnene i lista
' (visningskolonnene, status). Tomt sokefelt = alle rader. radNumre() holdes
' i takt med lista, slik at valgt listeindeks fortsatt gir riktig rad.
Private Sub FiltrerRader()
    Dim sok As String: sok = Trim(txtSok.Value)
    If Len(sok) = 0 Then Exit Sub
    Dim n As Long: n = lstRaderMatch.ListCount
    If n = 0 Then Exit Sub
    Dim i As Long, c As Long, tekst As String
    For i = n - 1 To 0 Step -1
        tekst = ""
        For c = 0 To lstRaderMatch.ColumnCount - 1
            tekst = tekst & " " & lstRaderMatch.List(i, c)
        Next c
        If InStr(1, tekst, sok, vbTextCompare) = 0 Then
            radNumre(i) = -1
            lstRaderMatch.RemoveItem i
        End If
    Next i
    Dim igjen As Long: igjen = lstRaderMatch.ListCount
    If igjen = 0 Then
        Erase radNumre
        Exit Sub
    End If
    Dim ny() As Long, k As Long
    ReDim ny(0 To igjen - 1)
    For i = 0 To n - 1
        If radNumre(i) <> -1 Then
            ny(k) = radNumre(i)
            k = k + 1
        End If
    Next i
    radNumre = ny
    If igjen = 1 Then lstRaderMatch.ListIndex = 0
End Sub

' Sokefeltet filtrerer lista over rader mens du skriver.
Private Sub txtSok_Change()
    LastRader
End Sub

' Dobbeltklikk pa en uavklart e-post apner den i Outlook, sa du ser hva den
' handler om for du velger hvilken rad den skal kobles til.
Private Sub lstUavklarte_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstUavklarte.ListIndex < 0 Then Exit Sub
    If modKontaktsentralen.KtsSisteUavklarte Is Nothing Then Exit Sub
    Dim u As Object
    Set u = modKontaktsentralen.KtsSisteUavklarte(lstUavklarte.ListIndex + 1)
    modKontaktsentralen.KtsAapneMeldingMedId CStr(u("EntryID")), CStr(u("StoreID"))
End Sub

' Bekvemmelighet: hvis dette vinduet ble apnet fra en spesifikk rads
' detaljside (frmKtsRad), forhandsvelg akkurat den raden i lista til
' hoyre - Håkon slipper a lete den opp pa nytt.
Private Sub ForhaandsvelgSisteRad()
    On Error Resume Next
    Dim mal As Long: mal = modKontaktsentralen.KtsValgtRad
    If mal > 0 Then
        Dim i As Long
        For i = 0 To UBound(radNumre)
            If radNumre(i) = mal Then
                lstRaderMatch.ListIndex = i
                Exit For
            End If
        Next i
    End If
    On Error GoTo 0
End Sub

Private Sub cmdMatch_Click()
    Dim uavklartIndeks As Long, radIndeks As Long
    uavklartIndeks = lstUavklarte.ListIndex
    radIndeks = lstRaderMatch.ListIndex

    If uavklartIndeks < 0 Then
        MsgBox "Velg en uavklart e-post i lista til venstre først.", vbExclamation, "Match uavklarte"
        Exit Sub
    End If
    If radIndeks < 0 Then
        MsgBox "Velg hvilken rad e-posten faktisk hører til, i lista til høyre.", vbExclamation, "Match uavklarte"
        Exit Sub
    End If

    Dim rad As Long
    rad = radNumre(radIndeks)
    ' Collection er 1-basert, ListIndex er 0-basert.
    modKontaktsentralen.KtsMatchUavklartTilRad uavklartIndeks + 1, rad

    LastUavklarte
    LastRader
End Sub

Private Sub cmdTilbake_Click()
    ' Apnet fra en rads detaljside (frmKtsRad)? Da skal vi bare lukke
    ' oss selv og la den siden - som fortsatt star der under - ta over
    ' igjen, IKKE ogsa apne hoveddashbordet oppa den.
    If modKontaktsentralen.KtsUavklartApnetFraRad Then
        Unload Me
        Exit Sub
    End If
    Unload Me
    On Error Resume Next
    Err.Clear
    AppActivate Application.Caption
    frmKontaktsentralen.Show 0
    If Err.Number <> 0 Then
        MsgBox "Klarte ikke å åpne dashbordet igjen (feil " & Err.Number & ": " & Err.Description & ")." & vbCrLf & vbCrLf & _
               "Trykk på ""Kontaktsentralen""-knappen på arket for å åpne på nytt.", vbExclamation, "Kontaktsentralen"
    End If
    On Error GoTo 0
End Sub
'@

function Sjekk-UtfPS1BOM {
    param([string]$FilSti)
    $bytes = [System.IO.File]::ReadAllBytes($FilSti)
    if ($bytes.Length -lt 3 -or $bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) {
        Write-Warning "Denne .ps1-fila mangler UTF-8 BOM - norske tegn kan bli feil. Se excel-vba-installer-pattern i minnet."
    }
}
Sjekk-UtfPS1BOM -FilSti $MyInvocation.MyCommand.Path

# ---------------------------------------------------------------
# 1) Velg malfila (.xlsm) - ALLTID en bevisst dialog, aldri gjetting
# ---------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms | Out-Null

if (-not $Path) {
    $dlgXlsm = New-Object System.Windows.Forms.OpenFileDialog
    $dlgXlsm.Title = "Velg Excel-fila (.xlsm) som skal fa/oppdatere Kontaktsentralen"
    $dlgXlsm.Filter = "Makroaktiverte Excel-filer (*.xlsm)|*.xlsm"
    if ($Mappe -and (Test-Path $Mappe)) { $dlgXlsm.InitialDirectory = $Mappe }
    if ($dlgXlsm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Output "Avbrutt - ingen fil valgt."
        exit
    }
    $Path = $dlgXlsm.FileName
}
if (-not (Test-Path $Path)) {
    throw "Fant ikke fila: $Path"
}
Write-Output "Excel-fil valgt: $Path"

# ---------------------------------------------------------------
# 2) Koble til Excel og finne/apne malfila
# ---------------------------------------------------------------
try {
    if ($NyExcelInstans) { $excel = $null } else { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
} catch {
    $excel = $null
}

$wb = $null
if ($excel) {
    foreach ($w in $excel.Workbooks) {
        if ($w.FullName -eq $Path) { $wb = $w; break }
    }
}
if (-not $wb) {
    if (-not $excel) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
    }
    $wb = $excel.Workbooks.Open($Path)
}

# --- Duplikat-sjekk: ikke installer i en fil som er apnet dobbelt opp ---
$duplikatMonster = $wb.Name + ':*'
$treff = @()
foreach ($w in $excel.Workbooks) {
    if ($w.Name -eq $wb.Name -or $w.Name -like $duplikatMonster) { $treff += $w.Name }
}
if ($treff.Count -gt 1) {
    throw "Fila ser ut til a vaere apnet flere ganger samtidig ($($treff -join ', ')). Lukk alle kopier og apne den pa nytt (evt. restart Excel) for du installerer - se excel-com-shared-instance-risk i minnet."
}
$antallExcelProsesser = (Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Measure-Object).Count
if ($antallExcelProsesser -gt 1) {
    Write-Warning "Flere Excel.exe-prosesser kjorer samtidig pa maskinen - kan ikke se workbooks i andre prosesser herfra. Vaer sikker pa at du redigerer riktig kopi."
}

try {
    $vbProject = $wb.VBProject
    $null = $vbProject.VBComponents.Count
} catch {
    throw "Fikk ikke tilgang til VBA-prosjektet. Aktiver 'Trust access to the VBA project object model' i Trust Center (Fil > Alternativer > Sikkerhetssenter > Innstillinger for sikkerhetssenter > Makroinnstillinger), og prov igjen."
}

# ---------------------------------------------------------------
# Dobbeltklikk-integrasjon i ThisWorkbook (samme monster som Makromeny sin
# "vis som fane"): en liten, tydelig MERKET kodeblokk settes inn i
# ThisWorkbook, som ikke kan fjernes/byttes ut som de andre komponentene og
# kan inneholde andre verktoys kode. Merkene gjor at akkurat denne blokken
# kan finnes og fjernes igjen uten a rore resten av ThisWorkbook.
# ---------------------------------------------------------------
$dblMarkerStart = "'===KONTAKTSENTRALEN_DOBBELTKLIKK_START==="
$dblMarkerSlutt = "'===KONTAKTSENTRALEN_DOBBELTKLIKK_SLUTT==="

function Get-ThisWorkbookKomponent {
    try { return $vbProject.VBComponents.Item("ThisWorkbook") } catch {}
    try { return $vbProject.VBComponents.Item($wb.CodeName) } catch {}
    return $null
}

function Remove-DobbeltklikkFraThisWorkbook {
    $komp = Get-ThisWorkbookKomponent
    if (-not $komp) { return }
    $cm = $komp.CodeModule
    $antall = $cm.CountOfLines
    $startLinje = 0; $sluttLinje = 0
    for ($i = 1; $i -le $antall; $i++) {
        $linje = $cm.Lines($i, 1).Trim()
        if ($linje -eq $dblMarkerStart) { $startLinje = $i }
        if ($linje -eq $dblMarkerSlutt) { $sluttLinje = $i }
    }
    if ($startLinje -gt 0 -and $sluttLinje -ge $startLinje) {
        $cm.DeleteLines($startLinje, ($sluttLinje - $startLinje + 1))
        Write-Output "  Dobbeltklikk-integrasjon fjernet fra ThisWorkbook"
    }
}

function Add-DobbeltklikkTilThisWorkbook {
    $komp = Get-ThisWorkbookKomponent
    if (-not $komp) {
        Write-Output "  ADVARSEL: fant ikke ThisWorkbook - dobbeltklikk i Kontrollkolonnen ble ikke satt opp."
        return
    }
    $cm = $komp.CodeModule
    $antall = $cm.CountOfLines
    $eksisterende = if ($antall -gt 0) { $cm.Lines(1, $antall) } else { "" }
    if ($eksisterende -match 'Sub\s+Workbook_SheetBeforeDoubleClick') {
        Write-Output "  ADVARSEL: ThisWorkbook har allerede en egen Workbook_SheetBeforeDoubleClick-hendelse."
        Write-Output "  Hopper over dobbeltklikk-integrasjonen for aa ikke odelegge eksisterende kode - alt annet fungerer som normalt."
        return
    }
    $blokk = @(
        $dblMarkerStart,
        "Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)",
        "    On Error Resume Next",
        "    modKontaktsentralen.KtsHaandterDobbeltklikk Sh, Target, Cancel",
        "    On Error GoTo 0",
        "End Sub",
        $dblMarkerSlutt
    ) -join "`r`n"
    $cm.InsertLines($antall + 1, $blokk)
    Write-Output "  Dobbeltklikk i Kontrollkolonnen lagt til i ThisWorkbook"
}

# ---------------------------------------------------------------
# 3) Avinstaller - fjerner KUN egne, eksakt navngitte komponenter
# ---------------------------------------------------------------
if ($Uninstall) {
    # Fjerner KUN de noyaktige, kjente Kts-komponentnavnene og knappen
    # med denne makroens EGET OnAction-navn, aldri et generisk filter -
    # arket kan ha andre makroers knapper/komponenter side om side
    # (f.eks. Mail-utsender eller Kolonnevelger). Den skjulte
    # KontaktsentralenKonfig-fana la staa urort med vilje - en senere
    # reinstallasjon gjenbruker da gamle innstillinger.
    Write-Output "Fjerner Kontaktsentralen fra $($wb.Name) ..."
    try { Remove-DobbeltklikkFraThisWorkbook } catch {}
    foreach ($navn in @("modKontaktsentralen", "frmKontaktsentralen", "frmKtsOppsett", "frmKtsUavklart", "frmKtsRad")) {
        try {
            $eksisterende = $vbProject.VBComponents.Item($navn)
            $vbProject.VBComponents.Remove($eksisterende)
            Write-Output "  Fjernet $navn"
        } catch {}
    }
    try {
        foreach ($ws in $wb.Worksheets) {
            foreach ($btn in @($ws.Buttons())) {
                if ($btn.OnAction -eq "VisKontaktsentralen" -or $btn.OnAction -like "*!VisKontaktsentralen") { $btn.Delete() }
            }
        }
    } catch {}
    try { $wb.Save() } catch {
        Write-Output "FEIL ved lagring: $($_.Exception.Message)"
        exit 1
    }
    Write-Output "Ferdig - Kontaktsentralen er fjernet fra $($wb.Name)."
    exit 0
}

# ---------------------------------------------------------------
# 4) Sjekk installert versjon - hopp over hvis allerede oppdatert
# ---------------------------------------------------------------
$installertVersjon = $null
$erFerskInstall = $true
try {
    $eksisterendeModul = $vbProject.VBComponents.Item("modKontaktsentralen")
    $erFerskInstall = $false
    $kilde = $eksisterendeModul.CodeModule.Lines(1, $eksisterendeModul.CodeModule.CountOfLines)
    if ($kilde -match 'KTS_VERSION\s*As\s*String\s*=\s*"([^"]+)"') {
        $installertVersjon = $Matches[1]
    }
} catch {}

$nyVersjon = $KtsVersion

if ($installertVersjon -eq $nyVersjon -and -not $Force) {
    Write-Output "Kontaktsentralen er allerede installert og oppdatert (versjon $installertVersjon) i $($wb.Name)."
} else {
    Write-Output "Oppdaterer Kontaktsentralen: versjon $installertVersjon -> $nyVersjon i $($wb.Name) ..."

    $moduleCode = $moduleCodeKontaktsentralen.Replace('{{VERSION}}', $nyVersjon)

    $tempDir = Join-Path $env:TEMP ("Kontaktsentralen_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    $tempExcel = $null
    $tempWb = $null
    try {
        $tempExcel = New-Object -ComObject Excel.Application
        $tempExcel.Visible = $false
        $tempExcel.DisplayAlerts = $false
        $tempWb = $tempExcel.Workbooks.Add()

        # --- modKontaktsentralen (standard-modul) ---
        $modKomponent = $tempWb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        $modKomponent.Name = "modKontaktsentralen_tmp"
        $modKomponent.CodeModule.AddFromString($moduleCode)
        $modBas = Join-Path $tempDir "modKontaktsentralen.bas"
        $modKomponent.Export($modBas)

        # --- frmKontaktsentralen (UserForm - dashboard) ---
        $frmKts = $tempWb.VBProject.VBComponents.Add(3)  # 3 = vbext_ct_MSForm
        $frmKts.Name = "frmKontaktsentralen"
        $frmKts.Properties("Caption").Value = "Kontaktsentralen"
        $frmKts.Properties("Width").Value = 780
        $frmKts.Properties("Height").Value = 420

        $lblKtsTittel = $frmKts.Designer.Controls.Add("Forms.Label.1")
        $lblKtsTittel.Name = "lblTittel"
        $lblKtsTittel.Caption = ""
        $lblKtsTittel.Left = 10; $lblKtsTittel.Top = 8; $lblKtsTittel.Width = 170; $lblKtsTittel.Height = 20

        # Sokefelt som filtrerer lista (erstatter musehjulet som navigasjon i en
        # lang liste - MSForms-kontroller uten hjulstotte kan ikke ta imot det).
        $lblKtsSok = $frmKts.Designer.Controls.Add("Forms.Label.1")
        $lblKtsSok.Name = "lblSok"
        $lblKtsSok.Caption = "Søk:"
        $lblKtsSok.Left = 200; $lblKtsSok.Top = 10; $lblKtsSok.Width = 30; $lblKtsSok.Height = 16
        $txtKtsSok = $frmKts.Designer.Controls.Add("Forms.TextBox.1")
        $txtKtsSok.Name = "txtSok"
        $txtKtsSok.Left = 232; $txtKtsSok.Top = 7; $txtKtsSok.Width = 220; $txtKtsSok.Height = 19

        $lblKtsVersjon = $frmKts.Designer.Controls.Add("Forms.Label.1")
        $lblKtsVersjon.Name = "lblVersjon"
        $lblKtsVersjon.Caption = ""
        $lblKtsVersjon.Left = 470; $lblKtsVersjon.Top = 10; $lblKtsVersjon.Width = 300; $lblKtsVersjon.Height = 16

        # Egenbygd, farget liste (en vanlig ListBox kan ikke fargelegge rader):
        # en hvit ramme, ANTALL_RADER rader a 7 etiketter (stripe, 4
        # visningskolonner, status, e-poster) og et rullefelt. Posisjon,
        # bredde og innhold settes ved kjoring - se TegnRad i skjemakoden.
        $lblRamme = $frmKts.Designer.Controls.Add("Forms.Label.1")
        $lblRamme.Name = "lblListeRamme"
        $lblRamme.Caption = ""
        $lblRamme.Left = 10; $lblRamme.Top = 50; $lblRamme.Width = 744; $lblRamme.Height = 211
        $lblRamme.BackColor = 16777215; $lblRamme.BorderStyle = 1; $lblRamme.BorderColor = 10921638
        foreach ($rNr in 1..11) {
            $rr = "{0:00}" -f $rNr
            foreach ($nm in (@("lblS$rr", "lblT$rr", "lblM$rr") + (1..4 | ForEach-Object { "lblC$rr$_" }))) {
                $rl = $frmKts.Designer.Controls.Add("Forms.Label.1")
                $rl.Name = $nm
                $rl.Caption = ""
                $rl.WordWrap = $false
                $rl.Left = 11; $rl.Top = 51; $rl.Width = 60; $rl.Height = 19
                $rl.Visible = $false
            }
        }
        $sbKtsRader = $frmKts.Designer.Controls.Add("Forms.ScrollBar.1")
        $sbKtsRader.Name = "sbRader"
        $sbKtsRader.Left = 754; $sbKtsRader.Top = 50; $sbKtsRader.Width = 16; $sbKtsRader.Height = 211
        # Overskriftsrad: seks etiketter som plasseres/fylles ved kjoring
        # (visningskolonnene 1-4 + Status + E-poster) etter kolonnebreddene.
        foreach ($hn in 1..6) {
            $lblHdr = $frmKts.Designer.Controls.Add("Forms.Label.1")
            $lblHdr.Name = "lblHdr$hn"
            $lblHdr.Caption = ""
            $lblHdr.Left = 14; $lblHdr.Top = 34; $lblHdr.Width = 100; $lblHdr.Height = 14
        }

        $lblKtsAntall = $frmKts.Designer.Controls.Add("Forms.Label.1")
        $lblKtsAntall.Name = "lblAntall"
        $lblKtsAntall.Caption = ""
        $lblKtsAntall.Left = 10; $lblKtsAntall.Top = 266; $lblKtsAntall.Width = 760; $lblKtsAntall.Height = 16

        $cmdKtsSkann = $frmKts.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdKtsSkann.Name = "cmdSkann"
        $cmdKtsSkann.Caption = "Skann nå"
        $cmdKtsSkann.Left = 10; $cmdKtsSkann.Top = 292; $cmdKtsSkann.Width = 130; $cmdKtsSkann.Height = 28

        $cmdKtsInnstillinger = $frmKts.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdKtsInnstillinger.Name = "cmdInnstillinger"
        $cmdKtsInnstillinger.Caption = "Innstillinger"
        $cmdKtsInnstillinger.Left = 150; $cmdKtsInnstillinger.Top = 292; $cmdKtsInnstillinger.Width = 130; $cmdKtsInnstillinger.Height = 28

        $cmdKtsUavklarte = $frmKts.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdKtsUavklarte.Name = "cmdUavklarte"
        $cmdKtsUavklarte.Caption = "Vis uavklarte"
        $cmdKtsUavklarte.Left = 290; $cmdKtsUavklarte.Top = 292; $cmdKtsUavklarte.Width = 160; $cmdKtsUavklarte.Height = 28

        Start-Sleep -Milliseconds 400
        $frmKts.CodeModule.AddFromString($formCodeKontaktsentralen)
        $frmKtsFrm = Join-Path $tempDir "frmKontaktsentralen.frm"
        $frmKts.Export($frmKtsFrm)

        # --- frmKtsOppsett (UserForm - innstillinger) ---
        $frmOpp = $tempWb.VBProject.VBComponents.Add(3)
        $frmOpp.Name = "frmKtsOppsett"
        $frmOpp.Properties("Caption").Value = "Kontaktsentralen - Oppsett"
        $frmOpp.Properties("Width").Value = 890
        $frmOpp.Properties("Height").Value = 400

        $lblOppTittel = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppTittel.Name = "lblTittel"
        $lblOppTittel.Caption = ""
        $lblOppTittel.Left = 10; $lblOppTittel.Top = 8; $lblOppTittel.Width = 420; $lblOppTittel.Height = 16

        $lblOppEpost = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppEpost.Name = "lblEpost"
        $lblOppEpost.Caption = "E-post-kolonner:"
        $lblOppEpost.Left = 10; $lblOppEpost.Top = 32; $lblOppEpost.Width = 420; $lblOppEpost.Height = 16

        $lstOppEpost = $frmOpp.Designer.Controls.Add("Forms.ListBox.1")
        $lstOppEpost.Name = "lstEpost"
        $lstOppEpost.Left = 10; $lstOppEpost.Top = 50; $lstOppEpost.Width = 420; $lstOppEpost.Height = 60
        $lstOppEpost.MultiSelect = 1  # fmMultiSelectMulti

        $lblOppGjenkjenning = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppGjenkjenning.Name = "lblGjenkjenning"
        $lblOppGjenkjenning.Caption = "Gjenkjenningskolonner (f.eks. skole-ID):"
        $lblOppGjenkjenning.Left = 10; $lblOppGjenkjenning.Top = 116; $lblOppGjenkjenning.Width = 420; $lblOppGjenkjenning.Height = 16

        $lstOppGjenkjenning = $frmOpp.Designer.Controls.Add("Forms.ListBox.1")
        $lstOppGjenkjenning.Name = "lstGjenkjenning"
        $lstOppGjenkjenning.Left = 10; $lstOppGjenkjenning.Top = 134; $lstOppGjenkjenning.Width = 420; $lstOppGjenkjenning.Height = 60
        $lstOppGjenkjenning.MultiSelect = 1

        $lblOppPostboks = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppPostboks.Name = "lblPostboks"
        $lblOppPostboks.Caption = "Delt postboks:"
        $lblOppPostboks.Left = 10; $lblOppPostboks.Top = 202; $lblOppPostboks.Width = 140; $lblOppPostboks.Height = 18

        $cboOppPostboks = $frmOpp.Designer.Controls.Add("Forms.ComboBox.1")
        $cboOppPostboks.Name = "cboPostboks"
        $cboOppPostboks.Left = 155; $cboOppPostboks.Top = 200; $cboOppPostboks.Width = 275; $cboOppPostboks.Height = 18
        $cboOppPostboks.Style = 2  # fmStyleDropDownList

        $lblOppStartpunkt = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppStartpunkt.Name = "lblStartpunkt"
        $lblOppStartpunkt.Caption = "Startpunkt (valgfritt):"
        $lblOppStartpunkt.Left = 10; $lblOppStartpunkt.Top = 226; $lblOppStartpunkt.Width = 140; $lblOppStartpunkt.Height = 18

        $txtOppStartpunkt = $frmOpp.Designer.Controls.Add("Forms.TextBox.1")
        $txtOppStartpunkt.Name = "txtStartpunkt"
        $txtOppStartpunkt.Left = 155; $txtOppStartpunkt.Top = 224; $txtOppStartpunkt.Width = 150; $txtOppStartpunkt.Height = 18

        $lblOppStartpunktHint = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppStartpunktHint.Name = "lblStartpunktHint"
        $lblOppStartpunktHint.Caption = "Format: dd.mm.åååå - tomt = skann alt som ligger i innboksen"
        $lblOppStartpunktHint.Left = 155; $lblOppStartpunktHint.Top = 244; $lblOppStartpunktHint.Width = 275; $lblOppStartpunktHint.Height = 14

        $lblOppNokkel = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppNokkel.Name = "lblNokkel"
        $lblOppNokkel.Caption = "Radnøkkel (unik ID per rad):"
        $lblOppNokkel.Left = 10; $lblOppNokkel.Top = 268; $lblOppNokkel.Width = 180; $lblOppNokkel.Height = 18

        $cboOppNokkel = $frmOpp.Designer.Controls.Add("Forms.ComboBox.1")
        $cboOppNokkel.Name = "cboNokkel"
        $cboOppNokkel.Left = 195; $cboOppNokkel.Top = 266; $cboOppNokkel.Width = 235; $cboOppNokkel.Height = 18
        $cboOppNokkel.Style = 2

        $lblOppNokkelHint = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppNokkelHint.Name = "lblNokkelHint"
        $lblOppNokkelHint.Caption = "Kolonnen som gjør hver rad unik (f.eks. skole-ID). Da kan tabellen sorteres og filtreres i Excel uten at e-post havner på feil rad."
        $lblOppNokkelHint.Left = 10; $lblOppNokkelHint.Top = 290; $lblOppNokkelHint.Width = 420; $lblOppNokkelHint.Height = 28
        $lblOppNokkelHint.WordWrap = $true

        $lblOppRespons = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppRespons.Name = "lblRespons"
        $lblOppRespons.Caption = "Kontrollkolonne i tabellen:"
        $lblOppRespons.Left = 10; $lblOppRespons.Top = 324; $lblOppRespons.Width = 180; $lblOppRespons.Height = 18

        $cboOppRespons = $frmOpp.Designer.Controls.Add("Forms.ComboBox.1")
        $cboOppRespons.Name = "cboRespons"
        $cboOppRespons.Left = 195; $cboOppRespons.Top = 322; $cboOppRespons.Width = 235; $cboOppRespons.Height = 18
        $cboOppRespons.Style = 2

        $lblOppKolonneHint = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppKolonneHint.Name = "lblKolonneHint"
        $lblOppKolonneHint.Caption = "Kontrollkolonnen fylles automatisk med kategorien som har høyest prioritet blant e-postene som er matchet til raden (se prioritetsliste under)."
        $lblOppKolonneHint.Left = 10; $lblOppKolonneHint.Top = 346; $lblOppKolonneHint.Width = 420; $lblOppKolonneHint.Height = 28
        $lblOppKolonneHint.WordWrap = $true

        # Tabellvelger oppe til hoyre (nar arbeidsboka har flere tabeller).
        $lblOppTabell = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppTabell.Name = "lblTabell"
        $lblOppTabell.Caption = "Tabell i arbeidsboka:"
        $lblOppTabell.Left = 450; $lblOppTabell.Top = 8; $lblOppTabell.Width = 120; $lblOppTabell.Height = 16
        $cboOppTabell = $frmOpp.Designer.Controls.Add("Forms.ComboBox.1")
        $cboOppTabell.Name = "cboTabell"
        $cboOppTabell.Left = 575; $cboOppTabell.Top = 6; $cboOppTabell.Width = 295; $cboOppTabell.Height = 18
        $cboOppTabell.Style = 2
        $cboOppTabell.ColumnCount = 2
        $cboOppTabell.ColumnWidths = "170 pt;115 pt"

        # --- Hoyre kolonne (x = 450): kolonner som vises i lista + kategoriprioritet ---
        $lblOppVisning = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppVisning.Name = "lblVisning"
        $lblOppVisning.Caption = "Kolonner som vises i lista (maks 4). Huk av, og bruk Opp/Ned for rekkefølgen:"
        $lblOppVisning.Left = 450; $lblOppVisning.Top = 32; $lblOppVisning.Width = 420; $lblOppVisning.Height = 28
        $lblOppVisning.WordWrap = $true
        $lstOppVisning = $frmOpp.Designer.Controls.Add("Forms.ListBox.1")
        $lstOppVisning.Name = "lstVisning"
        $lstOppVisning.Left = 450; $lstOppVisning.Top = 62; $lstOppVisning.Width = 320; $lstOppVisning.Height = 130
        $lstOppVisning.MultiSelect = 1
        $lstOppVisning.ListStyle = 1  # fmListStyleOption: avkrysningsbokser
        $visOpp = $frmOpp.Designer.Controls.Add("Forms.CommandButton.1"); $visOpp.Name="cmdVisOpp"; $visOpp.Caption="Opp"; $visOpp.Left=780; $visOpp.Top=67; $visOpp.Width=90; $visOpp.Height=25
        $visNed = $frmOpp.Designer.Controls.Add("Forms.CommandButton.1"); $visNed.Name="cmdVisNed"; $visNed.Caption="Ned"; $visNed.Left=780; $visNed.Top=97; $visNed.Width=90; $visNed.Height=25

        $lblOppKat = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblOppKat.Name = "lblKategorier"
        $lblOppKat.Caption = "Outlook-kategorier - øverst har høyest prioritet i Kontrollkolonnen:"
        $lblOppKat.Left = 450; $lblOppKat.Top = 202; $lblOppKat.Width = 420; $lblOppKat.Height = 16
        $lstOppKat = $frmOpp.Designer.Controls.Add("Forms.ListBox.1")
        $lstOppKat.Name = "lstKategorier"
        $lstOppKat.Left = 450; $lstOppKat.Top = 220; $lstOppKat.Width = 320; $lstOppKat.Height = 100
        $katOpp = $frmOpp.Designer.Controls.Add("Forms.CommandButton.1"); $katOpp.Name="cmdKatOpp"; $katOpp.Caption="Opp"; $katOpp.Left=780; $katOpp.Top=225; $katOpp.Width=90; $katOpp.Height=25
        $katNed = $frmOpp.Designer.Controls.Add("Forms.CommandButton.1"); $katNed.Name="cmdKatNed"; $katNed.Caption="Ned"; $katNed.Left=780; $katNed.Top=255; $katNed.Width=90; $katNed.Height=25

        $cmdOppLagre = $frmOpp.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdOppLagre.Name = "cmdLagre"
        $cmdOppLagre.Caption = "Lagre"
        $cmdOppLagre.Left = 330; $cmdOppLagre.Top = 390; $cmdOppLagre.Width = 110; $cmdOppLagre.Height = 28

        $cmdOppTilbake = $frmOpp.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdOppTilbake.Name = "cmdTilbake"
        $cmdOppTilbake.Caption = "Tilbake"
        $cmdOppTilbake.Left = 450; $cmdOppTilbake.Top = 390; $cmdOppTilbake.Width = 110; $cmdOppTilbake.Height = 28

        Start-Sleep -Milliseconds 400
        $frmOpp.CodeModule.AddFromString($formCodeKtsOppsett)
        $frmOppFrm = Join-Path $tempDir "frmKtsOppsett.frm"
        $frmOpp.Export($frmOppFrm)

        # --- frmKtsRad (UserForm - detaljside per rad) ---
        $frmRad = $tempWb.VBProject.VBComponents.Add(3)
        $frmRad.Name = "frmKtsRad"
        $frmRad.Properties("Caption").Value = "Kontaktsentralen - Radkontroll"
        $frmRad.Properties("Width").Value = 760
        $frmRad.Properties("Height").Value = 430
        $lblRadT = $frmRad.Designer.Controls.Add("Forms.Label.1"); $lblRadT.Name="lblTittel"; $lblRadT.Left=10; $lblRadT.Top=8; $lblRadT.Width=720; $lblRadT.Height=20
        foreach ($hNr in 1..5) { $h = $frmRad.Designer.Controls.Add("Forms.Label.1"); $h.Name="lblHdr$hNr"; $h.Caption="Kolonne $hNr"; $h.Left=10; $h.Top=34; $h.Width=100; $h.Height=14 }
        $lstRad = $frmRad.Designer.Controls.Add("Forms.ListBox.1"); $lstRad.Name="lstMeldinger"; $lstRad.Left=10; $lstRad.Top=50; $lstRad.Width=720; $lstRad.Height=150; $lstRad.ColumnCount=5; $lstRad.ColumnWidths="88 pt;140 pt;190 pt;115 pt;150 pt"
        $lblKat = $frmRad.Designer.Controls.Add("Forms.Label.1"); $lblKat.Name="lblKategorier"; $lblKat.Caption="Velg en e-post i lista for å endre kategoriene."; $lblKat.Left=10; $lblKat.Top=210; $lblKat.Width=720; $lblKat.Height=18
        foreach ($katNr in 1..30) { $kb = $frmRad.Designer.Controls.Add("Forms.CommandButton.1"); $kb.Name="cmdKat$katNr"; $kb.Caption="Kategori $katNr"; $kb.Left=10; $kb.Top=232; $kb.Width=90; $kb.Height=24; $kb.Visible=$false }
        $b1=$frmRad.Designer.Controls.Add("Forms.CommandButton.1"); $b1.Name="cmdLagreKategori"; $b1.Caption="Lagre i Outlook"; $b1.Left=10; $b1.Top=300; $b1.Width=130; $b1.Height=28
        $b6=$frmRad.Designer.Controls.Add("Forms.CommandButton.1"); $b6.Name="cmdAngre"; $b6.Caption="Angre"; $b6.Left=146; $b6.Top=300; $b6.Width=90; $b6.Height=28
        $b2=$frmRad.Designer.Controls.Add("Forms.CommandButton.1"); $b2.Name="cmdAapneOutlook"; $b2.Caption="Åpne e-post i Outlook"; $b2.Left=250; $b2.Top=300; $b2.Width=150; $b2.Height=28
        $b5=$frmRad.Designer.Controls.Add("Forms.CommandButton.1"); $b5.Name="cmdMatchUavklart"; $b5.Caption="Match uavklart e-post hit"; $b5.Left=410; $b5.Top=300; $b5.Width=190; $b5.Height=28
        $b4=$frmRad.Designer.Controls.Add("Forms.CommandButton.1"); $b4.Name="cmdHovedmeny"; $b4.Caption="Hovedmeny"; $b4.Left=610; $b4.Top=300; $b4.Width=120; $b4.Height=28
        Start-Sleep -Milliseconds 300
        $frmRad.CodeModule.AddFromString($formCodeKtsRad)
        $frmRadFrm = Join-Path $tempDir "frmKtsRad.frm"
        $frmRad.Export($frmRadFrm)

        # --- frmKtsUavklart (UserForm - match uavklarte e-poster til rad) ---
        $frmUav = $tempWb.VBProject.VBComponents.Add(3)
        $frmUav.Name = "frmKtsUavklart"
        $frmUav.Properties("Caption").Value = "Kontaktsentralen - Uavklarte"
        $frmUav.Properties("Width").Value = 660
        $frmUav.Properties("Height").Value = 430

        $lblUavTittel = $frmUav.Designer.Controls.Add("Forms.Label.1")
        $lblUavTittel.Name = "lblTittel"
        $lblUavTittel.Caption = ""
        $lblUavTittel.Left = 10; $lblUavTittel.Top = 8; $lblUavTittel.Width = 400; $lblUavTittel.Height = 20

        $lblUavUavklarteLabel = $frmUav.Designer.Controls.Add("Forms.Label.1")
        $lblUavUavklarteLabel.Name = "lblUavklarteLabel"
        $lblUavUavklarteLabel.Caption = "Uavklarte e-poster fra siste skanning (dobbeltklikk for å åpne e-posten i Outlook):"
        $lblUavUavklarteLabel.Left = 10; $lblUavUavklarteLabel.Top = 34; $lblUavUavklarteLabel.Width = 620; $lblUavUavklarteLabel.Height = 16

        $lstUavUavklarte = $frmUav.Designer.Controls.Add("Forms.ListBox.1")
        $lstUavUavklarte.Name = "lstUavklarte"
        $lstUavUavklarte.Left = 10; $lstUavUavklarte.Top = 52; $lstUavUavklarte.Width = 620; $lstUavUavklarte.Height = 120
        $lstUavUavklarte.ColumnCount = 3
        $lstUavUavklarte.ColumnWidths = "260 pt;180 pt;120 pt"
        $lstUavUavklarte.MultiSelect = 0

        $lblUavRaderLabel = $frmUav.Designer.Controls.Add("Forms.Label.1")
        $lblUavRaderLabel.Name = "lblRaderLabel"
        $lblUavRaderLabel.Caption = "Velg hvilken rad e-posten faktisk hører til:"
        $lblUavRaderLabel.Left = 10; $lblUavRaderLabel.Top = 178; $lblUavRaderLabel.Width = 240; $lblUavRaderLabel.Height = 16

        # Sokefelt pa samme linje som overskriften over radlista. Pa skjermer
        # med annen skalering (blandet DPI) forskyves vanlige kontroller ca. 20 %
        # nedover i forhold til listene - derfor er det god luft (ca. 60 pt) mellom
        # listene, og feltet ligger i venstre halvdel der forskyvningen er liten.
        $lblUavSok = $frmUav.Designer.Controls.Add("Forms.Label.1")
        $lblUavSok.Name = "lblSok"
        $lblUavSok.Caption = "Søk rad:"
        $lblUavSok.Left = 258; $lblUavSok.Top = 178; $lblUavSok.Width = 45; $lblUavSok.Height = 16
        $txtUavSok = $frmUav.Designer.Controls.Add("Forms.TextBox.1")
        $txtUavSok.Name = "txtSok"
        $txtUavSok.Left = 306; $txtUavSok.Top = 175; $txtUavSok.Width = 204; $txtUavSok.Height = 19

        $lstUavRaderMatch = $frmUav.Designer.Controls.Add("Forms.ListBox.1")
        $lstUavRaderMatch.Name = "lstRaderMatch"
        $lstUavRaderMatch.Left = 10; $lstUavRaderMatch.Top = 242; $lstUavRaderMatch.Width = 620; $lstUavRaderMatch.Height = 120
        $lstUavRaderMatch.ColumnCount = 3
        $lstUavRaderMatch.ColumnWidths = "220 pt;150 pt;180 pt"
        $lstUavRaderMatch.MultiSelect = 0

        $lblUavStatus = $frmUav.Designer.Controls.Add("Forms.Label.1")
        $lblUavStatus.Name = "lblStatus"
        $lblUavStatus.Caption = ""
        $lblUavStatus.Left = 10; $lblUavStatus.Top = 368; $lblUavStatus.Width = 350; $lblUavStatus.Height = 16

        $cmdUavMatch = $frmUav.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdUavMatch.Name = "cmdMatch"
        $cmdUavMatch.Caption = "Match valgt e-post til valgt rad"
        $cmdUavMatch.Left = 10; $cmdUavMatch.Top = 390; $cmdUavMatch.Width = 220; $cmdUavMatch.Height = 26

        $cmdUavTilbake = $frmUav.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdUavTilbake.Name = "cmdTilbake"
        $cmdUavTilbake.Caption = "Tilbake til dashbord"
        $cmdUavTilbake.Left = 500; $cmdUavTilbake.Top = 390; $cmdUavTilbake.Width = 130; $cmdUavTilbake.Height = 26

        Start-Sleep -Milliseconds 400
        $frmUav.CodeModule.AddFromString($formCodeKtsUavklart)
        $frmUavFrm = Join-Path $tempDir "frmKtsUavklart.frm"
        $frmUav.Export($frmUavFrm)

    } finally {
        if ($tempWb) { $tempWb.Close($false) }
        if ($tempExcel) {
            $tempExcel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($tempExcel) | Out-Null
        }
    }

    # --- Importer/oppdater i malfila ---
    # Fjern dobbeltklikk-blokken FORST (den refererer modKontaktsentralen),
    # og sett den inn igjen SIST, etter at modulen er importert pa nytt.
    try { Remove-DobbeltklikkFraThisWorkbook } catch {}
    foreach ($navn in @("modKontaktsentralen", "frmKontaktsentralen", "frmKtsOppsett", "frmKtsUavklart", "frmKtsRad")) {
        try {
            $eksisterende = $vbProject.VBComponents.Item($navn)
            $vbProject.VBComponents.Remove($eksisterende)
        } catch {}
    }
    $vbProject.VBComponents.Import($modBas) | Out-Null
    $vbProject.VBComponents.Import($frmKtsFrm) | Out-Null
    $vbProject.VBComponents.Import($frmRadFrm) | Out-Null
    $vbProject.VBComponents.Import($frmOppFrm) | Out-Null
    $vbProject.VBComponents.Import($frmUavFrm) | Out-Null

    # Gi Excel litt tid til a fullfore registreringen av de nettopp
    # importerte VBA-komponentene for vi setter en Button sin OnAction
    # (som ma kunne resolve makronavnet) - uten denne pausen kan
    # tilordningen feile med en COM-feil rett etter en fersk import.
    Start-Sleep -Milliseconds 500

    # Ingen automatisk ark-knapp lenger, verken ved fersk installasjon eller
    # oppdatering (Håkons eksplisitte onske 2026-09-22: en makro legger seg
    # bare til - en knapp/inngang i arket er et eget, manuelt valg brukeren
    # gjor selv, typisk via Makromeny). $erFerskInstall brukes fortsatt andre
    # steder i dette scriptet og er ikke fjernet.

    try { Add-DobbeltklikkTilThisWorkbook } catch {
        Write-Output "  ADVARSEL: kunne ikke sette opp dobbeltklikk i Kontrollkolonnen: $($_.Exception.Message)"
    }

    Write-Output "  modKontaktsentralen installert"
    Write-Output "  frmKontaktsentralen installert"
    Write-Output "  frmKtsOppsett installert"
    Write-Output "  frmKtsUavklart installert"
    Write-Output "  frmKtsRad installert"

    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ""
Write-Output "Ferdig. Husk aa lagre filen selv (Ctrl+S) naar du er klar."
Write-Output "Trykk 'Kontaktsentralen' pa arket. Forste gang apnes Innstillinger automatisk. Velg delt postboks, matchingkolonner, Kontrollkolonne og kategoriprioritet."
Write-Output "Dobbeltklikk en celle i Kontrollkolonnen i tabellen for a apne radsiden for den raden."
