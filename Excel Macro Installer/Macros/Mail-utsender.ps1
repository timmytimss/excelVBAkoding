param(
    [string]$Path,
    [string]$Mappe,
    [string]$MalSti,
    [string]$MalUndermappe = "",
    [string]$MalFilnavn = "",
    [switch]$Force,
    [switch]$Uninstall,
    # Kun for testing: start en EGEN Excel-prosess i stedet for a koble til den
    # som allerede kjorer. Samme monster som Kontaktsentralen.ps1.
    [switch]$NyExcelInstans
)

$ErrorActionPreference = 'Stop'

# Bump denne ved ENHVER endring i VBA-kilden under (moduleCodeTemplate/
# formCodeSendMail/formCodeOppsett) - se excel-vba-installer-pattern i
# minnet for hvorfor: uten en bump gjor installeren ingenting pa neste
# kjoring selv om koden faktisk er endret.
$GeneriskVersion = "0.20.1"

# ============================================================
# Mail-utsender - alt-i-en installer/oppdaterer.
# (Tidligere kalt "Generisk Mail-utsendelse" under utvikling - omdopt
# 2026-09-08 idet den erstatter det TIMSS-spesifikke "Massutsendelse V1".)
#
# FORSTE UTKAST (0.1.0-draft) - dette er en generalisert versjon av
# "Massutsendelse V1.ps1" (TIMSS-mail-verktoyet i naboprosjektet).
# Forskjellen: INGENTING er hardkodet til et bestemt regneark-oppsett.
# Bade hvilke kolonner som er e-post/status, OG hvilke plassholdere
# (AAA, BBB, CCC ...) i mal-fila som skal byttes ut med hvilken
# kolonne, velges av brukeren i et eget "Oppsett"-vindu i selve
# Excel-fila (frmOppsett) - og lagres i en skjult fane i arbeidsboka
# ("MailKonfig"), sa verktoyet kan gjenbrukes pa hvilken som helst
# Excel-fil + mal-fil-kombinasjon, ikke bare TIMSS-hovedfila.
#
# Plassholdere i mal-fila skrives som samme bokstav gjentatt 2+
# ganger (AAA, BBB, CCC, DD, ...) - fritt antall og valgfrie
# bokstaver, koden finner dem automatisk ved a lese mal-fila.
#
# Samme grunnmonster som Massutsendelse V1.ps1:
#  - Bygger VBA-komponenter i en midlertidig arbeidsbok, eksporterer
#    dem, og importerer i mal-filen. Ingen andre filer trengs.
#  - Excel-fila (.xlsm) VELGES ALLTID BEVISST via filvelger-dialog -
#    aldri automatisk gjetting. Mal-fila (.msg) derimot velges IKKE
#    her i installeren lenger (fjernet 2026-09-08) - det gjores
#    utelukkende via "Oppsett"-knappen inne i selve Excel-fila etter
#    installasjon (samme relative-sti-utregning, bare i VBA i stedet
#    for PowerShell). -MalSti finnes fortsatt som parameter for
#    skriptet/automatisert bruk, uten a apne noen dialog.
#  - Portabel sti-hakndtering (fungerer bade lokalt og pa
#    OneDrive/SharePoint, se LokalStiFraSky i modGenerisk).
#  - Test-modus (send alt til en testadresse i stedet for ekte mottakere)
#    fantes tidligere, men er fjernet igjen 2026-09-10 pa Håkons eksplisitte
#    onske - verktoyet er na ferdig utprovd og i fast produksjonsbruk, og
#    trenger ikke lenger dette sikkerhetsnettet.
#
# Bruk:
#   .\Mail-utsender.ps1
# ============================================================

# ---------------------------------------------------------------
# VBA-kildekode: modGenerisk (standard-modul)
# ---------------------------------------------------------------
$moduleCodeTemplate = @'
Attribute VB_Name = "modGenerisk"
Option Explicit

' ============================================================
' Mail-utsender
' Installeres/oppdateres av Mail-utsender.ps1 (se den fila for hele
' installer-monsteret). Erstatter det TIMSS-spesifikke
' "Massutsendelse V1.ps1" - se generisk-mail-verktoy i minnet for
' bakgrunnen.
'
' Forskjellen fra TIMSS-verktoyet: INGEN kolonnenavn eller
' plassholdere er hardkodet i koden. Alt konfigureres av brukeren
' i "Oppsett"-vinduet (frmOppsett) og lagres i en skjult fane i
' arbeidsboka ("MailKonfig"), sa denne .ps1-fila kan installeres
' pa hvilken som helst Excel-fil + mal-fil-kombinasjon.
'
' Plassholdere i mal-fila (emnefelt og/eller brodtekst) skrives som
' samme bokstav gjentatt 2 eller flere ganger, f.eks. AAA, BBB, CCC,
' DD - fritt antall og valgfrie bokstaver, koden finner dem
' automatisk (se FinnPlassholdere).
' ============================================================

Public Const GENERISK_VERSION As String = "{{VERSION}}"

Private Const KONFIG_FANENAVN As String = "MailKonfig"
Public Const MAKS_PLASSHOLDERE As Long = 8

' Husker hvilket vindu Oppsett ble apnet fra ("frmSendMail" eller
' "frmMasseSend"), sa "Lagre"-knappen i Oppsett kan sende brukeren
' TILBAKE dit, i stedet for alltid a hoppe til hovedmenyen. Et Public
' modulniva-variabel er et vanlig, trygt monster for slik "hvilket
' vindu apnet dette"-kommunikasjon mellom skjemaer i VBA.
Public OppsettReturnerTil As String
Public FeilsjekkReturnerTil As String

' --- Portabel sti-hakndtering (OneDrive/SharePoint) -----------------

Private Function DekodUrlTegn(ByVal s As String) As String
    ' Oversetter %20 osv. tilbake til vanlige tegn, i tilfelle den
    ' lokale OneDrive-registreringen skulle inneholde URL-koding.
    Dim i As Long, kode As String, res As String
    i = 1
    Do While i <= Len(s)
        If Mid(s, i, 1) = "%" And i + 2 <= Len(s) Then
            kode = Mid(s, i + 1, 2)
            res = res & Chr(CLng("&H" & kode))
            i = i + 3
        Else
            res = res & Mid(s, i, 1)
            i = i + 1
        End If
    Loop
    DekodUrlTegn = res
End Function

Public Function LokalStiFraSky(ByVal sti As String) As String
    ' Nar en arbeidsbok ligger pa OneDrive/SharePoint (og Autolagre er
    ' PA), rapporterer Excel ofte ThisWorkbook.Path som en nettadresse
    ' (https://...) i stedet for en lokal filsti - selv om filen OGSA
    ' finnes lokalt synkronisert. Denne funksjonen oversetter URL-en
    ' til den lokale, synkroniserte mappen ved a sla opp OneDrive sin
    ' egen liste over synkroniserte biblioteker i registeret. Kopiert
    ' uendret fra "Excel Kontaktskjema mail-generering"-prosjektet -
    ' se minnet "vba-onedrive-url-path-bug" der for hele bakgrunnen og
    ' hvorfor den er bygget akkurat slik (bl.a. hvorfor MountPoint sitt
    ' SISTE mappenavn ma sammenlignes som strengsuffiks, ikke som
    ' siste "\"-ledd).
    Const HKEY_CURRENT_USER As Long = &H80000001
    Dim reg As Object
    Dim baseKey As String
    Dim underNokler As Variant
    Dim i As Long
    Dim urlNamespace As String, mountPoint As String
    Dim bestUrl As String, bestMount As String, bestLen As Long
    Dim rest As String

    LokalStiFraSky = sti

    If LCase(Left(sti, 4)) <> "http" Then Exit Function

    On Error GoTo Ferdig
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    baseKey = "Software\SyncEngines\Providers\OneDrive"

    reg.EnumKey HKEY_CURRENT_USER, baseKey, underNokler
    If IsEmpty(underNokler) Then GoTo Ferdig

    bestLen = 0
    For i = LBound(underNokler) To UBound(underNokler)
        urlNamespace = ""
        mountPoint = ""
        reg.GetStringValue HKEY_CURRENT_USER, baseKey & "\" & underNokler(i), "UrlNamespace", urlNamespace
        reg.GetStringValue HKEY_CURRENT_USER, baseKey & "\" & underNokler(i), "MountPoint", mountPoint
        urlNamespace = DekodUrlTegn(urlNamespace)
        If Len(urlNamespace) > 0 And Len(mountPoint) > 0 Then
            If Len(sti) >= Len(urlNamespace) Then
                If LCase(Left(sti, Len(urlNamespace))) = LCase(urlNamespace) Then
                    If Len(urlNamespace) > bestLen Then
                        bestLen = Len(urlNamespace)
                        bestUrl = urlNamespace
                        bestMount = mountPoint
                    End If
                End If
            End If
        End If
    Next i

    If bestLen > 0 Then
        rest = Mid(sti, bestLen + 1)
        rest = DekodUrlTegn(rest)
        rest = Replace(rest, "/", "\")

        Dim restUtenSkrastrek As String, forsteSegment As String, posSkille As Long
        restUtenSkrastrek = rest
        If Left(restUtenSkrastrek, 1) = "\" Then restUtenSkrastrek = Mid(restUtenSkrastrek, 2)

        posSkille = InStr(restUtenSkrastrek, "\")
        If posSkille > 0 Then
            forsteSegment = Left(restUtenSkrastrek, posSkille - 1)
        Else
            forsteSegment = restUtenSkrastrek
        End If

        If Len(forsteSegment) > 0 And Len(forsteSegment) <= Len(bestMount) Then
            If LCase(Right(bestMount, Len(forsteSegment))) = LCase(forsteSegment) Then
                rest = Mid(restUtenSkrastrek, Len(forsteSegment) + 1)
            End If
        End If

        ' Sikre at det alltid star en "\" mellom MountPoint og resten av
        ' stien. UrlNamespace i registeret kan ha ELLER mangle en
        ' avsluttende "/" avhengig av hvordan akkurat DEN synk-
        ' registreringen ble opprettet - det avgjor om "rest" over
        ' allerede har fatt med seg skilletegnet fra url-en eller ikke.
        ' Uten denne sjekken kan stien bli feil sammensatt uten "\" i
        ' det hele tatt (f.eks. "...OsloDokumenter..." i stedet for
        ' "...Oslo\Dokumenter...").
        If Len(rest) > 0 And Left(rest, 1) <> "\" Then
            rest = "\" & rest
        End If

        LokalStiFraSky = bestMount & rest
    End If

Ferdig:
    On Error GoTo 0
End Function

' --- Konfig-lesing/skriving (skjult fane "MailKonfig") ----------------

Private Function KonfigFane() As Worksheet
    On Error Resume Next
    Set KonfigFane = ThisWorkbook.Worksheets(KONFIG_FANENAVN)
    On Error GoTo 0
    If KonfigFane Is Nothing Then
        Set KonfigFane = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        KonfigFane.Name = KONFIG_FANENAVN
        KonfigFane.Range("A1").Value = "Nokkel"
        KonfigFane.Range("B1").Value = "Verdi"
        KonfigFane.Visible = xlSheetVeryHidden
    End If
End Function

Public Function LesKonfig() As Object
    Dim ws As Worksheet, dict As Object, r As Long
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = KonfigFane()
    r = 2
    Do While ws.Cells(r, 1).Value <> ""
        dict(CStr(ws.Cells(r, 1).Value)) = CStr(ws.Cells(r, 2).Value)
        r = r + 1
    Loop
    Set LesKonfig = dict
End Function

Public Sub SkrivKonfigVerdi(ByVal nokkel As String, ByVal verdi As String)
    Dim ws As Worksheet, r As Long
    Set ws = KonfigFane()
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

' --- Plassholder-gjenkjenning -------------------------------------

Public Function FinnPlassholdere(ByVal tekst As String) As Collection
    ' Finner alle distinkte "samme bokstav gjentatt 2+ ganger"-tokens
    ' i teksten (AAA, BBB, CCC, DD, ...), i den rekkefolgen de forst
    ' dukker opp. Bruker en Collection (ikke et array) nettopp for a
    ' unnga arrays-med-variabel-storrelse-problematikken som stoppet
    ' den forrige avanserte versjonen av dette verktoyet.
    Dim resultat As New Collection
    Dim re As Object, treff As Object, m As Object
    Dim funnet As Object
    Set funnet = CreateObject("Scripting.Dictionary")

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.Pattern = "\b([A-Z])\1+\b"

    Set treff = re.Execute(tekst)
    For Each m In treff
        If Not funnet.Exists(m.Value) Then
            funnet(m.Value) = True
            resultat.Add m.Value
        End If
    Next m

    Set FinnPlassholdere = resultat
End Function

' --- Lagrede avsender-adresser (for "Send fra"-nedtrekksmenyen i Oppsett) --

Public Function HentLagredeAvsendere() As Collection
    Dim resultat As New Collection
    Dim konfig As Object, nokkel As Variant
    Set konfig = LesKonfig()
    For Each nokkel In konfig.Keys
        If Left(nokkel, 9) = "AVSENDER:" Then
            resultat.Add konfig(nokkel)
        End If
    Next nokkel
    Set HentLagredeAvsendere = resultat
End Function

' --- Tabell-hjelpefunksjoner ---------------------------------------

Public Function AktivTabell() As ListObject
    Dim ws As Worksheet, tbl As ListObject
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> KONFIG_FANENAVN Then
            For Each tbl In ws.ListObjects
                Set AktivTabell = tbl
                Exit Function
            Next tbl
        End If
    Next ws
End Function

Public Function KolonneNr(ByVal tbl As ListObject, ByVal kolonnenavn As String) As Long
    Dim kol As ListColumn
    If kolonnenavn = "" Then Exit Function
    For Each kol In tbl.ListColumns
        If kol.Name = kolonnenavn Then
            KolonneNr = kol.Index
            Exit Function
        End If
    Next kol
End Function

Public Function SynligeKolonneIndekser(ByVal tbl As ListObject) As Collection
    ' Respekterer skjulte kolonner (f.eks. fra Kolonnevelger-makroen) OG
    ' Mail-utsenders egen "Velg kolonner"-liste (frmKolonnevisning) - mail-
    ' listene viser da bare de kolonnene brukeren faktisk har valgt a se,
    ' i stedet for alltid a vise absolutt alt.
    Dim resultat As New Collection
    Dim kol As ListColumn
    For Each kol In tbl.ListColumns
        If Not kol.Range.EntireColumn.Hidden Then
            If Not KolonneErSkjultILista(kol.Name) Then
                resultat.Add kol.Index
            End If
        End If
    Next kol
    Set SynligeKolonneIndekser = resultat
End Function

Public Function KolonneErSkjultILista(ByVal kolonnenavn As String) As Boolean
    ' Config-nokkel SKJULTE_LISTEKOLONNER: "|"-separert liste med
    ' kolonnenavn brukeren har valgt a SKJULE fra utvalgslistene (ikke fra
    ' selve arket). En kolonne som er FJERNET/omdopt i tabellen siden sist
    ' matcher rett og slett ingenting her lenger - ingen feilhandtering
    ' nodvendig, den forsvinner bare stille fra den gamle lista.
    Dim konfig As Object, liste As String
    Set konfig = LesKonfig()
    If Not konfig.Exists("SKJULTE_LISTEKOLONNER") Then Exit Function
    liste = konfig("SKJULTE_LISTEKOLONNER")
    If liste = "" Then Exit Function
    KolonneErSkjultILista = (InStr(1, "|" & liste & "|", "|" & kolonnenavn & "|") > 0)
End Function

' --- Ressursmappe / mal-fil (lest fra konfig, ikke hardkodet) --------

Public Function RessursMappe() As String
    Dim konfig As Object, undermappe As String
    Set konfig = LesKonfig()
    undermappe = ""
    If konfig.Exists("MAL_UNDERMAPPE") Then undermappe = konfig("MAL_UNDERMAPPE")
    If undermappe = "" Then
        RessursMappe = LokalStiFraSky(ThisWorkbook.Path)
    Else
        RessursMappe = LokalStiFraSky(ThisWorkbook.Path) & "\" & undermappe
    End If
End Function

Public Function MalSti() As String
    Dim konfig As Object, filnavn As String
    Set konfig = LesKonfig()
    filnavn = ""
    If konfig.Exists("MAL_FILNAVN") Then filnavn = konfig("MAL_FILNAVN")
    If filnavn = "" Then
        MalSti = ""
    Else
        MalSti = RessursMappe() & "\" & filnavn
    End If
End Function

' --- Inngangspunkter (knappene pa arket) ------------------------------

Sub MailUtsendelse()
    Dim konfig As Object
    Set konfig = LesKonfig()
    If Not konfig.Exists("KOL_EPOST") Or Not konfig.Exists("KOL_STATUS") Or MalSti() = "" Then
        MsgBox "Oppsettet er ikke fullført ennå." & vbCrLf & _
               "Åpner ""Oppsett"" - velg mal-fil og kolonner først.", vbInformation, "Mail-utsendelse"
        ' Ma settes eksplisitt her - ellers kan OppsettReturnerTil sta igjen
        ' med en gammel verdi fra en TIDLIGERE navigering i samme Excel-
        ' okt (f.eks. "frmMasseSend"), og Lagre/Tilbake i Oppsett ville da
        ' feilaktig prove a apne DEN i stedet for frmSendMail.
        modGenerisk.OppsettReturnerTil = "frmSendMail"
        frmOppsett.Show
        Exit Sub
    End If
    frmSendMail.Show
End Sub

Sub VisOppsett()
    frmOppsett.Show
End Sub

' --- Outlook-hjelpefunksjoner (samme monster som TIMSS-verktoyet) --

Public Function HentOutlookApp() As Object
    On Error Resume Next
    Set HentOutlookApp = GetObject(, "Outlook.Application")
    If HentOutlookApp Is Nothing Then
        Set HentOutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
End Function

Private Function SettAvsender(ByVal olApp As Object, ByVal olMail As Object, ByVal smtpAdresse As String) As Boolean
    Dim acc As Object
    For Each acc In olApp.Session.Accounts
        If LCase(acc.SmtpAddress) = LCase(smtpAdresse) Then
            Set olMail.SendUsingAccount = acc
            On Error Resume Next
            olMail.SentOnBehalfOfName = acc.SmtpAddress
            On Error GoTo 0
            SettAvsender = True
            Exit Function
        End If
    Next acc

    ' Fant ingen EGEN konto med denne adressen - vanlig og helt normalt
    ' for delte postbokser/"Send som"-tilgang (se mail_sender_identity
    ' i naboprosjektets minne for hele bakgrunnen). Sett Fra-feltet
    ' direkte pa adressen likevel.
    On Error Resume Next
    olMail.SentOnBehalfOfName = smtpAdresse
    On Error GoTo 0
    SettAvsender = False
End Function

Private Function ErProsjektVerktoyfil(ByVal fil As String) As Boolean
    ' Nar mal-fila ligger i SAMME mappe som selve .xlsm-fila (ikke en
    ' egen ressurs-undermappe) - noe dette verktoyet, i motsetning til
    ' TIMSS-verktoyet, faktisk tillater - havner ofte prosjektfiler som
    ' IKKE skal sendes til mottakere (installer-skriptet, referanse-
    ' kopier, Excel sine midlertidige laasefiler) i samme mappe som
    ' skannes for vedlegg. Utelukk dem her.
    Dim ext As String
    ext = LCase(Mid(fil, InStrRev(fil, ".") + 1))
    Select Case ext
        Case "ps1", "bas", "frm", "cls", "xlsm", "xlsx", "xlsb", "xls"
            ErProsjektVerktoyfil = True
        Case Else
            ErProsjektVerktoyfil = (Left(fil, 2) = "~$")
    End Select
End Function

Public Function HentVedleggsliste() As Collection
    ' Eneste kilde til sannhet for "hvilke filer blir lagt ved" - brukt
    ' bade av selve utsendelsen (LeggVedVedlegg under) OG av
    ' forhandsvisningene i Oppsett/frmSendMail/frmMasseSend, sa det som
    ' vises aldri kan avvike fra det som faktisk skjer.
    Dim resultat As New Collection
    Dim mappe As String, fil As String, malFilnavn As String
    Dim konfig As Object
    Set konfig = LesKonfig()
    malFilnavn = ""
    If konfig.Exists("MAL_FILNAVN") Then malFilnavn = konfig("MAL_FILNAVN")

    mappe = RessursMappe()
    On Error Resume Next
    fil = Dir(mappe & "\*.*")
    On Error GoTo 0
    Do While fil <> ""
        If LCase(fil) <> LCase(malFilnavn) And Not ErProsjektVerktoyfil(fil) Then
            resultat.Add fil
        End If
        fil = Dir()
    Loop

    Set HentVedleggsliste = resultat
End Function

Private Sub LeggVedVedlegg(ByVal olMail As Object)
    Dim mappe As String, fil As Variant
    mappe = RessursMappe()
    For Each fil In HentVedleggsliste()
        olMail.Attachments.Add mappe & "\" & fil
    Next fil
End Sub

Public Function VedleggTekst() As String
    ' Klar-til-a-vise tekst for forhandsvisninger - en fil per linje,
    ' eller en tydelig melding hvis ressursmappen mangler/er tom.
    Dim liste As Collection, fil As Variant, tekst As String

    On Error Resume Next
    Set liste = HentVedleggsliste()
    On Error GoTo 0

    If liste Is Nothing Then
        VedleggTekst = "(fant ikke ressursmappen)"
        Exit Function
    End If

    If liste.Count = 0 Then
        VedleggTekst = "(ingen vedlegg funnet i ressursmappen)"
        Exit Function
    End If

    tekst = ""
    For Each fil In liste
        tekst = tekst & fil & vbCrLf
    Next fil
    VedleggTekst = Left(tekst, Len(tekst) - Len(vbCrLf))
End Function

' --- Utsatt levering ("Levering" i Oppsett) - under utprovning ----------

Public Function BeregnLeveringstidspunkt(ByRef feilmelding As String) As Variant
    ' Returnerer Empty (= send med en gang) eller en Dato a bruke for
    ' olMail.DeferredDeliveryTime, avhengig av hva som er valgt i
    ' Oppsett. Ved en tolkningsfeil settes feilmelding og Empty
    ' returneres - da sendes mailen med en gang i stedet for a feile
    ' helt, men den som ser sammendraget far vite hvorfor.
    Dim konfig As Object, leveringType As String
    Set konfig = LesKonfig()
    leveringType = ""
    If konfig.Exists("LEVERING_TYPE") Then leveringType = konfig("LEVERING_TYPE")

    feilmelding = ""

    Select Case leveringType
        Case "MINUTTER"
            Dim minutter As Long
            minutter = 0
            If konfig.Exists("LEVERING_MINUTTER") Then minutter = Val(konfig("LEVERING_MINUTTER"))
            If minutter > 0 Then
                BeregnLeveringstidspunkt = Now + TimeSerial(0, CLng(minutter), 0)
            End If

        Case "TIDSPUNKT"
            ' Dag/Maned/Ar lagres hver for seg (tre kaskaderende
            ' nedtrekksmenyer i Oppsett, se OppdaterDagListe der) i stedet
            ' for en enkelt kombinert datostreng.
            Dim dagStr As String, manedStr As String, aarStr As String, klokkeStr As String
            dagStr = "": manedStr = "": aarStr = "": klokkeStr = ""
            If konfig.Exists("LEVERING_DAG") Then dagStr = konfig("LEVERING_DAG")
            If konfig.Exists("LEVERING_MANED") Then manedStr = konfig("LEVERING_MANED")
            If konfig.Exists("LEVERING_AAR") Then aarStr = konfig("LEVERING_AAR")
            If konfig.Exists("LEVERING_KLOKKESLETT") Then klokkeStr = konfig("LEVERING_KLOKKESLETT")

            If dagStr <> "" And manedStr <> "" And aarStr <> "" Then
                Dim tidspunkt As Date
                On Error Resume Next
                tidspunkt = DateSerial(CLng(aarStr), CLng(manedStr), CLng(dagStr))
                If klokkeStr <> "" Then tidspunkt = tidspunkt + TimeValue(klokkeStr)
                If Err.Number <> 0 Then
                    feilmelding = "klarte ikke å tolke dato/klokkeslett (""" & dagStr & "." & manedStr & "." & aarStr & " " & klokkeStr & """)"
                ElseIf tidspunkt <= Now Then
                    feilmelding = "angitt leveringstidspunkt er allerede passert"
                Else
                    BeregnLeveringstidspunkt = tidspunkt
                End If
                On Error GoTo 0
            End If
    End Select
End Function

Public Function LeveringsBeskrivelse() As String
    ' Menneskelesbar tekst til bekreftelsesdialoger/forhandsvisninger.
    Dim feilmelding As String
    Dim tid As Variant
    tid = BeregnLeveringstidspunkt(feilmelding)

    If feilmelding <> "" Then
        LeveringsBeskrivelse = "Sendes med en gang (" & feilmelding & " - sender med en gang i stedet)"
    ElseIf IsEmpty(tid) Then
        LeveringsBeskrivelse = "Sendes med en gang"
    Else
        LeveringsBeskrivelse = "Utsettes til " & Format(tid, "dd.mm.yyyy \k\l. hh:mm")
    End If
End Function

Public Function AvsenderBeskrivelse() As String
    ' Menneskelesbar tekst til bekreftelsesdialoger.
    Dim konfig As Object, valgt As String
    Set konfig = LesKonfig()
    valgt = ""
    If konfig.Exists("AVSENDER_VALGT") Then valgt = konfig("AVSENDER_VALGT")
    If valgt = "" Then
        AvsenderBeskrivelse = "Automatisk (kontoen som kjører makroen)"
    Else
        AvsenderBeskrivelse = valgt
    End If
End Function

' --- Status-kolonne: retning (normal/omvendt) og tekst -----------------
' Lagt til 2026-09-23 etter Håkons ønske om å kunne reversere hvilke rader
' som vises: normalt vises TOMME rader (og sending SKRIVER teksten inn),
' omvendt vises rader MED den valgte teksten (og sending FJERNER teksten
' igjen). Selve teksten er fritt konfigurerbar i Oppsett i begge retninger
' (ikke lenger hardkodet "Ja") og sammenlignes alltid uten hensyn til
' store/små bokstaver (skrevet "X" i Oppsett matcher "x" i arket).

Public Function StatusOmvendtModus() As Boolean
    Dim konfig As Object
    Set konfig = LesKonfig()
    StatusOmvendtModus = konfig.Exists("STATUS_OMVENDT") And konfig("STATUS_OMVENDT") = "1"
End Function

Public Function StatusTekst() As String
    ' "Ja" brukes kun som et sikkerhetsnett for eldre oppsett som ikke er
    ' lagret på nytt siden dette feltet ble lagt til - selve Oppsett-feltet
    ' skal ALDRI vise "Ja" forhåndsutfylt (eksplisitt ønske fra Håkon).
    Dim konfig As Object
    Set konfig = LesKonfig()
    If konfig.Exists("STATUS_TEKST") Then
        If Trim(konfig("STATUS_TEKST")) <> "" Then
            StatusTekst = konfig("STATUS_TEKST")
            Exit Function
        End If
    End If
    StatusTekst = "Ja"
End Function

Public Function StatusRadKlasse(ByVal statusVerdi As String) As String
    ' Returnerer "AKTUELL" (gjenstår - skal vises/sendes), "SENDT" (ferdig
    ' behandlet) eller "FEIL" (uventet innhold - verken tomt eller den
    ' gjeldende teksten). Hvilken verdi som betyr hva avhenger av retningen.
    Dim v As String, tekst As String
    v = Trim(statusVerdi)
    tekst = StatusTekst()
    If StatusOmvendtModus() Then
        If LCase(v) = LCase(tekst) Then
            StatusRadKlasse = "AKTUELL"
        ElseIf v = "" Then
            StatusRadKlasse = "SENDT"
        Else
            StatusRadKlasse = "FEIL"
        End If
    Else
        If v = "" Then
            StatusRadKlasse = "AKTUELL"
        ElseIf LCase(v) = LCase(tekst) Then
            StatusRadKlasse = "SENDT"
        Else
            StatusRadKlasse = "FEIL"
        End If
    End If
End Function

Public Function StatusRadErAktuell(ByVal statusVerdi As String) As Boolean
    StatusRadErAktuell = (StatusRadKlasse(statusVerdi) = "AKTUELL")
End Function

Public Sub StatusMarkerSendt(ByVal celle As Range)
    ' Kalles etter en bekreftet sending - skriver teksten inn (normal
    ' retning) eller fjerner den igjen (omvendt retning).
    If StatusOmvendtModus() Then
        celle.Value = ""
    Else
        celle.Value = StatusTekst()
    End If
End Sub

Public Function StatusSendtBeskrivelse() As String
    ' Menneskelesbar tekst til "allerede sendt"-bekreftelsesdialogen.
    If StatusOmvendtModus() Then
        StatusSendtBeskrivelse = "tom - teksten er fjernet"
    Else
        StatusSendtBeskrivelse = """" & StatusTekst() & """"
    End If
End Function

Public Function StatusSendtLabel() As String
    ' Etiketten til den midterste tallboksen i frmSendMail (Gjenstår/?/
    ' Feil/Totalt). I normal retning betyr "SENDT"-klassen faktisk at
    ' verktøyet har sendt raden - "Sendt" er da en riktig etikett. I
    ' omvendt retning betyr "SENDT"-klassen bare "tom celle", som i
    ' praksis kan bety enten "allerede sendt og tekst fjernet" ELLER
    ' "aldri hadde teksten i utgangspunktet" - vi kan ikke skille disse
    ' fra hverandre ut fra statuskolonnen alene. Å kalle den "Sendt" ville
    ' derfor vært misvisende (Håkon, 2026-09-23: "denne er ikke helt
    ' riktig" i omvendt modus) - "Ikke aktuelt" er ærlig om usikkerheten.
    If StatusOmvendtModus() Then
        StatusSendtLabel = "Ikke aktuelt"
    Else
        StatusSendtLabel = "Sendt"
    End If
End Function

' --- Feilsjekk (frmFeilsjekk) - under utprovning -----------------------

Public Function FinnManglendeData() As Collection
    ' Finner to slags feil pa hver "aktuelle" rad (status-kolonnen tom, dvs.
    ' ikke allerede sendt/feilmarkert - samme "gjenstar"-kriterium som
    ' frmSendMail sin egen liste bruker):
    '  1) Raden mangler e-postadresse - da blir den stille hoppet over av
    '     den vanlige sendingen, uten at noen far vite at akkurat den
    '     kontakten aldri fikk mail. Flagges bare hvis raden faktisk har
    '     ANNEN data (ikke en helt blank/ubrukt tabellrad).
    '  2) Raden HAR e-post, men en plassholders mappede kolonne har en TOM
    '     celle for akkurat den raden - ville blitt sendt med et blankt
    '     hull, uten at "ukoblet plassholder"-sjekken i BygMailForRad (som
    '     bare fanger plassholdere som ALDRI ble koblet i Oppsett) fanger
    '     det opp.
    ' Hvert funn er en Scripting.Dictionary med RadNr/KolNr/Plassholder/
    ' KolonneNavn - nok til bade a vise en lesbar liste OG hoppe til
    ' akkurat den cellen i frmFeilsjekk.
    Dim resultat As New Collection
    Dim tbl As ListObject, konfig As Object
    Dim nokkel As Variant, plassholder As String, kolonnenavn As String, kolNr As Long
    Dim r As Long, epostKol As Long, statusKol As Long, verdi As String
    Dim funn As Object

    Set FinnManglendeData = resultat

    Set tbl = AktivTabell()
    If tbl Is Nothing Then Exit Function
    Set konfig = LesKonfig()
    If Not konfig.Exists("KOL_EPOST") Then Exit Function
    epostKol = KolonneNr(tbl, konfig("KOL_EPOST"))
    statusKol = 0
    If konfig.Exists("KOL_STATUS") Then statusKol = KolonneNr(tbl, konfig("KOL_STATUS"))

    For r = 1 To tbl.DataBodyRange.Rows.Count
        ' Rader som allerede er sendt eller feilmarkert er ikke "aktuelle"
        ' lenger og skal ikke nages om - se StatusRadErAktuell for hvordan
        ' dette avgjøres (avhenger av valgt retning i Oppsett).
        If statusKol > 0 Then
            If Not StatusRadErAktuell(CStr(tbl.DataBodyRange.Cells(r, statusKol).Value)) Then GoTo NesteRad
        End If

        If epostKol > 0 Then
            If Trim(CStr(tbl.DataBodyRange.Cells(r, epostKol).Value)) = "" Then
                If Application.WorksheetFunction.CountA(tbl.DataBodyRange.Rows(r)) > 0 Then
                    Set funn = CreateObject("Scripting.Dictionary")
                    funn("RadNr") = r
                    funn("KolNr") = epostKol
                    funn("Plassholder") = "e-postadresse"
                    funn("KolonneNavn") = konfig("KOL_EPOST")
                    resultat.Add funn
                End If
                GoTo NesteRad
            End If
        End If

        For Each nokkel In konfig.Keys
            If Left(nokkel, 3) = "PH:" Then
                plassholder = Mid(nokkel, 4)
                kolonnenavn = konfig(nokkel)
                kolNr = KolonneNr(tbl, kolonnenavn)
                If kolNr > 0 Then
                    verdi = Trim(CStr(tbl.DataBodyRange.Cells(r, kolNr).Value))
                    If verdi = "" Then
                        Set funn = CreateObject("Scripting.Dictionary")
                        funn("RadNr") = r
                        funn("KolNr") = kolNr
                        funn("Plassholder") = plassholder
                        funn("KolonneNavn") = kolonnenavn
                        resultat.Add funn
                    End If
                End If
            End If
        Next nokkel
NesteRad:
    Next r
End Function

' --- Selve utsendelsen -------------------------------------------------

Public Function BygMailForRad(ByVal radNr As Long, ByRef olMail As Object, ByRef feilmelding As String) As Boolean
    ' Felles "bygg mailen"-logikk for bade enkelt-sending (SendMailForRad,
    ' som viser feil i en MsgBox og kaller .Display) og masse-sending
    ' (frmMasseSend, som IKKE skal vise en dialogboks per rad - den samler
    ' feilmeldinger stille via feilmelding-parameteren og kaller .Send).
    ' Returnerer False + en feilmelding (ikke en MsgBox her) ved problemer.
    Dim tbl As ListObject
    Dim konfig As Object
    Dim colEpost As Long, colStatus As Long
    Dim epost As String
    Dim olApp As Object
    Dim ressursMappeFunnet As Boolean, malFilFunnet As Boolean
    Dim nokkel As Variant
    Dim plassholder As String, kolonnenavn As String, kolNr As Long, verdi As String

    BygMailForRad = False
    feilmelding = ""

    Set tbl = AktivTabell()
    If tbl Is Nothing Then
        feilmelding = "Fant ingen tabell i arbeidsboka."
        Exit Function
    End If

    Set konfig = LesKonfig()
    If Not konfig.Exists("KOL_EPOST") Or Not konfig.Exists("KOL_STATUS") Then
        feilmelding = "Oppsettet er ikke fullført (e-post-/statuskolonne mangler)."
        Exit Function
    End If

    colEpost = KolonneNr(tbl, konfig("KOL_EPOST"))
    colStatus = KolonneNr(tbl, konfig("KOL_STATUS"))
    If colEpost = 0 Or colStatus = 0 Then
        feilmelding = "Fant ikke kolonnene fra Oppsett i tabellen."
        Exit Function
    End If

    epost = Trim(CStr(tbl.DataBodyRange.Cells(radNr, colEpost).Value))
    If epost = "" Then
        feilmelding = "Mangler e-postadresse."
        Exit Function
    End If

    On Error Resume Next
    ressursMappeFunnet = (Dir(RessursMappe(), vbDirectory) <> "")
    If Err.Number <> 0 Then ressursMappeFunnet = False
    On Error GoTo 0
    If Not ressursMappeFunnet Then
        feilmelding = "Finner ikke ressursmappen: " & RessursMappe()
        Exit Function
    End If

    ' Dir("") oppfores upredikerbart i VBA (fortsetter et evt. tidligere
    ' Dir()-sok i stedet for a bare returnere tomt) - sjekk derfor
    ' eksplisitt for tom MalSti() FOR vi i det hele tatt kaller Dir() pa
    ' den, i stedet for a stole pa at Dir("") oppforer seg fornuftig.
    If MalSti() = "" Then
        feilmelding = "Ingen mal-fil er valgt ennå."
        Exit Function
    End If

    On Error Resume Next
    malFilFunnet = (Dir(MalSti()) <> "")
    If Err.Number <> 0 Then malFilFunnet = False
    On Error GoTo 0
    If Not malFilFunnet Then
        feilmelding = "Finner ikke mal-filen: " & MalSti()
        Exit Function
    End If

    Set olApp = HentOutlookApp()
    If olApp Is Nothing Then
        feilmelding = "Klarte ikke å starte/koble til Outlook."
        Exit Function
    End If

    Set olMail = olApp.CreateItemFromTemplate(MalSti())

    For Each nokkel In konfig.Keys
        If Left(nokkel, 3) = "PH:" Then
            plassholder = Mid(nokkel, 4)
            kolonnenavn = konfig(nokkel)
            kolNr = KolonneNr(tbl, kolonnenavn)
            If kolNr > 0 Then
                verdi = Trim(CStr(tbl.DataBodyRange.Cells(radNr, kolNr).Value))
                olMail.Subject = Replace(olMail.Subject, plassholder, verdi)
                olMail.HTMLBody = Replace(olMail.HTMLBody, plassholder, verdi)
            End If
        End If
    Next nokkel

    ' Sjekk om noen plassholdere star ukoblet igjen i selve mailen etter
    ' erstatningen - enten fordi de aldri ble koblet i Oppsett, eller
    ' fordi kolonnen de peker til er slettet/omdopt i ettertid. Uten
    ' denne sjekken kunne en rad "AAA" eller "BBB" blitt sendt rett til
    ' en mottaker uten at noen la merke til det - spesielt kritisk ved
    ' masse-sending der ingen ser gjennom hver enkelt mail forst.
    Dim gjenstaende As Collection, p As Variant, liste As String
    Set gjenstaende = FinnPlassholdere(olMail.Subject & " " & olMail.HTMLBody)
    If gjenstaende.Count > 0 Then
        liste = ""
        For Each p In gjenstaende
            liste = liste & p & ", "
        Next p
        If Len(liste) > 2 Then liste = Left(liste, Len(liste) - 2)
        feilmelding = "Ukoblet(e) plassholder(e) står fortsatt igjen: " & liste
        Exit Function
    End If

    ' CC-kolonnen er valgfri - kan mangle helt i konfig (eldre oppsett som
    ' ikke er lagret pa nytt ennaa), stå tom (bevisst "ingen CC"), eller
    ' peke pa en kolonne der akkurat DENNE raden har en tom celle. Alle tre
    ' skal bare bety "ingen CC pa denne mailen", aldri en feil.
    Dim colCC As Long, ccVerdi As String
    colCC = 0
    ccVerdi = ""
    If konfig.Exists("KOL_CC") Then
        If Len(konfig("KOL_CC")) > 0 Then
            colCC = KolonneNr(tbl, konfig("KOL_CC"))
            If colCC > 0 Then
                ccVerdi = Trim(CStr(tbl.DataBodyRange.Cells(radNr, colCC).Value))
            End If
        End If
    End If

    olMail.To = epost
    If Len(ccVerdi) > 0 Then olMail.CC = ccVerdi
    Dim avsenderValgt As String
    avsenderValgt = ""
    If konfig.Exists("AVSENDER_VALGT") Then avsenderValgt = konfig("AVSENDER_VALGT")
    If Len(avsenderValgt) > 0 Then SettAvsender olApp, olMail, avsenderValgt

    LeggVedVedlegg olMail

    Dim leveringFeil As String, leveringstid As Variant
    leveringstid = BeregnLeveringstidspunkt(leveringFeil)
    If Not IsEmpty(leveringstid) Then
        olMail.DeferredDeliveryTime = leveringstid
    End If

    BygMailForRad = True
End Function

Function SendMailForRad(ByVal radNr As Long) As Boolean
    Dim tbl As ListObject, konfig As Object, colStatus As Long, status As String
    Dim olMail As Object, feilmelding As String
    Dim brodtekstUtenSignatur As String
    Dim svar As VbMsgBoxResult

    SendMailForRad = False

    ' "Allerede sendt"-bekreftelsen horer hjemme her (enkelt-sending), ikke
    ' i BygMailForRad - masse-sending viser uansett bare rader som IKKE
    ' er sendt fra for, sa den situasjonen oppstar ikke der.
    Set tbl = AktivTabell()
    Set konfig = LesKonfig()
    If Not tbl Is Nothing And konfig.Exists("KOL_STATUS") Then
        colStatus = KolonneNr(tbl, konfig("KOL_STATUS"))
        If colStatus > 0 Then
            status = Trim(CStr(tbl.DataBodyRange.Cells(radNr, colStatus).Value))
            If StatusRadKlasse(status) = "SENDT" Then
                svar = MsgBox("Denne raden er allerede markert som sendt (" & StatusSendtBeskrivelse() & ")." & vbCrLf & _
                               "Vil du likevel åpne en ny mail for denne raden?", vbYesNo + vbQuestion, "Mail-utsendelse")
                If svar = vbNo Then Exit Function
            End If
        End If
    End If

    If Not BygMailForRad(radNr, olMail, feilmelding) Then
        MsgBox feilmelding, vbCritical, "Mail-utsendelse"
        Exit Function
    End If

    ' Ta vare pa leveringstidspunktet BygMailForRad eventuelt har satt, for
    ' Display() - se samme monster rett under for brodteksten. Bekreftet
    ' (Håkon, 2026-09-08): Display() nullstiller ogsa DeferredDeliveryTime,
    ' ikke bare signaturen - trolig fordi Display() initialiserer om noen
    ' av kompose-vinduets standardverdier. Masse-sending (frmMasseSend)
    ' rammes ikke, siden den aldri kaller Display(), bare Send direkte.
    Dim harLevering As Boolean, leveringSattFor As Date
    Dim feilmeldingLevering As String
    harLevering = Not IsEmpty(modGenerisk.BeregnLeveringstidspunkt(feilmeldingLevering))
    If harLevering Then leveringSattFor = olMail.DeferredDeliveryTime

    ' Se LokalStiFraSky-kommentaren for bakgrunn pa monsteret under -
    ' Outlook kan lime inn brukerens personlige signatur ved Display(),
    ' selv for en mail laget fra mal. Ta vare pa/gjenopprett brodteksten.
    brodtekstUtenSignatur = olMail.HTMLBody
    olMail.Display
    If olMail.HTMLBody <> brodtekstUtenSignatur Then
        olMail.HTMLBody = brodtekstUtenSignatur
    End If

    If harLevering Then
        If olMail.DeferredDeliveryTime <> leveringSattFor Then
            olMail.DeferredDeliveryTime = leveringSattFor
        End If
        ' Lagt til 2026-09-10 etter en observert forvirring - en utsatt
        ' enkelt-mail forsvant sporløst fordi kompose-vinduet krever et
        ' manuelt Send-trykk (Display() sender ALDRI automatisk), og det
        ' trykket ble glemt i en rask testrunde.
        MsgBox "Levering er satt til " & Format(leveringSattFor, "dd.mm.yyyy \k\l. hh:mm") & "." & vbCrLf & _
               "Husk å trykke Send i vinduet som er åpnet - mailen sendes IKKE automatisk før du gjør det.", vbInformation, "Utsatt levering"
    End If

    SendMailForRad = True
End Function
'@

# ---------------------------------------------------------------
# VBA-kildekode: frmSendMail (utvalgsvindu)
# ---------------------------------------------------------------
$formCodeSendMail = @'
Option Explicit

Private Sub UserForm_Initialize()
    ' Fet skrift/farger settes her i VBA (kjorer trygt i Excel selv) -
    ' IKKE fra PowerShell-COM ved bygging, som krasjer pa .Font.
    lblTittel.Font.Bold = True
    lblStatGjenstar.Font.Bold = True
    lblStatSendt.Font.Bold = True
    lblStatFeil.Font.Bold = True
    lblStatSendt.ForeColor = RGB(40, 140, 70)
    lblStatFeil.ForeColor = RGB(190, 50, 50)
    lblStatTotalt.ForeColor = RGB(110, 110, 110)
    lblLevering.Font.Bold = True
    lblLevering.ForeColor = RGB(190, 120, 20)
    cmdSendMail.Font.Bold = True
    cmdSendTilAlle.Font.Bold = True

    lblVedlegg.Caption = "Vedlegg som legges ved: " & Replace(modGenerisk.VedleggTekst(), vbCrLf, ", ")

    Dim leveringTekst As String
    leveringTekst = modGenerisk.LeveringsBeskrivelse()
    If leveringTekst <> "Sendes med en gang" Then
        lblLevering.Caption = "Utsatt levering er på (Oppsett): " & leveringTekst
        lblLevering.Visible = True
    Else
        lblLevering.Visible = False
    End If

    On Error Resume Next
    TilpassVindusstorrelse
    On Error GoTo 0
    FyllListe
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    lstRader.SetFocus
    On Error GoTo 0
End Sub

Private Sub TilpassVindusstorrelse()
    ' Krymper vinduet til skjermen om nodvendig - OG krymper listen med,
    ' sa den ikke henger igjen utenfor det synlige vinduet.
    Dim maksHoyde As Single
    Dim differanse As Single

    maksHoyde = Application.UsableHeight - 30
    If Me.Height > maksHoyde Then
        differanse = Me.Height - maksHoyde
        Me.Height = maksHoyde
        lstRader.Height = lstRader.Height - differanse
        lblVedlegg.Top = lblVedlegg.Top - differanse
        lblLevering.Top = lblLevering.Top - differanse
        cmdOppsett.Top = cmdOppsett.Top - differanse
        cmdSendMail.Top = cmdSendMail.Top - differanse
        cmdFeilsjekk.Top = cmdFeilsjekk.Top - differanse
        cmdSendTilAlle.Top = cmdSendTilAlle.Top - differanse
    End If
End Sub

Private Sub FyllListe()
    Dim tbl As ListObject
    Dim konfig As Object
    Dim colStatus As Long
    Dim antKolonner As Long, antRader As Long, r As Long, c As Long
    Dim data() As Variant
    Dim radIndekser() As Long
    Dim n As Long, antSendt As Long, antFeil As Long
    Dim bredder As String
    Dim status As String
    Dim i As Long

    lblStatGjenstar.Visible = False
    lblStatSendt.Visible = False
    lblStatFeil.Visible = False
    lblStatTotalt.Visible = False
    lblInfo.Visible = True

    Set tbl = modGenerisk.AktivTabell()
    If tbl Is Nothing Then
        lblInfo.Caption = "Fant ingen tabell i arbeidsboka."
        cmdSendMail.Enabled = False
        Exit Sub
    End If

    Set konfig = modGenerisk.LesKonfig()
    If Not konfig.Exists("KOL_STATUS") Then
        lblInfo.Caption = "Oppsettet er ikke fullført - trykk ""Oppsett""."
        cmdSendMail.Enabled = False
        Exit Sub
    End If

    colStatus = modGenerisk.KolonneNr(tbl, konfig("KOL_STATUS"))
    If colStatus = 0 Then
        lblInfo.Caption = "Statuskolonnen fra Oppsett finnes ikke lenger i tabellen."
        cmdSendMail.Enabled = False
        Exit Sub
    End If

    If tbl.DataBodyRange Is Nothing Then
        lblInfo.Caption = "Tabellen har ingen datarader ennå."
        cmdSendMail.Enabled = False
        Exit Sub
    End If

    Dim synligeKol As Collection, kolIdx As Variant
    Dim antTotaltRader As Long
    Set synligeKol = modGenerisk.SynligeKolonneIndekser(tbl)
    antKolonner = synligeKol.Count
    antRader = tbl.DataBodyRange.Rows.Count

    ReDim data(1 To antRader, 1 To antKolonner)
    ReDim radIndekser(1 To antRader)
    n = 0
    antSendt = 0
    antFeil = 0
    antTotaltRader = 0

    ' Status-kolonnen kan ha tre slags verdier: aktuell (gjenstar - vises i
    ' lista), sendt (skjules), eller noe helt annet (behandles som en
    ' feilmarkering - skjules ogsa, men telles for seg, slik TIMSS-verktoyet
    ' gjorde med "ERROR"-rader). Hvilken verdi som betyr hva avhenger av
    ' retningen valgt i Oppsett - se StatusRadKlasse. Skjulte/filtrerte
    ' rader (f.eks. via autofilter) telles heller ikke med noe sted - de er
    ' jo ikke synlige i selve arket akkurat na, sa "Totalt" telles ogsa kun
    ' blant synlige rader for a stemme med de tre andre tallene.
    For r = 1 To antRader
        If Not tbl.DataBodyRange.Rows(r).EntireRow.Hidden Then
            antTotaltRader = antTotaltRader + 1
            status = Trim(CStr(tbl.DataBodyRange.Cells(r, colStatus).Value))
            Select Case modGenerisk.StatusRadKlasse(status)
                Case "AKTUELL"
                    n = n + 1
                    c = 0
                    For Each kolIdx In synligeKol
                        c = c + 1
                        data(n, c) = CStr(tbl.DataBodyRange.Cells(r, CLng(kolIdx)).Value)
                    Next kolIdx
                    radIndekser(n) = r
                Case "SENDT"
                    antSendt = antSendt + 1
                Case Else
                    antFeil = antFeil + 1
            End Select
        End If
    Next r

    Me.Tag = ""
    For i = 1 To n
        Me.Tag = Me.Tag & radIndekser(i) & ","
    Next i

    lstRader.ColumnCount = antKolonner
    bredder = ""
    For c = 1 To antKolonner
        bredder = bredder & "90 pt;"
    Next c
    lstRader.ColumnWidths = bredder

    If n = 0 Then
        lstRader.Clear
    Else
        Dim visning() As Variant
        ReDim visning(1 To n, 1 To antKolonner)
        For r = 1 To n
            For c = 1 To antKolonner
                visning(r, c) = data(r, c)
            Next c
        Next r
        lstRader.List = visning
    End If

    lblInfo.Visible = False
    lblStatGjenstar.Caption = "Gjenstår: " & n
    lblStatSendt.Caption = modGenerisk.StatusSendtLabel() & ": " & antSendt
    lblStatFeil.Caption = "Feil: " & antFeil
    lblStatTotalt.Caption = "Totalt: " & antTotaltRader
    lblStatGjenstar.Visible = True
    lblStatSendt.Visible = True
    lblStatFeil.Visible = True
    lblStatTotalt.Visible = True
    cmdSendMail.Enabled = (n > 0)

    Dim antallManglendeData As Long
    On Error Resume Next
    antallManglendeData = modGenerisk.FinnManglendeData().Count
    On Error GoTo 0
    If antallManglendeData > 0 Then
        cmdFeilsjekk.Caption = "Undersøk feil (" & antallManglendeData & ")"
        cmdFeilsjekk.Font.Bold = True
        cmdFeilsjekk.ForeColor = RGB(190, 50, 50)
    Else
        cmdFeilsjekk.Caption = "Undersøk feil"
        cmdFeilsjekk.Font.Bold = False
        cmdFeilsjekk.ForeColor = RGB(0, 0, 0)
    End If
    ' Bredden var fast (160pt), designet for den lengste varianten av
    ' teksten - ga et unaturlig stort/tomt utseende for korte varianter og
    ' feil avstand til knappen etter (Håkon, 2026-09-11: "plassert rart").
    ' Beregn bredden fra selve teksten i stedet, og flytt knappen etter med.
    cmdFeilsjekk.Width = Len(cmdFeilsjekk.Caption) * 7 + 24
    If cmdFeilsjekk.Width < 110 Then cmdFeilsjekk.Width = 110
    cmdSendTilAlle.Left = cmdFeilsjekk.Left + cmdFeilsjekk.Width + 10
End Sub

Private Sub cmdSendMail_Click()
    If lstRader.ListIndex = -1 Then
        MsgBox "Velg en rad først.", vbExclamation, "Mail-utsendelse"
        Exit Sub
    End If

    Dim radIndekser() As String
    radIndekser = Split(Me.Tag, ",")
    Dim ekteRad As Long
    ekteRad = CLng(radIndekser(lstRader.ListIndex))

    If modGenerisk.SendMailForRad(ekteRad) Then
        Dim svar As VbMsgBoxResult
        svar = MsgBox("Ble mailen sendt?", vbYesNo + vbQuestion, "Mail-utsendelse")
        If svar = vbYes Then
            Dim tbl As ListObject, konfig As Object, colStatus As Long
            Set tbl = modGenerisk.AktivTabell()
            Set konfig = modGenerisk.LesKonfig()
            colStatus = modGenerisk.KolonneNr(tbl, konfig("KOL_STATUS"))
            modGenerisk.StatusMarkerSendt tbl.DataBodyRange.Cells(ekteRad, colStatus)
            FyllListe
        End If
    End If
End Sub

Private Sub cmdOppsett_Click()
    modGenerisk.OppsettReturnerTil = "frmSendMail"
    Unload Me
    frmOppsett.Show
End Sub

Private Sub cmdSendTilAlle_Click()
    Unload Me
    frmMasseSend.Show
End Sub

Private Sub cmdFeilsjekk_Click()
    ' frmSendMail (modalt skjema) ma lukkes forst - VBA tillater ikke a
    ' vise et modelost skjema mens et modalt skjema fortsatt er synlig
    ' (feilkode 401). Excel forblir synlig/interaktivt mens frmFeilsjekk
    ' star apen etterpa, som var hele poenget med a gjore den modelos.
    modGenerisk.FeilsjekkReturnerTil = "frmSendMail"
    Unload Me
    frmFeilsjekk.Show vbModeless
End Sub
'@

# ---------------------------------------------------------------
# VBA-kildekode: frmOppsett (koblingsvindu for plassholdere/kolonner)
# ---------------------------------------------------------------
$formCodeOppsett = @'
Option Explicit

Private Const PH_RAD_HOYDE As Long = 22
Private Const PH_TOPP_PADDING As Long = 16
Private Const PH_BUNN_PADDING As Long = 10
Private Const PH_MIN_HOYDE As Long = 30
Private Const LEVER_FARGE_AV As Long = 15921906  ' samme grå som andre "lockede" felt i Oppsett
Private Const LEVER_FARGE_PA As Long = 16777215  ' hvit

Private Sub UserForm_Initialize()
    Dim tbl As ListObject
    Dim konfig As Object
    Dim i As Long

    ' Fet skrift settes her i VBA (kjorer trygt i Excel selv) - IKKE
    ' fra PowerShell-COM ved bygging, som krasjer pa .Font.
    lblTittel.Font.Bold = True
    fraMalfil.Font.Bold = True
    fraPlassholdere.Font.Bold = True
    fraEpostStatus.Font.Bold = True
    fraAvsender.Font.Bold = True
    fraLevering.Font.Bold = True
    cmdLagre.Font.Bold = True

    lblMalFil.Caption = "(ingen valgt)"
    Set konfig = modGenerisk.LesKonfig()
    If konfig.Exists("MAL_UNDERMAPPE") Or konfig.Exists("MAL_FILNAVN") Then
        Dim undermappe As String, filnavn As String
        undermappe = ""
        filnavn = ""
        If konfig.Exists("MAL_UNDERMAPPE") Then undermappe = konfig("MAL_UNDERMAPPE")
        If konfig.Exists("MAL_FILNAVN") Then filnavn = konfig("MAL_FILNAVN")
        If filnavn <> "" Then
            lblMalFil.Caption = IIf(undermappe = "", filnavn, undermappe & "\" & filnavn)
        End If
    End If
    OppdaterMalfilLayout

    For i = 1 To modGenerisk.MAKS_PLASSHOLDERE
        Me.Controls("lblPH" & i).Visible = False
        Me.Controls("cboKol" & i).Visible = False
    Next i
    OppdaterLayout 0

    Set tbl = modGenerisk.AktivTabell()
    If tbl Is Nothing Then
        lblMelding.Caption = "OBS: fant ingen tabell i arbeidsboka. Marker dataene som en tabell (Sett inn > Tabell) først."
        cmdLagre.Enabled = False
        Exit Sub
    End If

    FyllKolonnelister tbl
    ' cboEpost/cboStatus/cboCC er fmStyleDropDownList (Style=2, kun listede
    ' verdier er gyldige) - en tidligere lagret kolonnenavn som ikke lenger
    ' finnes i tabellen (tabellstrukturen endret seg) gir ellers en
    ' ufanget "Could not set the Value property"-feil her og krasjer
    ' Oppsett. Samme fallgruve/losning som cboLeverManed/cboLeverAar/
    ' cboLeverKlokkeslett lenger ned i denne samme UserForm_Initialize -
    ' hopp stille over i stedet for a feile pa en ugyldig tilordning.
    On Error Resume Next
    If konfig.Exists("KOL_EPOST") Then cboEpost.Value = konfig("KOL_EPOST")
    If konfig.Exists("KOL_STATUS") Then cboStatus.Value = konfig("KOL_STATUS")
    If konfig.Exists("KOL_CC") Then cboCC.Value = konfig("KOL_CC")
    On Error GoTo 0

    ' Retning/tekst for status-kolonnen - txtStatusTekst skal ALDRI vise
    ' "Ja" forhåndsutfylt (kun det som faktisk er lagret fra før).
    Dim statusOmvendt As Boolean
    statusOmvendt = konfig.Exists("STATUS_OMVENDT") And konfig("STATUS_OMVENDT") = "1"
    optRetningOmvendt.Value = statusOmvendt
    optRetningNormal.Value = Not statusOmvendt
    If konfig.Exists("STATUS_TEKST") Then
        txtStatusTekst.Value = konfig("STATUS_TEKST")
    Else
        txtStatusTekst.Value = ""
    End If
    OppdaterStatusTekstLabel

    FyllAvsenderListe
    If konfig.Exists("AVSENDER_VALGT") And konfig("AVSENDER_VALGT") <> "" Then
        cboAvsender.Value = konfig("AVSENDER_VALGT")
    Else
        cboAvsender.Value = "Automatisk"
    End If

    FyllLeveringslister

    Dim leveringType As String
    leveringType = ""
    If konfig.Exists("LEVERING_TYPE") Then leveringType = konfig("LEVERING_TYPE")
    Select Case leveringType
        Case "MINUTTER"
            optLeverMinutter.Value = True
        Case "TIDSPUNKT"
            optLeverTidspunkt.Value = True
        Case Else
            optLeverEngang.Value = True
    End Select
    If konfig.Exists("LEVERING_MINUTTER") Then txtLeverMinutter.Value = konfig("LEVERING_MINUTTER")

    ' cboLeverManed/cboLeverAar/cboLeverKlokkeslett er fmStyleDropDownList
    ' (kun de listede verdiene er gyldige) - en tidligere lagret verdi som
    ' na faller utenfor det viste utvalget (f.eks. et arstall langt fram i
    ' tid) finnes rett og slett ikke i lista lenger, sa vi hopper stille
    ' over i stedet for a feile pa en ugyldig tilordning.
    On Error Resume Next
    If konfig.Exists("LEVERING_MANED") Then cboLeverManed.ListIndex = CLng(konfig("LEVERING_MANED")) - 1
    If konfig.Exists("LEVERING_AAR") Then cboLeverAar.Value = konfig("LEVERING_AAR")
    On Error GoTo 0

    ' Standardvalg (maned/ar) dersom ingenting er lagret ennaa - inneværende
    ' maned/ar, siden nesten all bruk gjelder naer fremtid. Feltene er uansett
    ' gratonet/deaktivert med mindre "Bestemt tidspunkt" faktisk er valgt.
    If cboLeverManed.ListIndex = -1 Then cboLeverManed.ListIndex = Month(Date) - 1
    If cboLeverAar.Value = "" Then cboLeverAar.Value = CStr(Year(Date))

    OppdaterDagListe
    On Error Resume Next
    If konfig.Exists("LEVERING_DAG") Then cboLeverDag.Value = konfig("LEVERING_DAG")
    If konfig.Exists("LEVERING_KLOKKESLETT") Then cboLeverKlokkeslett.Value = konfig("LEVERING_KLOKKESLETT")
    On Error GoTo 0

    OppdaterLeveringsFelter

    If modGenerisk.MalSti() <> "" Then
        LastPlassholdereFraFil modGenerisk.MalSti()
    End If
End Sub

Private Sub optLeverEngang_Click()
    OppdaterLeveringsFelter
End Sub

Private Sub optLeverMinutter_Click()
    OppdaterLeveringsFelter
End Sub

Private Sub optLeverTidspunkt_Click()
    OppdaterLeveringsFelter
End Sub

Private Sub optRetningNormal_Click()
    OppdaterStatusTekstLabel
End Sub

Private Sub optRetningOmvendt_Click()
    OppdaterStatusTekstLabel
End Sub

Private Sub OppdaterStatusTekstLabel()
    If optRetningOmvendt.Value Then
        lblStatusTekstLabel.Caption = "Tekst som regnes som ""fylt"":"
    Else
        lblStatusTekstLabel.Caption = "Tekst som skrives ved sending:"
    End If
End Sub

Private Sub cmdVelgKolonner_Click()
    ' Modal-over-modal er trygt i VBA (i motsetning til modelos-over-
    ' modal, som cmdFeilsjekk_Click andre steder ma ta hensyn til) - ingen
    ' Unload Me/navigasjon nodvendig, frmKolonnevisning lukker bare seg
    ' selv og gir kontrollen tilbake hit.
    frmKolonnevisning.Show
End Sub

Private Sub cboLeverManed_Change()
    OppdaterDagListe
End Sub

Private Sub cboLeverAar_Change()
    OppdaterDagListe
End Sub

Private Sub FyllLeveringslister()
    ' Nedtrekksmenyer i stedet for frie tekstfelt - umulig a skrive inn et
    ' ugyldig format eller en ikke-eksisterende dato da (tilbakemelding fra
    ' Håkon). Dag/Maned/Ar er tre sma, kaskaderende menyer i stedet for en
    ' enkelt lang liste med fulle datoer (som viste seg uoversiktlig å
    ' bla i) - Dag-lista bygges pa nytt i OppdaterDagListe hver gang Maned
    ' eller Ar endres, sa den alltid stemmer (skuddar/manedslengde).
    Dim manedNavn As Variant, i As Long, h As Long, m As Long

    cboLeverManed.Clear
    manedNavn = Array("Januar", "Februar", "Mars", "April", "Mai", "Juni", _
                       "Juli", "August", "September", "Oktober", "November", "Desember")
    For i = 0 To 11
        cboLeverManed.AddItem manedNavn(i)
    Next i

    cboLeverAar.Clear
    For i = 0 To 2
        cboLeverAar.AddItem CStr(Year(Date) + i)
    Next i

    cboLeverKlokkeslett.Clear
    For h = 0 To 23
        For m = 0 To 30 Step 30
            cboLeverKlokkeslett.AddItem Format(TimeSerial(h, m, 0), "hh:mm")
        Next m
    Next h
End Sub

Private Function AntallDagerIManed(ByVal maned As Long, ByVal aar As Long) As Long
    ' DateSerial(aar, maned+1, 0) = siste dag i "maned" - handterer
    ' skuddar (februar) og manedslengde (30/31 dager) helt automatisk.
    AntallDagerIManed = Day(DateSerial(aar, maned + 1, 0))
End Function

Private Sub OppdaterDagListe()
    If cboLeverManed.ListIndex = -1 Or cboLeverAar.Value = "" Then Exit Sub

    Dim maned As Long, aar As Long, antallDager As Long, d As Long
    Dim dagValgtForst As String

    dagValgtForst = cboLeverDag.Value
    maned = cboLeverManed.ListIndex + 1
    aar = CLng(cboLeverAar.Value)
    antallDager = AntallDagerIManed(maned, aar)

    cboLeverDag.Clear
    For d = 1 To antallDager
        cboLeverDag.AddItem Format(d, "00")
    Next d

    ' Behold tidligere valgt dag hvis den fortsatt er gyldig (f.eks. bytte
    ' fra en 31-dagers til en 30-dagers maned uten a miste valget unodig).
    ' NB: VBA "And" er ikke kortslutta (begge sider evalueres alltid), sa
    ' CLng(dagValgtForst) MA sta i en egen, nestet If - ellers feiler
    ' CLng("") med "Type mismatch" hver gang ingen dag er valgt fra for
    ' (f.eks. helt forste gang skjemaet apnes).
    If dagValgtForst <> "" Then
        If CLng(dagValgtForst) <= antallDager Then
            cboLeverDag.Value = dagValgtForst
        End If
    End If
End Sub

Private Sub OppdaterLeveringsFelter()
    ' Bade Enabled OG BackColor - Enabled alene er ikke alltid tydelig nok
    ' a se ved et blikk at feltet ikke er i bruk (tilbakemelding fra Håkon:
    ' "husk å gå ut utsatt levering om send med en gang er valgt").
    txtLeverMinutter.Enabled = optLeverMinutter.Value
    txtLeverMinutter.BackColor = IIf(optLeverMinutter.Value, LEVER_FARGE_PA, LEVER_FARGE_AV)

    cboLeverDag.Enabled = optLeverTidspunkt.Value
    cboLeverManed.Enabled = optLeverTidspunkt.Value
    cboLeverAar.Enabled = optLeverTidspunkt.Value
    cboLeverKlokkeslett.Enabled = optLeverTidspunkt.Value
    cboLeverDag.BackColor = IIf(optLeverTidspunkt.Value, LEVER_FARGE_PA, LEVER_FARGE_AV)
    cboLeverManed.BackColor = IIf(optLeverTidspunkt.Value, LEVER_FARGE_PA, LEVER_FARGE_AV)
    cboLeverAar.BackColor = IIf(optLeverTidspunkt.Value, LEVER_FARGE_PA, LEVER_FARGE_AV)
    cboLeverKlokkeslett.BackColor = IIf(optLeverTidspunkt.Value, LEVER_FARGE_PA, LEVER_FARGE_AV)
End Sub

Private Sub FyllAvsenderListe()
    Dim lagrede As Collection, e As Variant
    cboAvsender.Clear
    cboAvsender.AddItem "Automatisk"
    Set lagrede = modGenerisk.HentLagredeAvsendere()
    For Each e In lagrede
        cboAvsender.AddItem e
    Next e
End Sub

Private Sub cmdLeggTilAvsender_Click()
    Dim ny As String
    ny = Trim(txtNyAvsender.Value)
    If ny = "" Then Exit Sub
    If InStr(ny, "@") = 0 Then
        MsgBox "Det ser ikke ut som en gyldig e-postadresse.", vbExclamation, "Oppsett"
        Exit Sub
    End If
    modGenerisk.SkrivKonfigVerdi "AVSENDER:" & ny, ny
    FyllAvsenderListe
    cboAvsender.Value = ny
    txtNyAvsender.Value = ""
End Sub

Private Sub OppdaterMalfilLayout()
    ' Plasser "Velg"-knappen rett etter filnavn-teksten (som er AutoSize
    ' og dermed skifter bredde med innholdet), og la emne-forhandsvisningen
    ' strekke seg helt bort til samme hoyre kant som brodtekst-boksen under.
    cmdVelgMalFil.Left = lblMalFil.Left + lblMalFil.Width + 10
    txtEmne.Left = lblEmneLabel.Left + lblEmneLabel.Width + 8
    txtEmne.Width = (txtBrodtekst.Left + txtBrodtekst.Width) - txtEmne.Left

    ' Brodtekst-boksen skal ogsa bare vaere sa hoy som innholdet faktisk
    ' krever, ikke alltid en fast, ofte for stor standardstorrelse - egen
    ' rullefelt tar seg av uvanlig lange maler. Talte oppr. bare eksplisitte
    ' linjeskift (vbLf) - det undervurderte hoyden for LANGE enkeltlinjer
    ' som Excel selv bryter over flere synlige linjer (ordbryting), og
    ' klippet sa vidt siste linje pa enkelte skjermer/fonter (rapportert av
    ' Håkon 2026-09-10, "bare så vidt", "under dynamiske rammer"). Anslar na
    ' ogsa ordbryting ut fra tegn-per-linje, pluss en litt rausere hoyde-
    ' konstant enn for - bedre med litt for mye luft enn a klippe igjen.
    Dim brodtekstTekst As String, brodtekstLinjer As Long, brodtekstHoyde As Long
    Dim brLinjeArr() As String, brI As Long, brRadLengde As Long
    Const BR_TEGN_PER_LINJE As Long = 70
    brodtekstTekst = Replace(Replace(txtBrodtekst.Value, vbCrLf, vbLf), vbCr, vbLf)
    brLinjeArr = Split(brodtekstTekst, vbLf)
    brodtekstLinjer = 0
    For brI = LBound(brLinjeArr) To UBound(brLinjeArr)
        brRadLengde = Len(brLinjeArr(brI))
        If brRadLengde = 0 Then
            brodtekstLinjer = brodtekstLinjer + 1
        Else
            brodtekstLinjer = brodtekstLinjer + Int((brRadLengde - 1) \ BR_TEGN_PER_LINJE) + 1
        End If
    Next brI
    If brodtekstLinjer < 1 Then brodtekstLinjer = 1
    brodtekstHoyde = brodtekstLinjer * 15 + 14
    If brodtekstHoyde < 40 Then brodtekstHoyde = 40
    If brodtekstHoyde > 220 Then brodtekstHoyde = 220 ' svaert lange maler far egen rullefelt i stedet
    txtBrodtekst.Height = brodtekstHoyde

    ' Vedlegg-boksen skal bare vaere sa hoy som det faktiske antallet
    ' filer krever (ikke alltid reservere plass til flere linjer enn
    ' det er innhold) - den frigjorte plassen gar til brodtekst-
    ' forhandsvisningen i stedet, som er viktigere a se hele av.
    Dim vedleggListe As Collection
    Set vedleggListe = modGenerisk.HentVedleggsliste()
    txtVedlegg.Value = modGenerisk.VedleggTekst()

    Dim antallLinjer As Long
    antallLinjer = vedleggListe.Count
    If antallLinjer < 1 Then antallLinjer = 1 ' plass til "(ingen vedlegg ...)"-meldingen

    Dim vedleggHoyde As Long
    vedleggHoyde = antallLinjer * 15 + 12
    If vedleggHoyde < 24 Then vedleggHoyde = 24
    If vedleggHoyde > 60 Then vedleggHoyde = 60 ' mer enn ca. 4 linjer - fa egen rullefelt i stedet
    txtVedlegg.Height = vedleggHoyde

    lblVedleggLabel.Top = txtBrodtekst.Top + txtBrodtekst.Height + 8
    txtVedlegg.Top = lblVedleggLabel.Top + 18
    fraMalfil.Height = txtVedlegg.Top + txtVedlegg.Height + 12

    ' fraPlassholdere sin posisjon avledes fra fraMalfil sin faktiske
    ' hoyde, ikke en fast verdi - unngar at de to kommer i utakt om
    ' fraMalfil sin hoyde noensinne endres igjen.
    fraPlassholdere.Top = fraMalfil.Top + fraMalfil.Height + 10
End Sub

Private Sub OppdaterLayout(ByVal antallSynligePlassholdere As Long)
    ' Plassholder-rammen skal kun vaere sa hoy som det faktiske antallet
    ' plassholdere krever - resten (e-post/status-rammen, meldingen og
    ' lagre-knappen) flyttes opp/ned deretter, sa det ikke blir staende
    ' igjen tomrom nar det er fa plassholdere.
    Dim beregnetHoyde As Long
    beregnetHoyde = PH_TOPP_PADDING + antallSynligePlassholdere * PH_RAD_HOYDE + PH_BUNN_PADDING
    fraPlassholdere.Height = IIf(beregnetHoyde > PH_MIN_HOYDE, beregnetHoyde, PH_MIN_HOYDE)

    fraEpostStatus.Top = fraPlassholdere.Top + fraPlassholdere.Height + 10
    fraListevisning.Top = fraEpostStatus.Top + fraEpostStatus.Height + 10
    fraAvsender.Top = fraListevisning.Top + fraListevisning.Height + 10
    fraLevering.Top = fraAvsender.Top + fraAvsender.Height + 10
    lblMelding.Top = fraLevering.Top + fraLevering.Height + 8

    ' Meldingsfeltet skal ikke reservere tomrom nar det ikke har noe a
    ' vise (f.eks. det vanlige "fant N plassholdere"-tilfellet, som ikke
    ' lenger vises siden radene under allerede sier det samme) - da
    ' flyttes Lagre-knappen rett opp i tomrommet i stedet.
    Dim meldingHoyde As Long
    If lblMelding.Caption = "" Then
        meldingHoyde = 0
    Else
        meldingHoyde = lblMelding.Height + 8
    End If
    cmdLagre.Top = lblMelding.Top + meldingHoyde
    cmdTilbake.Top = cmdLagre.Top

    ' Vinduet vokser vanligvis for a passe innholdet, men hvis det
    ' likevel blir hoyere enn skjermen (mange plassholdere pa en liten
    ' skjerm), krymp vinduet og gjor resten tilgjengelig via rullefelt
    ' i stedet for at det forsvinner utenfor skjermen.
    Dim onsketHoyde As Long, maksHoyde As Long
    onsketHoyde = cmdLagre.Top + cmdLagre.Height + 40
    maksHoyde = Application.UsableHeight - 30

    Me.ScrollHeight = onsketHoyde
    If onsketHoyde > maksHoyde Then
        Me.Height = maksHoyde
        Me.ScrollBars = fmScrollBarsVertical
    Else
        Me.Height = onsketHoyde
        Me.ScrollBars = fmScrollBarsNone
    End If
End Sub

Private Sub FyllKolonnelister(ByVal tbl As ListObject)
    Dim kol As ListColumn
    Dim i As Long

    cboEpost.Clear
    cboStatus.Clear
    cboCC.Clear
    cboEpost.AddItem ""
    cboStatus.AddItem ""
    cboCC.AddItem ""
    For Each kol In tbl.ListColumns
        cboEpost.AddItem kol.Name
        cboStatus.AddItem kol.Name
        cboCC.AddItem kol.Name
    Next kol

    For i = 1 To modGenerisk.MAKS_PLASSHOLDERE
        Me.Controls("cboKol" & i).Clear
        Me.Controls("cboKol" & i).AddItem ""
        For Each kol In tbl.ListColumns
            Me.Controls("cboKol" & i).AddItem kol.Name
        Next kol
    Next i
End Sub

Private Sub cmdVelgMalFil_Click()
    Dim dlg As FileDialog
    Dim valgtSti As String
    Dim basisSti As String

    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    dlg.Title = "Velg mal-fila (.msg)"
    dlg.Filters.Clear
    dlg.Filters.Add "Outlook-mal", "*.msg"
    dlg.AllowMultiSelect = False
    If dlg.Show <> -1 Then Exit Sub
    valgtSti = dlg.SelectedItems(1)

    basisSti = modGenerisk.LokalStiFraSky(ThisWorkbook.Path)
    If LCase(Left(valgtSti, Len(basisSti))) <> LCase(basisSti) Then
        MsgBox "Mal-fila må ligge i samme mappe som denne Excel-fila, eller i en undermappe av den." & vbCrLf & _
               "Flytt fila dit først, og prøv igjen.", vbExclamation, "Oppsett"
        Exit Sub
    End If

    Dim relativ As String
    relativ = Mid(valgtSti, Len(basisSti) + 2)
    Dim posSisteSkille As Long
    posSisteSkille = InStrRev(relativ, "\")
    Dim undermappe As String, filnavn As String
    If posSisteSkille > 0 Then
        undermappe = Left(relativ, posSisteSkille - 1)
        filnavn = Mid(relativ, posSisteSkille + 1)
    Else
        undermappe = ""
        filnavn = relativ
    End If

    modGenerisk.SkrivKonfigVerdi "MAL_UNDERMAPPE", undermappe
    modGenerisk.SkrivKonfigVerdi "MAL_FILNAVN", filnavn
    lblMalFil.Caption = IIf(undermappe = "", filnavn, undermappe & "\" & filnavn)
    OppdaterMalfilLayout

    LastPlassholdereFraFil valgtSti
End Sub

Private Sub LastPlassholdereFraFil(ByVal sti As String)
    Dim olApp As Object, olMail As Object
    Dim tekst As String
    Dim plassholdere As Collection
    Dim konfig As Object
    Dim i As Long
    Dim p As Variant
    Dim antallVist As Long

    Set olApp = modGenerisk.HentOutlookApp()
    If olApp Is Nothing Then
        lblMelding.Caption = "Klarte ikke å koble til Outlook for å lese mal-fila."
        Exit Sub
    End If

    On Error Resume Next
    Set olMail = olApp.CreateItemFromTemplate(sti)
    On Error GoTo 0
    If olMail Is Nothing Then
        lblMelding.Caption = "Klarte ikke å åpne mal-fila."
        Exit Sub
    End If

    tekst = olMail.Subject & " " & olMail.HTMLBody

    ' Vis emne og brødtekst (som ren tekst - .Body er HTMLBody uten
    ' HTML-koder) så brukeren ser sammenhengen mens plassholdere kobles.
    txtEmne.Value = olMail.Subject
    txtBrodtekst.Value = olMail.Body

    olMail.Close 1 ' olDiscard - vi skal bare lese, ikke lagre noe utkast

    Set plassholdere = modGenerisk.FinnPlassholdere(tekst)
    Set konfig = modGenerisk.LesKonfig()

    For i = 1 To modGenerisk.MAKS_PLASSHOLDERE
        Me.Controls("lblPH" & i).Visible = False
        Me.Controls("cboKol" & i).Visible = False
    Next i

    If plassholdere.Count = 0 Then
        lblMelding.Caption = "Fant ingen plassholdere (f.eks. AAA, BBB) i mal-fila."
        OppdaterLayout 0
        Exit Sub
    End If

    antallVist = plassholdere.Count
    If antallVist > modGenerisk.MAKS_PLASSHOLDERE Then
        antallVist = modGenerisk.MAKS_PLASSHOLDERE
        lblMelding.Caption = "Fant " & plassholdere.Count & " plassholdere, men kun de første " & _
                              modGenerisk.MAKS_PLASSHOLDERE & " kan kobles i dette vinduet."
    Else
        ' Normaltilfellet trenger ingen egen melding - plassholder-radene
        ' vises jo allerede rett under. Tom melding gir OppdaterLayout
        ' beskjed om a flytte Lagre-knappen opp i det tomrommet.
        lblMelding.Caption = ""
    End If

    i = 0
    For Each p In plassholdere
        i = i + 1
        If i > modGenerisk.MAKS_PLASSHOLDERE Then Exit For
        Me.Controls("lblPH" & i).Caption = p & ":"
        Me.Controls("lblPH" & i).Visible = True
        Me.Controls("cboKol" & i).Visible = True
        Me.Controls("cboKol" & i).Tag = p
        If konfig.Exists("PH:" & p) Then
            ' Samme fallgruve som cboEpost/cboStatus/cboCC over: cboKol er
            ' ogsa Style=2 (fmStyleDropDownList) - en lagret kolonnekobling
            ' som ikke lenger finnes i tabellen skal hoppes stille over,
            ' ikke krasje Oppsett.
            On Error Resume Next
            Me.Controls("cboKol" & i).Value = konfig("PH:" & p)
            On Error GoTo 0
        Else
            Me.Controls("cboKol" & i).Value = ""
        End If
    Next p

    ' Ma kjores pa nytt her - brodtekst/vedlegg-innholdet (og dermed
    ' hoydene deres) er akkurat na fylt inn, sa den forrige beregningen
    ' fra Initialize var basert pa tom/gammel tekst.
    OppdaterMalfilLayout

    OppdaterLayout antallVist
End Sub

Private Sub cmdLagre_Click()
    Dim i As Long
    Dim p As String

    If cboEpost.Value = "" Or cboStatus.Value = "" Then
        MsgBox "Velg både e-post- og statuskolonne først.", vbExclamation, "Oppsett"
        Exit Sub
    End If

    If Trim(txtStatusTekst.Value) = "" Then
        MsgBox "Skriv inn hvilken tekst som skal brukes i statuskolonnen.", vbExclamation, "Oppsett"
        Exit Sub
    End If

    modGenerisk.SkrivKonfigVerdi "KOL_EPOST", cboEpost.Value
    modGenerisk.SkrivKonfigVerdi "KOL_STATUS", cboStatus.Value
    modGenerisk.SkrivKonfigVerdi "KOL_CC", cboCC.Value
    modGenerisk.SkrivKonfigVerdi "STATUS_OMVENDT", IIf(optRetningOmvendt.Value, "1", "0")
    modGenerisk.SkrivKonfigVerdi "STATUS_TEKST", Trim(txtStatusTekst.Value)

    For i = 1 To modGenerisk.MAKS_PLASSHOLDERE
        If Me.Controls("cboKol" & i).Visible Then
            p = Me.Controls("cboKol" & i).Tag
            If Me.Controls("cboKol" & i).Value <> "" Then
                modGenerisk.SkrivKonfigVerdi "PH:" & p, Me.Controls("cboKol" & i).Value
            End If
        End If
    Next i

    Dim avsenderAaLagre As String
    avsenderAaLagre = cboAvsender.Value
    If avsenderAaLagre = "Automatisk" Then avsenderAaLagre = ""
    modGenerisk.SkrivKonfigVerdi "AVSENDER_VALGT", avsenderAaLagre

    If optLeverMinutter.Value Then
        If Not IsNumeric(txtLeverMinutter.Value) Or Val(txtLeverMinutter.Value) <= 0 Then
            MsgBox "Skriv inn et gyldig antall minutter (større enn 0).", vbExclamation, "Oppsett"
            Exit Sub
        End If
        modGenerisk.SkrivKonfigVerdi "LEVERING_TYPE", "MINUTTER"
        modGenerisk.SkrivKonfigVerdi "LEVERING_MINUTTER", CStr(Val(txtLeverMinutter.Value))
    ElseIf optLeverTidspunkt.Value Then
        If cboLeverDag.Value = "" Or cboLeverManed.ListIndex = -1 Or cboLeverAar.Value = "" Or cboLeverKlokkeslett.Value = "" Then
            MsgBox "Velg dag, måned, år og klokkeslett for utsatt levering.", vbExclamation, "Oppsett"
            Exit Sub
        End If
        Dim tidspunktTest As Date
        On Error Resume Next
        tidspunktTest = DateSerial(CLng(cboLeverAar.Value), cboLeverManed.ListIndex + 1, CLng(cboLeverDag.Value)) + TimeValue(cboLeverKlokkeslett.Value)
        If Err.Number <> 0 Then
            On Error GoTo 0
            MsgBox "Klarte ikke å tolke dato/klokkeslett.", vbExclamation, "Oppsett"
            Exit Sub
        End If
        On Error GoTo 0
        If tidspunktTest <= Now Then
            MsgBox "Det angitte leveringstidspunktet er allerede passert. Velg et tidspunkt frem i tid.", vbExclamation, "Oppsett"
            Exit Sub
        End If
        modGenerisk.SkrivKonfigVerdi "LEVERING_TYPE", "TIDSPUNKT"
        modGenerisk.SkrivKonfigVerdi "LEVERING_DAG", cboLeverDag.Value
        modGenerisk.SkrivKonfigVerdi "LEVERING_MANED", CStr(cboLeverManed.ListIndex + 1)
        modGenerisk.SkrivKonfigVerdi "LEVERING_AAR", cboLeverAar.Value
        modGenerisk.SkrivKonfigVerdi "LEVERING_KLOKKESLETT", cboLeverKlokkeslett.Value
    Else
        modGenerisk.SkrivKonfigVerdi "LEVERING_TYPE", "MED_ENGANG"
    End If

    ' Unload Me FORST, deretter navigasjonen INLINE i denne samme Sub-en -
    ' IKKE via et kall til en delt hjelpe-Sub (slik det var frem til
    ' 2026-09-10). Håkon bekreftet at et tidligere forsok med rekkefolgen
    ' "vis neste skjema FOR Unload Me" (via en delt Sub) fikk neste skjema
    ' til a apne seg riktig, MEN Oppsett selv ble ikke lastet ut (hang
    ' igjen bak, siden Unload Me matte vente pa at det nye modale skjemaet
    ' ble lukket forst). Alle ANDRE "ga tilbake"-knapper i appen (frmMasse-
    ' Send og frmFeilsjekk sin Tilbake) gjor navigasjonen direkte INLINE i
    ' samme Sub rett etter Unload Me, uten et eget Sub-kall - og de virker
    ' feilfritt. Samme mønster brukt her na.
    Unload Me
    On Error Resume Next
    Err.Clear
    AppActivate Application.Caption
    If modGenerisk.OppsettReturnerTil = "frmMasseSend" Then
        frmMasseSend.Show
    Else
        frmSendMail.Show
    End If
    If Err.Number <> 0 Then
        MsgBox "Oppsettet ble lagret, men klarte ikke å åpne vinduet igjen etterpå (feil " & Err.Number & ": " & Err.Description & ")." & vbCrLf & vbCrLf & _
               "Trykk på ""Mail-utsender""-knappen på arket for å åpne på nytt - ingenting du gjorde nå er tapt.", vbExclamation, "Oppsett"
    End If
    On Error GoTo 0
End Sub

Private Sub cmdTilbake_Click()
    ' Samme navigasjon som Lagre, men UTEN a lagre noe forst - for a bla
    ' tilbake uten a gjore endringer (Håkons eksplisitte onske). Se
    ' kommentaren i cmdLagre_Click for hvorfor dette gjores inline her og
    ' ikke via en delt Sub.
    Unload Me
    On Error Resume Next
    Err.Clear
    AppActivate Application.Caption
    If modGenerisk.OppsettReturnerTil = "frmMasseSend" Then
        frmMasseSend.Show
    Else
        frmSendMail.Show
    End If
    If Err.Number <> 0 Then
        MsgBox "Klarte ikke å åpne vinduet igjen etterpå (feil " & Err.Number & ": " & Err.Description & ")." & vbCrLf & vbCrLf & _
               "Trykk på ""Mail-utsender""-knappen på arket for å åpne på nytt.", vbExclamation, "Oppsett"
    End If
    On Error GoTo 0
End Sub
'@

# ---------------------------------------------------------------
# VBA-kildekode: frmMasseSend (velg og send flere mailer samtidig)
# I FAST PRODUKSJONSBRUK (bekreftet 2026-09-09, brukt til a sende 196
# ekte mailer). Bryter bevisst med prinsippet resten av verktoyet folger
# (aldri sende automatisk uten at et menneske ser pa hver enkelt mail
# forst) - derfor: tydelig bekreftelsesdialog med antall, stille
# feilhandtering (ingen MsgBox per rad), liten pause mellom hver sending
# (Exchange-vennlig).
# ---------------------------------------------------------------
$formCodeMasseSend = @'
Option Explicit

Dim valgt() As Boolean

' --- Rad-klikk-sporing (klikk i avkrysningsboksen huker av/pa) ---
' SYVENDE forsok (2026-09-11) - se [[vba_listbox_ctrl_shift_saga]] i minnet
' for de seks foregaende forsokene og hvorfor de feilet, inkludert funnet om
' at Click-hendelsen rett og slett ikke fyrer her - alt haandteres na i
' MouseUp (se lstMailer_MouseUp).
' sisteReferanseIndeks: forrige rad klikket (i avkrysningsboksen ELLER
' andre steder) - utgangspunktet for et shift-utvalg i avkrysningsboksen.
Dim sisteReferanseIndeks As Long

' Onsket lstMailer.Height fra TilpassVindusstorrelse, satt av
' UserForm_Activate - se der for hvorfor det ma vaere sa sent.
Dim onsketListeHoyde As Single

' Bredden pa selve avkrysningskolonnen (kolonne 0) - se FyllMailerListe sin
' "bredder = "24 pt;..."" - matcher her sa MouseDown vet om X-koordinaten
' traff avkrysningsboksen eller resten av raden.
Private Const AVKRYSNING_BREDDE As Single = 24

Private Sub UserForm_Initialize()
    lblTittel.Font.Bold = True
    fraForhandsvisning.Font.Bold = True
    cmdSendTilAlle.Font.Bold = True
    cmdSendTilAlle.BackColor = RGB(241, 58, 48)
    cmdSendTilAlle.ForeColor = RGB(255, 255, 255)

    sisteReferanseIndeks = -1

    ' fmMultiSelectExtended (satt femte forsok, 2026-09-10) beholdes selv om
    ' avkrysning na styres av klikkeposisjon, ikke native markering - den
    ' gir fortsatt palitelig .ListIndex ved Shift-klikk (se
    ' [[vba_listbox_ctrl_shift_saga]]), og den blaa fargemarkeringen er en
    ' harmlos, nyttig visuell bivirkning (viser hvilken rad som sist ble
    ' klikket/shift-utvidet).
    lstMailer.MultiSelect = fmMultiSelectExtended

    On Error Resume Next
    onsketListeHoyde = TilpassVindusstorrelse()
    On Error GoTo 0
    FyllMailerListe
End Sub

Private Sub UserForm_Activate()
    ' Funnet 2026-09-11 (COM-diagnostikk): a sette lstMailer.Height ETTER
    ' bade Me.Height og FyllMailerListe, men FORTSATT inne i
    ' UserForm_Initialize, var IKKE nok - Height hoppet likevel tilbake til
    ' sin opprinnelige storrelse. Det betyr at selve re-layouten som
    ' tilbakestiller den skjer nar VBA aktiverer/tegner vinduet for forste
    ' gang - ETTER at Initialize er helt ferdig. UserForm_Activate fyres
    ' etter dette, sa det er forste sted en satt Height faktisk blir
    ' staende. Se TilpassVindusstorrelse for selve beregningen.
    If onsketListeHoyde > 0 Then lstMailer.Height = onsketListeHoyde
End Sub

' Returnerer onsket sluttHoyde for lstMailer, eller 0 hvis ingen krymping
' trengs. MA kalles fra UserForm_Initialize, og selve lstMailer.Height MA
' settes av KALLEREN - HELT TIL SLUTT, etter bade Me.Height OG
' FyllMailerListe. Se forklaring under for hvorfor.
Private Function TilpassVindusstorrelse() As Single
    ' Byttet 2026-09-11 til SAMME monster som frmSendMail bruker (der det
    ' aldri har vaert et problem) - krymper selve LISTEN (som far sitt eget
    ' naturlige rullefelt) og flytter alt UNDER den opp tilsvarende, i
    ' stedet for a stole pa at Me.ScrollBars/Me.ScrollHeight pa hele
    ' skjemaet fungerer palitelig. Det gjorde det ikke i praksis - klippet
    ' bunn-knapperaden gjentatte ganger (Håkon, 2026-09-11).
    '
    ' Feil funnet 2026-09-11 (Håkon: "hele vinduet ikke synes"): denne
    ' kaltes tidligere med "On Error Resume Next" rundt seg og satte selv
    ' lstMailer.Height direkte - hvis krympingen pa et TRANGT skjerm/DPI-
    ' oppsett trakk Height under 0 (ingen nedre grense fantes), feilet
    ' akkurat DEN linja og resten av Subet (alle .Top-justeringene under)
    ' ble ALDRI kjort. Fikset med nedre grenser (se under).
    '
    ' Funnet 2026-09-11 (COM-diagnostikk, presist innsirklet linje for
    ' linje): SELV MED nedre grense holdt ikke lstMailer.Height seg -
    ' bade "Me.Height = ..." (a krympe SKJEMAET selv) OG FyllMailerListe
    ' sin .List/.ColumnWidths-tildeling (kalt av UserForm_Initialize rett
    ' etter denne funksjonen) fikk uavhengig av hverandre lstMailer til a
    ' HOPPE TILBAKE til sin opprinnelige storrelse - bekreftet ved a lese
    ' Height tilbake umiddelbart for og etter hver av disse to linjene.
    ' IntegralHeight=False endret ingenting (kvirken skjer ikke via
    ' radhoyde-snapping, men et sted i skjemaets egen re-layout nar enten
    ' DETS EGEN Height ELLER lstMailer sitt eget innhold endres). Eneste
    ' palitelige fiks: la denne funksjonen KUN beregne og returnere onsket
    ' hoyde, og la kalleren sette lstMailer.Height HELT TIL SLUTT - etter
    ' bade Me.Height og FyllMailerListe - da er det ingenting igjen som kan
    ' tilbakestille den.
    Dim maksHoyde As Single, differanse As Single
    Dim kuttListe As Single, kuttBrodtekst As Single

    lstMailer.IntegralHeight = False

    maksHoyde = Application.UsableHeight - 30
    If Me.Height <= maksHoyde Then Exit Function

    differanse = Me.Height - maksHoyde

    kuttListe = differanse
    If kuttListe > lstMailer.Height - 60 Then kuttListe = lstMailer.Height - 60
    If kuttListe < 0 Then kuttListe = 0

    lblHint.Top = lblHint.Top - kuttListe
    lblAntallValgt.Top = lblAntallValgt.Top - kuttListe
    lblMarkerX.Top = lblMarkerX.Top - kuttListe
    txtMarkerX.Top = txtMarkerX.Top - kuttListe
    cmdMarkerX.Top = cmdMarkerX.Top - kuttListe
    cmdVelgAlle.Top = cmdVelgAlle.Top - kuttListe
    cmdFjernAlle.Top = cmdFjernAlle.Top - kuttListe
    fraForhandsvisning.Top = fraForhandsvisning.Top - kuttListe

    kuttBrodtekst = differanse - kuttListe
    If kuttBrodtekst > txtPrevBrodtekst.Height - 40 Then kuttBrodtekst = txtPrevBrodtekst.Height - 40
    If kuttBrodtekst < 0 Then kuttBrodtekst = 0
    txtPrevBrodtekst.Height = txtPrevBrodtekst.Height - kuttBrodtekst
    fraForhandsvisning.Height = fraForhandsvisning.Height - kuttBrodtekst
    lblPrevVedleggLabel.Top = lblPrevVedleggLabel.Top - kuttBrodtekst
    txtPrevVedlegg.Top = txtPrevVedlegg.Top - kuttBrodtekst

    Dim brukt As Single
    brukt = kuttListe + kuttBrodtekst
    TilpassVindusstorrelse = lstMailer.Height - kuttListe

    Me.Height = Me.Height - brukt
    lblStatus.Top = lblStatus.Top - brukt
    cmdOppsett.Top = cmdOppsett.Top - brukt
    cmdLukk.Top = cmdLukk.Top - brukt
    cmdFeilsjekk.Top = cmdFeilsjekk.Top - brukt
    cmdSendTilAlle.Top = cmdSendTilAlle.Top - brukt
End Function

Private Sub FyllMailerListe()
    Dim tbl As ListObject
    Dim konfig As Object
    Dim colStatus As Long
    Dim antKolonner As Long, antRader As Long, r As Long, c As Long
    Dim data() As Variant
    Dim maxLengder() As Long
    Dim radIndekser() As Long
    Dim n As Long
    Dim bredder As String
    Dim status As String
    Dim i As Long

    cmdSendTilAlle.Enabled = False

    Set tbl = modGenerisk.AktivTabell()
    If tbl Is Nothing Then
        lblStatus.Caption = "Fant ingen tabell i arbeidsboka."
        Exit Sub
    End If

    Set konfig = modGenerisk.LesKonfig()
    If Not konfig.Exists("KOL_STATUS") Then
        lblStatus.Caption = "Oppsettet er ikke fullført."
        Exit Sub
    End If

    colStatus = modGenerisk.KolonneNr(tbl, konfig("KOL_STATUS"))
    If colStatus = 0 Then
        lblStatus.Caption = "Statuskolonnen fra Oppsett finnes ikke lenger i tabellen."
        Exit Sub
    End If

    If tbl.DataBodyRange Is Nothing Then
        lblStatus.Caption = "Tabellen har ingen datarader ennå."
        Exit Sub
    End If

    Dim synligeKol As Collection, kolIdx As Variant
    Set synligeKol = modGenerisk.SynligeKolonneIndekser(tbl)
    antKolonner = synligeKol.Count
    antRader = tbl.DataBodyRange.Rows.Count

    ReDim data(1 To antRader, 1 To antKolonner)
    ReDim radIndekser(1 To antRader)
    ReDim maxLengder(1 To antKolonner)
    n = 0

    ' Skjulte/filtrerte rader (f.eks. via autofilter, eller kolonner
    ' skjult med Kolonnevelger-makroen) er ikke synlige i selve arket
    ' akkurat na, sa de tas heller ikke med her.
    For r = 1 To antRader
        If Not tbl.DataBodyRange.Rows(r).EntireRow.Hidden Then
            status = Trim(CStr(tbl.DataBodyRange.Cells(r, colStatus).Value))
            If modGenerisk.StatusRadErAktuell(status) Then
                n = n + 1
                c = 0
                For Each kolIdx In synligeKol
                    c = c + 1
                    data(n, c) = CStr(tbl.DataBodyRange.Cells(r, CLng(kolIdx)).Value)
                    If Len(data(n, c)) > maxLengder(c) Then maxLengder(c) = Len(data(n, c))
                Next kolIdx
                radIndekser(n) = r
            End If
        End If
    Next r

    Me.Tag = ""
    For i = 1 To n
        Me.Tag = Me.Tag & radIndekser(i) & ","
    Next i

    ReDim valgt(0 To IIf(n > 0, n - 1, 0))

    ' Kolonnebredde basert pa faktisk innhold (lengste verdi i kolonnen) i
    ' stedet for en fast "90 pt" for alle - unngar at lange verdier (f.eks.
    ' skolenavn) klippes nar det er mange rader/kolonner. Minimum 30pt,
    ' maks 200pt (en uvanlig lang enkeltverdi skal ikke sprenge hele lista),
    ' ca 6pt per tegn - en grov, men god nok tilnaerming for standard
    ' listeskrift.
    lstMailer.ColumnCount = antKolonner + 1
    bredder = "24 pt;"
    For c = 1 To antKolonner
        Dim kolBredde As Long
        kolBredde = maxLengder(c) * 6
        If kolBredde < 30 Then kolBredde = 30
        If kolBredde > 200 Then kolBredde = 200
        bredder = bredder & kolBredde & " pt;"
    Next c
    lstMailer.ColumnWidths = bredder

    ' Rader med manglende data (mangler e-post, eller en plassholder-kolonne
    ' er tom) far et varseltegn foran forste kolonne - MSForms.ListBox har
    ' ingen per-rad skriftfarge (kun ETT ForeColor for hele kontrollen), sa
    ' et symbol i selve teksten er den praktiske maten a flagge dem pa.
    Dim manglendeRader As Object
    Set manglendeRader = CreateObject("Scripting.Dictionary")
    Dim funn As Object
    On Error Resume Next
    For Each funn In modGenerisk.FinnManglendeData()
        manglendeRader(CLng(funn("RadNr"))) = True
    Next funn
    On Error GoTo 0

    If n = 0 Then
        lstMailer.Clear
    Else
        Dim visning() As Variant
        ReDim visning(1 To n, 1 To antKolonner + 1)
        For r = 1 To n
            visning(r, 1) = ChrW(9744) ' tomt avkrysningsboks-symbol
            For c = 1 To antKolonner
                visning(r, c + 1) = data(r, c)
            Next c
            If manglendeRader.Exists(radIndekser(r)) Then
                visning(r, 2) = ChrW(9888) & " " & visning(r, 2)
            End If
        Next r
        lstMailer.List = visning
    End If

    lblStatus.Caption = "Gjenstår å velge blant: " & n
    If manglendeRader.Count > 0 Then lblStatus.Caption = lblStatus.Caption & "   (⚠ = raden mangler data)"
    cmdSendTilAlle.Enabled = (n > 0)

    If manglendeRader.Count > 0 Then
        cmdFeilsjekk.Caption = "Undersøk feil (" & manglendeRader.Count & ")"
        cmdFeilsjekk.Font.Bold = True
        cmdFeilsjekk.ForeColor = RGB(190, 50, 50)
    Else
        cmdFeilsjekk.Caption = "Undersøk feil"
        cmdFeilsjekk.Font.Bold = False
        cmdFeilsjekk.ForeColor = RGB(0, 0, 0)
    End If
    ' Samme dynamiske bredde som frmSendMail (se der for hvorfor) - her er
    ' det ingen knapp rett etter a flytte (cmdSendTilAlle star bevisst
    ' langt til hoyre), sa kun selve bredden justeres.
    cmdFeilsjekk.Width = Len(cmdFeilsjekk.Caption) * 7 + 24
    If cmdFeilsjekk.Width < 110 Then cmdFeilsjekk.Width = 110
End Sub

' SYVENDE forsok (2026-09-11). Sjette forsok (klikk i egen avkrysningsboks,
' se [[vba_listbox_ctrl_shift_saga]]) hadde riktig logikk, bekreftet via
' selvtest - men fungerte fortsatt ikke for et EKTE klikk hos Håkon. Direkte
' COM-diagnostikk under et ekte, fysisk klikk viste noe overraskende: selve
' MouseDown-hendelsen fyrte alltid korrekt (riktig X-koordinat, riktig
' gjenkjent som "i avkrysningsboksen"), men Click-hendelsen fyrte ALDRI i
' det hele tatt - bekreftet bade for synteriske klikk (sendt direkte til
' vinduets meldingsko, se under) OG for Håkons eget fysiske klikk (en
' modulvariabel satt helt forst i Click forble False). Dette er selve
' rotarsaken til hele denne sagaen - ikke noe galt med toggle-logikken.
' Fiks: bruk MouseUp i stedet for Click. MouseUp er en rå, lavniva
' musehendelse (som MouseDown, som alltid har fyrt palitelig), ikke en
' hoyniva "gest"-hendelse - og gir X/Shift direkte, sa MouseDown trengs
' ikke lenger i det hele tatt for denne funksjonaliteten.
Private Sub lstMailer_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Presisert terminologi (Håkon, 2026-09-11): "huke av" = fa hake i
    ' avkrysningsboksen (valgt-arrayet/ChrW 9745), "markere" = vises i
    ' forhandsvisningen under (VisForhandsvisning). De to skjer ALDRI
    ' samtidig av samme klikk:
    '   - Klikk I avkrysningsboksen (kolonne 0): KUN huk av/pa DENNE raden
    '     - ALDRI forhandsvis. Shift+klikk her huker av/pa HELE intervallet
    '     siden forrige klikk, hver rad uavhengig av de andre.
    '   - Klikk ANDRE STEDER i raden: KUN forhandsvis - ALDRI huk av,
    '     uansett Shift.
    Dim indeks As Long
    indeks = lstMailer.ListIndex
    If indeks = -1 Then Exit Sub

    Dim iAvkrysningsboks As Boolean, shiftHoldt As Boolean
    iAvkrysningsboks = (X < AVKRYSNING_BREDDE)
    shiftHoldt = (Shift And 1) = 1

    If iAvkrysningsboks Then
        If shiftHoldt And sisteReferanseIndeks <> -1 Then
            Dim fraI As Long, tilI As Long, k As Long, tmp As Long
            fraI = sisteReferanseIndeks
            tilI = indeks
            If fraI > tilI Then
                tmp = fraI: fraI = tilI: tilI = tmp
            End If
            For k = fraI To tilI
                valgt(k) = Not valgt(k)
                lstMailer.List(k, 0) = IIf(valgt(k), ChrW(9745), ChrW(9744))
            Next k
        Else
            valgt(indeks) = Not valgt(indeks)
            lstMailer.List(indeks, 0) = IIf(valgt(indeks), ChrW(9745), ChrW(9744))
        End If
        sisteReferanseIndeks = indeks
        OppdaterAntallValgt
        Exit Sub
    End If

    sisteReferanseIndeks = indeks

    ' DoEvents FOR VisForhandsvisning (Håkon, 2026-09-11, video: "markering
    ' henger igjen fra vanlig trykk") - vanlig klikk bygger en EKTE Outlook-
    ' mail via COM i VisForhandsvisning (treg); DoEvents tvinger ventende
    ' tegne-meldinger til a bli behandlet FOR den trege delen starter, sa
    ' markoringen rekker a flytte seg visuelt med en gang.
    DoEvents
    VisForhandsvisning indeks
End Sub

Private Sub OppdaterAntallValgt()
    Dim i As Long, n As Long
    For i = 0 To lstMailer.ListCount - 1
        If valgt(i) Then n = n + 1
    Next i
    lblAntallValgt.Caption = n & " valgt"
End Sub

Private Sub cmdVelgAlle_Click()
    Dim i As Long
    For i = 0 To lstMailer.ListCount - 1
        valgt(i) = True
        lstMailer.List(i, 0) = ChrW(9745)
    Next i
    OppdaterAntallValgt
End Sub

Private Sub cmdFjernAlle_Click()
    Dim i As Long
    For i = 0 To lstMailer.ListCount - 1
        valgt(i) = False
        lstMailer.List(i, 0) = ChrW(9744)
    Next i
    OppdaterAntallValgt
End Sub

Private Sub cmdMarkerX_Click()
    ' Huker av de forste X radene (ovenfra og ned) - X skrives inn i
    ' txtMarkerX. Nyttig for a raskt markere en gitt mengde uten a
    ' dobbeltklikke/shift-klikke seg gjennom mange rader manuelt.
    If Not IsNumeric(txtMarkerX.Value) Or Val(txtMarkerX.Value) <= 0 Then
        MsgBox "Skriv inn et gyldig antall (større enn 0) i feltet ved siden av knappen.", vbExclamation, "Marker X"
        Exit Sub
    End If

    Dim antall As Long, i As Long
    antall = CLng(Val(txtMarkerX.Value))
    If antall > lstMailer.ListCount Then antall = lstMailer.ListCount

    For i = 0 To antall - 1
        valgt(i) = True
        lstMailer.List(i, 0) = ChrW(9745)
    Next i
    OppdaterAntallValgt
End Sub

Private Sub VisForhandsvisning(ByVal i As Long)
    Dim radIndekser() As String
    Dim ekteRad As Long
    Dim olMail As Object, feilmelding As String
    Dim fraVisning As String

    radIndekser = Split(Me.Tag, ",")
    ekteRad = CLng(radIndekser(i))

    txtPrevEmne.Value = ""
    txtPrevTil.Value = ""
    txtPrevCC.Value = ""
    txtPrevFra.Value = ""
    txtPrevBrodtekst.Value = ""
    txtPrevVedlegg.Value = ""

    If Not modGenerisk.BygMailForRad(ekteRad, olMail, feilmelding) Then
        txtPrevBrodtekst.Value = "(Kan ikke forhåndsvise: " & feilmelding & ")"
        Exit Sub
    End If

    txtPrevEmne.Value = olMail.Subject
    txtPrevTil.Value = olMail.To

    On Error Resume Next
    txtPrevCC.Value = olMail.CC
    On Error GoTo 0
    If txtPrevCC.Value = "" Then txtPrevCC.Value = "(ingen)"

    fraVisning = ""
    On Error Resume Next
    fraVisning = olMail.SentOnBehalfOfName
    On Error GoTo 0
    If fraVisning = "" Then fraVisning = "(Automatisk - kontoen som kjører makroen)"
    txtPrevFra.Value = fraVisning

    ' olMail.Body er Outlook sin EGEN, automatiske HTML-til-ren-tekst-
    ' konvertering av HTMLBody (som er det BygMailForRad faktisk setter) -
    ' den konverteringen legger ofte til flere tomme linjer pa SLUTTEN enn
    ' det som er synlig ved forste oyekast (typisk fra avsnitts-/div-tagger
    ' i malen), noe som gjorde forhandsvisningsboksen mye storre enn selve
    ' e-posten (Håkon, 2026-09-11: "brødtekst feeltet er fortsatt større
    ' enn selve mailen"). Etterslepende tomme linjer trimmes derfor bort
    ' FOR den vises OG FOR hoyden regnes ut - ikke bare i hoyde-formelen,
    ' sa selve den viste teksten heller ikke har usynlig tomrom man matte
    ' rulle forbi.
    Dim fpTekst As String
    Dim fpLinjeArr() As String, fpI As Long, fpRadLengde As Long, fpSisteLinje As Long
    fpTekst = Replace(Replace(olMail.Body, vbCrLf, vbLf), vbCr, vbLf)
    fpLinjeArr = Split(fpTekst, vbLf)
    fpSisteLinje = UBound(fpLinjeArr)
    Do While fpSisteLinje > LBound(fpLinjeArr) And Trim(fpLinjeArr(fpSisteLinje)) = ""
        fpSisteLinje = fpSisteLinje - 1
    Loop
    fpTekst = ""
    For fpI = LBound(fpLinjeArr) To fpSisteLinje
        If fpI > LBound(fpLinjeArr) Then fpTekst = fpTekst & vbCrLf
        fpTekst = fpTekst & fpLinjeArr(fpI)
    Next fpI
    txtPrevBrodtekst.Value = fpTekst

    ' Hoyden tilpasses den (na trimmede) teksten - samme teknikk/
    ' begrunnelse som OppdaterMalfilLayout i frmOppsett. Vedlegg-raden
    ' flyttes tilsvarende opp etter boksens nye, mindre hoyde.
    Dim fpLinjer As Long, fpHoyde As Long
    Const FP_TEGN_PER_LINJE As Long = 95
    fpLinjer = 0
    For fpI = LBound(fpLinjeArr) To fpSisteLinje
        fpRadLengde = Len(fpLinjeArr(fpI))
        If fpRadLengde = 0 Then
            fpLinjer = fpLinjer + 1
        Else
            fpLinjer = fpLinjer + Int((fpRadLengde - 1) \ FP_TEGN_PER_LINJE) + 1
        End If
    Next fpI
    If fpLinjer < 1 Then fpLinjer = 1
    fpHoyde = fpLinjer * 15 + 14
    If fpHoyde < 40 Then fpHoyde = 40
    If fpHoyde > 110 Then fpHoyde = 110 ' svaert lange mailer far egen rullefelt i stedet
    txtPrevBrodtekst.Height = fpHoyde
    lblPrevVedleggLabel.Top = txtPrevBrodtekst.Top + txtPrevBrodtekst.Height + 10
    txtPrevVedlegg.Top = lblPrevVedleggLabel.Top

    txtPrevVedlegg.Value = Replace(modGenerisk.VedleggTekst(), vbCrLf, ", ")

    olMail.Close 1 ' olDiscard - kun forhandsvisning, ikke la den bli et hengende utkast
End Sub

Private Sub cmdSendTilAlle_Click()
    Dim antallValgt As Long, i As Long
    Dim radIndekser() As String
    Dim tbl As ListObject, konfig As Object, colStatus As Long
    Dim svar As VbMsgBoxResult
    Dim antallSendt As Long, antallHoppetOver As Long, hoppOverMelding As String
    Dim olMail As Object, feilmelding As String
    Dim ekteRad As Long
    Dim malAntall As Long

    antallValgt = 0
    For i = 0 To lstMailer.ListCount - 1
        If valgt(i) Then antallValgt = antallValgt + 1
    Next i

    If antallValgt = 0 Then
        MsgBox "Ingen mailer er huket av. Dobbeltklikk en rad for å huke den av.", vbExclamation, "Send til flere"
        Exit Sub
    End If

    svar = MsgBox("Dette vil sende " & antallValgt & " e-post(er) med det samme, uten flere bekreftelser underveis." & vbCrLf & vbCrLf & _
                   "Avsender: " & modGenerisk.AvsenderBeskrivelse() & vbCrLf & _
                   "Levering: " & modGenerisk.LeveringsBeskrivelse() & vbCrLf & vbCrLf & _
                   "Er du helt sikker på at du vil fortsette?", vbYesNo + vbExclamation + vbDefaultButton2, "Send til flere - bekreft")
    If svar <> vbYes Then Exit Sub

    Set tbl = modGenerisk.AktivTabell()
    Set konfig = modGenerisk.LesKonfig()
    colStatus = modGenerisk.KolonneNr(tbl, konfig("KOL_STATUS"))

    radIndekser = Split(Me.Tag, ",")
    antallSendt = 0
    antallHoppetOver = 0
    hoppOverMelding = ""
    malAntall = antallValgt

    cmdSendTilAlle.Enabled = False
    cmdVelgAlle.Enabled = False
    cmdFjernAlle.Enabled = False

    ' Fet og farget mens sendingen pagar - tydeligere a folge med pa enn
    ' vanlig, tynn svart tekst (Håkon, 2026-09-10, etter a ha sendt 196
    ' mailer og ma ha oyet pa denne teksten en god stund).
    lblStatus.Font.Bold = True
    lblStatus.ForeColor = RGB(0, 90, 170)

    For i = 0 To lstMailer.ListCount - 1
        If valgt(i) Then
            ekteRad = CLng(radIndekser(i))
            lblStatus.Caption = "Sender " & (antallSendt + antallHoppetOver + 1) & " av " & malAntall & " ..."
            DoEvents

            If modGenerisk.BygMailForRad(ekteRad, olMail, feilmelding) Then
                olMail.Send
                If colStatus > 0 Then modGenerisk.StatusMarkerSendt tbl.DataBodyRange.Cells(ekteRad, colStatus)
                antallSendt = antallSendt + 1
            Else
                antallHoppetOver = antallHoppetOver + 1
                hoppOverMelding = hoppOverMelding & "Rad " & ekteRad & ": " & feilmelding & vbCrLf
            End If

            ' Liten pause mellom hver sending - unngar at Exchange
            ' oppfatter en rask serie som mistenkelig/spam-lignende.
            Application.Wait Now + TimeSerial(0, 0, 1)
        End If
    Next i

    Dim sluttmelding As String
    sluttmelding = "Ferdig." & vbCrLf & antallSendt & " sendt."
    If antallHoppetOver > 0 Then
        sluttmelding = sluttmelding & vbCrLf & antallHoppetOver & " hoppet over:" & vbCrLf & hoppOverMelding
    End If
    MsgBox sluttmelding, vbInformation, "Send til flere"

    Unload Me
    frmSendMail.Show
End Sub

Private Sub cmdOppsett_Click()
    modGenerisk.OppsettReturnerTil = "frmMasseSend"
    Unload Me
    frmOppsett.Show
End Sub

Private Sub cmdLukk_Click()
    Unload Me
    On Error Resume Next
    AppActivate Application.Caption ' se cmdLagre_Click i frmOppsett for hvorfor
    On Error GoTo 0
    frmSendMail.Show
End Sub

Private Sub cmdFeilsjekk_Click()
    ' Samme grunn som i frmSendMail - modalt skjema ma lukkes forst for et
    ' modelost skjema kan vises (feilkode 401).
    modGenerisk.FeilsjekkReturnerTil = "frmMasseSend"
    Unload Me
    frmFeilsjekk.Show vbModeless
End Sub
'@

# ---------------------------------------------------------------
# VBA-kildekode: frmFeilsjekk (finn rader med manglende plassholder-data)
# ---------------------------------------------------------------
$formCodeFeilsjekk = @'
Option Explicit

Private funn As Collection
Private gjeldendeIndeks As Long

Private Sub UserForm_Initialize()
    lblTittel.Font.Bold = True

    Set funn = modGenerisk.FinnManglendeData()
    gjeldendeIndeks = -1

    lstFeil.Clear
    Dim f As Object
    For Each f In funn
        lstFeil.AddItem "Rad " & f("RadNr") & ": mangler " & f("Plassholder") & " (kolonne '" & f("KolonneNavn") & "')"
    Next f

    If funn.Count = 0 Then
        lblTittel.Caption = "Ingen manglende data funnet - alle rader ser komplette ut."
        cmdGaTil.Enabled = False
        cmdForrige.Enabled = False
        cmdNeste.Enabled = False
    Else
        lblTittel.Caption = "Fant " & funn.Count & " rad(er)/kolonne(r) med manglende data:"
        gjeldendeIndeks = 0
        lstFeil.ListIndex = 0
    End If
    OppdaterStatus
End Sub

Private Sub OppdaterStatus()
    If funn Is Nothing Then Exit Sub
    If funn.Count = 0 Then
        lblStatus.Caption = ""
    Else
        lblStatus.Caption = "Feil " & (gjeldendeIndeks + 1) & " av " & funn.Count
    End If
End Sub

Private Sub GaTilFunn(ByVal indeks As Long)
    If funn Is Nothing Then Exit Sub
    If funn.Count = 0 Then Exit Sub
    If indeks < 0 Then indeks = funn.Count - 1
    If indeks > funn.Count - 1 Then indeks = 0
    gjeldendeIndeks = indeks
    lstFeil.ListIndex = indeks

    Dim f As Object
    Set f = funn(indeks + 1) ' VBA Collection er 1-basert, indeks er 0-basert
    Dim tbl As ListObject
    Set tbl = modGenerisk.AktivTabell()
    If tbl Is Nothing Then Exit Sub

    ' Beste-forsok a fa Excel-vinduet fremst og markere riktig celle - ikke
    ' kritisk om AppActivate feiler pa et system der vindustittelen avviker,
    ' derfor On Error Resume Next rundt hele blokken. Bruker .Select (ikke
    ' Application.Goto med Scroll:=True) - Select ruller vinduet akkurat nok
    ' til at cellen blir synlig, samme oppforsel som vanlig celle-navigering,
    ' i stedet for a tvinge cellen opp i ovre venstre hjorne av vinduet hver
    ' gang.
    On Error Resume Next
    Dim malCelle As Range
    Set malCelle = tbl.DataBodyRange.Cells(f("RadNr"), f("KolNr"))
    malCelle.Worksheet.Activate
    malCelle.Select
    AppActivate Application.Caption
    On Error GoTo 0

    OppdaterStatus
End Sub

Private Sub cmdGaTil_Click()
    If lstFeil.ListIndex = -1 Then Exit Sub
    GaTilFunn lstFeil.ListIndex
End Sub

Private Sub cmdForrige_Click()
    GaTilFunn gjeldendeIndeks - 1
End Sub

Private Sub cmdNeste_Click()
    GaTilFunn gjeldendeIndeks + 1
End Sub

Private Sub cmdLukk_Click()
    Unload Me
    On Error Resume Next
    AppActivate Application.Caption ' se cmdLagre_Click i frmOppsett for hvorfor
    On Error GoTo 0
    If modGenerisk.FeilsjekkReturnerTil = "frmMasseSend" Then
        frmMasseSend.Show
    Else
        frmSendMail.Show
    End If
End Sub
'@

# ---------------------------------------------------------------
# VBA-kildekode: frmKolonnevisning (velg hvilke kolonner som vises i
# utvalgslistene - egen, liten dialog apnet fra Oppsett) - NY 2026-09-23
# ---------------------------------------------------------------
$formCodeKolonnevisning = @'
Option Explicit

Private Sub UserForm_Initialize()
    lblTittel.Font.Bold = True
    cmdLagre.Font.Bold = True
    lblMelding.Caption = ""

    Dim tbl As ListObject
    Set tbl = modGenerisk.AktivTabell()
    If tbl Is Nothing Then
        lblMelding.Caption = "Fant ingen tabell i arbeidsboka."
        lstKolonner.Enabled = False
        cmdMerkAlle.Enabled = False
        cmdAvmerkAlle.Enabled = False
        cmdLagre.Enabled = False
        Exit Sub
    End If

    If tbl.ListColumns.Count = 0 Then
        lblMelding.Caption = "Tabellen har ingen kolonner."
        lstKolonner.Enabled = False
        cmdMerkAlle.Enabled = False
        cmdAvmerkAlle.Enabled = False
        cmdLagre.Enabled = False
        Exit Sub
    End If

    Dim konfig As Object, skjulte As String
    Set konfig = modGenerisk.LesKonfig()
    skjulte = ""
    If konfig.Exists("SKJULTE_LISTEKOLONNER") Then skjulte = konfig("SKJULTE_LISTEKOLONNER")

    lstKolonner.Clear
    lstKolonner.ListStyle = fmListStyleOption
    lstKolonner.MultiSelect = fmMultiSelectMulti

    Dim kol As ListColumn
    For Each kol In tbl.ListColumns
        lstKolonner.AddItem kol.Name
    Next kol

    Dim i As Long
    For i = 0 To lstKolonner.ListCount - 1
        lstKolonner.Selected(i) = (InStr(1, "|" & skjulte & "|", "|" & lstKolonner.List(i) & "|") = 0)
    Next i
End Sub

Private Sub cmdMerkAlle_Click()
    Dim i As Long
    For i = 0 To lstKolonner.ListCount - 1
        lstKolonner.Selected(i) = True
    Next i
End Sub

Private Sub cmdAvmerkAlle_Click()
    Dim i As Long
    For i = 0 To lstKolonner.ListCount - 1
        lstKolonner.Selected(i) = False
    Next i
End Sub

Private Sub cmdLagre_Click()
    Dim i As Long, skjulteListe As String, antallValgt As Long

    skjulteListe = ""
    antallValgt = 0
    For i = 0 To lstKolonner.ListCount - 1
        If lstKolonner.Selected(i) Then
            antallValgt = antallValgt + 1
        Else
            If skjulteListe <> "" Then skjulteListe = skjulteListe & "|"
            skjulteListe = skjulteListe & lstKolonner.List(i)
        End If
    Next i

    ' Minst en kolonne ma vaere valgt - en tom liste ville gitt en tom
    ' utvalgsliste i frmSendMail/frmMasseSend, ikke en krasj i seg selv,
    ' men et forvirrende "ingen rader vises noensinne"-utseende. Stoppes
    ' her i stedet, samme prinsipp som e-post-/statuskolonne-kravet i
    ' Oppsett (Håkon, 2026-09-23: innstillinger skal aldri kunne lagres i
    ' en tilstand som krasjer eller forvirrer senere).
    If antallValgt = 0 Then
        MsgBox "Minst én kolonne må være valgt.", vbExclamation, "Velg kolonner"
        Exit Sub
    End If

    modGenerisk.SkrivKonfigVerdi "SKJULTE_LISTEKOLONNER", skjulteListe
    Unload Me
End Sub

Private Sub cmdAvbryt_Click()
    Unload Me
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
    $dlgXlsm.Title = "Velg Excel-fila (.xlsm) som skal fa/oppdatere Mail-utsender"
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
# 1b) Mal-fil (.msg) - IKKE lenger en dialog her i installeren.
# Dette gjores na utelukkende via "Oppsett"-knappen inne i selve
# Excel-fila (frmOppsett.cmdVelgMalFil_Click, som regner ut samme
# relative sti selv) - installeren sin eneste jobb er a installere
# koden. -MalSti kan fortsatt sendes inn direkte for skriptet/
# automatisert bruk (f.eks. testing), men det apnes ingen dialog
# for det lenger.
# ---------------------------------------------------------------
if ($MalSti) {
    $xlsmMappe = (Get-Item (Split-Path $Path -Parent)).FullName
    $msgMappe = (Get-Item (Split-Path $MalSti -Parent)).FullName
    $baseUri = New-Object System.Uri(($xlsmMappe.TrimEnd('\') + '\'))
    $malUri = New-Object System.Uri(($msgMappe.TrimEnd('\') + '\'))
    $relUri = $baseUri.MakeRelativeUri($malUri)
    $relStr = [System.Uri]::UnescapeDataString($relUri.ToString()).TrimEnd('/').Replace('/', '\')

    if ($relStr.StartsWith('..')) {
        Write-Warning "Mal-fila ligger UTENFOR Excel-fila sin mappe-struktur ($relStr) - dette vil ikke fungere portabelt. Flytt mal-fila til samme mappe (eller en undermappe) som Excel-fila."
    }

    $MalUndermappe = $relStr
    $MalFilnavn = Split-Path $MalSti -Leaf
    Write-Output "  -> relativ undermappe (fra Excel-fila): '$MalUndermappe'"
    Write-Output "  -> filnavn: '$MalFilnavn'"
} else {
    Write-Output "Ingen mal-fil angitt her - sett den via 'Oppsett'-knappen inne i selve Excel-fila etter installasjon."
}

# ---------------------------------------------------------------
# 2) Koble til Excel og finne/apne malfila
# ---------------------------------------------------------------
try {
    if ($NyExcelInstans) {
        $excel = $null
    } else {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    }
} catch {
    $excel = $null
}

$wb = $null
if ($excel) {
    foreach ($w in $excel.Workbooks) {
        if ($w.FullName -eq $Path) { $wb = $w; break }
    }
}
$lukkVedSlutt = $false
if (-not $wb) {
    if (-not $excel) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
    }
    $wb = $excel.Workbooks.Open($Path)
    $lukkVedSlutt = $false
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

if ($Uninstall) {
    # Avinstallerer - fjerner KUN de nøyaktige, kjente komponentnavnene og
    # knappen med denne makroens EGET OnAction-navn, aldri et generisk
    # filter (se lesson om Kolonnevelger-knapp-hendelsen i minnet - et ark
    # kan ha flere uavhengig installerte makroers knapper side om side).
    # Den skjulte MailKonfig-fana la staa urort, med vilje - en senere
    # reinstallasjon gjenbruker da gamle innstillinger i stedet for at
    # brukeren ma sette opp alt pa nytt.
    Write-Output "Fjerner Mail-utsender fra $($wb.Name) ..."
    foreach ($navn in @("modGenerisk", "frmSendMail", "frmOppsett", "frmMasseSend", "frmFeilsjekk", "frmKolonnevisning")) {
        try {
            $eksisterende = $vbProject.VBComponents.Item($navn)
            $vbProject.VBComponents.Remove($eksisterende)
            Write-Output "  Fjernet $navn"
        } catch {}
    }
    try {
        foreach ($ws in $wb.Worksheets) {
            foreach ($btn in @($ws.Buttons())) {
                if ($btn.OnAction -eq "MailUtsendelse" -or $btn.OnAction -like "*!MailUtsendelse") { $btn.Delete() }
            }
        }
    } catch {}
    try { $wb.Save() } catch {
        Write-Output "FEIL ved lagring: $($_.Exception.Message)"
        exit 1
    }
    Write-Output "Ferdig - Mail-utsender er fjernet fra $($wb.Name)."
    exit 0
}

# ---------------------------------------------------------------
# 3) Sjekk installert versjon - hopp over hvis allerede oppdatert
# ---------------------------------------------------------------
$installertVersjon = $null
try {
    $eksisterendeModul = $vbProject.VBComponents.Item("modGenerisk")
    $kilde = $eksisterendeModul.CodeModule.Lines(1, $eksisterendeModul.CodeModule.CountOfLines)
    if ($kilde -match 'GENERISK_VERSION\s*As\s*String\s*=\s*"([^"]+)"') {
        $installertVersjon = $Matches[1]
    }
} catch {}

$nyVersjon = $GeneriskVersion

if ($installertVersjon -eq $nyVersjon -and -not $Force) {
    Write-Output "Mail-utsender er allerede installert og oppdatert (versjon $installertVersjon) i $($wb.Name)."
    Write-Output "(Innstillinger overskrives ikke automatisk naar versjonen er lik - kjor med -Force om du bare vil pushe pa nytt mal-fil-valg.)"
} else {
    Write-Output "Oppdaterer Mail-utsender: versjon $installertVersjon -> $nyVersjon i $($wb.Name) ..."

    $moduleCode = $moduleCodeTemplate.Replace('{{VERSION}}', $nyVersjon)

    $tempDir = Join-Path $env:TEMP ("GeneriskMail_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    $tempExcel = $null
    $tempWb = $null
    try {
        $tempExcel = New-Object -ComObject Excel.Application
        $tempExcel.Visible = $false
        $tempExcel.DisplayAlerts = $false
        $tempWb = $tempExcel.Workbooks.Add()

        # --- modGenerisk (standard-modul) ---
        $modKomponent = $tempWb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        $modKomponent.Name = "modGenerisk_tmp"
        $modKomponent.CodeModule.AddFromString($moduleCode)
        $modBas = Join-Path $tempDir "modGenerisk.bas"
        $modKomponent.Export($modBas)

        # --- frmSendMail (UserForm) ---
        $frmSend = $tempWb.VBProject.VBComponents.Add(3)  # 3 = vbext_ct_MSForm
        $frmSend.Name = "frmSendMail"
        $frmSend.Properties("Caption").Value = "Mail-utsender"
        $frmSend.Properties("Width").Value = 560
        $frmSend.Properties("Height").Value = 400

        $lbl = $frmSend.Designer.Controls.Add("Forms.Label.1")
        $lbl.Name = "lblTittel"
        $lbl.Caption = "Velg en rad og trykk Send mail"
        $lbl.Left = 10; $lbl.Top = 8; $lbl.Width = 400; $lbl.Height = 16

        $lblInfo = $frmSend.Designer.Controls.Add("Forms.Label.1")
        $lblInfo.Name = "lblInfo"
        $lblInfo.Caption = ""
        $lblInfo.Left = 10; $lblInfo.Top = 28; $lblInfo.Width = 480; $lblInfo.Height = 16

        # Fargede statusfelt (i stedet for én sammenslatt tekstlinje) -
        # gjor det raskere a skanne hvor mange som gjenstar/er sendt/har
        # feil ved et blikk. Farge+fet skrift settes i VBA ved Initialize
        # (IKKE her via PowerShell-COM - .Font-underegenskaper satt fra
        # PowerShell krasjer, se powershell-com-font-bug i minnet).
        $lblStatGjenstar = $frmSend.Designer.Controls.Add("Forms.Label.1")
        $lblStatGjenstar.Name = "lblStatGjenstar"
        $lblStatGjenstar.Caption = ""
        $lblStatGjenstar.Left = 10; $lblStatGjenstar.Top = 28; $lblStatGjenstar.Width = 120; $lblStatGjenstar.Height = 16
        $lblStatGjenstar.Visible = $false

        $lblStatSendt = $frmSend.Designer.Controls.Add("Forms.Label.1")
        $lblStatSendt.Name = "lblStatSendt"
        $lblStatSendt.Caption = ""
        $lblStatSendt.Left = 140; $lblStatSendt.Top = 28; $lblStatSendt.Width = 110; $lblStatSendt.Height = 16
        $lblStatSendt.Visible = $false

        $lblStatFeil = $frmSend.Designer.Controls.Add("Forms.Label.1")
        $lblStatFeil.Name = "lblStatFeil"
        $lblStatFeil.Caption = ""
        $lblStatFeil.Left = 260; $lblStatFeil.Top = 28; $lblStatFeil.Width = 100; $lblStatFeil.Height = 16
        $lblStatFeil.Visible = $false

        $lblStatTotalt = $frmSend.Designer.Controls.Add("Forms.Label.1")
        $lblStatTotalt.Name = "lblStatTotalt"
        $lblStatTotalt.Caption = ""
        $lblStatTotalt.Left = 370; $lblStatTotalt.Top = 28; $lblStatTotalt.Width = 120; $lblStatTotalt.Height = 16
        $lblStatTotalt.Visible = $false

        # Bruker listeboksens EGEN innebygde loddrette rullefelt (dukker
        # automatisk opp nar det er flere rader enn synlig hoyde) i stedet
        # for en separat ScrollBar-kontroll - en frittstaende kontroll ser
        # alltid ut som en losrevet boks ved siden av lista (egen ramme/
        # kant), uansett plassering, siden det rett og slett er en annen
        # type kontroll enn listeboksens egen (innebygde, integrerte)
        # vannrette rullefelt. Fjernet etter gjentatt tilbakemelding om at
        # den "ikke var en del av vinduet".
        $lst = $frmSend.Designer.Controls.Add("Forms.ListBox.1")
        $lst.Name = "lstRader"
        $lst.Left = 10; $lst.Top = 50; $lst.Width = 500; $lst.Height = 224

        $lblVedleggSend = $frmSend.Designer.Controls.Add("Forms.Label.1")
        $lblVedleggSend.Name = "lblVedlegg"
        $lblVedleggSend.Caption = ""
        $lblVedleggSend.ForeColor = 6579300
        $lblVedleggSend.Left = 10; $lblVedleggSend.Top = 276; $lblVedleggSend.Width = 500; $lblVedleggSend.Height = 18

        # Skjules med mindre utsatt levering er satt pa i Oppsett - da vil
        # brukeren fortsatt kunne se det med en gang uten en ekstra
        # bekreftelsesdialog ved hver enkelt-sending (se generisk-mail-
        # verktoy i minnet - eksplisitt onske om a unnga flere klikk her).
        $lblLeveringSend = $frmSend.Designer.Controls.Add("Forms.Label.1")
        $lblLeveringSend.Name = "lblLevering"
        $lblLeveringSend.Caption = ""
        $lblLeveringSend.Left = 10; $lblLeveringSend.Top = 296; $lblLeveringSend.Width = 500; $lblLeveringSend.Height = 18
        $lblLeveringSend.Visible = $false

        $cmdOppsett = $frmSend.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdOppsett.Name = "cmdOppsett"
        $cmdOppsett.Caption = "Oppsett"
        $cmdOppsett.Left = 10; $cmdOppsett.Top = 320; $cmdOppsett.Width = 90; $cmdOppsett.Height = 26

        $cmdSend = $frmSend.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdSend.Name = "cmdSendMail"
        $cmdSend.Caption = "Send mail"
        $cmdSend.BackColor = 5287936
        $cmdSend.ForeColor = 16777215
        $cmdSend.Left = 110; $cmdSend.Top = 320; $cmdSend.Width = 90; $cmdSend.Height = 26

        # Apner frmFeilsjekk (modelost - se den egne kildekode-seksjonen
        # under) - teksten/synligheten oppdateres i VBA ved FyllListe, ikke
        # her via PowerShell-COM.
        $cmdFeilsjekk = $frmSend.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdFeilsjekk.Name = "cmdFeilsjekk"
        $cmdFeilsjekk.Caption = "Undersøk feil"
        $cmdFeilsjekk.Left = 210; $cmdFeilsjekk.Top = 320; $cmdFeilsjekk.Width = 160; $cmdFeilsjekk.Height = 26

        # "Send til flere" - egen, tydelig farge (rod/oransje) siden dette
        # er en langt mer alvorlig handling enn a apne ett enkelt utkast.
        # I fast produksjonsbruk (bekreftet 2026-09-09, 196 ekte mailer).
        $cmdSendAlle = $frmSend.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdSendAlle.Name = "cmdSendTilAlle"
        $cmdSendAlle.Caption = "Send til flere"
        $cmdSendAlle.BackColor = 3160817
        $cmdSendAlle.ForeColor = 16777215
        $cmdSendAlle.Left = 380; $cmdSendAlle.Top = 320; $cmdSendAlle.Width = 100; $cmdSendAlle.Height = 26

        Start-Sleep -Milliseconds 400
        $frmSend.CodeModule.AddFromString($formCodeSendMail)
        $frmSendFrm = Join-Path $tempDir "frmSendMail.frm"
        $frmSend.Export($frmSendFrm)

        # --- frmOppsett (UserForm) ---
        $frmOpp = $tempWb.VBProject.VBComponents.Add(3)
        $frmOpp.Name = "frmOppsett"
        $frmOpp.Properties("Caption").Value = "Oppsett - Mail-utsender"
        $frmOpp.Properties("Width").Value = 525
        $frmOpp.Properties("Height").Value = 560

        $lblT = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblT.Name = "lblTittel"
        $lblT.Caption = "Koble plassholdere og kolonner"
        $lblT.Left = 10; $lblT.Top = 8; $lblT.Width = 400; $lblT.Height = 16

        # --- Ramme 1: Mal-fil (fil, emne, brodtekst-forhandsvisning) ---
        $fraMalfil = $frmOpp.Designer.Controls.Add("Forms.Frame.1")
        $fraMalfil.Name = "fraMalfil"
        $fraMalfil.Caption = "Mal-fil"
        $fraMalfil.Left = 10; $fraMalfil.Top = 30; $fraMalfil.Width = 470; $fraMalfil.Height = 320

        $lblMalLabel = $fraMalfil.Controls.Add("Forms.Label.1")
        $lblMalLabel.Name = "lblMalFilLabel"
        $lblMalLabel.Caption = "Mal-fil:"
        $lblMalLabel.AutoSize = $true
        $lblMalLabel.WordWrap = $false
        $lblMalLabel.Left = 10; $lblMalLabel.Top = 16

        $lblMal = $fraMalfil.Controls.Add("Forms.Label.1")
        $lblMal.Name = "lblMalFil"
        $lblMal.Caption = "(ingen valgt)"
        $lblMal.AutoSize = $true
        $lblMal.WordWrap = $false
        $lblMal.Left = 60; $lblMal.Top = 16

        $cmdVelg = $fraMalfil.Controls.Add("Forms.CommandButton.1")
        $cmdVelg.Name = "cmdVelgMalFil"
        $cmdVelg.Caption = "Velg"
        $cmdVelg.Left = 200; $cmdVelg.Top = 13; $cmdVelg.Width = 55; $cmdVelg.Height = 20

        $lblEmneLabel = $fraMalfil.Controls.Add("Forms.Label.1")
        $lblEmneLabel.Name = "lblEmneLabel"
        $lblEmneLabel.Caption = "Emne i malen:"
        $lblEmneLabel.AutoSize = $true
        $lblEmneLabel.WordWrap = $false
        $lblEmneLabel.Left = 10; $lblEmneLabel.Top = 42

        $txtEmne = $fraMalfil.Controls.Add("Forms.TextBox.1")
        $txtEmne.Name = "txtEmne"
        $txtEmne.Left = 105; $txtEmne.Top = 42; $txtEmne.Width = 350; $txtEmne.Height = 18
        $txtEmne.Locked = $true
        $txtEmne.BackColor = 15921906

        $lblBrodtekstLabel = $fraMalfil.Controls.Add("Forms.Label.1")
        $lblBrodtekstLabel.Name = "lblBrodtekstLabel"
        $lblBrodtekstLabel.Caption = "Brødtekst (som tekst - kun til forhåndsvisning):"
        $lblBrodtekstLabel.AutoSize = $true
        $lblBrodtekstLabel.WordWrap = $false
        $lblBrodtekstLabel.Left = 10; $lblBrodtekstLabel.Top = 66

        $txtBrodtekst = $fraMalfil.Controls.Add("Forms.TextBox.1")
        $txtBrodtekst.Name = "txtBrodtekst"
        $txtBrodtekst.Left = 10; $txtBrodtekst.Top = 84; $txtBrodtekst.Width = 445; $txtBrodtekst.Height = 150
        $txtBrodtekst.MultiLine = $true
        $txtBrodtekst.ScrollBars = 2
        $txtBrodtekst.Locked = $true
        $txtBrodtekst.BackColor = 15921906

        $lblVedleggLabel = $fraMalfil.Controls.Add("Forms.Label.1")
        $lblVedleggLabel.Name = "lblVedleggLabel"
        $lblVedleggLabel.Caption = "Vedlegg som legges ved (filer i samme mappe som mal-fila):"
        $lblVedleggLabel.AutoSize = $true
        $lblVedleggLabel.WordWrap = $false
        $lblVedleggLabel.Left = 10; $lblVedleggLabel.Top = 158

        $txtVedlegg = $fraMalfil.Controls.Add("Forms.TextBox.1")
        $txtVedlegg.Name = "txtVedlegg"
        $txtVedlegg.Left = 10; $txtVedlegg.Top = 176; $txtVedlegg.Width = 445; $txtVedlegg.Height = 40
        $txtVedlegg.MultiLine = $true
        $txtVedlegg.ScrollBars = 2
        $txtVedlegg.Locked = $true
        $txtVedlegg.BackColor = 15921906

        # --- Ramme 2: Plassholdere (hoyde justeres dynamisk i VBA) ---
        $fraPlassholdere = $frmOpp.Designer.Controls.Add("Forms.Frame.1")
        $fraPlassholdere.Name = "fraPlassholdere"
        $fraPlassholdere.Caption = "Plassholdere"
        $fraPlassholdere.Left = 10; $fraPlassholdere.Top = 225; $fraPlassholdere.Width = 470; $fraPlassholdere.Height = 30

        $radTop = 16
        for ($i = 1; $i -le 8; $i++) {
            $lblPH = $fraPlassholdere.Controls.Add("Forms.Label.1")
            $lblPH.Name = "lblPH$i"
            $lblPH.Caption = ""
            $lblPH.Left = 10; $lblPH.Top = $radTop; $lblPH.Width = 60; $lblPH.Height = 18
            $lblPH.Visible = $false

            $cboKol = $fraPlassholdere.Controls.Add("Forms.ComboBox.1")
            $cboKol.Name = "cboKol$i"
            $cboKol.Left = 75; $cboKol.Top = $radTop; $cboKol.Width = 200; $cboKol.Height = 18
            $cboKol.Style = 2  # fmStyleDropDownList
            $cboKol.Visible = $false

            $radTop += 22
        }

        # --- Ramme 3: E-post og status ---
        $fraEpostStatus = $frmOpp.Designer.Controls.Add("Forms.Frame.1")
        $fraEpostStatus.Name = "fraEpostStatus"
        $fraEpostStatus.Caption = "E-post og status"
        $fraEpostStatus.Left = 10; $fraEpostStatus.Top = 265; $fraEpostStatus.Width = 470; $fraEpostStatus.Height = 186

        $lblEpostLabel = $fraEpostStatus.Controls.Add("Forms.Label.1")
        $lblEpostLabel.Name = "lblEpostLabel"
        $lblEpostLabel.Caption = "E-post-kolonne:"
        $lblEpostLabel.Left = 10; $lblEpostLabel.Top = 16; $lblEpostLabel.Width = 100; $lblEpostLabel.Height = 18

        $cboEpost = $fraEpostStatus.Controls.Add("Forms.ComboBox.1")
        $cboEpost.Name = "cboEpost"
        $cboEpost.Left = 115; $cboEpost.Top = 16; $cboEpost.Width = 200; $cboEpost.Height = 18
        $cboEpost.Style = 2

        $lblStatusLabel = $fraEpostStatus.Controls.Add("Forms.Label.1")
        $lblStatusLabel.Name = "lblStatusLabel"
        $lblStatusLabel.Caption = "Status-kolonne:"
        $lblStatusLabel.Left = 10; $lblStatusLabel.Top = 42; $lblStatusLabel.Width = 100; $lblStatusLabel.Height = 18

        $cboStatus = $fraEpostStatus.Controls.Add("Forms.ComboBox.1")
        $cboStatus.Name = "cboStatus"
        $cboStatus.Left = 115; $cboStatus.Top = 42; $cboStatus.Width = 200; $cboStatus.Height = 18
        $cboStatus.Style = 2

        # CC-kolonne - valgfri (kan stå blank = ingen CC), samme mønster som
        # de to over. Under utprøving 2026-09-09 etter Håkons forespørsel.
        $lblCCLabel = $fraEpostStatus.Controls.Add("Forms.Label.1")
        $lblCCLabel.Name = "lblCCLabel"
        $lblCCLabel.Caption = "CC-kolonne (valgfritt):"
        $lblCCLabel.Left = 10; $lblCCLabel.Top = 68; $lblCCLabel.Width = 100; $lblCCLabel.Height = 18

        $cboCC = $fraEpostStatus.Controls.Add("Forms.ComboBox.1")
        $cboCC.Name = "cboCC"
        $cboCC.Left = 115; $cboCC.Top = 68; $cboCC.Width = 200; $cboCC.Height = 18
        $cboCC.Style = 2

        # Retning (normal/omvendt) og tekst for status-kolonnen - lagt til
        # 2026-09-23 etter Håkons ønske om å kunne reversere hvilke rader
        # som vises, og om å slippe unna en hardkodet "Ja".
        $optRetningNormal = $fraEpostStatus.Controls.Add("Forms.OptionButton.1")
        $optRetningNormal.Name = "optRetningNormal"
        $optRetningNormal.GroupName = "StatusRetningGruppe"
        $optRetningNormal.Caption = "Tomme rader vises - teksten skrives inn når mail sendes"
        $optRetningNormal.Left = 10; $optRetningNormal.Top = 98; $optRetningNormal.Width = 450; $optRetningNormal.Height = 18

        $optRetningOmvendt = $fraEpostStatus.Controls.Add("Forms.OptionButton.1")
        $optRetningOmvendt.Name = "optRetningOmvendt"
        $optRetningOmvendt.GroupName = "StatusRetningGruppe"
        $optRetningOmvendt.Caption = "Rader med teksten vises - teksten fjernes når mail sendes"
        $optRetningOmvendt.Left = 10; $optRetningOmvendt.Top = 122; $optRetningOmvendt.Width = 450; $optRetningOmvendt.Height = 18

        $lblStatusTekstLabel = $fraEpostStatus.Controls.Add("Forms.Label.1")
        $lblStatusTekstLabel.Name = "lblStatusTekstLabel"
        $lblStatusTekstLabel.Caption = "Tekst som skrives ved sending:"
        $lblStatusTekstLabel.Left = 10; $lblStatusTekstLabel.Top = 150; $lblStatusTekstLabel.Width = 180; $lblStatusTekstLabel.Height = 18

        $txtStatusTekst = $fraEpostStatus.Controls.Add("Forms.TextBox.1")
        $txtStatusTekst.Name = "txtStatusTekst"
        $txtStatusTekst.Left = 195; $txtStatusTekst.Top = 148; $txtStatusTekst.Width = 120; $txtStatusTekst.Height = 18

        # --- Ramme 3b: Kolonner i hovedlisten - NY 2026-09-23, apner
        # --- frmKolonnevisning (egen liten dialog med avkrysningsliste) ---
        $fraListevisning = $frmOpp.Designer.Controls.Add("Forms.Frame.1")
        $fraListevisning.Name = "fraListevisning"
        $fraListevisning.Caption = "Kolonner i hovedlisten"
        $fraListevisning.Left = 10; $fraListevisning.Top = 451; $fraListevisning.Width = 470; $fraListevisning.Height = 56

        $lblListevisningInfo = $fraListevisning.Controls.Add("Forms.Label.1")
        $lblListevisningInfo.Name = "lblListevisningInfo"
        $lblListevisningInfo.Caption = "Velg hvilke kolonner som skal vises i utvalgslistene."
        $lblListevisningInfo.Left = 10; $lblListevisningInfo.Top = 20; $lblListevisningInfo.Width = 300; $lblListevisningInfo.Height = 18

        $cmdVelgKolonner = $fraListevisning.Controls.Add("Forms.CommandButton.1")
        $cmdVelgKolonner.Name = "cmdVelgKolonner"
        $cmdVelgKolonner.Caption = "Velg kolonner..."
        $cmdVelgKolonner.Left = 320; $cmdVelgKolonner.Top = 16; $cmdVelgKolonner.Width = 130; $cmdVelgKolonner.Height = 26

        # --- Ramme 4: Avsender ("Send fra") - NY, under utprovning ---
        # Alternativ C (2026-09-08): "Automatisk" (dagens oppforsel, ingen
        # overstyring) + nedtrekksmeny med tidligere lagrede adresser +
        # fritekstfelt for a legge til nye. Se generisk-mail-verktoy i
        # minnet for bakgrunn (fritekst alene garanterer ikke sende-
        # rettighet, og delte postbokser vises IKKE i Outlook sine egne
        # kontoer - derfor bade dropdown OG fritekst, ikke bare en av dem).
        $fraAvsender = $frmOpp.Designer.Controls.Add("Forms.Frame.1")
        $fraAvsender.Name = "fraAvsender"
        $fraAvsender.Caption = "Avsender (send fra)"
        $fraAvsender.Left = 10; $fraAvsender.Top = 355; $fraAvsender.Width = 470; $fraAvsender.Height = 72

        $lblAvsenderLabel = $fraAvsender.Controls.Add("Forms.Label.1")
        $lblAvsenderLabel.Name = "lblAvsenderLabel"
        $lblAvsenderLabel.Caption = "Send fra:"
        $lblAvsenderLabel.Left = 10; $lblAvsenderLabel.Top = 16; $lblAvsenderLabel.Width = 70; $lblAvsenderLabel.Height = 18

        $cboAvsender = $fraAvsender.Controls.Add("Forms.ComboBox.1")
        $cboAvsender.Name = "cboAvsender"
        $cboAvsender.Left = 85; $cboAvsender.Top = 16; $cboAvsender.Width = 200; $cboAvsender.Height = 18
        $cboAvsender.Style = 2

        $lblNyAvsenderLabel = $fraAvsender.Controls.Add("Forms.Label.1")
        $lblNyAvsenderLabel.Name = "lblNyAvsenderLabel"
        $lblNyAvsenderLabel.Caption = "Legg til ny:"
        $lblNyAvsenderLabel.Left = 10; $lblNyAvsenderLabel.Top = 42; $lblNyAvsenderLabel.Width = 70; $lblNyAvsenderLabel.Height = 18

        $txtNyAvsender = $fraAvsender.Controls.Add("Forms.TextBox.1")
        $txtNyAvsender.Name = "txtNyAvsender"
        $txtNyAvsender.Left = 85; $txtNyAvsender.Top = 42; $txtNyAvsender.Width = 200; $txtNyAvsender.Height = 18

        $cmdLeggTilAvsender = $fraAvsender.Controls.Add("Forms.CommandButton.1")
        $cmdLeggTilAvsender.Name = "cmdLeggTilAvsender"
        $cmdLeggTilAvsender.Caption = "Legg til"
        $cmdLeggTilAvsender.Left = 295; $cmdLeggTilAvsender.Top = 39; $cmdLeggTilAvsender.Width = 70; $cmdLeggTilAvsender.Height = 20

        # --- Ramme 5: Levering (utsatt sending) - NY, under utprovning ---
        # Bruker Outlooks DeferredDeliveryTime (mailen legges i Utboks til
        # angitt tidspunkt - krever at Outlook er apent/tilkoblet da). Se
        # generisk-mail-verktoy i minnet for bakgrunn/avklaringer rundt dette.
        # v2 (etter tilbakemelding): alle verdifelt star i en felles, rett
        # kolonne (Left=190) i stedet for a henge rett etter hver radioknapp-
        # tekst - ryddigere a skanne. Klokkeslett er nedtrekksmeny
        # (fmStyleDropDownList - umulig a skrive feil format i) i stedet for
        # fritekstfelt. Felt for ikke-valgt alternativ er badte deaktivert
        # OG gratonet (OppdaterLeveringsFelter i VBA), sa "Send med en gang"
        # tydelig viser at resten av rammen ikke er i bruk.
        # v3 (etter tilbakemelding): en enkelt dato-nedtrekksmeny med 61
        # fulle datoer var uoversiktlig a bla i. Erstattet med tre sma,
        # kaskaderende menyer - Dag/Maned/Ar - der Dag-lista bygges pa nytt
        # (OppdaterDagListe) hver gang Maned eller Ar endres, sa den alltid
        # viser riktig antall dager for akkurat den maneden/aret (skuddar
        # handteres automatisk via DateSerial(aar, maned+1, 0)-trikset).
        $fraLevering = $frmOpp.Designer.Controls.Add("Forms.Frame.1")
        $fraLevering.Name = "fraLevering"
        $fraLevering.Caption = "Levering"
        $fraLevering.Left = 10; $fraLevering.Top = 435; $fraLevering.Width = 480; $fraLevering.Height = 108

        $optEngang = $fraLevering.Controls.Add("Forms.OptionButton.1")
        $optEngang.Name = "optLeverEngang"
        $optEngang.Caption = "Send med en gang"
        $optEngang.Left = 10; $optEngang.Top = 16; $optEngang.Width = 170; $optEngang.Height = 18

        $optMinutter = $fraLevering.Controls.Add("Forms.OptionButton.1")
        $optMinutter.Name = "optLeverMinutter"
        $optMinutter.Caption = "Send om"
        $optMinutter.Left = 10; $optMinutter.Top = 44; $optMinutter.Width = 90; $optMinutter.Height = 18

        $txtMinutter = $fraLevering.Controls.Add("Forms.TextBox.1")
        $txtMinutter.Name = "txtLeverMinutter"
        $txtMinutter.Left = 190; $txtMinutter.Top = 42; $txtMinutter.Width = 40; $txtMinutter.Height = 18

        $lblMinutter = $fraLevering.Controls.Add("Forms.Label.1")
        $lblMinutter.Name = "lblLeverMinutterEtikett"
        $lblMinutter.Caption = "minutter"
        $lblMinutter.AutoSize = $true
        $lblMinutter.WordWrap = $false
        $lblMinutter.Left = 235; $lblMinutter.Top = 44

        $optTidspunkt = $fraLevering.Controls.Add("Forms.OptionButton.1")
        $optTidspunkt.Name = "optLeverTidspunkt"
        $optTidspunkt.Caption = "Bestemt tidspunkt"
        $optTidspunkt.Left = 10; $optTidspunkt.Top = 72; $optTidspunkt.Width = 170; $optTidspunkt.Height = 18

        $cboDag = $fraLevering.Controls.Add("Forms.ComboBox.1")
        $cboDag.Name = "cboLeverDag"
        $cboDag.Left = 190; $cboDag.Top = 70; $cboDag.Width = 40; $cboDag.Height = 18
        $cboDag.Style = 2  # fmStyleDropDownList - kan ikke skrive inn ugyldig verdi

        $cboManed = $fraLevering.Controls.Add("Forms.ComboBox.1")
        $cboManed.Name = "cboLeverManed"
        $cboManed.Left = 233; $cboManed.Top = 70; $cboManed.Width = 85; $cboManed.Height = 18
        $cboManed.Style = 2

        $cboAar = $fraLevering.Controls.Add("Forms.ComboBox.1")
        $cboAar.Name = "cboLeverAar"
        $cboAar.Left = 321; $cboAar.Top = 70; $cboAar.Width = 50; $cboAar.Height = 18
        $cboAar.Style = 2

        $lblKl = $fraLevering.Controls.Add("Forms.Label.1")
        $lblKl.Name = "lblLeverKlEtikett"
        $lblKl.Caption = "kl."
        $lblKl.AutoSize = $true
        $lblKl.WordWrap = $false
        $lblKl.Left = 375; $lblKl.Top = 72

        $cboKlokkeslett = $fraLevering.Controls.Add("Forms.ComboBox.1")
        $cboKlokkeslett.Name = "cboLeverKlokkeslett"
        $cboKlokkeslett.Left = 396; $cboKlokkeslett.Top = 70; $cboKlokkeslett.Width = 55; $cboKlokkeslett.Height = 18
        $cboKlokkeslett.Style = 2

        $lblMelding = $frmOpp.Designer.Controls.Add("Forms.Label.1")
        $lblMelding.Name = "lblMelding"
        $lblMelding.Caption = ""
        $lblMelding.Left = 10; $lblMelding.Top = 353; $lblMelding.Width = 460; $lblMelding.Height = 32

        $cmdLagre = $frmOpp.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdLagre.Name = "cmdLagre"
        $cmdLagre.Caption = "Lagre"
        $cmdLagre.BackColor = 5287936
        $cmdLagre.ForeColor = 16777215
        $cmdLagre.Left = 10; $cmdLagre.Top = 393; $cmdLagre.Width = 100; $cmdLagre.Height = 28

        # Tilbake UTEN a lagre - ny knapp per Håkons onske (kun Lagre fantes
        # fra for). Bruker samme "ga tilbake til kalleren"-logikk som Lagre.
        $cmdTilbake = $frmOpp.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdTilbake.Name = "cmdTilbake"
        $cmdTilbake.Caption = "Tilbake"
        $cmdTilbake.Left = 120; $cmdTilbake.Top = 393; $cmdTilbake.Width = 100; $cmdTilbake.Height = 28

        Start-Sleep -Milliseconds 400
        $frmOpp.CodeModule.AddFromString($formCodeOppsett)
        $frmOppFrm = Join-Path $tempDir "frmOppsett.frm"
        $frmOpp.Export($frmOppFrm)

        # --- frmMasseSend (UserForm) - "Send til flere" ---
        $frmMasse = $tempWb.VBProject.VBComponents.Add(3)
        $frmMasse.Name = "frmMasseSend"
        $frmMasse.Properties("Caption").Value = "Send til flere"
        $frmMasse.Properties("Width").Value = 700
        # Ekstra raus margin (2026-09-11) - siste knapperad ble klippet ved
        # forrige, strammere hoyde (572). UserForm.Height ser ut til a
        # bruke mer av totalen til vindusramme/tittellinje enn antatt, sa
        # gir na god margin i stedet for a fortsette a gjette presist.
        $frmMasse.Properties("Height").Value = 610

        $lblMTittel = $frmMasse.Designer.Controls.Add("Forms.Label.1")
        $lblMTittel.Name = "lblTittel"
        $lblMTittel.Caption = "Velg hvilke mailer som skal sendes"
        $lblMTittel.Left = 10; $lblMTittel.Top = 8; $lblMTittel.Width = 500; $lblMTittel.Height = 16

        $lstMailer = $frmMasse.Designer.Controls.Add("Forms.ListBox.1")
        $lstMailer.Name = "lstMailer"
        $lstMailer.Left = 10; $lstMailer.Top = 28; $lstMailer.Width = 670; $lstMailer.Height = 180

        # Hele denne rada (hint, antall valgt, Marker X, Velg/Fjern alle) er
        # samlet her og strammet inn 2026-09-10 sa alt faktisk far plass pa
        # én linje - Marker-feltet stod for langt nede (nede ved Oppsett/
        # Tilbake-knappene) og gjorde den rada bredere enn selve vinduet.
        # Jevne 10px mellomrom mellom hvert element (Håkon 2026-09-11: "kan
        # også være litt mer symmetrisk") - hint, antall valgt, Marker-
        # feltet, og Velg/Fjern alle etter hverandre.
        $lblHint = $frmMasse.Designer.Controls.Add("Forms.Label.1")
        $lblHint.Name = "lblHint"
        $lblHint.Caption = "Huk av i boksen. Shift = flere."
        $lblHint.Left = 10; $lblHint.Top = 214; $lblHint.Width = 240; $lblHint.Height = 16
        $lblHint.ForeColor = 6579300

        $lblAntallValgt = $frmMasse.Designer.Controls.Add("Forms.Label.1")
        $lblAntallValgt.Name = "lblAntallValgt"
        $lblAntallValgt.Caption = "0 valgt"
        $lblAntallValgt.Left = 260; $lblAntallValgt.Top = 214; $lblAntallValgt.Width = 55; $lblAntallValgt.Height = 16

        $lblMarkerXM = $frmMasse.Designer.Controls.Add("Forms.Label.1")
        $lblMarkerXM.Name = "lblMarkerX"
        $lblMarkerXM.Caption = "Marker:"
        $lblMarkerXM.Left = 325; $lblMarkerXM.Top = 214; $lblMarkerXM.Width = 45; $lblMarkerXM.Height = 16

        $txtMarkerXM = $frmMasse.Designer.Controls.Add("Forms.TextBox.1")
        $txtMarkerXM.Name = "txtMarkerX"
        $txtMarkerXM.Left = 370; $txtMarkerXM.Top = 210; $txtMarkerXM.Width = 35; $txtMarkerXM.Height = 20

        $cmdMarkerXM = $frmMasse.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdMarkerXM.Name = "cmdMarkerX"
        $cmdMarkerXM.Caption = "Marker X"
        $cmdMarkerXM.Left = 415; $cmdMarkerXM.Top = 208; $cmdMarkerXM.Width = 70; $cmdMarkerXM.Height = 24

        $cmdVelgAlleM = $frmMasse.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdVelgAlleM.Name = "cmdVelgAlle"
        $cmdVelgAlleM.Caption = "Velg alle"
        $cmdVelgAlleM.Left = 495; $cmdVelgAlleM.Top = 208; $cmdVelgAlleM.Width = 75; $cmdVelgAlleM.Height = 24

        $cmdFjernAlleM = $frmMasse.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdFjernAlleM.Name = "cmdFjernAlle"
        $cmdFjernAlleM.Caption = "Fjern alle"
        $cmdFjernAlleM.Left = 580; $cmdFjernAlleM.Top = 208; $cmdFjernAlleM.Width = 75; $cmdFjernAlleM.Height = 24

        $fraForhandsvisning = $frmMasse.Designer.Controls.Add("Forms.Frame.1")
        $fraForhandsvisning.Name = "fraForhandsvisning"
        $fraForhandsvisning.Caption = "Forhåndsvisning"
        $fraForhandsvisning.Left = 10; $fraForhandsvisning.Top = 238; $fraForhandsvisning.Width = 670; $fraForhandsvisning.Height = 252

        $lblPrevEmneLabel = $fraForhandsvisning.Controls.Add("Forms.Label.1")
        $lblPrevEmneLabel.Name = "lblPrevEmneLabel"
        $lblPrevEmneLabel.Caption = "Emne:"
        $lblPrevEmneLabel.Left = 10; $lblPrevEmneLabel.Top = 16; $lblPrevEmneLabel.Width = 60; $lblPrevEmneLabel.Height = 18

        $txtPrevEmne = $fraForhandsvisning.Controls.Add("Forms.TextBox.1")
        $txtPrevEmne.Name = "txtPrevEmne"
        $txtPrevEmne.Left = 75; $txtPrevEmne.Top = 16; $txtPrevEmne.Width = 580; $txtPrevEmne.Height = 18
        $txtPrevEmne.Locked = $true
        $txtPrevEmne.BackColor = 15921906

        $lblPrevTilLabel = $fraForhandsvisning.Controls.Add("Forms.Label.1")
        $lblPrevTilLabel.Name = "lblPrevTilLabel"
        $lblPrevTilLabel.Caption = "Til:"
        $lblPrevTilLabel.Left = 10; $lblPrevTilLabel.Top = 40; $lblPrevTilLabel.Width = 60; $lblPrevTilLabel.Height = 18

        $txtPrevTil = $fraForhandsvisning.Controls.Add("Forms.TextBox.1")
        $txtPrevTil.Name = "txtPrevTil"
        $txtPrevTil.Left = 75; $txtPrevTil.Top = 40; $txtPrevTil.Width = 300; $txtPrevTil.Height = 18
        $txtPrevTil.Locked = $true
        $txtPrevTil.BackColor = 15921906

        $lblPrevFraLabel = $fraForhandsvisning.Controls.Add("Forms.Label.1")
        $lblPrevFraLabel.Name = "lblPrevFraLabel"
        $lblPrevFraLabel.Caption = "Fra:"
        $lblPrevFraLabel.Left = 385; $lblPrevFraLabel.Top = 40; $lblPrevFraLabel.Width = 40; $lblPrevFraLabel.Height = 18

        $txtPrevFra = $fraForhandsvisning.Controls.Add("Forms.TextBox.1")
        $txtPrevFra.Name = "txtPrevFra"
        $txtPrevFra.Left = 425; $txtPrevFra.Top = 40; $txtPrevFra.Width = 230; $txtPrevFra.Height = 18
        $txtPrevFra.Locked = $true
        $txtPrevFra.BackColor = 15921906

        $lblPrevCCLabel = $fraForhandsvisning.Controls.Add("Forms.Label.1")
        $lblPrevCCLabel.Name = "lblPrevCCLabel"
        $lblPrevCCLabel.Caption = "Cc:"
        $lblPrevCCLabel.Left = 10; $lblPrevCCLabel.Top = 64; $lblPrevCCLabel.Width = 60; $lblPrevCCLabel.Height = 18

        $txtPrevCC = $fraForhandsvisning.Controls.Add("Forms.TextBox.1")
        $txtPrevCC.Name = "txtPrevCC"
        $txtPrevCC.Left = 75; $txtPrevCC.Top = 64; $txtPrevCC.Width = 580; $txtPrevCC.Height = 18
        $txtPrevCC.Locked = $true
        $txtPrevCC.BackColor = 15921906

        $lblPrevBrodtekstLabel = $fraForhandsvisning.Controls.Add("Forms.Label.1")
        $lblPrevBrodtekstLabel.Name = "lblPrevBrodtekstLabel"
        $lblPrevBrodtekstLabel.Caption = "Brødtekst:"
        $lblPrevBrodtekstLabel.Left = 10; $lblPrevBrodtekstLabel.Top = 88; $lblPrevBrodtekstLabel.Width = 100; $lblPrevBrodtekstLabel.Height = 16

        $txtPrevBrodtekst = $fraForhandsvisning.Controls.Add("Forms.TextBox.1")
        $txtPrevBrodtekst.Name = "txtPrevBrodtekst"
        $txtPrevBrodtekst.Left = 10; $txtPrevBrodtekst.Top = 106; $txtPrevBrodtekst.Width = 645; $txtPrevBrodtekst.Height = 110
        $txtPrevBrodtekst.MultiLine = $true
        $txtPrevBrodtekst.ScrollBars = 2
        $txtPrevBrodtekst.Locked = $true
        $txtPrevBrodtekst.BackColor = 15921906

        $lblPrevVedleggLabel = $fraForhandsvisning.Controls.Add("Forms.Label.1")
        $lblPrevVedleggLabel.Name = "lblPrevVedleggLabel"
        $lblPrevVedleggLabel.Caption = "Vedlegg:"
        $lblPrevVedleggLabel.Left = 10; $lblPrevVedleggLabel.Top = 224; $lblPrevVedleggLabel.Width = 60; $lblPrevVedleggLabel.Height = 18

        $txtPrevVedlegg = $fraForhandsvisning.Controls.Add("Forms.TextBox.1")
        $txtPrevVedlegg.Name = "txtPrevVedlegg"
        $txtPrevVedlegg.Left = 75; $txtPrevVedlegg.Top = 224; $txtPrevVedlegg.Width = 580; $txtPrevVedlegg.Height = 18
        $txtPrevVedlegg.Locked = $true
        $txtPrevVedlegg.BackColor = 15921906

        $lblStatusM = $frmMasse.Designer.Controls.Add("Forms.Label.1")
        $lblStatusM.Name = "lblStatus"
        $lblStatusM.Caption = ""
        $lblStatusM.Left = 10; $lblStatusM.Top = 500; $lblStatusM.Width = 500; $lblStatusM.Height = 16

        $cmdOppsettM = $frmMasse.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdOppsettM.Name = "cmdOppsett"
        $cmdOppsettM.Caption = "Oppsett"
        $cmdOppsettM.Left = 10; $cmdOppsettM.Top = 522; $cmdOppsettM.Width = 90; $cmdOppsettM.Height = 30

        $cmdLukkM = $frmMasse.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdLukkM.Name = "cmdLukk"
        $cmdLukkM.Caption = "Tilbake"
        $cmdLukkM.Left = 110; $cmdLukkM.Top = 522; $cmdLukkM.Width = 90; $cmdLukkM.Height = 30

        $cmdFeilsjekkM = $frmMasse.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdFeilsjekkM.Name = "cmdFeilsjekk"
        $cmdFeilsjekkM.Caption = "Undersøk feil"
        $cmdFeilsjekkM.Left = 220; $cmdFeilsjekkM.Top = 522; $cmdFeilsjekkM.Width = 150; $cmdFeilsjekkM.Height = 30

        $cmdSendAlleM = $frmMasse.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdSendAlleM.Name = "cmdSendTilAlle"
        $cmdSendAlleM.Caption = "Send til flere"
        $cmdSendAlleM.Left = 570; $cmdSendAlleM.Top = 522; $cmdSendAlleM.Width = 110; $cmdSendAlleM.Height = 30

        Start-Sleep -Milliseconds 400
        $frmMasse.CodeModule.AddFromString($formCodeMasseSend)
        $frmMasseFrm = Join-Path $tempDir "frmMasseSend.frm"
        $frmMasse.Export($frmMasseFrm)

        # --- frmFeilsjekk (UserForm) - "Undersøk feil", under utprovning ---
        $frmFeil = $tempWb.VBProject.VBComponents.Add(3)
        $frmFeil.Name = "frmFeilsjekk"
        $frmFeil.Properties("Caption").Value = "Undersøk feil"
        $frmFeil.Properties("Width").Value = 450
        $frmFeil.Properties("Height").Value = 360

        $lblFeilTittel = $frmFeil.Designer.Controls.Add("Forms.Label.1")
        $lblFeilTittel.Name = "lblTittel"
        $lblFeilTittel.Caption = ""
        $lblFeilTittel.Left = 10; $lblFeilTittel.Top = 8; $lblFeilTittel.Width = 420; $lblFeilTittel.Height = 32
        $lblFeilTittel.WordWrap = $true

        $lstFeil = $frmFeil.Designer.Controls.Add("Forms.ListBox.1")
        $lstFeil.Name = "lstFeil"
        $lstFeil.Left = 10; $lstFeil.Top = 44; $lstFeil.Width = 420; $lstFeil.Height = 190

        $lblFeilStatus = $frmFeil.Designer.Controls.Add("Forms.Label.1")
        $lblFeilStatus.Name = "lblStatus"
        $lblFeilStatus.Caption = ""
        $lblFeilStatus.Left = 10; $lblFeilStatus.Top = 250; $lblFeilStatus.Width = 150; $lblFeilStatus.Height = 20

        $cmdForrigeFeil = $frmFeil.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdForrigeFeil.Name = "cmdForrige"
        $cmdForrigeFeil.Caption = "< Forrige feil"
        $cmdForrigeFeil.Left = 10; $cmdForrigeFeil.Top = 282; $cmdForrigeFeil.Width = 100; $cmdForrigeFeil.Height = 26

        $cmdGaTil = $frmFeil.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdGaTil.Name = "cmdGaTil"
        $cmdGaTil.Caption = "Gå til celle"
        $cmdGaTil.Left = 120; $cmdGaTil.Top = 282; $cmdGaTil.Width = 100; $cmdGaTil.Height = 26

        $cmdNesteFeil = $frmFeil.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdNesteFeil.Name = "cmdNeste"
        $cmdNesteFeil.Caption = "Neste feil >"
        $cmdNesteFeil.Left = 230; $cmdNesteFeil.Top = 282; $cmdNesteFeil.Width = 100; $cmdNesteFeil.Height = 26

        $cmdLukkFeil = $frmFeil.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdLukkFeil.Name = "cmdLukk"
        $cmdLukkFeil.Caption = "Tilbake"
        $cmdLukkFeil.Left = 350; $cmdLukkFeil.Top = 282; $cmdLukkFeil.Width = 80; $cmdLukkFeil.Height = 26

        Start-Sleep -Milliseconds 400
        $frmFeil.CodeModule.AddFromString($formCodeFeilsjekk)
        $frmFeilFrm = Join-Path $tempDir "frmFeilsjekk.frm"
        $frmFeil.Export($frmFeilFrm)

        # --- frmKolonnevisning (UserForm) - "Velg kolonner", NY 2026-09-23 ---
        $frmKolVis = $tempWb.VBProject.VBComponents.Add(3)
        $frmKolVis.Name = "frmKolonnevisning"
        $frmKolVis.Properties("Caption").Value = "Velg kolonner"
        $frmKolVis.Properties("Width").Value = 320
        $frmKolVis.Properties("Height").Value = 448

        $lblKolVisTittel = $frmKolVis.Designer.Controls.Add("Forms.Label.1")
        $lblKolVisTittel.Name = "lblTittel"
        $lblKolVisTittel.Caption = "Velg kolonner som skal vises i listen:"
        $lblKolVisTittel.Left = 10; $lblKolVisTittel.Top = 8; $lblKolVisTittel.Width = 280; $lblKolVisTittel.Height = 18

        $lblKolVisMelding = $frmKolVis.Designer.Controls.Add("Forms.Label.1")
        $lblKolVisMelding.Name = "lblMelding"
        $lblKolVisMelding.Caption = ""
        $lblKolVisMelding.ForeColor = 255  # rod (vbRed som Long)
        $lblKolVisMelding.Left = 10; $lblKolVisMelding.Top = 28; $lblKolVisMelding.Width = 280; $lblKolVisMelding.Height = 32
        $lblKolVisMelding.WordWrap = $true

        $cmdKolVisMerkAlle = $frmKolVis.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdKolVisMerkAlle.Name = "cmdMerkAlle"
        $cmdKolVisMerkAlle.Caption = "Merk alle"
        $cmdKolVisMerkAlle.Left = 10; $cmdKolVisMerkAlle.Top = 44; $cmdKolVisMerkAlle.Width = 135; $cmdKolVisMerkAlle.Height = 22

        $cmdKolVisAvmerkAlle = $frmKolVis.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdKolVisAvmerkAlle.Name = "cmdAvmerkAlle"
        $cmdKolVisAvmerkAlle.Caption = "Avmerk alle"
        $cmdKolVisAvmerkAlle.Left = 155; $cmdKolVisAvmerkAlle.Top = 44; $cmdKolVisAvmerkAlle.Width = 135; $cmdKolVisAvmerkAlle.Height = 22

        $lstKolVisKolonner = $frmKolVis.Designer.Controls.Add("Forms.ListBox.1")
        $lstKolVisKolonner.Name = "lstKolonner"
        $lstKolVisKolonner.Left = 10; $lstKolVisKolonner.Top = 72; $lstKolVisKolonner.Width = 280; $lstKolVisKolonner.Height = 300

        $cmdKolVisLagre = $frmKolVis.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdKolVisLagre.Name = "cmdLagre"
        $cmdKolVisLagre.Caption = "Lagre"
        $cmdKolVisLagre.Left = 120; $cmdKolVisLagre.Top = 382; $cmdKolVisLagre.Width = 80; $cmdKolVisLagre.Height = 26

        $cmdKolVisAvbryt = $frmKolVis.Designer.Controls.Add("Forms.CommandButton.1")
        $cmdKolVisAvbryt.Name = "cmdAvbryt"
        $cmdKolVisAvbryt.Caption = "Avbryt"
        $cmdKolVisAvbryt.Left = 210; $cmdKolVisAvbryt.Top = 382; $cmdKolVisAvbryt.Width = 80; $cmdKolVisAvbryt.Height = 26

        Start-Sleep -Milliseconds 400
        $frmKolVis.CodeModule.AddFromString($formCodeKolonnevisning)
        $frmKolVisFrm = Join-Path $tempDir "frmKolonnevisning.frm"
        $frmKolVis.Export($frmKolVisFrm)

    } finally {
        if ($tempWb) { $tempWb.Close($false) }
        if ($tempExcel) {
            $tempExcel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($tempExcel) | Out-Null
        }
    }

    # --- Importer/oppdater i malfila ---
    foreach ($navn in @("modGenerisk", "frmSendMail", "frmOppsett", "frmMasseSend", "frmFeilsjekk", "frmKolonnevisning")) {
        try {
            $eksisterende = $vbProject.VBComponents.Item($navn)
            $vbProject.VBComponents.Remove($eksisterende)
        } catch {}
    }
    $vbProject.VBComponents.Import($modBas) | Out-Null
    $vbProject.VBComponents.Import($frmSendFrm) | Out-Null
    $vbProject.VBComponents.Import($frmMasseFrm) | Out-Null
    $vbProject.VBComponents.Import($frmOppFrm) | Out-Null
    $vbProject.VBComponents.Import($frmFeilFrm) | Out-Null
    $vbProject.VBComponents.Import($frmKolVisFrm) | Out-Null

    # Gi Excel litt tid til a fullfore registreringen av de nettopp
    # importerte VBA-komponentene for vi setter en Button sin OnAction
    # (som ma kunne resolve makronavnet) - uten denne pausen kan
    # tilordningen feile med en COM-feil ("Kan ikke angi OnAction-
    # egenskapen") rett etter en fersk import.
    Start-Sleep -Milliseconds 500

    # --- Sa inn MAL_UNDERMAPPE/MAL_FILNAVN i den skjulte MailKonfig-fana,
    #     men bare hvis de ikke allerede er satt (ikke overskriv en
    #     kollega sine Oppsett-valg ved en vanlig oppdatering) ---
    $konfigFane = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq 'MailKonfig') { $konfigFane = $s; break } }
    if (-not $konfigFane) {
        $konfigFane = $wb.Worksheets.Add()
        $konfigFane.Name = 'MailKonfig'
        $konfigFane.Cells.Item(1, 1).Value2 = 'Nokkel'
        $konfigFane.Cells.Item(1, 2).Value2 = 'Verdi'
        $konfigFane.Visible = 2  # xlSheetVeryHidden
    }

    function Finn-KonfigRad {
        param($ws, $nokkel)
        $rad = 2
        while ($ws.Cells.Item($rad, 1).Value2) {
            if ($ws.Cells.Item($rad, 1).Value2 -eq $nokkel) { return $rad }
            $rad++
        }
        return $null
    }
    function Sett-KonfigVerdiHvisTom {
        param($ws, $nokkel, $verdi)
        if ([string]::IsNullOrEmpty($verdi)) { return }
        $eksisterendeRad = Finn-KonfigRad $ws $nokkel
        if ($eksisterendeRad) { return }  # ikke overskriv
        $rad = 2
        while ($ws.Cells.Item($rad, 1).Value2) { $rad++ }
        $ws.Cells.Item($rad, 1).Value2 = $nokkel
        $ws.Cells.Item($rad, 2).Value2 = $verdi
    }
    Sett-KonfigVerdiHvisTom $konfigFane 'MAL_UNDERMAPPE' $MalUndermappe
    Sett-KonfigVerdiHvisTom $konfigFane 'MAL_FILNAVN' $MalFilnavn

    # Ingen automatisk ark-knapp lenger (Håkons eksplisitte onske
    # 2026-09-22: en makro legger seg bare til - en knapp/inngang i arket
    # er et eget, manuelt valg brukeren gjor selv, typisk via Makromeny).
    # Rydder fortsatt bort en eventuell gammel "VisOppsett"-knapp fra en
    # tidligere versjon (fjerning av en dod knapp, ikke automatisk
    # tillegging av en ny - i strid med ingenting av det Håkon ba om).
    $wb.Activate()
    $aktivtArk = $wb.ActiveSheet
    foreach ($btn in @($aktivtArk.Buttons())) {
        if ($btn.OnAction -eq "VisOppsett" -or $btn.OnAction -like "*!VisOppsett") { $btn.Delete() }
    }

    Write-Output "  modGenerisk installert"
    Write-Output "  frmSendMail installert"
    Write-Output "  frmOppsett installert"
    Write-Output "  frmMasseSend installert"
    Write-Output "  frmFeilsjekk installert"
    Write-Output "  frmKolonnevisning installert"
    if ($MalUndermappe -or $MalFilnavn) {
        Write-Output "  (Mal-fil satt til: '$MalUndermappe\$MalFilnavn' - kun hvis dette IKKE var satt fra før)"
    }

    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ""
Write-Output "Ferdig. Husk aa lagre filen selv (Ctrl+S) naar du er klar."
Write-Output "Trykk 'Mail-utsender' pa arket - forste gang apnes 'Oppsett' automatisk for aa koble plassholdere og kolonner."
