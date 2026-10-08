param(
    [string]$Path,
    [switch]$Uninstall,
    # Tvinger reinstallasjon selv om versjonen er lik (brukes av EMI "Reinstaller alle").
    [switch]$Force,
    # Kun for testing: start en EGEN Excel-prosess i stedet for a koble til den
    # som allerede kjorer (ellers apnes testfila inne i brukerens egen Excel).
    # Samme monster som Kontaktsentralen.ps1.
    [switch]$NyExcelInstans
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# Makromeny -- alt-i-en installer/oppdaterer.
# Installerer en knapp + en meny (UserForm) formet som en "app-meny" -
# fargede, klikkbare fliser, en per allerede installert TIMSS-makro
# (Kolonnevelger, Mail-utsender, Kontaktsentralen, Nettskjema-henter, ...).
# Ett klikk kjorer makroen og lukker menyen med det samme.
#
# Deteksjon skjer ved a skanne Buttons()/OnAction pa alle ark og sla opp
# kjente OnAction-navn i en liten innebygd liste (HentKjentMakroInfo i
# modMakroMeny) - ikke en manuelt utfylt tabell. Nytt verktoy som skal
# vises i menyen ma legges til i den lista (med farge) her, OG i den
# tilsvarende fargepaletten i Excel Macro Installer sin egen kildekode
# (Hent-IkonBakgrunn / $script:IkonBakgrunnPerMakro), slik at fargene
# stemmer overens mellom installeren og selve menyen.
#
# VIKTIG OM NORSKE TEGN: Windows PowerShell 5.1 (powershell.exe, IKKE
# PowerShell 7/pwsh) leser en .ps1-fil uten BOM med systemets ANSI-
# kodeside, ikke UTF-8. Skrives ae/o-med-strek/aa-med-ring direkte i denne
# kildefila, blir de feiltolket og korrupte for de i det hele tatt naar
# VBA-koden kjorer (samme grunnfeil som er kjent fra tidligere med
# .bas/.frm-eksport, men denne gangen er det selve .ps1-fila som rammes).
# Losning brukt her: alle norske spesialtegn i UI-tekst bygges fra Unicode-
# tegnkoder ($oe/$aa/$ae under) og settes inn i kildekoden med -replace
# rett foer den sendes til AddFromString - aldri skrevet direkte i denne
# fila.
#
# Ingen andre filer trengs eller etterlates.
# ============================================================================

$oe = [char]0x00F8
$aa = [char]0x00E5
$ae = [char]0x00E6

# ---- Innebygd VBA-kildekode ----

$moduleCodeRaw = @'
Option Explicit

Public Const MAKROMENY_VERSION As String = "2.7.0"
' Navnet paa en evt. gammel "baandknapp"-verktoylinje fra en tidligere
' versjon (2.4.0-2.4.2, siden fjernet igjen - Haakons eksplisitte onske
' 2026-09-23: han provde funksjonen, likte ikke resultatet). Konstanten
' brukes KUN til opprydding (se -Uninstall/oppdatering lenger ned) - det
' finnes ingen kode lenger som LEGGER TIL en slik knapp.
Private Const BAANDKNAPP_NAVN As String = "Makromeny"

Public Sub VisMakroMeny()
    On Error GoTo Feil
    frmMakroMeny.Show
    Exit Sub
Feil:
    MsgBox "Kunne ikke {AA}pne makromenyen." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Makromeny"
End Sub

' Offentlig inngang til info-vinduet, i tillegg til "Installert info..."-
' knappen i selve menyen - gjor det mulig a apne det direkte (bl.a. for
' testing) uten a maatte klikke seg gjennom frmMakroMeny forst.
Public Sub VisMakroInfo()
    On Error GoTo Feil
    frmMakroInfo.Show
    Exit Sub
Feil:
    MsgBox "Kunne ikke {AA}pne info-vinduet." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Makromeny"
End Sub

' Kjorer en makro via navnet paa OnAction-en den er koblet til, og setter
' Excel tilbake til normal tilstand etterpaa selv om makroen feiler.
Public Sub KjorValgtMakro(ByVal Makronavn As String)
    Dim ForrigeScreenUpdating As Boolean
    Dim ForrigeEnableEvents As Boolean
    Dim ForrigeDisplayAlerts As Boolean

    If Len(Trim$(Makronavn)) = 0 Then Exit Sub

    ForrigeScreenUpdating = Application.ScreenUpdating
    ForrigeEnableEvents = Application.EnableEvents
    ForrigeDisplayAlerts = Application.DisplayAlerts

    On Error GoTo Feil

    Application.Run Makronavn

    GjenopprettTilstand ForrigeScreenUpdating, ForrigeEnableEvents, ForrigeDisplayAlerts
    Exit Sub

Feil:
    GjenopprettTilstand ForrigeScreenUpdating, ForrigeEnableEvents, ForrigeDisplayAlerts
    MsgBox "Det oppstod en feil under kj{OE}ring av:" & vbCrLf & Makronavn & vbCrLf & vbCrLf & _
           "Feil " & Err.Number & ":" & vbCrLf & Err.Description, vbCritical, "Makrofeil"
End Sub

' NB: setter bevisst IKKE Application.Cursor = xlWait rundt Application.Run -
' den kjorte makroen kan selv Aapne et eget, langvarig vindu (f.eks.
' Kolonnevelger sin egen dialog), og da hadde ventemarkoren stAatt urort helt
' til BRUKEREN lukker det vinduet, uten at noe faktisk prosesserer i
' mellomtiden - misvisende. Den kjorte makroen styrer sin egen markor om den
' trenger det.
Private Sub GjenopprettTilstand(ByVal SU As Boolean, ByVal EE As Boolean, ByVal DA As Boolean)
    Application.ScreenUpdating = SU
    Application.EnableEvents = EE
    Application.DisplayAlerts = DA
End Sub

' Kjente makroer denne menyen kan finne og kjore. Deteksjon skjer via
' KOMPONENTNAVNET paa makroens egen hovedmodul i VBA-prosjektet - ikke via
' en knapp i arket - saa menyen fungerer helt uavhengig av om det finnes
' noen knapp, eller om brukeren har slettet den (f.eks. fordi "vis som
' fane" gjor knappen overfloedig). Krever at "Trust access to the VBA
' project object model" er aktivert (samme innstilling som allerede
' kreves for at installer-scriptene i det hele tatt skal kunne installere
' noe, saa den boer alt vaere paa hvis Makromeny er installert).
'
' Kall med stigende Indeks fra 1 til funksjonen returnerer False - legg
' til en ny Case med neste ledige indeks naar et nytt verktoy skal kunne
' vises i Makromeny, sammen med samme fargekode som brukes for makroen i
' Excel Macro Installer.
Public Function HentKjentMakroInfo(ByVal Indeks As Long, ByRef Komponentnavn As String, ByRef OnActionNavn As String, ByRef Visningsnavn As String, ByRef Beskrivelse As String, ByRef Farge As Long) As Boolean
    Select Case Indeks
        Case 1
            Komponentnavn = "modColumnPicker"
            OnActionNavn = "ShowColumnPicker"
            Visningsnavn = "Kolonnevelger"
            Beskrivelse = "Vis/skjul kolonner i tabellen, pluss lagrede loadouts."
            Farge = RGB(237, 233, 254)
        Case 2
            Komponentnavn = "modGenerisk"
            OnActionNavn = "MailUtsendelse"
            Visningsnavn = "Mail-utsender"
            Beskrivelse = "Send e-post via Outlook-mal til valgte rader i tabellen."
            Farge = RGB(219, 234, 254)
        Case 3
            Komponentnavn = "modKontaktsentralen"
            OnActionNavn = "VisKontaktsentralen"
            Visningsnavn = "Kontaktsentralen"
            Beskrivelse = "Spor svar, avvisninger og automatiske svar fra en delt postboks."
            Farge = RGB(220, 252, 231)
        Case 4
            Komponentnavn = "modNettskjema"
            OnActionNavn = "AapneNettskjemaHenter"
            Visningsnavn = "Nettskjema-henter"
            Beskrivelse = "Hent svar fra Nettskjema (UiO) inn i egne ark/tabeller."
            Farge = RGB(254, 243, 199)
        Case 5
            Komponentnavn = "modStatistikkern"
            OnActionNavn = "AapneStatistikkern"
            Visningsnavn = "Statistikkern"
            Beskrivelse = "Statistikk over unike verdier i valgte kolonner, vist som fargede fliser."
            Farge = RGB(254, 226, 226)
        Case Else
            HentKjentMakroInfo = False
            Exit Function
    End Select
    HentKjentMakroInfo = True
End Function

' Sjekker om en gitt VBA-komponent (modul/skjema) finnes i denne
' arbeidsbokens VBA-prosjekt. Stille False (ikke feil) hvis "Trust access
' to the VBA project object model" ikke skulle vaere aktivert, siden
' ByggMeny sjekker tilgangen samlet paa forhaand og gir en tydelig
' statusmelding om det.
Public Function VBAKomponentFinnes(ByVal Navn As String) As Boolean
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(Navn)
    On Error GoTo 0
    VBAKomponentFinnes = Not comp Is Nothing
End Function

' True hvis "Trust access to the VBA project object model" er aktivert -
' brukes til aa gi en tydelig samlet statusmelding i stedet for at
' deteksjonen bare stille finner null makroer uten forklaring.
Public Function HarVBAProjektTilgang() As Boolean
    Dim n As Long
    On Error Resume Next
    n = ThisWorkbook.VBProject.VBComponents.Count
    HarVBAProjektTilgang = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function Base64ToBytes(ByVal B64 As String) As Byte()
    Dim oXML As Object
    Dim oNode As Object
    Set oXML = CreateObject("MSXML2.DOMDocument")
    Set oNode = oXML.createElement("b64")
    oNode.DataType = "bin.base64"
    oNode.Text = B64
    Base64ToBytes = oNode.nodeTypedValue
End Function

Private Function LastIkonFraBase64(ByVal B64 As String) As IPictureDisp
    Dim bytes() As Byte
    Dim sti As String
    Dim iFile As Integer

    If Len(B64) = 0 Then Exit Function

    On Error GoTo Feil

    bytes = Base64ToBytes(B64)

    sti = Environ$("TEMP") & "\MakromenyIkon_" & Format$(Timer * 1000, "0") & CStr(Int(Rnd * 10000)) & ".bmp"
    iFile = FreeFile
    Open sti For Binary Access Write As #iFile
    Put #iFile, 1, bytes
    Close #iFile

    Set LastIkonFraBase64 = LoadPicture(sti)

    On Error Resume Next
    Kill sti
    On Error GoTo 0
    Exit Function

Feil:
    Set LastIkonFraBase64 = Nothing
End Function

Public Function HentMakroIkon(ByVal OnActionNavn As String) As IPictureDisp
    Dim b64 As String
    Select Case OnActionNavn
        Case "ShowColumnPicker": b64 = IconB64_Kolonnevelger()
        Case "MailUtsendelse": b64 = IconB64_MailUtsender()
        Case "VisKontaktsentralen": b64 = IconB64_Kontaktsentralen()
        Case "AapneNettskjemaHenter": b64 = IconB64_NettskjemaHenter()
        Case "AapneStatistikkern": b64 = IconB64_Statistikkern()
        Case Else: b64 = ""
    End Select
    If Len(b64) > 0 Then
        Set HentMakroIkon = LastIkonFraBase64(b64)
    End If
End Function

Private Function IconB64_Kolonnevelger() As String
    Dim s As String
    s = s & "Qk02GQAAAAAAADYAAAAoAAAAKAAAACgAAAABACAAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7GSRHPdhliLoYY4i6FqOIuhajiLoWo4i6FKOIuhShyLoUoci4VKHIuhShyLoS4ci4UuHIuFLfyLhS38i"
    s = s & "4UN/IuFDfyLhQ38i4UN/IuE8eCLhPHgi90OHItI8eBEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1FV/DPJsmmPvZ5ez7maUxehikfTpYZD85l6O/OZcjfzlWoz85VmK/ORXifzkVIj841KH/ONQhvziT4X84k2E/OFL"
    s = s & "g/zhSYL84UeB/OBGf/zgRH/830N9/N9BfPzfP3v83j56/OI9e/zdOnjm5Tt6xeY6e53gNnZLAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5WeSO/RwnNz+cqL/+W6e//Zrmv/tZpT/62OT/+xikv/rYJH/6l6Q/+pcj//pWo3/6ViM/+hXi//nVYr/"
    s = s & "51OJ/+ZRiP/mT4f/5k6G/+VLhP/lSoP/5EiC/+RGgf/kRYD/40N+/+NBff/iP3z/5j9+/+o+f//wPoH/7Dt+/+E3d7fMM2YPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA6XKZOvNynvH6dKL/6muW/+pplf/pZ5T/6WWT/+lkkv/oYpH/6GCQ/+dej//nXI7/5luN/+ZZ"
    s = s & "i//lV4r/5FWJ/+RTiP/kUof/41CG/+NOhf/iTIT/4kqD/+JJgv/hR4D/4UV//+BDfv/gQn3/4EB8/98/e//fPXr/3jt4/986eP/zPYL/4jd4xd8vbxAAAAAAAAAAAAAAAAAAAAAA7m6ZHvB0nfH4dqL/6m6Y/+ttmP/qa5b/6mmV/+lnlP/pZZP/6WSS/+hikf/oYJD/"
    s = s & "516P/+dcjv/mW43/5lmL/+VXiv/kVYn/5FOI/+RSh//jUIb/406F/+JMhP/iSoP/4kmC/+FHgP/hRX//4EN+/+BCff/gQHz/3z97/989ev/eO3n/3Tl3//M9gv/bNnS6AAAAAAAAAAAAAAAAAAAAAO10nJH/f6v/63Ga/+twmv/rbpn/622X/+prlv/qaZX/6WeU/+ll"
    s = s & "k//pZJL/6GKR/+hgkP/nXo//51yO/+ZbjP/mWYv/5VeK/+RVif/kU4j/5FKH/+NQhv/jToX/4kyE/+NJhP/jRoP/4UaB/+FFf//gQ37/4EJ9/+BAfP/fPnv/3z16/947ef/fOnn/7Dt+/9o3dkUAAAAAAAAAAO9/nxDtd53l9nqj/+xznP/scpv/63Ca/+tumf/qbJf/"
    s = s & "6muW/+pplf/pZ5T/6WWT/+hjkf/oYpH/6GCQ/+dej//nXI7/5lqM/+ZZi//lV4r/5FWJ/+RTiP/jUob/5U2J/+ZHiv/fToD/4FB//+RChf/jQIT/4UV//+BEfv/gQn3/30B7/98+e//fPXr/3jt5/+8+gf/cN3akAAAAAAAAAADue6Qf7Hie//B4oP/tdZ3/7HSc/+xy"
    s = s & "m//rcJr/626Z/+ttl//qa5b/6mmV/+lnlP/pZZP/6WSS/+hikf/oYJD/516P/+dcjv/mW4z/5lmL/+VXiv/kVoj/6E+N/9xYff+8jEz/tbQ8/7u9Q//GqFL/2mV0/+M+hP/hRX//4ER+/+BCff/gQHz/3z97/989ev/rP4D/3Tl3xwAAAAAAAAAA8H2jNfJ9o/zweqD/"
    s = s & "7Xee/+11nf/sdJz/7HKb/+twmv/rbpn/622Y/+prlv/qaZX/6WeU/+llk//pZJL/6GKR/+hgkP/nXo//51yO/+Zbjf/mWYv/51SO/95bgf+opTL/qc8o/7TPO/+wyzP/sNM2/7jOPf/Ya3H/4z+E/+FFf//gQ37/4EJ9/+BAfP/fP3v/5j9+/907eOblP38UAAAAAO57"
    s = s & "n2v/h67/7nqg/+14n//td57/7XWd/+x0nP/scpv/622Y/+1nlP/rZJL/6WGQ/+lfjv/oXY3/6FuM/+hai//pWoz/6ViL/+hWiv/oVIn/51GH/+tRkf+5iUv/ocYg/73OUf/p7Nn/zdmH/67EKv+x0Tf/v7dJ/+FJgf/iRYH/4UV//+BDfv/gQn3/4EB8/+E/fP/gPXr/"
    s = s & "3D51JQAAAADwfqNn/4iu/+98of/ueqD/7Xif/+13nv/tdZ3/7HGa/+qCpf/kl7H/65q1//Ceuf/unLj/7pq3/+6at//rlrP/4oyq/+KLqf/iiqn/4omo/+OIqv/kZJD/pZ8t/63JMv/n69X/8fDv//Tz+//L2IL/sMcy/7bOPP/bZHb/40KF/+FHgP/hRX//4EN+/+BC"
    s = s & "ff/jQX3/3j56/OE8eCIAAAAA8ICjZ/+KsP/vfqL/73yh/+56oP/teJ//7Xee/+5tmP/lobj/0eHc/+Xu6//1/fr/8/v4//P7+P/1/fr/7PTx/9La1//T29j/09vY/9Pc2f/W397/4Hqc/6WcKv+uyTP/xtRw/7fKRv/Q25H/9vP8/77PWf+0zTP/2ml1/+RDhv/hSIH/"
    s = s & "4UeA/+FFf//gQ37/40N+/98/e/zhPHgiAAAAAPCApWf/jLH/8H+j/+9+ov/vfKH/7nqg/+55n//ucJn/5p+3/8/Y1f/i5OT/8fLy/+/w8P/v8PD/8vLy/+jp6f/Q0dH/0dLS/9HS0v/R0tL/0tzZ/+WKsv+0hz3/psgr/7PINf+yxjX/ssc2/8HQY/+zzz3/vL9D/+FT"
    s = s & "gf/jSYb/4kqD/+JJgf/hR4D/4UV//+RFgP/fQXz84UN/IgAAAADwhaVn/46y//CBpP/wf6P/736i/+98of/ueqD/73Ga/+agt//Q2NX/4+Tk//Ly8v/w8PD/8PDw//Ly8v/p6en/0dHR/9LS0v/S0tL/0tLS/9HY1f/evM7/22d6/6OuJv+u0DT/tMs8/7LJOv+uzC//"
    s = s & "stQ1/9N8Zv/mR4v/406F/+JMhP/iSoP/4kmC/+FHgP/kRoH/30N9/OFDfyIAAAAA8IWlZ/+Qs//xg6X/8IGk//B/o//vfqL/73yh/+9zm//mobj/z9fU/+jp6f/6+vr/+Pj4//f39//z8/P/6enp/9DQ0P/V1dX/2dnZ/9jY2P/Y2Nj/093Z/+Shwf/XZnL/sZo3/7LC"
    s = s & "Nv+3yTr/vbdB/9R2Y//th7P/53Wf/+JJgf/jToX/4kyE/+JKg//iSYL/5EiC/+BEf/zhQ38iAAAAAPKIqGf/kbX/8YWm//GDpf/wgaT/8H+j/+9+ov/wdZz/5qK4/9Pb2P/LzMz/x8fH/8fHx//Ly8v/7+/v/+rq6v/U1NT/xcXF/6+vr/+ysrL/sbGx/8rLyv/X4Nz/"
    s = s & "9Mbe/+2Bqv/ddIP/33eC/+l3m//xqs3/+Pr7/+l8o//jSID/41CG/+NOhf/iTIT/4kqD/+VKg//gRn/84UN/IgAAAADyiqhn/5O2//KHqP/xhab/8YOl//CBpP/wf6P/8Hed/+ejuf/T29j/zMzM/8jIyP/IyMj/y8vL/+/v7//q6ur/09PT/8XFxf+wsLD/srKy/7Gx"
    s = s & "sf/Kysr/1dbW/+/39P/3+vv//+36///o9//++///9f/9//P49v/peaH/40uC/+NRhv/jUIb/406F/+JMhP/lS4X/4UeB/OFLfyIAAAAA8oqqZ/+Vt//yiKn/8oan//GFpv/xg6X/8IGk//F5nv/npLn/1NvZ/8fIx/+/v7//wMDA/8TExP/v7+//6+vr/9TU1P/Dw8P/"
    s = s & "qamp/6ysrP+rq6v/ycnJ/9bW1v/x8fH/5Obl/7a6uP+3vLr/s7W1/9na2v/4+vr/6Xui/+RNg//kU4j/5FGG/+NQhv/jToX/5k2G/+FJgvzhS38iAAAAAPKNrWf/mLn/8ouq//KJqf/yh6j/8Yan//GEpv/ye6D/6KW7/9LZ1//V1NT/19fX/9XV1f/Y2Nj/8PDw/+rq"
    s = s & "6v/T09P/ycnJ/7y8vP+9vb3/vb29/83Nzf/V1dX/8PDw/+vr6//V1dX/1tbW/9TU1P/l5+b/9vj3/+p9pP/kUIX/5FaK/+RUiP/kU4f/41GG/+ZQiP/hTIT84UuHIgAAAADykq1n/5y7//OOrf/yjav/8ouq//KJqf/xh6j/8n+j/+invP/V3Nr/v8C//7Gxsf+ysrL/"
    s = s & "uLi4/+7u7v/r6+v/1dXV/76+vv+dnZ3/oKCg/5+fn//Hx8f/19fX//Ly8v/k5OT/srKy/7Ozs/+vr6//2NnZ//n7+v/qgKb/5VSI/+ZajP/lWYv/5VeK/+RVif/nVYr/41CG/OFShyIAAAAA8pKtZ/+evP/zkq//85Ct//KOrP/yjKv/8ouq//KDpf/oqb3/0tjW/9nZ"
    s = s & "2f/e3t7/3d3d/9/f3//x8fH/6urq/9LS0v/MzMz/wsLC/8PDw//Dw8P/zs7O/9XV1f/v7+//7e3t/93d3f/d3d3/3Nzc/+jp6f/19/f/7IOo/+ZYi//nXo//5lyO/+Vbjf/lWYz/6FiM/+NSh/zhUociAAAAAPKUr2f/oL3/9JSw//STsP/0ka//85Cu//OOrf/0hqj/"
    s = s & "6ay+/9Xc2v++vr7/rq6u/6+vr/+1tbX/7e3t/+vr6//V1dX/vb29/5ubm/+enp7/nZ2d/8bGxv/X19f/8vLy/+Pj4/+urq7/sLCw/6ysrP/W2Nf/+fr6/+yGqv/nXI3/52KS/+dgkP/mX5D/5V2O/+hbjv/jVIj84VKHIgAAAAD1lK9n/6K///SVsf/1lrH/9JWx//ST"
    s = s & "sP/zka//9Yqq/+mtwP/R2Nb/19jY/93d3f/b29v/3d3d//Dw8P/q6ur/0tLS/8vLy//BwcH/wsLC/8HBwf/Ozs7/1dXV/+/v7//t7e3/29vb/9zc3P/b29v/5+no//X49//siaz/52CQ/+hllP/nZJP/52KS/+ZfkP/pXI//5FeJ/OhShyIAAAAA9ZmyZ/+kwP/0mLL/"
    s = s & "9Zey//WXsv/0lrH/9JSx//WNrP/pr8H/1NrZ/8PDw/+3t7f/uLi4/729vf/u7u7/6+vr/9TU1P/AwMD/oqKi/6Wlpf+kpKT/yMjI/9fX1//x8fH/5eXl/7i4uP+5ubn/tbW1/9nb2v/4+fn/7Yyt/+ljkv/paZb/6WeW/+dkk//nYZH/6l6Q/+RZi/zoWociAAAAAPWZ"
    s = s & "smf/psH/9Zmz//WZs//1mLL/9Ziy//SXs//2kK7/6rHC/9LY1v/R0dH/0NDQ/9DQ0P/T09P/8PDw/+rq6v/T09P/yMjI/7e3t/+5ubn/uLi4/8zMzP/V1dX/8PDw/+rq6v/Q0ND/0dHR/8/Pz//j5eT/9vj3/+2Or//pZ5X/6myZ/+lplv/oZpT/52KS/+pgkf/lWoz8"
    s = s & "6FqOIgAAAAD1m7Rn/6fC//WbtP/2mrT/9Zq0//WZs//1mbP/9pOw/+21xP/c4dj/zczF/8TDvP/FxL3/ycjB//b27v/z8Or/3tbT/8vDwP+wqKT/s6uo/7Kqp//Tysb/39jV//Lz9//m6O3/vb7D/77Axf+6vMH/3N/j//j7///ukbL/6mqW/+pumv/qa5f/6GeV/+hk"
    s = s & "k//rYZL/5VyN/OhajiIAAAAA9Z60Z/+pxP/1nLX/9py1//abtf/1m7X/9Zq0//iXsv/eq8X/sbvV/83R6//g5P7/3uH7/93g+v/X2fT/zNXs/6zM1/+x0Nz/ttXh/7XU4P+11eD/rs7a/7PO2P/s59j/9Oza//nz4f/58uH/+vPi//Pu2//z7tz/7pKt/+ptmv/rcJr/"
    s = s & "6myY/+lpl//oZpT/62OU/+ZejvzoWo4iAAAAAPeet2f/q8X/9p63//aetv/2nbb/9p22//Wctv//orL/pXnJ/yI+3/8/U+7/Sl/8/0he+v9JXvr/Ul36/z50+v8At/r/ALT6/wCx+/8Asvv/ALH7/wCy/v8Vte7/4cds//bIXv/wxl7/8MZe//DGXf/vx2D/9M1k/++M"
    s = s & "k//qcaH/63Kb/+pumf/pa5f/6WiV/+xllf/mYI/86GGOIgAAAAD3oLdn/63G//eguP/2oLj/95+3//aet//2nrf//6Sz/6V6yv8cON//annr/5mk+P+Tn/b/jJn2/1Zh+v87cvr/ALT7/yi++f9sz/b/Zsz2/2jN9v8QuP7/DbPv/+HFaP/2y2v/8Nif//DYnf/w2aH/"
    s = s & "7854//TLXv/vjZT/63Oi/+tznf/qcJr/6m2Y/+lplv/tZ5b/52GQ/OhhjiIAAAAA9aG2bf+wyP/3orj/9qG4//eguP/3oLj/9qC4//+mtP+hes3/FDjk/0Bb8/9Wb///U2z+/1Fr//9LXf//NHT//wC7//8Au///Cb3//we9//8Ivf//ALn//wW68//hzWn/+NBe//LS"
    s = s & "bP/y0mv/8tJs//HQZP/21GH/75CU/+x0pP/sdZ7/63Kb/+pumf/qa5f/7WmW/+djkv7pZZEjAAAAAPihukr7q8L9+qS7//ejuf/3orn/96K5//ehuf/7pLf/0pHC/5d2z/+lgNT/q4bY/6qE2P+qhNj/roLX/6OM1v+Fqtb/g6nV/4Cm1f+AptT/gKXT/4Kk1P+Opc3/"
    s = s & "7KyT//WsjP/zq4v/9KqL//Ooiv/yp4r/86eK/+6Gnv/seKP/7Hee/+tznP/qcJr/6myY/+5rmP/oZJL17WCVHQAAAAD2p7gd96O4/vynvf/4pLr/96S6//ejuv/3o7r/9qK6//yluP//qLb//6a1//+ltP//pbT//6S0//+is///oLL//5my//+XsP//lq///5Su//+T"
    s = s & "rf//kaz//4+s//KMtP/wirT/8Imz/++Gsf/ug6//7X+t/+18q//tfqX/7Xyi/+x4oP/rdZ3/63Kb/+pumf/2cJ7/6WeTygAAAAAAAAAA9K23Gfimu/P9qsD/+Ka7//imu//4pbv/+KW7//eku//3pLv/9qO6//aju//2o7r/9qK6//aju//1orr/9aC5//Sfuf/0nbj/"
    s = s & "9Jy3//Obtv/zmbX/85e0//OWs//ylLL/8pOy//GQsP/xja3/74qr/++Hqf/uhKj/7oGm/+1+o//teqH/7Hef/+t0nP/rcJr/+nOi/+polbcAAAAAAAAAAAAAAAD5p7yx/7XL//mnvP/4p7z/+Ka8//imvP/3prz/96W8//elvP/2pbv/9qS7//aku//2pLv/9qS7//Wj"
    s = s & "vP/1orv/9KC6//Sfuf/0nbj/9Jy3//Satv/zmbX/85e0//KUsv/ykrD/8Y6u/++LrP/viKr/74Wo/+6Cpv/uf6T/7Xyi/+x4oP/sdZ7/63Kb//93pv/oa5ZkAAAAAAAAAAAAAAAA9qi9Pv+1y//9rMH/+ai8//movf/5p73/+Ke9//envf/3pr3/96a9//amvP/3prz/"
    s = s & "9qa8//alvP/2pbz/9aW9//WjvP/1obv/9aC6//Wfuf/0nbj/9Jy3//OZtf/zlrP/8pOx//GQr//wjK3/74mr/++Hqf/vhKf/7oGl/+19o//seqH/63ae//t7p//qb5rk53OiCwAAAAAAAAAAAAAAAAAAAAD6qbx3/7jN//2twv/5qb3/+am9//mpvv/4qb7/96i+//en"
    s = s & "vf/3p73/96e9//envf/2pr3/9qa9//amvf/1pb3/9aS8//WjvP/1obv/9aC6//SduP/0mrb/85e0//KUsv/xkbD/8Y6u//CLrP/viKr/74Wo/++Cpv/tfqT/7Huh//l/qP/0eaLx7XKZOgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPqrwHf/t8z//7jO//+vw//9rMH/"
    s = s & "+qvA//ipv//4qb//96m///epv//3qb7/96i+//aovv/2qL7/9qi+//anv//2pr7/9aW9//Wiu//1n7n/9Jy3//OZtf/ylrP/8pOx//GPrv/xjK3/8Imr//KIq//yhan/94Wr//+Jsv/2f6fs6XSbOwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA+63BQv2y"
    s = s & "xq36r8Tu+q3C+/+yx///utD//7jO//+4zv//uM7//7fP//+3zv//ts7//7bO//+2zv//tc3//7TN//+yzP//r8r//6zI//+pxv//p8T//6TC//+gwP//nb3//5m7//+Xuv/ziqz/8oeq+/WFq93ygqaT53uaIQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAA/rHHF/650CH4rMBK/LTIZ/eqvmf3qr5n96q+Z/eqvmf3qL5n96i+Z/elvGf1pbxn9aO8Z/WjuWf1oLln9aC5Z/Ket2fym7dn8pm0Z/KXtGfylLJn8pKvZ/CPrWf6krJn8IutNf6SuSH/mbIKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAA=="
    IconB64_Kolonnevelger = s
End Function
Private Function IconB64_MailUtsender() As String
    Dim s As String
    s = s & "Qk02GQAAAAAAADYAAAAoAAAAKAAAACgAAAABACAAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7H9IHPeHUiLoeEsi6HhLIuh4SyLoeEMi6HhDIuh4QyLocEMi6HBDIuFwPCLhcDwi4Wk8IuFpPCLhaTQi"
    s = s & "4Wk0IuFpNCLhYTQi4WE0IuFhNCLhYS0i92k0ItJaLREAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1H9VDPKIVGPthVKz7IJRxeZ+TfTofkz85XtJ/OV6SPzleEb85HdF/OR1RPzjdEL843JB/OJxP/zibz784W49/OFs"
    s = s & "O/zgazr84Go4/N9oN/zfZzX83mU0/N5kMvzeYjD83WAv/OFhL/zcXizm418txeVfK53cWyhLAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5YVSO/SMWdz9kFv/+IxX//WIVf/sg1H/6oFP/+p/Tf/qfkz/6nxK/+l7SP/oeUf/6HhG/+h2RP/ndEP/"
    s = s & "5nNB/+ZyQP/mcD7/5W89/+RtO//kbDr/42o4/+NpN//jZzX/4mUz/+JjMv/hYjD/5WMw/+ljL//uZC//62Et/+FbKbfMVSIPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA6YhXOvOOWvH6kVz/6odV/+mFVP/og1L/6IJR/+iAT//nf07/535M/+Z8Sv/me0n/5nlH/+V4"
    s = s & "Rv/ldkX/5HRD/+RzQv/jcUD/43A//+JvPv/hbTz/4Ww7/+FqOf/gaTj/4Gc2/99lNP/fZDP/32Ix/95hMP/eYC//3V4t/95dLP/xZC7/4Vspxd9PLxAAAAAAAAAAAAAAAAAAAAAA7pBdHvCOXPH4kl7/6olX/+qIVv/qh1X/6YVT/+iDUv/oglH/6IBP/+2BRf/sf0T/"
    s = s & "5nxL/+Z7Sf/meUf/5XhG/+V2RP/kdEP/5HNC/+NxQP/jcD//4m49/+FtPP/hbDv/4Wo5/+BpOP/gZzb/32U0/99kM//fYjH/3mEw/95fLv/dXi3/3Fws//JkLv/ZWSe6AAAAAAAAAAAAAAAAAAAAAO2OXZH/nGX/64xa/+uLWf/rilj/6ohW/+qGVf/phVP/6oRQ/+aB"
    s = s & "U/+daLf/ynV2//KBPf/mfEr/5ntJ/+V4SP/kdEf/5XZE/+R0Q//kc0L/43FA/+NwP//ibj3/4W08/+FsO//hajn/4Gg3/+BnNv/fZTT/32Qz/99iMf/eYTD/3l8u/91eLf/eXSz/62Et/9pYKEUAAAAAAAAAAO+PXxDtkF7l9pVh/+yOXP/sjFr/64tZ/+uKWP/qiFb/"
    s = s & "6oZV/+mFU//shUz/uXKR/3Rb8P+8cYn/8oE9/+Z5Sf/ul0H/7ZJA/+NzR//ldkT/5HRD/+RzQv/jcUD/43A//+JuPf/hbTz/4Ww7/+BqOf/gaDf/4Gc2/99lNP/fZDP/3mIx/95hMP/eXy7/3V4t/+5kL//bWymkAAAAAAAAAADulGIf7JFg//CTYP/tj13/7I5c/+yN"
    s = s & "W//ri1n/64pY/+qIVv/qh1X/6YVU//OHQ//OeXT/d1zq/6hrpv/vfkD/7YdC//euOv/tlD//5HVG/+Z2Rf/mc0T/43JC/+NxQP/jcD//4m8+/+JtPP/hbDv/4Wo5/+BpOP/gZzb/32Y1/99kM//fYjH/3mEw/95gL//qYy//3FwrxwAAAAAAAAAA8JVlNfKWZPzwk2L/"
    s = s & "7ZFf/+2PXf/sjlz/7I1b/+uLWf/rilj/6ohW/+qHVf/phVT/8YZG/99/Xf+CX9z/k2TD/+h7Sv/uhkD/7JJC/+V4SP/idkX/331A/+V0RP/kc0L/43FA/+NwP//ibz7/4W08/+FsO//hajn/4Gk4/+BnNv/fZTT/32Qz/99iMf/eYTD/5WIw/9xeLOblWTMUAAAAAO6T"
    s = s & "ZGv/oWz/7pNi/+2SYP/tkV//7Y9d/+yOXP/sjFr/64tZ/+uKWP/qiFb/6odV/+iFVP/vhUn/6oJN/5NlxP+CX9n/3Hla/+t6RP/rfEv/yWxA/7yGMP/xjUH/429E/+RzQv/jcUD/43A//+JuPf/hbTz/4Ww7/+FqOf/gaTj/4Gc2/99lNP/fZDP/32Ix/+BiMf/fYC//"
    s = s & "3GApJQAAAADtl2Vn/6Jt/+6VY//uk2L/7ZJg/+2RX//tj13/7I5c/+yMWv/ri1n/64pY/+qIVv/qhlX/6YVU/+yETf/whET/p2up/3td5v/aeV7/8YBI/6RgNf+dcin/+7M4/+d+Q//jcUT/5HNC/+NxQP/ibj3/4Gg5/+BoOf/hbDv/4Wo5/+BoN//gZzb/32U0/99k"
    s = s & "M//iZDL/3WAv/OFhLSIAAAAA7ZlnZ/+kb//vlmX/7pVj/+6TYv/tkmD/7ZFf/+yPXf/sjlz/7Ixa/+uLWf/rilj/6ohW/+qGVf/phVT/6oRQ//CERf/Jdnr/3nxg/+J8Rv+DVir/kWQp//KyNP/zpjj/5XVC/+JsP//jbT7/5XdE/+uQU//nhEv/4Gk5/+FsO//gajn/"
    s = s & "4Gg3/+BnNv/fZTT/4mUz/95iMPzhYS0iAAAAAPCZZ2f/pnD/75hn/++WZf/ulWT/7pRi/+2SYf/tkV//7Y9d/+yOXP/sjVv/64tZ/+uKWP/qiFb/6odV/+mFVP/phFL/7oNI//KET/+8bz//dlEl/4dbKf/rqDL/9bM0//GjSP/um1//9bZu//vLev//3Yb/97xx/+Fp"
    s = s & "Ov/hbDv/4Ww7/+FqOf/gaTj/4Gc2/+NnNf/eZDL84WE0IgAAAADwm2pn/6dy/++aaP/vmGf/75Zl/+6VZP/ulGL/7ZJg/+2RX//tj13/7I5c/+yNW//ri1n/64pY/+qIVv/qhlX/6YNS/+l/UP/kekz/jlgr/3pTJv97UST/zpxK///Xe//+1H///tiD//7Xgv/904D/"
    s = s & "/NF///3TgP/ngUr/4Wo6/+FtPP/hbDv/4Wo5/+BpOP/jaTf/3mU0/OFhNCIAAAAA8J5qZ/+pc//wm2r/75lo/++YZ//vlmX/7pVk/+6TYv/tkmD/7ZFf/+2PXf/sjlz/7Ixa/+uLWf/riVj/6ohW/+uOWf/xnmL/665p/593P/91SyH/eE4i/31VK//Hn17//9uG///T"
    s = s & "gP/8z37//NB+//zQfv/+2YT/8aRi/+FoOv/ibj3/4W08/+FsO//hajn/42o4/99nNfzhaTQiAAAAAPCebGf/q3X/8J1r//Cbaf/vmWj/8phl//iXYP/3ll3/95Rc//aTWv/2kln/9pBX//WPVv/0i1P/9ZNX//+7bP//13v//9t+///bgv//1oL/zaVg/4xiMf9yRx7/"
    s = s & "c0ke/7OKTP/91oL//9aC//zPfv/80H7//dSB//nGd//jckH/4m8+/+JuPf/hbTz/4Ww7/+RsOv/faDf84Wk0IgAAAADwoG9n/6x2//GebP/wnGv/85tn/+6Yaf/Ok33/zJN8/8ySev/MkHj/y493/8uOdv/LjXT/y4tz/8uLcv/MkXX/0qV//9a9jf/30ID//9l1///f"
    s = s & "gf/4znz/uZBR/39UJ/9qQBj/nnQ9//TMe///2YT//M9+//zQfv/+14P/649U/+JqPP/jcD//4m49/+FtPP/kbTv/4Go4/OFpNCIAAAAA8qBvZ/+teP/xoG3/855r//Odaf9om8f/HoLu/yZ+5v8mgOf/JoDn/yaA5/8mgOf/JoDm/yaA5/8mgOf/Jn7m/yV64/8ieub/"
    s = s & "QYrX/46osP/cxYr//9l5///ifv/twnH/pXxD/25DG/+HXSz/5r5y///chv/90H7//tiD//W0bP/ibD3/43FA/+NwP//ibz7/5W48/+BrOvzhaTwiAAAAAPKjcWf/sHv/8qFw//ega//hn3f/KKDw/xeU9P8ifeb/In3m/yJ+5/8ifuf/In7n/yJ+5/8ifuf/In7n/yJ+"
    s = s & "5/8ifuf/I33m/xt66v8Sdu//JX7l/2GWxv+ytp7/+tWE///he//is2L/jWIw/3ZLIf/QqGL//92H//7TgP/8z33/5n9K/+NxQf/jckH/43FA/+ZxQP/hbT384XA8IgAAAADypXZn/7N///KkdP/3o27/46F5/zCe6v8Zn/f/IZLu/yV85P8lfeX/JX7l/yV+5f8lfuX/"
    s = s & "JX7l/yV+5f8lfuX/JX7l/yV+5f8lfuX/JX7l/yJ95/8WeO3/Fnjt/zuH2v+HpbX/282V///Wef/Omkz/g1Uk/7WMTv//2IT//92G/++gYP/jcEL/5XZF/+R0RP/ndET/4nBA/OFwPCIAAAAA8qV2Z/+0gf/zp3f/+KZy/+Ojff8wnur/GZz1/x+f8/8hj+z/JXvk/yV+"
    s = s & "5f8lfuX/JX7l/yV+5f8lfuX/JX7l/yV+5f8lfuX/JX7l/yV+5f8lfuX/JX7l/yV+5f8de+n/Enbu/yF65v9ZotL/scuw/+DDev+yfzj/q4FE//HPfv/+yHr/5XhJ/+V4SP/ld0j/6HdH/+NyQvzhcEMiAAAAAPKodmf/toL/86l5//qodv/lpoD/MJ7q/xmc9f8fnPL/"
    s = s & "H5/z/yKM6/8le+T/JX7l/yV+5f8lfuX/JX7l/yV+5f8lfuX/JX7l/yV+5f8lfuX/JX7l/yV+5f8lfuX/JX7l/yZ75P8giez/EZn6/w6Y+/8zpe3/h73E/8ygYv+8kU//6sp4/++XXf/ld0r/5XpL/+h5Sf/jdEP84XBDIgAAAAD1qnln/7iE//Sqe//6qnj/5qiD/zCf"
    s = s & "6/8ZnPX/H5zy/x+c8v8fnvP/I4nq/yV75P8lfuX/JX7l/yV+5f8lfuX/JX7l/yV95f8lfuX/JX7l/yV+5f8lfuX/JX7l/yV85P8jh+n/H53z/x+d8v8fnPL/E5r5/xaZ+P/SkHT/8ZNX/9WbXP/oomP/53xP/+Z8Tf/pe0r/5HVF/OhwQyIAAAAA9ap7Z/+6hv/0rHz/"
    s = s & "+qt6/+aqhf8wn+v/GZz1/x+c8v8fnPL/H53y/x+e8/8jh+n/JXvk/yV+5f8lfuX/Jnrk/yZ65f8lf+X/Jnvk/yZ55P8lfuX/JX7l/yV75P8khOj/H5zy/x+d8v8fnPL/H5zy/xuc9P8mnvD/zItv//GCUP/qg1b/5oRV/+eCUv/mfk//6nxM/+R3RvzoeEMiAAAAAPWt"
    s = s & "e2f/u4f/9a1+//ute//mq4b/MJ/r/xmb9f8fnPL/H5zy/x+c8v8fnfP/H53y/ySE6P8meuT/Jnjk/yKO5/8csu3/Grzv/xy07v8ikej/Jnjk/yZ65P8kg+f/H5vy/x+e8/8fnPL/H5zy/x+c8v8bnPT/Jp7x/86Oc//xiFb/6Yha/+iEVv/nglP/539Q/+p+Tf/leEf8"
    s = s & "6HhDIgAAAAD1r35n/72I//Wvf//7r3z/56yH/zCf6/8Zm/X/H5zy/x+c8v8fnPL/H5zy/x+e8/8gl/H/JILn/x+i6v8Yx/D/GMrx/xjH8P8YyfH/GMjx/x6m6/8kguf/IJbw/x+e8/8fnPL/H5zy/x+c8v8fnPL/G5z0/yaf8f/OkXb/8YpZ/+mKXP/phlj/54NV/+eB"
    s = s & "Uf/qf07/5XpJ/Oh4SyIAAAAA9a9+Z/+/iv/1sID/+7B+/+euiP8xn+z/GZv1/x+c8v8fnPL/H5zy/x+b8v8fmPL/HqTy/xq98f8YyvH/GcXw/xnE8P8ZxPD/GcTw/xnF8P8YyvH/Gr/x/x2m8v8fmPL/H5vy/x+c8v8fnPL/H5zy/xuc9P8mn/H/zpN5//KNW//pil3/"
    s = s & "6YhZ/+iFVv/nglP/64FQ/+V7SvzoeEsiAAAAAPeygGf/wIz/9rKC//yyf//nsIr/MZ/s/xmb9f8fnPL/H5zy/x+a8v8fmvL/HK/x/xnF8P8Zx/D/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnH8P8ZxvD/HLHx/x+b8v8fmfL/H5zy/x+c8v8bnPT/Jp/x/8+U"
    s = s & "e//yj13/6oxf/+qJW//oh1j/6IRU/+uCUv/lfUz86H9LIgAAAAD3tINn/8KN//azg//8s4D/6LGL/zCg7P8Zm/X/H5zy/x+Y8v8eofL/G7vx/xnI8P8ZxfD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/GcXw/xnI8P8avPH/HqPy/x+Y8v8fm/L/"
    s = s & "G5z0/yaf8f/Qlnz/85Be/+uNYP/qilz/6YdZ/+iFVv/shFP/5n5N/Oh/SyIAAAAA9bSDbf/Ej//3tIT//LSC/+iyjP8xoOz/GZn1/x+Z8v8drPH/GcTw/xnH8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/Gcfw/xnF"
    s = s & "8P8crvH/H5ry/xua9P8mn/H/0Jd9//SRX//rj2H/6oxe/+mJW//phlf/7IVV/+aAT/7pg0gjAAAAAPi2hkr7vov9+beG//22g//os43/MZzs/xie9f8buPH/Gcfw/xnF8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/"
    s = s & "GcTw/xnE8P8ZxfD/Gcjw/xu78P8aoPT/J5vx/9GYfv/0k2D/7JBi/+qNX//pilz/6YdZ/+2HV//ngVD17YNPHQAAAAD2uIMd9rWF/vy6if/9t4T/6LSP/y6t7P8TwvP/Gcjw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE"
    s = s & "8P8ZxPD/GcTw/xnE8P8ZxPD/GcTw/xnE8P8Zx/D/FcPy/ySu8f/RmH//9ZRh/+yRZP/rj2D/6oxd/+mJWv/1jVv/6INTygAAAAAAAAAA9LeOGfe4iPP9vYz//biF/+i4kP8iyO//Dcb1/xTE8v8UxPL/FMTy/xTE8v8UxPL/FMTy/xTE8v8UxPL/FMTy/xTE8v8UxPL/"
    s = s & "FMTy/xTE8v8UxPL/FMTy/xTF8v8UxfP/FMXz/xTF8/8UxfP/FMXz/w/G9f8XyvX/0KCA//aUY//tk2X/65Bi/+uOX//qilz/+ZBd/+eEVLcAAAAAAAAAAAAAAAD5u4qx/8mV//m5if/6uIf/csPL/xzG8f8lxOz/JcTs/yXE7P8lxOz/JcTs/yXE7P8lxOz/JcTs/yXE"
    s = s & "7P8lxOz/JcTs/yXE6/8lxOv/JcTr/yXD6/8lw+v/JcPr/yXD6/8lw+v/JcPq/yTD6v8dxfD/YLrK/+6abf/xl2j/7ZRm/+ySY//rj2D/64xd//6VYf/ohFZkAAAAAAAAAAAAAAAA9rmLPv/Jlf/8v47/+bqK//u5if/lupT/5bqU/+W5lP/luZT/5LmV/+S5lf/luJb/"
    s = s & "5LiW/+S4lv/juJb/47eX/+O2lv/itZT/4rST/+Kzkv/hspH/4bCQ/+Cujf/grIv/4KqI/9+nhf/epYL/3aN///GdcP/ymmz/7phr/+2WaP/sk2T/65Bh//uXZP/qilrk/otcCwAAAAAAAAAAAAAAAAAAAAD6uot3/8uY//zAkP/6u4v//7qJ//66iv/+uor//rmK//25"
    s = s & "iv/9uIv//biM//24jP/9t4z//LeN//y3jv/8t43/+7WM//u0iv/6s4n/+rKI//mvhf/5rYL/+Kt///iofP/3pXn/9qN1//agcv/ynnH/75xw/++abP/tl2n/7JRl//mZaP/0lGPx6YxcOgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPq+j3f/ypj//8uZ///Bkf/8vpD/"
    s = s & "+ryP//i6j//3uo//97qP//e6kP/3upD/97mQ//a5kf/2uZH/9riS//a4kv/1t5H/9baQ//S0jf/0son/86+G//OthP/yqoH/8ah+//Glev/wonf/8KB0//Kfcv/ynW//955u//+kcP/2mWfs6Y5fOwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA972OQv3F"
    s = s & "lK36wJHu+b6R+//Elv//zJz//8qb///KnP//ypz//8qd///Jnf//yJ3//8id///Inv//x57//8ed///Fm///wpj//8CV//++kv//u4///7mM//+2if//s4X//7CC//+vf//zoXT/8p5w+/Web93ymmmT55JkIQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAA/seQF/7ImiH4vZBK/MaXZ/e8kmf3vI9n97ySZ/e8kmf3uZJn97mPZ/e5j2f1t41n9bSNZ/W0jWf1sopn9bKIZ/KviGfyr4Vn8q2DZ/Kqg2fwqIBn8KV7Z/CjeWf6qHtn8KNzNf6peyH/sn8KAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAA=="
    IconB64_MailUtsender = s
End Function
Private Function IconB64_Kontaktsentralen() As String
    Dim s As String
    s = s & "Qk02GQAAAAAAADYAAAAoAAAAKAAAACgAAAABACAAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf6MtHIesLSJ4pS0ieKUtInidLSJ4nSUieJ0lInidJSJ4nSUicJ0lInCWJSJwlh4icJYeInCWHiJwlh4i"
    s = s & "aY4eImmOFiJpjhYiaY4WImmOFiJphxYicJYWIlqHDxEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAapQqDIixNmOErTOzgqoyxX6mL/R+pS78e6Is/HqhK/x5nyr8eJ4o/HecJ/x1myX8dJok/HOYI/xylyL8cZYh/G+U"
    s = s & "IPxukx78bZEd/GyQG/xrjxr8ao0Z/GiMF/xnihb8ZokU/GaKE/xjhhHmZokQxWSKEJ1ihA1LAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgag4O4m0ONyPujn/irU2/4eyNP+CqjH/gKgw/3+nL/9+pS3/fKQs/3uiK/96oSn/eaAo/3ieJ/92nSX/"
    s = s & "dZsk/3SaI/9zmSL/cZcg/3CWH/9vlB3/bpMc/22SG/9rkBn/ao8Y/2mNFv9njBX/aI0U/2mOE/9qkBL/aIwQ/2KFDbdVdxEPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAiK85Oou1O/GOujv/hK02/4OsNf+CqjP/gKky/3+nMf9+pi//faUu/3yjLf96oiz/eaAq/3if"
    s = s & "Kf93nij/dpwm/3SbJf9zmST/cpgi/3GXIf9wlSD/bpQf/22SHf9skRz/a5Ab/2qOGf9ojRj/Z4sW/2aKFf9liRT/ZIcS/2OGEf9rkRH/YoUOxV9/DxAAAAAAAAAAAAAAAAAAAAAAiLI7Hou1PPGPuj3/hq85/4GuOf9/rTj/fqs2/32qNf97qTT/eqcz/3mmMf94pDD/"
    s = s & "dqIv/3WhLf90oCz/c58r/3GdKf9wmyf/cJsm/3CaJf9umST/bZci/2uVIf9qlCD/aZIe/2iRHf9mjxv/ZY4a/2SNGP9jixf/YooW/2CIFP9ghhL/Y4UR/2uREf9fgg26AAAAAAAAAAAAAAAAAAAAAIq1PZGYxUP/h7I8/4ywOP/BrR7/xqsa/8OqGv/Dqhr/w6kZ/8Kp"
    s = s & "Gf/CqRn/wagY/8GoGP/Bpxj/wKcX/8CnF//Aphf/wacW/7GaFf+mkhT/qJMU/6eTE/+nkhP/ppIS/6aSEv+mkRL/pZER/6WQEf+lkBD/pJAQ/6SPEP+ojw//k44Q/2CGEv9jhhH/aI0Q/2CBDkUAAAAAAAAAAH+vPxCMtT/lkbtA/4izPv+Msjr/260T/+WnC//iqA3/"
    s = s & "4qgN/+KoDf/iqA3/4qgN/+OoDf/jqQ3/46kN/+OpDf/jqQ3/46kN/+WrDf/PmQz/vo0L/8GQDP/BkAz/wZAM/8GQDP/BkAz/wZAM/8KQDP/CkAz/wpAM/8KQDP/CkQz/yZIN/6iQD/9giBT/Y4cS/2qQEv9hhA+kAAAAAAAAAACLvUEfjbZB/4+4QP+LtT7/hbQ//8Wu"
    s = s & "Hf/iqAz/3KgP/9yoD//cqA//3KgQ/9ukBf/aowH/2qMB/9ulCP/cqBD/3KgP/9yoD//eqg//yZkO/7mODf+8kA3/vJAN/7yQDf+8kA3/vJAN/7yQDf+8kA3/vJAN/7yQDf+8kA3/vJAN/8SSDf+TjxL/YIkW/2WJFP9qjxP/YoUQxwAAAAAAAAAAkLtDNZK9Q/yPuUL/"
    s = s & "jbZA/4a1Qf+osi7/5akM/9yoD//cqA//3KgP/9ynDf/htjj/47tI/+O8SP/gsi3/26YL/9yoD//cqA//3qoP/8mZDv+5jg3/vJAN/7yQDf+8kA3/vJAN/7yQDf+8kA3/vJAN/7yQDf+8kA3/vJAN/72QDf+/kw3/d44W/2SLF/9mihX/aI0U/2OGEeZmjAwUAAAAAJG5"
    s = s & "Qmudykn/j7hC/463Qf+MtkD/irU//86uGv/jqAz/3KgP/9yoDv/bpQX/9OrS//3////+////7tid/9qhAP/cqA//3KgP/96qD//JmQ7/uY4N/7yQDf+8kA3/vJAN/7yQDf+8kA3/vJAN/7yQDf+8kA3/vJAN/7uQDf/FkQz/npIT/2SOGv9ojRj/aIsW/2eMFf9liRT/"
    s = s & "Z4INJQAAAACSvEVnncpK/5C6Q/+PuEL/jrdB/4i2Qv+ZtDj/3awR/+CnDP/bqA7/26cL/+bDYP/qzn3/6s59/+O8Sv/bpQj/3KgP/9yoDv/jqgH/z5kA/7yOAv+8kA3/vJAN/7yQDf+8kA3/vJAN/7yQDf+8kA3/vJAN/7yQDf/DkAv/tJQQ/3CRG/9pjxv/ao4Z/2mN"
    s = s & "GP9pjRb/ZokU/GmHFiIAAAAAkrxFZ57LS/+Su0T/kLpD/4+4Qv+Ot0H/hrZD/5+zNP/crBL/5KgL/9ynD//aowH/2qEA/9qhAP/bpAX/3KgR/92oDP/fqAX/yKlO/6mZcf+vjzX/wJAA/7yQDf+8kA3/vJAN/7yQDf+8kA3/vJAN/7yPDf/EkAv/tpUQ/3iUHf9pkh7/"
    s = s & "bJEc/2uPGv9qjhn/ao8Y/2eKFvxpjhYiAAAAAJS+R2egzU3/k71G/5G7RP+QukP/j7lC/463Qf+HtkP/lbQ5/8mvHf/jqg3/46gM/9+nDv/cpw//3KgP/9yoDv/hqAT/saaM/4am//9+mf//fJHo/6aRWv/AkAH/vJAN/7yPDf+8jw3/vo8M/8OQC//Bkg3/pZYV/3WX"
    s = s & "IP9rlSH/b5Qf/22SHf9skRz/a5Ab/2uQGf9ojBf8aY4WIgAAAACUwUpnoc9O/5S+R/+TvUb/krtF/5C6Q/+PuUL/jrdB/4i2Qv+HtUD/orMx/8SvH//YrBT/4KsP/+OqDv/nqgL/06k2/4mk//+Mpv//hZno/3uR6f+FktD/vpET/8KRB//Ckgz/v5IN/7eUEf+llxf/"
    s = s & "hpog/2+ZJf9vmCP/cZch/3CVIP9vlB//bpId/2yRHP9tkhv/ao0Z/GmOFiIAAAAAl8FKZ6LQUP+VwEn/lL5H/5O9Rv+Su0X/kLpD/4+4Qv+Ot0H/jLZA/4a1Qf+FtED/jrI5/5uxMv+hry3/oq4e/5usa/+Mpv//jqb8/4WZ6f9/kOD/gZTf/4+ePv+Rnxz/j54j/4if"
    s = s & "Jf97nyj/cp4p/3KcJ/90myX/dJkk/3KYIv9xliH/cJUg/2+UH/9ukh3/bpMc/2uPGvxpjhYiAAAAAJfDTGej0VH/lsFK/5W/SP+Uvkf/k71G/5G7RP+QukP/j7hC/463Qf+NtkD/i7Q+/4izPv+Fsj3/grE8/3+wL/+GrnX/j6f//46m/P+Fmen/f5Dg/4KV3f97pUr/"
    s = s & "eKYp/3ejL/94pC7/eqQs/3ifKf93nSf/dpwm/3WbJf9zmSP/cpgi/3GWIf9wlSD/b5Qf/2+UHf9skBv8aY4eIgAAAACZw0xnpdNT/5jCS/+WwUr/lb9I/5S+R/+TvUb/kbtE/5C6Q/+PuEL/jrdB/422QP+LtD7/irM9/4myPP+HsS3/ia5x/4uk//+Ko/z/gZbn/3yM"
    s = s & "3f9/ktv/f6VF/36nKf9+pi//eJwq/3WXKP98piz/eJ8p/3edJ/92nCb/dZsl/3OZI/9ymCL/cZYh/3CVIP9wlh//bZEd/HCOHiIAAAAAm8ZPZ6bUVP+ZxEz/l8JL/5bBSv+VwEn/lL5H/5O8Rf+Ru0T/kLpD/4+5Qv+Ot0H/jbU//4u0Pv+Ksz3/hbAo/5Czev+jt///"
    s = s & "prn+/5qt9P+TpfD/kaTo/4KoTP99pyP/g64z/2d8H/89LQP/ZXkd/32mLP94nyn/d54n/3acJv91miT/c5kj/3KYIv9xlyH/cZcg/26THvxwlh4iAAAAAJvIUWeo1lf/m8VP/5nETv+Ywkz/l8FL/5bASv+Vv0j/k71H/5K8Rv+RukT/kLlD/4+3Qv+NtkH/ibQ1/5i7"
    s = s & "af++z+r/wtD//8LQ//+1xv//q7///63B//+nvtb/hqxG/4GrLP91lCr/QTQF/0M4B/95nSv/fKMt/3mfK/94nin/d50n/3abJv91miX/c5kk/3SZI/9wlSH8cJYeIgAAAACeyFRnq9lb/53IU/+cxlH/m8VQ/5rDT/+Ywk7/l8BM/5a/Sv+Vvkn/lL1I/5O7R/+Rukb/"
    s = s & "j7hC/5C3Rf+6zeP/wc7//73M//++zf//scP//6e8//+pvf//rsD//6C7vP+BqSz/h7M4/1xlGf85JwD/c5Ep/4CpMv98oi7/e6Et/3qfK/95nir/d50p/3abKP93nCf/c5gk/HCWJSIAAAAAnstWZ63aXf+gylf/nshV/53HVP+cxlP/m8RR/5rDUP+ZwU7/l8BN/5a/"
    s = s & "TP+Uvkv/k7xK/4+6QP+cv2z/wdD//73M//+9zP//vs3//7HD//+nvP//qb3//6q9//+pv+3/iK5E/4i0OP9daBz/NyQA/3WRLP+DrDb/f6Uz/36jMf99ojD/e6Au/3qfLf95niz/ep8r/3WZJvxwliUiAAAAAKDLVmeu3F//ostZ/6HKWf+gyVf/nshW/53HVf+cxVP/"
    s = s & "m8NS/5rCUP+XwVH/mMBO/5i+TP+Uu0H/nsN1/8HQ//+9y///vcz//77N//+xw///p7z//6m9//+qvf//qL/w/4qyTP+NtTr/YWod/zYkAP91kS7/h686/4KnNv+BpjX/f6Q0/36jMv99ojH/fKAw/3ugLf92mif8eJ0lIgAAAACgzVlnr91g/6LNWv+jzFv/ostb/6HK"
    s = s & "Wv+gyFj/nsdX/53GVf+ZxVf/qr9G/9avGf/UrBf/1qsM/8e4UP+8z///vcz//73M//++zf//scP//6e8//+pvf//qb7//6m+7P+xnSj/uJUO/7KMD/+feQr/g5wy/4avPf+Fqjr/g6g5/4KnOP+Bpjb/gKQ1/32iMv98oi7/d5wp/HidJSIAAAAAo89ZZ7DfYf+kzlv/"
    s = s & "pM1c/6TMXf+jzF3/ostb/6HJWv+fyFj/mchc/7i7O//mpwj/36cM/+GlAP/Zrjz/vM3//73M//+9zP//vs3//7HD//+nvP//qb3//6i+//+ruun/vpEY/76NBf/AkAv/xZYO/5arNv+FrkH/h6w9/4arPP+FqTv/g6g5/4GmN/9+pDP/fqMv/3idKvx4nSUiAAAAAKPP"
    s = s & "W2ex4GP/pc9d/6XOXf+lzl7/pM1e/6TMX/+jy13/oslc/5zKYP+3vT//4qgK/9yoD//epgL/1q8//7zN//+9zP//vcz//77N//+xw///p7z//6m9//+ovv//qrrp/7uTG/+9jwf/vI8M/76RDf+Wqzj/iLFE/4quQf+JrUD/h6w//4WpO/+Cpzj/gKQ0/3+kMf95nyv8"
    s = s & "eJ0lIgAAAACl0ltns+Jk/6bRXv+m0F//ps9f/6XOYP+mzmH/ps1h/6TLX/+ezGP/uL5B/+KnCv/cqA//3qYC/9avP/+8zf//vcz//77N///A0P//tMf//6m///+pvf//qL7//6q66f+7kxv/vY8H/7yPDP++kQ3/mK08/4uzSP+MsUT/i69E/4mtQP+Gqjz/g6g5/4Gm"
    s = s & "Nv+ApjL/eqAs/HidLSIAAAAApdJbZ7TjZf+n0l//p9Fg/6fQYP+nz2H/p89i/6fOY/+mzmP/oc9o/7y/Qv/oqwr/4awP/+GkAP/YrTv/vs///8TX///B0f//ucbu/6i35P+jt/P/r8X//7DI//+ruuz/vZAX/7+OBP/ClA3/w5QM/5uuPv+Otk3/j7NI/4yxRf+KrkH/"
    s = s & "h6w9/4WpOv+Cpzf/gacz/3yhLvx4pS0iAAAAAKjUXme15Wf/qNRg/6jTYf+o0mL/qNFi/6jQY/+o0GT/qM9l/6bQaf+xxFL/x5cM/8udEP/ZshT/ybxW/7nL8v+VlJb/c2VB/2dTHf9PPxP/Rjsf/1xXT/+IkbD/q8Ht/7WkL/+6mRT/qH8J/6WFEv+XtUr/kbhP/5G1"
    s = s & "Sv+Nskb/irBC/4itPv+Gqzv/g6k4/4KpNf99oy/8f6UtIgAAAACo12Bnt+Zo/6nVYf+p1GL/qdNj/6nSY/+p0mX/qdFl/6nRZv+q0Wf/qNJr/2xkG/9dRQb/nrxW/6vZdv95cjz/UDIA/1g/AP9dRgD/STYA/zcpAP81JQD/LxoA/3F6Tf+o1Wj/f5Q8/zMfAP9gaCX/"
    s = s & "m8NY/5S4T/+Rtkv/j7NH/4yxQ/+Jrj//h6w8/4SqOf+Eqjb/fqMx/H+lLSIAAAAAqNdfbbnpav+q1mP/qtVj/6rUZP+q1GT/qtNl/6rSZ/+q0Wf/qtFo/6/ZcP+Qokf/VDUA/3RwJP+t2W7/la1J/2VSDf9ZPgD/XEEA/0UvAP80IwD/NiQA/01FD/+Us0z/nMBa/0lB"
    s = s & "D/85KAL/iqVJ/5vBVv+VuVD/krdM/5C0SP+NskX/irBB/4iuPf+Gqzr/hKs3/3+mMv6DpysjAAAAAKzZY0q04mf9rNll/6vXZf+r1mX/q9Vl/6vUZv+r1Gf/q9Np/6vSaf+s02v/r9Zv/3dyJf9TMwD/fn4t/67XcP+nymb/jJlA/4WMN/96hDb/cn00/4CSQP+ly2T/"
    s = s & "oMNe/1RQGP8vGQD/bngv/6HIXv+YvFT/lrpR/5O4Tv+Rtkr/jrNG/4uxQv+Jrz7/hq07/4atOP+BpzL1g6crHQAAAACv22kdq9dj/q/cZ/+t2Gb/rddm/6zWZv+s1Wf/rNVp/6zUav+s02r/rNJr/6/XcP+pzGn/b2Mb/1MzAP9uYxr/ma5S/63Vbv+w23T/st1z/7Lb"
    s = s & "cf+nzGb/g5VD/0c9DP8vGQD/ZGgm/6LGYP+dwVv/mb1V/5e7Uv+UuU7/krdL/4+0R/+NskP/irBA/4iuPP+MtTv/gqk1ygAAAAAAAAAArdZlGa3aZfOy32n/rtln/63YZ/+u2Gj/rddp/63Wav+u1Wv/rdVr/63UbP+t023/sdlz/6zPbf9/fzD/WT4A/1Y5AP9nVxH/"
    s = s & "eHIl/25vKf9aWR//RDgJ/zAbAP89LgP/dH82/6XKZf+ixmD/nL9a/5q+V/+YvFT/lrpQ/5O4TP+Qtkn/jrRF/4uyQv+Jrz7/j7g9/4KpNrcAAAAAAAAAAAAAAACv3Gexvexw/6/aaP+v2mn/rtlp/6/Yav+v12v/r9ds/6/Wbf+u1W3/r9Vu/67Tb/+x2HT/tNt3/5+2"
    s = s & "XP+AgTL/a1wV/2JMCP9KNwP/PzIG/1JMFv9yejX/mrZb/6zTbf+kx2T/oMNf/57BW/+bv1j/mb5V/5a7Uf+UuU7/kbdK/4+1Rv+Ns0P/irE//5O+QP+EqjhkAAAAAAAAAAAAAAAArNlmPr3scP+04Gz/r9pq/7Daa/+w2Wv/sNls/7DYbf+w123/r9Zu/6/Wb/+w1XD/"
    s = s & "sNRx/7DUcv+02nn/td17/6/Rcv+ox2n/pMNn/6XEaP+tz2//std0/6zQbf+myWb/pMZj/6HEYP+fwlz/nMBZ/5q+Vv+YvFL/lbpP/5O4S/+Qtkf/jbNE/5S+RP+Irzzki7lFCwAAAAAAAAAAAAAAAAAAAACv3Gh3v+9z/7Xgbf+w22z/sdps/7Habf+x2W7/sNhu/7DX"
    s = s & "b/+w13D/sdZx/7HWcv+x1XP/sdR0/7HUdf+y1Xf/stZ3/7LVdv+x1HX/rtFy/6vNbv+py2r/p8ln/6XIZP+ixWH/n8Nd/53BWv+bv1f/mb1U/5a7UP+UuUz/kbdJ/5a/Sf+SukTxjLQ9OgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALHebXe+7HT/wO12/7bhcP+03nD/"
    s = s & "s9xw/7HZcP+y2XD/sthx/7LYc/+y13T/std0/7LWdf+y1Xf/s9V3/7PUeP+y03f/sdJ1/67Qc/+szm//qsxr/6jLaP+myWX/o8di/6HFXv+ew1v/nMBY/5zBVv+Zv1L/msFQ/6DKUf+WvknsirFAOwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAsdxsQrnl"
    s = s & "cq224HHutN9x+7rkdf/C7nv/wet7/8Hqe//B6n3/wel+/8Dofv/B6H//weeA/8Hmgf/B5oH/wOWB/7/jfv+84Xv/uuB4/7jddP+23HL/tNpv/7HYa/+v1mf/rNNj/6rSYf+dwlf/m8FU+5rBUd2WvUyTi7FFIQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAvN1uF8HveyGz3HFKvOZ2Z7LZcWey2XFnstd0Z7LZdGey13Rnstd0Z6/UdGev1HFnr9JxZ63Sb2etz29nqs9sZ6rPbGeozWpnqMtqZ6XLZ2elyGNnoMZgZ6DDXmejy15nmcBWNanQXCGZzGYKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAA=="
    IconB64_Kontaktsentralen = s
End Function
Private Function IconB64_NettskjemaHenter() As String
    Dim s As String
    s = s & "Qk02GQAAAAAAADYAAAAoAAAAKAAAACgAAAABACAAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArZo2HLSdNCKsljQipY40IqWONCKljjQinY4tIp2HLSKdhy0inYctIp2HLSKWhy0iln8tIpZ/JSKWfyUi"
    s = s & "ln8lIo54JSKOeCUijnglIo54HiKHcB4iln8lIodpHhEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlJQqDLaiO2OznjqzsZs4xauWNvSrlTb8p5M0/KWRM/yjjzL8oo0x/KCLMPyeii/8nYgu/JuGLfyZhSz8mIMr/JaC"
    s = s & "KvyVgCn8k34o/JJ8J/yQeib8j3kk/I13I/yLdSL8inMh/IpzIfyHbyDmiXEfxYhxHp2EbB5LAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAsZs8O7ulPtzBq0D/vKY+/7iiO/+xmzn/rpk4/62XN/+rljb/qZQ0/6eSM/+lkDP/pI4y/6KNMf+giy//"
    s = s & "noku/52ILf+bhi3/moQr/5iCKv+WgCn/lX8o/5N9J/+SeyX/kHkk/453I/+MdSL/jXYi/493Iv+QdyH/jHQf/4VuHbd3ZhEPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAtKI9Or2oP/HBq0H/tJ88/7KdO/+wmzr/r5o5/62YOP+slzf/qpU2/6iTNf+mkTT/pZAz/6OO"
    s = s & "Mv+hjDH/n4ow/56JL/+chy7/m4Yt/5mELP+Xgiv/loAq/5R/Kf+TfSf/kXsm/495Jf+OdyT/jHYj/4t0Iv+JciH/h3Ag/4ZvH/+QdyD/hW4dxX9vHxAAAAAAAAAAAAAAAAAAAAAAu6pEHr2pQfHCrUL/tqI9/7WhPf+0nzz/sp07/7CbOv+vmjn/rZg4/6yXN/+qlTb/"
    s = s & "qJM1/6aRNP+lkDP/o44y/6GMMf+fijD/nokv/5yHLv+bhi3/mYQs/5eCK/+WgCr/lH8p/5N9J/+Reyb/j3kl/453JP+MdiP/i3Qi/4lyIf+HcCD/hW4f/5B3IP+Cahy6AAAAAAAAAAAAAAAAAAAAALyqQZHOuEf/uaVA/7ikP/+3oj7/taE9/7OfPP+znTr/tJs1/7OZ"
    s = s & "NP+xmDP/r5Yy/66UMf+skzD/qpEv/6mPLv+njS3/pYss/6OJK/+hiCr/oIYp/56EKP+cgyf/m4Em/5l/Jf+Wfib/kn0o/5F7Jv+PeSX/jnck/4x2I/+KdCL/iXIh/4dwIP+GcB//jHQg/4FrHUUAAAAAAAAAAL+fPxC+q0LlxLBE/7uoQf+6pkD/uKQ//7eiPv+1oDz/"
    s = s & "s588/6mfRf+nnkT/ppxD/6SaQv+jmUH/oZdA/6CWP/+elD7/nJI9/5uRPP+Zjzv/mI06/5aMOf+Uijj/k4k4/5GHN/+Phjb/lIEs/5V+KP+SfSf/kXsm/495Jf+OdyT/jHUi/4p0Iv+JciH/h3Ag/5B4If+CbB2kAAAAAAAAAAC9rEEfv61D/8KuRP+9qkL/vKhB/7qm"
    s = s & "QP+5pD//wqAy/3mqff8VuOj/FLXm/xS15f8UtOX/FLTl/xS05f8UtOX/E7Tl/xO05f8Ts+T/E7Pk/xOz5P8Ss+T/ErLk/xKy5P8SsuT/Ebbn/0uek/+ffR3/lH8o/5N9J/+Reyb/kHkl/453JP+MdiP/i3Qi/4lyIf+PdyL/hW4exwAAAAAAAAAAxbJINcazRvzDr0X/"
    s = s & "v6tD/72qQv+8qEH/uqZA/8aiMf9pro//ALj//wC0//8AtP//ALX//wC1//8Atf//ALX//wC1//8Atf//ALX//wC1//8Atf//ALX//wC1//8Atf//ALX//wC5//86pK//on8c/5aAKf+Ufyn/k30n/5F7Jv+PeSX/jnck/4x2I/+LdCL/jXYi/4dvH+aMchkUAAAAAMOw"
    s = s & "RWvVwEz/wq9F/8CtRP+/q0P/vapC/7yoQf/HpDL/ba6M/wC3//8AtPr/ALT6/wC0+v8AtPr/ALT6/wC0+v8AtPr/ALT6/wC0+v8AtPr/ALT6/wC0+v8AtPr/ALT6/wC0+v8AuP//PqOp/6OBHv+Ygir/loAq/5R/Kf+TfSf/kXsm/495Jf+OdyT/jHYj/4x1Iv+KciH/"
    s = s & "gm4bJQAAAADGskdn1cBM/8SxRv/Cr0X/wK1E/7+rQ/+9qkH/yaYz/26vjf8Atv//ALv//wDC//8Awf//AMH//wDB//8Awf//AMH//wDC//8AuP//ALT6/wC2/P8Atfv/ALT6/wC0+v8AtPr/ALj//z+kqf+lgh//mYQr/5eCK/+WgCr/lH8o/5J9J/+Reyb/j3kl/453"
    s = s & "JP+OdyP/inMh/IdwHiIAAAAAyLRHZ9fCTf/Gskb/xLBG/8KvRf/ArUT/v6tD/8qoNP9vsY7/AL7//xGJuv8iWnT/IGB5/yBgef8gYHn/IGB5/yBfef8iXHf/CprS/wC7//8Asf3/ALL5/wC2/P8AtPr/ALT6/wC4//8/par/p4Qg/5uFLP+ZhCz/l4Ir/5aAKv+Ufij/"
    s = s & "kn0n/5F7Jv+PeSX/kHkk/4t1IvyOeCUiAAAAAMi0SmfZxE//x7RI/8ayR//EsUb/wq9F/8GtRP/MqjX/cLGO/wC9//8Qibj/IVhy/x9eeP8fXnj/H154/x9eeP8fXnf/IVp2/wmi1f8WaIr/JVx8/xut7v8Ar/z/ALf9/wC0+v8AuP//QKWr/6mGIf+dhy3/m4Yt/5mE"
    s = s & "LP+Ygiv/loAq/5R/KP+TfSf/kXsm/5J7Jf+NdyP8jnglIgAAAADLuUpn28ZQ/8q1Sf/HtEj/xrJH/8SxRv/Cr0X/zqw3/3Gzj/8Atv//ALn//wC///8Avv//AL7//wC+//8Avv//AL7//wC///8Av///IGJ+/5OKs/+LhY//L3l4/wCy//8At/7/ALf//0Gmq/+qiCL/"
    s = s & "noku/5yHLv+bhi3/mYQs/5eCK/+WgCr/lH8p/5N9J/+TfSf/j3kk/I54JSIAAAAAzblKZ93IUf/Lt0r/ybVJ/8e0SP/Gskf/xLFG/9CuOP9ytI//ALP//2PM9v+65PL/tODy/7Tg8v+04PL/tODy/7Tg8v+04PL/tOH0/7nr+/+Df5j/UgKb/3c8JP8zgX//ALL//wC6"
    s = s & "//9Cp6v/rIkj/6CKL/+eiS//nIcu/5uGLf+ZhCz/l4Ir/5aAKv+Ufyn/lX8o/5B6JvyOeCUiAAAAAM28TGfeyVL/zblL/8u3Sv/JtUn/x7RI/8ayRv/Srzn/c7SQ/wCy//+d2/T///nu///07///9O////Tv///07///9O////Tv///18P//+/r/taRv/2EtMv9TCbf/"
    s = s & "eD0j/zOBfv8Atf//Qaqv/66LJP+hjDD/n4ow/56JL/+chy7/moUt/5mELP+Xgiv/loAq/5aAKf+SfCf8ln8lIgAAAADPvkxn4MtT/867S//MuUv/y7dK/8m1Sf/HtEj/07E5/3S1kf8Asv//m9r0///47v/78+//+/Pv//vz7//78+//+/Pv//vz7//78+///PTx////"
    s = s & "///Fr3L/bDAy/1UJt/93PSP/KIWO/zqmt/+wjyf/o44x/6GMMf+fijD/nokv/5yHLf+ahS3/mYQs/5eCK/+Ygir/k34o/JZ/JSIAAAAA0r5PZ+LNVP/QvE3/zrpL/8y5S//Lt0r/ybVJ/9WzO/91t5H/ALH//2/U/f/N9P//xfD//8Xw///F8P//xfD//8Xw///E8P//"
    s = s & "x+n3/8jl8f/K5/X/0PT//5Sic/9qMDL/Vgm3/2s+MP9wgkH/rJc2/6WQMv+jjjL/oYwx/5+KMP+diC7/nIct/5qFLf+ZhCz/mYQr/5WAKfyWfyUiAAAAANTBUWfk0Ff/0r9P/9C9Tv/Ou03/zbpM/8u4S//YtT3/driS/wC6//8HmNL/Fnum/xV+qv8Vfqr/FX6q/xV+"
    s = s & "qv8Vfqn/Fnyo/wij4f8Bt/z/AbP6/wG1/f8Iv///P4x1/3YzMv9UCbj/dkEn/5yCF/+qlTX/pZAz/6SPM/+ijTL/oIsx/56JMP+diC//m4Yu/5yHLv+Xgiz8loctIgAAAADXw1Rn5tJa/9TBU//SwFH/0b5Q/8+8T//Ouk//2rhB/3i5lf8Av///Cnik/xY3Sf8UP1D/"
    s = s & "FD9Q/xQ/UP8UP1D/FD5P/xc5Tf8Ai8T/ALb+/wCu+/8Arvr/ALD9/wC4//8xiHX/dzMx/1IKu/9tOCz/noco/7CiOP+mkjb/pZA2/6OONf+hjTT/oIsz/56KMv+fijL/moUv/J2HLSIAAAAA18ZUZ+jVXP/WxFb/1MJV/9PBVP/Rv1P/0L1R/927RP95u5b/ALP//zfH"
    s = s & "//9x3P//a9n//2vZ//9r2f//a9n//2vZ//9r2v//bdP7/27O9f9uz/X/bs/1/27P9f900vf/RND//zuLdP9tKjL/YjS0/5u20v+Ue3z/q5Uy/6iTOf+mkjj/pJA3/6KONv+hjTb/oY01/5yHMfydhy0iAAAAANnIVmfq117/2MZY/9fFWP/Vw1f/08FW/9K/Vf/fvUf/"
    s = s & "eryY/wCx//+c2vT///vv///18P//9fD///Xw///18P//9fD///Xw///17///9e////Xv///17///9e////ju/8no9f8AuP//UJF8/6Oy0f9lRen/SgDR/56BVv+unTb/qZQ8/6eTO/+lkTr/o485/6SPN/+diTH8nYctIgAAAADcyFZn7Nlf/9rIWf/Yx1r/18Za/9bE"
    s = s & "Wf/Uwlj/4cBL/3y+mv8Asf//ltn0///27//z8fD/8/Hw//Px8P/z8fD/8/Hw//Px8P/z8fD/8/Hw//Px8P/z8fD/8/Hw///07//C5PH/ALT//0rAvP+gd4P/RQDU/2sxqv+sl0b/r5w+/6yYP/+qlj7/qJU+/6WRO/+lkTf/n4oy/J2HLSIAAAAA3MtZZ+3aYP/bylr/"
    s = s & "2shb/9nHXP/Yxlz/1sRb/+PDTv99v5z/ALD//47Y9//9+vX/8vT1//L09f/y9PX/8vT1//L09f/y9PX/8/Lx//Px8P/z8fD/8/Hw//Px8P/99O//uuLy/wC0//9Ir7L/wp40/6CAZP+vmkr/tKJA/6+cRP+um0L/rJlB/6qWPv+nkzv/p5I4/6CMM/ydjjQiAAAAAN7N"
    s = s & "WWfv3GH/3Mtb/9vKXP/ayV3/2che/9jHXv/lxVH/f8Ge/wC1//8Xs/H/NrTk/zW05f81tOX/NbTl/zW05f81tOX/NbTl/zG88/8wv/n/ML74/zC++P8wvvj/Mb/4/x67+f8Atf//SrC2/8WkPv+8q0L/t6VG/7OgSP+yn0f/sJ1G/66bQv+rlz//qZU8/6mUOf+ijjT8"
    s = s & "pY40IgAAAADezVln8d1i/97NXP/dzF3/28te/9vJX//ayWD/6MhV/3/CoP8Av///FHGX/ykoNP8lMTj/JTE4/yUxOP8lMTj/JTA3/ygrN/8EiLv/ALr//wCx+v8Asvv/ALP8/wCy/P8As/z/ALf//0yzuP/Hp0L/uqZM/7ikTP+2o0v/taJJ/7KfR/+wnEP/rZlA/6qW"
    s = s & "Pf+qljr/pI81/KWONCIAAAAA4c9bZ/LfY//fzl3/3s1e/93MX//cy2D/28ph/+nKV/+Bw6H/ALX//wC0+f8CsvX/ArP1/wKz9f8Cs/X/ArP1/wKz9f8Csvb/ALX6/wC0+v8AtPr/ALP5/wCt8/8ArfP/AK3z/wCt+v9Krrb/yqpG/7ypUP+6p0//uaZO/7ajS/+0oEf/"
    s = s & "sZ1E/6+bQf+smD7/rJc7/6WRNvyljjQiAAAAAOPSW2f04WT/4dBe/+DPX//fzmD/3s1h/93MYv/rzFn/gsSi/wC5//8Km9P/FH6r/xOCrP8Tgqz/E4Ks/xOCrP8Tgqz/FICs/wak4v8AuPz/ALX7/wCf5f8AmuD/AJ7k/wCd4/8AnfT/Uqep/82uSf++q1P/vKpT/7qn"
    s = s & "UP+4pUz/taJJ/7OfRf+xnUP/rpo//66ZPf+nkzf8pZY0IgAAAADj0l5n9uNl/+LRX//h0GD/4M9h/9/OYv/ezmP/7c5a/4LFo/8Avf//Fnuj/y09S/8qRFD/KkRQ/ypEUP8qRFD/KkRP/y0/Tv8OkcP/ALz+/wC0+v8AnOL/AK/0/wC2/P8Atv//Krna/7azZP/GsFP/"
    s = s & "wK5W/76sU/+7qVD/uaZN/7ejSv+0oUb/sp5D/7CcQP+vmz7/qZU4/KyWNCIAAAAA5dRdbfrmZ//k0mD/49Jh/+LRYv/g0GP/389k/+7PW/+Dx6P/ALP//wC6//8Av///AL7//wC+//8Avv//AL7//wC+//8Av///ALj//wCz+v8AtPr/AJ3j/wCu9P8Atf//K7rb/762"
    s = s & "Yv/PtFL/wrNb/8GwV/+/rVT/vatR/7uoTv+4pUv/tqJH/7OgRP+xnUH/sZw//6uXOv6umDIjAAAAAObVXUry32T96NZi/+TTYv/j0mP/4tFk/+HQZf/w0Vz/g8el/wCz//8As/v/ALT7/wC0+/8AtPv/ALT7/wC0+/8AtPv/ALT7/wCz+/8As/v/ALP7/wCc4/8Arv//"
    s = s & "LLvc/8K6Zv/TuFb/xrZf/8W0W//DsVj/wK9V/76sUv+8qU//uadL/7ekSP+1oUX/sp5C/7SfP/+slzr1r5U9HQAAAADt22Ad5tVf/uvZY//m1WP/5dRk/+PTZP/j0mX/8dNc/4XIpf8Atf//ALX//wC1//8Atf//ALX//wC1//8Atf//ALX//wC2//8Atv//ALb//wC2"
    s = s & "//8AnfD/L7XX/8W/av/XvFv/yrtj/8m4X//GtVz/xLNZ/8KwVv/ArlP/vatP/7uoTP+5pUn/tqNH/7SgQ/+7pkL/r5o8ygAAAAAAAAAA6tZbGevZYfPv3WX/6Ndk/+fWZf/l1Gb/5NRn/+bTZf/Q0HX/pMuS/6PLk/+jypP/osqU/6LJlf+hyZX/oMeV/5/GlP+exZP/"
    s = s & "ncST/5zDkv+bwpL/mruP/8a/bv/ZwGD/zb5n/8y8ZP/KumD/yLdd/8a1Wf/Eslf/wbBU/7+tUf+9qk3/uqdJ/7ilR/+1oUT/v6lE/7CcPbcAAAAAAAAAAAAAAADs2mOx/uts/+nYZf/o12b/59Zn/+XVaP/j1Gr/6NRn/+/TYv/u02P/7dJl/+zRZv/r0Gf/6s9o/+nO"
    s = s & "aP/ozWf/5stm/+XJZv/kyGT/4sdk/+LGY//YxGj/0MJq/9DAZ//OvmT/y7th/8m5Xv/Htlr/xbRY/8OxVf/ArlL/vqxO/7ypS/+5pkj/t6NF/8SuR/+vmz9kAAAAAAAAAAAAAAAA6tliPv3rbP/w32j/6dhm/+nYZ//n12n/5tZq/+TUa//k1Gz/4tNt/+LSbv/h0W//"
    s = s & "4NFx/9/Pcf/ez3L/3c5z/9zNcv/ay3H/2clw/9jJb//Xx2//1cVt/9PDav/RwWj/z79l/828Yv/Kul//yLhb/8a1WP/Es1b/wrBS/7+tT/+9qkz/uqdI/8WwSv+2okPkuaJFCwAAAAAAAAAAAAAAAAAAAADr2mR3/+1v//Dfaf/q2Wj/6dhp/+fXav/m1mz/5dVt/+TU"
    s = s & "bf/j02//4tNw/+HScv/g0XL/39Bz/97QdP/dz3X/3M10/9vMc//ay3L/2Mlx/9fHb//VxGv/08Jp/9DAZv/PvmP/zLtf/8q5XP/Ht1n/xrRX/8OxU//Br1D/vqtM/8azTv/BrEnxuKdBOgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO3cZnf962///uxx//Dfbf/s3G3/"
    s = s & "6dlt/+bWbf/m1m7/5NVv/+PUcf/j1HL/4dNz/+HSdP/g0Xb/39B3/97Qd//dznf/3M11/9rLcv/YyG//1sZs/9TEav/TwWf/0L9j/869YP/Mul3/ybha/8m4WP/HtlX/yrdU/9K+Vf/Fsk3suahJOwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA69xoQvbk"
    s = s & "ba3x323u7d1t+/Picv/863j/+eh4//jnef/35nv/9eV8//Tkff/05H7/8+N///LigP/w4YH/8OCA/+7efv/s3Hv/6tl4/+jXdf/m1XP/5dJw/+LQbf/fzWn/3ctl/9zJY//Lulr/ybdW+8u4VN3Fsk+TualNIQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAA891uF/7vcyHq2W5K9eN0Z+bXb2fm13Fn5tdxZ+PUcWfj0nFn4dJxZ+HScWfez3Fn3s9vZ9zNb2fcy29n2ctsZ9fIbGfXxmpn1MZqZ9LDZ2fSwWVnz75gZ82+XmfUwV5nyrtbNdjIXCHMzGYKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAA=="
    IconB64_NettskjemaHenter = s
End Function

Private Function IconB64_Statistikkern() As String
    Dim s As String
    s = s & "Qk02GQAAAAAAADYAAAAoAAAAKAAAACgAAAABACAAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIyM9GCJifbviYn2/4eH9v+Fhfb/g4P1/4GB9f9/f/X/fX31/3t79f95efT/d3f0/3V19P9zc/T/cXH0/3Bw8/9ubvP/bGzz/2pq8/9oaPP/"
    s = s & "Zmby/2Rk8v9iYvL/YGDy/15e8v9cXPH/Wlrx/1hY8f9WVvH/VFTx/1JS8P9PT+/3Tk7weAAAAAAAAAAAAAAAAAAAAAAAAAAAj4/vEI+P+LePj/f/jY32/4uL9v+Jifb/h4f2/4WF9v+Dg/X/gYH1/39/9f99ffX/e3v1/3l59P93d/T/dXX0/3Nz9P9xcfT/cHDz/25u"
    s = s & "8/9sbPP/amrz/2ho8/9mZvL/ZGTy/2Ji8v9gYPL/Xl7y/1xc8f9aWvH/WFjx/1ZW8f9UVPH/UlLw/1BQ8P9OTu/HT0/vIAAAAAAAAAAAAAAAAJWV9Z+Tk/f/kZH3/4+P9/+Njfb/i4v2/4mJ9v+Hh/b/hYX2/4OD9f+BgfX/f3/1/3199f97e/X/eXn0/3d39P91dfT/"
    s = s & "c3P0/3Fx9P9wcPP/bm7z/2xs8/9qavP/aGjz/2Zm8v9kZPL/YmLy/2Bg8v9eXvL/XFzx/1pa8f9YWPH/Vlbx/1RU8f9SUvD/UFDw/05O78cAAAAAAAAAAJiY9EiXl/f/lZX3/5OT9/+Rkff/j4/3/42N9v+Li/b/iYn2/4eH9v+Fhfb/g4P1/4GB9f9/f/X/fX31/3t7"
    s = s & "9f95efT/d3f0/3V19P9zc/T/cXH0/3Bw8/9ubvP/bGzz/2pq8/9oaPP/Zmby/2Rk8v9iYvL/YGDy/15e8v9cXPH/Wlrx/1hY8f9WVvH/VFTx/1JS8P9QUPD/Tk7weAAAAACbm/XfmZn3/5eX9/+Vlff/k5P3/5GR9/+Pj/f/jY32/4uL9v+Jifb/h4f2/4WF9v+Dg/X/"
    s = s & "gYH1/39/9f99ffX/e3v1/3l59P93d/T/dXX0/3Nz9P9xcfT/cHDz/25u8/9sbPP/amrz/2ho8/9mZvL/ZGTy/2Ji8v9gYPL/Xl7y/1xc8f9aWvH/WFjx/1ZW8f9UVPH/UlLw/09P7/efn+8QnZ33/5ub9/+Zmff/l5f3/5WV9/+Tk/f/kZH3/4+P9/+Njfb/i4v2/4mJ"
    s = s & "9v+Hh/b/hYX2/4OD9f+BgfX/f3/1/3199f97e/X/eXn0/3d39P91dfT/c3P0/3Fx9P9wcPP/bm7z/2xs8/9qavP/aGjz/2Zm8v9kZPL/YmLy/2Bg8v9eXvL/XFzx/1pa8f9YWPH/Vlbx/1RU8f9SUvD/n5/0SJ+f9/+dnff/m5v3/5mZ9/+Xl/f/lZX3/5OT9/+Rkff/"
    s = s & "j4/3/42N9v+Li/b/iYn2/4eH9v+Fhfb/g4P1/4GB9f9/f/X/fX31/3t79f95efT/d3f0/3V19P9zc/T/cXH0/3Bw8/9ubvP/bGzz/2pq8/9oaPP/Zmby/2Rk8v9iYvL/YGDy/15e8v9cXPH/Wlrx/1hY8f9WVvH/VFTx/6Gh9nihoff/n5/3/52d9/+bm/f/mZn3/5eX"
    s = s & "9/+Vlff/k5P3/5GR9/+Pj/f/jY32/4uL9v+Jifb/h4f2/4WF9v+Dg/X/gYH1/39/9f99ffX/e3v1/3l59P93d/T/dXX0/3Nz9P9xcfT/cHDz/25u8/9sbPP/amrz/2ho8/9mZvL/ZGTy/2Ji8v9gYPL/Xl7y/1xc8f9aWvH/WFjx/1ZW8f+lpfeAo6P4/6Gh9/+fn/f/"
    s = s & "nZ33/5ub9/+Zmff/l5f3/5WV9/+Tk/f/kZH3/4+P9/+Njfb/i4v2/4mJ9v+Hh/b/hYX2/4OD9f+BgfX/f3/1/3199f97e/X/eXn0/3d39P91dfT/c3P0/3Fx9P9wcPP/bm7z/2xs8/9qavP/aGjz/2Zm8v9kZPL/YmLy/2Bg8v9eXvL/XFzx/1pa8f9YWPH/p6f3gKWl"
    s = s & "+P+jo/j/oaH3/5+f9/+dnff/m5v3/5mZ9/+kpPj/sLD5/66u+f+trfn/q6v5/6qq+P+Pj/b/iYn2/4eH9v+Njff/oqL4/6Gh+P+fn/j/np74/5yc+P+GhvX/d3f0/3V19P9zc/T/lZX3/5SU9v+Skvb/kZH2/4+P9v97e/X/Zmby/2Rk8v9iYvL/YGDy/15e8v9cXPH/"
    s = s & "Wlrx/6mp94Cnp/j/paX4/6Oj+P+hoff/n5/3/52d9/+bm/f/zMz7////////////////////////////m5v3/4uL9v+Jifb/paX4////////////////////////////ra35/3l59P93d/T/dXX0////////////////////////////tbX5/2ho8/9mZvL/ZGTy/2Ji"
    s = s & "8v9gYPL/Xl7y/1xc8f+rq/eAqan4/6en+P+lpfj/o6P4/6Gh9/+fn/f/nZ33/83N+////////////////////////////52d+P+Njfb/i4v2/6en+P///////////////////////////66u+f97e/X/eXn0/3d39P///////////////////////////7a2+f9qavP/"
    s = s & "aGjz/2Zm8v9kZPL/YmLy/2Bg8v9eXvL/ra35gKur+P+pqfj/p6f4/6Wl+P+jo/j/oaH3/5+f9//Ozvv///////////////////////////+fn/j/j4/3/42N9v+oqPj///////////////////////////+vr/n/fX31/3t79f95efT/////////////////////////"
    s = s & "//+3t/n/bGzz/2pq8/9oaPP/Zmby/2Rk8v9iYvL/YGDy/6+v+YCtrfn/q6v4/6mp+P+np/j/paX4/6Oj+P+hoff/z8/7////////////////////////////oaH4/5GR9/+Pj/f/qqr4////////////////////////////sLD5/39/9f99ffX/e3v1////////////"
    s = s & "////////////////uLj5/25u8/9sbPP/amrz/2ho8/9mZvL/ZGTy/2Ji8v+xsfmAr6/5/62t+f+rq/j/qan4/6en+P+lpfj/o6P4/9DQ+////////////////////////////6Ki+P+Tk/f/kZH3/6ur+f///////////////////////////7Ky+f+BgfX/f3/1/319"
    s = s & "9f///////////////////////////7i4+v9wcPP/bm7z/2xs8/9qavP/aGjz/2Zm8v9kZPL/s7P5gLGx+f+vr/n/ra35/6ur+P+pqfj/p6f4/6Wl+P/R0fz///////////////////////////+kpPj/lZX3/5OT9/+trfn///////////////////////////+zs/n/"
    s = s & "g4P1/4GB9f9/f/X///////////////////////////+5ufr/cXH0/3Bw8/9ubvP/bGzz/2pq8/9oaPP/Zmby/7W1+YCzs/n/sbH5/6+v+f+trfn/q6v4/6mp+P+np/j/0tL8////////////////////////////pqb4/5eX9/+Vlff/rq75////////////////////"
    s = s & "////////tLT5/4WF9v+Dg/X/gYH1////////////////////////////urr6/3Nz9P9xcfT/cHDz/25u8/9sbPP/amrz/2ho8/+3t/mAtbX5/7Oz+f+xsfn/r6/5/62t+f+rq/j/qan4/9PT/P///////////////////////////6io+P+Zmff/l5f3/7Cw+f//////"
    s = s & "/////////////////////7W1+f+Hh/b/hYX2/4OD9f///////////////////////////7u7+v91dfT/c3P0/3Fx9P9wcPP/bm7z/2xs8/9qavP/ubn5gLe3+v+1tfn/s7P5/7Gx+f+vr/n/ra35/6ur+P/U1Pz///////////////////////////+pqfj/m5v3/5mZ"
    s = s & "9/+xsfn///////////////////////////+3t/n/iYn2/4eH9v+Fhfb///////////////////////////+8vPr/d3f0/3V19P9zc/T/cXH0/3Bw8/9ubvP/bGzz/7u7+YC5ufr/t7f6/7W1+f+zs/n/sbH5/6+v+f+trfn/zc37////////////////////////////"
    s = s & "qKj4/52d9/+bm/f/s7P5////////////////////////////uLj5/4uL9v+Jifb/h4f2////////////////////////////vb36/3l59P93d/T/dXX0/3Nz9P9xcfT/cHDz/25u8/+9vfmAu7v6/7m5+v+3t/r/tbX5/7Oz+f+xsfn/r6/5/7Cw+f/i4v3/////////"
    s = s & "///5+f//0dH8/6Gh9/+fn/f/nZ33/7S0+f///////////////////////////7m5+v+Njfb/i4v2/4mJ9v///////////////////////////76++v97e/X/eXn0/3d39P91dfT/c3P0/3Fx9P9wcPP/v7/5gL29+v+7u/r/ubn6/7e3+v+1tfn/s7P5/7Gx+f+vr/n/"
    s = s & "ra35/6ur+P+pqfj/p6f4/6Wl+P+jo/j/oaH3/5+f9/+2tvn///////////////////////////+6uvr/j4/3/42N9v+Li/b///////////////////////////+/v/r/fX31/3t79f95efT/d3f0/3V19P9zc/T/cXH0/7+/+4C/v/r/vb36/7u7+v+5ufr/t7f6/7W1"
    s = s & "+f+zs/n/sbH5/6+v+f+trfn/q6v4/6mp+P+np/j/paX4/6Oj+P+hoff/t7f5////////////////////////////vLz6/5GR9/+Pj/f/jY32////////////////////////////wMD6/39/9f99ffX/e3v1/3l59P93d/T/dXX0/3Nz9P/BwfuAwMD7/7+/+v+9vfr/"
    s = s & "u7v6/7m5+v+3t/r/tbX5/7Oz+f+xsfn/r6/5/62t+f+rq/j/qan4/6en+P+lpfj/o6P4/7m5+f///////////////////////////729+v+Tk/f/kZH3/4+P9////////////////////////////8HB+v+BgfX/f3/1/3199f97e/X/eXn0/3d39P91dfT/w8P7gMLC"
    s = s & "+//AwPv/v7/6/729+v+7u/r/ubn6/7e3+v+1tfn/s7P5/7Gx+f+vr/n/ra35/6ur+P+pqfj/p6f4/6Wl+P+6uvr///////////////////////////++vvr/lZX3/5OT9/+Rkff////////////////////////////Cwvv/g4P1/4GB9f9/f/X/fX31/3t79f95efT/"
    s = s & "d3f0/8XF+4DExPv/wsL7/8DA+/+/v/r/vb36/7u7+v+5ufr/t7f6/7W1+f+zs/n/sbH5/6+v+f+trfn/q6v4/6mp+P+np/j/vLz6////////////////////////////v7/6/5eX9/+Vlff/k5P3////////////////////////////w8P7/4WF9v+Dg/X/gYH1/39/"
    s = s & "9f99ffX/e3v1/3l59P/Hx/uAxsb7/8TE+//Cwvv/wMD7/7+/+v+9vfr/u7v6/7m5+v+3t/r/tbX5/7Oz+f+xsfn/r6/5/62t+f+rq/j/qan4/7i4+f///////////////////////////76++v+Zmff/l5f3/5WV9////////////////////////////8TE+/+Hh/b/"
    s = s & "hYX2/4OD9f+BgfX/f3/1/3199f97e/X/ycn7gMjI+//Gxvv/xMT7/8LC+//AwPv/v7/6/729+v+7u/r/ubn6/7e3+v+1tfn/s7P5/7Gx+f+vr/n/ra35/6ur+P+pqfj/5ub9//////////////////Pz/v+goPf/m5v3/5mZ9/+Xl/f/////////////////////////"
    s = s & "///Fxfv/iYn2/4eH9v+Fhfb/g4P1/4GB9f9/f/X/fX31/8vL+4DKyvz/yMj7/8bG+//ExPv/wsL7/8DA+/+/v/r/vb36/7u7+v+5ufr/t7f6/7W1+f+zs/n/sbH5/6+v+f+trfn/q6v4/6mp+P+1tfn/vLz6/7S0+f+hoff/n5/3/52d9/+bm/f/mZn3////////////"
    s = s & "////////////////xsb7/4uL9v+Jifb/h4f2/4WF9v+Dg/X/gYH1/39/9f/NzfuAzMz8/8rK/P/IyPv/xsb7/8TE+//Cwvv/wMD7/7+/+v+9vfr/u7v6/7m5+v+3t/r/tbX5/7Oz+f+xsfn/r6/5/62t+f+rq/j/qan4/6en+P+lpfj/o6P4/6Gh9/+fn/f/nZ33/5ub"
    s = s & "9////////////////////////////8fH+/+Njfb/i4v2/4mJ9v+Hh/b/hYX2/4OD9f+BgfX/z8/7gM7O/P/MzPz/ysr8/8jI+//Gxvv/xMT7/8LC+//AwPv/v7/6/729+v+7u/r/ubn6/7e3+v+1tfn/s7P5/7Gx+f+vr/n/ra35/6ur+P+pqfj/p6f4/6Wl+P+jo/j/"
    s = s & "oaH3/5+f9/+dnff////////////////////////////IyPv/j4/3/42N9v+Li/b/iYn2/4eH9v+Fhfb/g4P1/9HR+nDQ0Pz/zs78/8zM/P/Kyvz/yMj7/8bG+//ExPv/wsL7/8DA+/+/v/r/vb36/7u7+v+5ufr/t7f6/7W1+f+zs/n/sbH5/6+v+f+trfn/q6v4/6mp"
    s = s & "+P+np/j/paX4/6Oj+P+hoff/n5/3////////////////////////////ycn7/5GR9/+Pj/f/jY32/4uL9v+Jifb/h4f2/4WF9v/R0fo40tL8/9DQ/P/Ozvz/zMz8/8rK/P/IyPv/xsb7/8TE+//Cwvv/wMD7/7+/+v+9vfr/u7v6/7m5+v+3t/r/tbX5/7Oz+f+xsfn/"
    s = s & "r6/5/62t+f+rq/j/qan4/6en+P+lpfj/o6P4/6Gh9////////////////////////////8rK+/+Tk/f/kZH3/4+P9/+Njfb/i4v2/4mJ9v+Hh/b/39//CNTU/f/S0vz/0ND8/87O/P/MzPz/ysr8/8jI+//Gxvv/xMT7/8LC+//AwPv/v7/6/729+v+7u/r/ubn6/7e3"
    s = s & "+v+1tfn/s7P5/7Gx+f+vr/n/ra35/6ur+P+pqfj/p6f4/6Wl+P+jo/j/7e39//////////////////////+7u/r/lZX3/5OT9/+Rkff/j4/3/42N9v+Li/b/iYn2/wAAAADV1f2/1NT9/9LS/P/Q0Pz/zs78/8zM/P/Kyvz/yMj7/8bG+//ExPv/wsL7/8DA+/+/v/r/"
    s = s & "vb36/7u7+v+5ufr/t7f6/7W1+f+zs/n/sbH5/6+v+f+trfn/q6v4/6mp+P+np/j/paX4/6ys+f/Nzfv/z8/7/87O+/++vvr/mZn3/5eX9/+Vlff/k5P3/5GR9/+Pj/f/jY32/4uL9N8AAAAA2Nj/KNTU+/fU1P3/0tL8/9DQ/P/Ozvz/zMz8/8rK/P/IyPv/xsb7/8TE"
    s = s & "+//Cwvv/wMD7/7+/+v+9vfr/u7v6/7m5+v+3t/r/tbX5/7Oz+f+xsfn/r6/5/62t+f+rq/j/qan4/6en+P+lpfj/o6P4/6Gh9/+fn/f/nZ33/5ub9/+Zmff/l5f3/5WV9/+Tk/f/kZH3/4+P9/+Li/dAAAAAAAAAAADX1/xo1tb9/9TU/f/S0vz/0ND8/87O/P/MzPz/"
    s = s & "ysr8/8jI+//Gxvv/xMT7/8LC+//AwPv/v7/6/729+v+7u/r/ubn6/7e3+v+1tfn/s7P5/7Gx+f+vr/n/ra35/6ur+P+pqfj/p6f4/6Wl+P+jo/j/oaH3/5+f9/+dnff/m5v3/5mZ9/+Xl/f/lZX3/5OT9/+RkfiXAAAAAAAAAAAAAAAAAAAAANfX/GDU1Pv31NT9/9LS"
    s = s & "/P/Q0Pz/zs78/8zM/P/Kyvz/yMj7/8bG+//ExPv/wsL7/8DA+/+/v/r/vb36/7u7+v+5ufr/t7f6/7W1+f+zs/n/sbH5/6+v+f+trfn/q6v4/6mp+P+np/j/paX4/6Oj+P+hoff/n5/3/52d9/+bm/f/mZn3/5eX9/+VlfeHn5//CAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "2Nj/KNbW/a/U1P3/0tL8/9DQ/P/Ozvz/zMz8/8rK/P/IyPv/xsb7/8TE+//Cwvv/wMD7/7+/+v+9vfr/u7v6/7m5+v+3t/r/tbX5/7Oz+f+xsfn/r6/5/62t+f+rq/j/qan4/6en+P+lpfj/o6P4/6Gh9/+fn/f/nZ33/5qa+LeWlvU4AAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAANTU+TDR0fxgz8/7gM3N+4DLy/uAycn7gMfH+4DFxfuAw8P7gMHB+4C/v/uAv7/5gL29+YC7u/mAubn5gLe3+YC1tfmAs7P5gLGx+YCvr/mAra35gKur94CpqfeAp6f3gKWl94Ciovdgn5/0MAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAA=="
    IconB64_Statistikkern = s
End Function
'@

$classCode = @'
Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Public OnActionNavn As String

Private Sub Btn_Click()
    ' Lagre navnet i en lokal variabel og lukk MENYEN FOeR makroen kjoeres -
    ' ikke etter. Mange makroer (f.eks. Kolonnevelger) Aapner sitt eget,
    ' langvarige vindu naar de kjoeres - hadde vi ventet med Unload til
    ' ETTER Application.Run (som da forst returnerer naar BRUKEREN lukker
    ' det andre vinduet), ville Makromeny blitt staaende synlig i
    ' bakgrunnen hele den tiden.
    Dim Navn As String
    Navn = OnActionNavn
    Unload frmMakroMeny
    KjorValgtMakro Navn
End Sub
'@

$formCodeRaw = @'
Option Explicit

Private mTiles As Collection

Private Const TILE_W As Single = 150
Private Const TILE_H As Single = 90
Private Const TILE_GAP As Single = 10
Private Const KOLONNER As Long = 2

Private Sub UserForm_Initialize()
    Me.Caption = "Makromeny"
    ByggMeny
End Sub

Private Sub ByggMeny()
    Dim Komponentnavn As String
    Dim OnActionNavn As String
    Dim Visningsnavn As String
    Dim Beskrivelse As String
    Dim Farge As Long
    Dim FunnedeMakroer As New Collection
    Dim i As Long
    Dim TilgangOK As Boolean

    If Not mTiles Is Nothing Then
        For i = mTiles.Count To 1 Step -1
            frameMeny.Controls.Remove mTiles(i).Btn.Name
        Next i
    End If
    Set mTiles = New Collection

    ' Deteksjon skjer via VBA-komponentnavn (se HentKjentMakroInfo), IKKE via
    ' knapper i arket. Krever "Trust access to the VBA project object
    ' model" - sjekk tilgangen samlet her forst for en tydelig melding i
    ' stedet for at hvert enkelt oppslag bare stille feiler.
    TilgangOK = HarVBAProjektTilgang()

    If TilgangOK Then
        i = 1
        Do While HentKjentMakroInfo(i, Komponentnavn, OnActionNavn, Visningsnavn, Beskrivelse, Farge)
            If VBAKomponentFinnes(Komponentnavn) Then
                FunnedeMakroer.Add Array(OnActionNavn, Visningsnavn, Beskrivelse, Farge)
            End If
            i = i + 1
        Loop
    End If

    ' Kun de to varselsituasjonene vises - ingen "X makro(er) funnet"-tekst
    ' i det vanlige tilfellet (Haakons eksplisitte onske 2026-09-23).
    If Not TilgangOK Then
        lblStatus.Caption = "Krever 'Trust access to the VBA project object model' i Excel sitt Klareringssenter."
    ElseIf FunnedeMakroer.Count = 0 Then
        lblStatus.Caption = "Fant ingen andre installerte makroer i denne fila."
    Else
        lblStatus.Caption = ""
    End If

    Dim idx As Long
    Dim r As Long, c As Long
    Dim info As Variant
    Dim tileHandler As clsMakroTile
    Dim newBtn As MSForms.CommandButton
    Dim ico As IPictureDisp

    For idx = 1 To FunnedeMakroer.Count
        info = FunnedeMakroer(idx)
        r = (idx - 1) \ KOLONNER
        c = (idx - 1) Mod KOLONNER

        Set newBtn = frameMeny.Controls.Add("Forms.CommandButton.1", "tileBtn" & idx, True)
        With newBtn
            .Left = c * (TILE_W + TILE_GAP)
            .Top = r * (TILE_H + TILE_GAP)
            .Width = TILE_W
            .Height = TILE_H
            .Caption = info(1)
            .BackColor = info(3)
            .ForeColor = RGB(35, 38, 46)
            .Font.Size = 11
            .Font.Bold = True
            .ControlTipText = info(2)
        End With

        Set ico = HentMakroIkon(CStr(info(0)))
        If Not ico Is Nothing Then
            Set newBtn.Picture = ico
            newBtn.PicturePosition = fmPicturePositionAboveCenter
        End If

        Set tileHandler = New clsMakroTile
        Set tileHandler.Btn = newBtn
        tileHandler.OnActionNavn = CStr(info(0))
        mTiles.Add tileHandler
    Next idx

    Dim totalRows As Long
    If FunnedeMakroer.Count = 0 Then
        totalRows = 0
    Else
        totalRows = Int((FunnedeMakroer.Count - 1) / KOLONNER) + 1
    End If

    Dim neededHeight As Single
    neededHeight = totalRows * (TILE_H + TILE_GAP)

    If neededHeight > 280 Then
        frameMeny.ScrollBars = fmScrollBarsVertical
        frameMeny.ScrollHeight = neededHeight
        frameMeny.Height = 280
    Else
        frameMeny.ScrollBars = fmScrollBarsNone
        frameMeny.Height = IIf(neededHeight > 0, neededHeight, 40)
    End If

    PosisjonerBunnKontroller
End Sub

Private Sub PosisjonerBunnKontroller()
    Dim topp As Single
    topp = frameMeny.Top + frameMeny.Height + 10
    lblStatus.Top = topp
    cmdInfo.Top = topp + 20
    Me.Height = cmdInfo.Top + cmdInfo.Height + 40
End Sub

Private Sub cmdInfo_Click()
    frmMakroInfo.Show
End Sub
'@

$moduleCode = $moduleCodeRaw -replace '\{AA\}', $aa -replace '\{OE\}', $oe -replace '\{AE\}', $ae
$formCode = $formCodeRaw -replace '\{AA\}', $aa -replace '\{OE\}', $oe -replace '\{AE\}', $ae

# ---- Innebygd VBA-kildekode: frmMakroInfo (info-vindu, apnes via "Installert
# info..."-knappen i frmMakroMeny) - viser VBA-komponenter og innholdet i
# alle "very hidden"-ark i arbeidsboken, sa Haakon kan se hva EMI-makroene
# faktisk har lagt til bak kulissene. Rent lesevindu, ingen skriving. ----
$infoFormCodeRaw = @'
Option Explicit

' Holder klasseinstansene for vis/skjul-knappene i live (WithEvents-handlere
' slutter aa fyre hvis eneste referanse forsvinner) - samme monster som
' mTiles i frmMakroMeny/clsMakroTile.
Private mArkTogglere As Collection

Private Sub UserForm_Initialize()
    Me.Caption = "Installert info"
    lblKomponenterHeader.Font.Bold = True
    lblArkHeader.Font.Bold = True
    lstKomponenter.Font.Size = 8
    FyllKomponentliste
    FyllSkjulteArk
End Sub

Private Sub cmdLukk_Click()
    Unload Me
End Sub

Private Sub FyllKomponentliste()
    Dim comp As Object, typeTekst As String
    lstKomponenter.Clear

    Dim rader As New Collection
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1: typeTekst = "Modul"
            Case 2: typeTekst = "Klasse"
            Case 3: typeTekst = "Skjema"
            Case Else: typeTekst = ""
        End Select
        If Len(typeTekst) > 0 Then rader.Add typeTekst & "|" & comp.Name
    Next comp

    Dim n As Long: n = rader.Count
    If n = 0 Then
        lblKomponenterHeader.Caption = "VBA-komponenter (fant ingen - mangler VBA-tilgang?)"
        Exit Sub
    End If

    Dim arr() As String, a As Long, b As Long, tmp As String
    ReDim arr(1 To n)
    For a = 1 To n: arr(a) = rader(a): Next a
    For a = 1 To n - 1
        For b = a + 1 To n
            If arr(b) < arr(a) Then tmp = arr(a): arr(a) = arr(b): arr(b) = tmp
        Next b
    Next a

    For a = 1 To n
        Dim deler() As String: deler = Split(arr(a), "|")
        lstKomponenter.AddItem deler(0)
        lstKomponenter.List(lstKomponenter.ListCount - 1, 1) = deler(1)
    Next a
    lblKomponenterHeader.Caption = "VBA-komponenter (" & n & ")"
End Sub

Private Sub FyllSkjulteArk()
    Dim i As Long
    For i = frameArk.Controls.Count - 1 To 0 Step -1
        frameArk.Controls.Remove frameArk.Controls(i).Name
    Next i
    Set mArkTogglere = New Collection

    Dim ws As Worksheet
    Dim topp As Single: topp = 5
    Dim funnet As Long: funnet = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVeryHidden Then
            funnet = funnet + 1
            topp = LeggTilArkSeksjon(ws, topp)
        End If
    Next ws

    If funnet = 0 Then
        Dim lblIngen As MSForms.Label
        Set lblIngen = frameArk.Controls.Add("Forms.Label.1", "lblIngenArk", True)
        lblIngen.Left = 5: lblIngen.Top = 5: lblIngen.Width = 370: lblIngen.Height = 18
        lblIngen.Font.Italic = True
        lblIngen.Caption = "Ingen skjulte innstillingsark funnet i denne arbeidsboken."
        topp = 28
    End If

    If topp > 220 Then
        frameArk.ScrollBars = fmScrollBarsVertical
        frameArk.ScrollHeight = topp
        frameArk.Height = 220
    Else
        frameArk.ScrollBars = fmScrollBarsNone
        frameArk.Height = IIf(topp > 0, topp, 30)
    End If

    lblArkHeader.Caption = "Skjulte innstillingsark (very hidden) - " & funnet & " funnet"
    PosisjonerBunn
End Sub

' Bygger EN seksjon (arknavn + innhold) i frameArk fra Top=startTop, og
' returnerer Top-posisjonen rett etter seksjonen.
Private Function LeggTilArkSeksjon(ByVal ws As Worksheet, ByVal startTop As Single) As Single
    Dim topp As Single: topp = startTop

    Dim lblNavn As MSForms.Label
    Set lblNavn = frameArk.Controls.Add("Forms.Label.1", "lblArkNavn" & (frameArk.Controls.Count + 1), True)
    lblNavn.Left = 5: lblNavn.Top = topp: lblNavn.Width = 290: lblNavn.Height = 16
    lblNavn.Font.Bold = True
    lblNavn.BackColor = RGB(226, 232, 240)
    lblNavn.Caption = " " & ws.Name

    ' Vis/skjul-knapp for aa kunne aapne det skjulte arket rett i Excel for
    ' aa se paa det direkte, og skjule det igjen etterpaa - Haakons
    ' eksplisitte onske 2026-09-22. clsArkToggle holder styr paa hvilket
    ' ark KNAPPEN gjelder og bytter Visible-status ved klikk.
    Dim btnToggle As MSForms.CommandButton
    Set btnToggle = frameArk.Controls.Add("Forms.CommandButton.1", "btnArkToggle" & (frameArk.Controls.Count + 1), True)
    btnToggle.Left = 300: btnToggle.Top = topp: btnToggle.Width = 77: btnToggle.Height = 16
    btnToggle.Caption = "Vis ark"

    Dim toggler As clsArkToggle
    Set toggler = New clsArkToggle
    Set toggler.Btn = btnToggle
    toggler.ArkNavn = ws.Name
    mArkTogglere.Add toggler

    topp = topp + 18

    Dim brukt As Range
    On Error Resume Next
    Set brukt = ws.UsedRange
    On Error GoTo 0
    If brukt Is Nothing Then
        topp = topp + LeggTilInfoLinje(topp, "(tomt ark)")
        LeggTilArkSeksjon = topp + 6
        Exit Function
    End If

    Dim rader As Long, kolonner As Long
    rader = brukt.Rows.Count
    kolonner = brukt.Columns.Count

    ' Nokkel/Verdi-format (kolonne A/B) er moensteret alle fire makroenes
    ' konfigfunksjoner bruker - se KtsSkrivKonfigVerdi/SkrivKonfigVerdi i de
    ' respektive kildefilene. Cache-/registerark hopper bevisst over KV-
    ' visning (JSON/raadata som ikke er lesbart som nokkel/verdi-rader).
    If kolonner = 2 And rader <= 30 And InStr(1, ws.Name, "Cache", vbTextCompare) = 0 Then
        Dim r As Long, nokkel As String, verdi As String
        For r = 1 To rader
            nokkel = CStr(brukt.Cells(r, 1).Value)
            verdi = CStr(brukt.Cells(r, 2).Value)
            If Len(nokkel) > 0 Or Len(verdi) > 0 Then
                topp = topp + LeggTilNokkelVerdiLinje(topp, nokkel, verdi)
            End If
        Next r
    Else
        topp = topp + LeggTilInfoLinje(topp, rader & " rad(er) x " & kolonner & " kolonne(r) (ikke nokkel/verdi-format)")
    End If

    LeggTilArkSeksjon = topp + 8
End Function

Private Function LeggTilNokkelVerdiLinje(ByVal topp As Single, ByVal nokkel As String, ByVal verdi As String) As Single
    Dim lblN As MSForms.Label, lblV As MSForms.Label
    Set lblN = frameArk.Controls.Add("Forms.Label.1", "lblN" & (frameArk.Controls.Count + 1), True)
    lblN.Left = 12: lblN.Top = topp: lblN.Width = 140: lblN.Height = 14
    lblN.Font.Size = 8
    lblN.Caption = nokkel
    Set lblV = frameArk.Controls.Add("Forms.Label.1", "lblV" & (frameArk.Controls.Count + 1), True)
    lblV.Left = 155: lblV.Top = topp: lblV.Width = 220: lblV.Height = 14
    lblV.Font.Size = 8
    lblV.ForeColor = RGB(71, 85, 105)
    lblV.Caption = verdi
    LeggTilNokkelVerdiLinje = 15
End Function

Private Function LeggTilInfoLinje(ByVal topp As Single, ByVal tekst As String) As Single
    Dim lbl As MSForms.Label
    Set lbl = frameArk.Controls.Add("Forms.Label.1", "lblInfo" & (frameArk.Controls.Count + 1), True)
    lbl.Left = 12: lbl.Top = topp: lbl.Width = 360: lbl.Height = 14
    lbl.Font.Size = 8
    lbl.Font.Italic = True
    lbl.Caption = tekst
    LeggTilInfoLinje = 16
End Function

Private Sub PosisjonerBunn()
    Dim topp As Single
    topp = frameArk.Top + frameArk.Height + 10
    cmdLukk.Top = topp
    Me.Height = cmdLukk.Top + cmdLukk.Height + 40
End Sub
'@
$infoFormCode = $infoFormCodeRaw -replace '\{AA\}', $aa -replace '\{OE\}', $oe -replace '\{AE\}', $ae

# ---- Innebygd VBA-kildekode: clsArkToggle (WithEvents-handler for hver
# vis/skjul-knapp i frmMakroInfo - samme monster som clsMakroTile over) ----
$arkToggleCodeRaw = @'
Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Public ArkNavn As String

Private Sub Btn_Click()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ArkNavn)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If ws.Visible = xlSheetVisible Then
        ws.Visible = xlSheetVeryHidden
        Btn.Caption = "Vis ark"
    Else
        ws.Visible = xlSheetVisible
        ws.Activate
        Btn.Caption = "Skjul ark"
    End If
End Sub
'@

function Get-MenuVersion {
    param([string]$Code)
    if ($Code -match 'MAKROMENY_VERSION\s*As\s*String\s*=\s*"([^"]+)"') {
        return $Matches[1]
    }
    return $null
}

# ---- Opprydding av den nedlagte "Ark-knapp"-integrasjonen (2.4.x og
# tidligere satte inn en liten, tydelig MERKET kodeblokk i ThisWorkbook -
# funksjonen er fjernet helt 2026-09-23, men merkene brukes fortsatt til A
# FINNE OG FJERNE en evt. gammel blokk fra en fil installert med en eldre
# versjon, se Remove-ArkKnappFraThisWorkbook under). Ingen kode setter
# lenger NOE inn i ThisWorkbook - kun opprydding gjenstar.
$arkKnappMarkerStart = "'===MAKROMENY_ARKKNAPP_START==="
$arkKnappMarkerSlutt = "'===MAKROMENY_ARKKNAPP_SLUTT==="

function Remove-ArkKnappFraThisWorkbook {
    param($ThisWbComponent)
    $antall = $ThisWbComponent.CodeModule.CountOfLines
    if ($antall -eq 0) { return }
    $tekst = $ThisWbComponent.CodeModule.Lines(1, $antall)
    $rader = $tekst -split "`r`n"

    $startLinje = -1
    $sluttLinje = -1
    for ($i = 0; $i -lt $rader.Count; $i++) {
        if ($rader[$i] -eq $arkKnappMarkerStart) { $startLinje = $i }
        if ($rader[$i] -eq $arkKnappMarkerSlutt) { $sluttLinje = $i }
    }
    if ($startLinje -ge 0 -and $sluttLinje -ge $startLinje) {
        $ThisWbComponent.CodeModule.DeleteLines($startLinje + 1, ($sluttLinje - $startLinje + 1))
        Write-Output "  'Vis som fane'-integrasjon fjernet fra ThisWorkbook"
    }
}

$sourceVersion = Get-MenuVersion $moduleCode

# ---- 1. Finn maalfilen ----

if (-not $Path) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel-filer (*.xlsm)|*.xlsm|Alle filer (*.*)|*.*"
    $dlg.Title = "Velg Excel-fil som skal faa Makromeny"
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
    Write-Output "FEIL: Faar ikke tilgang til VBA-prosjektet."
    Write-Output "Sjekk at 'Klarer tilgang til VBA-prosjektobjektmodellen' er huket av i Excel sitt Klareringssenter (Filer > Alternativer > Klareringssenter > Innstillinger for klareringssenteret > Makroinnstillinger)."
    exit 1
}

if ($Uninstall) {
    Write-Output "Fjerner Makromeny fra $($wb.Name) ..."
    foreach ($navn in @("modMakroMeny", "clsMakroTile", "frmMakroMeny", "frmMakroInfo", "clsArkToggle")) {
        try {
            $eksisterende = $vbproj.VBComponents.Item($navn)
            $vbproj.VBComponents.Remove($eksisterende)
            Write-Output "  Fjernet $navn"
        } catch {}
    }
    try {
        foreach ($ws in $wb.Worksheets) {
            foreach ($btn in @($ws.Buttons())) {
                if ($btn.OnAction -eq "VisMakroMeny" -or $btn.OnAction -like "*!VisMakroMeny") { $btn.Delete() }
            }
        }
    } catch {}
    try {
        $thisWbComp = $vbproj.VBComponents.Item("ThisWorkbook")
        Remove-ArkKnappFraThisWorkbook $thisWbComp
    } catch {}
    try {
        if ($wb.Worksheets | Where-Object { $_.Name -eq "Makromeny" }) {
            $wb.Worksheets.Item("Makromeny").Delete()
            Write-Output "  Fjernet 'Makromeny'-fanearket"
        }
    } catch {}
    try {
        # En evt. gammel baandknapp (Application.CommandBars, funksjonen er
        # fjernet - se historikken) er en global, per-PC-innstilling - fjern
        # den KUN hvis den faktisk peker til DENNE fila, ellers kan den
        # tilhore en annen arbeidsbok som fortsatt bruker Makromeny.
        $cb = $excel.CommandBars | Where-Object { $_.Name -eq "Makromeny" }
        if ($cb -and $cb.Controls.Count -gt 0 -and $cb.Controls.Item(1).OnAction -like "*$($wb.Name)*") {
            $cb.Delete()
            Write-Output "  Fjernet baandknappen (pekte til denne fila)"
        }
    } catch {}
    try { $wb.Save() } catch {
        Write-Output "FEIL ved lagring: $($_.Exception.Message)"
        exit 1
    }
    Write-Output "Ferdig - Makromeny er fjernet fra $($wb.Name)."
    exit 0
}

# ---- 3. Sjekk hva som allerede er installert ----

$existingModule = $null
foreach ($comp in $vbproj.VBComponents) {
    if ($comp.Name -eq "modMakroMeny") { $existingModule = $comp }
}

$installedVersion = $null
if ($existingModule) {
    $existingCode = $existingModule.CodeModule.Lines(1, $existingModule.CodeModule.CountOfLines)
    $installedVersion = Get-MenuVersion $existingCode
}

$isFreshInstall = ($null -eq $existingModule)

if (-not $isFreshInstall -and $installedVersion -eq $sourceVersion -and -not $Force) {
    Write-Output "Makromeny er allerede installert og oppdatert (versjon $installedVersion) i $($wb.Name)."
    exit 0
}

if ($isFreshInstall) {
    Write-Output "Installerer Makromeny (versjon $sourceVersion) i $($wb.Name) ..."
} else {
    Write-Output "Oppdaterer Makromeny: versjon $installedVersion -> $sourceVersion i $($wb.Name) ..."
}

# ---- 4. Bygg komponentene i en midlertidig arbeidsbok, eksporter til en temp-mappe ----

$tempDir = Join-Path $env:TEMP ("Makromeny_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir | Out-Null

try {
    $buildWb = $excel.Workbooks.Add()
    $buildProj = $buildWb.VBProject

    $mod = $buildProj.VBComponents.Add(1)
    $mod.Name = "modMakroMeny"
    $mod.CodeModule.AddFromString($moduleCode)

    $cls = $buildProj.VBComponents.Add(2)
    $cls.Name = "clsMakroTile"
    $cls.CodeModule.AddFromString($classCode)

    $frmComp = $buildProj.VBComponents.Add(3)
    $frmComp.Name = "frmMakroMeny"
    $frmComp.Properties("Width").Value = 360
    $frmComp.Properties("Height").Value = 160
    $frmComp.Properties("Caption").Value = "Makromeny"
    $designer = $frmComp.Designer

    $frameMeny = $designer.Controls.Add("Forms.Frame.1", "frameMeny", $true)
    $frameMeny.Left = 10; $frameMeny.Top = 10; $frameMeny.Width = 310; $frameMeny.Height = 40
    $frameMeny.Caption = ""
    $frameMeny.BorderStyle = 0

    $lblStatus = $designer.Controls.Add("Forms.Label.1", "lblStatus", $true)
    $lblStatus.Left = 10; $lblStatus.Top = 54; $lblStatus.Width = 310; $lblStatus.Height = 18
    $lblStatus.Caption = "Klar"

    $cmdInfo = $designer.Controls.Add("Forms.CommandButton.1", "cmdInfo", $true)
    $cmdInfo.Left = 90; $cmdInfo.Top = 76; $cmdInfo.Width = 150; $cmdInfo.Height = 26
    $cmdInfo.Caption = "Installert info"

    $frmComp.CodeModule.AddFromString($formCode)

    # ---- frmMakroInfo (nytt info-vindu, se "Installert info..."-knappen over) ----
    $infoComp = $buildProj.VBComponents.Add(3)
    $infoComp.Name = "frmMakroInfo"
    $infoComp.Properties("Width").Value = 410
    $infoComp.Properties("Height").Value = 420
    $infoComp.Properties("Caption").Value = "Installert info"
    $infoDesigner = $infoComp.Designer

    $lblKompHeader = $infoDesigner.Controls.Add("Forms.Label.1", "lblKomponenterHeader", $true)
    $lblKompHeader.Left = 10; $lblKompHeader.Top = 10; $lblKompHeader.Width = 370; $lblKompHeader.Height = 16
    $lblKompHeader.Caption = "VBA-komponenter"

    $lstKomponenter = $infoDesigner.Controls.Add("Forms.ListBox.1", "lstKomponenter", $true)
    $lstKomponenter.Left = 10; $lstKomponenter.Top = 28; $lstKomponenter.Width = 380; $lstKomponenter.Height = 110
    $lstKomponenter.ColumnCount = 2
    $lstKomponenter.ColumnWidths = "60pt;300pt"

    $lblArkHeader = $infoDesigner.Controls.Add("Forms.Label.1", "lblArkHeader", $true)
    $lblArkHeader.Left = 10; $lblArkHeader.Top = 146; $lblArkHeader.Width = 370; $lblArkHeader.Height = 16
    $lblArkHeader.Caption = "Skjulte innstillingsark (very hidden)"

    $frameArk = $infoDesigner.Controls.Add("Forms.Frame.1", "frameArk", $true)
    $frameArk.Left = 10; $frameArk.Top = 164; $frameArk.Width = 385; $frameArk.Height = 40
    $frameArk.Caption = ""

    $cmdLukk = $infoDesigner.Controls.Add("Forms.CommandButton.1", "cmdLukk", $true)
    $cmdLukk.Left = 10; $cmdLukk.Top = 214; $cmdLukk.Width = 100; $cmdLukk.Height = 26
    $cmdLukk.Caption = "Lukk"

    $infoComp.CodeModule.AddFromString($infoFormCode)

    $arkToggleCls = $buildProj.VBComponents.Add(2)
    $arkToggleCls.Name = "clsArkToggle"
    $arkToggleCls.CodeModule.AddFromString($arkToggleCodeRaw)

    $mod.Export((Join-Path $tempDir "modMakroMeny.bas"))
    $cls.Export((Join-Path $tempDir "clsMakroTile.cls"))
    $frmComp.Export((Join-Path $tempDir "frmMakroMeny.frm"))
    $infoComp.Export((Join-Path $tempDir "frmMakroInfo.frm"))
    $arkToggleCls.Export((Join-Path $tempDir "clsArkToggle.cls"))

    $buildWb.Close($false)

    # ---- 5. Bytt ut / legg til komponentene i maal-arbeidsboken ----

    $componentsToReplace = @(
        @{ Name = "modMakroMeny"; File = "modMakroMeny.bas" },
        @{ Name = "clsMakroTile"; File = "clsMakroTile.cls" },
        @{ Name = "frmMakroMeny"; File = "frmMakroMeny.frm" },
        @{ Name = "frmMakroInfo"; File = "frmMakroInfo.frm" },
        @{ Name = "clsArkToggle"; File = "clsArkToggle.cls" }
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

    # ---- 5b. Opprydding av "vis som fane"/baandknapp-funksjonene, som er
    # ---- FJERNET HELT (Haakons eksplisitte onske 2026-09-23 - provde
    # ---- begge, likte ikke resultatet). Kjores ved HVER install/oppdatering
    # ---- (ikke bare -Uninstall), slik at en fil som allerede har rester fra
    # ---- en tidligere versjon (2.4.x) blir ryddet for dem automatisk uten
    # ---- at brukeren maa avinstallere/reinstallere manuelt.
    try {
        $thisWbComp = $vbproj.VBComponents.Item("ThisWorkbook")
        Remove-ArkKnappFraThisWorkbook $thisWbComp
    } catch {}
    try {
        if ($wb.Worksheets | Where-Object { $_.Name -eq "Makromeny" }) {
            $wb.Worksheets.Item("Makromeny").Delete()
            Write-Output "  Fjernet gammelt 'Makromeny'-fanearket (funksjonen er fjernet)"
        }
    } catch {}
    try {
        $cb = $excel.CommandBars | Where-Object { $_.Name -eq "Makromeny" }
        if ($cb -and $cb.Controls.Count -gt 0 -and $cb.Controls.Item(1).OnAction -like "*$($wb.Name)*") {
            $cb.Delete()
            Write-Output "  Fjernet gammel baandknapp (funksjonen er fjernet)"
        }
    } catch {}

    Write-Output ""
    Write-Output "Ferdig. Husk aa lagre filen selv (Ctrl+S) naar du er klar."
} finally {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
