# ============================================================================
# Excel Macro Installer — sentralt sted for å installere/oppdatere alle de
# ferdige Excel-makroene i denne mappa (Kolonnevelger, Mail-utsender, osv.),
# og for å se hvilke Excel-filer som allerede har fått hvilken versjon.
#
# To faner ("Installerte filer" er standard/venstre fane):
#   "Installerte filer" - ett kort per fil (+ Legg til fil for å legge inn
#                         en fil før du velger makro), med hver installerte
#                         makro som egen linje (+ Legg til makro, Oppdater,
#                         Fjern - Fjern kaller makroens egen -Uninstall og
#                         fjerner FAKTISK VBA-koden/knappen fra fila).
#   "Makroer"          - ett kort per makro (ikon, nyeste versjon, sist
#                         endret, kort beskrivelse), med en "Installer/
#                         Oppdater"-knapp som selv spør om målfil.
#
# Mappestruktur (portabel - ingen hardkodede stier):
#   Excel Macro Installer\
#   |- Excel Macro Installer.exe   <- kompilert (ps2exe) - dette er filen som
#   |                                  faktisk dobbeltklikkes/festes til oppgavelinjen
#   |- Macros\                     <- de ferdige makro-installerne som listes/startes
#   |                                  (må støtte -Path og -Uninstall for full funksjonalitet)
#   |- Resources\                  <- icon.ico + ett ikon per makro
#   |- Data\                       <- macro-metadata.json (beskrivelser),
#   |                                  installed-files.json (register over faktiske
#   |                                  installasjoner) og known-files.json (filer lagt
#   |                                  til via "+ Legg til fil") - launcheren skriver alt dette selv
#   \- Installer\
#      \- Excel Macro Installer.ps1  <- kildekoden (denne fila) - rediger her,
#                                        kompiler sa pa nytt til .exe-en over
# ============================================================================

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Uten dette regner Windows programmet som "DPI-uvitende" og skalerer HELE
# vinduet som et strukket bitmap for a matche skjermens skalering (f.eks.
# 125%) - det er derfor gamle/uoppdaterte programmer ser uskarpe/piksellerte
# ut pa hoy-DPI skjermer. Ma kalles FOR noe som helst vindu/kontroll
# opprettes, derfor helt her oppe. -4 = DPI_AWARENESS_CONTEXT_PER_MONITOR_
# AWARE_V2 (mest robust variant, handterer ogsa flere skjermer med ulik
# skalering).
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DpiFix {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
}
"@
    [void][DpiFix]::SetProcessDpiAwarenessContext([IntPtr](-4))
} catch { }

# Bump denne OG "-version"-verdien i Invoke-ps2exe-kommandoen samtidig hver
# gang programmet rekompileres - selvoppdaterings-sjekken lenger nede
# sammenligner disse to (bakt-inn-versjon vs. fersk lest fra .exe-fila pa
# disk) for a varsle om en nyere versjon ligger klar.
$LauncherVersion = "2.9.0.0"

# Kjort som kompilert .exe (ps2exe, GUI-undersystem - ingen konsoll/
# terminal-vindu i det hele tatt, uansett Windows sine Terminal-
# innstillinger) for et helt rent dobbeltklikk-oppstart. En ufanget feil
# ville ellers forsvunnet helt usett uten noen konsoll a vise den i - hele
# programmet kjores derfor i en try/catch som viser en vanlig MessageBox
# hvis noe uventet skjer, sa en feil alltid blir synlig for Håkon.
try {

# $PSScriptRoot er tom i en ps2exe-kompilert .exe (ingen ekte skriptfil pa
# disk lenger) - bruk da exe-filens egen plassering i stedet. Og siden
# .exe-en ligger direkte i "Excel Macro Installer\" mens kildekoden ligger
# ett niva dypere i "...\Installer\", finner vi riktig basismappe ved a
# sjekke om Macros/Resources faktisk finnes rett under - fungerer likt om
# man kjorer den ferdige .exe-en ELLER denne .ps1-kilden direkte for testing.
$erKompilert = [string]::IsNullOrEmpty($PSScriptRoot) -or -not (Test-Path $PSScriptRoot)
if ($erKompilert) {
    $eigenMappe = Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Parent
} else {
    $eigenMappe = $PSScriptRoot
}

if ((Test-Path (Join-Path $eigenMappe "Macros")) -and (Test-Path (Join-Path $eigenMappe "Resources"))) {
    $basisMappe = $eigenMappe
} else {
    $basisMappe = Split-Path $eigenMappe -Parent
}
$makroMappe = Join-Path $basisMappe "Macros"
$ressursMappe = Join-Path $basisMappe "Resources"
$dataMappe = Join-Path $basisMappe "Data"
if (-not (Test-Path $dataMappe)) { New-Item -ItemType Directory -Path $dataMappe | Out-Null }

# ---- Oppdatering fra GitHub (v2.9.0.0) ----
# Pakken hentes fra det offentlige prosjektrepoet (zip av main). KUN
# programfilene (exe, Macros\*.ps1, Resources\, Data\macro-metadata.json,
# versjon.json) fra mappa "Excel Macro Installer\" i repoet kopieres inn -
# aldri noe annet fra repoet. Ny versjon publiseres ved a oke versjon.json
# (Prosjektfiler\_publisering\Publiser-EMI.ps1) og pushe. versjon.json i basismappa er
# den installerte PAKKE-versjonen (uavhengig av $LauncherVersion, siden en ny
# makroversjon ogsa er en ny pakke uten at exe-en endres). Mangler fila
# (installasjon fra for v2.9.0.0) regnes den som 0.0.0.0, sa forste
# publiserte pakke alltid tilbys.
$script:OppdateringRepo = "timmytimss/excelVBAkoding"
$script:OppdateringVersjonUrl = "https://raw.githubusercontent.com/$($script:OppdateringRepo)/main/Excel%20Macro%20Installer/versjon.json"
# Hviteliste over filer (relativt til "Excel Macro Installer\") som en
# oppdatering har lov a skrive. Alt annet i repo-zipen ignoreres.
$script:OppdateringHviteliste = '^(Excel Macro Installer\.exe|versjon\.json|Data\\macro-metadata\.json|Macros\\[^\\]+\.ps1|Resources\\.+)$'
$script:OppdateringZipUrl = "https://codeload.github.com/$($script:OppdateringRepo)/zip/refs/heads/main"
$script:lokalPakkeVersjon = [version]"0.0.0.0"
try {
    $lokalVersjonFil = Join-Path $basisMappe "versjon.json"
    if (Test-Path $lokalVersjonFil) {
        $script:lokalPakkeVersjon = [version]((Get-Content -Raw -Encoding UTF8 $lokalVersjonFil | ConvertFrom-Json).versjon)
    }
} catch { }
# En utviklingskopi skal ALDRI overskrive seg selv med den publiserte
# pakken - der er det kildekoden som er fasit (ellers kunne «Oppdater» i
# OneDrive-prosjektet rulle tilbake upubliserte makroendringer). Kjennetegn:
# kildekoden Installer\Excel Macro Installer.ps1 ligger ved siden av exe-en
# (OneDrive-prosjektet og git-klonene) - den er aldri med i en ekte
# installasjon, siden oppdateringens hviteliste ikke omfatter Installer\.
$script:oppdateringMulig = $erKompilert -and
    -not (Test-Path (Join-Path $basisMappe "Installer\Excel Macro Installer.ps1")) -and
    -not (Test-Path (Join-Path (Split-Path $basisMappe -Parent) ".git"))
# Rester fra forrige oppdatering: filer som var i bruk (den kjorende exe-en,
# lastede DLL-er) ble omdopt til *.gammel i stedet for overskrevet, og kan
# forst slettes nar den gamle prosessen er borte - altsa na.
try {
    Get-ChildItem -Path $basisMappe -Recurse -Filter "*.gammel" -File -ErrorAction SilentlyContinue |
        ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch { } }
} catch { }

$ikonSti = Join-Path $ressursMappe "icon.ico"

# ---- WebView2 (kunnskapsbase-fanen) ----
# Egne DLL-er i Resources\WebView2\ (IKKE en NuGet-pakke installert pa
# maskinen - EMI er ikke et vanlig .NET-prosjekt med pakkehandtering, sa
# DLL-ene er hentet ut fra NuGet-pakken Microsoft.Web.WebView2 manuelt og
# ligger her permanent, samme prinsipp som ikonene i Resources for ovrig).
# WebView2Loader.dll (native) MA vaere i prosessens DLL-sokesti for a bli
# funnet - AddDllDirectory gjor dette uten a matte kopiere den til siden av
# selve .exe-en. Feiler HELE denne blokken (DLL-ene mangler, feil
# arkitektur, WebView2 Runtime ikke installert): $script:webView2Tilgjengelig
# forblir $false og kunnskapsbase-fanen viser en enkel "apne i nettleser"-
# knapp i stedet - appen skal aldri krasje pa grunn av dette.
$script:webView2Tilgjengelig = $false
$webView2Mappe = Join-Path $ressursMappe "WebView2"
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DllSok {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool SetDllDirectory(string lpPathName);
}
"@
    [void][DllSok]::SetDllDirectory($webView2Mappe)
    Add-Type -Path (Join-Path $webView2Mappe "Microsoft.Web.WebView2.Core.dll")
    Add-Type -Path (Join-Path $webView2Mappe "Microsoft.Web.WebView2.WinForms.dll")
    $script:webView2Tilgjengelig = $true
} catch {
    $script:webView2Tilgjengelig = $false
}

# Selve kunnskapsbase-HTML-fila ligger i prosjektroten, ETT nivaa over
# "Excel Macro Installer\" (se mappestruktur-kommentaren helt ovst i fila).
$kunnskapsbaseSti = Join-Path (Split-Path $basisMappe -Parent) "Excel VBA Koding.html"

# ---- Datatilgang ----

function Finn-Versjon {
    param([string]$innhold)

    # Prioritet 1: en toppnivå PowerShell-variabel som "$GeneriskVersion = "0.13.2""
    # — den mest pålitelige kilden, siden noen installere (som Mail-utsender)
    # bruker en {{VERSION}}-mal inni selve VBA-kildekoden som IKKE er et ekte
    # tall før installasjon (den byttes ut da) og derfor ikke kan brukes her.
    $m = [regex]::Match($innhold, '\$\w*[Vv]ersion\s*=\s*"([^"]+)"')
    if ($m.Success) { return $m.Groups[1].Value }

    # Prioritet 2: en VBA Const med et ekte tall direkte i kildekoden
    # (f.eks. "COLUMN_PICKER_VERSION As String = "1.0.1"").
    $m = [regex]::Match($innhold, '\w*VERSION\w*\s+As\s+String\s*=\s*"([^"{]+)"')
    if ($m.Success) { return $m.Groups[1].Value }

    return $null
}

function Hent-Makroliste {
    $resultat = New-Object System.Collections.Generic.List[object]
    Get-ChildItem -Path $makroMappe -Filter '*.ps1' -File | Sort-Object Name | ForEach-Object {
        $innhold = Get-Content -Raw -Encoding UTF8 $_.FullName
        $versjon = Finn-Versjon $innhold
        $navn = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        $resultat.Add([PSCustomObject]@{
            Navn       = $navn
            Versjon    = $versjon
            Sti        = $_.FullName
            SistEndret = $_.LastWriteTime
        })
    }
    # Ledende komma - IKKE bare stil: uten den "ruller PowerShell ut" en
    # List[object] nar den sendes til pipelinen via return, og en TOM eller
    # ETT-ELEMENTS liste ville da blitt hhv. $null eller det bare elementet
    # i stedet for selve List-objektet hos mottakeren (klassisk PowerShell-
    # fallgruve). Se excel-macro-installer-launcher i minnet.
    return ,$resultat
}

function Hent-Metadata {
    $metaSti = Join-Path $dataMappe "macro-metadata.json"
    if (Test-Path $metaSti) {
        try {
            $raw = Get-Content -Raw -Encoding UTF8 $metaSti
            $obj = ConvertFrom-Json $raw
            if ($obj) { return $obj }
        } catch { }
    }
    return [PSCustomObject]@{}
}

function Les-InstallertRegister {
    param([switch]$MeldFraOmFeil)

    $regSti = Join-Path $dataMappe "installed-files.json"
    $liste = New-Object System.Collections.Generic.List[object]
    $script:sisteLesingOk = $true

    if (Test-Path $regSti) {
        # Prov flere ganger med en liten pause mellom - fila ligger i en
        # OneDrive-synket mappe og kan vaere kortvarig last rett etter at
        # DENNE launcheren selv nettopp skrev til den (les-etter-skriv rett
        # etter installasjon). Uten dette kunne en forbigaende lesefeil bli
        # tolket som "registeret er tomt" og deretter SKREVET tilbake tomt
        # av Oppdater-InstallertRegister - dvs. faktisk data-tap. Se
        # excel-macro-installer-launcher i minnet.
        $forsok = 0
        $lest = $false
        $sisteFeil = $null
        while (-not $lest -and $forsok -lt 5) {
            try {
                $raw = Get-Content -Raw -Encoding UTF8 $regSti
                $data = ConvertFrom-Json $raw
                foreach ($rad in @($data)) { if ($rad) { $liste.Add($rad) } }
                $lest = $true
            } catch {
                $sisteFeil = $_
                $forsok++
                Start-Sleep -Milliseconds 150
            }
        }
        if (-not $lest) {
            $script:sisteLesingOk = $false
            if ($MeldFraOmFeil) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Klarte ikke å lese inn registeret over installerte filer etter flere forsøk (kan skyldes at OneDrive synkroniserer fila akkurat nå). Prøv igjen om et par sekunder." + "`n`n" + $sisteFeil.Exception.Message,
                    "Excel Macro Installer", "OK", "Warning") | Out-Null
            }
        }
    }

    # Ledende komma - se forklaring i Hent-Makroliste over. Uten den ville
    # et TOMT register (helt normalt forste gang programmet kjores) blitt
    # $null hos mottakeren i stedet for en tom liste.
    return ,$liste
}

function Hent-KjenteFiler {
    # Filer Håkon eksplisitt har lagt til via "+ Legg til fil" i Installerte
    # filer-fanen, UAVHENGIG av om noen makro er installert på dem ennå -
    # gjør at en fil kan dukke opp i oversikten før man har valgt en makro
    # å installere på den. Slås sammen med de faktiske installasjonene
    # (installed-files.json) i Last-FilerFane.
    $stiFil = Join-Path $dataMappe "known-files.json"
    $liste = New-Object System.Collections.Generic.List[string]
    if (Test-Path $stiFil) {
        try {
            $raw = Get-Content -Raw -Encoding UTF8 $stiFil
            $data = ConvertFrom-Json $raw
            foreach ($f in @($data)) { if ($f) { $liste.Add([string]$f) } }
        } catch { }
    }
    return ,$liste
}

function Lagre-KjentFil {
    param([string]$Fil)
    $stiFil = Join-Path $dataMappe "known-files.json"
    $liste = Hent-KjenteFiler
    if ($liste -notcontains $Fil) {
        $liste.Add($Fil)
        $json = ConvertTo-Json -InputObject $liste.ToArray() -Depth 3
        [System.IO.File]::WriteAllText($stiFil, $json, (New-Object System.Text.UTF8Encoding($false)))
    }
}

function Skriv-InstallertRegister {
    param([System.Collections.Generic.List[object]]$Liste)
    $regSti = Join-Path $dataMappe "installed-files.json"
    # .ToArray() - IKKE @($Liste). @() pa en ekte List[object] treffer en
    # obskur PowerShell-binder-feil ("Argumenttypene samsvarer ikke", inni
    # PSToObjectArrayBinder) i denne PowerShell-versjonen. .ToArray() gir et
    # ekte System.Object[] direkte fra .NET og unngar hele problemet - og
    # ConvertTo-Json serialiserer det riktig som JSON-array bade nar det er
    # tomt og nar det har innhold.
    $json = ConvertTo-Json -InputObject $Liste.ToArray() -Depth 5
    [System.IO.File]::WriteAllText($regSti, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Oppdater-InstallertRegister {
    param([string]$Fil, [string]$Makro, [string]$Versjon)
    $liste = Les-InstallertRegister -MeldFraOmFeil
    if (-not $script:sisteLesingOk) {
        # IKKE skriv tilbake her - en mislykket lesing ville gitt en (nesten)
        # tom liste, og a skrive DEN tilbake ville slettet alle radene som
        # faktisk star i fila. Bedre a bare hoppe over denne logg-
        # oppdateringen (installasjonen skjedde uansett fint) enn a risikere
        # data-tap i registeret.
        return
    }
    $eksisterende = $liste | Where-Object { $_.Fil -eq $Fil -and $_.Makro -eq $Makro } | Select-Object -First 1
    if ($eksisterende) {
        $eksisterende.Versjon = $Versjon
        $eksisterende.Tidspunkt = (Get-Date).ToString("s")
    } else {
        $liste.Add([PSCustomObject]@{ Fil = $Fil; Makro = $Makro; Versjon = $Versjon; Tidspunkt = (Get-Date).ToString("s") })
    }
    Skriv-InstallertRegister $liste
}

function Fjern-FraInstallertRegister {
    # Fjerner KUN sporingsraden (registeret), rører ikke selve Excel-fila -
    # brukes etter en bekreftet vellykket -Uninstall (se Fjern-MakroFraFil),
    # og indirekte via Fjern-FilHelt.
    param([string]$Fil, [string]$Makro)
    $liste = Les-InstallertRegister -MeldFraOmFeil
    if (-not $script:sisteLesingOk) { return }
    $ny = New-Object System.Collections.Generic.List[object]
    foreach ($rad in $liste) {
        if (-not ($rad.Fil -eq $Fil -and $rad.Makro -eq $Makro)) { $ny.Add($rad) }
    }
    Skriv-InstallertRegister $ny
}

function Fjern-KjentFil {
    param([string]$Fil)
    $liste = Hent-KjenteFiler
    $ny = New-Object System.Collections.Generic.List[string]
    foreach ($f in $liste) { if ($f -ne $Fil) { $ny.Add($f) } }
    $stiFil = Join-Path $dataMappe "known-files.json"
    $json = ConvertTo-Json -InputObject $ny.ToArray() -Depth 3
    [System.IO.File]::WriteAllText($stiFil, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Fjern-FilHelt {
    # "Fjern fil" - ren sporing, rører ALDRI selve Excel-fila. Fjerner fra
    # både known-files.json og alle registerrader for den fila.
    param([string]$Fil)
    Fjern-KjentFil -Fil $Fil
    $liste = Les-InstallertRegister -MeldFraOmFeil
    if (-not $script:sisteLesingOk) { return }
    $ny = New-Object System.Collections.Generic.List[object]
    foreach ($rad in $liste) { if ($rad.Fil -ne $Fil) { $ny.Add($rad) } }
    Skriv-InstallertRegister $ny
}

function Fjern-MakroFraFil {
    # "Fjern makro" - EKTE avinstallering: kaller makroens egen .ps1 med
    # -Uninstall, som fjerner VBA-komponentene/knappen FRA SELVE fila (se
    # Del B i denne makro-familiens installere). Kjøres SYNKRONT (-Wait),
    # i motsetning til vanlig installasjon som er asynkron - avinstallering
    # er rask og trenger ingen dialogvinduer, og vi vil vite at det faktisk
    # lyktes (riktig avslutningskode) FØR raden fjernes fra registeret.
    param($Makro, [string]$FilSti, [System.Windows.Forms.Label]$StatusLabel)

    $StatusLabel.Text = "Fjerner " + $Makro.Navn + " fra: " + $FilSti + " ..."
    [System.Windows.Forms.Application]::DoEvents()

    $proc = Start-Process powershell.exe -ArgumentList @(
        "-ExecutionPolicy", "Bypass",
        "-File", ('"' + $Makro.Sti + '"'),
        "-Path", ('"' + $FilSti + '"'),
        "-Uninstall"
    ) -Wait -PassThru

    if ($proc.ExitCode -eq 0) {
        Fjern-FraInstallertRegister -Fil $FilSti -Makro $Makro.Navn
        $StatusLabel.Text = "Fjernet " + $Makro.Navn + " fra: " + $FilSti
        return $true
    } else {
        $StatusLabel.Text = ""
        [System.Windows.Forms.MessageBox]::Show(
            "Klarte ikke å fjerne " + $Makro.Navn + " fra fila (avsluttet med feilkode $($proc.ExitCode)). Sjekk konsollvinduet som åpnet seg for detaljer.",
            "Excel Macro Installer", "OK", "Warning") | Out-Null
        return $false
    }
}

function Ny-BryttLabel {
    # Bygger en Label som VOKSER I HOYDEN for a passe teksten ved en fast
    # bredde, i stedet for a klippe den - erstatter tidligere faste/gjettede
    # Height-verdier som ikke strakk til for lengre beskrivelser/stier.
    # AutoSize + MaximumSize(bredde, 0) er den vanlige WinForms-knepet for
    # "ordbryt ved denne bredden, la hoyden folge automatisk".
    param([string]$Tekst, [int]$Bredde, [System.Drawing.Font]$Font = $null)
    $lbl = New-Object System.Windows.Forms.Label
    if ($Font) { $lbl.Font = $Font }
    $lbl.AutoSize = $true
    $lbl.MaximumSize = New-Object System.Drawing.Size($Bredde, 0)
    $lbl.Text = $Tekst
    return $lbl
}

function Velg-ExcelFil {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Velg Excel-fila makroen skal installeres/oppdateres i"
    $dlg.Filter = "Excel-filer (*.xlsm;*.xlsx;*.xlsb)|*.xlsm;*.xlsx;*.xlsb|Alle filer (*.*)|*.*"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.FileName
    }
    return $null
}

function Start-MakroMedFil {
    param($Makro, [string]$FilSti, [switch]$Tving)
    # Egen, uavhengig prosess - launcheren fortsetter å virke mens den
    # valgte installeren kjører (egne dialogvinduer for eventuelle andre
    # valg, egen levetid). -Path sendes eksplisitt slik at IKKE installeren
    # sin egen fildialog vises - launcheren har allerede spurt.
    # -Tving sender -Force (alle fire installere stotter det siden 2.7.0.0) sa
    # installasjonen kjores pa nytt selv om versjonen er lik ("Reinstaller
    # alle"). Returnerer prosessen sa en ko kan vente pa at den er ferdig.
    $args2 = @(
        "-ExecutionPolicy", "Bypass",
        "-File", ('"' + $Makro.Sti + '"'),
        "-Path", ('"' + $FilSti + '"')
    )
    if ($Tving) { $args2 += "-Force" }
    $prosess = Start-Process powershell.exe -ArgumentList $args2 -PassThru
    Oppdater-InstallertRegister -Fil $FilSti -Makro $Makro.Navn -Versjon $Makro.Versjon
    return $prosess
}

# Installasjonskø: "Oppdater alle"/"Reinstaller alle" kjorer makroene EN og EN
# (aldri samtidig - to installere som redigerer samme arbeidsbok/VBA-prosjekt
# parallelt kan kollidere). Et Timer-tikk sjekker om forrige prosess er ferdig
# og starter neste, sa UI-et aldri fryser mens man venter.
$script:installKo = New-Object System.Collections.Queue
$script:installProsess = $null
$script:installTimer = New-Object System.Windows.Forms.Timer
$script:installTimer.Interval = 700
$script:installTimer.Add_Tick({
    if ($script:installProsess -and -not $script:installProsess.HasExited) { return }
    if ($script:installKo.Count -eq 0) {
        $script:installTimer.Stop()
        $script:installProsess = $null
        $lblStatusBunn.Text = "Ferdig - alle valgte makroer er kjørt."
        Last-FilerFane
        return
    }
    $jobb = $script:installKo.Dequeue()
    $lblStatusBunn.Text = "Installerer " + $jobb.Makro.Navn + " (" + $script:installKo.Count + " igjen) ..."
    $script:installProsess = Start-MakroMedFil -Makro $jobb.Makro -FilSti $jobb.Fil -Tving:([bool]$jobb.Tving)
})

function Test-InstallKoAktiv {
    return ($script:installKo.Count -gt 0 -or ($script:installProsess -and -not $script:installProsess.HasExited))
}

function Start-InstallKo {
    param([string]$FilSti, $Rader, [bool]$Tving)
    if (Test-InstallKoAktiv) {
        [System.Windows.Forms.MessageBox]::Show("En installasjon pågår allerede - vent til den er ferdig.", "Excel Macro Installer", "OK", "Information") | Out-Null
        return
    }
    foreach ($rad in @($Rader)) {
        $makro = $makroer | Where-Object { $_.Navn -eq $rad.Makro } | Select-Object -First 1
        if ($makro) { $script:installKo.Enqueue([PSCustomObject]@{ Makro = $makro; Fil = $FilSti; Tving = $Tving }) }
    }
    if ($script:installKo.Count -gt 0) { $script:installTimer.Start() }
}

# ---- Vindu ----
# Fargepalett og smaa hjelpefunksjoner for et flatere, mer moderne utseende
# (avrundede hjorner/piller via Region-klipping, farge-badges, statusprikker)
# - inspirert av en HTML-mockup Håkon fikk laget, gjenskapt sa naert som
# WinForms tillater uten a ga til fullt egendefinert tegning av hver kontroll.

$script:FargePrimaer      = [System.Drawing.Color]::FromArgb(79, 99, 210)
$script:FargeSideBg       = [System.Drawing.Color]::FromArgb(246, 247, 249)
$script:FargeTittel       = [System.Drawing.Color]::FromArgb(25, 28, 36)
$script:FargeUndertittel  = [System.Drawing.Color]::FromArgb(110, 116, 128)
$script:FargeMuted        = [System.Drawing.Color]::FromArgb(150, 155, 165)
$script:FargeKortKant     = [System.Drawing.Color]::FromArgb(228, 230, 235)
$script:FargeBadgeBg      = [System.Drawing.Color]::FromArgb(253, 236, 207)
$script:FargeBadgeTekst   = [System.Drawing.Color]::FromArgb(161, 98, 7)
$script:FargeGronnPrikk   = [System.Drawing.Color]::FromArgb(34, 197, 94)
$script:FargeOransjePrikk = [System.Drawing.Color]::FromArgb(245, 158, 11)
$script:FargePillBg       = [System.Drawing.Color]::FromArgb(240, 241, 245)
$script:FargePillTekst    = [System.Drawing.Color]::FromArgb(90, 95, 105)
$script:FargeLenkeRod     = [System.Drawing.Color]::FromArgb(160, 40, 40)

function Set-RundedeHjorner {
    # Klipper en kontroll til en avrundet rektangel-form via en GraphicsPath-
    # Region. MA kalles ETTER at kontrollens endelige Width/Height er satt.
    param([System.Windows.Forms.Control]$Ctrl, [int]$Radius)
    $r = [Math]::Min($Radius, [Math]::Min($Ctrl.Width, $Ctrl.Height) / 2)
    if ($r -le 0) { return }
    $d = $r * 2
    $rect = New-Object System.Drawing.Rectangle(0, 0, $Ctrl.Width, $Ctrl.Height)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($rect.X, $rect.Y, $d, $d, 180, 90)
    $path.AddArc($rect.Right - $d, $rect.Y, $d, $d, 270, 90)
    $path.AddArc($rect.Right - $d, $rect.Bottom - $d, $d, $d, 0, 90)
    $path.AddArc($rect.X, $rect.Bottom - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    $Ctrl.Region = New-Object System.Drawing.Region($path)
}

function Ny-PillLabel {
    # Liten "badge"/versjonstagg med farget, helt avrundet (piller-formet)
    # bakgrunn - storrelsen males direkte fra teksten (TextRenderer), sa den
    # ikke er avhengig av AutoSize-timing (som krever at kontrollen alt er
    # lagt til et Controls-sett for a male riktig - se Ny-BryttLabel).
    param([string]$Tekst, [System.Drawing.Color]$Bakgrunn, [System.Drawing.Color]$Tekstfarge, [bool]$Fet = $true, [single]$FontStorrelse = 8)
    $stil = if ($Fet) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $font = New-Object System.Drawing.Font("Segoe UI", $FontStorrelse, $stil)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Tekst
    $lbl.Font = $font
    $lbl.ForeColor = $Tekstfarge
    $lbl.BackColor = $Bakgrunn
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    # Bredde/hoyde-margin og avrundingsradius ma stemme overens - full
    # piller-avrunding (radius = halve hoyden, som forr) spiser inn i selve
    # hjornene av teksten sin bounding box nar innvendig margin er mindre
    # enn radiusen, og klipper hjornene av tegn som star helt inntil kanten
    # (typisk tall - rapportert av Håkon 2026-09-10, "se tallene"). Fast,
    # beskjeden radius (ikke Height/2) + storre margin unngar dette, og ser
    # fortsatt tydelig avrundet ut.
    # 2026-09-24 (Haakon: "pillen er litt for stor for selve teksten"):
    # strammere innvendig margin - TextRenderer.MeasureText legger allerede
    # til ca. 6 px horisontal padding selv, sa +8/+4 gir en pille som sitter
    # tett rundt teksten uten a klippe hjornene (radius 6).
    $maalt = [System.Windows.Forms.TextRenderer]::MeasureText($Tekst, $font)
    $lbl.Width = $maalt.Width + 8
    $lbl.Height = $maalt.Height + 4
    Set-RundedeHjorner -Ctrl $lbl -Radius 6
    return $lbl
}

function Ny-StatusPrikk {
    param([System.Drawing.Color]$Farge, [int]$Storrelse = 8)
    $p = New-Object System.Windows.Forms.Panel
    $p.Width = $Storrelse; $p.Height = $Storrelse
    $p.BackColor = $Farge
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddEllipse(0, 0, $Storrelse, $Storrelse)
    $p.Region = New-Object System.Drawing.Region($path)
    return $p
}

function Ny-IkonKvadrat {
    # Avrundet, fargelagt kvadrat med makroens eget ikon inni - erstatter den
    # rene PictureBox-en fra for for a matche mockupens "ikon-plate"-stil.
    param([string]$IkonSti, [System.Drawing.Color]$Bakgrunn, [int]$Storrelse = 40)
    $p = New-Object System.Windows.Forms.Panel
    $p.Width = $Storrelse; $p.Height = $Storrelse
    $p.BackColor = $Bakgrunn
    Set-RundedeHjorner -Ctrl $p -Radius 9
    if ($IkonSti -and (Test-Path $IkonSti)) {
        $innStorrelse = [int]($Storrelse * 0.55)
        $pic = New-Object System.Windows.Forms.PictureBox
        $pic.Width = $innStorrelse; $pic.Height = $innStorrelse
        $pic.Left = [int](($Storrelse - $innStorrelse) / 2); $pic.Top = $pic.Left
        $pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $pic.BackColor = [System.Drawing.Color]::Transparent
        $tmpIcon = New-Object System.Drawing.Icon($IkonSti)
        $pic.Image = $tmpIcon.ToBitmap()
        $p.Controls.Add($pic)
    }
    return $p
}

function Ny-Knapp {
    # Stil: "primaer" (fylt bla, avrundet), "outline" (hvit m/ grontt/graa
    # kant), "lenke" (bare farget tekst, ingen kant/bakgrunn - for "Fjern").
    # $Bredde = 0 (eller utelatt) betyr AUTO-bredde - malt fra selve teksten
    # (TextRenderer.MeasureText) + fast innvendig margin, i stedet for en
    # gjettet konstant som kan bli for smal og klippe teksten pa en annen
    # DPI/skjerm/font-rendering enn den ble testet pa (se tilsvarende
    # lardom for Ny-PillLabel og "+ Legg til makro"-knappen tidligere).
    param([string]$Tekst, [string]$Stil = "outline", [int]$Bredde = 0, [int]$Hoyde = 30)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Tekst
    $b.Height = $Hoyde
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $b.UseVisualStyleBackColor = $false
    switch ($Stil) {
        "primaer" {
            $b.BackColor = $script:FargePrimaer
            $b.ForeColor = [System.Drawing.Color]::White
            $b.FlatAppearance.BorderSize = 0
            $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(66, 84, 190)
            $b.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        }
        "outline" {
            $b.BackColor = [System.Drawing.Color]::White
            $b.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 70)
            $b.FlatAppearance.BorderSize = 1
            $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(210, 212, 218)
            $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(245, 246, 248)
            $b.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        }
        "gronn" {
            $b.BackColor = [System.Drawing.Color]::FromArgb(30, 150, 90)
            $b.ForeColor = [System.Drawing.Color]::White
            $b.FlatAppearance.BorderSize = 0
            $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(24, 128, 76)
            $b.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        }
        "aksent" {
            $b.BackColor = [System.Drawing.Color]::FromArgb(232, 236, 255)
            $b.ForeColor = $script:FargePrimaer
            $b.FlatAppearance.BorderSize = 1
            $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(170, 182, 240)
            $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(218, 224, 252)
            $b.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        }
        "lenke" {
            $b.BackColor = [System.Drawing.Color]::White
            $b.ForeColor = $script:FargeLenkeRod
            $b.FlatAppearance.BorderSize = 0
            $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(250, 240, 240)
            $b.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        }
    }
    if ($Bredde -gt 0) {
        $b.Width = $Bredde
    } else {
        $maalt = [System.Windows.Forms.TextRenderer]::MeasureText($Tekst, $b.Font)
        $b.Width = $maalt.Width + $(if ($Stil -eq "lenke") { 16 } else { 32 })
    }
    if ($Stil -eq "primaer" -or $Stil -eq "gronn") { Set-RundedeHjorner -Ctrl $b -Radius 6 }
    return $b
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel Macro Installer"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
$form.ClientSize = New-Object System.Drawing.Size(900, 760)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.MinimumSize = New-Object System.Drawing.Size(820, 550)
$form.BackColor = [System.Drawing.Color]::White

if (Test-Path $ikonSti) {
    $form.Icon = New-Object System.Drawing.Icon($ikonSti)
}

$sideMargin = 28
$innholdBredde = $form.ClientSize.Width - ($sideMargin * 2)
# Trekker fra bredden pa et vertikalt rullefelt i tillegg til den vanlige
# 16 px marginen - uten dette blir kortene EN anelse bredere enn det som
# faktisk er plass til sa snart flow-panelet trenger a vise et vertikalt
# rullefelt (fler filer/makroer enn det som er synlig samtidig), som i sin
# tur uonsket trigger et VANNRETT rullefelt ogsa. Reserverer plassen
# proaktivt i stedet for a oppdage det for sent (funnet ved resize-testing
# 2026-09-22, Nils).
$kortBredde = $innholdBredde - 16 - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth

$makroer = Hent-Makroliste
$metadata = Hent-Metadata
$tooltipFiler = New-Object System.Windows.Forms.ToolTip

function Hent-MakroIkonSti {
    param($Makro)
    $ikonNavn = $null
    if ($metadata.PSObject.Properties.Name -contains $Makro.Navn) {
        $meta = $metadata.($Makro.Navn)
        if ($meta.Ikon) { $ikonNavn = $meta.Ikon }
    }
    if ($ikonNavn) {
        $kandidat = Join-Path $ressursMappe $ikonNavn
        if (Test-Path $kandidat) { return $kandidat }
    }
    if (Test-Path $ikonSti) { return $ikonSti }
    return $null
}

function Vis-MakroVelger {
    # Liten modal picker - ett kort per makro med ikon, brukt bade fra
    # "+ Legg til makro" pa et filkort og kan gjenbrukes andre steder.
    $velgerForm = New-Object System.Windows.Forms.Form
    $velgerForm.Text = "Velg makro å legge til"
    $velgerForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    $velgerForm.Width = 360
    $velgerForm.Height = 420
    $velgerForm.StartPosition = "CenterParent"
    $velgerForm.FormBorderStyle = "FixedDialog"
    $velgerForm.MaximizeBox = $false
    $velgerForm.MinimizeBox = $false
    if (Test-Path $ikonSti) { $velgerForm.Icon = New-Object System.Drawing.Icon($ikonSti) }

    $velgerFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $velgerFlow.Left = 8; $velgerFlow.Top = 8; $velgerFlow.Width = 328; $velgerFlow.Height = 330
    $velgerFlow.AutoScroll = $true
    $velgerFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $velgerFlow.WrapContents = $false

    $script:makroValgtIVelger = $null
    $klikkHandler = {
        $t = $this.Tag
        if (-not $t) { $t = $this.Parent.Tag }
        $script:makroValgtIVelger = $t
        $script:velgerFormRef.Close()
    }
    $script:velgerFormRef = $velgerForm

    foreach ($makro in $makroer) {
        $rad = New-Object System.Windows.Forms.Panel
        $rad.Width = 300; $rad.Height = 48
        $rad.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
        $rad.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $rad.Cursor = [System.Windows.Forms.Cursors]::Hand
        $rad.Tag = $makro

        $picSti = Hent-MakroIkonSti $makro
        $pic = New-Object System.Windows.Forms.PictureBox
        $pic.Left = 6; $pic.Top = 6; $pic.Width = 36; $pic.Height = 36
        $pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $pic.Tag = $makro
        if ($picSti) { $pic.Image = (New-Object System.Drawing.Icon($picSti)).ToBitmap() }

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $makro.Navn + $(if ($makro.Versjon) { " (v$($makro.Versjon))" } else { "" })
        $lbl.Left = 50; $lbl.Top = 14; $lbl.Width = 235; $lbl.Height = 22
        $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $lbl.Tag = $makro

        $rad.Controls.AddRange(@($pic, $lbl))
        $rad.Add_Click($klikkHandler)
        $pic.Add_Click($klikkHandler)
        $lbl.Add_Click($klikkHandler)

        $velgerFlow.Controls.Add($rad)
    }

    $cmdAvbryt = New-Object System.Windows.Forms.Button
    $cmdAvbryt.Text = "Avbryt"
    $cmdAvbryt.Left = 8; $cmdAvbryt.Top = 346; $cmdAvbryt.Width = 90; $cmdAvbryt.Height = 28
    $cmdAvbryt.Add_Click({ $script:makroValgtIVelger = $null; $script:velgerFormRef.Close() })

    $velgerForm.Controls.AddRange(@($velgerFlow, $cmdAvbryt))
    [void]$velgerForm.ShowDialog($form)
    return $script:makroValgtIVelger
}

function Test-MakroUtdatert {
    param($Rad)
    $nyeste = $makroer | Where-Object { $_.Navn -eq $Rad.Makro } | Select-Object -First 1
    if ($nyeste -and $nyeste.Versjon -and $Rad.Versjon) {
        try { return ([version]$Rad.Versjon -lt [version]$nyeste.Versjon) }
        catch { return ($Rad.Versjon -ne $nyeste.Versjon) }
    }
    return $false
}

# Fargeplate for ikon-platen bak hver makro sitt ikon. De fem kjente makroene
# far en EKSPLISITT, unik farge hver (matchet mot fargetonen i deres eget
# .ico-ikon), i stedet for a stole pa et hash-oppslag som kan kollidere naar
# antall makroer naermer seg antall farger i paletten (skjedde i praksis med
# den gamle 4-fargers paletten og 5 makroer). Nye, ikke-eksplisitt tildelte
# makroer faller tilbake til et deterministisk (ikke GetHashCode() - den er
# tilfeldig randomisert per prosess i .NET av sikkerhetsgrunner) hash-oppslag
# over hele paletten.
$script:IkonBakgrunner = @(
    [System.Drawing.Color]::FromArgb(237, 233, 254),  # lilla   - Kolonnevelger
    [System.Drawing.Color]::FromArgb(219, 234, 254),  # blaa    - Mail-utsender
    [System.Drawing.Color]::FromArgb(220, 252, 231),  # gronn   - Kontaktsentralen
    [System.Drawing.Color]::FromArgb(254, 243, 199),  # gul     - Nettskjema-henter
    [System.Drawing.Color]::FromArgb(204, 251, 241),  # turkis  - Makromeny
    [System.Drawing.Color]::FromArgb(254, 226, 226)   # rosa    - Statistikkern
)
$script:IkonBakgrunnPerMakro = @{
    "Kolonnevelger"     = 0
    "Mail-utsender"     = 1
    "Kontaktsentralen"  = 2
    "Nettskjema-henter" = 3
    "Makromeny"         = 4
    "Statistikkern"     = 5
}
function Hent-IkonBakgrunn {
    param([string]$Navn)
    if ($script:IkonBakgrunnPerMakro.ContainsKey($Navn)) {
        return $script:IkonBakgrunner[$script:IkonBakgrunnPerMakro[$Navn]]
    }
    $sum = 0
    foreach ($ch in $Navn.ToCharArray()) { $sum += [int]$ch }
    return $script:IkonBakgrunner[$sum % $script:IkonBakgrunner.Count]
}

# ---- Header: ikon + tittel + versjon ----

$headerH = 64
$tabH = 40
$bottomH = 60
$contentTop = $headerH + $tabH
$contentHeight = $form.ClientSize.Height - $contentTop - $bottomH

$pnlHeaderIkon = New-Object System.Windows.Forms.Panel
$pnlHeaderIkon.Left = $sideMargin; $pnlHeaderIkon.Top = 16; $pnlHeaderIkon.Width = 32; $pnlHeaderIkon.Height = 32
$pnlHeaderIkon.BackColor = $script:FargePrimaer
Set-RundedeHjorner -Ctrl $pnlHeaderIkon -Radius 8
if (Test-Path $ikonSti) {
    $picHeaderIkon = New-Object System.Windows.Forms.PictureBox
    $picHeaderIkon.Width = 20; $picHeaderIkon.Height = 20; $picHeaderIkon.Left = 6; $picHeaderIkon.Top = 6
    $picHeaderIkon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $picHeaderIkon.Image = (New-Object System.Drawing.Icon($ikonSti)).ToBitmap()
    $pnlHeaderIkon.Controls.Add($picHeaderIkon)
}

$lblAppTittel = New-Object System.Windows.Forms.Label
$lblAppTittel.Text = "Excel Macro Installer"
$lblAppTittel.AutoSize = $true
$lblAppTittel.Left = $sideMargin + 44; $lblAppTittel.Top = 20
$lblAppTittel.Font = New-Object System.Drawing.Font("Segoe UI", 12.5, [System.Drawing.FontStyle]::Bold)
$lblAppTittel.ForeColor = $script:FargeTittel

# Bredden males eksplisitt (TextRenderer, ikke AutoSize) siden Text/Font kan
# endres igjen senere (selvoppdaterings-varselet) og .Left ma regnes ut pa
# nytt da uansett for a forbli hoyrejustert - se Sett-VersjonToppTekst.
function Sett-VersjonToppTekst {
    param([string]$Tekst, [System.Drawing.Font]$Font, [System.Drawing.Color]$Farge)
    $script:lblVersjonTopp.Text = $Tekst
    $script:lblVersjonTopp.Font = $Font
    $script:lblVersjonTopp.ForeColor = $Farge
    $maalt = [System.Windows.Forms.TextRenderer]::MeasureText($Tekst, $Font)
    $script:lblVersjonTopp.Width = $maalt.Width + 4
    $script:lblVersjonTopp.Height = $maalt.Height
    $script:lblVersjonTopp.Left = $form.ClientSize.Width - $sideMargin - $script:lblVersjonTopp.Width
}

$lblVersjonTopp = New-Object System.Windows.Forms.Label
$lblVersjonTopp.Top = 24
Sett-VersjonToppTekst -Tekst "v$LauncherVersion" -Font (New-Object System.Drawing.Font("Segoe UI", 9)) -Farge $script:FargeMuted

# ---- Fane-header: to klikkbare "faner" med understrek under den aktive ----

$fontFaneAktiv = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)
$fontFaneInaktiv = New-Object System.Drawing.Font("Segoe UI", 10.5)

$lblTabFiler = New-Object System.Windows.Forms.Label
$lblTabFiler.Text = "Installerte filer"
$lblTabFiler.AutoSize = $true
$lblTabFiler.Left = $sideMargin; $lblTabFiler.Top = $headerH + 8
$lblTabFiler.Cursor = [System.Windows.Forms.Cursors]::Hand

$lblTabMakroer = New-Object System.Windows.Forms.Label
$lblTabMakroer.Text = "Makroer"
$lblTabMakroer.AutoSize = $true
$lblTabMakroer.Top = $headerH + 8
$lblTabMakroer.Cursor = [System.Windows.Forms.Cursors]::Hand

$lblTabKunnskap = New-Object System.Windows.Forms.Label
$lblTabKunnskap.Text = "Kunnskapsbase"
$lblTabKunnskap.AutoSize = $true
$lblTabKunnskap.Top = $headerH + 8
$lblTabKunnskap.Cursor = [System.Windows.Forms.Cursors]::Hand

$pnlFaneUnderstrek = New-Object System.Windows.Forms.Panel
$pnlFaneUnderstrek.Height = 3
$pnlFaneUnderstrek.BackColor = $script:FargePrimaer

$pnlFaneDeler = New-Object System.Windows.Forms.Panel
$pnlFaneDeler.Left = 0; $pnlFaneDeler.Top = $contentTop - 1; $pnlFaneDeler.Width = $form.ClientSize.Width; $pnlFaneDeler.Height = 1
$pnlFaneDeler.BackColor = $script:FargeKortKant

# Legges til FOR Sett-AktivFane noensinne kalles - AutoSize-bredden pa
# fane-labelene (brukt til a plassere understreken og nabofanen) males
# ikke riktig for kontrollen faktisk er i et Controls-sett (se samme
# lardom i Ny-BryttLabel-kommentaren andre steder i denne fila).
$form.Controls.AddRange(@($pnlHeaderIkon, $lblAppTittel, $lblVersjonTopp, $lblTabFiler, $lblTabMakroer, $lblTabKunnskap, $pnlFaneUnderstrek, $pnlFaneDeler))

# ---- De tre fane-innholdspanelene (erstatter nativ TabControl helt, for a
# fa full kontroll over det flate, uten-kant utseendet i mockupen) ----

$pnlFaneFiler = New-Object System.Windows.Forms.Panel
$pnlFaneFiler.Left = 0; $pnlFaneFiler.Top = $contentTop; $pnlFaneFiler.Width = $form.ClientSize.Width; $pnlFaneFiler.Height = $contentHeight

$pnlFaneMakroer = New-Object System.Windows.Forms.Panel
$pnlFaneMakroer.Left = 0; $pnlFaneMakroer.Top = $contentTop; $pnlFaneMakroer.Width = $form.ClientSize.Width; $pnlFaneMakroer.Height = $contentHeight
$pnlFaneMakroer.Visible = $false

$pnlFaneKunnskap = New-Object System.Windows.Forms.Panel
$pnlFaneKunnskap.Left = 0; $pnlFaneKunnskap.Top = $contentTop; $pnlFaneKunnskap.Width = $form.ClientSize.Width; $pnlFaneKunnskap.Height = $contentHeight
$pnlFaneKunnskap.Visible = $false

function Sett-AktivFane {
    param([string]$Navn)
    $script:aktivFane = $Navn
    $alleTabber = @(
        @{ Lbl = $lblTabFiler; Pnl = $pnlFaneFiler; Navn = "Filer" },
        @{ Lbl = $lblTabMakroer; Pnl = $pnlFaneMakroer; Navn = "Makroer" },
        @{ Lbl = $lblTabKunnskap; Pnl = $pnlFaneKunnskap; Navn = "Kunnskap" }
    )
    $forrigeHoyre = $sideMargin
    $aktivLbl = $null
    foreach ($t in $alleTabber) {
        $t.Lbl.Left = $forrigeHoyre
        $erAktiv = ($t.Navn -eq $Navn)
        $t.Lbl.Font = if ($erAktiv) { $fontFaneAktiv } else { $fontFaneInaktiv }
        $t.Lbl.ForeColor = if ($erAktiv) { $script:FargeTittel } else { $script:FargeMuted }
        $t.Pnl.Visible = $erAktiv
        if ($erAktiv) { $aktivLbl = $t.Lbl }
        $forrigeHoyre = $t.Lbl.Right + 28
    }
    $pnlFaneUnderstrek.Left = $aktivLbl.Left
    $pnlFaneUnderstrek.Width = $aktivLbl.Width
    $pnlFaneUnderstrek.Top = $aktivLbl.Bottom + 4

    # Kunnskapsbase-fanen bygges FORST na den faktisk vises forste gang
    # (ikke ved oppstart) - unngar a starte WebView2 (relativt tregt a
    # initialisere) for brukeren i det hele tatt har bedt om a se den.
    if ($Navn -eq "Kunnskap" -and -not $script:kunnskapFaneBygget) {
        Bygg-KunnskapFane
        $script:kunnskapFaneBygget = $true
    }
}
$lblTabFiler.Add_Click({ Sett-AktivFane -Navn "Filer" })
$lblTabMakroer.Add_Click({ Sett-AktivFane -Navn "Makroer" })
$lblTabKunnskap.Add_Click({ Sett-AktivFane -Navn "Kunnskap" })

# ---- Fane "Kunnskapsbase" - viser Excel VBA Koding.html direkte i EMI via
# WebView2 (ekte Chromium-rendring, samme motor som Edge - den store,
# moderne HTML-siden bruker CSS grid/variabler/morkt-modus-sporring som en
# gammeldags WebBrowser-kontroll (Internet Explorer-motoren) ikke ville
# rendret riktig. Se HTML-sidens egen "Apne punkter"-seksjon for EMI for
# begrunnelsen). Bygges FORST forste gang fanen faktisk apnes (se
# Sett-AktivFane over), ikke ved oppstart. ----
$script:kunnskapFaneBygget = $false

# Venter pa en .NET Task UTEN a fryse/laase UI-tradet - se lange forklaring
# i Bygg-KunnskapFane. Pumper Application.DoEvents() i en lokke i stedet
# for a blokkere blindt, slik at Windows-meldinger (inkl. WebView2 sine
# egne async-fullforinger) fortsatt kan behandles mens vi venter.
function Vent-Task {
    param($Task, [int]$TimeoutMs = 20000)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not $Task.IsCompleted) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 10
        if ($sw.ElapsedMilliseconds -gt $TimeoutMs) {
            throw "Tidsavbrudd - asynkron operasjon ble ikke ferdig innen $TimeoutMs ms."
        }
    }
    return $Task.GetAwaiter().GetResult()
}

function Bygg-KunnskapFane {
    if (-not $script:webView2Tilgjengelig) {
        Vis-KunnskapFallback -Arsak "WebView2-komponentene mangler eller kunne ikke lastes (se Resources\WebView2\)."
        return
    }
    if (-not (Test-Path $kunnskapsbaseSti)) {
        Vis-KunnskapFallback -Arsak "Fant ikke Excel VBA Koding.html i prosjektroten."
        return
    }

    try {
        $wv2 = New-Object Microsoft.Web.WebView2.WinForms.WebView2
        $wv2.Left = 0; $wv2.Top = 0
        $wv2.Width = $pnlFaneKunnskap.Width; $wv2.Height = $pnlFaneKunnskap.Height
        $wv2.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

        $lblLaster = New-Object System.Windows.Forms.Label
        $lblLaster.Text = "Laster kunnskapsbasen ..."
        $lblLaster.AutoSize = $true
        $lblLaster.Left = $sideMargin; $lblLaster.Top = 16
        $lblLaster.ForeColor = $script:FargeUndertittel
        $pnlFaneKunnskap.Controls.Add($lblLaster)

        # Egen brukerdata-mappe under Data\ - WebView2 MA ha en skrivbar
        # mappe for sin egen profil/buffer (cache, cookies for selve
        # visningen - ingen ekte nettleserdata av interesse her, siden
        # siden er en lokal fil uten innlogging).
        $userDataMappe = Join-Path $dataMappe "WebView2UserData"
        if (-not (Test-Path $userDataMappe)) { New-Item -ItemType Directory -Path $userDataMappe | Out-Null }
        # $null for miljo-valgene (siste parameter) - de er kun for avanserte
        # scenarier (egne kommandolinjeargumenter til Chromium-motoren, ikke
        # nodvendig her). New-Object PA CoreWebView2EnvironmentOptions feilet
        # med "Cannot find an appropriate constructor" under testing
        # 2026-09-22 (PowerShell sin refleksjonsbaserte New-Object fant ikke
        # en brukbar konstruktor for denne COM/managed-hybridtypen) - $null
        # er baade enklere og det som faktisk fungerer.
        #
        # VIKTIG: rar .GetAwaiter().GetResult() her (en ren, meldingslopeblind
        # blokkering) LASER SEG FOR ALLTID - bekreftet ved faktisk testing
        # 2026-09-22. WebView2 sin async-fullforing skjer via Windows sin
        # vanlige meldingskо - og en .GetResult()-blokkering av UI-tradet
        # pumper ALDRI meldinger mens den venter, sa fullforingen kan aldri
        # naa fram til det samme tradet den venter pa. Vent-Task under pumper
        # Application.DoEvents() i en lokke i stedet - samme grunnleggende
        # knep som brukes for a vente pa synkrone COM-kall andre steder i
        # kodebasen uten a fryse UI-et.
        $envTask = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync($null, $userDataMappe, $null)
        $env = Vent-Task -Task $envTask
        $ensureTask = $wv2.EnsureCoreWebView2Async($env)
        Vent-Task -Task $ensureTask | Out-Null

        # Haakon 2026-09-24: teksten i kunnskapsbasen var for stor i forhold
        # til resten av EMI - 80 % zoom gir omtrent samme tekststorrelse.
        $wv2.ZoomFactor = 0.8
        $pnlFaneKunnskap.Controls.Remove($lblLaster)
        $pnlFaneKunnskap.Controls.Add($wv2)
        $wv2.CoreWebView2.Navigate("file:///" + ($kunnskapsbaseSti -replace '\\', '/'))
    } catch {
        try { $pnlFaneKunnskap.Controls.Clear() } catch {}
        Vis-KunnskapFallback -Arsak ("Kunne ikke starte WebView2: " + $_.Exception.Message)
    }
}

function Vis-KunnskapFallback {
    param([string]$Arsak)
    $pnlFaneKunnskap.Controls.Clear()

    $lbl = Ny-BryttLabel -Tekst "Kunnskapsbasen kunne ikke vises direkte i EMI.`r`n$Arsak" -Bredde ($script:innholdBredde - 20)
    $lbl.Left = $sideMargin; $lbl.Top = 20
    $lbl.ForeColor = $script:FargeUndertittel
    $pnlFaneKunnskap.Controls.Add($lbl)

    $cmdApne = Ny-Knapp -Tekst "Åpne i nettleser" -Stil "primaer" -Bredde 180 -Hoyde 34
    $cmdApne.Left = $sideMargin; $cmdApne.Top = $lbl.Top + $lbl.Height + 16
    $cmdApne.Add_Click({
        if (Test-Path $kunnskapsbaseSti) { Start-Process $kunnskapsbaseSti }
        else { [System.Windows.Forms.MessageBox]::Show("Fant ikke Excel VBA Koding.html.", "Excel Macro Installer", "OK", "Warning") | Out-Null }
    })
    $pnlFaneKunnskap.Controls.Add($cmdApne)
}

# ---- Bunnlinje: status + "Start programmet på nytt" (ingen egen Lukk-
# knapp lenger - Håkon onsker heller a alltid ha restart-knappen liggende
# klar, siden den ogsa er nyttig etter endringer i Macros-mappa som ikke
# nodvendigvis trigger versjonsvarselet under). Bygges FOR Last-FilerFane
# kalles - den skriver til $lblStatusBunn.Text, sa den kontrollen ma finnes
# for det forste kallet.

$pnlBunnDeler = New-Object System.Windows.Forms.Panel
$pnlBunnDeler.Left = 0; $pnlBunnDeler.Top = $form.ClientSize.Height - $bottomH; $pnlBunnDeler.Width = $form.ClientSize.Width; $pnlBunnDeler.Height = 1
$pnlBunnDeler.BackColor = $script:FargeKortKant

$lblStatusBunn = New-Object System.Windows.Forms.Label
$lblStatusBunn.Left = $sideMargin; $lblStatusBunn.Top = $form.ClientSize.Height - $bottomH + 18; $lblStatusBunn.Width = 500; $lblStatusBunn.Height = 24
$lblStatusBunn.ForeColor = $script:FargeUndertittel

$cmdStartNytt = Ny-Knapp -Tekst "Start programmet på nytt" -Stil "outline" -Bredde 210 -Hoyde 34
$cmdStartNytt.Left = $form.ClientSize.Width - $sideMargin - 210; $cmdStartNytt.Top = $form.ClientSize.Height - $bottomH + 13
$cmdStartNytt.Add_Click({
    try {
        if ($erKompilert -and $script:egenExeSti) { Start-Process $script:egenExeSti }
        else { Start-Process powershell.exe -ArgumentList @("-ExecutionPolicy", "Bypass", "-File", ('"' + $PSCommandPath + '"')) }
    } catch { }
    $form.Close()
})

try {
    # Sjekken gir bare mening nar dette faktisk er den kompilerte .exe-en -
    # kjort som ra .ps1 (f.eks. under testing) er "MainModule" powershell.exe
    # sin egen fil, som naturlig nok alltid vil ha en annen versjon enn
    # $LauncherVersion og ville gitt et falskt varsel.
    if ($erKompilert) {
        $egenExeSti = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($egenExeSti -and (Test-Path $egenExeSti)) {
            $friskVersjon = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($egenExeSti).FileVersion
            if ($friskVersjon -and $friskVersjon -ne $LauncherVersion) {
                Sett-VersjonToppTekst -Tekst "v$LauncherVersion — ny versjon klar" -Font (New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)) -Farge $script:FargeOransjePrikk
            }
        }
    }
} catch { }

# ---- Oppdatering fra GitHub: knapp + sjekk ved oppstart + installasjon ----
# Sjekken gjores i bakgrunnen (HttpClient-Task som en WinForms-Timer poller -
# INGEN PowerShell-scriptblokk kjores pa et annet trad, som ellers kan krasje
# runspacet). Svarer ikke GitHub innen 8 s (ingen nett, proxy e.l.) skjer
# ingenting - EMI fungerer som for. Knappen vises kun nar en nyere pakke finnes.
$script:nyPakke = $null
$cmdOppdater = Ny-Knapp -Tekst "Oppdater EMI" -Stil "primaer" -Bredde 150 -Hoyde 34
$cmdOppdater.Left = $cmdStartNytt.Left - 10 - 150; $cmdOppdater.Top = $form.ClientSize.Height - $bottomH + 13
$cmdOppdater.Visible = $false

function Ny-HttpKlient {
    Add-Type -AssemblyName System.Net.Http
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    $klient = New-Object System.Net.Http.HttpClient
    $klient.Timeout = [TimeSpan]::FromSeconds(60)
    $klient.DefaultRequestHeaders.UserAgent.ParseAdd("ExcelMacroInstaller/$LauncherVersion")
    return $klient
}

function Vis-OppdateringTilgjengelig {
    param($Info)
    $script:nyPakke = $Info
    Sett-VersjonToppTekst -Tekst "v$LauncherVersion — ny versjon tilgjengelig" -Font (New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)) -Farge $script:FargeOransjePrikk
    $cmdOppdater.Visible = $true
}

function Start-OppdateringSjekk {
    if (-not $script:oppdateringMulig) { return }
    try {
        $script:oppdKlient = Ny-HttpKlient
        $script:oppdKlient.Timeout = [TimeSpan]::FromSeconds(8)
        # Tidsstempel i URL-en omgar mellomlagring hos proxyer (raw.githubusercontent
        # har i tillegg ca. 5 min egen cache - helt greit her).
        $script:oppdTask = $script:oppdKlient.GetStringAsync($script:OppdateringVersjonUrl + "?t=" + [DateTime]::UtcNow.Ticks)
        $script:oppdTimer = New-Object System.Windows.Forms.Timer
        $script:oppdTimer.Interval = 250
        $script:oppdTimer.Add_Tick({
            if (-not $script:oppdTask.IsCompleted) { return }
            $script:oppdTimer.Stop()
            try {
                if ($script:oppdTask.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion) {
                    $info = $script:oppdTask.Result | ConvertFrom-Json
                    if ($info.versjon -and ([version]$info.versjon -gt $script:lokalPakkeVersjon)) { Vis-OppdateringTilgjengelig -Info $info }
                }
            } catch { }
            try { $script:oppdKlient.Dispose() } catch { }
        })
        $script:oppdTimer.Start()
    } catch { }
}

function Installer-Oppdatering {
    $info = $script:nyPakke
    if (-not $info) { return }
    if (Test-InstallKoAktiv) {
        [System.Windows.Forms.MessageBox]::Show("Vent til makroinstallasjonene som kjorer er ferdige, og prov igjen.", "Oppdater EMI", "OK", "Information") | Out-Null
        return
    }
    $endringer = if ($info.endringer) { "`r`n`r`nNytt i denne versjonen:`r`n" + $info.endringer } else { "" }
    $svar = [System.Windows.Forms.MessageBox]::Show(
        "Versjon $($info.versjon) av Excel Macro Installer er tilgjengelig.$endringer`r`n`r`nEMI laster ned oppdateringen, erstatter programfilene og starter seg selv på nytt. Listen over dine filer og installerte makroer berøres ikke. Excel-filene dine endres ikke - bruk «Oppdater» på hver fil etterpå hvis en makro har fått ny versjon.`r`n`r`nOppdatere nå?",
        "Oppdater EMI", "YesNo", "Question")
    if ($svar -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $cmdOppdater.Enabled = $false
    $lblStatusBunn.Text = "Laster ned oppdatering ..."
    $arbeid = Join-Path ([System.IO.Path]::GetTempPath()) "EMI-oppdatering"
    $utfort = New-Object System.Collections.Generic.List[object]   # for tilbakerulling
    try {
        if (Test-Path $arbeid) { Remove-Item -LiteralPath $arbeid -Recurse -Force }
        New-Item -ItemType Directory -Path $arbeid | Out-Null
        $zipSti = Join-Path $arbeid "pakke.zip"

        # 1) Last ned (Vent-Task pumper meldingskoen, sa vinduet ikke fryser).
        $klient = Ny-HttpKlient
        # NB: hent .Result direkte - et byte[] som returneres fra en PowerShell-
        # funksjon (Vent-Task) blir rullet ut til object[] element for element.
        try {
            $lastTask = $klient.GetByteArrayAsync($script:OppdateringZipUrl)
            Vent-Task -Task $lastTask -TimeoutMs 120000 | Out-Null
            $bytes = $lastTask.Result
        } finally { $klient.Dispose() }
        [System.IO.File]::WriteAllBytes($zipSti, $bytes)

        # 2) Pakk ut til en egen mappe og kontroller innholdet FOR noe rores.
        $lblStatusBunn.Text = "Pakker ut ..."
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $utpakket = Join-Path $arbeid "pakke"
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipSti, $utpakket)
        $rot = @(Get-ChildItem -LiteralPath $utpakket -Directory)
        if ($rot.Count -ne 1) { throw "Uventet innhold i nedlastet pakke." }
        $rot = Join-Path $rot[0].FullName "Excel Macro Installer"
        if (-not (Test-Path -LiteralPath $rot)) { throw "Nedlastet pakke mangler mappa 'Excel Macro Installer'." }
        foreach ($kreves in @("Excel Macro Installer.exe", "Macros", "Resources", "versjon.json")) {
            if (-not (Test-Path (Join-Path $rot $kreves))) { throw "Nedlastet pakke mangler '$kreves'." }
        }
        $pakkeVersjon = (Get-Content -Raw -Encoding UTF8 (Join-Path $rot "versjon.json") | ConvertFrom-Json).versjon
        if ([version]$pakkeVersjon -le $script:lokalPakkeVersjon) { throw "Nedlastet pakke ($pakkeVersjon) er ikke nyere enn installert versjon." }

        # 3) Bytt ut filene. Kun filer som faktisk er endret rores. Brukerdata
        #    (installed-files.json, known-files.json, WebView2UserData) er
        #    aldri med i pakken og blir derfor aldri overskrevet. En fil som er
        #    i bruk (den kjorende exe-en, lastede DLL-er) kan ikke overskrives,
        #    men KAN omdopes - den blir *.gammel og slettes ved neste oppstart.
        $lblStatusBunn.Text = "Installerer versjon $pakkeVersjon ..."
        $kildeFiler = Get-ChildItem -LiteralPath $rot -Recurse -File
        foreach ($fil in $kildeFiler) {
            $rel = $fil.FullName.Substring($rot.Length).TrimStart('\')
            if ($rel -notmatch $script:OppdateringHviteliste) { continue }
            $maal = Join-Path $basisMappe $rel
            if (Test-Path -LiteralPath $maal) {
                $likt = (Get-FileHash -LiteralPath $maal -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $fil.FullName -Algorithm SHA256).Hash
                if ($likt) { continue }
                $gammel = $maal + ".gammel"
                if (Test-Path -LiteralPath $gammel) { Remove-Item -LiteralPath $gammel -Force }
                Move-Item -LiteralPath $maal -Destination $gammel
                $utfort.Add(@{ Maal = $maal; Gammel = $gammel })
            } else {
                $mappe = Split-Path $maal -Parent
                if (-not (Test-Path $mappe)) { New-Item -ItemType Directory -Path $mappe | Out-Null }
                $utfort.Add(@{ Maal = $maal; Gammel = $null })
            }
            Copy-Item -LiteralPath $fil.FullName -Destination $maal -Force
        }

        # 4) Start den nye versjonen og lukk denne.
        $lblStatusBunn.Text = "Ferdig - starter på nytt ..."
        Start-Process (Join-Path $basisMappe "Excel Macro Installer.exe")
        $form.Close()
    } catch {
        # Rull tilbake i omvendt rekkefolge: fjern nye kopier, legg originalene tilbake.
        $feil = $_.Exception.Message
        for ($i = $utfort.Count - 1; $i -ge 0; $i--) {
            $steg = $utfort[$i]
            try {
                if (Test-Path -LiteralPath $steg.Maal) { Remove-Item -LiteralPath $steg.Maal -Force }
                if ($steg.Gammel -and (Test-Path -LiteralPath $steg.Gammel)) { Move-Item -LiteralPath $steg.Gammel -Destination $steg.Maal }
            } catch { }
        }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $cmdOppdater.Enabled = $true
        $lblStatusBunn.Text = "Oppdateringen feilet - ingenting er endret."
        [System.Windows.Forms.MessageBox]::Show("Oppdateringen kunne ikke fullføres, og EMI er satt tilbake slik den var.`r`n`r`nFeil: $feil", "Oppdater EMI", "OK", "Warning") | Out-Null
    }
}

$cmdOppdater.Add_Click({ Installer-Oppdatering })

# ---- Fane "Installerte filer" (den man motes av) ----
# Ett kort per FIL - filnavnet vises stort/fett med hele stien synlig
# (bryter over sa mange linjer den faktisk trenger - Ny-BryttLabel - i
# stedet for a klippes/kuttes av) under. Alle makroene som er installert pa
# akkurat den fila listes som egne linjer, alltid i samme (alfabetiske)
# rekkefolge, hver med "Oppdater" (kjor installeren pa nytt) og "Fjern"
# (kaller makroens EGEN -Uninstall - fjerner faktisk VBA-koden/knappen fra
# fila, ikke bare sporingen). "Legg til fil" lar deg legge til en fil FOR
# du velger noen makro for den; "+ Makro" og "Fjern fil" pa hvert kort
# apner hhv. ikon-velgeren (Vis-MakroVelger) og fjerner KUN sporingen for
# hele fila (rorer aldri selve Excel-fila).

$lblHeadingFiler = New-Object System.Windows.Forms.Label
$lblHeadingFiler.Text = "Installerte filer"
$lblHeadingFiler.AutoSize = $true
$lblHeadingFiler.Left = $sideMargin; $lblHeadingFiler.Top = 16
$lblHeadingFiler.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$lblHeadingFiler.ForeColor = $script:FargeTittel
$pnlFaneFiler.Controls.Add($lblHeadingFiler)

# NB: $lblHeadingFiler.Height males FORST etter at den er lagt til et
# Controls-sett over (AutoSize-hoyden er 0 helt til da) - sa undertittelen
# under IKKE ligger pa en gjettet, for hoy posisjon igjen (samme feil som
# nettopp ble fikset for selve overskriften, na videreflyttet hit).
$lblSubFiler = New-Object System.Windows.Forms.Label
$lblSubFiler.AutoSize = $true
$lblSubFiler.Left = $sideMargin; $lblSubFiler.Top = $lblHeadingFiler.Top + $lblHeadingFiler.Height + 2
$lblSubFiler.ForeColor = $script:FargeUndertittel
$pnlFaneFiler.Controls.Add($lblSubFiler)

# Samme lardom videreflyttet enda et hakk ned: listen under skal starte
# etter undertittelens FAKTISKE (malte) posisjon+hoyde, ikke en gjettet
# konstant - ellers gjentar overlapp-bugen seg her ogsa nar fonten/DPI-en
# gjor undertittelen litt hoyere enn antatt.
$flateInnholdTop = $lblSubFiler.Top + $lblSubFiler.Height + 14

$cmdLeggTilFil = Ny-Knapp -Tekst "Legg til fil" -Stil "primaer" -Bredde 140 -Hoyde 34
$cmdLeggTilFil.Left = $form.ClientSize.Width - $sideMargin - 140; $cmdLeggTilFil.Top = 16
$cmdLeggTilFil.Add_Click({
    $filSti = Velg-ExcelFil
    if ($filSti) {
        Lagre-KjentFil -Fil $filSti
        Last-FilerFane
    }
})

$cmdNyFil = Ny-Knapp -Tekst "Opprett ny makroaktivert Excel" -Stil "gronn" -Hoyde 34
$cmdNyFil.Left = $cmdLeggTilFil.Left - 10 - $cmdNyFil.Width; $cmdNyFil.Top = 16
$cmdNyFil.Add_Click({
    # Lager en tom .xlsm i en EGEN, skjult Excel-instans (aldri en som
    # allerede kjorer hos brukeren), lagrer den der brukeren velger, og
    # legger den til i Filer-fanen.
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = "Opprett ny makroaktivert Excel-fil"
    $dlg.Filter = "Excel-arbeidsbok med makroer (*.xlsm)|*.xlsm"
    $dlg.DefaultExt = "xlsm"
    $dlg.AddExtension = $true
    $dlg.OverwritePrompt = $true
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $nyFilSti = $dlg.FileName

    $excelNy = $null; $bokNy = $null
    $gammelMarkor = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $excelNy = New-Object -ComObject Excel.Application
        $excelNy.Visible = $false
        $excelNy.DisplayAlerts = $false
        $bokNy = $excelNy.Workbooks.Add()
        $bokNy.SaveAs($nyFilSti, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
        $bokNy.Close($false)
        $bokNy = $null
        Lagre-KjentFil -Fil $nyFilSti
        $lblStatusBunn.Text = "Opprettet ny makroaktivert Excel-fil: " + $nyFilSti
        Last-FilerFane
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Klarte ikke å opprette filen:`n`n" + $_.Exception.Message, "Excel Macro Installer", "OK", "Warning") | Out-Null
    } finally {
        try { if ($bokNy) { $bokNy.Close($false) } } catch {}
        try { if ($excelNy) { $excelNy.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelNy) | Out-Null } } catch {}
        $form.Cursor = $gammelMarkor
    }
})

$flowFiler = New-Object System.Windows.Forms.FlowLayoutPanel
$flowFiler.Left = $sideMargin; $flowFiler.Top = $flateInnholdTop; $flowFiler.Width = $innholdBredde; $flowFiler.Height = $contentHeight - $flateInnholdTop - 8
$flowFiler.AutoScroll = $true
$flowFiler.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$flowFiler.WrapContents = $false

$lblIngenFiler = New-Object System.Windows.Forms.Label
$lblIngenFiler.Left = $sideMargin; $lblIngenFiler.Top = $flateInnholdTop; $lblIngenFiler.Width = $innholdBredde; $lblIngenFiler.Height = 24
$lblIngenFiler.Text = "Ingen filer ennå - trykk ""Legg til fil"" eller ""Installer / oppdater"" i Makroer-fanen først."
$lblIngenFiler.ForeColor = $script:FargeUndertittel
$lblIngenFiler.Visible = $false

function Bygg-FilKort {
    param([string]$FilSti, $Rader, [int]$Bredde)

    # Alfabetisk pa makronavn - samme rekkefolge pa tvers av alle filkort,
    # uavhengig av hvilken rekkefolge de faktisk ble installert i.
    $radeneListe = @(@($Rader) | Sort-Object Makro)
    $noenUtdatert = $false
    foreach ($rad in $radeneListe) { if (Test-MakroUtdatert $rad) { $noenUtdatert = $true } }

    $ytre = New-Object System.Windows.Forms.Panel
    $ytre.Width = $Bredde
    $ytre.BackColor = $script:FargeKortKant
    $ytre.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)

    $kort = New-Object System.Windows.Forms.Panel
    $kort.Left = 1; $kort.Top = 1; $kort.Width = $Bredde - 2
    $kort.BackColor = [System.Drawing.Color]::White
    $ytre.Controls.Add($kort)

    $navnTekst = Split-Path $FilSti -Leaf
    $fontNavn = New-Object System.Drawing.Font("Segoe UI", 11.5, [System.Drawing.FontStyle]::Bold)
    $lblNavn = New-Object System.Windows.Forms.Label
    $lblNavn.Text = $navnTekst
    $lblNavn.AutoSize = $true
    $lblNavn.Left = 16; $lblNavn.Top = 14
    $lblNavn.Font = $fontNavn
    $lblNavn.ForeColor = $script:FargeTittel
    $kort.Controls.Add($lblNavn)

    if ($noenUtdatert) {
        $navnMaalt = [System.Windows.Forms.TextRenderer]::MeasureText($navnTekst, $fontNavn)
        $badge = Ny-PillLabel -Tekst "OPPDATERING" -Bakgrunn $script:FargeBadgeBg -Tekstfarge $script:FargeBadgeTekst -FontStorrelse 7.5
        $badge.Left = 16 + $navnMaalt.Width + 12; $badge.Top = 16
        $kort.Controls.Add($badge)
    }

    $cmdFjernFil = Ny-Knapp -Tekst "Fjern fil" -Stil "lenke" -Bredde 70 -Hoyde 26
    $cmdFjernFil.Left = $Bredde - 2 - 16 - 70; $cmdFjernFil.Top = 12
    $cmdFjernFil.Tag = $FilSti
    $cmdFjernFil.Add_Click({
        $filSti2 = $this.Tag
        $svar = [System.Windows.Forms.MessageBox]::Show(
            "Fjerne """ + (Split-Path $filSti2 -Leaf) + """ fra denne oversikten?" + "`n`n" +
            "Dette endrer INGENTING i selve Excel-fila - fjerner bare sporingen her. Bruk ""Fjern"" på en enkelt makro-linje for å faktisk avinstallere fra fila.",
            "Excel Macro Installer", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($svar -eq [System.Windows.Forms.DialogResult]::Yes) {
            Fjern-FilHelt -Fil $filSti2
            Last-FilerFane
        }
    })
    $kort.Controls.Add($cmdFjernFil)

    $cmdLeggTilMakro = Ny-Knapp -Tekst "+ Makro" -Stil "aksent" -Bredde 90 -Hoyde 26
    $cmdLeggTilMakro.Left = $cmdFjernFil.Left - 8 - 90; $cmdLeggTilMakro.Top = 12
    $cmdLeggTilMakro.Tag = $FilSti
    $cmdLeggTilMakro.Add_Click({
        $filSti2 = $this.Tag
        $valgtMakro = Vis-MakroVelger
        if ($valgtMakro) {
            Start-MakroMedFil -Makro $valgtMakro -FilSti $filSti2
            $lblStatusBunn.Text = "Startet " + $valgtMakro.Navn + " mot: " + $filSti2
            Last-FilerFane
        }
    })
    $kort.Controls.Add($cmdLeggTilMakro)

    # Hele stien, ordbrutt over sa mange linjer den faktisk trenger (Ny-
    # BryttLabel maler den dynamisk) i stedet for avkuttet med "..." eller
    # klippet av en for lav fast boks - pluss verktoytips som ekstra
    # sikkerhetsnett. Toppen regnes ut fra $lblNavn sin FAKTISKE malte
    # hoyde (ikke en gjettet fast Top=40) - den var for lav og overlappet
    # filnavnet pa 2026-09-10 (rapportert av Håkon, bekreftet via UI
    # Automation: filnavn Bottom=565 vs sti Top=560).
    $lblSti = Ny-BryttLabel -Tekst $FilSti -Bredde ($Bredde - 32) -Font (New-Object System.Drawing.Font("Segoe UI", 8))
    $lblSti.Left = 16; $lblSti.Top = $lblNavn.Top + $lblNavn.Height + 6
    $lblSti.ForeColor = $script:FargeMuted
    $tooltipFiler.SetToolTip($lblSti, $FilSti)
    $kort.Controls.Add($lblSti)

    # NB: $lblSti.Height males FORST etter at den faktisk er lagt til et
    # Controls-sett over (AutoSize-hoyden er 0 helt til da) - se lignende
    # kommentar i Bygg-MakroKort.
    $y = $lblSti.Top + $lblSti.Height + 12

    # Fil-niva-knapper (2026-09-24, Haakons onske): "Oppdater alle makroer"
    # (grayet ut nar ingen makro pa fila har ny versjon) og "Reinstaller alle
    # makroer" (kjorer ALLE om igjen med -Force, og far dermed ogsa med
    # nyeste versjon hvis en oppdatering finnes). Bare nar fila faktisk har
    # minst en installert makro.
    if ($radeneListe.Count -gt 0) {
        $cmdOppdaterAlle = Ny-Knapp -Tekst "Oppdater alle makroer" -Stil "primaer" -Hoyde 28
        $cmdOppdaterAlle.Left = 16; $cmdOppdaterAlle.Top = $y
        $cmdOppdaterAlle.Tag = [PSCustomObject]@{ Fil = $FilSti; Rader = @($radeneListe | Where-Object { Test-MakroUtdatert $_ }) }
        if (-not $noenUtdatert) {
            $cmdOppdaterAlle.Enabled = $false
            $cmdOppdaterAlle.BackColor = [System.Drawing.Color]::FromArgb(236, 237, 241)
            $cmdOppdaterAlle.ForeColor = [System.Drawing.Color]::FromArgb(165, 170, 180)
        }
        $cmdOppdaterAlle.Add_Click({
            $info = $this.Tag
            Start-InstallKo -FilSti $info.Fil -Rader $info.Rader -Tving $false
        })
        $kort.Controls.Add($cmdOppdaterAlle)

        $cmdReinstallerAlle = Ny-Knapp -Tekst "Reinstaller alle makroer" -Stil "outline" -Hoyde 28
        $cmdReinstallerAlle.Left = $cmdOppdaterAlle.Right + 10; $cmdReinstallerAlle.Top = $y
        $cmdReinstallerAlle.Tag = [PSCustomObject]@{ Fil = $FilSti; Rader = @($radeneListe) }
        $cmdReinstallerAlle.Add_Click({
            $info = $this.Tag
            $svar = [System.Windows.Forms.MessageBox]::Show(
                "Reinstallere ALLE makroer på """ + (Split-Path $info.Fil -Leaf) + """ med nyeste versjon?" + "`n`n" +
                "Innstillinger beholdes (samme som når du trykker Oppdater på en enkelt makro).",
                "Excel Macro Installer", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($svar -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-InstallKo -FilSti $info.Fil -Rader $info.Rader -Tving $true
            }
        })
        $kort.Controls.Add($cmdReinstallerAlle)
        $y = $y + 28 + 10
    }

    foreach ($rad in $radeneListe) {
        $utdatert = Test-MakroUtdatert $rad
        $makroNyeste = $makroer | Where-Object { $_.Navn -eq $rad.Makro } | Select-Object -First 1

        $tidspunktTekst = $rad.Tidspunkt
        try { $tidspunktTekst = (Get-Date $rad.Tidspunkt).ToString("dd.MM.yyyy HH:mm") } catch { }

        $ikonSti2 = Hent-MakroIkonSti ([PSCustomObject]@{ Navn = $rad.Makro })
        $ikonKvadrat = Ny-IkonKvadrat -IkonSti $ikonSti2 -Bakgrunn (Hent-IkonBakgrunn $rad.Makro) -Storrelse 36
        $ikonKvadrat.Left = 16; $ikonKvadrat.Top = $y + 8
        $kort.Controls.Add($ikonKvadrat)

        # Alt pa EN linje, vertikalt sentrert mot ikonet (2026-09-24, Haakon:
        # "mye luft mellom tittel og sist installert"): status-prikk, navn,
        # versjonspille, evt. "-> tilgjengelig" og "Installert ..." ligger
        # side om side i stedet for stablet over hverandre.
        $radMidt = $ikonKvadrat.Top + [int]($ikonKvadrat.Height / 2)

        $prikkFarge = if ($utdatert) { $script:FargeOransjePrikk } else { $script:FargeGronnPrikk }
        $prikk = Ny-StatusPrikk -Farge $prikkFarge -Storrelse 9
        $prikk.Left = $ikonKvadrat.Right + 10; $prikk.Top = $radMidt - 4
        $kort.Controls.Add($prikk)

        $fontMakroNavn = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $navnMaalt2 = [System.Windows.Forms.TextRenderer]::MeasureText($rad.Makro, $fontMakroNavn)
        $lblMakroNavn = New-Object System.Windows.Forms.Label
        $lblMakroNavn.Text = $rad.Makro
        $lblMakroNavn.AutoSize = $true
        $lblMakroNavn.Left = $prikk.Right + 8; $lblMakroNavn.Top = $radMidt - [int]($navnMaalt2.Height / 2)
        $lblMakroNavn.Font = $fontMakroNavn
        $lblMakroNavn.ForeColor = $script:FargeTittel
        $kort.Controls.Add($lblMakroNavn)

        $versjonPill = Ny-PillLabel -Tekst "v$($rad.Versjon)" -Bakgrunn $script:FargePillBg -Tekstfarge $script:FargePillTekst -Fet $false -FontStorrelse 8
        $versjonPill.Left = $lblMakroNavn.Left + $navnMaalt2.Width + 6; $versjonPill.Top = $radMidt - [int]($versjonPill.Height / 2)
        $kort.Controls.Add($versjonPill)
        $hoyreKant = $versjonPill.Right

        $fontLitenFet = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
        $fontLiten = New-Object System.Drawing.Font("Segoe UI", 8.5)
        if ($utdatert -and $makroNyeste) {
            $pilTekst = "→ v$($makroNyeste.Versjon) tilgjengelig"
            $pilMaalt = [System.Windows.Forms.TextRenderer]::MeasureText($pilTekst, $fontLitenFet)
            $lblPil = New-Object System.Windows.Forms.Label
            $lblPil.Text = $pilTekst
            $lblPil.AutoSize = $true
            $lblPil.Left = $hoyreKant + 8; $lblPil.Top = $radMidt - [int]($pilMaalt.Height / 2)
            $lblPil.ForeColor = $script:FargeOransjePrikk
            $lblPil.Font = $fontLitenFet
            $kort.Controls.Add($lblPil)
            $hoyreKant = $lblPil.Left + $pilMaalt.Width
        }

        # "Installert ..." ved siden av, med mindre det ikke er plass foran
        # knappene lenger til hoyre (smalt vindu) - da faller den tilbake til
        # en egen linje UNDER navnet, sa den aldri overlapper knappene.
        $installertTekst = "Installert " + $tidspunktTekst
        $installertMaalt = [System.Windows.Forms.TextRenderer]::MeasureText($installertTekst, $fontLiten)
        $knappeVenstre = $Bredde - 2 - 16 - 60 - 10 - 110
        $lblInstallert = New-Object System.Windows.Forms.Label
        $lblInstallert.Text = $installertTekst
        $lblInstallert.AutoSize = $true
        $lblInstallert.ForeColor = $script:FargeUndertittel
        $lblInstallert.Font = $fontLiten
        $installertEgenLinje = (($hoyreKant + 14 + $installertMaalt.Width) -gt ($knappeVenstre - 8))
        if ($installertEgenLinje) {
            $lblInstallert.Left = $lblMakroNavn.Left; $lblInstallert.Top = $lblMakroNavn.Top + $navnMaalt2.Height + 1
        } else {
            $lblInstallert.Left = $hoyreKant + 14; $lblInstallert.Top = $radMidt - [int]($installertMaalt.Height / 2)
        }
        $kort.Controls.Add($lblInstallert)

        $cmdFjernMakro = Ny-Knapp -Tekst "Fjern" -Stil "lenke" -Bredde 60 -Hoyde 30
        $cmdFjernMakro.Left = $Bredde - 2 - 16 - 60; $cmdFjernMakro.Top = $radMidt - 15
        $cmdFjernMakro.Tag = [PSCustomObject]@{ Makro = $rad.Makro; Fil = $FilSti }
        $cmdFjernMakro.Add_Click({
            $info = $this.Tag
            $makro = $makroer | Where-Object { $_.Navn -eq $info.Makro } | Select-Object -First 1
            if (-not $makro) {
                [System.Windows.Forms.MessageBox]::Show("Fant ikke makroen '$($info.Makro)' i Macros-mappa lenger.", "Excel Macro Installer", "OK", "Warning") | Out-Null
                return
            }
            $svar = [System.Windows.Forms.MessageBox]::Show(
                "Fjerne " + $makro.Navn + " FRA SELVE EXCEL-FILA """ + (Split-Path $info.Fil -Leaf) + """?" + "`n`n" +
                "Dette fjerner VBA-koden og knappen fra arbeidsboken - ikke bare sporingen her. Innstillinger beholdes (en senere installasjon gjenbruker dem).",
                "Excel Macro Installer - bekreft fjerning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($svar -eq [System.Windows.Forms.DialogResult]::Yes) {
                Fjern-MakroFraFil -Makro $makro -FilSti $info.Fil -StatusLabel $lblStatusBunn
                Last-FilerFane
            }
        })
        $kort.Controls.Add($cmdFjernMakro)

        $cmdRad = Ny-Knapp -Tekst $(if ($utdatert) { "Oppdater" } else { "Oppdatert" }) -Stil $(if ($utdatert) { "primaer" } else { "outline" }) -Bredde 110 -Hoyde 32
        $cmdRad.Left = $cmdFjernMakro.Left - 10 - 110; $cmdRad.Top = $radMidt - 16
        $cmdRad.Tag = [PSCustomObject]@{ Makro = $rad.Makro; Fil = $FilSti }
        $cmdRad.Add_Click({
            $info = $this.Tag
            $makro = $makroer | Where-Object { $_.Navn -eq $info.Makro } | Select-Object -First 1
            if (-not $makro) {
                [System.Windows.Forms.MessageBox]::Show("Fant ikke makroen '$($info.Makro)' i Macros-mappa lenger.", "Excel Macro Installer", "OK", "Warning") | Out-Null
                return
            }
            Start-MakroMedFil -Makro $makro -FilSti $info.Fil
            $lblStatusBunn.Text = "Startet " + $makro.Navn + " mot: " + $info.Fil
            Last-FilerFane
        })
        $kort.Controls.Add($cmdRad)

        # Radens hoyde avledes av det FAKTISK malte bunnpunktet til ALLE
        # kontrollene i rada (ikon-plate, versjonspille, evt. "-> tilgjengelig"-
        # teksten, og "Installert ..."-teksten) - ikke bare et par av dem -
        # i stedet for en gjettet fast konstant. Forrige runde av denne
        # fiksen glemte versjonspillen og pil-teksten, som fortsatt klippet
        # pa "utdatert"-rader (de har begge i tillegg til de to andre) -
        # rapportert av Håkon 2026-09-10.
        $radBunn = $ikonKvadrat.Top + $ikonKvadrat.Height
        $radBunn = [Math]::Max($radBunn, $versjonPill.Top + $versjonPill.Height)
        $radBunn = [Math]::Max($radBunn, $lblInstallert.Top + $installertMaalt.Height)
        if ($utdatert -and $makroNyeste) {
            $radBunn = [Math]::Max($radBunn, $lblPil.Top + $pilMaalt.Height)
        }
        $y = [Math]::Max($y + 52, $radBunn + 10)
    }

    $kort.Height = $y + 10
    $ytre.Height = $kort.Height + 2
    return $ytre
}

function Last-FilerFane {
    $flowFiler.Controls.Clear()
    $register = Les-InstallertRegister
    $kjente = Hent-KjenteFiler

    # Union: filer eksplisitt lagt til via "+ Legg til fil" OG filer som har
    # minst én faktisk installasjon i registeret - en fil kan altsa vises
    # her selv om den (ennå) ikke har noen makro installert.
    $alleFiler = New-Object System.Collections.Generic.List[string]
    foreach ($f in $kjente) { if ($alleFiler -notcontains $f) { $alleFiler.Add($f) } }
    foreach ($rad in $register) { if ($alleFiler -notcontains $rad.Fil) { $alleFiler.Add($rad.Fil) } }

    $antallUtdatert = 0
    foreach ($rad in $register) { if (Test-MakroUtdatert $rad) { $antallUtdatert++ } }

    $lblSubFiler.Text = "$($alleFiler.Count) fil$(if ($alleFiler.Count -ne 1) { 'er' })" + $(if ($antallUtdatert -gt 0) { " · $antallUtdatert makro$(if ($antallUtdatert -ne 1) { 'er' }) har ny versjon" } else { "" })
    $lblStatusBunn.Text = if ($antallUtdatert -gt 0) { "$antallUtdatert oppdatering$(if ($antallUtdatert -ne 1) { 'er' }) tilgjengelig" } else { "" }

    $lblIngenFiler.Visible = ($alleFiler.Count -eq 0)
    $flowFiler.Visible = ($alleFiler.Count -gt 0)
    if ($alleFiler.Count -eq 0) { return }

    $grupper = @{}
    foreach ($rad in $register) {
        if (-not $grupper.ContainsKey($rad.Fil)) { $grupper[$rad.Fil] = New-Object System.Collections.Generic.List[object] }
        $grupper[$rad.Fil].Add($rad)
    }

    # Nyeste aktivitet forst; en fil uten noen installasjon ennå (bare "kjent")
    # har ingen tidspunkt å sortere på og havner naturlig sist.
    $sortert = $alleFiler | Sort-Object -Descending -Property {
        if ($grupper.ContainsKey($_)) {
            ($grupper[$_] | ForEach-Object { $_.Tidspunkt } | Sort-Object -Descending | Select-Object -First 1)
        } else {
            ""
        }
    }

    foreach ($filSti in $sortert) {
        $rader = if ($grupper.ContainsKey($filSti)) { $grupper[$filSti] } else { @() }
        $flowFiler.Controls.Add((Bygg-FilKort -FilSti $filSti -Rader $rader -Bredde $kortBredde))
    }
}

$pnlFaneFiler.Controls.AddRange(@($cmdNyFil, $cmdLeggTilFil, $flowFiler, $lblIngenFiler))

# ---- Fane "Makroer" ----

$lblHeadingMakroer = New-Object System.Windows.Forms.Label
$lblHeadingMakroer.Text = "Tilgjengelige makroer"
$lblHeadingMakroer.AutoSize = $true
$lblHeadingMakroer.Left = $sideMargin; $lblHeadingMakroer.Top = 16
$lblHeadingMakroer.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$lblHeadingMakroer.ForeColor = $script:FargeTittel
$pnlFaneMakroer.Controls.Add($lblHeadingMakroer)

# Samme lardom som i Filer-fanen: undertittel og liste plasseres ut fra
# FAKTISK malt hoyde pa det som kommer over, ikke gjettede konstanter.
$lblSubMakroer = New-Object System.Windows.Forms.Label
$lblSubMakroer.Text = "Installer i en valgt fil, eller oppdater til nyeste versjon."
$lblSubMakroer.AutoSize = $true
$lblSubMakroer.Left = $sideMargin; $lblSubMakroer.Top = $lblHeadingMakroer.Top + $lblHeadingMakroer.Height + 2
$lblSubMakroer.ForeColor = $script:FargeUndertittel
$pnlFaneMakroer.Controls.Add($lblSubMakroer)

$flateMakroerTop = $lblSubMakroer.Top + $lblSubMakroer.Height + 14

$flow = New-Object System.Windows.Forms.FlowLayoutPanel
$flow.Left = $sideMargin; $flow.Top = $flateMakroerTop; $flow.Width = $innholdBredde; $flow.Height = $contentHeight - $flateMakroerTop - 8
$flow.AutoScroll = $true
$flow.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$flow.WrapContents = $false

function Bygg-MakroKort {
    param($Makro, $Metadata, [int]$Bredde)

    $beskrivelse = "Ingen beskrivelse tilgjengelig ennå."
    if ($Metadata.PSObject.Properties.Name -contains $Makro.Navn) {
        $meta = $Metadata.($Makro.Navn)
        if ($meta.Beskrivelse) { $beskrivelse = $meta.Beskrivelse }
    }

    $ytre = New-Object System.Windows.Forms.Panel
    $ytre.Width = $Bredde
    $ytre.BackColor = $script:FargeKortKant
    $ytre.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)

    $kort = New-Object System.Windows.Forms.Panel
    $kort.Left = 1; $kort.Top = 1; $kort.Width = $Bredde - 2
    $kort.BackColor = [System.Drawing.Color]::White
    $ytre.Controls.Add($kort)

    $ikonKvadrat = Ny-IkonKvadrat -IkonSti (Hent-MakroIkonSti $Makro) -Bakgrunn (Hent-IkonBakgrunn $Makro.Navn) -Storrelse 44
    $ikonKvadrat.Left = 16; $ikonKvadrat.Top = 16
    $kort.Controls.Add($ikonKvadrat)

    $fontMakroNavn = New-Object System.Drawing.Font("Segoe UI", 11.5, [System.Drawing.FontStyle]::Bold)
    $lblNavn = New-Object System.Windows.Forms.Label
    $lblNavn.Text = $Makro.Navn
    $lblNavn.AutoSize = $true
    $lblNavn.Left = $ikonKvadrat.Right + 16; $lblNavn.Top = 14
    $lblNavn.Font = $fontMakroNavn
    $lblNavn.ForeColor = $script:FargeTittel
    $kort.Controls.Add($lblNavn)

    $navnMaalt = [System.Windows.Forms.TextRenderer]::MeasureText($Makro.Navn, $fontMakroNavn)
    if ($Makro.Versjon) {
        $versjonPill = Ny-PillLabel -Tekst "v$($Makro.Versjon)" -Bakgrunn $script:FargePillBg -Tekstfarge $script:FargePillTekst -Fet $false -FontStorrelse 8
        $versjonPill.Left = $lblNavn.Left + $navnMaalt.Width + 6
        $versjonPill.Top = $lblNavn.Top + [int](($navnMaalt.Height - $versjonPill.Height) / 2)
        $kort.Controls.Add($versjonPill)
        $dateLeft = $versjonPill.Right + 10
    } else {
        $dateLeft = $lblNavn.Left + $navnMaalt.Width + 10
    }

    $lblDato = New-Object System.Windows.Forms.Label
    $lblDato.Text = "Sist endret " + $Makro.SistEndret.ToString("dd.MM.yyyy")
    $lblDato.AutoSize = $true
    $datoMaalt = [System.Windows.Forms.TextRenderer]::MeasureText($lblDato.Text, (New-Object System.Drawing.Font("Segoe UI", 8.5)))
    $lblDato.Left = $dateLeft; $lblDato.Top = $lblNavn.Top + [int](($navnMaalt.Height - $datoMaalt.Height) / 2)
    $lblDato.ForeColor = $script:FargeUndertittel
    $lblDato.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $kort.Controls.Add($lblDato)

    # Beskrivelsen males FORST (dynamisk hoyde) - kortets totale hoyde
    # avledes av det faktiske resultatet, ikke en gjettet konstant. NB:
    # AutoSize-hoyden er ikke palitelig for kontrollen faktisk er lagt til
    # et Controls-sett (Height leses som 0 for den) - derfor legges den til
    # kortet FOR $kort.Height leses ut helt til slutt.
    $lblBeskrivelse = Ny-BryttLabel -Tekst $beskrivelse -Bredde ($Bredde - $ikonKvadrat.Right - 210)
    $lblBeskrivelse.Left = $ikonKvadrat.Right + 16; $lblBeskrivelse.Top = $lblNavn.Top + $navnMaalt.Height + 6
    $lblBeskrivelse.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 75)
    $kort.Controls.Add($lblBeskrivelse)

    $cmd = Ny-Knapp -Tekst "Installer / oppdater" -Stil "outline" -Hoyde 34
    $cmd.Left = $Bredde - 2 - 16 - $cmd.Width; $cmd.Top = 16
    $cmd.Tag = $Makro
    $cmd.Add_Click({
        $valgtMakro = $this.Tag
        $filSti = Velg-ExcelFil
        if ($filSti) {
            Start-MakroMedFil -Makro $valgtMakro -FilSti $filSti
            $lblStatusBunn.Text = "Startet " + $valgtMakro.Navn + " mot: " + $filSti
            Last-FilerFane
        }
    })
    $kort.Controls.Add($cmd)

    $kort.Height = [Math]::Max(96, $lblBeskrivelse.Top + $lblBeskrivelse.Height + 18)
    $ytre.Height = $kort.Height + 2
    return $ytre
}

function Last-MakroerFane {
    # Bygges pa nytt (samme monster som Last-FilerFane) i stedet for kun en
    # gang ved oppstart - trengs bade for at "Bredde" skal stemme etter en
    # vindusstorrelse-endring (se resize-handleren under) og gjor fanen
    # konsistent med Filer-fanen sin oppdaterbare oppforsel.
    $flow.Controls.Clear()
    if ($makroer.Count -eq 0) {
        $lblIngen = New-Object System.Windows.Forms.Label
        $lblIngen.Text = "Fant ingen .ps1-installere i Macros-mappa."
        $lblIngen.Width = $script:kortBredde; $lblIngen.Height = 24
        $flow.Controls.Add($lblIngen)
    } else {
        foreach ($makro in $makroer) {
            $flow.Controls.Add((Bygg-MakroKort -Makro $makro -Metadata $metadata -Bredde $script:kortBredde))
        }
    }
}
Last-MakroerFane

$pnlFaneMakroer.Controls.Add($flow)

Last-FilerFane
Sett-AktivFane -Navn "Filer"

$form.Controls.AddRange(@($pnlFaneFiler, $pnlFaneMakroer, $pnlFaneKunnskap, $pnlBunnDeler, $lblStatusBunn, $cmdStartNytt, $cmdOppdater))
$form.Add_Shown({ $form.Activate(); Start-OppdateringSjekk })

# ---- Gjor vinduet faktisk resizable: reposisjoner/bygg om innholdet nar
# storrelsen endres, i stedet for a bare la WinForms strekke det statisk
# plasserte innholdet urort (som ville gitt tomrom/avkuttet innhold). To
# hastigheter: de faste, billige kontrollene (skillelinjer, panelbredder,
# hoyrejusterte knapper/tekst) flyttes LOPENDE under selve draget (Resize),
# mens den tyngre jobben - bygge alle kortene pa nytt med ny bredde - skjer
# bare NAR draget slippes eller vinduet maksimeres (ResizeEnd), for a unnga
# a bygge hundrevis av kontroller pa nytt for hver piksel under draget.
$script:sisteLayoutBredde = $form.ClientSize.Width

function Oppdater-FasteKontroller {
    # [Math]::Max-gulv overalt her: MinimumSize BOR hindre vinduet i a bli sa
    # lite at disse blir negative, men er observert (testing 2026-09-22,
    # Nils) a kunne omgas av DPI-virtualisering ved programmatisk
    # storrelsesendring fra en prosess med annen DPI-bevissthet enn EMI selv
    # (SetWindowPos fra et eksternt testverktoy). Uten gulvet ville en
    # negativ Height gitt paneler som ikke lenger klipper barna sine riktig
    # -> kortinnhold overlappet bunnlinja. Gulvverdiene er bevisst romslige
    # nok til at AutoScroll i flow-panelene uansett tar over gracefully i
    # stedet for a klippe midt i en rad.
    # Gulvet pa 736/720 er ikke vilkarlig: under ca. 720 px kortbredde har
    # ikke makro-raden nok plass til bade den lengste statusteksten ("->
    # vX.Y.Z tilgjengelig" for Nettskjema-henter, lengste makronavnet) OG de
    # hoyrejusterte Oppdater/Fjern-knappene samtidig - bekreftet overlapp
    # visuelt (PrintWindow) under testing 2026-09-22. $form.MinimumSize
    # (820 px klientbredde) skal normalt hindre at vinduet i det hele tatt
    # kommer hit, men gulvet star som et sikkerhetsnett i tillegg.
    $script:innholdBredde = [Math]::Max(736, $form.ClientSize.Width - ($sideMargin * 2))
    $script:kortBredde = [Math]::Max(700, $script:innholdBredde - 16 - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth)
    $nyContentHeight = [Math]::Max(150, $form.ClientSize.Height - $contentTop - $bottomH)

    $pnlFaneDeler.Width = $form.ClientSize.Width
    $pnlFaneFiler.Width = $form.ClientSize.Width; $pnlFaneFiler.Height = $nyContentHeight
    $pnlFaneMakroer.Width = $form.ClientSize.Width; $pnlFaneMakroer.Height = $nyContentHeight
    # WebView2-kontrollen (hvis bygget) folger automatisk via sin egen
    # Anchor til alle fire kanter av pnlFaneKunnskap - trenger ingen egen
    # kode her.
    $pnlFaneKunnskap.Width = $form.ClientSize.Width; $pnlFaneKunnskap.Height = $nyContentHeight

    Sett-VersjonToppTekst -Tekst $script:lblVersjonTopp.Text -Font $script:lblVersjonTopp.Font -Farge $script:lblVersjonTopp.ForeColor
    $cmdLeggTilFil.Left = $form.ClientSize.Width - $sideMargin - $cmdLeggTilFil.Width
    $cmdNyFil.Left = $cmdLeggTilFil.Left - 10 - $cmdNyFil.Width

    $flowFiler.Width = $script:innholdBredde; $flowFiler.Height = [Math]::Max(60, $nyContentHeight - $flateInnholdTop - 8)
    $flow.Width = $script:innholdBredde; $flow.Height = [Math]::Max(60, $nyContentHeight - $flateMakroerTop - 8)

    $pnlBunnDeler.Top = $form.ClientSize.Height - $bottomH; $pnlBunnDeler.Width = $form.ClientSize.Width
    $lblStatusBunn.Top = $form.ClientSize.Height - $bottomH + 18
    $cmdStartNytt.Left = $form.ClientSize.Width - $sideMargin - $cmdStartNytt.Width
    $cmdStartNytt.Top = $form.ClientSize.Height - $bottomH + 13
    $cmdOppdater.Left = $cmdStartNytt.Left - 10 - $cmdOppdater.Width
    $cmdOppdater.Top = $cmdStartNytt.Top
}

$form.Add_Resize({
    if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) { return }
    Oppdater-FasteKontroller

    # Kortene bygges pa nytt med riktig bredde her i Resize selv (IKKE i
    # ResizeEnd) - .NET sitt Form.ResizeEnd/WM_EXITSIZEMOVE viste seg
    # upalitelig a trigge ved programmatisk (ikke-interaktiv musedrag)
    # storrelsesendring under testing 2026-09-22, noe som lot kortene sta
    # igjen i FEIL (gammel) bredde og forarsaket ekte overlapp mot
    # bunnlinja. Bredde-sjekken pa $script:sisteLayoutBredde hindrer
    # unodvendig ombygging nar bare HOYDEN endres (typisk forste tick i et
    # museddrag for bredden faktisk rekker a endre seg).
    if ($form.ClientSize.Width -ne $script:sisteLayoutBredde) {
        $script:sisteLayoutBredde = $form.ClientSize.Width
        Last-FilerFane
        Last-MakroerFane
        Sett-AktivFane -Navn $script:aktivFane
    }
})

[void]$form.ShowDialog()

} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Excel Macro Installer stotte pa en uventet feil:`n`n$($_.Exception.Message)",
        "Excel Macro Installer - feil",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}
