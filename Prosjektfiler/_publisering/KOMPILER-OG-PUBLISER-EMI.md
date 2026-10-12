# Kompiler og publiser EMI (lokal økt på Håkons PC)

Brukes når `Excel Macro Installer/Installer/Excel Macro Installer.ps1` er endret og `.exe`-en må kompileres på
nytt. Det kan bare gjøres på Windows (ps2exe). Sky-økter endrer kilden og bumper `$LauncherVersion`; en lokal økt
gjør resten med stegene under. Det finnes **ingen fast git-klone** på PC-en: klon til `%TEMP%`, gjør jobben, slett
klonen. Håkons installerte EMI (`Documents\Excel Macro Installer`) skal **ikke** røres direkte – den henter
oppdateringen selv via «Oppdater EMI».

Kjør stegene i Windows PowerShell 5.1, ett og ett, og stopp hvis noe feiler.

```powershell
# 1. Midlertidig klone
$kl = Join-Path $env:TEMP "evk-bygg"
if (Test-Path $kl) { Remove-Item $kl -Recurse -Force }
git clone https://timmytimss@github.com/timmytimss/excelVBAkoding.git $kl
$emi = Join-Path $kl "Excel Macro Installer"
$kilde = Join-Path $emi "Installer\Excel Macro Installer.ps1"

# 2. Versjon fra kilden + syntakssjekk (feiler kallet, er kilden ødelagt - stopp)
$v = ([regex]::Match((Get-Content -Raw -Encoding UTF8 $kilde), '\$LauncherVersion\s*=\s*"([^"]+)"')).Groups[1].Value
$v
[void][ScriptBlock]::Create((Get-Content -Raw -Encoding UTF8 $kilde))

# 3. Kompiler (-noConsole er ESSENSIELT, se gotcha g-emi-w1-terminal-problemet)
Import-Module ps2exe   # mangler den: Install-Module ps2exe -Scope CurrentUser
Invoke-ps2exe -inputFile $kilde -outputFile (Join-Path $emi "Excel Macro Installer.exe") `
    -iconFile (Join-Path $emi "Resources\icon.ico") -noConsole -STA `
    -title "Excel Macro Installer" -company "Håkon" -version $v

# 4. Kontroll: FileVersion skal være lik $v
[System.Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $emi "Excel Macro Installer.exe")).FileVersion
```

5. Røyktest: start `"$emi\Excel Macro Installer.exe"`. Sjekk at ett vindu åpner (ingen konsoll/Terminal ved siden
   av), og lukk det. Klonen regnes som utviklingskopi, så den sjekker ikke etter oppdateringer – det er riktig.

```powershell
# 6. Commit exe-en og publiser (Publiser-EMI sjekker at exe og kilde har samme versjon, øker versjon.json og pusher)
git -C $kl add "Excel Macro Installer/Excel Macro Installer.exe"
git -C $kl -c user.name=timmytimss -c user.email=timmytimss@users.noreply.github.com commit -m "EMI $v`: kompilert exe"
git -C $kl push origin main
powershell -ExecutionPolicy Bypass -File (Join-Path $kl "Prosjektfiler\_publisering\Publiser-EMI.ps1") -Repo $kl -Endringer "<kort tekst om hva som er nytt>"

# 7. Rydd
Remove-Item $kl -Recurse -Force
```

8. Be Håkon starte EMI og trykke «Oppdater EMI». Oppdater versjonshistorikken i `Excel VBA Koding.html`
   (`#mac-emi-historikk`) og `KOORDINERING.md` kort.

## Ventende jobb

Ingen. (Sist: EMI 2.9.1.0, publisert 2026-10-12 som pakke `2026.10.12.2`.)
