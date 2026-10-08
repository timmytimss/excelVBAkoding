# _testverktoy - Nils sine test- og byggehjelpere

Brukes av Nils (lokal koder-session) til å teste VBA/PowerShell-makroer mot `Prosjektfiler\Testfiler\Testfil.xlsm`
og til å lage ikoner. **Aldri mot Håkons ekte arbeidsbøker** (absolutt personvernregel, se CLAUDE.md).
Skriptene er skrevet for Windows PowerShell 5.1. Kopier gjerne til en egen arbeidsmappe før bruk.

| Fil | Hva den gjør |
|---|---|
| `get-testfil-wb.psm1` | `Import-Module` -> `Get-TestfilExcelApp` / `Get-TestfilWorkbook`. Finner riktig Excel-instans via Running Object Table på filnavnet «Testfil.xlsm» (aldri `GetActiveObject`, som er tvetydig når flere Excel-prosesser kjører). Har sikkerhetssjekk som stopper hvis arbeidsboken ikke heter Testfil.xlsm. |
| `findwin.ps1 -Title "..."` | Skriver HWND (tall) til vindu med EKSAKT tittel, f.eks. `"Statistikkern"`. |
| `screenshot-window.ps1 -Hwnd N -OutPath x.png` | Skjermbilde av ett vindu (PrintWindow). Bildets piksler = koordinater for `click-at.ps1`. |
| `click-at.ps1 -Hwnd N -X px -Y px` | Ekte museklikk relativt til vinduets øvre venstre hjørne (flytter markøren tilbake etterpå). Koordinatene leses fra skjermbildet, IKKE regnes ut fra VBA-punkter (ca. 1,67 px per punkt hos Håkon). |
| `sendkey.ps1 -Hwnd N -Key Down/Up/Enter/Escape` | Sender én tast til et (PID-/tittelverifisert) vindu. |
| `closewin.ps1 -Hwnd N` | Lukker et vindu pent (WM_CLOSE). Bruk dette, ikke kill, for modale skjema. |
| `setup-test-v3.ps1` | Eksempel: lager ark `StatTest` med tabell + fargede rader og en oppsett-rad til Statistikkern. Mal for egne testdata. |
| `write-raw-ico.ps1`, `make-statistikkern-icon2.ps1` | Skriver ekte 32-bit ICO uten `GetHicon()` (som gir gråtonet ikon i automatisering) og lager makro-ikonet + 40x40 BMP-base64 til Makromeny. |

## Standard testløp for en makro-endring

1. Rediger kilden i OneDrive-prosjektet (`Excel Macro Installer\Macros\<Makro>.ps1`). Bump versjonen i VBA-konstanten.
2. Drep gammel testinstans: `Stop-Process -Name EXCEL -Force` (bare når du vet at ingen ekte fil er åpen - sjekk `Get-Process EXCEL` og vindustitler først, spør Håkon ved tvil).
3. Installer i en FERSK, isolert instans:
   `& '.\Excel Macro Installer\Macros\<Makro>.ps1' -NyExcelInstans -Force -Path '...\Prosjektfiler\Testfiler\Testfil.xlsm'` (kjør i bakgrunn hvis kallet kan ta >2 min).
4. **Skjul Excel umiddelbart** (`$app.Visible=$false`, og `$app.VBE.MainWindow.Visible=$false`) - verktøyene kjører på Håkons EKTE maskin, ikke i en sandkasse.
5. Kompileringssjekk: `$app.VBE.CommandBars.FindControl(1,578).Execute()` i try/catch. **Ikke uttømmende** (har gitt «OK» på reelle feil, se HTML-sidens gotcha `g-vba-const-midt-i-modul`).
6. Funksjonstest av rene funksjoner: `$app.Run("modX.Funksjon")` direkte (fungerer). **Ikke** via scratch-moduler som kaller andre moduler (upålitelig), og aldri `Run` på noe som åpner modalt skjema og så forvente korrekt rendring.
7. Skjema-test: åpne via `Run` (i bakgrunnsjobb - den blokkerer til skjemaet lukkes), finn vinduet (`findwin`), ta skjermbilde, klikk med `click-at` på koordinater fra skjermbildet, nytt skjermbilde. Dynamisk innhold bygget i `UserForm_Initialize` rendres ikke alltid via ekstern Run; en ekte klikk-kjede gjør det (se `g-application-run-form-method`).
8. Oppstår VBE-feil: `$app.VBE.ActiveCodePane.GetSelection(...)` + `CodeModule.Lines(n,1)` viser linjen. Klikk OK på dialogen, drep instansen, fiks, start på nytt (en instans i feiltilstand er ubrukelig).
9. Rydd: slett testdata (`StatTest`-ark, oppsett-rader), `$wb.Save()`, `$app.Quit()`, `Stop-Process`. Ikke la vinduer ligge igjen på Håkons skjerm.
10. Commit/push fra `C:\Users\haaklun\Documents\excelVBAkoding-git-nils\` (kopier endrede filer dit først). Oppdater `Excel VBA Koding.html` PROPORSJONALT (se CLAUDE.md). Nye `.ps1`-filer MÅ ha UTF-8 med BOM og LF.

## Kjente fallgruver ved COM-bygde UserForms (detaljer i HTML-siden)

- Egenskaper på `.Font.*` og `CommandButton.Style` kan IKKE settes fra PowerShell-designeren (feiler) - sett dem i VBA (`UserForm_Initialize`).
- Symbolske MSForms-konstanter (`fmButtonStyleGraphical`) er ikke pålitelig tilgjengelige i STANDARDMODULER - bruk tallverdi. I skjemamoduler fungerer de.
- Modul-nivå `Const`/`Dim` MÅ stå før første Sub/Function - også i nye tillegg.
- `Dim x As New Collection` inne i en løkke gjenbruker samme objekt; bruk `Set x = New Collection` per runde.
- `ListObjects`/farger: bruk `Range.DisplayFormat.Interior` for faktisk vist farge (betinget formatering vinner).
- Auto-modus-sikkerhetssjekken svarer av og til «ingen verdikt» på Bash/Edit/Write: vent og prøv igjen senere i stedet for å hamre.
