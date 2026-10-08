# Beredskapsplan: gjenopprett arbeidsflyten etter PC-bytte eller tap

Oppdatert 2026-10-08 da OneDrive-mappa ble avviklet. **Alt prosjektinnhold ligger på GitHub**, så en ny PC
trenger bare dette:

**Forutsetning:** Håkon har tilgang til GitHub-kontoen `timmytimss` (innlogging timssbooking@gmail.com) og til
Claude. Sørg for en pålitelig gjenopprettingsvei inn i GitHub-kontoen (gjenopprettings-e-post, 2FA-backupkoder) –
det er den ene lenken denne planen ikke kan sikre.

## 1. Arbeidsmappen
```
git clone https://timmytimss@github.com/timmytimss/excelVBAkoding.git C:\Users\<bruker>\Documents\excelVBAkoding
git -C C:\Users\<bruker>\Documents\excelVBAkoding config user.name timmytimss
git -C C:\Users\<bruker>\Documents\excelVBAkoding config user.email timmytimss@users.noreply.github.com
```
Første push åpner en nettleserinnlogging (Git Credential Manager) – logg inn som `timmytimss`.

## 2. Excel Macro Installer (programmet)
Følg «Installere EMI» i repoets `README.md` (last ned zip, fjern blokkering, kopier mappa `Excel Macro Installer`
uten `Installer\` til `C:\Users\<bruker>\Documents\Excel Macro Installer`). Registeret over Håkons filer
(`Data\installed-files.json`/`known-files.json`) er lokalt og finnes ikke på GitHub – mangler det, legg filene til
på nytt med «Legg til fil». Makroinstallatørene oppdager selv det som allerede ligger i arbeidsbøkene.

## 3. Claude-økter
- **Lokal økt** (testing i Excel): start Claude Code i `Documents\excelVBAkoding`. Les `CLAUDE.md` og «NY SESSION? Start her» i `KOORDINERING.md`.
- **Sky-økt** (kode/dokumentasjon): claude.ai/code → repoet `timmytimss/excelVBAkoding`. Claude-appen må være installert på `timmytimss` på GitHub.
- Minnefilene fra de gamle øktene (bakgrunn om Håkon og arbeidsmåten) ligger privat i `timmytimss/timssboka`:
  `kilder/2026-10-08-claude-minne/` og `arkiv/excel-vba-koding/minne-backup/`.

## 4. Verktøy for lokal utvikling
- ps2exe (for å kompilere EMI): `Install-Module ps2exe -Scope CurrentUser` (se kunnskapsbasen, `#mac-emi-bygg`).
- Excel: «Klarer tilgang til objektmodellen for VBA-prosjektet» må være på.

## 5. Gammel historikk
Full git-historikk fra før repoet ble offentlig (2026-10-08): `timssboka/arkiv/excel-vba-koding/git-historikk-for-offentlig-2026-10-08.bundle`
(`git clone <bundle> gammel-historikk`).
