# Excel VBA Koding – Excel Macro Installer (EMI)

En verktøykasse av Excel-makroer (VBA) og et lite Windows-program, **Excel Macro Installer (EMI)**, som installerer, oppdaterer og fjerner dem i dine egne arbeidsbøker – uten at du trenger å åpne VBA-editoren.

## Installere EMI

1. Last ned repoet som zip: **[excelVBAkoding-main.zip](https://github.com/timmytimss/excelVBAkoding/archive/refs/heads/main.zip)** (eller den grønne **Code**-knappen → **Download ZIP**).
2. Høyreklikk zip-fila → **Egenskaper** → huk av **Fjern blokkering** (*Unblock*) → **OK**. Ellers kan Windows blokkere makroene som er lastet ned fra internett.
3. Pakk ut, og kopier **bare mappa `Excel Macro Installer`** til et fast sted, f.eks. `C:\Users\<deg>\AppData\Local\Programs\Excel Macro Installer` (helst *ikke* i en OneDrive-synkronisert mappe). Slett mappa `Installer` inni den – den er kildekode og trengs ikke (og uten den kan EMI oppdatere seg selv).
4. Start `Excel Macro Installer.exe`. Første gang kan Windows SmartScreen si «Windows beskyttet PC-en» – klikk **Mer informasjon** → **Kjør likevel** (programmet er ikke kodesignert).
5. Lag gjerne en snarvei på skrivebordet.

**Forutsetninger:** Windows 10/11 med Excel (skrivebordsversjonen). I Excel: *Fil → Alternativer → Klareringssenter → Innstillinger for klareringssenter → Makroinnstillinger* → huk av **Klarer tilgang til objektmodellen for VBA-prosjektet**.

## Oppdateringer

EMI sjekker ved oppstart om `Excel Macro Installer/versjon.json` her er nyere enn installasjonen. Da vises **Oppdater EMI** nederst i vinduet. Oppdateringen erstatter kun programfilene (exe, `Macros`, `Resources`, `Data/macro-metadata.json`); listen over dine filer og installerte makroer beholdes. Bruk deretter **Oppdater** på hver fil for å få nye makroversjoner inn i arbeidsbøkene.

## Innhold i repoet

| Mappe/fil | Hva |
|---|---|
| `Excel Macro Installer/` | Programmet (`.exe`), makroinstallatørene (`Macros/`), ikoner og WebView2 (`Resources/`), kildekoden (`Installer/`) |
| `Excel VBA Koding.html` | Kunnskapsbasen: arkitektur, hver makro, fallgruver og testmetoder |
| `Prosjektfiler/` | Byggefragmenter for kunnskapsbasen, testverktøy, testfiler (oppdiktede data), makroer under utvikling, publiseringsskript |
| `CLAUDE.md`, `KOORDINERING.md` | Arbeidsinstruksjoner og koordinering for Claude Code-øktene som utvikler prosjektet |
