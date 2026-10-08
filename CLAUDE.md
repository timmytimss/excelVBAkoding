# Excel VBA Koding — prosjektinstruksjoner

Dette er Håkons Excel/VBA/PowerShell-verktøykasse: Excel Macro Installer (EMI) + alle makroene som plugger inn i den. Les dette FØR du gjør noe annet.

## Hvor prosjektet bor (fra 2026-10-08)

- **GitHub er den eneste fasiten:** `timmytimss/excelVBAkoding` (OFFENTLIG), gren `main`. Det finnes ingen OneDrive-kopi lenger.
- **Sky-økter** (claude.ai/code) jobber direkte i repoet: kildekode, dokumentasjon, planlegging.
- **Lokale økter** (på Håkons PC, nødvendig for testing i Excel/EMI) jobber i `C:\Users\haaklun\Documents\excelVBAkoding` (en vanlig git-klone). Start ALLTID med `git pull origin main`, og commit/push når en fungerende endring er klar — da ser sky-øktene den også.
- Push som `timmytimss` (remote `https://timmytimss@github.com/timmytimss/excelVBAkoding.git`), commit som `timmytimss <timmytimss@users.noreply.github.com>`.
- **Håkons installerte EMI:** `C:\Users\haaklun\Documents\Excel Macro Installer`. Den oppdaterer seg selv fra repoet. Utviklingskopier (git-klonen) oppdateres aldri automatisk.
- **Privat materiale** som ikke skal være offentlig (gamle overleveringsnotater, minne-backup, kontoflyttingsplan, gammel git-historikk) ligger i det private repoet `timmytimss/timssboka` under `arkiv/excel-vba-koding/`.
- Kontoer og infrastruktur (GitHub, Cloudflare, Google) er beskrevet i `timssboka/kilder/2026-10-08-kontoer-infrastruktur-og-lardommer.md`.

## Den ene kunnskapsfila: `Excel VBA Koding.html`

Alt som er lært om arkitektur, hver modul, tekniske fallgruver, designbeslutninger og testmetoder står i `Excel VBA Koding.html`. **Ikke les hele fila** — den er stor (800+ KB). Bruk den slik:

1. **Ny økt:** les kun innledningen (`<section id="oversikt">`, `<section id="arbeidsprosess">`, `<section id="arbeidsmodell">` — starten av fila, før `<section … id="mac-emi"`).
2. **Under arbeid:** `Grep` etter makronavn/funksjon/feilmelding/tema, og les KUN den relevante blokken (`offset`/`limit`). Seksjoner og fallgruver har stabile `id`-er (f.eks. `g-listbox-click-never-fires`, `mac-emi-selvoppdatering`).
3. **Nye funn** (fallgruve, endret arkitektur): skriv dem inn i fila med en gang, målrettet.

NB: noen eldre avsnitt i HTML-fila beskriver fortsatt OneDrive-stier og rollene Kjetil/Nils/Eva fra før 2026-10-08 — dette dokumentet går foran ved motstrid.

## Hvordan oppdatere `Excel VBA Koding.html`

- **Liten endring:** presis, direkte tekstendring i den ferdige fila. Ikke gjenoppbygg siden.
- **Versjonshistorikk-rader skal matche endringens størrelse:** en triviell rettelse får HØYST én kort setning. Fulle rader og egne fallgruver (`<div class="gotcha">`) er for faktiske atferdsendringer eller genuint ikke-opplagte funn. I tvil: spør før du skriver et helt avsnitt.
- **Stor, flerseksjons omskrivning:** bruk fragment-metoden i `Prosjektfiler/_bygg/` (se `SPEC.md`), sett sammen ved ren filkonkatenering, og kontroller etterpå at hver `href="#…"` har en matchende `id` og at ingen `id` er duplisert.
- **Ingen modul er noensinne «ferdig»:** bruk kun «I drift · vX.Y.Z» eller «Under utvikling».
- **Skriv generisk** («en delt postboks», «hans hovedarbeidsbok») fremfor spesifikke prosjekt-/organisasjonsnavn.

## Mappene i repoet

- `Excel Macro Installer/` — programmet: `Installer/Excel Macro Installer.ps1` (kilde, kompileres med ps2exe til `.exe`), `Macros/*.ps1` (én installatør per makro), `Resources/`, `Data/macro-metadata.json`, `versjon.json` (publisert pakkeversjon).
- `Prosjektfiler/Testfiler/` — VARIGE, gjenbrukbare test-arbeidsbøker (oppdiktede data). Ikke lag kastefiler per testrunde.
- `Prosjektfiler/Ikke-utgitte makroer/` — makroer under utvikling.
- `Prosjektfiler/_testverktoy/` — testskript for lokal testing (se `README.md` der).
- `Prosjektfiler/_publisering/Publiser-EMI.ps1` — publiserer ny EMI-versjon til brukerne.
- `Prosjektfiler/_bygg/` — `SPEC.md` + fragmentene HTML-siden bygges fra ved store ombygginger.
- `Prosjektfiler/_beredskap/GJENOPPRETTING.md` — hvordan sette opp på nytt etter PC-bytte.
- `KOORDINERING.md` — logg over større runder og beslutninger mellom økter (start med «NY SESSION? Start her»).

## Publisere en endring til brukerne

1. Endre kilden, test lokalt (EMI-kilden må rekompileres med ps2exe og `$LauncherVersion` + `-version` bumpes hvis `Installer/Excel Macro Installer.ps1` endres; makroer bumper sin egen versjonsvariabel).
2. Commit og push til `main`.
3. Kjør `Prosjektfiler/_publisering/Publiser-EMI.ps1 -Repo C:\Users\haaklun\Documents\excelVBAkoding -Endringer "…"` (øker `versjon.json` og pusher). EMI hos brukerne tilbyr da «Oppdater EMI» ved neste oppstart.

## Sesjonshygiene

Lange samtaler er dyre (hele historikken sendes med hver melding, og alle økter deler én usage-pott). Alt som betyr noe skal ligge i filer: oppdater `CLAUDE.md`/HTML-siden/`KOORDINERING.md` ved slutten av en oppgave, commit/push, og start gjerne en ny økt. Hold «Status»-avsnittet i `KOORDINERING.md` oppdatert ved større runder. Varsle andre økter sjelden (kun når noe trenger beslutning); vanlige commits trenger ingen varsling — git-historikken ER loggen.

## Absolutte regler (aldri overstyres av noe annet)

- **Repoet er OFFENTLIG.** Aldri commit personopplysninger (andres e-postadresser, stier til Håkons ekte arbeidsbøker, `installed-files.json`/`known-files.json`), minnefiler, hemmeligheter, tokens eller ekte data. Privat materiale hører hjemme i `timssboka/arkiv/` (privat repo).
- **Personvern:** les aldri innhold fra Håkons ekte, levende Excel-filer eller Outlook-postboks — ikke engang «bare struktur». Gjelder også subagenter. Test alltid mot testfilene i `Prosjektfiler/Testfiler/`.
- **Lokal vs. sky:** kun en lokal økt med faktisk Excel/Outlook-tilgang skal kjøre COM-automatisering eller ekte klikk-testing. Sky-økter holder seg til kode, dokumentasjon og planlegging.
- **Delegering/subagenter:** spør Håkon FØR du setter i gang parallelle subagenter.
- **Ressursbruk:** tilpass testing og svarlengde til hvor stor og risikabel oppgaven er. Ikke gjenta verifisering «for sikkerhets skyld». Planlegg før du handler, og samle relaterte oppgaver.
