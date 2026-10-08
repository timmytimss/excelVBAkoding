# Beredskapsplan: gjenopprett hele flyten etter et tap

**Denne fila skal fungere selv om den eneste overlevende kopien er git-klonen på GitHub (`timmytimss/excelVBAkoding`).** Skriv den derfor alltid slik at den ikke forutsetter at OneDrive-mappa finnes — det er nettopp poenget med Krisescenario 2 under. Hvis du er Eva og leser dette fordi Håkon har orientert deg om at noe har gått tapt: dette er din komplette instruks, du trenger ikke noe annet.

**Forutsetning som IKKE er dekket her:** at Håkon fortsatt har tilgang til GitHub-kontoen som eier repoet (`timmytimss`, innlogging timssbooking@gmail.com; Håkons private `actualpope` er bare samarbeidspartner siden 2026-10-08) og til Claude. Uten det hjelper ingenting av dette. Verdt å nevne til Håkon en gang: sørg for at du har en pålitelig gjenopprettingsvei inn i GitHub-kontoen din (recovery-e-post, 2FA-backupkoder) — det er den ene lenken i kjeden denne planen ikke kan sikre for deg.

## To krisescenarioer

- **Krisescenario 1 — UiO-PC-en er ødelagt, men OneDrive-kontoen/innholdet er intakt.** Enklest: alt kommer tilbake automatisk så snart OneDrive er satt opp igjen på en ny maskin. Se «Scenario 1» under.
- **Krisescenario 2 — selve OneDrive-innholdet er borte** (slettet, mistet tilgang, eller Håkon bevisst bytter til en annen konto/lokal base). Da finnes INGENTING igjen utenom det som er i git-repoet. Se «Scenario 2» under — dette er den grundigere gjenoppbyggingen.

Begge scenarioene kan kombineres (PC-en ØG OneDrive-innholdet borte samtidig) — da følger du Scenario 2 i sin helhet, siden den ikke forutsetter noe fra den gamle PC-en eller OneDrive-kontoen.

## Hva som faktisk ligger i git-repoet (bekreft dette stemmer når du leser)

Repoet er en FULL, rå kopi av hele `Excel VBA Koding\`-mappa, ikke bare kildekode:
- `Excel VBA Koding.html` — hele kunnskapsbasen (arkitektur, alle makroer, fallgruver, testteknikker, arbeidsmodell).
- `CLAUDE.md` — prosjektinstruksjonene, lastes automatisk av Claude Code når arbeidskatalogen er denne mappa (eller en klone av den).
- `KOORDINERING.md` — hele koordineringsloggen mellom Kjetil/Nils/Eva, med full historikk.
- `Excel Macro Installer\` — all EMI-kildekode og kompilert `.exe`.
- `Prosjektfiler\` — testfiler, ikke-utgitte makroer, historisk arkiv (`_overlevering\`), byggefragmenter (`_bygg\`), og denne `_beredskap\`-mappa med minnebackup.

Det ENESTE som IKKE ligger i git, og som du (Eva) derfor må hjelpe Håkon huske finnes andre steder: ingenting av betydning, egentlig — se «Det som ikke kan gjenopprettes» nederst for de få unntakene.

## Scenario 1: PC-en er ødelagt, OneDrive er intakt

1. Sett opp OneDrive på den nye maskinen med samme konto. Vent til `Excel VBA Koding\`-mappa (inkl. denne `_beredskap\`-mappa) er ferdig synket ned — dette skjer automatisk, ingenting Claude-spesifikt kreves.
2. Opprett to nye Claude Code-sessions i «Jobbrelatert»-gruppen, begge med arbeidskatalog satt til `Excel VBA Koding\Prosjektfiler`.
3. Gi dem nøyaktig samme navn som før: `EVK (OD): Koordinator "Kjetil"` og `EVK (OD): Koder "Nils"` (`set_session_title`).
4. Kopier minnefilene manuelt fra `Prosjektfiler\_beredskap\minne-backup\` inn i hver ny sessions faktiske minnemappe (`C:\Users\haaklun\.claude\projects\<ny-sti-hash>\memory\` — spør sessionen selv om nøyaktig sti). **Dette skjer ikke automatisk** — kjent, dokumentert fallgruve.
5. Klon git-arbeidskopien på nytt: `git clone https://github.com/timmytimss/excelVBAkoding.git` inn i `Documents\excelVBAkoding-git\` (utenfor OneDrive, som før — se hvorfor i Arbeidsmodell-seksjonen i HTML-siden). Sett lokal `git config user.name`/`user.email` på nytt.
6. Be de nye sessionene lese `CLAUDE.md` (auto), så hele `KOORDINERING.md`, og evt. si «Saml troppene» for en fersk statussjekk mot Evas nyeste git-historie.

## Scenario 2: OneDrive-innholdet er borte — gjenoppbygg fra git alene

Ingen OneDrive-mappe å vente på. Håkon velger en NY base — kan være en ren lokal mappe uten skysynk, eller en annen (f.eks. personlig) OneDrive-konto. Uansett hvilken:

1. **Klon repoet rett inn i den nye basen:** `git clone https://github.com/timmytimss/excelVBAkoding.git "<ny-base>\Excel VBA Koding"`. Dette gir deg HELE strukturen tilbake — HTML-siden, CLAUDE.md, KOORDINERING.md, EMI, Prosjektfiler, denne beredskapsmappa. `.git`-mappa havner nå INNI den nye basen, som er annerledes enn dagens oppsett (der git-kopien bevisst ligger utenfor OneDrive) — det er greit for en ren lokal base, men **hvis den nye basen er skysynket** (en annen OneDrive-konto), flytt `.git`-mappa ut av den synkede mappa etterpå, av samme grunn som i dag (se `id="arbeidsmodell-git"` i HTML-siden — synk + git-interne filer kan låse hverandre).
2. **Gjenopprett minnet:** kopier `.md`-filene fra `Prosjektfiler\_beredskap\minne-backup\` (nå en del av den ferske klonen) inn i minnemappa til to nye Claude Code-sessions, akkurat som i Scenario 1, punkt 4.
3. **Opprett to nye lokale sessions** (`Jobbrelatert`-gruppen), arbeidskatalog satt til `<ny-base>\Excel VBA Koding\Prosjektfiler`, navngitt akkurat som før: `EVK (OD): Koordinator "Kjetil"` og `EVK (OD): Koder "Nils"`. (Navnetagget «OD» i navnene refererer historisk til OneDrive — behold navnet uansett for gjenkjennelighet, selv om basen nå er en annen; det er kun et kallenavn-tag, ikke en teknisk avhengighet.)
4. **Fortell Eva (hvis hun ikke allerede vet det) at basen har flyttet.** Hun trenger ikke gjøre noe teknisk selv — hun har allerede full tilgang via git — men bør vite det for å unngå å referere til den gamle OneDrive-stien i noe hun skriver videre.
5. **Verifiser at EMI faktisk fungerer fra den nye basen** før du stoler på den — samme lette test som ble gjort 2026-09-22 (`Kontaktsentralen.ps1 -NyExcelInstans` mot en testfil i `Prosjektfiler\Testfiler\`, isolert Excel-instans). Se `Excel VBA Koding.html` → Migrering-seksjonen for hvordan det ble gjort forrige gang.
6. Les `CLAUDE.md` og hele `KOORDINERING.md`, og si gjerne «Saml troppene» for en fersk statussjekk.

## Det som ikke kan gjenopprettes (vær ærlig om dette, ikke bare fortell Håkon at alt er trygt)

- **Ferske endringer siden siste push.** Hvor mye som går tapt avhenger av hvor lenge det er siden forrige synk-runde (se «Synk-hyppighet»-regelen i `KOORDINERING.md` — sjelden, batchet per fullført runde). En kodeendring Nils akkurat har gjort, men ikke rapportert/synket ennå, er IKKE i git og er derfor tapt i Scenario 2 hvis den bare fantes i OneDrive-mappa. Minimer denne risikoen ved å faktisk følge synk-rutinen, ikke la runder ligge uforløst for lenge.
- **Selve samtalehistorikken** til de gamle Kjetil/Nils-sessionene (den løpende dialogen med Håkon) — minnefilene fanger opp LÆRDOMMEN og BESLUTNINGENE, ikke ordrett hva som ble sagt. Det er en bevisst avveining (kompresjon > råtekst), ikke noe denne planen prøver å fikse.
- **Ekte, levende Excel/Outlook-data** — var aldri i git eller OneDrive-backupen i utgangspunktet (personvernregelen), så det er ikke noe "tap" knyttet til denne planen — bare en påminnelse om at Håkons faktiske arbeidsfiler (f.eks. hovedarbeidsboken) er et helt separat spørsmål, utenfor denne beredskapsplanens omfang.
