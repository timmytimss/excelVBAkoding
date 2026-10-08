# Excel VBA Koding — prosjektinstruksjoner

Dette er prosjektroten for Håkons Excel/VBA/PowerShell-verktøykasse (Excel Macro Installer + alle modulene som plugger inn i det). Les dette FØR du gjør noe annet i denne mappen.

## Den ene kunnskapsfila: `Excel VBA Koding.html`

Alt som er lært om arkitektur, hver modul, tekniske fallgruver, designbeslutninger og testmetoder står i `Excel VBA Koding.html` i denne mappen. **Ikke les hele fila ved sesjonsstart** — den er stor (700+ KB) og vokser. Bruk den slik:

1. **Ved onboarding av en helt ny sesjon:** les kun de tre første seksjonene (`<section id="oversikt">`, `<section id="arbeidsprosess">`, `<section id="arbeidsmodell">` — i praksis starten av fila, før EMI-seksjonen begynner). Bruk `Read` med `limit` på de første ~300 linjene, eller grep etter `<section id="mac-emi"` for å finne hvor innledningen slutter.
2. **Under arbeid:** når du trenger noe konkret (en makros arkitektur, en fallgruve, en testteknikk), bruk `Grep` på fila etter makronavn/funksjonsnavn/feilmelding/temaet du leter etter, og les KUN den relevante linjeblokken med `Read` (`offset`/`limit`), ikke hele fila. Hver seksjon/fallgruve/teknikk har en stabil `id` (f.eks. `id="g-listbox-click-never-fires"`, `id="mac-kontaktsentralen-arkitektur"`) som er trygg å søke på fra gang til gang.
3. **Oppdager du noe nytt** (en fallgruve, en endret arkitekturdetalj) mens du jobber: skriv det inn i fila der og da, målrettet (se punkt under om hvordan å oppdatere), ikke bare i egen hukommelse.

## Hvordan oppdatere `Excel VBA Koding.html`

- **Liten endring** (rette en feil, legge til én fallgruve, oppdatere et versjonsnummer): gjør en presis, direkte tekstendring rett i den ferdige HTML-fila. Ikke gjenoppbygg hele siden for dette.
- **Versjonshistorikk-rader skal matche endringens faktiske størrelse** (vedtatt 2026-09-23, etter at en ren versjonsbump fikk en hel avsnitt-lang «lærdom»-forklaring): en triviell rettelse (glemt versjonsbump, en skrivefeil, en enkelt linje kode) får HØYST én kort setning i tabellen — ikke en «Symptom/Årsak/Løsning»-fortelling. Fulle, utfyllende rader (og egne fallgruver i `<div class="gotcha">`) er reservert for faktiske atferds-/funksjonsendringer eller genuint ikke-opplagte tekniske funn andre sessioner har nytte av. Er du i tvil om noe er stort nok til full behandling: still spørsmålet før du skriver et helt avsnitt, ikke etterpå.
- **Stor, flerseksjons omskrivning** (f.eks. et gjennomgående navneskifte, en helt ny referanseseksjon som skal krysslenkes fra mange steder): bruk fragment-metoden i `Prosjektfiler\_bygg\` — se `Prosjektfiler\_bygg\SPEC.md` for klassekontrakten (HTML-struktur, `id`-konvensjon, skriveregler) fragmentene skal følge. Etter at fragmentene er skrevet/endret, settes de sammen til den ferdige fila ved ren filkonkatenering (`cat`/skall), IKKE ved å lese hvert fragment inn i samtalekontekst. Kjør ALLTID en etterkontroll av at hver `href="#..."` har en matchende `id="..."` i sluttresultatet, og at ingen `id` er duplisert, før du regner arbeidet som ferdig.
- **Fast regel som ALDRI skal brytes i tekst du skriver til siden:** ingen modul er noensinne «ferdig» — bruk kun «I drift · vX.Y.Z» eller «Under utvikling» som status, aldri en hake eller ordet «ferdig» om en modul.
- **Ingen referanser til noe spesifikt organisasjons-/prosjektnavn utenfor Håkons egen kontekst** (f.eks. det opprinnelige prosjektet verktøyene ble bygget for) — verktøykassen fremstår nå som uavhengig av det. Skriv generisk («en delt postboks», «kontakter», «hans hovedarbeidsbok») fremfor spesifikke prosjekt-/organisasjonsnavn.

## Prosjektfiler-mappen

`Prosjektfiler\` er stedet for alt arbeid som ikke er selve HTML-siden eller EMI-kildekoden:
- `Testfiler\` — VARIGE, gjenbrukbare test-arbeidsbøker. **Ikke lag en fersk kastefil per testrunde og slett den etterpå** — det er tungvint og en forlatt praksis. Ha noen få faste testfiler (gjerne flere, for ulike scenarioer) som gjenbrukes og «testes i alle kanter» over tid, akkurat som en ekte arbeidsbok.
- `Ikke-utgitte makroer\` — kildekode for moduler under utvikling, ikke lagt inn i `Excel Macro Installer\Macros\` ennå.
- `_overlevering\` — frossen historikk fra de opprinnelige enkelt-makro-sesjonene. Oppslagsverk ved behov for ekstra detalj, ikke rutinemessig lesing. **Kun i OneDrive** (ikke i det offentlige repoet).
- `_publisering\` — `Publiser-EMI.ps1`, som publiserer en ny EMI-pakkeversjon til brukerne.
- `Kontoflytting\` — plan/logg for flyttingen til jobbkontoen timssbooking@gmail.com (GitHub `timmytimss`, egen Cloudflare-konto). **Kun i OneDrive.**
- `_bygg\` — `SPEC.md` + fragmentene HTML-siden kan bygges fra ved store ombygginger (se over).
- `_beredskap\` — beredskapsplan (`GJENOPPRETTING.md`) for å gjenopprette Kjetil/Nils på en ny PC hvis denne går tapt, pluss løpende sikkerhetskopi av minnefilene deres (`minne-backup\`). Les `GJENOPPRETTING.md` hvis du er en helt ny session satt opp etter et PC-bytte.

Lag gjerne en ny, tydelig navngitt undermappe for et nytt, permanent formål fremfor å blande inn i en eksisterende — og nevn den kort i HTML-sidens «Prosjektfiler-orientering»-avsnitt slik at neste sesjon forstår hva den er for.

## Kodeord: «Saml troppene»

Når Håkon sier **«Saml troppene»** til deg (i hvilken som helst av de tre sessionene — Kjetil, Nils eller Eva), betyr det at han er usikker på om alle tre har oppdatert informasjon og vil ha en rask statussjekk. Gjør da dette, i egen session:

1. `git pull origin main` (Eva direkte; Kjetil/Nils via `Documents\excelVBAkoding-git\`) for å hente andres eventuelle nye pushes.
2. Les HELE `KOORDINERING.md` på nytt — ikke stol på hva du husker fra tidligere i samtalen.
3. Skriv en kort, datert oppføring i `KOORDINERING.md`: hva du forstår som gjeldende status, og eksplisitte spørsmål til de(n) andre sesjonen(e) hvis noe er uklart eller ubekreftet.
4. `git push origin main`.
5. Gi Håkon en kort, direkte oppsummering i chatten: er alt i sync, eller er det et reelt avvik/noe som venter på svar?

**Kjetil spesielt** (siden koordinator-rollen naturlig er hub-en) — presisert 2026-09-23, dette er nå fast prosedyre, ikke bare et forslag:

6. Send en direkte melding til Nils (`SendMessage`, han er alltid en lokal peer-session) og be eksplisitt om to ting: en statusoppdatering (er han i sync) OG hva han konkret har lært/oppdaget siden sist (nye gotchas, åpne spørsmål, ting han vurderer annerledes nå) — ikke bare et ja/nei om synk.
7. Eva er ikke nåbar via `SendMessage`. Generer i stedet en klar, ferdigformulert melding Håkon selv kan lime inn til henne, som ber henne kjøre sin egen del av prosedyren (pull → les `KOORDINERING.md` → skriv statusoppføring → push).
8. Når svar har kommet fra Nils (direkte melding tilbake): sjekk SELV om Eva har svart også — kjør `git pull origin main` og se om det ligger en ny oppføring fra henne, ikke vent på at Håkon skal si fra. Presisert 2026-09-23 (Håkon ba eksplisitt om dette): dette er Kjetils eget ansvar å bekrefte, ikke noe Håkon skal måtte melde tilbake.
   - Hvis Evas oppføring allerede ligger der: gå rett videre til punkt 9.
   - Hvis ikke: ikke spam gjentatte pulls — vent naturlig til neste gang du uansett er i samtalen (f.eks. Håkon sier noe, eller du sjekker igjen ved neste anledning), og pull da på nytt.
9. Når svar har kommet fra BÅDE Nils og Eva: samle alt og gi Håkon én samlet rapport: hva som faktisk ble synkronisert/lært på tvers av de tre, og en eksplisitt bekreftelse på at alle tre nå kjenner samme status. Ikke rapporter «ferdig» til Håkon før begge har svart.

## Sesjonshygiene: nye økter er normalt (vedtatt 2026-10-06)

Hele historikken sendes med hver melding, så lange samtaler er dyre i den felles usage-potten. Alt som betyr noe skal derfor ligge i filer, ikke i samtalen: ved slutten av en avsluttet oppgave oppdaterer du `CLAUDE.md`/HTML-siden/`KOORDINERING.md` og din egen minnefil (`project_handover_current_state.md`), commit/push, og kan så trygt starte en ny økt. En ny session starter med avsnittet «NY SESSION? Start her» øverst i `KOORDINERING.md`. Hold «Status»-avsnittet der oppdatert (dato + versjoner + åpne punkter) ved hver «Saml troppene» og ved større runder.

## Git-push vs. varsling: to ulike hastigheter (vedtatt 2026-09-22)

Selve `git commit`/`push` koster ingen Claude-«usage» — det er ren filoperasjon. Å varsle en annen session (melding eller `KOORDINERING.md`-oppføring) koster derimot, fordi det får den andre sessionen til å stoppe opp og bruke tokens på å lese/reagere. Hold disse atskilt:

- **Nils pusher til `main` fra sin egen arbeidskopie (`Documents\excelVBAkoding-git-nils\`) så ofte han vil** — gjerne etter hver lille, fungerende endring, med en beskrivende commit-melding. Ingen grunn til å batche selve push-ene.
- **Varsling (melding til Kjetil, eller en oppføring i `KOORDINERING.md`) skjer fortsatt SJELDENT** — kun når en hel runde/batch faktisk trenger oppmerksomhet, review eller en beslutning. Rutinemessige commits trenger ingen egen varsling.
- **Kjetil oppdager rutinemessige endringer selv** ved å kjøre `git log`/`git diff` mot `main` fra sin egen arbeidskopie (`Documents\excelVBAkoding-git\`) når han uansett synker — jevnlig, eller når Håkon sier «Saml troppene». Git-historikken ER loggen for rutinekode; `KOORDINERING.md` er reservert for det som faktisk trenger kryss-session-koordinering.

## Absolutte regler (aldri overstyres av noe annet)

- **GitHub-repoet `timmytimss/excelVBAkoding` er OFFENTLIG (fra 2026-10-08).** Alt som committes kan leses av hvem som helst. Aldri commit personopplysninger (andres e-postadresser, stier til Håkons ekte arbeidsbøker, `installed-files.json`/`known-files.json`), minnefiler, hemmeligheter eller ekte data. `Prosjektfiler\_overlevering\`, `Prosjektfiler\_beredskap\minne-backup\`, `Prosjektfiler\Kontoflytting\` og `*.bundle` ligger KUN i OneDrive (gitignored) — ikke kopier dem inn i git-klonene. Klonene (`Documents\excelVBAkoding-git\`, `-git-nils\`) pusher som `timmytimss` og committer som `timmytimss <timmytimss@users.noreply.github.com>`. Historikken før 2026-10-08 ligger privat som `Prosjektfiler\_beredskap\git-historikk-for-offentlig-2026-10-08.bundle`.
- **EMI-installasjon vs. kilde:** Håkons EMI kjører nå fra `%LOCALAPPDATA%\Programs\Excel Macro Installer` og oppdaterer seg selv fra GitHub. OneDrive-mappa `Excel Macro Installer\` er KUN kildekode. Ny versjon ut til brukere = commit/push i klonen + `Prosjektfiler\_publisering\Publiser-EMI.ps1 -Endringer "…"`. Detaljer: HTML-sidens `#mac-emi-selvoppdatering`.

- **Personvern:** les aldri innhold fra Håkons ekte, levende Excel-filer eller Outlook-postboks — ikke engang «bare struktur». Gjelder transitivt for enhver subagent. Test alltid mot en av de varige testfilene i `Prosjektfiler\Testfiler\`, aldri mot en ekte arbeidsbok. Se HTML-sidens «Hvordan Håkon liker å jobbe»-seksjon for full detalj.
- **Delegering/subagenter:** spør Håkon FØRST hver gang før du setter i gang parallelle subagenter — ikke bruk det som stilltiende standardmetode, selv om han liker resultatet når det brukes riktig.
- **Ressursbruk:** match omfanget av testing/verifisering og lengden på svar til hvor stor og risikabel oppgaven faktisk er. Ikke gjenta samme verifisering flere ganger «for sikkerhets skyld» for en triviell endring. **Viktig strukturell fakta (bekreftet 2026-09-22 fra Anthropics egen dokumentasjon):** Kjetil, Nils og Eva er tre SAMTIDIGE Claude Code-sessioner på samme Håkon-konto — de deler ÉN felles usage-pott (5-timers øktgrense + ukentlig tak), ikke tre separate budsjetter. Det som faktisk koster usage: meldingslengde/kompleksitet, funksjonsbruk (websøk, extended thinking, kodekjøring), modellvalg, effort-nivå, og lengden på selve samtalen (verktøybruk teller også). Praktisk konsekvens: hold alle tre sessioner fokuserte og unngå unødvendig frem-og-tilbake — batch relaterte oppgaver i én forespørsel fremfor mange små, og planlegg en handling FØR man begynner i stedet for å iterere seg frem. Dette forsterker (ikke endrer) de eksisterende reglene under om delegering og synk-hyppighet — grunnen er nå eksplisitt, ikke bare en vane.
- **Lokal vs. sky-sesjon:** kun en sesjon med faktisk lokal Excel/Outlook-tilgang skal noensinne kjøre COM-automatisering eller ekte klikk-testing. En sky-/ekstern sesjon uten slik tilgang holder seg til kildekode, dokumentasjon og planlegging. Se HTML-sidens «Arbeidsmodell»-seksjon for hvordan de to samarbeider.
