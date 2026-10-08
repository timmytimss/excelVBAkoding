# SPEC: HTML-fragmenter til «Excel VBA Koding.html»

## Bakgrunn (les først)
Håkon (UiO, product owner/tester, IKKE utvikler, norsk bokmål) bygger Excel/VBA-makroer (TIMSS-prosjektet er bare stedet han jobber). Alle Claude-«Macro Maker»-sessions legges ned og kildemappene slettes. All lærdom må derfor samles i EN selvforklarende HTML-fil, «Excel VBA Koding.html», som en helt ny AI-chat skal kunne onboardes med ved å lese bare den. Filen handler generelt om **Excel VBA-koding**; TIMSS er bare eksempelbruk.

Alle makroer er moduler i ett pågående hovedprosjekt, **Excel Macro Installer (EMI)**. **Aldri** formuler noe som «ferdig»/«done»/✅. Bruk «I drift», «Under utvikling», «Planlagt».

## Din jobb
Du får en liste kildefiler. Les dem grundig og skriv HTML-**fragmenter** (ikke hele sider) til `...\Prosjektfiler\_bygg\frag\`. Ikke slå sammen, ikke lag <html>/<head>/<body>/<style>/<script>; det gjør jeg.

## Regler
- Språk: norsk bokmål. Presis, konkret, teknisk. Skriv for en AI-leser som aldri har sett prosjektet.
- **Ikke tap detaljer.** Ta med alle navn (komponenter, funksjoner, konfignøkler, filer, stier), versjonsnummer, tall, feilkoder, symptomer, årsaker, løsninger, og HVORFOR bak beslutninger. Komprimer formuleringer, ikke innhold. Fjern bare ren gjentakelse og ren prosess-støy («Håkon sa ok»).
- Ingen ekte persondata fra Håkons filer finnes i kildene; ikke gjett eller finn på slikt heller.
- Les ALDRI Håkons ekte, levende Excel-filer eller postboks (f.eks. Hovedfila-2027.xlsm). Kun kildefilene du får oppgitt.
- Escape HTML i kode/tekst: `&lt;` `&gt;` `&amp;`. Kode i `<pre><code>…</code></pre>`, korte identifikatorer i `<code>…</code>`.
- Norske tegn (æøå) rett i teksten, filene skrives som UTF-8 (Write-verktøyet gjør det; bruk aldri PowerShell `Set-Content` uten `-Encoding utf8`).
- Ingen emoji-overbruk. Ingen markdown; ren HTML.
- Skriv fragmentene i flere Write-kall om nødvendig (ett fragment kan deles i `_a`, `_b`, `_c` filer som jeg konkatenerer i navnerekkefølge). Hold hver fil under ca. 25 KB.

## Klassekontrakt (bruk nøyaktig disse)
Makro-fragment (`frag\mac_<slug>_a.html` osv.):
```html
<section class="macro" id="mac-<slug>" data-title="Kontaktsentralen" data-status="I drift · v0.10.5">
  <h2>Kontaktsentralen</h2>
  <p class="lead">Én setning: hva den gjør.</p>
  <dl class="facts"><dt>Versjon</dt><dd>…</dd><dt>Kilde</dt><dd><code>sti</code></dd><dt>VBA-komponenter</dt><dd>…</dd><dt>EMI-ikon/farge</dt><dd>…</dd></dl>
  <h3 id="mac-<slug>-bruk">Formål og bruk</h3> …
  <h3 id="mac-<slug>-arkitektur">Arkitektur</h3> …
  <h3 id="mac-<slug>-historikk">Funksjoner og versjonshistorikk</h3> <table class="hist"><thead><tr><th>Versjon</th><th>Hva</th></tr></thead><tbody>…</tbody></table>
  <h3 id="mac-<slug>-design">Designbeslutninger og hvorfor</h3> …(inkl. forkastede retninger)
  <h3 id="mac-<slug>-test">Testing av denne makroen</h3> …
  <h3 id="mac-<slug>-apent">Åpne punkter, begrensninger og idéer</h3> …
</section>
```
Bruk `<div class="callout warn">…</div>` for viktige advarsler og `<div class="callout note">…</div>` for tips. Tabeller: `<table>` med `<thead>`.

Fallgruve-fragment (`frag\gotchas_<slug>.html`): kun rene oppføringer, ingen wrapper-section:
```html
<div class="gotcha" id="g-<kort-slug>" data-cat="<én av: vba, excel, com, powershell, msforms, outlook, encoding, installer, testing, api, process>">
  <h4>Kort tittel</h4>
  <dl><dt>Symptom</dt><dd>…</dd><dt>Årsak</dt><dd>…</dd><dt>Løsning</dt><dd>…</dd></dl>
  <p class="src">Kilde: <navn på makro/session></p>
</div>
```
Én oppføring per distinkt fallgruve (ikke slå ulike sammen). Generelle VBA/COM/PowerShell/MSForms/Office-lærdommer skal med i gotchas-filen (ikke bare i makro-fragmentet), og makro-fragmentet kan henvise dem med `<a href="#g-slug">…</a>`.

Testmetode-fragment (`frag\testing_<slug>.html`) kun hvis kildene har generelle testteknikker (kompileringssjekk, modal-skjema-testing, sikker Excel-instans osv.): `<div class="technique" id="t-<slug>"><h4>…</h4><p>…</p></div>`.

Preferanser/arbeidsstil om Håkon: `frag\hakon_<slug>.html` som en `<ul>` med korte punkter; jeg dedupliserer.

## Ferdigmelding
Svar med en kort liste: hvilke filer du skrev (filnavn + omtrent KB), antall gotcha-oppføringer, og eventuelle ting i kildene du var usikker på eller som virker motstridende. Ikke lim inn innholdet.
