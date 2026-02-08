# MatAPP — kratki vodič

Kratko: ovaj projekat sadrži CLI alate za:

- parsiranje XML `.db` fajla u JSON,
- generisanje XLSX template-a za uređivanje podataka,
- import iz izmenjenog XLSX-a nazad u `.db` (XML) uz backup originala,
- generisanje JSON report-a sa informacijama i listom duplikata.

Glavni fajlovi:

- `src/parseXml.js` — XML → JSON parser
- `src/generateTemplate.js` — JSON → XLSX generator
- `src/importXlsx.js` — XLSX → XML importer + report generator
- `src/cli.js` — jednostavan CLI za pokretanje operacija
- `src/reportServer.js` — (opcionalno) Express server koji služi UI i /open endpoint
- `ui/index.html` — mala Vue strana za pregled report-a i otvaranje Excel-a

Prerequisites
-------------

- Node.js (preporučeno v18+ / testovano na v20)
- Excel (ako želite da koristite /open endpoint koji pozicionira Excel na sheet/row)

Instalacija
-----------

U projektu (PowerShell):

```powershell
npm install
```

Brzi primeri (PowerShell)
-------------------------

CLI reference — kratka uputstva
--------------------------------

Kratko objašnjenje dostupnih CLI komandi i najčešće korišćenih opcija:

- Parsiranje `.db` u JSON
	- Svrha: brzo inspekcija XML strukture u čitljivom obliku.
	- Komanda:

```powershell
node src/cli.js parse "<putanja-do-db.db>" <out.json>
```

- Generisanje XLSX template-a
	- Svrha: napravi multi-sheet Excel fajl koji može uređivati neprogramer.
	- Komanda:

```powershell
node src/cli.js generate-template <parsed.json> <out-template.xlsx>
```

- Import iz XLSX nazad u `.db` (sa backup-om i izveštajem)
	- Svrha: učitaj izmenjeni Excel i napiši novu XML `.db` datoteku.
	- Osnovna komanda:

```powershell
node src/cli.js import-xlsx <in.xlsx> <out.db> "<original.db>" [--report] [--report-full] [--force] [--report-out <path>]
```

	- Najvažnije opcije:
		- `--report` — sačuvaj sažeti JSON report (`reports/report-<ts>.json`).
		- `--report-full` — sačuvaj detaljniji report sa primerima i duplikatima.
		- `--force` — prepiši `out.db` ako već postoji (koristiti pažljivo).
		- `--report-out <path>` — sačuvaj report na prilagođenu lokaciju.

- Pokretanje lokalnog UI/report servera

```powershell
node src/reportServer.js
# Otvori u browseru: http://localhost:3000/
```

Ovo je sažeta referenca; za brze primere pogledaj dole "Brzi primeri" koji sadrže konkretne komande i objašnjenja.

1) Parsiraj `.db` u JSON

```powershell
node src/cli.js parse "db files\materials_test.db" materials_parsed.json
```

2) Generiši XLSX template iz parsanog JSON-a

```powershell
node src/cli.js generate-template materials_parsed.json materials_template.xlsx
```

3) Importuj izmenjeni XLSX nazad u `.db` i generiši report (full)

```powershell
node src/cli.js import-xlsx materials_template.xlsx materials_new.db "db files\materials_test.db" --report --report-full
```

Ova komanda:
- napiše novi `materials_new.db` (XML),
- napravi backup originala `db files\materials_test.db.bak.<ts>`,
- sačuva report JSON u `reports/report-<ts>.json`.

4) Pokreni lokalni report server (servira `ui/index.html` i endpoint `/report` i `/open`)

```powershell
node src/reportServer.js
# zatim u browser-u otvori: http://localhost:3000/
```

Kako testirati UI i otvaranje Excel-a
-----------------------------------

1. Generiši report (`--report`) koristeći `import-xlsx` kao gore.
2. Pokreni server `node src/reportServer.js`.
3. Otvori u pretraživaču: `http://localhost:3000/`.
4. U UI-u možeš upisati ime report fajla (npr. `reports/report-1770371594772.json`) ili ostaviti prazno da se učita zadnji/standarni report.
5. Postavi putanju do Excel fajla u polje `Excel path` (npr. `C:\_smartWOP\_Projects\MatAPP\materials_template.xlsx`).
6. Klikni `Load report` pa `Open first` pored duplicate-a — server će pozvati `/open` endpoint koji pokuša da otvori Excel i pozicionira se na dati sheet/row.

Napomena o Windows `/open` ponašanju
----------------------------------

- Endpoint `/open` koristi PowerShell COM automaciju (Excel.Application). To radi samo na Windows mašinama sa instaliranim Excel-om i odgovarajućim dozvolama za COM.
- Ako Excel nije prisutan ili COM kreiranje ne uspe, server će vratiti JSON grešku. U tom slučaju možemo primeniti fallback: otvoriti datoteku u Explorer-u ili koristiti `Start-Process` za otvaranje fajla bez pozicioniranja.

Gde su report-i i backup-i
--------------------------

- JSON report-i su u `reports/` kao `report-<ts>.json`.
- Backup originalnog DB (pre zamene) se čuva pored originala kao `*.bak.<ts>`.

Troubleshooting (brzo)
----------------------

- Greška `ERR_REQUIRE_ESM` pri require('uuid') => rešenje: projekat koristi `crypto.randomUUID()` pa nema potrebe za `uuid` paketom.
- Ako server ne može da otvori Excel preko /open: proveri da li Excel postoji i da li PowerShell može kreirati COM objekat (pokreni PowerShell kao administrator kad testiraš).

Kako napraviti screenshot UI-a (ako želiš meni poslati)
-----------------------------------------------------

1. Otvori `http://localhost:3000/` i prikaži deo koji želiš.
2. U Windows-u pritisni Win+Shift+S i izaberi region (clipboard).
3. Sačuvaj sliku u fajl ili direktno nalepi u e-mail/poruku.

Šta dalje mogu da uradim za tebe
--------------------------------

- Dovršim UI (filteri, pretraga, download report-a).
- Dodam fallback za /open (Start-Process) ako COM ne radi.
- Napravim 2-3 unit testa za parser i importer.

Ako želiš, mogu odmah da: (a) testiram /open na tvojoj mašini (treba potvrda da imaš Excel), ili (b) doteram UI na bolje UX. Koju opciju biraš?

---
README generisan automatski — fajl: `README.md` u root-u projekta.
# MatAPP materials tool (MVP)

Ovo je jednostavan Node CLI alat koji će pomoći da se vaš `.db` (XML) fajl konvertuje u JSON/XLSX i nazad.

Brzi koraci za početak (PowerShell):

```powershell
# instaliraj zavisnosti (iz root foldera projekta)
npm install fast-xml-parser exceljs xmlbuilder2 commander uuid ajv --save
# opcionalno za testove
npm install --save-dev mocha chai

# za parsiranje primjera u workspace-u (pretpostavlja se da je materials_test.db u folderu projekta):
node src/cli.js parse materials_test.db materials_parsed.json
```

Dalji koraci: generisanje Excel template-a, importer i validacija. Radićemo taj korak-po-korak zajedno.