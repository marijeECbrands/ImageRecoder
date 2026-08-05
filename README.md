# Fashion Image Renamer v5

Windows-app voor het automatisch koppelen, classificeren en hernoemen van fashionfoto's.

## Primaire match

Een bestandsnaam zoals:

```text
EX13-261070_EX153.jpg
```

wordt op de eerste `_` gesplitst:

- `EX13-261070` wordt exact gezocht in `Artikelnr. Leverancier`;
- `EX153` moet voorkomen in `Externe omschrijving` van dezelfde Excel-regel.

De nieuwe naam wordt:

```text
Artikel nr.-Kleurcode-SoortBeeld.extensie
```

Wanneer geen eenduidige naammatch mogelijk is, vergelijkt de app de foto automatisch visueel met foto's die wel succesvol gekoppeld zijn. Alleen onbetrouwbare resultaten komen in `controle_nodig` of `geen_match`.

## Vereiste Excel-kolommen

- `Artikelnr. Leverancier`
- `Externe omschrijving`
- `Artikel nr.`
- `Kleurcode`

De kolomnamen en matchinstellingen zijn aanpasbaar in `matchconfig.json`.
De nummering voor soorten beeld staat in `fotocategorieen.json`.

## Interface

De app toont tijdens verwerking de huidige stap, een bewegende bezig-indicator, voortgang en verstreken tijd. In het resultaat blijft de kolom `Zekerheid` aanwezig. Resultaten kunnen worden gefilterd op vrije tekst en status.

## GitHub Actions

Plaats de projectbestanden in de hoofdmap van de repository. De workflows moeten staan in:

```text
.github/workflows/build-exe.yml
.github/workflows/build-setup.yml
```

Onder **Actions** verschijnen:

- `Build losse EXE`
- `Build Setup.exe`

De artifacts bevatten `FashionImageRenamer_v5.exe` en `FashionImageRenamer_v5_Setup.exe`.

## Correctie v5.1
De matcher gebruikt alleen het eerste en tweede deel van de bestandsnaam. Een eventueel derde deel (`_1`, `_2`, `_3`) is het fotonummer en wordt niet meegenomen in de Excel-match. Ook worden externe codes met en zonder `EX` als gelijk behandeld, bijvoorbeeld `707` en `EX707`.

## Nieuw in v6 – snelheidsoptimalisaties
- Excel-match gebruikt een index voor snelle lookup.
- AI krijgt automatisch een verkleinde JPEG-kopie (standaard max. 1200 px); de output blijft het originele bestand in volledige resolutie.
- Artikelgroepen worden parallel naar de AI gestuurd (standaard 4 tegelijk, instelbaar 1–8).
- Exacte duplicaten worden lokaal met SHA-256 herkend en niet naar de AI gestuurd.
- Bijna-identieke beelden worden alleen binnen hetzelfde reeds betrouwbaar gematchte artikel met een zeer strikte perceptuele hash gecontroleerd; uitgesloten originelen gaan naar `duplicaten`.
- AI-resultaten worden lokaal gecachet in `classification_cache.json`, zodat identieke beelden bij een volgende run niet opnieuw geclassificeerd hoeven te worden.
- Visuele fallback blijft pas actief wanneer de naammatch faalt en wordt eveneens parallel uitgevoerd.

### Nieuwe instellingen
- **Parallelle AI-taken**: standaard 4. Bij rate-limitfouten terugzetten naar 2 of 3.
- **AI max px**: standaard 1200. Dit verandert alleen wat naar OpenAI wordt gestuurd, niet de uiteindelijke outputfoto.
- **Exacte + bijna-identieke beelden uitfilteren**: standaard ingeschakeld. Duplicaten worden nooit verwijderd; ze worden naar de map `duplicaten` gekopieerd en in het CSV-rapport vermeld.


## v6.1
- Categorie 8: Artikel back.
- Dubbele Artikel front-classificaties worden extra gecontroleerd op front/back.
- Twee echte frontbeelden van hetzelfde artikel krijgen status DUPLICAAT_BEELDTYPE en gaan naar duplicaten.
- Onzekere dubbele categorieën krijgen HERCLASSIFICATIE_NODIG.

## v6.2
- Bugfix: `ai_data_url is not defined` opgelost.
- 5 = Detail artikel: mouw, kraag, knopen, zak, zoom, rits, borduring en andere constructiedetails.
- 6 = Detail stof: materiaal, textuur, weving, print/patroon en stofoppervlak.
- 7 = Artikel front.
- 8 = Artikel back.
- Categorie 5 en 6 mogen meerdere foto's per artikel hebben zonder automatisch als duplicaat te gelden.
- Categorie 8 wordt na hernoemen in de aparte map `artikel_back` geplaatst.
- De bestaande extra front/back-controle bij dubbele categorieën blijft actief.
