# Factuur Upload &amp; Extractie Tool

Eerste werkende versie. Doel: facturen uploaden (zonder de 20-bestanden-limiet van de
Claude chat), automatisch laten uitlezen door de Anthropic API, en als Excel downloaden.

## Wat dit MVP wel en niet doet

**Wel:**
- Upload pagina waar je meerdere foto's tegelijk kunt slepen
- Elke foto wordt los naar de Anthropic API gestuurd voor extractie (leverancier, datum,
  bedrag, BTW, factuurnummer, omschrijving)
- Resultaten verzamelen in een Excel-bestand om te downloaden

**Nog niet (bewust, voor latere uitbreiding):**
- Geen permanente database (resultaten verdwijnen als de server herstart — voor een
  eerste test is dat prima)
- Geen login/beveiliging — iedereen met de link kan de pagina gebruiken. **Niet
  geschikt om nu al naar klanten te sturen.**
- Geen automatische import-format voor e-Boekhouden (Excel-kolommen zijn generiek,
  kunnen later aangepast worden)
- Verwerkt foto's één voor één (niet parallel) — bij heel veel foto's duurt het dus
  even, maar je hoeft niet te wachten tot alles klaar is om te zien wat werkt

## Stap 1: Lokaal testen (optioneel, kan ook direct naar Railway)

```bash
cd factuur-app
pip install -r requirements.txt
export ANTHROPIC_API_KEY="jouw-api-key-hier"
python app.py
```
Ga naar `http://localhost:5000` in je browser.

## Stap 2: Naar GitHub zetten

Railway deployt vanuit een GitHub repository.

```bash
cd factuur-app
git init
git add .
git commit -m "Eerste versie factuur upload tool"
```

Maak een nieuwe (lege) repository aan op GitHub, bijvoorbeeld `factuur-app`, en volg
de instructies die GitHub toont om je lokale code daarheen te pushen (meestal iets als):

```bash
git remote add origin https://github.com/JOUW-GEBRUIKERSNAAM/factuur-app.git
git branch -M main
git push -u origin main
```

## Stap 3: Koppelen aan Railway

1. Log in op je Railway account (het account dat je al eerder had aangemaakt)
2. Klik op "New Project" → "Deploy from GitHub repo"
3. Selecteer de `factuur-app` repository
4. Railway herkent automatisch dat het een Python/Flask app is (via `railway.json`)
5. Ga naar het tabblad **Variables** van je service en voeg toe:
   - `ANTHROPIC_API_KEY` = jouw Anthropic API key (deze haal je op via
     console.anthropic.com → API Keys)
6. Railway geeft je daarna een publieke URL (iets als `factuur-app-production.up.railway.app`)

## Stap 4: Testen

Open de Railway-URL in je browser, sleep een paar testfacturen erin, klik op
"Verwerk facturen", en download het Excel-bestand om te checken of de extractie
goed gaat.

## Volgende stappen (pas als de basis werkt)

- Login/wachtwoord toevoegen zodat niet iedereen erbij kan
- Per-klant mapjes/sessies, zodat je meerdere klanten naast elkaar kunt verwerken
- Excel-kolommen aanpassen naar exact e-Boekhouden import-format
- "Uitzonderingen"-tabblad: facturen waarbij de extractie onzeker is (bijv. lage
  resolutie, geknipte foto) automatisch apart zetten in plaats van foute data door te
  laten — precies het uitzondering-systeem waar we het eerder over hadden
