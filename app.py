
"""
Factuur Upload & Extractie Tool
--------------------------------
Simpele webapp:
1. Upload pagina waar je meerdere factuurfoto's tegelijk kunt slepen (geen limiet zoals in de chat)
2. Elke foto wordt naar de Anthropic API gestuurd om gegevens te extraheren (leverancier, datum, bedrag, BTW)
3. Resultaten worden verzameld en als Excel-bestand gedownload
 
Gebouwd met Flask (lichtgewicht, makkelijk te draaien op Railway).
"""
 
import os
import io
import json
import base64
import uuid
from datetime import datetime
 
from flask import Flask, request, jsonify, send_file, render_template_string
from openpyxl import Workbook
import anthropic
import eboekhouden
 
app = Flask(__name__)
 
# Anthropic API key wordt via environment variable ingesteld (NIET in code hardcoden)
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY) if ANTHROPIC_API_KEY else None
 
# In-memory opslag van resultaten per sessie (voor een eerste werkende versie)
# Later kan dit naar een echte database, maar voor nu is dit prima om te testen.
SESSIONS = {}
 
# Relevante grootboekrekeningen voor BurgerMe Venlo (subset uit het volledige
# rekeningschema, alleen betaalmiddelen + inkoop-gerelateerde rekeningen).
# Voor nu hardcoded voor deze ene klant; bij uitbreiding naar andere klanten
# maken we dit per klant instelbaar.
BURGERME_BETAALREKENINGEN = [
    ("1000", "Kas"),
    ("1010", "Bank"),
    ("23102", "Pin"),
    ("23103", "Thuisbezorgd"),
    ("23104", "Mollie Web"),
    ("23106", "Uber"),
]
 
BURGERME_KOSTENREKENINGEN = [
    ("7021", "Inkoop BTW 21%"),
    ("7022", "Inkoop BTW 9%"),
    ("7024", "Inkoop BTW 0%"),
    ("7012", "Overige inkopen"),
    ("47001", "Huur milkshakemachine"),
    ("70000", "Frituurolie"),
    ("70001", "Keuken"),
    ("70003", "Milkshake"),
    ("70004", "Dranken 9% btw"),
    ("70005", "Dranken 21% btw"),
    ("70006", "Verpakkingen"),
    ("70007", "Speelgoed"),
    ("70008", "Emballage Sligro"),
    ("70009", "Emballage vanGelder"),
    ("70010", "Emballage Hanos"),
    ("70011", "Emballage flessen/blikjes"),
    ("42200", "Brandstof"),
    ("42950", "Parkeergelden"),
    ("45300", "Kantoorartikelen en drukwerk"),
    ("45351", "Softwarekosten"),
    ("44991", "Kosten Thuisbezorgd"),
    ("44992", "Kosten Mollie"),
    ("44996", "Kosten Uber"),
    ("4500", "Contributies en abonnementen"),
]
 
# Generieke kostenrekening per BTW-tarief - dit is de standaard voor élke
# leverancier, tenzij een specifieke regel (zie hieronder) iets anders aangeeft
# (bijv. een vast terugkerende kostenpost zoals een machine-huur).
BTW_KOSTENREKENING_DEFAULT = {
    21: "7021",  # Inkoop BTW 21%
    9: "7022",   # Inkoop BTW 9%
    0: "7024",   # Inkoop BTW 0%
}
 
# Vaste boekingsvoorbeelden per leverancier voor BurgerMe Venlo, aangeleverd door
# Wais. Zodra de herkende leverancier hierop matcht (ongeacht hoofdletters, op basis
# van een deel van de naam), gebruiken we relatiecode + betalingstermijn hieruit.
# Optionele sleutels per leverancier:
#   - kostenrekening: vaste rekening voor ALLE regels (bijv. een vaste huurpost)
#   - kostenrekening_per_btw: dict, override voor één specifiek BTW-tarief
#     (bijv. de 0%-regel van een emballage-leverancier gaat naar hun eigen
#     emballage-rekening i.p.v. de generieke 7024)
#   - kruispost: {"kostenrekening": ..., "kruispost_rekening": ...} - voor
#     bezorgplatforms (Thuisbezorgd/Uber Eats/Mollie): naast de kostenregel wordt
#     automatisch een tweede, negatieve verrekeningsregel geboekt tegen de
#     kruispost-rekening (zo staat het ook al in e-Boekhouden)
#   - eu_leverancier: True - gebruik de BTW-code voor "Leveringen/diensten van
#     binnen EU 0%" (verlegde BTW) i.p.v. de gewone "Geen btw"-code
# Sleutel: (deel van) leveranciersnaam in kleine letters.
BURGERME_LEVERANCIER_MAPPING = {
    "dupon": {
        "relatiecode": "9",
        "betalingstermijn": "14",
        # geen vaste kostenrekening: valt terug op BTW-tarief (7021/7022/7024)
        # let op: bij een specifieke huurpost (bijv. milkshakemachine) gebruikt
        # Wais handmatig 47001 i.p.v. de generieke rekening
    },
    "ijsexpress": {
        "relatiecode": "0121",
        "betalingstermijn": "14",
    },
    "thuisbezorgd": {
        "relatiecode": None,  # nog te bevestigen door Wais - laat leeg tot bekend
        "betalingstermijn": "0",
        "kruispost": {"kostenrekening": "44991", "kruispost_rekening": "23103"},
        "zoeknaam": "Thuisbezorgd",
    },
    "takeaway.com": {
        # Zelfde bedrijf als Thuisbezorgd.nl (Takeaway.com is de eigenaar/merknaam
        # op sommige facturen) - zelfde boekingsafspraken
        "relatiecode": None,  # nog te bevestigen door Wais - laat leeg tot bekend
        "betalingstermijn": "0",
        "kruispost": {"kostenrekening": "44991", "kruispost_rekening": "23103"},
        "zoeknaam": "Thuisbezorgd",  # de bestaande relatie heet vermoedelijk zo, niet 'Takeaway.com'
    },
    "uber eats": {
        "relatiecode": "0122",
        "betalingstermijn": "0",
        "kruispost": {"kostenrekening": "44996", "kruispost_rekening": "23106"},
    },
    "mollie": {
        "relatiecode": "0104",
        "betalingstermijn": "14",
        "kruispost": {"kostenrekening": "44992", "kruispost_rekening": "23104"},
    },
    "simplydelivery": {
        "relatiecode": "0113",
        "betalingstermijn": "14",
        "kostenrekening": "45351",  # Softwarekosten
        "eu_leverancier": True,
    },
    "van gelder": {
        "relatiecode": "0005",
        "betalingstermijn": "0",
        "kostenrekening_per_btw": {0: "70009"},  # Emballage van Gelder
    },
    "sligro": {
        "relatiecode": "0004",
        "betalingstermijn": "14",
        "kostenrekening_per_btw": {0: "70008"},  # Emballage Sligro
    },
    "meyer quick service": {
        "relatiecode": "0003",
        "betalingstermijn": "14",
        "eu_leverancier": True,
        # geen vaste kostenrekening: valt terug op BTW-tarief, met 70006
        # (verpakkingen) als handmatige keuze-optie voor verpakking-regels
    },
}
 
 
def zoek_leverancier_mapping(leverancier_naam):
    """Zoekt of de herkende leveranciersnaam matcht met een vaste mapping."""
    if not leverancier_naam:
        return None
    naam_lower = leverancier_naam.lower()
    for sleutel, mapping in BURGERME_LEVERANCIER_MAPPING.items():
        if sleutel in naam_lower:
            return mapping
    return None
 
UPLOAD_PAGE = """
<!DOCTYPE html>
<html lang="nl">
<head>
  <meta charset="UTF-8">
  <title>Factuur Upload</title>
  <style>
    body { font-family: Arial, sans-serif; max-width: 700px; margin: 40px auto; padding: 0 20px; }
    h1 { font-size: 22px; }
    #dropzone {
      border: 2px dashed #999; border-radius: 8px; padding: 40px;
      text-align: center; color: #666; cursor: pointer; margin-bottom: 20px;
    }
    #dropzone.dragover { background: #f0f7ff; border-color: #2563eb; }
    #fileInput { display: none; }
    #cameraInput { display: none; }
    #cameraRow { margin-bottom: 16px; }
    #cameraBtn {
      background: #111827; color: white; border: none; padding: 12px 20px;
      border-radius: 6px; cursor: pointer; font-size: 15px; width: 100%;
    }
    #fileList { margin-bottom: 20px; }
    .file-row { display: flex; justify-content: space-between; padding: 6px 0; border-bottom: 1px solid #eee; font-size: 14px; }
    .status-wait { color: #999; }
    .status-ok { color: #16a34a; }
    .status-err { color: #dc2626; }
    button { background: #2563eb; color: white; border: none; padding: 10px 18px; border-radius: 6px; cursor: pointer; font-size: 14px; }
    button:disabled { background: #aaa; cursor: not-allowed; }
    #downloadBtn { background: #16a34a; margin-top: 16px; display: none; }
    #boekBtn { background: #7c3aed; margin-top: 16px; margin-left: 8px; display: none; }
    #reviewTable { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px; display: none; }
    #reviewTable th, #reviewTable td { border: 1px solid #ddd; padding: 6px; text-align: left; }
    #reviewTable th { background: #f3f4f6; }
    #reviewTable select, #reviewTable input { width: 100%; font-size: 13px; border: 1px solid #ccc; border-radius: 4px; padding: 3px; }
    .boek-status { font-size: 12px; margin-top: 2px; }
  </style>
</head>
<body>
  <h1>Factuur Upload &amp; Extractie</h1>
  <p>Sleep meerdere factuurfoto's hieronder, of klik om te selecteren. Geen limiet op aantal bestanden.</p>
 
  <div id="dropzone">Sleep foto's hierheen, of klik om te kiezen</div>
  <input type="file" id="fileInput" multiple accept="image/*,.pdf">
 
  <div id="cameraRow">
    <button type="button" id="cameraBtn">📷 Bonnetje fotograferen</button>
  </div>
  <input type="file" id="cameraInput" accept="image/*" capture="environment">
 
  <div id="fileList"></div>
 
  <button id="processBtn" disabled>Verwerk facturen</button>
  <a id="downloadBtn" href="#"><button type="button">Download Excel</button></a>
  <button id="boekBtn" type="button">Boek in e-Boekhouden</button>
 
  <table id="reviewTable">
    <thead>
      <tr>
        <th>Bestand</th><th>Leverancier</th><th>Datum</th><th>Factuurnr</th>
        <th>BTW%</th><th>Excl. BTW</th><th>BTW bedrag</th>
        <th>Betaalrekening</th><th>Kostenrekening</th><th>Relatiecode</th><th>Betalingstermijn</th><th>Status</th>
      </tr>
    </thead>
    <tbody id="reviewBody"></tbody>
  </table>
 
  <script>
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const cameraBtn = document.getElementById('cameraBtn');
    const cameraInput = document.getElementById('cameraInput');
    const fileList = document.getElementById('fileList');
    const processBtn = document.getElementById('processBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const boekBtn = document.getElementById('boekBtn');
    const reviewTable = document.getElementById('reviewTable');
    const reviewBody = document.getElementById('reviewBody');
    let files = [];
    let sessionId = null;
    let processedRows = []; // bewaart de (bewerkbare) data per verwerkte factuur, incl. rekeningkeuzes
    let rekeningOpties = { betaalrekeningen: [], kostenrekeningen: [] };
 
    fetch('/api/rekeningen').then(r => r.json()).then(d => { rekeningOpties = d; });
 
    dropzone.addEventListener('click', () => fileInput.click());
    dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('dragover'); });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', e => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      addFiles(e.dataTransfer.files);
    });
    fileInput.addEventListener('change', e => addFiles(e.target.files));
 
    // Camera-knop: opent direct de camera op mobiel (geen omweg via fotorol)
    cameraBtn.addEventListener('click', () => cameraInput.click());
    cameraInput.addEventListener('change', e => {
      addFiles(e.target.files);
      cameraInput.value = ''; // reset zodat je meteen nog een foto kunt maken
    });
 
    function addFiles(newFiles) {
      for (const f of newFiles) files.push(f);
      renderList();
      processBtn.disabled = files.length === 0;
    }
 
    function renderList() {
      fileList.innerHTML = '';
      files.forEach((f, i) => {
        const row = document.createElement('div');
        row.className = 'file-row';
        row.id = 'row-' + i;
        row.innerHTML = `<span>${f.name}</span><span class="status-wait" id="status-${i}">wachtend</span>`;
        fileList.appendChild(row);
      });
    }
 
    processBtn.addEventListener('click', async () => {
      processBtn.disabled = true;
      processBtn.textContent = 'Bezig...';
 
      // Maak een nieuwe sessie aan
      const sRes = await fetch('/api/session', { method: 'POST' });
      const sData = await sRes.json();
      sessionId = sData.session_id;
 
      for (let i = 0; i < files.length; i++) {
        document.getElementById('status-' + i).textContent = 'verwerken...';
        const formData = new FormData();
        formData.append('file', files[i]);
        formData.append('session_id', sessionId);
        try {
          const res = await fetch('/api/process', { method: 'POST', body: formData });
          const data = await res.json();
          if (data.success) {
            document.getElementById('status-' + i).textContent = 'OK';
            document.getElementById('status-' + i).className = 'status-ok';
            processedRows.push(data.data);
          } else {
            document.getElementById('status-' + i).textContent = 'fout: ' + (data.error || 'onbekend');
            document.getElementById('status-' + i).className = 'status-err';
          }
        } catch (err) {
          document.getElementById('status-' + i).textContent = 'fout';
          document.getElementById('status-' + i).className = 'status-err';
        }
      }
 
      processBtn.textContent = 'Klaar';
      downloadBtn.href = '/api/download/' + sessionId;
      downloadBtn.style.display = 'inline-block';
 
      renderReviewTable();
      if (processedRows.length > 0) {
        boekBtn.style.display = 'inline-block';
      }
    });
 
    function maakSelect(opties, geselecteerd, onChange) {
      const select = document.createElement('select');
      opties.forEach(([code, omschrijving]) => {
        const opt = document.createElement('option');
        opt.value = code;
        opt.textContent = code + ' - ' + omschrijving;
        if (code === geselecteerd) opt.selected = true;
        select.appendChild(opt);
      });
      select.addEventListener('change', e => onChange(e.target.value));
      return select;
    }
 
    function renderReviewTable() {
      reviewBody.innerHTML = '';
      reviewTable.style.display = processedRows.length ? 'table' : 'none';
 
      processedRows.forEach((rij, rijIndex) => {
        const regels = (rij.btw_regels && rij.btw_regels.length) ? rij.btw_regels : [{}];
        regels.forEach((regel, regelIndex) => {
          const tr = document.createElement('tr');
 
          const bestandTd = document.createElement('td');
          bestandTd.textContent = regelIndex === 0 ? (rij.bestandsnaam || '') : '';
          tr.appendChild(bestandTd);
 
          [rij.leverancier, rij.factuurdatum, rij.factuurnummer].forEach(waarde => {
            const td = document.createElement('td');
            td.textContent = regelIndex === 0 ? (waarde ?? '') : '';
            tr.appendChild(td);
          });
 
          [regel.btw_percentage, regel.bedrag_excl_btw, regel.btw_bedrag].forEach(waarde => {
            const td = document.createElement('td');
            td.textContent = waarde ?? '';
            tr.appendChild(td);
          });
 
          const betaalTd = document.createElement('td');
          betaalTd.appendChild(maakSelect(
            rekeningOpties.betaalrekeningen, regel.betaalrekening || '1000',
            v => { regel.betaalrekening = v; }
          ));
          tr.appendChild(betaalTd);
 
          const kostenTd = document.createElement('td');
          kostenTd.appendChild(maakSelect(
            rekeningOpties.kostenrekeningen, regel.kostenrekening || '7012',
            v => { regel.kostenrekening = v; }
          ));
          tr.appendChild(kostenTd);
 
          const relatieTd = document.createElement('td');
          if (regelIndex === 0) {
            const relatieInput = document.createElement('input');
            relatieInput.type = 'text';
            relatieInput.value = rij.relatiecode || '';
            relatieInput.placeholder = 'verplicht bij factuurnr.';
            relatieInput.addEventListener('input', e => { rij.relatiecode = e.target.value; });
            relatieTd.appendChild(relatieInput);
          }
          tr.appendChild(relatieTd);
 
          const termijnTd = document.createElement('td');
          if (regelIndex === 0) {
            const termijnInput = document.createElement('input');
            termijnInput.type = 'text';
            termijnInput.value = rij.betalingstermijn || '0';
            termijnInput.addEventListener('input', e => { rij.betalingstermijn = e.target.value; });
            termijnTd.appendChild(termijnInput);
          }
          tr.appendChild(termijnTd);
 
          const statusTd = document.createElement('td');
          statusTd.className = 'boek-status';
          if (regelIndex === 0) statusTd.id = `boekstatus-${rijIndex}`;
          tr.appendChild(statusTd);
 
          reviewBody.appendChild(tr);
        });
      });
    }
 
    boekBtn.addEventListener('click', async () => {
      boekBtn.disabled = true;
      boekBtn.textContent = 'Bezig met boeken...';
 
      const res = await fetch('/api/boek/' + sessionId, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(processedRows),
      });
      const data = await res.json();
 
      (data.resultaten || []).forEach((resultaat, i) => {
        const statusEl = document.getElementById(`boekstatus-${i}`);
        if (!statusEl) return;
        if (resultaat.success) {
          statusEl.textContent = 'Geboekt (' + resultaat.mutatienummers.join(', ') + ')';
          statusEl.style.color = '#16a34a';
        } else {
          statusEl.textContent = 'Fout: ' + resultaat.error;
          statusEl.style.color = '#dc2626';
        }
      });
 
      boekBtn.textContent = 'Klaar met boeken';
    });
  </script>
</body>
</html>
"""
 
 
@app.route("/")
def index():
    return render_template_string(UPLOAD_PAGE)
 
 
@app.route("/api/session", methods=["POST"])
def create_session():
    session_id = str(uuid.uuid4())
    SESSIONS[session_id] = []
    return jsonify({"session_id": session_id})
 
 
@app.route("/api/process", methods=["POST"])
def process_invoice():
    if client is None:
        return jsonify({"success": False, "error": "ANTHROPIC_API_KEY niet ingesteld op de server"}), 500
 
    session_id = request.form.get("session_id")
    file = request.files.get("file")
 
    if not session_id or session_id not in SESSIONS:
        return jsonify({"success": False, "error": "ongeldige sessie"}), 400
    if not file:
        return jsonify({"success": False, "error": "geen bestand ontvangen"}), 400
 
    try:
        file_bytes = file.read()
        file_b64 = base64.standard_b64encode(file_bytes).decode("utf-8")
        is_pdf = (file.mimetype == "application/pdf") or (file.filename or "").lower().endswith(".pdf")
        media_type = "application/pdf" if is_pdf else (file.mimetype or "image/jpeg")
 
        prompt = (
            "Dit is een foto of PDF-document van een factuur of bon. In Nederland zijn er drie BTW-tarieven: "
            "21%, 9% en 0%. Een factuur kan meerdere tarieven tegelijk bevatten (bijvoorbeeld "
            "deels 21% en deels 9%). Kijk ook goed naar hoe er betaald is: staat er expliciet "
            "'contant', 'cash' of een kassabon-kenmerk op, of staan er juist bankgegevens "
            "(IBAN, rekeningnummer, 'overgemaakt', 'betaald via bank') op de factuur? "
            "Bepaal voor elke BTW-regel ook of het gaat om INKOOP VAN MATERIAAL/GOEDEREN "
            "die direct met de omzet te maken heeft (bijv. grondstoffen, dranken, "
            "verpakkingen, ingrediënten - dingen die doorverkocht of verwerkt worden in "
            "producten), of om een ALGEMENE KOST (bijv. huur, verzekering, abonnement, "
            "brandstof, kantoorkosten, marketing - kosten die niets met specifieke omzet "
            "te maken hebben). Haal de volgende gegevens eruit en geef ALLEEN een JSON "
            "object terug, niets anders, geen uitleg, geen markdown:\n"
            "{\n"
            '  "leverancier": "...",\n'
            '  "factuurdatum": "DD-MM-JJJJ",\n'
            '  "factuurnummer": "...",\n'
            '  "omschrijving": "...",\n'
            '  "betaalmethode": "contant" of "bank" of null als onduidelijk,\n'
            '  "btw_regels": [\n'
            '    {"btw_percentage": 21, "bedrag_excl_btw": 0.00, "btw_bedrag": 0.00, "kostentype": "inkoop_goederen" of "algemene_kost"},\n'
            '    {"btw_percentage": 9, "bedrag_excl_btw": 0.00, "btw_bedrag": 0.00, "kostentype": "inkoop_goederen" of "algemene_kost"}\n'
            "  ]\n"
            "}\n"
            "BELANGRIJK: maak voor elk BTW-tarief dat op de factuur voorkomt een apart object in "
            "de 'btw_regels' lijst, ook als er maar één tarief is (dan bevat de lijst één object). "
            "Gebruik alleen de tarieven 21, 9 of 0. Vul 'betaalmethode' alleen in als je het "
            "ECHT duidelijk op de bon/factuur ziet staan - gok niet, gebruik anders null. "
            "Als een ander veld niet leesbaar of niet aanwezig is, gebruik dan null. Gebruik "
            "een punt als decimaalteken."
        )
 
        content_block = (
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": file_b64,
                },
            }
            if is_pdf
            else {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": media_type,
                    "data": file_b64,
                },
            }
        )
 
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1000,
            messages=[
                {
                    "role": "user",
                    "content": [
                        content_block,
                        {"type": "text", "text": prompt},
                    ],
                }
            ],
        )
 
        raw_text = "".join(block.text for block in message.content if block.type == "text")
        raw_text = raw_text.strip()
        # Verwijder eventuele markdown-codeblocks als die er per ongeluk in zitten
        if raw_text.startswith("```"):
            raw_text = raw_text.strip("`")
            if raw_text.startswith("json"):
                raw_text = raw_text[4:]
            raw_text = raw_text.strip()
 
        data = json.loads(raw_text)
        data["bestandsnaam"] = file.filename
        # Standaardwaarden toevoegen aan elke BTW-regel, die de gebruiker in het
        # controlescherm zelf kan aanpassen vóór het boeken. Betaalrekening baseren
        # we op wat de AI zag staan (contant -> Kas, bank -> Bank); bij twijfel
        # gokken we niet en blijft Kas de gok totdat de gebruiker het corrigeert.
        betaalmethode = (data.get("betaalmethode") or "").lower()
        standaard_betaalrekening = "1010" if betaalmethode == "bank" else "1000"
 
        # Check of deze leverancier een vast boekingsvoorbeeld heeft - zo ja, dan
        # gebruiken we relatiecode/betalingstermijn daaruit. Kostenrekening komt,
        # tenzij de leverancier een vaste rekening (of tarief-specifieke
        # uitzondering) heeft, standaard uit het BTW-tarief van de regel
        # (21% -> 7021, 9% -> 7022, 0% -> 7024).
        mapping = zoek_leverancier_mapping(data.get("leverancier"))
        if mapping:
            data["relatiecode"] = mapping.get("relatiecode") or ""
            data["betalingstermijn"] = mapping["betalingstermijn"]
            kruispost = mapping.get("kruispost")
            # Bij het bezorgplatform-patroon hoort de kostenregel op de kruispost-
            # kostenrekening (bijv. 44991), niet op de generieke BTW-standaard.
            vaste_kostenrekening = mapping.get("kostenrekening") or (kruispost or {}).get("kostenrekening")
            kostenrekening_per_btw = mapping.get("kostenrekening_per_btw") or {}
            data["kruispost"] = kruispost
            data["eu_leverancier"] = mapping.get("eu_leverancier", False)
            # Voor de automatische relatie-opzoeking gebruiken we een kortere,
            # betrouwbaardere zoeknaam als die is opgegeven (de volledige
            # AI-herkende naam matcht vaak niet exact met hoe de relatie in
            # e-Boekhouden is vastgelegd).
            data["zoeknaam_relatie"] = mapping.get("zoeknaam") or data.get("leverancier")
        else:
            data["relatiecode"] = ""
            data["betalingstermijn"] = "0"
            vaste_kostenrekening = None
            kostenrekening_per_btw = {}
            data["kruispost"] = None
            data["eu_leverancier"] = False
            data["zoeknaam_relatie"] = data.get("leverancier")
 
        for regel in data.get("btw_regels", []) or []:
            btw_pct = regel.get("btw_percentage")
            kostentype = regel.get("kostentype")
            if btw_pct in kostenrekening_per_btw:
                regel["kostenrekening"] = kostenrekening_per_btw[btw_pct]
            elif vaste_kostenrekening:
                regel["kostenrekening"] = vaste_kostenrekening
            elif kostentype == "algemene_kost":
                # Algemene kosten (huur, verzekering, abonnementen, etc.) horen niet
                # op de inkoop-BTW-rekeningen thuis - "Overige inkopen" is hier een
                # duidelijke plek-houder die om handmatige controle vraagt.
                regel["kostenrekening"] = "7012"
            else:
                # Inkoop van materiaal/goederen die met de omzet te maken heeft
                regel["kostenrekening"] = BTW_KOSTENREKENING_DEFAULT.get(btw_pct, "7012")
            regel["betaalrekening"] = standaard_betaalrekening
        SESSIONS[session_id].append(data)
 
        return jsonify({"success": True, "data": data})
 
    except json.JSONDecodeError:
        return jsonify({"success": False, "error": "kon resultaat niet als JSON lezen"}), 500
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
 
 
@app.route("/api/download/<session_id>")
def download_excel(session_id):
    if session_id not in SESSIONS:
        return "Sessie niet gevonden", 404
 
    rows = SESSIONS[session_id]
 
    wb = Workbook()
    ws = wb.active
    ws.title = "Facturen"
 
    headers = [
        "Bestandsnaam", "Leverancier", "Factuurdatum", "Factuurnummer",
        "BTW %", "Bedrag excl. BTW", "BTW bedrag", "Bedrag incl. BTW", "Omschrijving",
    ]
    ws.append(headers)
 
    for row in rows:
        btw_regels = row.get("btw_regels") or []
        if not btw_regels:
            # Fallback: geen btw_regels gevonden, toch één lege regel zodat de factuur niet verdwijnt
            btw_regels = [{"btw_percentage": None, "bedrag_excl_btw": None, "btw_bedrag": None}]
 
        for regel in btw_regels:
            excl = regel.get("bedrag_excl_btw")
            btw = regel.get("btw_bedrag")
            incl = (excl + btw) if isinstance(excl, (int, float)) and isinstance(btw, (int, float)) else None
            ws.append([
                row.get("bestandsnaam"),
                row.get("leverancier"),
                row.get("factuurdatum"),
                row.get("factuurnummer"),
                regel.get("btw_percentage"),
                excl,
                btw,
                incl,
                row.get("omschrijving"),
            ])
 
    # Kolombreedte iets vergroten voor leesbaarheid
    for col_idx, header in enumerate(headers, start=1):
        ws.column_dimensions[chr(64 + col_idx)].width = max(15, len(header) + 2)
 
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
 
    filename = f"facturen_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
 
 
@app.route("/api/rekeningen")
def get_rekeningen():
    return jsonify({
        "betaalrekeningen": BURGERME_BETAALREKENINGEN,
        "kostenrekeningen": BURGERME_KOSTENREKENINGEN,
    })
 
 
@app.route("/api/boek/<session_id>", methods=["POST"])
def boek_facturen(session_id):
    if session_id not in SESSIONS:
        return jsonify({"success": False, "error": "sessie niet gevonden"}), 404
 
    # De frontend stuurt de (eventueel door de gebruiker aangepaste) rijen mee,
    # zodat we altijd boeken met wat er op het scherm staat, niet met de
    # oorspronkelijke AI-extractie.
    edited_rows = request.get_json(silent=True) or []
    resultaten = []
 
    for row in edited_rows:
        try:
            regels = row.get("btw_regels") or []
            betaalrekening = (regels[0].get("betaalrekening") if regels else None) or "1000"
            mutatienummers = eboekhouden.boek_bonnetje(
                klant_prefix="BURGERME",
                leverancier=row.get("leverancier"),
                factuurdatum_ddmmjjjj=row.get("factuurdatum"),
                factuurnummer=row.get("factuurnummer"),
                omschrijving=row.get("omschrijving"),
                betaalrekening=betaalrekening,
                regels=regels,
                relatiecode=row.get("relatiecode") or "",
                betalingstermijn=row.get("betalingstermijn") or "0",
                eu_leverancier=row.get("eu_leverancier", False),
                kruispost=row.get("kruispost"),
                zoeknaam_relatie=row.get("zoeknaam_relatie") or row.get("leverancier"),
            )
            resultaten.append({
                "bestandsnaam": row.get("bestandsnaam"),
                "success": True,
                "mutatienummers": mutatienummers,
            })
        except eboekhouden.EBoekhoudenError as e:
            resultaten.append({"bestandsnaam": row.get("bestandsnaam"), "success": False, "error": str(e)})
        except Exception as e:
            resultaten.append({"bestandsnaam": row.get("bestandsnaam"), "success": False, "error": str(e)})
 
    return jsonify({"resultaten": resultaten})
 
 
@app.route("/health")
def health():
    return jsonify({"status": "ok", "api_key_configured": client is not None})
 
 
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
 
