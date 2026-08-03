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
 
app = Flask(__name__)
 
# Anthropic API key wordt via environment variable ingesteld (NIET in code hardcoden)
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY) if ANTHROPIC_API_KEY else None
 
# In-memory opslag van resultaten per sessie (voor een eerste werkende versie)
# Later kan dit naar een echte database, maar voor nu is dit prima om te testen.
SESSIONS = {}
 
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
  </style>
</head>
<body>
  <h1>Factuur Upload &amp; Extractie</h1>
  <p>Sleep meerdere factuurfoto's hieronder, of klik om te selecteren. Geen limiet op aantal bestanden.</p>
 
  <div id="dropzone">Sleep foto's hierheen, of klik om te kiezen</div>
  <input type="file" id="fileInput" multiple accept="image/*">
 
  <div id="cameraRow">
    <button type="button" id="cameraBtn">📷 Bonnetje fotograferen</button>
  </div>
  <input type="file" id="cameraInput" accept="image/*" capture="environment">
 
  <div id="fileList"></div>
 
  <button id="processBtn" disabled>Verwerk facturen</button>
  <a id="downloadBtn" href="#"><button type="button">Download Excel</button></a>
 
  <script>
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const cameraBtn = document.getElementById('cameraBtn');
    const cameraInput = document.getElementById('cameraInput');
    const fileList = document.getElementById('fileList');
    const processBtn = document.getElementById('processBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    let files = [];
    let sessionId = null;
 
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
        image_bytes = file.read()
        image_b64 = base64.standard_b64encode(image_bytes).decode("utf-8")
        media_type = file.mimetype or "image/jpeg"
 
        prompt = (
            "Dit is een foto van een factuur of bon. In Nederland zijn er drie BTW-tarieven: "
            "21%, 9% en 0%. Een factuur kan meerdere tarieven tegelijk bevatten (bijvoorbeeld "
            "deels 21% en deels 9%). Haal de volgende gegevens eruit en geef ALLEEN een JSON "
            "object terug, niets anders, geen uitleg, geen markdown:\n"
            "{\n"
            '  "leverancier": "...",\n'
            '  "factuurdatum": "DD-MM-JJJJ",\n'
            '  "factuurnummer": "...",\n'
            '  "omschrijving": "...",\n'
            '  "btw_regels": [\n'
            '    {"btw_percentage": 21, "bedrag_excl_btw": 0.00, "btw_bedrag": 0.00},\n'
            '    {"btw_percentage": 9, "bedrag_excl_btw": 0.00, "btw_bedrag": 0.00}\n'
            "  ]\n"
            "}\n"
            "BELANGRIJK: maak voor elk BTW-tarief dat op de factuur voorkomt een apart object in "
            "de 'btw_regels' lijst, ook als er maar één tarief is (dan bevat de lijst één object). "
            "Gebruik alleen de tarieven 21, 9 of 0. Als een veld niet leesbaar of niet aanwezig is, "
            "gebruik dan null. Gebruik een punt als decimaalteken."
        )
 
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1000,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": image_b64,
                            },
                        },
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
 
 
@app.route("/health")
def health():
    return jsonify({"status": "ok", "api_key_configured": client is not None})
 
 
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
