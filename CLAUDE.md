# Document Review Tool — Claude Code Build Instructions

## Projektziel

Baue eine lokal laufende Flask-Webanwendung, die Dokumente anhand einer konfigurierbaren Rolle mit der Claude API reviewed und als kommentiertes DOCX mit farbiger Hervorhebung zurückgibt.

---

## Stack

- **Backend:** Python 3.11+, Flask
- **DOCX-Verarbeitung:** `python-docx`, `docx` (npm, für Output-Generierung)
- **KI:** Anthropic Python SDK (`anthropic`)
- **Frontend:** Vanilla HTML/CSS/JS (kein Framework)
- **Konfiguration:** JSON-Datei (`roles.json`)

---

## Projektstruktur

```
doc-review/
├── app.py                  # Flask-App, Routen
├── reviewer.py             # Kernlogik: Dokument → Claude → DOCX-Output
├── roles.json              # Rollenkonfigurationen
├── requirements.txt
├── templates/
│   ├── index.html          # Hauptseite (Upload + Review)
│   └── config.html         # Rollenverwaltung
├── static/
│   └── style.css
└── uploads/                # temporäre Uploads (gitignore)
```

---

## Funktionsumfang

### Hauptseite (`/`)

- Dropdown zur Rollenauswahl (aus `roles.json` befüllt)
- Zwei Input-Optionen (Tab-Switch):
  - **Datei-Upload:** DOCX oder TXT
  - **Texteingabe:** Textarea
- Button „Review starten"
- Nach Verarbeitung: Download-Button für Output-DOCX
- Statusanzeige während Verarbeitung (Spinner + Statustext)

### Config-Seite (`/config`)

- Liste aller vorhandenen Rollen
- Pro Rolle editierbar:
  - **Name** (Anzeigename im Dropdown)
  - **System Prompt** (Rollenbeschreibung für Claude, Textarea)
  - **Sprache** (Dropdown: Deutsch / Englisch)
- Buttons: Rolle hinzufügen, speichern, löschen
- Speicherung in `roles.json`

---

## `roles.json` Format

```json
[
  {
    "id": "it-contract-lawyer",
    "name": "IT-Vertragsanwalt",
    "language": "de",
    "system_prompt": "Du bist ein erfahrener IT-Vertragsanwalt mit Schwerpunkt auf Managed Services und Outsourcing-Verträgen. Du prüfst Dokumente auf: fehlende Leistungsdefinitionen, unklare Haftungsregelungen, fehlende KPIs, einseitige Kündigungsrechte, und Lücken im Eskalationsprozess. Sei präzise und praxisorientiert."
  },
  {
    "id": "compliance-officer",
    "name": "Compliance Officer",
    "language": "de",
    "system_prompt": "Du bist ein Compliance Officer mit Fokus auf DSGVO, IT-Sicherheit und regulatorische Anforderungen. Du identifizierst Datenschutzrisiken, fehlende Nachweispflichten und Verstöße gegen Best Practices."
  }
]
```

---

## Kernlogik: `reviewer.py`

### Ablauf

1. **Dokument lesen:**
   - DOCX: Text mit `python-docx` extrahieren (paragraphenweise, Struktur erhalten)
   - TXT/Textarea: direkt verwenden

2. **Claude API aufrufen:**

   Sende den Dokumenttext mit folgendem Prompt-Schema:

   ```
   [System Prompt der gewählten Rolle]

   Analysiere das folgende Dokument sorgfältig.

   Gib deine Analyse als strukturiertes JSON zurück. Keine weiteren Erklärungen, nur valides JSON.

   Format:
   {
     "summary": "Kurze Gesamtbewertung (2-4 Sätze)",
     "findings": [
       {
         "id": 1,
         "severity": "red" | "orange" | "green",
         "quote": "Exakte Textstelle aus dem Dokument (so kurz wie möglich, max. 200 Zeichen)",
         "comment": "Erläuterung des Befundes und Handlungsempfehlung"
       }
     ]
   }

   Severity-Bedeutung:
   - red: Kritisches Problem, sofortiger Handlungsbedarf
   - orange: Problematisch, sollte überarbeitet werden
   - green: Positiv hervorzuheben oder nur geringes Risiko

   Dokument:
   ---
   [DOKUMENTTEXT]
   ---
   ```

3. **JSON parsen** (Fehlerbehandlung: retry bei invalide JSON, max. 2 Versuche)

4. **DOCX generieren** mit `python-docx`:

   - Einfügen der Zusammenfassung am Anfang (Heading + normaler Paragraph, hellgrauer Hintergrund)
   - Dann den Originaltext paragraphenweise wiedergeben
   - Für jeden Fund (`finding`): die `quote`-Textstelle im Originaltext suchen und:
     - **Hervorhebung:** Hintergrundfarbe des Run setzen:
       - `red` → `FF0000` (Rot, 40% opacity via `FFB3B3` als Highlight)
       - `orange` → `FFA500` (Orange via `FFD9A0`)
       - `green` → `00AA00` (Grün via `C6EFCE`)
     - **Kommentar:** Word-Kommentar an dieser Stelle einfügen (mit `comment`-Text)
   - Am Ende: Anhang „Review-Zusammenfassung" mit tabellarischer Auflistung aller Findings (Severity | Textstelle | Kommentar)

### DOCX-Kommentare einfügen

Nutze die XML-Manipulation via `python-docx` direkt (nicht das npm `docx`-Paket für den Output, da wir ein bestehendes Dokument aufbauen):

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

def add_comment(doc, paragraph, run, comment_text, author="Document Reviewer"):
    """Fügt einen Word-Kommentar zu einem Run hinzu."""
    # comments.xml befüllen und Referenz im document.xml setzen
    # Implementiere gemäß python-docx XML-Manipulation
```

Implementiere `add_comment()` vollständig über direkte XML-Manipulation der `comments`-Part. Orientiere dich an der python-docx-Dokumentation für Custom XML Parts.

### Highlight-Farben setzen

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def highlight_run(run, hex_color):
    """Setzt Hintergrundfarbe eines Runs (Shading)."""
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    rPr.append(shd)
```

---

## Flask-Routen (`app.py`)

```
GET  /              → index.html
POST /review        → Verarbeitung, gibt DOCX-Download zurück
GET  /config        → config.html
POST /config/save   → roles.json speichern
POST /config/delete → Rolle löschen
GET  /roles         → JSON-Liste aller Rollen (für Dropdown)
```

### `/review` Endpunkt

- Akzeptiert: `multipart/form-data` mit `role_id`, optional `file` (DOCX/TXT) oder `text`
- Validierung: Entweder Datei oder Text muss vorhanden sein
- Ruft `reviewer.py` auf
- Gibt DOCX als `application/vnd.openxmlformats-officedocument.wordprocessingml.document` zurück
- Dateiname: `review_[rollenname]_[timestamp].docx`
- Bei Fehler: JSON `{"error": "..."}` mit HTTP 500

---

## UI-Details

### Farbcodierung im Frontend

Zeige nach dem Review eine kurze Legende:
- 🔴 Rot = Kritisch
- 🟠 Orange = Überarbeitungsbedarf
- 🟢 Grün = In Ordnung

### Stil

- Sauber, professionell, keine Spielereien
- Hintergrund weiß/hellgrau, Akzentfarbe dunkelblau (`#1a3a5c`)
- Responsive für Desktop (mobile ist optional)
- Font: System-UI / Segoe UI

---

## Umgebungsvariablen

```bash
ANTHROPIC_API_KEY=sk-...   # Pflicht
FLASK_SECRET_KEY=...        # Für Flask Sessions (optional, random default)
```

Lade via `python-dotenv` aus `.env` Datei.

---

## `requirements.txt`

```
flask>=3.0
anthropic>=0.25
python-docx>=1.1
python-dotenv>=1.0
```

---

## Fehlerbehandlung

- Kein API-Key → klare Fehlermeldung beim Start
- Invalide JSON-Antwort von Claude → 1 Retry, dann generischer Fehler
- Datei zu groß (>10MB) → HTTP 413 mit Hinweis
- Unsupported Dateiformat → HTTP 400
- `quote` nicht im Dokument gefunden → Kommentar trotzdem als Fußnote am Dokumentende anhängen, nicht überspringen

---

## Startbefehl (lokal ohne Docker)

```bash
python app.py
# App läuft auf http://localhost:5000
```

---

## Docker

### `Dockerfile`

```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads

EXPOSE 5000

CMD ["python", "app.py"]
```

### `docker-compose.yml`

```yaml
services:
  doc-review:
    build: .
    restart: unless-stopped
    ports:
      - "5000:5000"
    environment:
      - ANTHROPIC_API_KEY=${ANTHROPIC_API_KEY}
      - FLASK_SECRET_KEY=${FLASK_SECRET_KEY}
    volumes:
      - ./roles.json:/app/roles.json   # Rollen bleiben beim Rebuild erhalten
      - ./uploads:/app/uploads         # temporäre Uploads außerhalb des Containers
```

### `.env` auf dem Server

```bash
ANTHROPIC_API_KEY=sk-...
FLASK_SECRET_KEY=ein-langer-zufaelliger-string
```

### Build & Start

```bash
docker compose up -d --build
```

### Hinweise für Claude Code

- Flask muss auf `0.0.0.0` binden, nicht `127.0.0.1`:
  ```python
  app.run(host="0.0.0.0", port=5000)
  ```
- `uploads/`-Verzeichnis im Container anlegen (`RUN mkdir -p uploads`)
- `roles.json` als Volume mounten, damit Rollenänderungen über Container-Neustarts hinaus erhalten bleiben
- Kein Gunicorn notwendig für Single-User-Betrieb; bei Bedarf nachrüstbar

---

## Was Claude Code NICHT tun soll

- Kein gunicorn/nginx-Setup (Flask reicht für Single-User)
- Kein User-Auth / Login
- Keine Datenbank — alles dateibasiert
- Kein Test-Framework — nur funktionierende App

---

## Reihenfolge der Implementierung

1. `roles.json` mit zwei Beispielrollen anlegen
2. `app.py` mit allen Routen (erst Stubs)
3. `reviewer.py` — Claude-Integration + JSON-Parsing
4. DOCX-Generierung mit Highlighting
5. DOCX-Kommentare via XML-Manipulation
6. `index.html` + `config.html` + `style.css`
7. End-to-End-Test mit einem Beispieldokument

---

## Hinweise für Claude Code

- Bei DOCX-Kommentaren: Nutze **direkte XML-Manipulation** über `python-docx`'s `element`-API. Die High-Level-API unterstützt keine Kommentare nativ.
- Die `quote`-Suche im Dokument muss **fuzzy** sein (stripped whitespace, case-insensitive matching), da Leerzeichen und Zeilenumbrüche im Originaltext variieren können.
- Halte `reviewer.py` und `app.py` **strikt getrennt** — keine Flask-Imports in `reviewer.py`.
- Der Claude-Aufruf soll `claude-sonnet-4-5` verwenden mit `max_tokens=4096`.
