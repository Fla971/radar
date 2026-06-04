'use strict';

require('dotenv').config();

const Anthropic = require('@anthropic-ai/sdk');
const pdfParse = require('pdf-parse');
const XLSX = require('xlsx');

const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB
const MIN_TEXT_LENGTH = 200;

const SYSTEM_PROMPT = `Sei un controller di gestione esperto in PMI italiane manifatturiere con 20 anni di esperienza.

Stai analizzando una SITUAZIONE CONTABILE INFRANNUALE — dati provvisori e parziali, NON un bilancio definitivo.

REGOLE FONDAMENTALI:
- Analizza SOLO i dati effettivamente presenti nel documento
- Segnala SEMPRE quando un dato è assente o inaffidabile
- Non inventare mai dati mancanti
- Distingui sempre tra dati certi e stime/proiezioni
- Le rimanenze di magazzino infrannuali sono spesso non aggiornate — segnalalo esplicitamente se rilevi questa situazione
- Gli ammortamenti infrannuali sono spesso assenti — segnalalo e indica l'impatto sulla lettura del risultato

═══════════════════════════════════════════════════
FORMATO OUTPUT — REGOLE ASSOLUTE
═══════════════════════════════════════════════════
Produci ESCLUSIVAMENTE HTML valido e semantico.

NON usare MAI sintassi Markdown:
  ✗ NO asterischi per il grassetto  (**testo**)
  ✗ NO cancelletti per i titoli  (## Titolo)
  ✗ NO trattini per i separatori  (---)
  ✗ NO underscore per il corsivo  (_testo_)

Usa SOLO questi tag HTML:
  <h2>  titoli di sezione principali
  <h3>  sottotitoli di sezione
  <h4>  titoli minori
  <p>   paragrafi di testo
  <ul> <li>  liste puntate
  <ol> <li>  liste numerate
  <table> <thead> <tbody> <tr> <th> <td>  tabelle
  <strong>  testo in grassetto
  <em>  testo in corsivo
  <hr>  separatore orizzontale
  <div class="...">  contenitori con classi semantiche

Classi div disponibili (usa dove appropriato):
  <div class="alert-box">  per avvertenze importanti
  <div class="verdict IN_LINEA">  giudizio finale verde
  <div class="verdict ATTENZIONE">  giudizio finale arancione
  <div class="verdict URGENTE">  giudizio finale rosso
  <div class="kpi-grid">  griglia KPI affiancati
  <div class="kpi-card">  singola card KPI dentro kpi-grid

NON includere tag: html, head, body, style, script.
═══════════════════════════════════════════════════

## 1. INTESTAZIONE REPORT
- Ragione sociale (se presente)
- Periodo di riferimento rilevato dal documento
- Data di elaborazione (oggi)
- Avvertenza: "Dati provvisori — situazione infrannuale"

## 2. AFFIDABILITÀ DEI DATI
Tabella con semaforo per ogni area:
🟢 Dato affidabile / 🟡 Dato parziale / 🔴 Dato assente o inaffidabile

Aree da valutare:
- Ricavi
- Costi operativi
- Margine lordo
- Posizione bancaria
- Crediti clienti
- Debiti fornitori
- Rimanenze magazzino
- Ammortamenti
- Risultato netto stimato

## 3. SITUAZIONE ATTUALE (solo dati affidabili)

### Ricavi YTD
- Valore assoluto
- Proiezione fine anno (estrapolazione lineare con formula esplicita)
- Confronto anno precedente stesso periodo (solo se dato presente)

### Struttura Costi (se disponibile)
- Incidenza % principali voci di costo sui ricavi
- Anomalie rispetto a valori tipici di settore PMI

### Posizione Finanziaria
- Liquidità disponibile
- Debiti bancari a breve vs lungo termine
- Saldo netto

### Capitale Circolante (se dati disponibili)
- DSO — Giorni crediti clienti
  Soglie: 🟢 <60gg | 🟡 60-90gg | 🔴 >90gg
- DPO — Giorni debiti fornitori
  Soglie: 🟢 >60gg | 🟡 30-60gg | 🔴 <30gg

## 4. PROIEZIONI FINE ANNO
Segnala sempre: "Proiezione indicativa basata su andamento lineare — soggetta a variazioni"

- Ricavi stimati fine anno
- Risultato operativo stimato (solo se ammortamenti presenti o stimabili)
- Posizione finanziaria stimata fine anno
- Scenario ottimistico / base / pessimistico (±10% sui ricavi)

## 5. SEGNALI DI ALLERTA
Massimo 3, ordinati per urgenza.
Solo quelli supportati da dati concreti presenti nel documento. Con evidenza numerica.

## 6. AZIONI PRIORITARIE (entro 30-60 giorni)
Massimo 3 azioni concrete e specifiche.
Basate esclusivamente sui dati rilevati.
Con impatto atteso stimato.

## 7. DATI MANCANTI
Lista chiara di cosa richiedere al commercialista per completare l'analisi:
"Per una analisi più completa, richiedere:"
- [lista voci mancanti rilevate]

## 8. DOMANDE PER LA PROSSIMA RIUNIONE
3 domande operative specifiche basate sui numeri trovati — non domande generiche.

## 9. SINTESI DIREZIONALE
4 righe massimo.
Giudizio sulla traiettoria attuale:
IN LINEA CON OBIETTIVI / ATTENZIONE RICHIESTA / INTERVENTO URGENTE NECESSARIO`;

## 10. PASSO SUCCESSIVO CONSIGLIATO

Aggiungi sempre questa sezione finale:

---

PASSO SUCCESSIVO CONSIGLIATO

Questa analisi proietta la traiettoria di [Ragione Sociale] sulla base dei dati infrannuali disponibili.

Per trasformare queste proiezioni in azioni concrete con impatto misurabile sul margine operativo, il Check-Up Margine e Rischi di Castelli Consulting è il passo naturale.

Hai già il radar. Ora ti serve la rotta.

Scopri il Check-Up: https://www.castelliconsulting.it/check-up-strategico/
Scrivi a: fcastelli@castelliconsulting.it

---

function parseMultipart(body, boundary) {
  const parts = [];
  const boundaryBuffer = Buffer.from(`--${boundary}`);
  let start = 0;

  while (start < body.length) {
    const boundaryIdx = body.indexOf(boundaryBuffer, start);
    if (boundaryIdx === -1) break;

    const headerStart = boundaryIdx + boundaryBuffer.length + 2; // skip \r\n
    const headerEnd = body.indexOf(Buffer.from('\r\n\r\n'), headerStart);
    if (headerEnd === -1) break;

    const headers = body.slice(headerStart, headerEnd).toString();
    const contentStart = headerEnd + 4;

    const nextBoundaryIdx = body.indexOf(boundaryBuffer, contentStart);
    if (nextBoundaryIdx === -1) break;

    const contentEnd = nextBoundaryIdx - 2; // strip trailing \r\n
    const content = body.slice(contentStart, contentEnd);

    const nameMatch = headers.match(/name="([^"]+)"/);
    const filenameMatch = headers.match(/filename="([^"]+)"/);
    const contentTypeMatch = headers.match(/Content-Type:\s*([^\r\n]+)/i);

    parts.push({
      name: nameMatch ? nameMatch[1] : null,
      filename: filenameMatch ? filenameMatch[1] : null,
      contentType: contentTypeMatch ? contentTypeMatch[1].trim() : 'text/plain',
      data: content,
    });

    start = nextBoundaryIdx;
  }

  return parts;
}

function xlsxToText(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const lines = [];

  for (const sheetName of workbook.SheetNames) {
    lines.push(`\n=== FOGLIO: ${sheetName} ===\n`);
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    for (const row of rows) {
      if (row.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
        lines.push(row.map(c => String(c).trim()).join('\t'));
      }
    }
  }

  return lines.join('\n');
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Metodo non consentito' });
  }

  // Collect raw body
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }
  const rawBody = Buffer.concat(chunks);

  if (rawBody.length > MAX_FILE_SIZE) {
    return res.status(413).json({
      error: 'Il file supera il limite di 10 MB. Carica un file più piccolo.',
    });
  }

  const contentType = req.headers['content-type'] || '';
  const boundaryMatch = contentType.match(/boundary=([^\s;]+)/);
  if (!boundaryMatch) {
    return res.status(400).json({ error: 'Richiesta non valida: boundary multipart mancante.' });
  }

  const parts = parseMultipart(rawBody, boundaryMatch[1]);
  const filePart = parts.find(p => p.filename);

  if (!filePart) {
    return res.status(400).json({ error: 'Nessun file ricevuto.' });
  }

  const filename = filePart.filename.toLowerCase();
  const isPdf = filename.endsWith('.pdf') || filePart.contentType.includes('pdf');
  const isExcel =
    filename.endsWith('.xlsx') ||
    filename.endsWith('.xls') ||
    filePart.contentType.includes('spreadsheet') ||
    filePart.contentType.includes('excel');

  if (!isPdf && !isExcel) {
    return res.status(415).json({
      error:
        'Formato file non supportato. Carica un PDF o un file Excel (.xlsx, .xls) esportato dal gestionale.',
    });
  }

  let extractedText = '';

  try {
    if (isPdf) {
      const pdfData = await pdfParse(filePart.data);
      extractedText = pdfData.text;
    } else {
      extractedText = xlsxToText(filePart.data);
    }
  } catch (err) {
    console.error('Errore parsing file:', err);
    return res.status(422).json({
      error: 'Impossibile leggere il file. Verifica che non sia corrotto e riprova.',
    });
  }

  if (!extractedText || extractedText.trim().length < MIN_TEXT_LENGTH) {
    const message = isPdf
      ? 'Il PDF sembra essere scansionato (immagine) e non contiene testo selezionabile. ' +
        'Chiedi al commercialista il documento in formato digitale (PDF nativo o Excel).'
      : 'Il file Excel sembra vuoto o non contiene dati leggibili. ' +
        'Verifica che il file non sia protetto da password e che contenga dati nella prima scheda.';
    return res.status(422).json({ error: message });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    console.error('ANTHROPIC_API_KEY non configurata');
    return res.status(500).json({
      error: 'Configurazione server incompleta. Contatta l\'amministratore.',
    });
  }

  const client = new Anthropic({ apiKey });

  const userMessage = `Analizza la seguente situazione contabile infrannuale e produci il report completo secondo le istruzioni ricevute.

DOCUMENTO ESTRATTO:
---
${extractedText.slice(0, 80000)}
---`;

  try {
    const message = await client.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 4096,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userMessage }],
    });

    const reportHtml = message.content
      .filter(block => block.type === 'text')
      .map(block => block.text)
      .join('');

    return res.status(200).json({ html: reportHtml });
  } catch (err) {
    console.error('Errore Claude API:', err);

    if (err.status === 401) {
      return res.status(500).json({ error: 'Chiave API non valida. Contatta l\'amministratore.' });
    }
    if (err.status === 429) {
      return res.status(429).json({
        error: 'Troppe richieste. Attendi qualche secondo e riprova.',
      });
    }
    if (err.status === 529 || err.message?.includes('overloaded')) {
      return res.status(503).json({
        error: 'Il servizio AI è temporaneamente sovraccarico. Riprova tra qualche minuto.',
      });
    }

    return res.status(500).json({
      error: 'Errore durante l\'analisi AI. Riprova tra qualche istante.',
    });
  }
};
