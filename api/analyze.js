'use strict';

require('dotenv').config();

const Anthropic = require('@anthropic-ai/sdk');
const { createClient } = require('@supabase/supabase-js');
const pdfParse = require('pdf-parse');
const XLSX = require('xlsx');
const { spezzaReport, salvaAnalisiTemp } = require('./leadgate');

const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB
const MIN_TEXT_LENGTH = 200;

// Supabase client — stesso DB di BalanceScan, service_role key
const supabase = (process.env.SUPABASE_URL && process.env.SUPABASE_SERVICE_ROLE_KEY)
  ? createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY)
  : null;

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

Subito dopo la tabella affidabilità inserisci su una riga da solo, senza altri caratteri: <!--GATE-->

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
Scegli UNO dei tre giudizi e usa esattamente il blocco HTML corrispondente,
seguito da un paragrafo di testo normale (max 4 righe, nessun grassetto, nessun maiuscolo):

Se la situazione è positiva:
<div class="verdict IN_LINEA">IN LINEA CON GLI OBIETTIVI</div>
<p>[testo descrittivo della situazione, tono professionale e diretto]</p>

Se richiede attenzione:
<div class="verdict ATTENZIONE">ATTENZIONE RICHIESTA</div>
<p>[testo descrittivo della situazione, tono professionale e diretto]</p>

Se richiede intervento urgente:
<div class="verdict URGENTE">INTERVENTO URGENTE NECESSARIO</div>
<p>[testo descrittivo della situazione, tono professionale e diretto]</p>

Il report termina qui. NON aggiungere sezioni successive.

═══════════════════════════════════════════════════
FORMATO RISPOSTA FINALE — OBBLIGATORIO
═══════════════════════════════════════════════════
Rispondi ESCLUSIVAMENTE con un oggetto JSON valido in questo formato:
{
  "ragioneSociale": "Nome Azienda S.r.l.",
  "esito": "ATTENZIONE",
  "report": "...tutto l'HTML del report..."
}

- "ragioneSociale": ragione sociale rilevata dal documento, o "Azienda" se assente
- "esito": UNO di questi valori esatti → "IN_LINEA" | "ATTENZIONE" | "URGENTE"
- "report": l'intero report HTML (sezioni 1-9 con <!--GATE--> dopo la sezione 2), come stringa JSON escaped
- Nessun testo fuori dal JSON, nessun markdown wrapper (\`\`\`json)
═══════════════════════════════════════════════════`;

// ── CTA HTML: versione co-firmata (token valido) ──────────────────────────
function bloccoCoFirmato(nomeStudio, ragioneSociale) {
  return `
<hr>
<h2>Passo Successivo Consigliato</h2>
<p>Questa proiezione mostra dove sta andando <strong>${ragioneSociale}</strong> se la rotta non cambia, e segnala i punti da tenere d'occhio prima della chiusura d'anno.</p>
<p>I numeri dicono <em>cosa</em> sta succedendo. Per capire <em>perché</em> e intervenire in tempo, il passo naturale è una lettura operativa dei processi che generano questi risultati.</p>
<p>Ne parli con <strong>${nomeStudio}</strong>, che segue questa analisi.</p>
<p><em>Analisi operativa a cura di Castelli Consulting, partner tecnico per il margine e l'efficienza operativa.</em></p>
<hr>`;
}

// ── CTA HTML: versione diretta (nessun token) ─────────────────────────────
function ctaDiretta(ragioneSociale) {
  return `
<hr>
<h2>Passo Successivo Consigliato</h2>
<p>Questa proiezione mostra dove sta andando <strong>${ragioneSociale}</strong> se la rotta non cambia.</p>
<p>I numeri dicono cosa sta succedendo. Per capire perché e intervenire in tempo, il <strong>Check-Up Margine e Rischi di Castelli Consulting</strong> individua le cause operative e costruisce una roadmap concreta a 90 giorni.</p>
<p>Scopri il Check-Up: <a href="https://www.castelliconsulting.it/check-up-strategico/" target="_blank" rel="noopener">www.castelliconsulting.it/check-up-strategico/</a><br>Scrivi a: <a href="mailto:fcastelli@castelliconsulting.it">fcastelli@castelliconsulting.it</a></p>
<hr>`;
}

// ── Validazione token commercialista ─────────────────────────────────────
async function validaCommercialista(tokenGrezzo) {
  if (!supabase || !tokenGrezzo) return { valido: false, nomeStudio: null };

  const token = String(tokenGrezzo).replace(/[^a-f0-9]/gi, '').slice(0, 64);
  if (!token) return { valido: false, nomeStudio: null };

  const { data, error } = await supabase
    .from('commercialisti')
    .select('nome_studio, attivo, scadenza')
    .eq('token', token)
    .maybeSingle();

  if (error || !data) return { valido: false, nomeStudio: null };

  const nonScaduto = !data.scadenza || new Date(data.scadenza) >= new Date();
  const valido = data.attivo && nonScaduto;

  return { valido, nomeStudio: valido ? data.nome_studio : null };
}

// ── Log utilizzo (tabella condivisa con BalanceScan) ──────────────────────
async function logUtilizzo({ token, nomeStudio, ragioneSociale, esito }) {
  if (!supabase) return;
  try {
    await supabase.from('utilizzi_balancescan').insert({
      token:           token || null,
      nome_studio:     nomeStudio || null,
      ragione_sociale: ragioneSociale || null,
      esito:           esito || null,
      app:             'radar',
    });
  } catch (e) {
    console.error('Log utilizzo Radar fallito:', e.message);
  }
}

// ── Parsing multipart ─────────────────────────────────────────────────────
function parseMultipart(body, boundary) {
  const parts = [];
  const boundaryBuffer = Buffer.from(`--${boundary}`);
  let start = 0;

  while (start < body.length) {
    const boundaryIdx = body.indexOf(boundaryBuffer, start);
    if (boundaryIdx === -1) break;

    const headerStart = boundaryIdx + boundaryBuffer.length + 2;
    const headerEnd = body.indexOf(Buffer.from('\r\n\r\n'), headerStart);
    if (headerEnd === -1) break;

    const headers = body.slice(headerStart, headerEnd).toString();
    const contentStart = headerEnd + 4;

    const nextBoundaryIdx = body.indexOf(boundaryBuffer, contentStart);
    if (nextBoundaryIdx === -1) break;

    const contentEnd = nextBoundaryIdx - 2;
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

// ── Excel → testo strutturato ─────────────────────────────────────────────
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

// ── Handler principale ────────────────────────────────────────────────────
module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Metodo non consentito' });
  }

  // Token commercialista dalla querystring (?k=...)
  const tokenGrezzo = req.query.k || null;

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
      error: 'Formato file non supportato. Carica un PDF o un file Excel (.xlsx, .xls) esportato dal gestionale.',
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
      error: "Configurazione server incompleta. Contatta l'amministratore.",
    });
  }

  // Validazione token (in parallelo con niente — non blocca la generazione)
  const { valido, nomeStudio } = await validaCommercialista(tokenGrezzo);

  const client = new Anthropic({ apiKey });

  const userMessage = `Analizza la seguente situazione contabile infrannuale e produci il report completo secondo le istruzioni ricevute.

DOCUMENTO ESTRATTO:
---
${extractedText.slice(0, 80000)}
---`;

  let reportHtml    = '';
  let ragioneSociale = 'Azienda';
  let esitoReport   = '';   // IN_LINEA | ATTENZIONE | URGENTE
  let esito         = 'errore'; // ok | errore (per log Supabase)

  try {
    const message = await client.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 8192,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userMessage }],
    });

    const rawText = message.content
      .filter(block => block.type === 'text')
      .map(block => block.text)
      .join('');

    // Estrai JSON dalla risposta (Claude a volte aggiunge whitespace o backtick)
    const jsonMatch = rawText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      try {
        const parsed = JSON.parse(jsonMatch[0]);
        reportHtml    = parsed.report || rawText;
        ragioneSociale = (parsed.ragioneSociale || '').trim() || 'Azienda';
        esitoReport    = parsed.esito || '';
      } catch {
        reportHtml = rawText;
      }
    } else {
      reportHtml = rawText;
    }

    esito = 'ok';
  } catch (err) {
    console.error('Errore Claude API:', err);

    await logUtilizzo({ token: tokenGrezzo, nomeStudio, ragioneSociale: null, esito: 'errore_ai' });

    if (err.status === 401) {
      return res.status(500).json({ error: "Chiave API non valida. Contatta l'amministratore." });
    }
    if (err.status === 429) {
      return res.status(429).json({ error: 'Troppe richieste. Attendi qualche secondo e riprova.' });
    }
    if (err.status === 529 || err.message?.includes('overloaded')) {
      return res.status(503).json({ error: 'Il servizio AI è temporaneamente sovraccarico. Riprova tra qualche minuto.' });
    }
    return res.status(500).json({ error: "Errore durante l'analisi AI. Riprova tra qualche istante." });
  }

  // CTA deterministica (mai generata da Claude)
  const sezioneFinale = valido
    ? bloccoCoFirmato(nomeStudio, ragioneSociale)
    : ctaDiretta(ragioneSociale);

  const reportCompleto = reportHtml + sezioneFinale;

  // Log utilizzo su Supabase
  await logUtilizzo({ token: tokenGrezzo, nomeStudio, ragioneSociale, esito });

  // Token valido → report completo subito
  if (valido) {
    return res.status(200).json({ gated: false, html: reportCompleto });
  }

  // Nessun token → gate: salva report, restituisci solo anteprima
  try {
    const analisiId = await salvaAnalisiTemp(supabase, {
      reportCompleto,
      ragioneSociale,
      esito: esitoReport,
    });
    const { anteprima } = spezzaReport(reportCompleto);
    return res.status(200).json({ gated: true, analisiId, anteprima });
  } catch (e) {
    console.error('salvaAnalisiTemp fallito:', e.message);
    // Fallback: mostra il report completo se il salvataggio fallisce
    return res.status(200).json({ gated: false, html: reportCompleto });
  }
};
