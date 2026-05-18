# CLAUDE.md — RadarFinanziario™
## Castelli Consulting | Aggiornato: maggio 2026

> Questo file viene letto automaticamente da Claude Code ad ogni sessione.
> Contiene tutto il contesto necessario per lavorare su questo progetto
> senza ripetere errori già incontrati.

---

## 1. Cos'è questo progetto

**RadarFinanziario™** è una web app che analizza situazioni contabili
infrannuali (PDF o Excel) di PMI italiane tramite Claude AI e produce
proiezioni a fine anno, segnali di allerta e azioni prioritarie.

**Differenza chiave da BalanceScan:** non analizza il passato —
proietta la traiettoria futura e intercetta i problemi con mesi di anticipo.

**URL produzione:** https://radar.kpilot.tech
**Repository GitHub:** github.com/Fla971/radarfinanziario
**Deploy:** Vercel (Serverless Functions)

---

## 2. Stack tecnologico

| Componente | Dettaglio |
|---|---|
| Frontend | HTML + CSS + JavaScript (single file index.html) |
| Backend | Vercel Serverless Functions (/api/analyze.js) |
| AI | Anthropic Claude API — modello: claude-sonnet-4-20250514 |
| PDF parsing | pdf-parse |
| Excel parsing | SheetJS (xlsx) |
| Dipendenze | @anthropic-ai/sdk, pdf-parse, xlsx, dotenv |

---

## 3. Architettura — Decisioni già prese (NON cambiare)

- **Accetta PDF e Excel (.xlsx, .xls)** — doppio formato, rilevamento automatico
- **Nessun salvataggio file su disco** — tutto in memoria per privacy GDPR
- **Deploy su Vercel Serverless** — non Express classico
- **max_tokens API: 8192** — non ridurre mai
- **Excel → testo strutturato → Claude API** (SheetJS converte, non inviare binario)
- **Dati infrannuali sono provvisori** — il prompt deve sempre segnalarlo

---

## 4. Struttura file

```
/radarfinanziario
  index.html          ← frontend (drag&drop PDF/Excel + visualizzazione)
  /api
    analyze.js        ← Serverless Function principale
  package.json
  vercel.json
  .env                ← ANTHROPIC_API_KEY (mai committare)
  .env.example
```

---

## 5. Variabili d'ambiente richieste

```
ANTHROPIC_API_KEY=sk-ant-...
```

---

## 6. Differenze critiche rispetto al bilancio CEE

I dati infrannuali NON sono affidabili come il bilancio depositato:

| Dato | Affidabilità infrannuale |
|---|---|
| Ricavi YTD | 🟢 Alta |
| Costi operativi | 🟢 Alta |
| Crediti clienti | 🟢 Alta |
| Debiti fornitori | 🟢 Alta |
| Rimanenze magazzino | 🔴 Bassa — non aggiornate infrannualmente |
| Ammortamenti | 🔴 Spesso assenti |
| Risultato netto | 🔴 Inaffidabile senza ammortamenti e imposte |

**Il prompt DEVE segnalare esplicitamente questi limiti.**
**Non fingere precisione dove non c'è.**

---

## 7. Output del report — Sezioni obbligatorie

1. Intestazione (ragione sociale, periodo, tipo documento, data analisi reale)
2. Tabella affidabilità dati (semaforo per ogni area)
3. Situazione attuale (solo dati affidabili)
4. Proiezioni fine anno (3 scenari: ottimistico/base/pessimistico ±10%)
5. Segnali di allerta (max 3, con evidenza numerica)
6. Azioni prioritarie entro 30-60 giorni (max 3)
7. Dati mancanti da richiedere al commercialista
8. Domande per la prossima riunione (3, specifiche sui numeri trovati)
9. Sintesi direzionale (IN LINEA / ATTENZIONE / INTERVENTO URGENTE)
10. **CTA verso Check-Up Margine e Rischi** (sempre inclusa)

---

## 8. Sezione CTA — DA INCLUDERE SEMPRE in fondo al report

```html
<hr>
<h2>Passo Successivo Consigliato</h2>
<p>
Questa analisi proietta la traiettoria di [Ragione Sociale]
sulla base dei dati infrannuali disponibili.
</p>
<p>
Hai già il radar. Ora ti serve la rotta.
</p>
<p>
Per trasformare queste proiezioni in azioni concrete con impatto
misurabile sul margine operativo, il <strong>Check-Up Margine e Rischi</strong>
di Castelli Consulting è il passo naturale: identifica le cause operative
e costruisce una roadmap a 90 giorni.
</p>
<p>📧 flavio@castelliconsulting.it</p>
<hr>
```

---

## 9. Disclaimer obbligatori — sempre visibili in index.html

```
⚠️ I dati infrannuali sono provvisori.
   Le proiezioni hanno valore indicativo
   e non sostituiscono la consulenza professionale.

🔒 Il documento viene elaborato in memoria
   e non viene salvato. I dati non sono condivisi
   né usati per addestrare sistemi AI.
```

---

## 10. Errori da evitare

| Problema | Soluzione |
|---|---|
| Excel con struttura variabile | SheetJS → converti in testo strutturato leggibile, non assumere formato fisso |
| PDF scansionato → testo vuoto | Verificare > 200 chars prima di chiamare API, dare messaggio specifico |
| Proiezioni senza disclaimer | Aggiungere sempre "Proiezione indicativa basata su andamento lineare" |
| Inventare dati mancanti | Se un dato non è presente → segnalarlo esplicitamente, non stimare |
| Data analisi sbagliata | Usare `new Date().toLocaleDateString('it-IT')` sempre |

---

## 11. Sistema Token (in integrazione)

Stesso sistema di BalanceScan — repository `kpilot-token-system`.

Quando integrato, il parametro prodotto da passare è: `"radar"`

```javascript
{ token: tokenDaURL, prodotto: "radar" }
```

---

## 12. Relazione con gli altri prodotti

```
RadarFinanziario™ (questo progetto)
  → secondo livello del funnel (dopo BalanceScan)
  → monitoraggio mensile/trimestrale
  → €39/analisi standalone | €49/mese suite con BalanceScan
  → incluso gratuitamente nel Piano PERFORMANCE di KPilot™

BalanceScan™ (repo: balancescan)
  → analisi bilancio storico — prodotto complementare
  → URL: balancescan.kpilot.tech

KPilot™ (repo: App-KPilot)
  → piattaforma KPI/KRI — destinazione finale del funnel
  → URL: app.kpilot.tech
```

---

## 13. Prossimi task (in ordine di priorità)

- [ ] Completare build iniziale con Claude Code (prompt già pronto)
- [ ] Deploy su Vercel con dominio radar.kpilot.tech
- [ ] Verificare parsing Excel con file reale del commercialista
- [ ] Aggiungere banner disclaimer e privacy in index.html
- [ ] Integrare controllo token (dopo deploy token-system)
- [ ] Test con situazione contabile reale anonimizzata

---

## 14. Contesto business

**Target:** CFO, titolari PMI, commercialisti con clienti manifatturieri
**Problema risolto:** scoprire problemi finanziari a marzo quando sono già accaduti in ottobre
**Pricing:** €39/analisi | €49/mese suite FinanceIntel con BalanceScan
**Suite commercialista:** €199/mese clienti illimitati
**Privacy:** nessun dato salvato, nessun training AI, GDPR compliant

---

*Aggiorna questo file dopo ogni sessione se scopri nuovi errori o decisioni.*
