'use strict';

const BREVO_API = 'https://api.brevo.com/v3';

const MITTENTE = {
  name:  'Flavio Castelli – Castelli Consulting',
  email: 'fcastelli@castelliconsulting.it',
};

function gancioPerEsito(esito) {
  const e = (esito || '').toUpperCase();

  if (e.includes('URGENTE')) {
    return `
      <p>L'analisi segnala una traiettoria che richiede intervento immediato.</p>
      <p>I numeri dicono <em>cosa</em> sta succedendo. Per capire <em>perché</em> — quali processi
      stanno erodendo il margine e come invertire la rotta — il passo successivo è il
      <strong>Check-Up Margine e Rischi</strong>: una diagnosi che quantifica in euro quanto
      puoi recuperare e come, con una roadmap concreta a 90 giorni.</p>
      <p><a href="https://www.castelliconsulting.it/check-up-strategico/"
        style="color:#15803d;font-weight:600;">Scopri come funziona il Check-Up &rarr;</a></p>`;
  }
  if (e.includes('ATTENZIONE')) {
    return `
      <p>L'analisi segnala alcune aree che meritano attenzione nelle prossime settimane.</p>
      <p>I numeri indicano la direzione. Per capire come correggere la rotta in tempo,
      il <strong>Check-Up Margine e Rischi</strong> identifica le cause operative e costruisce
      una roadmap concreta a 90 giorni.</p>
      <p><a href="https://www.castelliconsulting.it/check-up-strategico/"
        style="color:#15803d;font-weight:600;">Scopri il Check-Up &rarr;</a></p>`;
  }
  return `
    <p>L'analisi mostra una traiettoria complessivamente positiva. Buona notizia.</p>
    <p>Anche le aziende in linea con gli obiettivi lasciano spesso margine sul tavolo senza
    accorgersene. Se vuoi capire dove, il <strong>Check-Up Margine e Rischi</strong> misura
    quanto puoi ancora recuperare e come tenerlo monitorato nel tempo.</p>
    <p><a href="https://www.castelliconsulting.it/check-up-strategico/"
      style="color:#15803d;font-weight:600;">Scopri il Check-Up &rarr;</a></p>`;
}

function corpoEmail({ nome, reportCompleto, esito, ragioneSociale }) {
  return `<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<style>
  body{margin:0;padding:0;background:#f1f5f9;}
  .outer{background:#f1f5f9;padding:32px 16px;}
  .card{background:#fff;max-width:640px;margin:0 auto;border-radius:12px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;color:#1e293b;}
  .hdr{background:#1a2744;padding:22px 32px;}
  .hdr-logo{font-size:21px;font-weight:700;color:#fff;letter-spacing:-.5px;}
  .hdr-logo span{color:#4ade80;}
  .hdr-sub{font-size:13px;color:#94a3b8;font-weight:400;margin-left:8px;}
  .body{padding:28px 32px;}
  .intro{font-size:15px;line-height:1.65;color:#475569;margin:0 0 24px;}
  .report-box{border:1px solid #e2e8f0;border-radius:10px;padding:24px 28px;background:#f8fafc;}
  .report-box h2{font-size:15px;font-weight:800;color:#1a2744;border-bottom:2px solid #e2e8f0;padding-bottom:7px;margin:26px 0 11px;text-transform:uppercase;letter-spacing:.3px;}
  .report-box h2:first-child{margin-top:0;}
  .report-box h3{font-size:13px;font-weight:700;color:#243158;margin:16px 0 6px;}
  .report-box p{font-size:14px;line-height:1.65;color:#1e293b;margin:6px 0;}
  .report-box ul{padding-left:20px;margin:7px 0;}
  .report-box li{font-size:14px;line-height:1.7;color:#1e293b;margin-bottom:5px;}
  .report-box table{width:100%;border-collapse:collapse;margin:12px 0;font-size:13px;}
  .report-box th{background:#1a2744;color:#fff;padding:9px 12px;text-align:left;font-weight:600;font-size:12px;}
  .report-box td{border:1px solid #e2e8f0;padding:8px 12px;color:#1e293b;vertical-align:top;}
  .report-box tr:nth-child(even) td{background:#f1f5f9;}
  .report-box hr{border:none;border-top:1px solid #e2e8f0;margin:18px 0;}
  .report-box strong{color:#1a2744;}
  .report-box a{color:#16a34a;}
  .cta{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:20px 24px;margin-top:24px;font-size:14px;line-height:1.65;color:#14532d;}
  .cta a{color:#15803d;font-weight:600;}
  .firma{margin-top:28px;font-size:14px;line-height:1.7;color:#475569;}
  .firma strong{color:#1a2744;}
  .firma a{color:#1a2744;}
  .footer{background:#f8fafc;border-top:1px solid #e2e8f0;padding:16px 32px;font-size:12px;color:#94a3b8;text-align:center;}
  .footer a{color:#64748b;}
</style>
</head>
<body>
<div class="outer"><div class="card">
  <div class="hdr">
    <span class="hdr-logo">Radar<span>Finanziario</span>™</span>
    <span class="hdr-sub">· Castelli Consulting</span>
  </div>
  <div class="body">
    <p class="intro">Ciao <strong>${nome}</strong>,<br>ecco l'analisi infrannuale completa che hai richiesto su RadarFinanziario™.</p>
    <div class="report-box">${reportCompleto}</div>
    <div class="cta">${gancioPerEsito(esito)}</div>
    <p class="firma">A presto,<br><strong>Flavio Castelli</strong><br>
    Castelli Consulting<br>
    <a href="mailto:fcastelli@castelliconsulting.it">fcastelli@castelliconsulting.it</a></p>
  </div>
  <div class="footer">
    Hai ricevuto questa email perché hai richiesto un'analisi tramite RadarFinanziario™.<br>
    Per qualsiasi domanda: <a href="mailto:fcastelli@castelliconsulting.it">fcastelli@castelliconsulting.it</a>
  </div>
</div></div>
</body></html>`;
}

async function inviaEmailReport({ nome, email, reportCompleto, esito, ragioneSociale }) {
  const subject = ragioneSociale
    ? `La tua analisi infrannuale di ${ragioneSociale} è pronta`
    : `La tua analisi RadarFinanziario è pronta`;

  const r = await fetch(`${BREVO_API}/smtp/email`, {
    method: 'POST',
    headers: {
      'api-key': process.env.BREVO_API_KEY,
      'Content-Type': 'application/json',
      accept: 'application/json',
    },
    body: JSON.stringify({
      sender:      MITTENTE,
      to:          [{ email, name: nome }],
      subject,
      htmlContent: corpoEmail({ nome, reportCompleto, esito, ragioneSociale }),
    }),
  });

  if (!r.ok) throw new Error('Invio email Brevo fallito: ' + (await r.text()));
}

async function iscriviContattoBrevo({ nome, email, ragioneSociale, esito }) {
  const r = await fetch(`${BREVO_API}/contacts`, {
    method: 'POST',
    headers: {
      'api-key': process.env.BREVO_API_KEY,
      'Content-Type': 'application/json',
      accept: 'application/json',
    },
    body: JSON.stringify({
      email,
      attributes:    { NOME: nome, AZIENDA: ragioneSociale || '', ESITO: esito || '' },
      listIds:       [Number(process.env.BREVO_LISTA_LEAD_RADAR)],
      updateEnabled: true,
    }),
  });

  if (!r.ok) throw new Error('Iscrizione contatto Brevo fallita: ' + (await r.text()));
}

module.exports = { inviaEmailReport, iscriviContattoBrevo };
