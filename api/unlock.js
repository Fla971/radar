'use strict';

require('dotenv').config();

const { createClient } = require('@supabase/supabase-js');
const { sbloccaReport } = require('./leadgate');
const { inviaEmailReport, iscriviContattoBrevo } = require('./brevo');

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Metodo non consentito' });
  }

  let body = {};
  try {
    const chunks = [];
    for await (const chunk of req) chunks.push(chunk);
    body = JSON.parse(Buffer.concat(chunks).toString());
  } catch {
    return res.status(400).json({ ok: false, motivo: 'Body non valido.' });
  }

  const { analisiId, nome, email, consenso } = body;
  const ip = req.headers['x-forwarded-for'] || req.socket?.remoteAddress;

  const supabase = createClient(
    process.env.SUPABASE_URL,
    process.env.SUPABASE_SERVICE_ROLE_KEY
  );

  const r = await sbloccaReport(supabase, {
    analisiId,
    nome,
    email,
    consenso: consenso === true,
    ip,
  });

  if (!r.ok) return res.status(400).json({ ok: false, motivo: r.motivo });

  try {
    await inviaEmailReport({
      nome,
      email,
      reportCompleto: r.reportCompleto,
      esito:          r.esito,
      ragioneSociale: r.ragioneSociale,
    });
    await iscriviContattoBrevo({
      nome,
      email,
      ragioneSociale: r.ragioneSociale,
      esito:          r.esito,
    });
  } catch (e) {
    console.error('Brevo unlock Radar:', e.message);
    // Non bloccare l'utente se Brevo fallisce
  }

  return res.json({ ok: true });
};
