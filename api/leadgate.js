'use strict';

function spezzaReport(html) {
  const idx = html.indexOf('<!--GATE-->');
  if (idx === -1) {
    // Fallback: mostra primo 35% se manca il marker
    return { anteprima: html.slice(0, Math.floor(html.length * 0.35)) };
  }
  return { anteprima: html.slice(0, idx) };
}

async function salvaAnalisiTemp(supabase, { reportCompleto, ragioneSociale, esito }) {
  const { data, error } = await supabase
    .from('analisi_temp')
    .insert({
      report_completo: reportCompleto,
      ragione_sociale: ragioneSociale || null,
      esito:           esito || null,
    })
    .select('id')
    .single();

  if (error) throw new Error('Salvataggio analisi fallito: ' + error.message);
  return data.id;
}

async function sbloccaReport(supabase, { analisiId, nome, email, consenso, ip }) {
  if (!analisiId)          return { ok: false, motivo: 'ID analisi mancante.' };
  if (!email || !nome)     return { ok: false, motivo: 'Nome e email obbligatori.' };
  if (!consenso)           return { ok: false, motivo: 'È necessario il consenso per procedere.' };

  const { data, error } = await supabase
    .from('analisi_temp')
    .select('report_completo, ragione_sociale, esito, usato')
    .eq('id', analisiId)
    .maybeSingle();

  if (error || !data) return { ok: false, motivo: 'Analisi non trovata o scaduta.' };
  if (data.usato)     return { ok: false, motivo: 'Questo report è già stato sbloccato.' };

  await supabase
    .from('analisi_temp')
    .update({ usato: true, email_unlock: email, nome_unlock: nome, ip_unlock: ip || null })
    .eq('id', analisiId);

  return {
    ok:             true,
    reportCompleto: data.report_completo,
    ragioneSociale: data.ragione_sociale,
    esito:          data.esito,
  };
}

module.exports = { spezzaReport, salvaAnalisiTemp, sbloccaReport };
