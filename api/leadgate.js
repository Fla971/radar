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
    .select('*')
    .eq('id', analisiId)
    .maybeSingle();

  if (error) {
    console.error('sbloccaReport SELECT error:', error.message);
    return { ok: false, motivo: 'Errore nel recupero dell\'analisi: ' + error.message };
  }
  if (!data) return { ok: false, motivo: 'Analisi non trovata o scaduta.' };

  // Supporta sia colonna "usato" che "utilizzato" o assenza della colonna
  if (data.usato === true || data.utilizzato === true) {
    return { ok: false, motivo: 'Questo report è già stato sbloccato.' };
  }

  // Aggiorna solo le colonne che sicuramente esistono
  const aggiornamenti = {};
  if ('usato'        in data) aggiornamenti.usato        = true;
  if ('utilizzato'   in data) aggiornamenti.utilizzato   = true;
  if ('email_unlock' in data) aggiornamenti.email_unlock = email;
  if ('nome_unlock'  in data) aggiornamenti.nome_unlock  = nome;
  if ('ip_unlock'    in data) aggiornamenti.ip_unlock    = ip || null;

  if (Object.keys(aggiornamenti).length > 0) {
    await supabase.from('analisi_temp').update(aggiornamenti).eq('id', analisiId);
  }

  return {
    ok:             true,
    reportCompleto: data.report_completo,
    ragioneSociale: data.ragione_sociale,
    esito:          data.esito,
  };
}

module.exports = { spezzaReport, salvaAnalisiTemp, sbloccaReport };
