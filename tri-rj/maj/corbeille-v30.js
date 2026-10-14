

/* =====================================================================
 * Supprimer un mail sans le ranger (v3.0 — 14/10/2026)
 * « 🗑 Supprimer » met le fil dans la Corbeille de Gmail (récupérable pendant 30 jours, Gmail le vide ensuite).
 * Aucune suppression définitive n'est faite par l'application. « ↩ Restaurer » le remet dans la boîte de réception.
 * ===================================================================== */
function mob_corbeille(jeton, id) {
  mob_verifier_(jeton);
  if (/^C-/.test(String(id))) throw new Error('Un courrier scanné n\'est pas un mail.');
  var verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    var f = mob_ligneFiche_(id);
    mob_gmailApi_('post', 'threads/' + id + '/trash');
    var note = '🗑 mis à la corbeille le ' + maintenant_();
    if (f.h['Date application']) ecrireTexte_(f.sh, f.ligne, f.h['Date application'], note);
    journal_('Mail mis à la corbeille (récupérable 30 jours)', id, '');
    return { ok: true, note: note };
  } finally {
    verrou.releaseLock();
  }
}

function mob_restaurer(jeton, id) {
  mob_verifier_(jeton);
  var verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    var f = mob_ligneFiche_(id);
    mob_gmailApi_('post', 'threads/' + id + '/untrash');
    mob_gmailApi_('post', 'threads/' + id + '/modify', { addLabelIds: ['INBOX'], removeLabelIds: [] });
    var note = '↩ remis en boîte de réception le ' + maintenant_();
    if (f.h['Date application']) ecrireTexte_(f.sh, f.ligne, f.h['Date application'], note);
    journal_('Mail sorti de la corbeille et remis en boîte de réception', id, '');
    return { ok: true, note: note };
  } finally {
    verrou.releaseLock();
  }
}
