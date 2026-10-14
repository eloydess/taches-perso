

/* =====================================================================
 * Annuler une erreur (v2.8 — 14/10/2026)
 *  - « Clôturés récemment » : dossiers clôturés du suivi et dossiers archivés depuis moins de 45 jours.
 *  - Rouvrir un dossier : remettre « en attente de livraison » (livraison → 🔴) ou rouvrir le paiement
 *    (date de renvoi de la preuve effacée). Un dossier archivé est d'abord ressorti des archives.
 *  - Remettre un mail rangé dans la boîte de réception (l'étiquette du laboratoire est retirée,
 *    le statut lu / non lu n'est pas modifié) ; il n'est plus rangé automatiquement ensuite.
 * ===================================================================== */
var RO = { JOURS: 45, PREMIERE_ARCH: 3, EN_ATTENTE: '🔴 En attente de livraison' };

function ro_jour_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, APP.FUSEAU, 'yyyy-MM-dd');
  var m = /(\d{2})\/(\d{2})\/(\d{4})/.exec(String(d || ''));
  return m ? m[3] + '-' + m[2] + '-' + m[1] : '';
}

/** Dossiers clôturés (suivi) et archivés récemment, les plus récents d'abord. */
function ro_clotures_() {
  var res = [];
  var su = feuille_(RJ.F_SUIVI), n = su.getLastRow() - RJ.PREMIERE + 1;
  if (n > 0) {
    su.getRange(RJ.PREMIERE, 1, n, 16).getDisplayValues().forEach(function (l) {
      if (!l[2] || !(/CL[ÔO]TUR/i.test(l[11]) || l[10])) return;
      res.push({ ou: 'suivi', cle: 'RJ|' + l[2] + '|' + l[3], tiers: l[2], refs: l[3], montant: tr_nombre_(l[6]), preuve: l[10], livraison: l[13], tri: ro_jour_(l[10]) });
    });
  }
  var ar = feuille_(RJ.F_ARCH), m = ar.getLastRow() - RO.PREMIERE_ARCH + 1;
  if (m > 0) {
    var limite = Utilities.formatDate(new Date(Date.now() - RO.JOURS * 86400000), APP.FUSEAU, 'yyyy-MM-dd');
    ar.getRange(RO.PREMIERE_ARCH, 1, m, 16).getDisplayValues().forEach(function (l, i) {
      if (!l[2]) return;
      var archive = ro_jour_(l[15]);
      if (!archive || archive < limite) return;
      res.push({ ou: 'archives', ligne: RO.PREMIERE_ARCH + i, cle: 'RJ|' + l[2] + '|' + l[3], tiers: l[2], refs: l[3], montant: tr_nombre_(l[6]),
        preuve: l[10], livraison: l[13], archive: l[15], tri: ro_jour_(l[10]) || archive });
    });
  }
  res.sort(function (a, b) { return a.tri < b.tri ? 1 : a.tri > b.tri ? -1 : 0; });
  return res.slice(0, 60);
}

function mob_clotures(jeton) {
  mob_verifier_(jeton);
  return ro_clotures_();
}

/** Ressort une ligne des archives vers le suivi (avec ses pièces si elles ont été conservées). Renvoie la ligne du suivi. */
function ro_desarchiver_(ligneArch, cle) {
  var ar = feuille_(RJ.F_ARCH), su = feuille_(RJ.F_SUIVI);
  var v = ar.getRange(ligneArch, 1, 1, 18).getValues()[0];
  var aff = ar.getRange(ligneArch, 1, 1, 18).getDisplayValues()[0];
  if ('RJ|' + aff[2] + '|' + aff[3] !== cle) throw new Error('Les archives ont changé entre-temps : rechargez l\'écran.');
  // première ligne libre après le dernier tiers renseigné
  var n = Math.max(su.getLastRow() - RJ.PREMIERE + 1, 0);
  var cs = n > 0 ? su.getRange(RJ.PREMIERE, 3, n, 1).getDisplayValues() : [];
  var r = RJ.PREMIERE;
  for (var i = cs.length - 1; i >= 0; i--) { if (cs[i][0]) { r = RJ.PREMIERE + i + 1; break; } }
  su.getRange(r, 2, 1, 10).setValues([v.slice(1, 11)]);
  su.getRange(r, 13).setValue(v[12]);
  su.getRange(r, 14).setValue(v[13]);
  [5, 6, 9, 10, 11].forEach(function (c) { su.getRange(r, c).setNumberFormat('dd/mm/yyyy'); });
  su.getRange(r, 7).setNumberFormat('#,##0.00');
  if (aff[16] || aff[17]) su.getRange(r, 19, 1, 2).setNumberFormat('@').setValues([[aff[16], aff[17]]]);
  rj_formules_(su, r, r);
  ar.deleteRow(ligneArch);
  try { vs_majSuivi_(); } catch (e) { /* colonnes versé / reste recalculées à la prochaine saisie */ }
  return r;
}

/**
 * Rouvre un dossier clôturé par erreur.
 * mode 'livraison' : la commande n'est pas livrée → livraison « 🔴 En attente de livraison » (le dossier reste payé).
 * mode 'paiement'  : le paiement n'est pas terminé → date de renvoi de la preuve effacée (le dossier revient dans les paiements).
 */
function mob_rouvrir(jeton, ou, cle, ligneArch, mode) {
  mob_verifier_(jeton);
  if (['livraison', 'paiement'].indexOf(mode) < 0) throw new Error('Action inconnue.');
  var verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    var su = feuille_(RJ.F_SUIVI), r;
    if (ou === 'archives') {
      r = ro_desarchiver_(Number(ligneArch), cle);
    } else {
      r = pv_ligne_(cle).ligne;
    }
    var avant = su.getRange(r, 11).getDisplayValue() + ' / ' + su.getRange(r, 14).getDisplayValue();
    if (mode === 'livraison') su.getRange(r, 14).setValue(RO.EN_ATTENTE);
    else su.getRange(r, 11).setValue('');
    SpreadsheetApp.flush();
    journal_('Dossier rouvert (' + (mode === 'livraison' ? 'livraison en attente' : 'paiement') + (ou === 'archives' ? ', ressorti des archives' : '') + ')', '', cle + ' — avant : ' + avant);
    return { ok: true, statut: su.getRange(r, 12).getDisplayValue() };
  } finally {
    verrou.releaseLock();
  }
}

/** Remet un mail rangé dans la boîte de réception (retire l'étiquette du laboratoire). */
function mob_desranger(jeton, id) {
  mob_verifier_(jeton);
  if (/^C-/.test(String(id))) throw new Error('Un courrier scanné n\'est pas un mail.');
  var verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    var f = mob_ligneFiche_(id);
    var etiq = f.h['Étiquette finale'] ? f.sh.getRange(f.ligne, f.h['Étiquette finale']).getDisplayValue() : '';
    var lab = rg_etiquettes_().filter(function (x) { return x.nom === etiq; })[0];
    mob_gmailApi_('post', 'threads/' + id + '/modify', { addLabelIds: ['INBOX'], removeLabelIds: lab ? [lab.id] : [] });
    var note = '↩ remis en boîte de réception le ' + maintenant_();
    if (f.h['Date application']) ecrireTexte_(f.sh, f.ligne, f.h['Date application'], note);   // empêche le rangement automatique de le reprendre
    journal_('Mail remis en boîte de réception' + (etiq ? ' (étiquette « ' + etiq + ' » retirée)' : ''), id, '');
    return { ok: true, note: note };
  } finally {
    verrou.releaseLock();
  }
}
