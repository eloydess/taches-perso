

/* =====================================================================
 * Livraison et soldes (v2.7 — 14/10/2026)
 * 1. Un dossier payé et clôturé reste visible (« 📦 Payés — en attente de livraison ») et n'est pas archivé
 *    tant que la livraison n'est pas notée « 🟢 Livrée en totalité » ou « ⚪ Sans livraison attendue ».
 *    Ses mails, eux, sont rangés dès que le virement est fait et la preuve envoyée (rangement Gmail).
 * 2. Le solde de chaque banque est actualisé : solde saisi − virements effectués depuis cette saisie
 *    (versements enregistrés, ou dossiers validés par l'administrateur / clôturés sans versement enregistré).
 * ===================================================================== */
var LS = { FINI: /^(🟢|⚪)/ };

function ls_livre_(livraison) { return LS.FINI.test(String(livraison || '')); }

/** Dossiers payés (clôturés) dont la livraison n'est pas encore notée comme terminée. */
function ls_attenteLivraison_() {
  var sh = feuille_(RJ.F_SUIVI), n = sh.getLastRow() - RJ.PREMIERE + 1;
  if (n < 1) return [];
  var res = [];
  sh.getRange(RJ.PREMIERE, 1, n, 16).getDisplayValues().forEach(function (l) {
    if (!l[2] || !/CL[ÔO]TUR/i.test(l[11]) || ls_livre_(l[13])) return;
    res.push({ cle: 'RJ|' + l[2] + '|' + l[3], tiers: l[2], refs: l[3], montant: tr_nombre_(l[6]), preuve: l[10], livraison: l[13], lien: l[15] });
  });
  return res;
}

function ls_jour_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, APP.FUSEAU, 'yyyy-MM-dd');
  var m = /(\d{2})\/(\d{2})\/(\d{4})/.exec(String(d || ''));
  return m ? m[3] + '-' + m[2] + '-' + m[1] : '';
}
function ls_instant_(t) {
  var m = /(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/.exec(String(t || ''));
  return m ? m[3] + '-' + m[2] + '-' + m[1] + ' ' + ('0' + (m[4] || '0')).slice(-2) + ':' + (m[5] || '00') : '';
}

/**
 * Virements effectués depuis la dernière saisie du solde de chaque banque.
 * - Versement enregistré (onglet « Versements ») saisi après le solde : son montant.
 * - Sinon, dossier validé par l'administrateur ou clôturé le jour du solde ou après : montant choisi (plan) ou montant TTC.
 */
function ls_virementsDepuis_(soldes) {
  var res = {};
  TR.BANQUES.forEach(function (b) { res[b] = { total: 0, liste: [], sansMontant: [] }; });
  var plan = {};
  var fp = tr_feuille_(TR.F_PLAN, TR.ENT_PLAN);
  if (fp.getLastRow() > 1) fp.getRange(2, 1, fp.getLastRow() - 1, 4).getDisplayValues().forEach(function (l) { plan[l[0]] = { montant: tr_nombre_(l[2]), banque: l[3] }; });
  // Versements enregistrés
  var vsCles = {};
  var fv = vs_feuille_();
  if (fv.getLastRow() > 1) {
    fv.getRange(2, 1, fv.getLastRow() - 1, VS.ENT.length).getDisplayValues().forEach(function (l) {
      if (!l[0]) return;
      vsCles[l[0]] = true;
      var b = tr_banque_(l[5]), s = soldes[b];
      if (!s || !s.date) return;
      if (ls_instant_(l[7]) <= ls_instant_(s.date)) return;
      var m = tr_nombre_(l[4]);
      if (m == null) return;
      res[b].total = Math.round((res[b].total + m) * 100) / 100;
      res[b].liste.push({ tiers: l[1], refs: l[2], montant: m, quand: l[7], origine: 'versement enregistré' });
    });
  }
  // Dossiers validés / clôturés sans versement enregistré (Suivi + Archives)
  var lire = function (nom, premiere) {
    var sh = feuille_(nom), n = sh.getLastRow() - premiere + 1;
    if (n < 1) return;
    var v = sh.getRange(premiere, 1, n, 14).getValues(), aff = sh.getRange(premiere, 1, n, 14).getDisplayValues();
    v.forEach(function (l, i) {
      var a = aff[i];
      if (!a[2]) return;
      var cle = 'RJ|' + a[2] + '|' + a[3];
      if (vsCles[cle]) return;
      var jour = ls_jour_(l[9]) || ls_jour_(l[10]);
      if (ls_jour_(l[9]) && ls_jour_(l[10]) && ls_jour_(l[10]) < ls_jour_(l[9])) jour = ls_jour_(l[10]);
      if (!jour || jour > Utilities.formatDate(new Date(), APP.FUSEAU, 'yyyy-MM-dd')) return;   // date absente ou future (erreur de saisie)
      var p = plan[cle] || {};
      var b = TR.BANQUES.indexOf(p.banque) >= 0 ? p.banque : (tr_banque_(a[7]) || 'CEBPL');
      var s = soldes[b];
      if (!s || !s.date || jour < ls_jour_(s.date)) return;
      var m = p.montant != null && p.montant > 0 ? p.montant : tr_nombre_(a[6]);
      var quoi = { tiers: a[2], refs: a[3], quand: a[9] || a[10], origine: a[9] ? 'validé par l\'administrateur' : 'preuve envoyée' };
      if (m == null || m <= 0) { res[b].sansMontant.push(quoi); return; }
      quoi.montant = m;
      res[b].total = Math.round((res[b].total + m) * 100) / 100;
      res[b].liste.push(quoi);
    });
  };
  lire(RJ.F_SUIVI, RJ.PREMIERE);
  try { lire(RJ.F_ARCH, 3); } catch (e) { /* pas d'archives */ }
  return res;
}

/** Soldes saisis + virements depuis + solde estimé actuel. */
function ls_soldes_() {
  var s = tr_soldes_();
  var d = ls_virementsDepuis_(s);
  Object.keys(s).forEach(function (b) {
    var x = d[b] || { total: 0, liste: [], sansMontant: [] };
    s[b].saisi = s[b].montant;
    s[b].depuis = x;
    if (s[b].montant != null) s[b].actuel = Math.round((s[b].montant - x.total) * 100) / 100;
  });
  return s;
}

function ls_soldesSur_() { try { return ls_soldes_(); } catch (e) { return tr_soldes_(); } }
function ls_attenteSur_() { try { return ls_attenteLivraison_(); } catch (e) { return []; } }

/**
 * Texte du mail « preuve de virement » adapté à l'interlocuteur (v2.7) :
 *  - loyer / bailleur (Carmila, Immobilière Carrefour…) : règlement du loyer, pas de commande ;
 *  - contrat de location, leasing, organisme (greffe, URSSAF…) : règlement, pas de commande ;
 *  - fournisseur déjà livré : règlement de la facture ;
 *  - fournisseur avec commande en attente (proforma) : demande d'expédition de la commande.
 */
function ls_genre_(l) {
  var t = (l[1] + ' ' + l[2] + ' ' + l[3]).toUpperCase();
  if (/LOYER|BAIL|[ÉE]CH[ÉE]ANCE \d|CARMILA|IMMOBILI|CARREFOUR PROPERTY/.test(t)) return 'loyer';
  if (/CONTRAT|LEASING|LEASECOM|GRENKE|LOCATION|CR[ÉE]DIT-BAIL|SIEMENS|ANAVEO/.test(t)) return 'contrat';
  if (/^ORGANISME|^TRAVAUX|GREFFE|URSSAF|TRIBUNAL|IMP[ÔO]T|TR[ÉE]SOR|CAVP|CCI/.test(t)) return 'organisme';
  if (/^🟢/.test(l[13])) return 'livre';
  if (/^⚪/.test(l[13])) return 'organisme';
  return 'commande';
}
function ls_texteEnvoi_(l, vs, montant, refs) {
  var genre = ls_genre_(l), total = tr_nombre_(l[6]);
  var de = montant != null ? ' de ' + pv_eur_(montant) : '';
  var partiel = vs && total != null && (vs.liste.length > 1 || Math.round((total - vs.total) * 100) > 0);
  var recap = partiel ? '\n\nMontant total : ' + pv_eur_(total) + ' ; déjà versé : ' + pv_eur_(vs.total) + ' ; reste dû : ' + pv_eur_(Math.max(0, total - vs.total)) + '.' : '';
  var corps;
  if (genre === 'loyer') {
    corps = 'Veuillez trouver ci-joint la preuve du virement' + de + ' effectué en règlement du loyer — ' + refs + '.' + recap +
      '\n\nPourriez-vous nous confirmer la bonne réception de ce règlement ?';
  } else if (genre === 'contrat' || genre === 'organisme') {
    corps = 'Veuillez trouver ci-joint la preuve du virement' + de + ' effectué en règlement de : ' + refs + '.' + recap +
      '\n\nPourriez-vous nous confirmer la bonne réception de ce règlement ?';
  } else if (genre === 'livre') {
    corps = 'Veuillez trouver ci-joint la preuve du virement' + de + ' effectué en règlement de : ' + refs + '.' + recap +
      '\n\nNous vous remercions de bien vouloir l\'affecter à notre compte.';
  } else {
    corps = 'Veuillez trouver ci-joint la preuve du virement' + de + ' pour ' + refs + '.' + recap +
      (partiel && Math.round((total - vs.total) * 100) > 0 ? '' : '\n\nMerci de bien vouloir procéder à l\'expédition de la commande.');
  }
  return { genre: genre, texte: 'Bonjour,\n\n' + corps + '\n\nBien cordialement.' };
}
