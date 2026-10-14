

/* =====================================================================
 * Rangement des mails des dossiers PAYÉS (v2.9 — 14/10/2026)
 * Dès que le virement est fait et la preuve envoyée (dossier clôturé), tous les mails de la boîte de réception
 * qui concernent ce dossier sont rangés dans l'étiquette du laboratoire, sortis de la boîte de réception
 * et marqués comme lus (demande de Pierre du 14/10/2026). Le dossier, lui, reste au suivi jusqu'à la livraison.
 * Mails retrouvés : (1) le mail d'origine du dossier (colonne P), (2) les fils de la boîte de réception dont
 * l'objet ou l'extrait cite un numéro de facture / proforma / commande du dossier (jamais un n° de client ou
 * de compte), (3) les fiches de la File de tri rattachées au dossier. Rien n'est supprimé.
 * ===================================================================== */
var RP = { JOURS_ARCH: 60, MAX_S: 70 };

/** Numéros forts d'un dossier : au moins 6 caractères dont 4 chiffres, hors n° de client / compte / code. */
function rp_numeros_(refs) {
  var res = [];
  var propre = String(refs || '').replace(/(compte\s+client|code\s+client|r[ée]f[ée]rence\s+client|n°\s*client|client|compte|code|cip|ship-to|siren|siret)\s*[:n°]*\s*[A-Z0-9][A-Z0-9\-]*(\s*[–-]\s*[A-Z0-9]+)?/gi, ' ');
  propre.split(/[;\/,\n]| et /i).forEach(function (seg) {
    (seg.match(/[A-Z0-9][A-Z0-9\-]{5,}/gi) || []).forEach(function (t) {
      t = t.replace(/^-+|-+$/g, '');
      if (/^20\d{2}(0[1-9]|1[0-2])\d{2}$/.test(t)) return;   // une date (aaaammjj), pas un numéro de pièce
      if (t.length >= 6 && (t.match(/\d/g) || []).length >= 4 && res.indexOf(t) < 0) res.push(t);
    });
  });
  return res;
}

function rp_filDepuisLien_(lien) {
  var m = /thread-f(?::|%3A)(\d+)/i.exec(String(lien || ''));
  return m ? pc_decHex_(m[1]) : '';
}

/** Étiquette du dossier : celle d'une fiche rattachée si elle existe dans Gmail, sinon devinée parmi les étiquettes existantes. */
function rp_etiquette_(tiers, proposee) {
  var noms = rg_etiquettes_().map(function (e) { return e.nom; });
  if (proposee && noms.indexOf(proposee) >= 0) return proposee;
  var e = rg_devine_(tiers);
  if (e) return e;
  return '';   // aucune étiquette existante : pas de création automatique, Pierre choisit avec « Ranger »
}

function rp_dossiersPayes_() {
  var res = [];
  var su = feuille_(RJ.F_SUIVI), n = su.getLastRow() - RJ.PREMIERE + 1;
  if (n > 0) su.getRange(RJ.PREMIERE, 1, n, 16).getDisplayValues().forEach(function (l) {
    if (l[2] && (/CL[ÔO]TUR/i.test(l[11]) || l[10])) res.push({ tiers: l[2], refs: l[3], lien: l[15] });
  });
  try {
    var ar = feuille_(RJ.F_ARCH), m = ar.getLastRow() - 2;
    var limite = Utilities.formatDate(new Date(Date.now() - RP.JOURS_ARCH * 86400000), APP.FUSEAU, 'yyyy-MM-dd');
    if (m > 0) ar.getRange(3, 1, m, 16).getDisplayValues().forEach(function (l) {
      var mm = /(\d{2})\/(\d{2})\/(\d{4})/.exec(l[10] || l[15]);
      if (l[2] && mm && (mm[3] + '-' + mm[2] + '-' + mm[1]) >= limite) res.push({ tiers: l[2], refs: l[3], lien: '' });
    });
  } catch (e) { /* pas d'archives */ }
  return res;
}

/** Fil en boîte de réception ? + objet et extraits (pour vérifier que le numéro y figure vraiment). */
function rp_fil_(id) {
  try {
    var t = mob_gmailApi_('get', 'threads/' + id + '?format=metadata&metadataHeaders=Subject');
    var dansReception = false, texte = '';
    (t.messages || []).forEach(function (m) {
      if ((m.labelIds || []).indexOf('INBOX') >= 0) dansReception = true;
      texte += ' ' + (m.snippet || '');
      ((m.payload && m.payload.headers) || []).forEach(function (h) { if (/^subject$/i.test(h.name)) texte += ' ' + h.value; });
    });
    return { reception: dansReception, texte: texte };
  } catch (e) { return { reception: false, texte: '' }; }
}

function rg_rangerPayes_() {
  var debut = Date.now(), nb = 0, vus = {};
  var tri = feuille_(APP.F_TRI), h = entetes_(tri), mt = tri.getLastRow() - 1;
  var fiches = mt > 0 ? tri.getRange(2, 1, mt, tri.getLastColumn()).getDisplayValues() : [];
  var g = function (l, nom) { return h[nom] ? l[h[nom] - 1] : ''; };
  var ligneFiche = {};
  fiches.forEach(function (l, i) { if (g(l, 'ID')) ligneFiche[g(l, 'ID')] = i; });
  var ranger = function (id, d, etiq, pourquoi) {
    if (!id || vus[id] || Date.now() - debut > RP.MAX_S * 1000) return;
    vus[id] = true;
    var i = ligneFiche[id];
    if (i !== undefined && /^↩/.test(g(fiches[i], 'Date application'))) return;   // remis en réception volontairement par Pierre
    var f = rp_fil_(id);
    if (!f.reception) return;
    if (pourquoi === 'numero' && !d.numeros.some(function (t) { return f.texte.indexOf(t) >= 0; })) return;
    try {
      rg_rangerFil_(id, etiq);
      if (i !== undefined) rg_noterFiche_(tri, h, i + 2, etiq);
      journal_('Dossier payé : mail rangé dans « ' + etiq + ' » et marqué comme lu', id, d.tiers + ' | ' + d.refs);
      nb++;
    } catch (e) { /* réessai à la prochaine ouverture */ }
  };
  // un numéro partagé par plusieurs dossiers ou par une fiche encore à valider est un n° de compte, pas de pièce : écarté
  var compte = {};
  var su = feuille_(RJ.F_SUIVI), ns = su.getLastRow() - RJ.PREMIERE + 1;
  (ns > 0 ? su.getRange(RJ.PREMIERE, 4, ns, 1).getDisplayValues() : []).forEach(function (r) {
    rp_numeros_(r[0]).forEach(function (t) { compte[t] = (compte[t] || 0) + 1; });
  });
  var aValider = fiches.filter(function (l) { return g(l, 'Statut') === 'À valider'; }).map(function (l) { return g(l, 'Références') + ' ' + g(l, 'Objet'); }).join(' ');
  rp_dossiersPayes_().forEach(function (d) {
    if (Date.now() - debut > RP.MAX_S * 1000) return;
    d.numeros = rp_numeros_(d.refs).filter(function (t) { return (compte[t] || 0) <= 1; });
    // fiches rattachées (même lien, ou mêmes tiers + références, ou un numéro du dossier dans les références de la fiche)
    var rattachees = fiches.filter(function (l) {
      var id = g(l, 'ID');
      if (!id || /^C-/.test(id) || g(l, 'Statut') === 'À valider') return false;
      if (d.lien && rp_filDepuisLien_(d.lien) === id) return true;
      if (g(l, 'Références') === d.refs && rg_norm_(g(l, 'Tiers')) === rg_norm_(d.tiers)) return true;
      var r = g(l, 'Références') + ' ' + g(l, 'Objet');
      return d.numeros.some(function (t) { return r.indexOf(t) >= 0; });
    });
    var proposee = rattachees.map(function (l) { return g(l, 'Étiquette finale') || g(l, 'Étiquette proposée'); }).filter(String)[0];
    var etiq = rp_etiquette_(d.tiers, proposee);
    if (!etiq) return;   // pas d'étiquette sûre : Pierre rangera avec « Ranger »
    ranger(rp_filDepuisLien_(d.lien), d, etiq, 'lien');
    rattachees.forEach(function (l) { ranger(g(l, 'ID'), d, etiq, 'fiche'); });
    d.numeros.forEach(function (t) {
      if (Date.now() - debut > RP.MAX_S * 1000) return;
      var r = mob_gmailApi_('get', 'threads?maxResults=10&q=' + encodeURIComponent('in:inbox "' + t + '"'));
      (r.threads || []).forEach(function (x) { ranger(x.id, d, etiq, 'numero'); });
    });
  });
  return nb;
}
