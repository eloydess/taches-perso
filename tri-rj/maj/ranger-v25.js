

/* =====================================================================
 * Rangement des mails (v2.5 — 14/10/2026)
 * Un mail est rangé dans l'étiquette de son laboratoire (LABORATOIRES/…, PRESTATAIRES/…) et sort de la boîte de réception :
 *   - automatiquement quand son dossier de paiement est clôturé (preuve renvoyée au tiers) ;
 *   - ou quand Pierre appuie sur « Ranger » sur la fiche (informations, publicités…).
 * Le statut lu / non lu n'est JAMAIS modifié (seules les étiquettes changent). Rien n'est supprimé.
 * ===================================================================== */
var RG = { PREFIXES: ['LABORATOIRES/', 'PRESTATAIRES/'] };

function rg_etiquettes_() {
  var c = CacheService.getScriptCache(), v = c.get('RG_ETIQ');
  if (v) return JSON.parse(v);
  var l = (mob_gmailApi_('get', 'labels').labels || []).filter(function (x) { return x.type === 'user'; })
    .map(function (x) { return { id: x.id, nom: x.name }; });
  try { c.put('RG_ETIQ', JSON.stringify(l), 600); } catch (e) { /* trop gros pour le cache */ }
  return l;
}

function rg_norm_(t) { return String(t || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Z0-9]/g, ''); }

/** Étiquette de laboratoire la plus probable pour un tiers (ou ''). */
function rg_devine_(tiers) {
  var n = rg_norm_(String(tiers || '').replace(/\(.*?\)/g, ' ').split(/ via | \/ |—/i)[0]);
  if (n.length < 3) return '';
  var meilleur = '', score = 0;
  rg_etiquettes_().forEach(function (e) {
    if (!RG.PREFIXES.some(function (p) { return e.nom.indexOf(p) === 0; })) return;
    var fin = rg_norm_(e.nom.split('/').slice(1).join(' '));
    if (!fin) return;
    var s = fin === n ? 3 : (n.indexOf(fin) === 0 || fin.indexOf(n) === 0) ? 2 : (n.indexOf(fin) >= 0 || fin.indexOf(n) >= 0) ? 1 : 0;
    if (fin.length < 4 && s < 3) s = 0;
    if (s > score || (s === score && s > 0 && e.nom.length < meilleur.length)) { score = s; meilleur = e.nom; }
  });
  return meilleur;
}

function rg_idEtiquette_(nom) {
  nom = String(nom || '').trim();
  if (!nom) throw new Error('Choisissez une étiquette.');
  var e = rg_etiquettes_().filter(function (x) { return x.nom === nom; })[0];
  if (e) return e.id;
  var cree = mob_gmailApi_('post', 'labels', { name: nom, labelListVisibility: 'labelShow', messageListVisibility: 'show' });
  CacheService.getScriptCache().remove('RG_ETIQ');
  journal_('Étiquette créée : ' + nom, '', '');
  return cree.id;
}

/** Range un fil : ajoute l'étiquette, retire « Boîte de réception ». Ne touche pas au statut lu / non lu. */
function rg_rangerFil_(id, etiquette) {
  var lab = rg_idEtiquette_(etiquette);
  mob_gmailApi_('post', 'threads/' + id + '/modify', { addLabelIds: [lab], removeLabelIds: ['INBOX'] });
}

function rg_noterFiche_(sh, h, ligne, etiquette) {
  if (h['Étiquette finale']) ecrireTexte_(sh, ligne, h['Étiquette finale'], etiquette);
  if (h['Date application']) ecrireTexte_(sh, ligne, h['Date application'], maintenant_());
}

/** Liste des étiquettes de la boîte (pour choisir dans l'application). */
function mob_etiquettes(jeton) {
  mob_verifier_(jeton);
  return rg_etiquettes_().map(function (e) { return e.nom; }).filter(function (n) { return !/^CATEGORY_/.test(n); }).sort();
}

/** Étiquette proposée pour une fiche (celle de la fiche si elle existe dans la boîte, sinon devinée d'après le tiers). */
function mob_etiquettePourFiche(jeton, id) {
  mob_verifier_(jeton);
  var f = mob_ligneFiche_(id), get = function (n) { return f.h[n] ? f.sh.getRange(f.ligne, f.h[n]).getDisplayValue() : ''; };
  var noms = rg_etiquettes_().map(function (e) { return e.nom; });
  var prop = [get('Étiquette finale'), get('Étiquette proposée'), get('Étiquette actuelle')].filter(function (n) { return n && noms.indexOf(n) >= 0; })[0];
  return { proposee: prop || rg_devine_(get('Tiers')), rangee: get('Date application') };
}

/** « Ranger » depuis l'application : étiquette choisie + sortie de la boîte de réception. */
function mob_ranger(jeton, id, etiquette) {
  mob_verifier_(jeton);
  if (/^C-/.test(String(id))) throw new Error('Un courrier scanné n\'est pas un mail : rien à ranger dans Gmail.');
  var verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    var f = mob_ligneFiche_(id);
    rg_rangerFil_(id, etiquette);
    rg_noterFiche_(f.sh, f.h, f.ligne, etiquette);
    journal_('Mail rangé dans « ' + etiquette + ' » (sorti de la réception)', id, '');
    return { ok: true, date: maintenant_(), etiquette: etiquette };
  } finally {
    verrou.releaseLock();
  }
}

/**
 * Dossiers clôturés : range tous les mails de la File de tri rattachés au dossier (même lien Gmail ou mêmes tiers + références)
 * qui ne sont pas encore rangés. Appelé à chaque ouverture de « Paiements RJ ».
 */
function rg_rangerClotures_() {
  var su = feuille_(RJ.F_SUIVI), n = su.getLastRow() - RJ.PREMIERE + 1;
  if (n < 1) return 0;
  var clos = su.getRange(RJ.PREMIERE, 1, n, 16).getDisplayValues().filter(function (l) {
    return l[2] && (/^✅/.test(l[11]) || /CL[ÔO]TUR/i.test(l[11]) || l[10]);
  });
  if (!clos.length) return 0;
  var tri = feuille_(APP.F_TRI), h = entetes_(tri), m = tri.getLastRow() - 1;
  if (m < 1 || !h['Date application']) return 0;
  var v = tri.getRange(2, 1, m, tri.getLastColumn()).getDisplayValues();
  var g = function (l, nom) { return h[nom] ? l[h[nom] - 1] : ''; };
  var nb = 0, debut = Date.now();
  v.forEach(function (l, i) {
    if (Date.now() - debut > 60000) return;
    var id = g(l, 'ID');
    if (!id || /^C-/.test(id) || g(l, 'Date application') || g(l, 'Statut') === 'À valider') return;
    var lien = g(l, 'Lien Gmail'), refs = g(l, 'Références'), tiers = rg_norm_(g(l, 'Tiers'));
    var dossier = clos.filter(function (d) {
      return (lien && d[15] && d[15] === lien) || (refs && d[3] === refs && rg_norm_(d[2]) === tiers);
    })[0];
    if (!dossier) return;
    var noms = rg_etiquettes_().map(function (e) { return e.nom; });
    var etiq = [g(l, 'Étiquette finale'), g(l, 'Étiquette proposée')].filter(function (x) { return x && noms.indexOf(x) >= 0; })[0] || rg_devine_(g(l, 'Tiers'));
    if (!etiq) return;   // pas d'étiquette sûre : on laisse en réception, Pierre rangera avec « Ranger »
    try {
      rg_rangerFil_(id, etiq);
      rg_noterFiche_(tri, h, i + 2, etiq);
      journal_('Dossier clôturé : mail rangé dans « ' + etiq + ' »', id, dossier[2] + ' | ' + dossier[3]);
      nb++;
    } catch (e) { /* on réessaiera à la prochaine ouverture */ }
  });
  return nb;
}
