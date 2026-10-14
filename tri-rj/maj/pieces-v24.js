

/* =====================================================================
 * Pièces des dossiers (v2.4 — 14/10/2026)
 * Colonne S du « Suivi factures » : preuves de virement (identifiants Drive séparés par des espaces).
 * Colonne T : factures et autres documents, notés « F:id » (facture) ou « D:id » (document).
 * Les fichiers sont rangés dans le dossier Drive « Preuves de virement — Application Pharmacie ».
 * La facture peut être retrouvée automatiquement dans le mail d'origine (ou le courrier scanné) du dossier.
 * ===================================================================== */
var PC = { COL: 20, MAX: 9 * 1024 * 1024, GENRES: ['preuve', 'facture', 'document'] };

function pc_entete_(sh) {
  if (!sh.getRange(3, PC.COL).getValue()) {
    sh.getRange(3, PC.COL).setValue('Factures et documents\n(fichiers Drive)').setFontWeight('bold').setWrap(true).setBackground('#e8eaed');
  }
}

/** Ligne du dossier (valeurs affichées A → T). */
function pc_ligne_(cle) {
  var sh = feuille_(RJ.F_SUIVI);
  pc_entete_(sh);
  var n = sh.getLastRow() - RJ.PREMIERE + 1;
  var v = n > 0 ? sh.getRange(RJ.PREMIERE, 1, n, PC.COL).getDisplayValues() : [];
  for (var i = 0; i < v.length; i++) {
    if ('RJ|' + v[i][2] + '|' + v[i][3] === cle) return { sh: sh, ligne: RJ.PREMIERE + i, l: v[i] };
  }
  throw new Error('Dossier introuvable dans « Suivi factures ». Rechargez l\'écran.');
}

/** [{ id, genre }] à partir des colonnes S et T d'une ligne. */
function pc_liste_(l) {
  var res = [];
  String(l[18] || '').split(' ').filter(String).forEach(function (id) { res.push({ id: id, genre: 'preuve' }); });
  String(l[19] || '').split(' ').filter(String).forEach(function (x) {
    var m = /^([FD]):(.+)$/.exec(x);
    if (m) res.push({ id: m[2], genre: m[1] === 'F' ? 'facture' : 'document' });
  });
  return res;
}

function pc_compte_(l) {
  var c = { preuve: 0, facture: 0, document: 0 };
  pc_liste_(l).forEach(function (p) { c[p.genre]++; });
  return c;
}

function pc_ajouter_(d, genre, id) {
  if (genre === 'preuve') {
    var s = d.l[18] ? d.l[18].split(' ').filter(String) : [];
    if (s.indexOf(id) < 0) s.push(id);
    d.sh.getRange(d.ligne, PV.COL).setNumberFormat('@').setValue(s.join(' '));
    d.l[18] = s.join(' ');
  } else {
    var t = d.l[19] ? d.l[19].split(' ').filter(String) : [];
    var jeton = (genre === 'facture' ? 'F:' : 'D:') + id;
    if (t.indexOf(jeton) < 0) t.push(jeton);
    d.sh.getRange(d.ligne, PC.COL).setNumberFormat('@').setValue(t.join(' '));
    d.l[19] = t.join(' ');
  }
}

/** Liste des pièces avec leur nom de fichier. */
function pc_infos_(l) {
  return pc_liste_(l).map(function (p) {
    var nom = '';
    try { nom = cr_driveApi_('get', 'https://www.googleapis.com/drive/v3/files/' + p.id + '?fields=name').name; }
    catch (e) { nom = '(fichier introuvable dans Drive)'; }
    return { id: p.id, genre: p.genre, nom: nom };
  });
}

function pc_nom_(genre, d, nomOrigine, ext) {
  var base = genre.toUpperCase() + ' ' + Utilities.formatDate(new Date(), APP.FUSEAU, 'yyyy-MM-dd') + ' — ' + d.l[2] + ' — ' + d.l[3] +
    (nomOrigine ? ' — ' + String(nomOrigine).replace(/\.[^.]+$/, '') : '');
  return base.replace(/[\\\/:*?"<>|\r\n]/g, ' ').slice(0, 150) + '.' + ext;
}

function pc_ext_(type) {
  return type === 'application/pdf' ? 'pdf' : (String(type).split('/')[1] === 'jpeg' ? 'jpg' : String(type).split('/')[1] || 'bin');
}

/** Ajoute une pièce (preuve, facture ou document) au dossier. f = { nom, type, b64 } */
function mob_piece(jeton, cle, genre, f) {
  mob_verifier_(jeton);
  if (PC.GENRES.indexOf(genre) < 0) throw new Error('Type de pièce inconnu.');
  if (!f || !f.b64) throw new Error('Aucun fichier.');
  if (!/^(application\/pdf|image\/(jpeg|png|heic|heif|webp))$/.test(String(f.type))) throw new Error('Format non accepté : ' + f.type);
  if (Math.floor(String(f.b64).length * 3 / 4) > PC.MAX) throw new Error('Fichier trop lourd (9 Mo au maximum).');
  var verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    var d = pc_ligne_(cle);
    var fi = cr_televerser_(pv_dossier_(), pc_nom_(genre, d, f.nom, pc_ext_(f.type)), f.type, f.b64);
    pc_ajouter_(d, genre, fi.id);
    journal_('Pièce ajoutée (' + genre + ') : ' + d.l[2], '', d.l[3]);
    return { pieces: pc_infos_(d.l), compte: pc_compte_(d.l) };
  } finally {
    verrou.releaseLock();
  }
}

function mob_pieces(jeton, cle) {
  mob_verifier_(jeton);
  var d = pc_ligne_(cle);
  return { pieces: pc_infos_(d.l), compte: pc_compte_(d.l) };
}

/** Contenu d'une pièce du dossier (pour l'afficher ou l'assembler dans un PDF). */
function mob_fichier(jeton, cle, id) {
  mob_verifier_(jeton);
  var d = pc_ligne_(cle);
  if (!pc_liste_(d.l).some(function (p) { return p.id === id; })) throw new Error('Ce fichier n\'est pas rattaché à ce dossier.');
  var meta = cr_driveApi_('get', 'https://www.googleapis.com/drive/v3/files/' + id + '?fields=name,mimeType,size');
  if (Number(meta.size || 0) > PC.MAX) throw new Error('Fichier trop lourd pour l\'affichage (9 Mo au maximum).');
  var rep = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '?alt=media',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (rep.getResponseCode() !== 200) throw new Error('Fichier illisible dans Drive (' + rep.getResponseCode() + ').');
  return { nom: meta.name, type: meta.mimeType, b64: Utilities.base64Encode(rep.getContent()) };
}

/** Retire une pièce du dossier (le fichier reste dans Drive). */
function mob_retirerPiece(jeton, cle, id) {
  mob_verifier_(jeton);
  var verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    var d = pc_ligne_(cle);
    var s = (d.l[18] || '').split(' ').filter(function (x) { return x && x !== id; });
    var t = (d.l[19] || '').split(' ').filter(function (x) { return x && x !== 'F:' + id && x !== 'D:' + id; });
    d.sh.getRange(d.ligne, PV.COL).setNumberFormat('@').setValue(s.join(' '));
    d.sh.getRange(d.ligne, PC.COL).setNumberFormat('@').setValue(t.join(' '));
    d.l[18] = s.join(' '); d.l[19] = t.join(' ');
    journal_('Pièce retirée du dossier : ' + d.l[2], '', d.l[3]);
    return { pieces: pc_infos_(d.l), compte: pc_compte_(d.l) };
  } finally {
    verrou.releaseLock();
  }
}

/** Identifiant décimal Gmail (« thread-f:… ») → identifiant hexadécimal du fil. */
function pc_decHex_(s) {
  var chiffres = String(s).split('').map(Number), out = '';
  while (chiffres.length && !(chiffres.length === 1 && chiffres[0] === 0)) {
    var r = 0, q = [];
    chiffres.forEach(function (x) { var v = r * 10 + x, t = Math.floor(v / 16); r = v % 16; if (q.length || t) q.push(t); });
    out = '0123456789abcdef'.charAt(r) + out;
    chiffres = q;
  }
  return out || '0';
}

/** Mail (id du fil) ou courrier (C-…) d'où vient le dossier. */
function pc_source_(d) {
  var lien = d.l[15] || '';
  var tri = feuille_(APP.F_TRI), h = entetes_(tri), n = tri.getLastRow() - 1;
  if (n > 0) {
    var v = tri.getRange(2, 1, n, tri.getLastColumn()).getDisplayValues();
    for (var i = 0; i < v.length; i++) {
      var id = v[i][h['ID'] - 1];
      if (!id) continue;
      if (lien && h['Lien Gmail'] && v[i][h['Lien Gmail'] - 1] === lien) return id;
      if (h['Références'] && v[i][h['Références'] - 1] && v[i][h['Références'] - 1] === d.l[3] &&
          String(v[i][h['Tiers'] - 1]).toUpperCase() === String(d.l[2]).toUpperCase()) return id;
    }
  }
  var m = /thread-f:(\d+)/.exec(lien);
  if (m) return pc_decHex_(m[1]);
  m = /#[a-z]+\/([0-9a-f]{16})\b/.exec(lien);
  return m ? m[1] : '';
}

/** Retrouve la facture du dossier et la range dans ses pièces. Renvoie le nombre de fichiers ajoutés. */
function pc_factureAuto_(d) {
  var src = pc_source_(d);
  if (!src) return 0;
  if (/^C-/.test(src)) {
    var cr = cr_feuille_(), n = cr.getLastRow() - 1;
    if (n < 1) return 0;
    var lig = cr.getRange(2, 1, n, 6).getDisplayValues().filter(function (r) { return r[0] === src; })[0];
    if (!lig) return 0;
    var ids = String(lig[4] || '').split(' ').filter(String);
    ids.forEach(function (id) { pc_ajouter_(d, 'facture', id); });
    return ids.length;
  }
  var fil;
  try { fil = GmailApp.getThreadById(src); } catch (e) { fil = null; }
  if (!fil) return 0;
  var pjs = [];
  fil.getMessages().forEach(function (msg) {
    msg.getAttachments({ includeInlineImages: false }).forEach(function (a) {
      var nom = a.getName() || '';
      if (!/pdf/i.test(a.getContentType()) && !/\.pdf$/i.test(nom)) return;
      if (/preuve|virement|rib|cgv|conditions|courrier.?proforma.?groupe/i.test(nom)) return;
      if (a.getSize() > PC.MAX) return;
      pjs.push(a);
    });
  });
  pjs = pjs.slice(-3);
  pjs.forEach(function (a) {
    var fi = cr_televerser_(pv_dossier_(), pc_nom_('facture', d, a.getName(), 'pdf'), 'application/pdf', Utilities.base64Encode(a.getBytes()));
    pc_ajouter_(d, 'facture', fi.id);
  });
  return pjs.length;
}

/**
 * Récupère automatiquement les factures manquantes.
 * cle : un seul dossier ; sans cle : tous les dossiers en cours qui n'ont pas encore de facture.
 */
function mob_facturesAuto(jeton, cle) {
  mob_verifier_(jeton);
  var debut = Date.now();
  var verrou = LockService.getScriptLock();
  verrou.waitLock(30000);
  try {
    var sh = feuille_(RJ.F_SUIVI);
    pc_entete_(sh);
    var n = sh.getLastRow() - RJ.PREMIERE + 1;
    var v = n > 0 ? sh.getRange(RJ.PREMIERE, 1, n, PC.COL).getDisplayValues() : [];
    var res = { trouvees: [], sansFacture: [] };
    for (var i = 0; i < v.length; i++) {
      var l = v[i];
      if (!l[2]) continue;
      var c = 'RJ|' + l[2] + '|' + l[3];
      if (cle ? (c !== cle || pc_compte_(l).facture > 0) : (/^✅/.test(l[11]) || /CL[ÔO]TUR/i.test(l[11]) || pc_compte_(l).facture > 0)) continue;
      if (Date.now() - debut > 240000) { res.sansFacture.push(l[2] + ' (temps dépassé : relancez)'); continue; }
      var d = { sh: sh, ligne: RJ.PREMIERE + i, l: l };
      var nb = 0;
      try { nb = pc_factureAuto_(d); } catch (e) { nb = 0; }
      if (nb) res.trouvees.push(l[2] + ' — ' + l[3] + ' (' + nb + ' fichier' + (nb > 1 ? 's' : '') + ')');
      else res.sansFacture.push(l[2] + ' — ' + l[3]);
    }
    if (res.trouvees.length) journal_('Factures retrouvées automatiquement : ' + res.trouvees.length, '', '');
    return res;
  } finally {
    verrou.releaseLock();
  }
}
