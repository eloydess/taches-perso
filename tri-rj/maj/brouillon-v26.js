

/* =====================================================================
 * Mail à l'administrateur (v2.6 — 14/10/2026)
 * Crée un BROUILLON Gmail avec les pièces jointes (position bancaire Excel + PDF des justificatifs).
 * Rien n'est envoyé : Pierre ouvre le brouillon dans Gmail, relit et envoie lui-même.
 * ===================================================================== */
function mob_brouillonAdmin(jeton, a, cc, objet, texte, pjs) {
  mob_verifier_(jeton);
  var adr = function (s) { return String(s || '').split(/[;,]\s*/).map(function (x) { return x.trim(); }).filter(String); };
  var valide = function (x) { return /^[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(x); };
  var dests = adr(a), copies = adr(cc);
  if (!dests.length || dests.concat(copies).some(function (x) { return !valide(x); })) throw new Error('Adresse invalide (À ou Copie).');
  texte = String(texte || '').trim();
  if (!texte) throw new Error('Le message est vide.');
  pjs = pjs || [];
  var total = 0;
  pjs.forEach(function (p) { total += Math.floor(String(p.b64 || '').length * 3 / 4); });
  if (total > 20 * 1024 * 1024) throw new Error('Pièces jointes trop lourdes (20 Mo au maximum).');
  var limite = 'limite_' + Utilities.getUuid().replace(/-/g, '');
  var L = ['To: ' + dests.join(', ')];
  if (copies.length) L.push('Cc: ' + copies.join(', '));
  L.push('Subject: ' + mob_enteteUtf8_(objet || 'Position bancaire'));
  L.push('MIME-Version: 1.0');
  L.push('Content-Type: multipart/mixed; boundary="' + limite + '"');
  L.push('');
  L.push('--' + limite);
  L.push('Content-Type: text/plain; charset="UTF-8"');
  L.push('Content-Transfer-Encoding: base64');
  L.push('');
  L.push(mob_b64Lignes_(Utilities.newBlob(texte).getBytes()));
  pjs.forEach(function (p) {
    var nom = String(p.nom || 'piece-jointe').replace(/["\\\r\n]/g, '');
    L.push('--' + limite);
    L.push('Content-Type: ' + String(p.type || 'application/octet-stream').replace(/[\r\n]/g, '') + '; name="' + mob_enteteUtf8_(nom) + '"');
    L.push('Content-Disposition: attachment; filename="' + mob_enteteUtf8_(nom) + '"');
    L.push('Content-Transfer-Encoding: base64');
    L.push('');
    L.push(String(p.b64).replace(/[^A-Za-z0-9+\/=]/g, '').replace(/(.{76})/g, '$1\r\n'));
  });
  L.push('--' + limite + '--');
  var r = mob_gmailApi_('post', 'drafts', { message: { raw: Utilities.base64EncodeWebSafe(L.join('\r\n')) } });
  var msgId = r.message && r.message.id;
  journal_('Brouillon pour l\'administrateur créé (' + pjs.length + ' pièce(s) jointe(s))', '', objet);
  return { id: r.id, url: 'https://mail.google.com/mail/?authuser=pharmaciecsudachat@gmail.com#drafts?compose=' + msgId };
}
