// ── Configuration ─────────────────────────────────────────────────────────
var CONTACT_NAME       = "Nicaise ATEKOSSI";
var CONTACT_PHONE      = "+229 0197082602";
var UPLOAD_FOLDER_NAME = "IFL – Pièces jointes";

// Colonnes 1-indexées (ordre réel du Sheet)
// Col1:Horodateur | Col2:Nom | Col3:Email | Col4:Adresse | Col5:Tél1 | Col6:Tél2
// Col7:Secteur | Col8:Poste | Col9:Niveau | Col10:Structure | Col11:Région
// Col12:District | Col13:Observation | Col14:CV | Col15:Prénoms | Col16:Photo
// Col17:Email(dup) | Col18:Profession
var COL_MAP = {
  nom:         2,
  adresse:     4,
  tel1:        5,
  tel2:        6,
  secteur:     7,
  poste:       8,
  niveau:      9,
  structure:   10,
  region:      11,
  district:    12,
  observation: 13,
  prenoms:     15,
  profession:  18
};

var EMAIL_COL_0 = 2; // index 0-basé → col 3 du sheet (Adresse e-mail)

// ── Google Drive – upload fichier ──────────────────────────────────────────
function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function saveFileToDrive(base64, mimeType, filename) {
  var folder = getOrCreateFolder(UPLOAD_FOLDER_NAME);
  var bytes  = Utilities.base64Decode(base64);
  var blob   = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', filename);
  var file   = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ── GET : recherche d'une fiche par email + actions admin ─────────────────
function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();
  if (action === 'lookup') {
    var email = (e.parameter.email || '').toLowerCase().trim();
    return lookupRecord(email);
  }
  if (action === 'adminlogin') {
    return adminLogin(e.parameter.user || '', e.parameter.pass || '');
  }
  if (action === 'getdata') {
    return getAdminData(e.parameter.user || '', e.parameter.pass || '');
  }
  if (action === 'listadmins') {
    return listAdminUsers(e.parameter.user || '', e.parameter.pass || '');
  }
  if (action === 'addadmin') {
    return addAdminUser(
      e.parameter.user    || '',
      e.parameter.pass    || '',
      e.parameter.newuser || '',
      e.parameter.newpass || ''
    );
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function lookupRecord(email) {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data   = sheet.getDataRange().getValues();
  var labels = {
    nom:         'Nom',
    prenoms:     'Prénoms',
    adresse:     'Adresse',
    tel1:        'Numéro de téléphone 1',
    tel2:        'Numéro de téléphone 2',
    secteur:     "Secteur d'activité",
    poste:       'Poste actuel',
    niveau:      "Niveau d'étude",
    structure:   'Structure ou Service',
    region:      'Région',
    district:    'District',
    observation: 'Observation',
    profession:  'Profession'
  };

  for (var i = 1; i < data.length; i++) {
    var rowEmail = (data[i][EMAIL_COL_0] || '').toString().toLowerCase().trim();
    if (rowEmail !== email) continue;

    var missing = [];
    Object.keys(COL_MAP).forEach(function(key) {
      var val = (data[i][COL_MAP[key] - 1] || '').toString().trim();
      if (!val) missing.push({ key: key, label: labels[key] });
    });

    return ContentService
      .createTextOutput(JSON.stringify({
        status:   'found',
        rowIndex: i + 1,
        missing:  missing
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'not_found' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST : soumission nouvelle fiche OU mise à jour ────────────────────────
function doPost(e) {
  try {
    // Lecture JSON (envoyé en text/plain) avec fallback url-encoded
    var p;
    try { p = JSON.parse(e.postData.contents); }
    catch(ex) { p = e.parameter; }

    var action = (p.action || 'submit').toLowerCase();

    if (action === 'update') {
      updateRecord(p);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'updated' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Upload CV et Photo vers Google Drive ──────────────────────────────
    var cvUrl    = '';
    var photoUrl = '';

    if (p.cvData && p.cvName) {
      try { cvUrl = saveFileToDrive(p.cvData, p.cvMime, p.cvName); } catch(fe) {}
    }
    if (p.photoData && p.photoName) {
      try { photoUrl = saveFileToDrive(p.photoData, p.photoMime, p.photoName); } catch(fe) {}
    }

    // ── Enregistrement dans le Sheet ──────────────────────────────────────
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    sheet.appendRow([
      new Date(),           // Col 1  : Horodateur
      p.nom         || "", // Col 2  : Nom
      p.email       || "", // Col 3  : Adresse e-mail
      p.adresse     || "", // Col 4  : Adresse
      p.tel1        || "", // Col 5  : Numéro de téléphone1
      p.tel2        || "", // Col 6  : Numéro de téléphone
      p.secteur     || "", // Col 7  : Secteur D'activité
      p.poste       || "", // Col 8  : Poste Actuel
      p.niveau      || "", // Col 9  : Niveau d'étude
      p.structure   || "", // Col 10 : Structure ou Service
      p.region      || "", // Col 11 : Région
      p.district    || "", // Col 12 : District
      p.observation || "", // Col 13 : Observation
      cvUrl,               // Col 14 : CV (lien Drive)
      p.prenoms     || "", // Col 15 : Prénoms
      photoUrl,            // Col 16 : Photo (lien Drive)
      p.email       || "", // Col 17 : Adresse e-mail (doublon)
      p.profession  || ""  // Col 18 : Profession
    ]);

    var fullName   = ((p.prenoms || "") + " " + (p.nom || "")).trim();
    var ownerEmail = Session.getActiveUser().getEmail();

    MailApp.sendEmail({
      to:       ownerEmail,
      subject:  "Nouvelle fiche reçue – " + fullName,
      htmlBody: buildOwnerEmail(p, fullName, cvUrl, photoUrl)
    });

    if (p.email) {
      MailApp.sendEmail({
        to:       p.email,
        subject:  "Confirmation de votre fiche – IFL Bénin",
        htmlBody: buildConfirmEmail(p, fullName)
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Mise à jour d'une fiche existante ─────────────────────────────────────
function updateRecord(p) {
  var sheet    = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowIndex = parseInt(p.rowIndex, 10);
  if (!rowIndex) throw new Error('rowIndex manquant');

  Object.keys(COL_MAP).forEach(function(key) {
    if (p[key] && p[key].toString().trim()) {
      sheet.getRange(rowIndex, COL_MAP[key]).setValue(p[key]);
    }
  });
}

// ── Email au responsable (avec liens CV & Photo) ───────────────────────────
function buildOwnerEmail(p, fullName, cvUrl, photoUrl) {
  var cvLink    = cvUrl    ? '<a href="' + cvUrl    + '" style="color:#0d6eb8">Télécharger le CV</a>'  : '–';
  var photoLink = photoUrl ? '<a href="' + photoUrl + '" style="color:#0d6eb8">Voir la photo</a>'      : '–';

  var fields = [
    ["Nom complet",          fullName],
    ["E-mail",               p.email       || "–"],
    ["Adresse",              p.adresse     || "–"],
    ["Téléphone 1",          p.tel1        || "–"],
    ["Téléphone 2",          p.tel2        || "–"],
    ["Profession",           p.profession  || "–"],
    ["Secteur d'activité",   p.secteur     || "–"],
    ["Poste actuel",         p.poste       || "–"],
    ["Niveau d'étude",       p.niveau      || "–"],
    ["Structure / Service",  p.structure   || "–"],
    ["Région",               p.region      || "–"],
    ["District",             p.district    || "–"],
    ["Observation",          p.observation || "–"],
    ["CV",                   cvLink],
    ["Photo",                photoLink]
  ];

  var rows = fields.map(function(f) {
    return '<tr>'
      + '<td style="padding:8px 14px;font-weight:700;color:#1a2d7d;background:#f5f6fc;'
      +   'border:1px solid #c5cae9;white-space:nowrap">' + f[0] + '</td>'
      + '<td style="padding:8px 14px;border:1px solid #c5cae9;color:#333">' + f[1] + '</td>'
      + '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;background:#e8eaf6;padding:24px;margin:0">'
    + '<div style="max-width:640px;margin:auto;background:#fff;border-radius:10px;overflow:hidden;'
    +   'box-shadow:0 4px 20px rgba(26,45,125,.15)">'
    + '<div style="background:linear-gradient(135deg,#1a2d7d,#0d6eb8);padding:24px 32px">'
    +   '<h2 style="color:#fff;margin:0;font-size:20px">Nouvelle fiche de renseignement</h2>'
    +   '<p style="color:#ffe082;margin:6px 0 0;font-size:13px">Ligue Internationale de l\'Amitié – Bénin</p>'
    + '</div>'
    + '<div style="padding:24px 32px">'
    +   '<p style="color:#333;font-size:15px;margin:0 0 16px">Nouvelle fiche soumise par <strong>' + fullName + '</strong>.</p>'
    +   '<table style="width:100%;border-collapse:collapse;font-size:14px">' + rows + '</table>'
    + '</div>'
    + '<div style="background:#1a2d7d;padding:14px 32px;text-align:center">'
    +   '<p style="color:#bbdefb;font-size:12px;margin:0">'
    +     'Formulaire géré par <strong style="color:#ffe082">' + CONTACT_NAME + '</strong>'
    +     ' &nbsp;|&nbsp; ' + CONTACT_PHONE
    +   '</p>'
    + '</div>'
    + '</div></body></html>';
}

// ── Email de confirmation au déclarant ────────────────────────────────────
function buildConfirmEmail(p, fullName) {
  return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;background:#e8eaf6;padding:24px;margin:0">'
    + '<div style="max-width:640px;margin:auto;background:#fff;border-radius:10px;overflow:hidden;'
    +   'box-shadow:0 4px 20px rgba(26,45,125,.15)">'
    + '<div style="background:linear-gradient(135deg,#1a2d7d,#0d6eb8);padding:24px 32px">'
    +   '<h2 style="color:#fff;margin:0;font-size:20px">Confirmation de votre fiche</h2>'
    +   '<p style="color:#ffe082;margin:6px 0 0;font-size:13px">Ligue Internationale de l\'Amitié – Bénin</p>'
    + '</div>'
    + '<div style="padding:28px 32px">'
    +   '<p style="color:#333;font-size:15px;margin:0">Bonjour <strong>' + fullName + '</strong>,</p>'
    +   '<p style="color:#333;font-size:15px;margin:14px 0 0">'
    +     'Vous venez de renseigner votre fiche auprès de la '
    +     '<strong>Ligue Internationale de l\'Amitié – Bénin</strong>. '
    +     'Vos informations et documents ont bien été enregistrés.'
    +   '</p>'
    +   '<div style="background:#f5f6fc;border-left:4px solid #f5a623;padding:16px 22px;border-radius:6px;margin-top:22px">'
    +     '<p style="margin:0;color:#1a2d7d;font-weight:700;font-size:14px">'
    +       'Pour modifier vos informations ou vous renseigner davantage :'
    +     '</p>'
    +     '<p style="margin:8px 0 0;color:#555;font-size:14px">Veuillez contacter <strong>' + CONTACT_NAME + '</strong></p>'
    +     '<p style="margin:6px 0 0;font-size:17px;font-weight:700;color:#0d6eb8">' + CONTACT_PHONE + '</p>'
    +   '</div>'
    + '</div>'
    + '<div style="background:#1a2d7d;padding:14px 32px;text-align:center">'
    +   '<p style="color:#bbdefb;font-size:12px;margin:0">IFL… Influencing lives for Christ</p>'
    + '</div>'
    + '</div></body></html>';
}

// ── Admin : gestion des accès ─────────────────────────────────────────────

var ADMIN_KEY    = 'IFL_ADMINS';
var DEFAULT_USER = 'admin';
var DEFAULT_PASS = 'IFL@Benin2024!';

function hashPass(pass) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, pass, Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

function getAdmins() {
  var props = PropertiesService.getScriptProperties();
  var json  = props.getProperty(ADMIN_KEY);
  if (!json) {
    var init = [{ user: DEFAULT_USER, passHash: hashPass(DEFAULT_PASS) }];
    props.setProperty(ADMIN_KEY, JSON.stringify(init));
    return init;
  }
  return JSON.parse(json);
}

function verifyAdmin(u, p) {
  var h = hashPass(p);
  return getAdmins().some(function(a) { return a.user === u && a.passHash === h; });
}

function adminLogin(u, p) {
  var ok = !!(u && p && verifyAdmin(u, p));
  return ContentService
    .createTextOutput(JSON.stringify(ok
      ? { status: 'ok' }
      : { status: 'error', message: 'Identifiants incorrects' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAdminData(u, p) {
  if (!verifyAdmin(u, p)) return ContentService
    .createTextOutput(JSON.stringify({ status: 'unauthorized' }))
    .setMimeType(ContentService.MimeType.JSON);
  var vals = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', data: vals }))
    .setMimeType(ContentService.MimeType.JSON);
}

function listAdminUsers(u, p) {
  if (!verifyAdmin(u, p)) return ContentService
    .createTextOutput(JSON.stringify({ status: 'unauthorized' }))
    .setMimeType(ContentService.MimeType.JSON);
  var names = getAdmins().map(function(a) { return a.user; });
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', admins: names }))
    .setMimeType(ContentService.MimeType.JSON);
}

function addAdminUser(u, p, nu, np) {
  if (!verifyAdmin(u, p)) return ContentService
    .createTextOutput(JSON.stringify({ status: 'unauthorized' }))
    .setMimeType(ContentService.MimeType.JSON);
  if (!nu || !np) return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: 'Champs manquants' }))
    .setMimeType(ContentService.MimeType.JSON);
  var admins = getAdmins();
  if (admins.some(function(a) { return a.user === nu; })) return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: 'Identifiant déjà existant' }))
    .setMimeType(ContentService.MimeType.JSON);
  admins.push({ user: nu, passHash: hashPass(np) });
  PropertiesService.getScriptProperties().setProperty(ADMIN_KEY, JSON.stringify(admins));
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Menu Google Sheets ────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Administration IFL')
    .addItem('Créer un administrateur', 'showCreateAdminDialog')
    .addSeparator()
    .addItem('Liste des administrateurs', 'showAdminsList')
    .addToUi();
}

function showAdminsList() {
  var admins = getAdmins();
  var lines  = admins.map(function(a, i) { return (i + 1) + '.  ' + a.user; }).join('\n');
  SpreadsheetApp.getUi().alert(
    'Administrateurs IFL (' + admins.length + ')',
    lines || 'Aucun administrateur trouvé.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showCreateAdminDialog() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8"/>'
    + '<style>'
    + '*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:Segoe UI,Arial,sans-serif;background:#f0f2ff;padding:16px}'
    + '.card{background:#fff;padding:24px;border-radius:12px;'
    +   'box-shadow:0 4px 20px rgba(26,45,125,.12)}'
    + 'h2{font-size:15px;font-weight:800;color:#1a2d7d;margin-bottom:18px;'
    +   'padding-bottom:10px;border-bottom:2px solid #e8eaf6}'
    + 'label{display:block;font-size:11px;font-weight:700;color:#1a2d7d;'
    +   'text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px}'
    + '.fg{margin-bottom:12px}'
    + 'input{width:100%;border:1.5px solid #9fa8da;border-radius:6px;'
    +   'padding:8px 11px;font-size:14px;color:#1a237e;background:#f5f6fc;'
    +   'outline:none;font-family:inherit}'
    + 'input:focus{border-color:#0d6eb8;box-shadow:0 0 0 3px rgba(13,110,184,.12);background:#fff}'
    + '.btn{width:100%;background:linear-gradient(135deg,#1a2d7d,#0d6eb8);'
    +   'color:#fff;border:none;border-radius:20px;padding:11px;font-size:14px;'
    +   'font-weight:700;cursor:pointer;font-family:inherit;margin-top:6px}'
    + '.btn:hover{opacity:.9}.btn:disabled{opacity:.55;cursor:default}'
    + '.msg{margin-top:10px;font-size:13px;text-align:center;min-height:18px}'
    + '.ok{color:#2e7d32}.er{color:#e53935}'
    + '</style></head><body>'
    + '<div class="card">'
    + '<h2>Nouvel administrateur</h2>'
    + '<div class="fg"><label>Identifiant <span style="color:#e53935">*</span></label>'
    + '<input type="text" id="u" placeholder="Nom d\'utilisateur" autocomplete="off"/></div>'
    + '<div class="fg"><label>Mot de passe <span style="color:#e53935">*</span></label>'
    + '<input type="password" id="p" placeholder="Minimum 8 caractères"/></div>'
    + '<div class="fg"><label>Confirmer <span style="color:#e53935">*</span></label>'
    + '<input type="password" id="p2" placeholder="Répéter le mot de passe"/></div>'
    + '<button class="btn" id="btn" onclick="save()">Créer l\'administrateur</button>'
    + '<p class="msg" id="msg"></p>'
    + '</div>'
    + '<script>'
    + 'function save(){'
    + 'var u=document.getElementById("u").value.trim();'
    + 'var p=document.getElementById("p").value;'
    + 'var p2=document.getElementById("p2").value;'
    + 'var msg=document.getElementById("msg");'
    + 'msg.textContent="";msg.className="msg";'
    + 'if(!u||!p||!p2){msg.textContent="Veuillez remplir tous les champs.";msg.className="msg er";return;}'
    + 'if(p.length<8){msg.textContent="Mot de passe : 8 caractères minimum.";msg.className="msg er";return;}'
    + 'if(p!==p2){msg.textContent="Les mots de passe ne correspondent pas.";msg.className="msg er";return;}'
    + 'document.getElementById("btn").disabled=true;'
    + 'msg.textContent="Création en cours…";'
    + 'google.script.run'
    + '.withSuccessHandler(function(r){'
    + 'document.getElementById("btn").disabled=false;'
    + 'if(r.ok){'
    +   'msg.textContent="✓ "+r.message;msg.className="msg ok";'
    +   'document.getElementById("u").value="";'
    +   'document.getElementById("p").value="";'
    +   'document.getElementById("p2").value="";'
    + '}else{msg.textContent=r.message;msg.className="msg er";}'
    + '})'
    + '.withFailureHandler(function(e){'
    + 'document.getElementById("btn").disabled=false;'
    + 'document.getElementById("msg").textContent="Erreur : "+e.message;'
    + 'document.getElementById("msg").className="msg er";'
    + '})'
    + '.createAdminFromSheet(u,p);'
    + '}'
    + 'document.getElementById("p2").addEventListener("keydown",function(e){if(e.key==="Enter")save();});'
    + '<\/script></body></html>'
  ).setWidth(360).setHeight(370);
  SpreadsheetApp.getUi().showModalDialog(html, 'Nouvel administrateur – IFL');
}

function createAdminFromSheet(user, pass) {
  try {
    if (!user || !pass) return { ok: false, message: 'Identifiants manquants.' };
    var admins = getAdmins();
    if (admins.some(function(a) { return a.user === user; })) {
      return { ok: false, message: 'L\'identifiant « ' + user + ' » existe déjà.' };
    }
    admins.push({ user: user, passHash: hashPass(pass) });
    PropertiesService.getScriptProperties().setProperty(ADMIN_KEY, JSON.stringify(admins));
    return { ok: true, message: 'Administrateur « ' + user + ' » créé avec succès !' };
  } catch(err) {
    return { ok: false, message: 'Erreur : ' + err.toString() };
  }
}
