// ── Configuration ─────────────────────────────────────────────────────────
var CONTACT_NAME  = "Nicaise ATEKOSSI";
var CONTACT_PHONE = "+229 0197082602";

// Colonnes 1-indexées (ordre réel du Sheet fourni par l'utilisateur)
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

// ── GET : recherche d'une fiche par email ──────────────────────────────────
function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();
  if (action === 'lookup') {
    var email = (e.parameter.email || '').toLowerCase().trim();
    return lookupRecord(email);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function lookupRecord(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data  = sheet.getDataRange().getValues();

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

  // Ligne 0 = en-têtes → données à partir de i = 1
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
        rowIndex: i + 1,  // 1-indexé pour Sheet
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
    var p      = e.parameter;
    var action = (p.action || 'submit').toLowerCase();

    if (action === 'update') {
      updateRecord(p);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'updated' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Soumission initiale (ordre du Sheet) ──────────────────────────────
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
      "",                  // Col 14 : CV (fichier – non transmis via HTML)
      p.prenoms     || "", // Col 15 : Prénoms
      "",                  // Col 16 : Photo (fichier – non transmis via HTML)
      p.email       || "", // Col 17 : Adresse e-mail (doublon)
      p.profession  || ""  // Col 18 : Profession
    ]);

    var fullName   = ((p.prenoms || "") + " " + (p.nom || "")).trim();
    var ownerEmail = Session.getActiveUser().getEmail();

    MailApp.sendEmail({
      to:       ownerEmail,
      subject:  "Nouvelle fiche reçue – " + fullName,
      htmlBody: buildOwnerEmail(p, fullName)
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

// ── Email au responsable ──────────────────────────────────────────────────
function buildOwnerEmail(p, fullName) {
  var fields = [
    ["Nom complet",         fullName],
    ["E-mail",              p.email       || "–"],
    ["Adresse",             p.adresse     || "–"],
    ["Téléphone 1",         p.tel1        || "–"],
    ["Téléphone 2",         p.tel2        || "–"],
    ["Profession",          p.profession  || "–"],
    ["Secteur d'activité",  p.secteur     || "–"],
    ["Poste actuel",        p.poste       || "–"],
    ["Niveau d'étude",      p.niveau      || "–"],
    ["Structure / Service", p.structure   || "–"],
    ["Région",              p.region      || "–"],
    ["District",            p.district    || "–"],
    ["Observation",         p.observation || "–"]
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
    +   '<p style="color:#333;font-size:15px;margin:0 0 16px">Une nouvelle fiche a été soumise par <strong>' + fullName + '</strong>.</p>'
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
    +     'Vos informations ont bien été enregistrées.'
    +   '</p>'
    +   '<div style="background:#f5f6fc;border-left:4px solid #f5a623;padding:16px 22px;'
    +     'border-radius:6px;margin-top:22px">'
    +     '<p style="margin:0;color:#1a2d7d;font-weight:700;font-size:14px">'
    +       'Pour modifier vos informations ou vous renseigner davantage :'
    +     '</p>'
    +     '<p style="margin:8px 0 0;color:#555;font-size:14px">'
    +       'Veuillez contacter <strong>' + CONTACT_NAME + '</strong>'
    +     '</p>'
    +     '<p style="margin:6px 0 0;font-size:17px;font-weight:700;color:#0d6eb8">'
    +       CONTACT_PHONE
    +     '</p>'
    +   '</div>'
    + '</div>'
    + '<div style="background:#1a2d7d;padding:14px 32px;text-align:center">'
    +   '<p style="color:#bbdefb;font-size:12px;margin:0">IFL… Influencing lives for Christ</p>'
    + '</div>'
    + '</div></body></html>';
}
