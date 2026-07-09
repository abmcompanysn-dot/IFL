// ── Configuration ─────────────────────────────────────────────────────────
var CONTACT_NAME       = "Nicaise ATEKOSSI";
var CONTACT_PHONE      = "+229 0197082602";
var UPLOAD_FOLDER_NAME = "IFL – Pièces jointes";

// Colonnes 1-indexées (ordre réel du Sheet)
// Col1:Horodateur | Col2:Nom | Col3:Email | Col4:Adresse | Col5:Tél1 | Col6:Tél2
// Col7:Secteur | Col8:Poste | Col9:Niveau | Col10:Structure | Col11:Région
// Col12:District | Col13:Observation | Col14:CV | Col15:Prénoms | Col16:Photo
// Col17:Email(dup) | Col18:Profession | Col19:Résumé automatique CV
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

// ── Extraction de texte depuis un fichier Drive (CV) ───────────────────────
function extractDriveFileIdFromUrl(url) {
  if (!url) return '';
  var m = url.toString().match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

function extractTextFromDriveFile(fileId) {
  // Nécessite le service avancé "Drive API" activé dans le projet Apps Script.
  var tempFile = null;
  try {
    var blob = DriveApp.getFileById(fileId).getBlob();

    if (typeof Drive === 'undefined') {
      return { text: '', error: "Le service avancé Drive n'est pas activé (identifiant global 'Drive' introuvable). Activez-le via Services (+) dans l'éditeur Apps Script." };
    }

    var isV3 = !!(Drive.Files && Drive.Files.create);
    if (isV3) {
      var resourceV3 = { name: 'tmp_ocr_' + fileId, mimeType: MimeType.GOOGLE_DOCS };
      tempFile = Drive.Files.create(resourceV3, blob, { ocrLanguage: 'fr' });
    } else {
      var resourceV2 = { title: 'tmp_ocr_' + fileId, mimeType: MimeType.GOOGLE_DOCS };
      tempFile = Drive.Files.insert(resourceV2, blob, { convert: true, ocr: true, ocrLanguage: 'fr' });
    }

    var text = DocumentApp.openById(tempFile.id).getBody().getText();
    return { text: (text || '').trim(), error: '' };
  } catch (err) {
    return { text: '', error: err && err.message ? err.message : err.toString() };
  } finally {
    if (tempFile && tempFile.id) {
      try {
        if (Drive.Files.remove) { Drive.Files.remove(tempFile.id); }
        else { Drive.Files.delete(tempFile.id); }
      } catch (ex) {}
    }
  }
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
  return createJsonResponse({ status: 'ok' });
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

    return createJsonResponse({
      status:   'found',
      rowIndex: i + 1,
      missing:  missing
    });
  }

  return createJsonResponse({ status: 'not_found' });
}

// ── POST : soumission nouvelle fiche OU mise à jour ────────────────────────
function createJsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON);
}

var AI_GROQ_KEY    = 'GROQ_API_KEY';
var AI_GROQ_MODEL  = 'GROQ_MODEL';
var AI_OPENAI_KEY  = 'OPENAI_API_KEY';
var AI_OPENAI_MODEL = 'OPENAI_MODEL';

function getAiConfig() {
  // Placez vos clés dans les propriétés de script de Google Apps Script :
  // GROQ_API_KEY, GROQ_MODEL ou OPENAI_API_KEY, OPENAI_MODEL.
  var props = PropertiesService.getScriptProperties();
  return {
    groqKey: props.getProperty(AI_GROQ_KEY),
    groqModel: props.getProperty(AI_GROQ_MODEL) || 'llama-3.3-70b-versatile',
    openaiKey: props.getProperty(AI_OPENAI_KEY),
    openaiModel: props.getProperty(AI_OPENAI_MODEL) || 'gpt-4o-mini'
  };
}

function callGroqModel(prompt, apiKey, model) {
  try {
    var url = 'https://api.groq.com/openai/v1/chat/completions';
    var body = {
      model: model,
      messages: [
        { role: 'system', content: 'Vous êtes un assistant qui résume des CV en français.' },
        { role: 'user', content: prompt }
      ],
      temperature: 0.2,
      max_tokens: 800
    };
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code !== 200) {
      return { content: null, error: 'Groq HTTP ' + code + ' : ' + response.getContentText().slice(0, 300) };
    }
    var result = JSON.parse(response.getContentText());
    if (result.choices && result.choices.length > 0 && result.choices[0].message) {
      return { content: result.choices[0].message.content, error: '' };
    }
    return { content: null, error: 'Réponse Groq inattendue : ' + response.getContentText().slice(0, 300) };
  } catch (err) {
    return { content: null, error: 'Exception Groq : ' + (err && err.message ? err.message : err.toString()) };
  }
}

function callOpenAiModel(prompt, model, apiKey) {
  try {
    var url = 'https://api.openai.com/v1/chat/completions';
    var body = {
      model: model,
      messages: [
        { role: 'system', content: 'Vous êtes un assistant qui résume des CV en français.' },
        { role: 'user', content: prompt }
      ],
      temperature: 0.2,
      max_tokens: 600
    };
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code !== 200) {
      return { content: null, error: 'OpenAI HTTP ' + code + ' : ' + response.getContentText().slice(0, 300) };
    }
    var result = JSON.parse(response.getContentText());
    if (result.choices && result.choices.length > 0 && result.choices[0].message) {
      return { content: result.choices[0].message.content, error: '' };
    }
    return { content: null, error: 'Réponse OpenAI inattendue : ' + response.getContentText().slice(0, 300) };
  } catch (err) {
    return { content: null, error: 'Exception OpenAI : ' + (err && err.message ? err.message : err.toString()) };
  }
}

function callAiModel(prompt) {
  var cfg = getAiConfig();
  if (cfg.groqKey) {
    return callGroqModel(prompt, cfg.groqKey, cfg.groqModel);
  }
  if (cfg.openaiKey) {
    return callOpenAiModel(prompt, cfg.openaiModel, cfg.openaiKey);
  }
  return { content: null, error: "Aucune clé API IA configurée (propriété de script GROQ_API_KEY ou OPENAI_API_KEY manquante)." };
}

function buildCvAiPrompt(cvText, formData) {
  var prompt = 'Analyse ce CV en français. Fais un résumé structuré des compétences, expériences, formations, et indique la correspondance avec les informations déclarées.';
  if (formData.profession) prompt += '\nProfession déclarée: ' + formData.profession;
  if (formData.secteur) prompt += '\nSecteur déclaré: ' + formData.secteur;
  if (formData.region) prompt += '\nRégion déclarée: ' + formData.region;
  prompt += '\n\nTexte du CV :\n' + cvText;
  prompt += '\n\nRéponds UNIQUEMENT avec un objet JSON strict (pas de texte autour), au format :'
    + '\n{"resume": "court résumé clair et une recommandation simple",'
    + ' "domaine_etude": "domaine d\'étude / de formation détecté dans le CV, chaîne vide si indéterminable",'
    + ' "annees_experience": "nombre d\'années d\'expérience professionnelle détecté, chaîne vide si indéterminable"}';
  return prompt;
}

function runCvAiAnalysis(cvText, formData) {
  if (!cvText) return { summary: '', domaineEtude: '', anneesExperience: '', error: '' };
  var prompt = buildCvAiPrompt(cvText, formData || {});
  var aiResult = callAiModel(prompt);
  if (!aiResult.content) {
    return { summary: '', domaineEtude: '', anneesExperience: '', error: aiResult.error || 'Aucune réponse IA.' };
  }

  var raw = aiResult.content.toString().trim();
  try {
    var jsonMatch = raw.match(/\{[\s\S]*\}/);
    var parsed = JSON.parse(jsonMatch ? jsonMatch[0] : raw);
    return {
      summary: (parsed.resume || '').toString().trim(),
      domaineEtude: (parsed.domaine_etude || '').toString().trim(),
      anneesExperience: (parsed.annees_experience || '').toString().trim(),
      error: ''
    };
  } catch (ex) {
    return { summary: raw, domaineEtude: '', anneesExperience: '', error: '' };
  }
}

function buildAdminAiPrompt(question, cvSummary, formData) {
  var prompt = 'Vous êtes un assistant RH. Réponds en français de façon claire et concise.';
  if (question) prompt += '\nQuestion : ' + question;
  if (cvSummary) prompt += '\nRésumé du CV : ' + cvSummary;
  if (formData.profession) prompt += '\nProfession déclarée : ' + formData.profession;
  if (formData.secteur) prompt += '\nSecteur déclaré : ' + formData.secteur;
  if (formData.region) prompt += '\nRégion déclarée : ' + formData.region;
  prompt += '\n\nDonne une réponse structurée et une recommandation simple.';
  return prompt;
}

function runAdminAiQuestion(question, cvSummary, formData) {
  if (!question) return 'Veuillez saisir une question pour l’IA.';
  var prompt = buildAdminAiPrompt(question, cvSummary || '', formData || {});
  var aiResult = callAiModel(prompt);
  return aiResult.content ? aiResult.content.toString().trim() : (aiResult.error || 'Aucune réponse IA disponible.');
}

function buildMatchPrompt(criteria, candidates) {
  var prompt = 'Vous êtes un assistant de recrutement. Voici une liste de profils enregistrés '
    + 'dans notre base de données interne (fiches et CV soumis). Analysez UNIQUEMENT cette liste : '
    + 'n\'inventez aucune information et ne recherchez rien en dehors des profils fournis ci-dessous.'
    + '\n\nPour le besoin de recrutement décrit, identifiez les profils qui correspondent le mieux. '
    + 'Tenez compte des années d\'expérience minimum demandées, et acceptez les domaines d\'étude '
    + 'proches ou connexes du domaine demandé (pas seulement une correspondance exacte de mot). '
    + 'Pour chaque profil retenu, indiquez : nom complet, e-mail, et une courte justification '
    + '(diplôme/domaine, années d\'expérience, éléments pertinents du CV). '
    + 'Si aucun profil ne correspond parfaitement, ne spéculez JAMAIS sur des domaines ou profils '
    + 'hypothétiques qui ne sont pas dans la liste ci-dessous. À la place, listez concrètement, '
    + 'nom par nom, les profils réellement présents dans la base dont le domaine d\'étude ou '
    + 'l\'expérience se rapprochent le plus (même partiellement) du besoin décrit, en citant leur '
    + 'domaine d\'étude et années d\'expérience réels tels qu\'indiqués ci-dessous. '
    + 'Toute affirmation doit être basée uniquement sur les données fournies, jamais sur des '
    + 'suppositions générales de domaines qui pourraient exister.';

  prompt += '\n\nBesoin de recrutement :\n' + criteria;

  prompt += '\n\nProfils de la base de données interne :\n';
  candidates.forEach(function(c, i) {
    var resume = (c.cvSummary || '').toString().trim();
    if (resume.length > 400) resume = resume.slice(0, 400) + '…';
    prompt += '\n' + (i + 1) + '. Nom: ' + (c.nom || '–')
      + ' | Email: ' + (c.email || '–')
      + ' | Profession: ' + (c.profession || '–')
      + ' | Secteur: ' + (c.secteur || '–')
      + ' | Niveau d\'étude déclaré: ' + (c.niveau || '–')
      + ' | Domaine d\'étude (détecté): ' + (c.cvDomaine || '–')
      + ' | Années d\'expérience (détecté): ' + (c.cvExperience || '–')
      + ' | Résumé CV: ' + (resume || '–');
  });

  return prompt;
}

function runCandidateMatch(criteria, candidates) {
  if (!criteria) return 'Veuillez décrire le besoin de recrutement.';
  if (!candidates || !candidates.length) return 'Aucun profil disponible dans la base de données.';
  var prompt = buildMatchPrompt(criteria, candidates);
  var aiResult = callAiModel(prompt);
  return aiResult.content ? aiResult.content.toString().trim() : (aiResult.error || 'Aucune réponse IA disponible.');
}

// ── Resynchronisation du résumé IA d'une fiche existante ────────────────────
function resyncCvSummary(rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row   = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  var cvUrl      = (row[13] || '').toString();
  var profession = (row[17] || '').toString();
  var secteur    = (row[6]  || '').toString();
  var region     = (row[10] || '').toString();

  if (!cvUrl) {
    return { status: 'error', message: 'Aucun CV enregistré pour cette fiche.' };
  }

  var fileId = extractDriveFileIdFromUrl(cvUrl);
  if (!fileId) {
    return { status: 'error', message: "Impossible d'identifier le fichier Drive à partir du lien CV : " + cvUrl };
  }

  var extraction = extractTextFromDriveFile(fileId);
  if (!extraction.text) {
    return { status: 'error', message: "Impossible d'extraire le texte du CV : " + (extraction.error || 'raison inconnue') };
  }

  var aiResult = runCvAiAnalysis(extraction.text, { profession: profession, secteur: secteur, region: region });
  if (!aiResult.summary) {
    return { status: 'error', message: "L'IA n'a pas pu générer de résumé pour ce CV : " + (aiResult.error || 'raison inconnue') };
  }

  sheet.getRange(rowIndex, 19).setValue(aiResult.summary);
  sheet.getRange(rowIndex, 20).setValue(aiResult.domaineEtude);
  sheet.getRange(rowIndex, 21).setValue(aiResult.anneesExperience);

  return {
    status: 'ok',
    summary: aiResult.summary,
    domaineEtude: aiResult.domaineEtude,
    anneesExperience: aiResult.anneesExperience
  };
}

// ── Test de synchronisation IA sur une seule fiche (avant lot complet) ──────
function syncOneCvSummaryTest() {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var ui      = SpreadsheetApp.getUi();

  for (var row = 2; row <= lastRow; row++) {
    var existingSummary = sheet.getRange(row, 19).getValue();
    if (existingSummary && existingSummary.toString().trim()) continue;

    var cvUrl = sheet.getRange(row, 14).getValue();
    if (!cvUrl || !cvUrl.toString().trim()) continue;

    var nom    = sheet.getRange(row, 2).getValue();
    var result = resyncCvSummary(row);

    if (result.status === 'ok') {
      ui.alert(
        'Test réussi – ligne ' + row + ' (' + nom + ')',
        'Résumé : ' + result.summary
          + '\n\nDomaine d\'étude : ' + (result.domaineEtude || '–')
          + '\nAnnées d\'expérience : ' + (result.anneesExperience || '–')
          + '\n\nSi ce résultat vous convient, lancez « Synchroniser tous les résumés IA » pour traiter le reste.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Échec du test – ligne ' + row + ' (' + nom + ')',
        result.message || 'Erreur inconnue.',
        ui.ButtonSet.OK
      );
    }
    return;
  }

  ui.alert('Test de synchronisation IA', 'Aucune fiche à tester : toutes les fiches ont déjà un résumé ou aucun CV n\'est disponible.', ui.ButtonSet.OK);
}

// ── Synchronisation en lot des résumés IA (fiches sans résumé) ──────────────
var CV_SYNC_TIME_BUDGET_MS = 5 * 60 * 1000; // marge sous la limite d'exécution Apps Script (~6 min)

function runBatchCvSync() {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var start   = new Date().getTime();

  var processed = 0, skippedNoCv = 0, errors = 0, stoppedEarly = false, row;
  var failedRows = [];

  for (row = 2; row <= lastRow; row++) {
    if (new Date().getTime() - start > CV_SYNC_TIME_BUDGET_MS) {
      stoppedEarly = true;
      break;
    }

    var existingSummary = sheet.getRange(row, 19).getValue();
    if (existingSummary && existingSummary.toString().trim()) continue;

    var cvUrl = sheet.getRange(row, 14).getValue();
    if (!cvUrl || !cvUrl.toString().trim()) { skippedNoCv++; continue; }

    var result = resyncCvSummary(row);
    if (result.status === 'ok') {
      processed++;
    } else {
      errors++;
      failedRows.push({
        row: row,
        nom: (sheet.getRange(row, 2).getValue() || '').toString().trim() || 'sans nom',
        message: result.message || 'Erreur inconnue.'
      });
    }
  }

  var remaining = 0;
  var scanFrom  = stoppedEarly ? row : lastRow + 1;
  for (var r = scanFrom; r <= lastRow; r++) {
    var s   = sheet.getRange(r, 19).getValue();
    var url = sheet.getRange(r, 14).getValue();
    if ((!s || !s.toString().trim()) && url && url.toString().trim()) remaining++;
  }

  return {
    processed: processed,
    skippedNoCv: skippedNoCv,
    errors: errors,
    remaining: remaining,
    stoppedEarly: stoppedEarly,
    failedRows: failedRows
  };
}

function syncAllCvSummaries() {
  var r = runBatchCvSync();

  var msg = r.processed + ' fiche(s) synchronisée(s).\n'
    + r.skippedNoCv + ' fiche(s) sans CV ignorée(s).\n'
    + r.errors + ' erreur(s) d\'extraction/analyse.';

  if (r.failedRows.length) {
    msg += '\n\nDétail des erreurs :\n' + r.failedRows.map(function(f) {
      return '- Ligne ' + f.row + ' (' + f.nom + ') : ' + f.message;
    }).join('\n');
  }

  msg += r.stoppedEarly
    ? '\n\n⏱ Limite de temps atteinte. Il reste au moins ' + r.remaining + ' fiche(s) à traiter.'
      + '\nRelancez « Synchroniser tous les résumés IA » pour continuer.'
    : '\n\n✓ Toutes les fiches ont été parcourues.';

  SpreadsheetApp.getUi().alert('Synchronisation des résumés IA', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function doPost(e) {
  try {
    // Lecture JSON (envoyé en text/plain) avec fallback url-encoded
    var p;
    try { p = JSON.parse(e.postData.contents); }
    catch(ex) { p = e.parameter; }

    var action = (p.action || 'submit').toLowerCase();

    if (action === 'update') {
      updateRecord(p);
      return createJsonResponse({ status: 'updated' });
    }

    if (action === 'adminai') {
      if (!verifyAdmin(p.user || '', p.pass || '')) {
        return createJsonResponse({ status: 'unauthorized' });
      }
      var answer = runAdminAiQuestion(p.question || '', p.cvSummary || '', {
        profession: p.profession || '',
        secteur: p.secteur || '',
        region: p.region || ''
      });
      return createJsonResponse({ status: 'ok', answer: answer || 'Aucune réponse IA disponible.' });
    }

    if (action === 'matchcandidates') {
      if (!verifyAdmin(p.user || '', p.pass || '')) {
        return createJsonResponse({ status: 'unauthorized' });
      }
      var matchResult = runCandidateMatch(p.criteria || '', p.candidates || []);
      return createJsonResponse({ status: 'ok', result: matchResult });
    }

    if (action === 'resynccv') {
      if (!verifyAdmin(p.user || '', p.pass || '')) {
        return createJsonResponse({ status: 'unauthorized' });
      }
      var rowIndex = parseInt(p.rowIndex, 10);
      if (!rowIndex) return createJsonResponse({ status: 'error', message: 'rowIndex manquant' });
      return createJsonResponse(resyncCvSummary(rowIndex));
    }

    if (action === 'resyncbatch') {
      if (!verifyAdmin(p.user || '', p.pass || '')) {
        return createJsonResponse({ status: 'unauthorized' });
      }
      var batchResult = runBatchCvSync();
      return createJsonResponse({
        status: 'ok',
        processed: batchResult.processed,
        skippedNoCv: batchResult.skippedNoCv,
        errors: batchResult.errors,
        remaining: batchResult.remaining,
        stoppedEarly: batchResult.stoppedEarly,
        failedRows: batchResult.failedRows
      });
    }

    // ── Enrichissement IA CV avant enregistrement ─────────────────────────
    if (p.cvText) {
      try {
        var aiResult = runCvAiAnalysis(p.cvText, {
          profession: p.profession || '',
          secteur: p.secteur || '',
          region: p.region || ''
        });
        if (aiResult.summary) {
          p.cvSummary = aiResult.summary;
        }
        p.cvDomaine    = aiResult.domaineEtude || '';
        p.cvExperience = aiResult.anneesExperience || '';
      } catch (ex) {
        // Ne pas bloquer l'enregistrement si l'IA échoue.
      }
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
      p.profession  || "", // Col 18 : Profession
      p.cvSummary   || "", // Col 19 : Résumé automatique extrait du CV
      p.cvDomaine   || "", // Col 20 : Domaine d'étude détecté (IA)
      p.cvExperience|| ""  // Col 21 : Années d'expérience détectées (IA)
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

    return createJsonResponse({ status: 'success' });

  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
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
    ["Résumé automatique CV",p.cvSummary   || "–"],
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
  return createJsonResponse(ok
    ? { status: 'ok' }
    : { status: 'error', message: 'Identifiants incorrects' });
}

function getAdminData(u, p) {
  if (!verifyAdmin(u, p)) return createJsonResponse({ status: 'unauthorized' });
  var vals = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();
  return createJsonResponse({ status: 'ok', data: vals });
}

function listAdminUsers(u, p) {
  if (!verifyAdmin(u, p)) return createJsonResponse({ status: 'unauthorized' });
  var names = getAdmins().map(function(a) { return a.user; });
  return createJsonResponse({ status: 'ok', admins: names });
}

function addAdminUser(u, p, nu, np) {
  if (!verifyAdmin(u, p)) return createJsonResponse({ status: 'unauthorized' });
  if (!nu || !np) return createJsonResponse({ status: 'error', message: 'Champs manquants' });
  var admins = getAdmins();
  if (admins.some(function(a) { return a.user === nu; })) return createJsonResponse({ status: 'error', message: 'Identifiant déjà existant' });
  admins.push({ user: nu, passHash: hashPass(np) });
  PropertiesService.getScriptProperties().setProperty(ADMIN_KEY, JSON.stringify(admins));
  return createJsonResponse({ status: 'ok' });
}

// ── Diagnostic : la clé Groq est-elle bien enregistrée ? ────────────────────
function debugCheckGroqKey() {
  var ui  = SpreadsheetApp.getUi();
  var cfg = getAiConfig();

  if (!cfg.groqKey && !cfg.openaiKey) {
    ui.alert(
      'Diagnostic clé IA',
      "Aucune clé trouvée.\n\nAllez dans l'éditeur Apps Script → ⚙ Paramètres du projet → Propriétés du script,"
        + " et ajoutez une propriété nommée exactement GROQ_API_KEY avec votre clé Groq comme valeur.",
      ui.ButtonSet.OK
    );
    return;
  }

  if (cfg.groqKey) {
    var k = cfg.groqKey;
    var masked = k.length > 8 ? (k.slice(0, 4) + '…' + k.slice(-4)) : '(valeur très courte, suspecte)';
    ui.alert(
      'Diagnostic clé IA',
      'Propriété GROQ_API_KEY trouvée : ' + masked + ' (' + k.length + ' caractères)'
        + '\nModèle configuré : ' + cfg.groqModel
        + (k !== k.trim() ? '\n\n⚠ La valeur contient des espaces en début/fin — corrigez-la dans Script Properties.' : ''),
      ui.ButtonSet.OK
    );
    return;
  }

  ui.alert('Diagnostic clé IA', 'Pas de clé Groq, mais une clé OpenAI est configurée (modèle : ' + cfg.openaiModel + ').', ui.ButtonSet.OK);
}

// ── Menu Google Sheets ────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Administration IFL')
    .addItem('Créer un administrateur', 'showCreateAdminDialog')
    .addSeparator()
    .addItem('Liste des administrateurs', 'showAdminsList')
    .addSeparator()
    .addItem('Diagnostiquer la clé IA', 'debugCheckGroqKey')
    .addItem('Tester la synchronisation IA (1 fiche)', 'syncOneCvSummaryTest')
    .addItem('Synchroniser tous les résumés IA', 'syncAllCvSummaries')
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
