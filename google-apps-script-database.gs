/*
  SCRIPT GOOGLE SHEETS — Base de données centralisée TUNARCO
  ============================================================
  Reçoit les inscriptions envoyées par le site et les répartit
  CHACUNE DANS SON PROPRE ONGLET, un onglet par programme :
    - "Génies Lab"              (formulaire principal)
    - "Compétitions"
    - "Summer Camp"
    - "Master Classe Vacances"
    - "Formation Personnalisée"
    - "Options"                 (choix facultatifs après une inscription
                                  Génies Lab)

  PANNEAU ADMIN "SANS CODE" — onglet "Planning Compétitions" :
  Ce n'est PAS l'onglet "Compétitions" ci-dessus (qui reçoit les
  inscriptions des visiteurs). "Planning Compétitions" est un onglet
  à part que VOUS éditez vous-même pour changer les dates, lieux,
  descriptions des compétitions affichées sur le site — sans toucher
  au code ni à GitHub. Le site va chercher ces données automatiquement
  à chaque visite (fonction doGet ci-dessous, appelée en GET depuis
  planning-competitions.html). Il se crée tout seul avec 5 lignes
  d'exemple au premier appel si l'onglet n'existe pas encore.
  Colonnes : Ordre | Jour | Mois | Type date (normal/tbd/own) | Titre |
  Texte badge | Type badge (limited/tbd/own/none) | Lieu | Description
  "Type date" contrôle l'affichage de la pastille de date :
    normal = date confirmée en couleur, tbd = "?" grisé, own = mis en
    avant (notre événement).
  "Type badge" contrôle la couleur du badge à côté du titre.
  Pour ajouter une compétition : ajoutez une ligne. Pour la retirer :
  supprimez la ligne (ou videz la colonne Titre). L'ordre d'affichage
  suit la colonne "Ordre" (nombres croissants).
  Chaque onglet a son propre rendu "tableau professionnel" :
    - une bannière-titre fusionnée en haut (ligne 1)
    - un en-tête de colonnes bleu marine, texte blanc, figé (ligne 2)
    - une colonne "N°" numérotée automatiquement
    - des bandes alternées bleu très clair / blanc sur les lignes de
      données, pour la lisibilité (comme un tableau Excel classique)
    - la cellule "Confirmé" elle-même colorée en rouge (Non) ou vert
      (Oui), pour repérer d'un coup d'œil qui n'a pas encore confirmé
    - bordures sur tout le tableau
    - la photo d'identité et l'extrait de naissance sont envoyés en
      base64 par le site, décodés ici, et envoyés directement par EMAIL
      (en pièce jointe) à ADMIN_ALERT_EMAIL — plus besoin de Google
      Drive. CES FICHIERS N'APPARAISSENT PAS DANS LE SHEET — le Sheet
      ne contient que les données d'inscription, les photos/documents
      arrivent uniquement par email.

  MISE À JOUR (si vous aviez déjà installé une version précédente) :
  1. Ouvre ton Google Sheet → Extensions → Apps Script.
  2. Sélectionne tout le code existant et efface-le.
  3. Colle TOUT le contenu de ce fichier à la place.
  4. Enregistre (icône disquette).
  5. Déployer → Gérer les déploiements → icône crayon (Edit) sur le
     déploiement existant → Version : "Nouvelle version" → Déployer.
     (L'URL /exec reste identique, rien à changer côté site.)
  6. IMPORTANT : cette version sépare les inscriptions par programme
     dans des onglets distincts (avant : tout était regroupé dans un
     seul onglet "Inscriptions"). Vous pouvez supprimer l'ancien
     onglet "Inscriptions" — les nouveaux onglets se créent tout
     seuls, proprement, au premier envoi de chaque type.

  INSTALLATION DE ZÉRO :
  Voir les instructions données précédemment (créer un Sheet, coller
  ce code, Déployer → Nouveau déploiement → Application Web →
  Exécuter en tant que Moi / Qui a accès : Tout le monde).
*/

// Email qui reçoit la photo d'identité + l'extrait de naissance en pièce
// jointe à chaque inscription (plus besoin de Google Drive).
const ADMIN_ALERT_EMAIL = "info.tunarco@gmail.com";
const NAVY = "#1a1035";
const BANDING_BLUE = "#eaf1fb";
const RED = "#fde2e2";
const GREEN = "#dcfce7";
const BORDER_COLOR = "#c7ccd8";

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;

  if (data.type === "options") {
    // Choix d'options facultatifs envoyés depuis la popup après une
    // inscription Génies Lab.
    sheet = getOrCreateSheet(
      ss, "Options", "TUNARCO — Options choisies",
      ["N°", "Date", "Élève", "Email", "Opt. Compétiteur", "Kit Programmation", "Kit Électronique", "Pull officiel", "Taille pull", "Mode paiement", "Total (DT)"]
    );
    sheet.appendRow([
      "=ROW()-2",
      data.date || "", data.nomEleve || "", data.email || "",
      data.optCompetiteur || "Non", data.optKitProg || "Non",
      data.optKitElec || "Non", data.optPull || "Non", data.pullTaille || "",
      data.modePaiement || "", data.total || ""
    ]);
    applyRowStyling(sheet);
    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (data.type === "competition") {
    sheet = getOrCreateSheet(
      ss, "Compétitions", "TUNARCO — Compétitions",
      ["N°", "Date", "Compétition", "Élève", "Téléphone", "Confirmé"]
    );
    sheet.appendRow([
      "=ROW()-2", data.date || "", data.competition || "",
      data.nomEleve || "", data.tel || "", "Non"
    ]);
  } else if (data.type === "summer-camp") {
    sheet = getOrCreateSheet(
      ss, "Summer Camp", "TUNARCO — Summer Camp",
      ["N°", "Date", "Élève", "Âge", "Naissance", "Ancien/Nouveau", "Parents", "Email", "Tél 1", "Tél 2", "Semaine", "Confirmé"]
    );
    sendDocumentsByEmail(
      data.photoIdentiteData, data.photoIdentite,
      data.extraitNaissanceData, data.extraitNaissance,
      data.nomEleve, "Summer Camp"
    );
    sheet.appendRow([
      "=ROW()-2", data.date || "", data.nomEleve || "", data.age || "",
      data.dateNaissance || "", data.statut || "", data.nomParents || "",
      data.email || "", data.tel1 || "", data.tel2 || "", data.semaine || "", "Non"
    ]);
  } else if (data.type === "master-classe") {
    sheet = getOrCreateSheet(
      ss, "Master Classe Vacances", "TUNARCO — Master Classe Vacances",
      ["N°", "Date", "Élève", "Âge", "Naissance", "Ancien/Nouveau", "Parents", "Email", "Tél 1", "Tél 2", "Période", "Confirmé"]
    );
    sendDocumentsByEmail(
      data.photoIdentiteData, data.photoIdentite,
      data.extraitNaissanceData, data.extraitNaissance,
      data.nomEleve, "Master Classe Vacances"
    );
    sheet.appendRow([
      "=ROW()-2", data.date || "", data.nomEleve || "", data.age || "",
      data.dateNaissance || "", data.statut || "", data.nomParents || "",
      data.email || "", data.tel1 || "", data.tel2 || "", data.periode || "", "Non"
    ]);
  } else if (data.type === "formation-personnalisee") {
    sheet = getOrCreateSheet(
      ss, "Formation Personnalisée", "TUNARCO — Formation Personnalisée",
      ["N°", "Date", "Élève", "Âge", "Naissance", "Ancien/Nouveau", "Parents", "Email", "Tél 1", "Tél 2", "Domaine souhaité", "Confirmé"]
    );
    sheet.appendRow([
      "=ROW()-2", data.date || "", data.nomEleve || "", data.age || "",
      data.dateNaissance || "", data.statut || "", data.nomParents || "",
      data.email || "", data.tel1 || "", data.tel2 || "", data.domaine || "", "Non"
    ]);
  } else {
    // Inscription Génies Lab (formulaire principal)
    sheet = getOrCreateSheet(
      ss, "Génies Lab", "TUNARCO — Génies Lab",
      ["N°", "Date", "Élève", "Naissance", "Ancien/Nouveau", "Parents", "Email", "Tél 1", "Tél 2", "Jour", "Horaire", "Compétition (accès)", "Droit à l'image", "Âge", "Confirmé"]
    );
    sendDocumentsByEmail(
      data.photoIdentiteData, data.photoIdentite,
      data.extraitNaissanceData, data.extraitNaissance,
      data.nomEleve, "Génies Lab"
    );
    sheet.appendRow([
      "=ROW()-2", data.dateInscription || "", data.nomEleve || "",
      data.dateNaissance || "", data.ancienAdherent || "", data.nomParents || "",
      data.email || "", data.tel1 || "", data.tel2 || "", data.jour || "",
      data.horaire || "", data.competition || "", data.droitImage || "", data.age || "", "Non"
    ]);
  }

  applyRowStyling(sheet);

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------
// PANNEAU ADMIN "SANS CODE" — le site interroge cette route en GET pour
// afficher les compétitions à jour, sans qu'on ait besoin de modifier le
// code ni de re-publier sur GitHub : il suffit d'éditer l'onglet
// "Planning Compétitions" dans ce Google Sheet.
// ---------------------------------------------------------------------
function doGet(e) {
  const type = (e && e.parameter && e.parameter.type) || "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (type === "planning-competitions") {
    return jsonResponse(getPlanningCompetitions(ss));
  }

  return jsonResponse({ error: "type inconnu" });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Lit l'onglet "Planning Compétitions" et renvoie un tableau d'objets
// prêt à être affiché par planning-competitions.html. Crée l'onglet avec
// 5 lignes d'exemple (les compétitions déjà sur le site) s'il n'existe
// pas encore, pour que l'admin voie tout de suite le format à respecter.
function getPlanningCompetitions(ss) {
  let sheet = ss.getSheetByName("Planning Compétitions");
  // Si l'onglet n'existe pas, OU existe mais n'a aucune ligne de données
  // (moins de 3 lignes = juste la bannière + l'en-tête, sans lignes de
  // compétition en dessous — par exemple si sa création a été interrompue
  // avant d'ajouter les 5 lignes d'exemple), on (re)crée/complète avec les
  // lignes d'exemple plutôt que de renvoyer une liste vide.
  if (!sheet || sheet.getLastRow() < 3) sheet = createPlanningCompetitionsSheet(ss);

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const values = sheet.getRange(3, 1, lastRow - 2, 9).getValues();
  const list = [];
  values.forEach(row => {
    const [ordre, jour, mois, dateType, titre, badgeText, badgeType, lieu, description] = row;
    if (!titre) return; // ligne vide ou supprimée par l'admin
    list.push({
      ordre: ordre || 999,
      jour: String(jour || ""),
      mois: String(mois || ""),
      dateType: String(dateType || "normal"),
      titre: String(titre || ""),
      badgeText: String(badgeText || ""),
      badgeType: String(badgeType || "none"),
      lieu: String(lieu || ""),
      description: String(description || "")
    });
  });
  list.sort((a, b) => a.ordre - b.ordre);
  return list;
}

function createPlanningCompetitionsSheet(ss) {
  const headers = ["Ordre", "Jour", "Mois", "Type date (normal/tbd/own)", "Titre", "Texte badge", "Type badge (limited/tbd/own/none)", "Lieu", "Description"];
  const sheet = getOrCreateSheet(ss, "Planning Compétitions", "TUNARCO — Planning Compétitions (admin)", headers);

  // Lignes d'exemple = les compétitions déjà présentes sur le site, pour
  // que rien ne disparaisse au premier chargement et que l'admin voie le
  // format attendu directement.
  const seed = [
    [1, "11", "Oct 2026", "normal", "RoboCup", "Places limitées", "limited", "ENSI, Manouba", "Compétition nationale de robotique — début d'année de formation, ouverte aux niveaux préparés."],
    [2, "?", "2026-27", "tbd", "Eurobot — Qualifications tunisiennes", "Date à confirmer", "tbd", "À confirmer", "Étape de qualification nationale pour la compétition internationale Eurobot."],
    [3, "?", "2026-27", "tbd", "MRC — Hammamet", "Date à confirmer", "tbd", "Hammamet", "Minoan Robotsports Competition — planning pas encore finalisé."],
    [4, "?", "Juin 2027", "own", "Festival de Robotique", "Notre événement", "own", "Les Petits Génies de la Tunisie", "Notre propre compétition, organisée par TUNARCO en fin d'année de formation — l'occasion pour tous nos adhérents de présenter leurs projets."],
    [5, "?", "2026-27", "tbd", "Compétition supplémentaire", "À venir", "tbd", "À confirmer", "Emplacement réservé pour une autre compétition prévue cette année."]
  ];
  seed.forEach(row => sheet.appendRow(row));
  sheet.autoResizeColumns(1, headers.length);
  applyBandingOnly(sheet); // pas de colonne "Confirmé" ici, donc pas applyRowStyling()
  return sheet;
}

// Version allégée de applyRowStyling() pour les onglets qui n'ont pas de
// colonne "Confirmé" (comme "Planning Compétitions") : bandes alternées
// + bordures, sans le surlignage rouge/vert.
function applyBandingOnly(sheet) {
  const lastRow = Math.max(sheet.getMaxRows(), 100);
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const dataRange = sheet.getRange(3, 1, lastRow - 2, lastCol);

  const bandingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=ISEVEN(ROW())`)
    .setBackground(BANDING_BLUE)
    .setRanges([dataRange])
    .build();
  sheet.setConditionalFormatRules([bandingRule]);

  const actualLastRow = sheet.getLastRow();
  if (actualLastRow >= 1) {
    sheet.getRange(1, 1, actualLastRow, lastCol)
      .setBorder(true, true, true, true, true, true, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  }
  if (actualLastRow > 2) {
    sheet.getRange(3, 1, actualLastRow - 2, 1).setHorizontalAlignment("center");
  }
}

// Décode un fichier envoyé en base64 (data URL) et le transforme en Blob
// Apps Script (utilisable comme pièce jointe email). Renvoie null si rien
// à envoyer ou si le format n'est pas reconnu.
function dataUrlToBlob(dataUrl, fileName) {
  if (!dataUrl || typeof dataUrl !== "string") return null;
  const match = dataUrl.match(/^data:(.*?);base64,(.*)$/);
  if (!match) return null;
  const mimeType = match[1] || "application/octet-stream";
  const bytes = Utilities.base64Decode(match[2]);
  return Utilities.newBlob(bytes, mimeType, fileName || "document");
}

// Envoie la photo d'identité + l'extrait de naissance en pièces jointes
// d'un seul email à ADMIN_ALERT_EMAIL. N'écrit rien dans le Sheet — les
// documents arrivent uniquement par email. Appelée pour Génies Lab,
// Summer Camp et Master Classe Vacances, dès que l'un des deux fichiers
// a été fourni (Compétition et Formation Personnalisée n'ont pas ces
// champs).
function sendDocumentsByEmail(photoDataUrl, photoName, extraitDataUrl, extraitName, studentName, programme) {
  const attachments = [];
  const photoBlob = dataUrlToBlob(photoDataUrl, photoName);
  if (photoBlob) attachments.push(photoBlob);
  const extraitBlob = dataUrlToBlob(extraitDataUrl, extraitName);
  if (extraitBlob) attachments.push(extraitBlob);

  if (attachments.length === 0) return;

  try {
    MailApp.sendEmail({
      to: ADMIN_ALERT_EMAIL,
      subject: "📎 Documents d'inscription — " + (programme ? programme + " — " : "") + (studentName || "élève"),
      body: "Ci-joint la photo d'identité et/ou l'extrait de naissance pour l'inscription " +
        (programme || "") + " de " + (studentName || "") + ".",
      attachments: attachments
    });
  } catch (err) {
    console.error("Échec envoi email des documents: " + err.message);
  }
}

// Crée la feuille avec : bannière-titre fusionnée (ligne 1) +
// en-tête de colonnes (ligne 2), figés, largeurs ajustées.
function getOrCreateSheet(ss, name, title, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    const lastCol = headers.length;

    // Ligne 1 — bannière-titre fusionnée
    sheet.getRange(1, 1, 1, lastCol).merge();
    const banner = sheet.getRange(1, 1);
    banner.setValue(title);
    banner.setBackground(NAVY);
    banner.setFontColor("#ffffff");
    banner.setFontWeight("bold");
    banner.setFontSize(13);
    banner.setHorizontalAlignment("center");
    banner.setVerticalAlignment("middle");
    sheet.setRowHeight(1, 34);

    // Ligne 2 — en-tête des colonnes
    sheet.getRange(2, 1, 1, lastCol).setValues([headers]);
    const headerRange = sheet.getRange(2, 1, 1, lastCol);
    headerRange.setBackground(NAVY);
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");
    headerRange.setWrap(true);
    sheet.setRowHeight(2, 34);

    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(2); // N° + Élève toujours visibles au scroll
    sheet.autoResizeColumns(1, lastCol);
    sheet.setColumnWidth(1, 45); // colonne N° plus étroite
  }
  return sheet;
}

// Applique bandes alternées + couleur de la cellule "Confirmé" + bordures.
function applyRowStyling(sheet) {
  const lastRow = Math.max(sheet.getMaxRows(), 300);
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const dataRange = sheet.getRange(3, 1, lastRow - 2, lastCol); // les données commencent ligne 3

  const bandingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=ISEVEN(ROW())`)
    .setBackground(BANDING_BLUE)
    .setRanges([dataRange])
    .build();

  // Couleur uniquement la cellule "Confirmé" (pas toute la ligne), pour ne
  // pas masquer les bandes alternées.
  const confirmRange = sheet.getRange(3, lastCol, lastRow - 2, 1);
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Non")
    .setBackground(RED)
    .setBold(true)
    .setRanges([confirmRange])
    .build();
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Oui")
    .setBackground(GREEN)
    .setBold(true)
    .setRanges([confirmRange])
    .build();

  // Ordre = priorité : rouge/vert (cellule) d'abord, bandes ensuite.
  sheet.setConditionalFormatRules([redRule, greenRule, bandingRule]);

  const actualLastRow = sheet.getLastRow();
  if (actualLastRow >= 1) {
    sheet.getRange(1, 1, actualLastRow, lastCol)
      .setBorder(true, true, true, true, true, true, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  }
  if (actualLastRow > 2) {
    sheet.getRange(3, 1, actualLastRow - 2, 1).setHorizontalAlignment("center"); // centrer la colonne N°
  }
}
