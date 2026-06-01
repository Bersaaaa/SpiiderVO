/**
 * ============================================================
 * SPIDERVO PRO — Système de Gestion VO pour Négociant Automobile
 * Version: 3.0 | Google Apps Script + Sheets
 * Compatible: Google Workspace, Drive, Gmail, WhatsApp (via API)
 * ============================================================
 * INSTRUCTIONS D'INSTALLATION :
 *  1. Ouvrir Google Sheets (nouveau fichier)
 *  2. Extensions > Apps Script
 *  3. Coller TOUT ce code dans l'éditeur
 *  4. Lancer : Menu SpiderVO > Initialiser le système
 *  5. Autoriser les permissions demandées
 * ============================================================
 */

// ============================================================
// CONSTANTES GLOBALES
// ============================================================
const CONFIG = {
  VERSION: '3.0',
  SPREADSHEET_NAME: 'SpiderVO Pro — Gestion VO',
  DOSSIER_PREFIX: 'VO-',
  CLIENT_PREFIX: 'CL-',
  COMMANDE_PREFIX: 'BC-',
  FACTURE_PREFIX: 'FA-',

  // Couleurs de la charte graphique (hex)
  COLORS: {
    PRIMARY: '#0A0F1E',       // Bleu nuit
    SECONDARY: '#1A2744',     // Bleu marine
    ACCENT: '#E63946',        // Rouge SpiderVO
    ACCENT2: '#F4A261',       // Orange chaud
    SUCCESS: '#2DC653',       // Vert succès
    WARNING: '#FFB703',       // Orange alerte
    DANGER: '#E63946',        // Rouge danger
    INFO: '#4895EF',          // Bleu info
    GOLD: '#F4D03F',          // Or
    WHITE: '#FFFFFF',
    LIGHT_BG: '#F8F9FC',
    DARK_TEXT: '#0A0F1E',
    MID_GRAY: '#6B7280',
    BORDER: '#E5E7EB',
    HEADER_BG: '#0A0F1E',
    ROW_ALT: '#F1F5FF',
  },

  // Statuts Stock
  STATUTS_STOCK: ['Disponible', 'Préparation', 'Publié', 'Réservé', 'Vendu', 'Livré', 'SAV'],
  STATUTS_COULEURS: {
    'Disponible': '#2DC653',
    'Préparation': '#F4A261',
    'Publié': '#4895EF',
    'Réservé': '#FFB703',
    'Vendu': '#E63946',
    'Livré': '#6D28D9',
    'SAV': '#78716C',
  },

  // Pipeline CRM
  PIPELINE_ETAPES: ['Nouveau lead', 'Contacté', 'RDV planifié', 'Essai', 'Négociation', 'Réservé', 'Vendu', 'Perdu'],

  // Carburants
  CARBURANTS: ['Essence', 'Diesel', 'Hybride', 'Hybride rechargeable', 'Électrique', 'GPL', 'GNV'],

  // Boîtes
  BOITES: ['Manuelle', 'Automatique', 'Semi-automatique', 'CVT'],

  // Provenances
  PROVENANCES: ['Particulier', 'Reprise', 'Enchères BCA', 'Enchères SVI', 'Mandataire', 'Concession', 'Leaser', 'Import', 'Centrale achat'],

  // Crit\'Air
  CRITAIR: ['Crit\'Air 0 (Électrique)', 'Crit\'Air 1', 'Crit\'Air 2', 'Crit\'Air 3', 'Crit\'Air 4', 'Crit\'Air 5', 'Non classé'],

  // TVA
  TVA_TAUX: 0.20,

  // Seuil véhicule dormant (jours)
  SEUIL_DORMANT: 45,
};

// ============================================================
// MENU PRINCIPAL
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚗 SpiderVO Pro')
    .addItem('🏠 Dashboard', 'allerAuDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('📦 Stock VO')
      .addItem('➕ Ajouter un véhicule', 'ouvrirFormulaireVehicule')
      .addItem('🔍 Recherche avancée', 'rechercheAvancee')
      .addItem('📊 Mise à jour stock', 'majStock')
      .addItem('⚠️ Véhicules dormants', 'voirVehiculesDormants')
      .addItem('🖨️ Fiche véhicule PDF', 'genererFicheVehicule'))
    .addSeparator()
    .addSubMenu(ui.createMenu('👥 CRM Clients')
      .addItem('➕ Nouveau client/lead', 'ouvrirFormulaireClient')
      .addItem('🔥 Relances chauds', 'voirLeadsChauds')
      .addItem('📧 Email de relance', 'envoyerEmailRelance')
      .addItem('💬 Message WhatsApp', 'genererMessageWhatsApp'))
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 Ventes')
      .addItem('📋 Bon de commande', 'genererBonCommande')
      .addItem('🧾 Facture', 'genererFacture')
      .addItem('🔄 Saisir une reprise', 'saisirReprise')
      .addItem('📈 Rapport mensuel', 'genererRapportMensuel'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📄 Administratif')
      .addItem('📋 Tableau de bord CERFA', 'voirCERFA')
      .addItem('⚠️ Dossiers incomplets', 'voirDossiersIncomplets')
      .addItem('🔔 Alertes CT / Garanties', 'voirAlertesCT'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Système')
      .addItem('🚀 Initialiser le système', 'initialiserSysteme')
      .addItem('🔄 Recalculer tout', 'recalculerTout')
      .addItem('💾 Sauvegarder sur Drive', 'sauvegarderDrive')
      .addItem('📧 Test email', 'testEmail')
      .addItem('ℹ️ À propos', 'aPropos'))
    .addToUi();
}

// ============================================================
// INITIALISATION COMPLÈTE DU SYSTÈME
// ============================================================
function initialiserSysteme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const reponse = ui.alert(
    '🚗 SpiderVO Pro — Initialisation',
    'Cette opération va créer tous les onglets et configurer le système.\n\nDurée estimée : 2-3 minutes.\n\nContinuer ?',
    ui.ButtonSet.YES_NO
  );

  if (reponse !== ui.Button.YES) return;

  // Renommer le spreadsheet
  ss.setName(CONFIG.SPREADSHEET_NAME);

  // Créer les onglets dans l'ordre
  creerOngletDashboard(ss);
  creerOngletStockVO(ss);
  creerOngletCRM(ss);
  creerOngletVentes(ss);
  creerOngletAdministratif(ss);
  creerOngletFinance(ss);
  creerOngletParametres(ss);

  // Supprimer la feuille par défaut si elle existe encore
  try {
    const feuillePardDefaut = ss.getSheetByName('Feuille 1') || ss.getSheetByName('Sheet1');
    if (feuillePardDefaut) ss.deleteSheet(feuillePardDefaut);
  } catch(e) {}

  // Configurer les déclencheurs automatiques
  configurerDeclencheurs();

  // Mise à jour initiale du dashboard
  majDashboard();

  // Aller au dashboard
  ss.setActiveSheet(ss.getSheetByName('DASHBOARD'));

  ui.alert(
    '✅ Installation réussie !',
    'SpiderVO Pro est prêt à l\'utilisation.\n\nUtilisez le menu 🚗 SpiderVO Pro pour naviguer.\n\nBonne gestion !',
    ui.ButtonSet.OK
  );
}

// ============================================================
// ONGLET DASHBOARD
// ============================================================
function creerOngletDashboard(ss) {
  let sheet = ss.getSheetByName('DASHBOARD');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('DASHBOARD', 0);

  sheet.setTabColor(CONFIG.COLORS.ACCENT);

  // Masquer les lignes/colonnes de grille
  sheet.setHiddenGridlines(true);

  // Définir les dimensions
  sheet.setColumnWidth(1, 20);    // Marge gauche
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 20);    // Marge droite

  for (let i = 1; i <= 60; i++) sheet.setRowHeight(i, 30);
  sheet.setRowHeight(1, 10);
  sheet.setRowHeight(2, 70);
  sheet.setRowHeight(3, 10);

  // ---- HEADER ----
  const rangeHeader = sheet.getRange('B2:G2');
  rangeHeader.merge();
  rangeHeader.setValue('🚗  SPIDERVO PRO  —  Gestion de Véhicules d\'Occasion');
  rangeHeader.setBackground(CONFIG.COLORS.PRIMARY);
  rangeHeader.setFontColor(CONFIG.COLORS.WHITE);
  rangeHeader.setFontSize(22);
  rangeHeader.setFontWeight('bold');
  rangeHeader.setHorizontalAlignment('center');
  rangeHeader.setVerticalAlignment('middle');

  // ---- SOUS-TITRE ----
  sheet.setRowHeight(4, 25);
  const rangeSub = sheet.getRange('B4:G4');
  rangeSub.merge();
  const now = new Date();
  rangeSub.setValue(`Tableau de bord — Mis à jour le ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy à HH:mm')}`);
  rangeSub.setBackground(CONFIG.COLORS.SECONDARY);
  rangeSub.setFontColor('#94A3B8');
  rangeSub.setFontSize(11);
  rangeSub.setHorizontalAlignment('center');
  rangeSub.setVerticalAlignment('middle');

  sheet.setRowHeight(5, 15);

  // ---- KPI ROW 1 : Stock ----
  const kpiRow1Start = 6;
  const kpiTitres1 = ['📦 VÉHICULES EN STOCK', '💰 VALEUR TOTALE STOCK', '📈 MARGE POTENTIELLE', '🏷️ PRIX MOYEN STOCK'];
  const kpiFormules1 = [
    '=COUNTIF(\'STOCK VO\'!L:L,"<>Vendu")-COUNTIF(\'STOCK VO\'!L:L,"<>Livré")-COUNTIF(\'STOCK VO\'!L:L,"")+COUNTIF(\'STOCK VO\'!L2:L5000,"Disponible")+COUNTIF(\'STOCK VO\'!L2:L5000,"Préparation")+COUNTIF(\'STOCK VO\'!L2:L5000,"Publié")+COUNTIF(\'STOCK VO\'!L2:L5000,"Réservé")',
    '=SUMPRODUCT((\'STOCK VO\'!L2:L5000<>"Vendu")*(\'STOCK VO\'!L2:L5000<>"Livré")*(\'STOCK VO\'!L2:L5000<>"")*(\'STOCK VO\'!AE2:AE5000))',
    '=SUMPRODUCT((\'STOCK VO\'!L2:L5000<>"Vendu")*(\'STOCK VO\'!L2:L5000<>"Livré")*(\'STOCK VO\'!L2:L5000<>"")*(\'STOCK VO\'!AH2:AH5000))',
    '=IFERROR(AVERAGEIF(\'STOCK VO\'!L2:L5000,"Publié",\'STOCK VO\'!AE2:AE5000),0)',
  ];
  const kpiFormats1 = ['#,##0', '#,##0 "€"', '#,##0 "€"', '#,##0 "€"'];
  const kpiColors1 = [CONFIG.COLORS.INFO, CONFIG.COLORS.SUCCESS, CONFIG.COLORS.ACCENT2, CONFIG.COLORS.GOLD];

  creerCarteKPI(sheet, kpiRow1Start, kpiTitres1, kpiFormules1, kpiFormats1, kpiColors1);

  sheet.setRowHeight(kpiRow1Start + 5, 15);

  // ---- KPI ROW 2 : Ventes ----
  const kpiRow2Start = kpiRow1Start + 6;
  const kpiTitres2 = ['🛒 VENDUS CE MOIS', '💵 CA MENSUEL', '💎 MARGE NETTE MOIS', '🔥 LEADS ACTIFS'];
  const kpiFormules2 = [
    '=COUNTIFS(\'VENTES\'!E2:E5000,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'VENTES\'!E2:E5000,"<="&EOMONTH(TODAY(),0))',
    '=SUMPRODUCT((\'VENTES\'!E2:E5000>=DATE(YEAR(TODAY()),MONTH(TODAY()),1))*(\'VENTES\'!E2:E5000<=EOMONTH(TODAY(),0))*(\'VENTES\'!J2:J5000))',
    '=SUMPRODUCT((\'VENTES\'!E2:E5000>=DATE(YEAR(TODAY()),MONTH(TODAY()),1))*(\'VENTES\'!E2:E5000<=EOMONTH(TODAY(),0))*(\'VENTES\'!M2:M5000))',
    '=COUNTIF(\'CRM CLIENTS\'!H2:H5000,"Nouveau lead")+COUNTIF(\'CRM CLIENTS\'!H2:H5000,"Contacté")+COUNTIF(\'CRM CLIENTS\'!H2:H5000,"RDV planifié")+COUNTIF(\'CRM CLIENTS\'!H2:H5000,"Essai")+COUNTIF(\'CRM CLIENTS\'!H2:H5000,"Négociation")',
  ];
  const kpiFormats2 = ['#,##0', '#,##0 "€"', '#,##0 "€"', '#,##0'];
  const kpiColors2 = [CONFIG.COLORS.ACCENT, CONFIG.COLORS.SUCCESS, CONFIG.COLORS.GOLD, CONFIG.COLORS.WARNING];

  creerCarteKPI(sheet, kpiRow2Start, kpiTitres2, kpiFormules2, kpiFormats2, kpiColors2);

  sheet.setRowHeight(kpiRow2Start + 5, 15);

  // ---- KPI ROW 3 : Alertes ----
  const kpiRow3Start = kpiRow2Start + 6;
  const kpiTitres3 = ['⏳ VÉHICULES DORMANTS', '📋 DOSSIERS INCOMPLETS', '🔧 EN PRÉPARATION', '⚠️ CT EXPIRANT'];
  const kpiFormules3 = [
    `=COUNTIFS(\'STOCK VO\'!L2:L5000,"<>Vendu",\'STOCK VO\'!L2:L5000,"<>Livré",\'STOCK VO\'!L2:L5000,"<>",\'STOCK VO\'!P2:P5000,"<"&TODAY()-${CONFIG.SEUIL_DORMANT})`,
    '=COUNTIF(\'ADMINISTRATIF\'!P2:P5000,"INCOMPLET")',
    '=COUNTIF(\'STOCK VO\'!L2:L5000,"Préparation")',
    '=COUNTIFS(\'ADMINISTRATIF\'!I2:I5000,"<>"&"",\'ADMINISTRATIF\'!I2:I5000,"<"&TODAY()+30)',
  ];
  const kpiFormats3 = ['#,##0', '#,##0', '#,##0', '#,##0'];
  const kpiColors3 = [CONFIG.COLORS.WARNING, CONFIG.COLORS.DANGER, CONFIG.COLORS.INFO, CONFIG.COLORS.DANGER];

  creerCarteKPI(sheet, kpiRow3Start, kpiTitres3, kpiFormules3, kpiFormats3, kpiColors3);

  sheet.setRowHeight(kpiRow3Start + 5, 15);

  // ---- BOUTONS DE NAVIGATION ----
  const btnRow = kpiRow3Start + 6;
  sheet.setRowHeight(btnRow, 8);
  sheet.setRowHeight(btnRow + 1, 40);
  sheet.setRowHeight(btnRow + 2, 8);

  const btnLabels = ['📦  STOCK VO', '👥  CRM CLIENTS', '💰  VENTES', '📄  ADMINISTRATIF'];
  const btnCols = ['B', 'C', 'E', 'G'];

  // Titre section
  const rangeBtnTitle = sheet.getRange(`B${btnRow}:G${btnRow}`);
  rangeBtnTitle.merge();
  rangeBtnTitle.setValue('NAVIGATION RAPIDE');
  rangeBtnTitle.setFontSize(9);
  rangeBtnTitle.setFontColor(CONFIG.COLORS.MID_GRAY);
  rangeBtnTitle.setFontWeight('bold');
  rangeBtnTitle.setHorizontalAlignment('center');
  rangeBtnTitle.setBackground(CONFIG.COLORS.LIGHT_BG);

  const btnColors = [CONFIG.COLORS.INFO, CONFIG.COLORS.SUCCESS, CONFIG.COLORS.ACCENT, CONFIG.COLORS.WARNING];
  const btnCols2 = [['B', 'C'], ['D', 'D'], ['E', 'F'], ['G', 'G']];
  const btnLinks = ['#gid=STOCK_VO', '#gid=CRM', '#gid=VENTES', '#gid=ADMIN'];

  // Boutons manuels (pas de DrawingService requis)
  const btnData = [
    { range: 'B' + (btnRow + 1) + ':C' + (btnRow + 1), label: '📦  STOCK VO', color: CONFIG.COLORS.INFO },
    { range: 'D' + (btnRow + 1) + ':D' + (btnRow + 1), label: '👥  CLIENTS', color: CONFIG.COLORS.SUCCESS },
    { range: 'E' + (btnRow + 1) + ':F' + (btnRow + 1), label: '💰  VENTES', color: CONFIG.COLORS.ACCENT },
    { range: 'G' + (btnRow + 1) + ':G' + (btnRow + 1), label: '📄  ADMIN', color: CONFIG.COLORS.WARNING },
  ];

  btnData.forEach(btn => {
    const r = sheet.getRange(btn.range);
    r.merge();
    r.setValue(btn.label);
    r.setBackground(btn.color);
    r.setFontColor(CONFIG.COLORS.WHITE);
    r.setFontWeight('bold');
    r.setFontSize(12);
    r.setHorizontalAlignment('center');
    r.setVerticalAlignment('middle');
  });

  // Footer
  sheet.setRowHeight(btnRow + 3, 10);
  sheet.setRowHeight(btnRow + 4, 25);
  const rangeFooter = sheet.getRange(`B${btnRow + 4}:G${btnRow + 4}`);
  rangeFooter.merge();
  rangeFooter.setValue(`SpiderVO Pro v${CONFIG.VERSION}  |  Système de gestion VO  |  © ${new Date().getFullYear()}`);
  rangeFooter.setBackground(CONFIG.COLORS.PRIMARY);
  rangeFooter.setFontColor('#475569');
  rangeFooter.setFontSize(9);
  rangeFooter.setHorizontalAlignment('center');
  rangeFooter.setVerticalAlignment('middle');

  // Figer les lignes de titre
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  // Protéger le dashboard (lecture seule sauf formules)
  // Note: la protection est optionnelle
}

// Fonction helper : créer une ligne de 4 cartes KPI
function creerCarteKPI(sheet, startRow, titres, formules, formats, couleurs) {
  const cols = [2, 3, 4, 5]; // B, C, D, E sauf qu'on merge par 1.5 col chaque
  const cardRanges = [
    `B${startRow}:B${startRow + 4}`,
    `C${startRow}:C${startRow + 4}`,
    `E${startRow}:E${startRow + 4}`,
    `G${startRow}:G${startRow + 4}`,
  ];

  // Layout 2+2 avec colonnes D et F comme séparateurs
  sheet.setColumnWidth(4, 10);  // Séparateur D
  sheet.setColumnWidth(6, 10);  // Séparateur F

  const cardDefs = [
    { mergeRange: `B${startRow}:C${startRow + 4}`, col: 'B' },
    { mergeRange: `D${startRow}:D${startRow + 4}`, col: 'D', separator: true },
    { mergeRange: `E${startRow}:F${startRow + 4}`, col: 'E' },
    { mergeRange: `G${startRow}:G${startRow + 4}`, col: 'G', separator: true },
  ];

  // Ajuster hauteurs
  for (let i = 0; i < 5; i++) sheet.setRowHeight(startRow + i, i === 0 ? 30 : i === 4 ? 40 : 30);

  let cardIdx = 0;
  const pairs = [
    [`B${startRow}:C${startRow + 4}`, 0],
    [`E${startRow}:F${startRow + 4}`, 1],
    [`B${startRow + 0}:C${startRow + 4}`, 0], // dummy
    [`G${startRow}:H${startRow + 4}`, 2],     // dummy
  ];

  // Approche simplifiée : 4 cartes côte à côte sur B C D E F G (2 cols chacune avec petit séparateur)
  // On refait proprement :
  const cardColStarts = [2, 3, 5, 6]; // col index (1-based): B=2, C=3, E=5, F=6 -> pair B:C, D sep, E:F, G sep? Non.
  // Structure: B:C = card1 | D = 2px sep | E:F = card2 | [sep] | col G pas dispo après le 6e
  // Ajustons pour 4 cartes sur 6 colonnes B à G :
  // B:C = carte1, D = sep, E:F = carte2, G = carte3 seul col, ...
  // Mieux : colonnes A(marge) B C D E F G H(marge) avec 4 cartes B:C, D vide, E:F, G vide + H
  // Déjà défini en haut. Faisons 4 blocs de 1 colonne avec les 6 disponibles.

  // Réalisation : on merge chaque carte sur 1.5 col → trop complexe.
  // Solution propre : 4 cartes sur colonnes B, C, E, G avec D et F comme mini-séparateurs
  const kpiCols = ['B', 'C', 'E', 'G'];

  // Séparateurs
  sheet.setColumnWidth(4, 8);  // D
  sheet.setColumnWidth(6, 8);  // F

  titres.forEach((titre, i) => {
    const col = kpiCols[i];
    const bg = couleurs[i];

    // Carte entière
    const rCard = sheet.getRange(`${col}${startRow}:${col}${startRow + 4}`);

    // Ligne 1 : indicateur coloré
    const rIndicateur = sheet.getRange(`${col}${startRow}`);
    rIndicateur.setBackground(bg);
    sheet.setRowHeight(startRow, 8);

    // Ligne 2 : titre
    const rTitre = sheet.getRange(`${col}${startRow + 1}`);
    rTitre.setValue(titre);
    rTitre.setBackground(CONFIG.COLORS.LIGHT_BG);
    rTitre.setFontColor(CONFIG.COLORS.MID_GRAY);
    rTitre.setFontSize(9);
    rTitre.setFontWeight('bold');
    rTitre.setHorizontalAlignment('center');
    rTitre.setVerticalAlignment('middle');
    sheet.setRowHeight(startRow + 1, 30);

    // Ligne 3 : valeur principale
    const rValeur = sheet.getRange(`${col}${startRow + 2}`);
    rValeur.setFormula(formules[i]);
    rValeur.setNumberFormat(formats[i]);
    rValeur.setBackground(CONFIG.COLORS.WHITE);
    rValeur.setFontColor(bg);
    rValeur.setFontSize(22);
    rValeur.setFontWeight('bold');
    rValeur.setHorizontalAlignment('center');
    rValeur.setVerticalAlignment('middle');
    sheet.setRowHeight(startRow + 2, 45);

    // Ligne 4 : variation (placeholder)
    const rVariation = sheet.getRange(`${col}${startRow + 3}`);
    rVariation.setValue('vs mois précédent');
    rVariation.setBackground(CONFIG.COLORS.WHITE);
    rVariation.setFontColor(CONFIG.COLORS.MID_GRAY);
    rVariation.setFontSize(8);
    rVariation.setHorizontalAlignment('center');
    sheet.setRowHeight(startRow + 3, 22);

    // Bordure de la carte
    rCard.setBorder(true, true, true, true, false, false,
      CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  });
}

// ============================================================
// ONGLET STOCK VO
// ============================================================
function creerOngletStockVO(ss) {
  let sheet = ss.getSheetByName('STOCK VO');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('STOCK VO');
  sheet.setTabColor(CONFIG.COLORS.INFO);

  // Définition des colonnes (44 colonnes)
  const colonnes = [
    // IDENTIFICATION
    { label: 'ID Dossier', width: 90, group: 'IDENTIFICATION' },          // A
    { label: 'Immatriculation', width: 110, group: 'IDENTIFICATION' },    // B
    { label: 'VIN', width: 160, group: 'IDENTIFICATION' },                // C
    { label: 'Marque', width: 100, group: 'IDENTIFICATION' },             // D
    { label: 'Modèle', width: 120, group: 'IDENTIFICATION' },             // E
    { label: 'Version', width: 140, group: 'IDENTIFICATION' },            // F
    { label: 'Finition', width: 120, group: 'IDENTIFICATION' },           // G
    { label: 'Année', width: 65, group: 'IDENTIFICATION' },               // H
    { label: 'Kilométrage', width: 95, group: 'IDENTIFICATION' },         // I
    { label: 'Carburant', width: 100, group: 'IDENTIFICATION' },          // J
    { label: 'Boîte', width: 90, group: 'IDENTIFICATION' },               // K
    { label: 'STATUT', width: 110, group: 'IDENTIFICATION' },             // L ★
    // CARACTÉRISTIQUES
    { label: 'Puissance Fisc.', width: 85, group: 'CARACTÉRISTIQUES' },   // M
    { label: 'Date MEC', width: 90, group: 'CARACTÉRISTIQUES' },          // N
    { label: 'Couleur', width: 90, group: 'CARACTÉRISTIQUES' },           // O
    { label: 'Date Achat', width: 90, group: 'CARACTÉRISTIQUES' },        // P (pour dormants)
    { label: 'Nb Clés', width: 60, group: 'CARACTÉRISTIQUES' },           // Q
    { label: 'Nb Propriétaires', width: 80, group: 'CARACTÉRISTIQUES' },  // R
    { label: 'Historique Entretien', width: 130, group: 'CARACTÉRISTIQUES' }, // S
    { label: 'Crit\'Air', width: 100, group: 'CARACTÉRISTIQUES' },        // T
    { label: 'CT Validité', width: 100, group: 'CARACTÉRISTIQUES' },      // U
    { label: 'Garantie (mois)', width: 95, group: 'CARACTÉRISTIQUES' },   // V
    // ACHAT
    { label: 'Provenance', width: 120, group: 'ACHAT' },                  // W
    { label: 'Fournisseur', width: 140, group: 'ACHAT' },                 // X
    { label: 'Prix Achat HT', width: 110, group: 'ACHAT' },              // Y
    { label: 'Frais Transport', width: 100, group: 'ACHAT' },             // Z
    { label: 'Réparations', width: 100, group: 'ACHAT' },                 // AA
    { label: 'Nettoyage', width: 85, group: 'ACHAT' },                    // AB
    { label: 'Carte Grise', width: 90, group: 'ACHAT' },                  // AC
    { label: 'Préparation', width: 95, group: 'ACHAT' },                  // AD
    { label: 'Divers', width: 80, group: 'ACHAT' },                       // AE
    // CALCULS AUTO
    { label: 'Coût de Revient', width: 115, group: 'CALCULS' },           // AF ★
    { label: 'Prix Conseillé', width: 110, group: 'CALCULS' },            // AG
    { label: 'Prix Affiché', width: 100, group: 'CALCULS' },              // AH
    { label: 'Prix Mini Acceptable', width: 130, group: 'CALCULS' },      // AI
    { label: 'Marge Brute €', width: 110, group: 'CALCULS' },             // AJ
    { label: 'Marge Brute %', width: 110, group: 'CALCULS' },             // AK
    { label: 'Marge Nette €', width: 110, group: 'CALCULS' },             // AL
    { label: 'TVA/Marge', width: 90, group: 'CALCULS' },                  // AM
    { label: 'ROI %', width: 80, group: 'CALCULS' },                      // AN
    // VENTE
    { label: 'Date Vente', width: 90, group: 'VENTE' },                   // AO
    { label: 'Prix Vente', width: 100, group: 'VENTE' },                  // AP
    { label: 'Client', width: 140, group: 'VENTE' },                      // AQ
    // NOTES
    { label: 'Jours en Stock', width: 90, group: 'INFO' },                // AR
    { label: 'Notes', width: 200, group: 'INFO' },                        // AS
  ];

  // Appliquer les largeurs
  colonnes.forEach((col, i) => {
    sheet.setColumnWidth(i + 1, col.width);
  });

  // En-tête ligne 1 : groupes de colonnes
  const groupes = {};
  colonnes.forEach((col, i) => {
    if (!groupes[col.group]) groupes[col.group] = { start: i + 1, end: i + 1 };
    else groupes[col.group].end = i + 1;
  });

  // Couleurs par groupe
  const groupColors = {
    'IDENTIFICATION': CONFIG.COLORS.PRIMARY,
    'CARACTÉRISTIQUES': CONFIG.COLORS.SECONDARY,
    'ACHAT': '#1A3A2A',
    'CALCULS': '#2D1A3A',
    'VENTE': '#3A1A1A',
    'INFO': '#2A2A2A',
  };

  sheet.setRowHeight(1, 28);
  Object.entries(groupes).forEach(([groupe, pos]) => {
    const r = sheet.getRange(1, pos.start, 1, pos.end - pos.start + 1);
    r.merge();
    r.setValue(groupe);
    r.setBackground(groupColors[groupe] || CONFIG.COLORS.PRIMARY);
    r.setFontColor(CONFIG.COLORS.WHITE);
    r.setFontWeight('bold');
    r.setFontSize(9);
    r.setHorizontalAlignment('center');
    r.setVerticalAlignment('middle');
  });

  // En-tête ligne 2 : noms des colonnes
  sheet.setRowHeight(2, 40);
  const headerRange = sheet.getRange(2, 1, 1, colonnes.length);
  const headerValues = colonnes.map(c => c.label);
  headerRange.setValues([headerValues]);
  headerRange.setBackground(CONFIG.COLORS.HEADER_BG);
  headerRange.setFontColor(CONFIG.COLORS.WHITE);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(9);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrap(true);

  // Colonne STATUT (L=12) en rouge pour la différencier
  sheet.getRange(2, 12).setBackground(CONFIG.COLORS.ACCENT);

  // Formules automatiques (à partir de la ligne 3)
  const maxRows = 5000;

  // Coût de revient = Achat + Transport + Réparations + Nettoyage + Carte grise + Préparation + Divers
  // Col AF = 32 = Y(25) + Z(26) + AA(27) + AB(28) + AC(29) + AD(30) + AE(31)
  for (let row = 3; row <= 52; row++) { // Pré-remplir 50 lignes de formules
    const r = row;

    // Coût de revient (AF = col 32)
    sheet.getRange(r, 32).setFormula(`=IF(A${r}="","",SUM(Y${r}:AE${r}))`);
    sheet.getRange(r, 32).setNumberFormat('#,##0 "€"');

    // Prix conseillé (AG = col 33) : coût * 1.20 minimum
    sheet.getRange(r, 33).setFormula(`=IF(AF${r}="","",ROUND(AF${r}*1.20/100,0)*100)`);
    sheet.getRange(r, 33).setNumberFormat('#,##0 "€"');

    // Prix mini acceptable (AI = col 35) : coût * 1.05
    sheet.getRange(r, 35).setFormula(`=IF(AF${r}="","",ROUND(AF${r}*1.05/100,0)*100)`);
    sheet.getRange(r, 35).setNumberFormat('#,##0 "€"');

    // Marge brute € (AJ = col 36) : Prix affiché - Coût de revient
    sheet.getRange(r, 36).setFormula(`=IF(OR(AH${r}="",AF${r}=""),"",AH${r}-AF${r})`);
    sheet.getRange(r, 36).setNumberFormat('#,##0 "€"');

    // Marge brute % (AK = col 37)
    sheet.getRange(r, 37).setFormula(`=IF(OR(AF${r}="",AF${r}=0),"",AJ${r}/AF${r})`);
    sheet.getRange(r, 37).setNumberFormat('0.0"%"');

    // TVA sur marge (AM = col 39) = (Prix affiché TTC - Coût achat TTC) / 1.20 * 0.20
    sheet.getRange(r, 39).setFormula(`=IF(OR(AH${r}="",AF${r}=""),"",ROUND((AH${r}-AF${r})/1.20*0.20,2))`);
    sheet.getRange(r, 39).setNumberFormat('#,##0.00 "€"');

    // Marge nette (AL = col 38) : Marge brute - TVA
    sheet.getRange(r, 38).setFormula(`=IF(AJ${r}="","",AJ${r}-AM${r})`);
    sheet.getRange(r, 38).setNumberFormat('#,##0 "€"');

    // ROI % (AN = col 40)
    sheet.getRange(r, 40).setFormula(`=IF(OR(AF${r}="",AF${r}=0),"",AL${r}/AF${r})`);
    sheet.getRange(r, 40).setNumberFormat('0.0"%"');

    // Jours en stock (AR = col 44)
    sheet.getRange(r, 44).setFormula(`=IF(OR(A${r}="",P${r}=""),"",IF(AO${r}<>"",AO${r}-P${r},TODAY()-P${r}))`);
    sheet.getRange(r, 44).setNumberFormat('#,##0 "j"');
  }

  // Mise en forme conditionnelle
  const dataRange = sheet.getRange('A3:AS5002');

  // Formatage conditionnel STATUT (colonne L)
  const statutRange = sheet.getRange('L3:L5002');
  const rules = sheet.getConditionalFormatRules();

  Object.entries(CONFIG.STATUTS_COULEURS).forEach(([statut, couleur]) => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(statut)
      .setBackground(couleur)
      .setFontColor(CONFIG.COLORS.WHITE)
      .setRanges([sheet.getRange('L3:L5002')])
      .build();
    rules.push(rule);
  });

  // Marge brute rouge si négative
  const ruleMarge = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#FEE2E2')
    .setFontColor(CONFIG.COLORS.DANGER)
    .setRanges([sheet.getRange('AJ3:AJ5002')])
    .build();
  rules.push(ruleMarge);

  // Jours en stock : rouge si > 45j
  const ruleDormant = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(CONFIG.SEUIL_DORMANT)
    .setBackground('#FEF9C3')
    .setFontColor('#92400E')
    .setRanges([sheet.getRange('AR3:AR5002')])
    .build();
  rules.push(ruleDormant);

  sheet.setConditionalFormatRules(rules);

  // Validations données
  const validations = [
    { col: 10, values: CONFIG.CARBURANTS },        // J - Carburant
    { col: 11, values: CONFIG.BOITES },            // K - Boîte
    { col: 12, values: CONFIG.STATUTS_STOCK },     // L - Statut
    { col: 20, values: CONFIG.CRITAIR },           // T - Crit'Air
    { col: 23, values: CONFIG.PROVENANCES },       // W - Provenance
    { col: 19, values: ['Complet', 'Carnet partiel', 'Absent', 'Non applicable'] }, // S - Historique
  ];

  validations.forEach(v => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(v.values, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(3, v.col, maxRows, 1).setDataValidation(rule);
  });

  // Formats numériques
  sheet.getRange('I3:I5002').setNumberFormat('#,##0 " km"');    // Kilométrage
  sheet.getRange('N3:N5002').setNumberFormat('dd/mm/yyyy');     // Date MEC
  sheet.getRange('P3:P5002').setNumberFormat('dd/mm/yyyy');     // Date achat
  sheet.getRange('U3:U5002').setNumberFormat('dd/mm/yyyy');     // CT
  sheet.getRange('AO3:AO5002').setNumberFormat('dd/mm/yyyy');   // Date vente
  sheet.getRange('Y3:AE5002').setNumberFormat('#,##0 "€"');     // Prix achats
  sheet.getRange('AH3:AH5002').setNumberFormat('#,##0 "€"');    // Prix affiché
  sheet.getRange('AP3:AP5002').setNumberFormat('#,##0 "€"');    // Prix vente

  // Alternance couleurs lignes
  const altRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=MOD(ROW(),2)=1')
    .setBackground(CONFIG.COLORS.ROW_ALT)
    .setRanges([sheet.getRange('A3:AS5002')])
    .build();
  const rulesWithAlt = sheet.getConditionalFormatRules();
  rulesWithAlt.push(altRule);
  sheet.setConditionalFormatRules(rulesWithAlt);

  // Figer les 2 premières lignes et les 2 premières colonnes
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(2);

  // Filtre automatique
  sheet.getRange('A2:AS2').createFilter();

  // Hauteur des lignes de données
  for (let r = 3; r <= 52; r++) sheet.setRowHeight(r, 28);
}

// ============================================================
// ONGLET CRM CLIENTS
// ============================================================
function creerOngletCRM(ss) {
  let sheet = ss.getSheetByName('CRM CLIENTS');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('CRM CLIENTS');
  sheet.setTabColor(CONFIG.COLORS.SUCCESS);

  const colonnes = [
    { label: 'ID Client', width: 90 },
    { label: 'Nom Prénom', width: 160 },
    { label: 'Téléphone', width: 120 },
    { label: 'Email', width: 180 },
    { label: 'Ville', width: 110 },
    { label: 'Source', width: 120 },
    { label: 'Véhicule Recherché', width: 180 },
    { label: 'STATUT PIPELINE', width: 140 },
    { label: 'Score', width: 70 },
    { label: 'Budget Max', width: 100 },
    { label: 'Date Entrée', width: 100 },
    { label: 'Dernier Contact', width: 100 },
    { label: 'Prochain RDV', width: 100 },
    { label: 'Nb Contacts', width: 80 },
    { label: 'Véhicule Attribué', width: 140 },
    { label: 'Vendeur', width: 110 },
    { label: 'Notes / Historique', width: 250 },
    { label: 'Relance Auto', width: 100 },
    { label: 'Priorité', width: 80 },
  ];

  colonnes.forEach((col, i) => sheet.setColumnWidth(i + 1, col.width));

  // En-tête
  sheet.setRowHeight(1, 40);
  const headerRange = sheet.getRange(1, 1, 1, colonnes.length);
  headerRange.setValues([colonnes.map(c => c.label)]);
  headerRange.setBackground(CONFIG.COLORS.HEADER_BG);
  headerRange.setFontColor(CONFIG.COLORS.WHITE);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrap(true);

  // Colonne pipeline en vert
  sheet.getRange(1, 8).setBackground(CONFIG.COLORS.SUCCESS);

  // Validations
  const rulePipeline = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.PIPELINE_ETAPES, true)
    .build();
  sheet.getRange('H2:H5002').setDataValidation(rulePipeline);

  const ruleScore = SpreadsheetApp.newDataValidation()
    .requireValueInList(['🔥 Chaud', '🌡️ Tiède', '❄️ Froid', '💀 Perdu'], true)
    .build();
  sheet.getRange('I2:I5002').setDataValidation(ruleScore);

  const rulePriorite = SpreadsheetApp.newDataValidation()
    .requireValueInList(['🔴 Urgente', '🟠 Haute', '🟡 Normale', '🟢 Basse'], true)
    .build();
  sheet.getRange('S2:S5002').setDataValidation(rulePriorite);

  const ruleSource = SpreadsheetApp.newDataValidation()
    .requireValueInList(['LeBonCoin', 'Autosphere', 'LaArgus', 'Bouche à oreille', 'Réseaux sociaux', 'Site web', 'Walk-in', 'Tel entrant', 'Partenaire', 'Autre'], true)
    .build();
  sheet.getRange('F2:F5002').setDataValidation(ruleSource);

  // Formats dates
  sheet.getRange('K2:K5002').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('L2:L5002').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('M2:M5002').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('J2:J5002').setNumberFormat('#,##0 "€"');

  // Formatage conditionnel pipeline
  const rules = [];
  const pipelineColors = {
    'Nouveau lead': ['#EFF6FF', '#1D4ED8'],
    'Contacté': ['#F0FDF4', '#15803D'],
    'RDV planifié': ['#FFF7ED', '#C2410C'],
    'Essai': ['#FFF7ED', '#B45309'],
    'Négociation': ['#FFFBEB', '#92400E'],
    'Réservé': ['#F5F3FF', '#6D28D9'],
    'Vendu': ['#F0FDF4', '#166534'],
    'Perdu': ['#FEF2F2', '#991B1B'],
  };

  Object.entries(pipelineColors).forEach(([statut, [bg, fg]]) => {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(statut)
      .setBackground(bg)
      .setFontColor(fg)
      .setRanges([sheet.getRange('H2:H5002')])
      .build());
  });

  // Relance : rouge si dernier contact > 7j et pas vendu/perdu
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(L2<>"",TODAY()-L2>7,H2<>"Vendu",H2<>"Perdu")')
    .setBackground('#FEE2E2')
    .setRanges([sheet.getRange('L2:L5002')])
    .build());

  sheet.setConditionalFormatRules(rules);

  // Filtre + figer
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.getRange('A1:S1').createFilter();

  for (let r = 2; r <= 52; r++) sheet.setRowHeight(r, 28);
}

// ============================================================
// ONGLET VENTES
// ============================================================
function creerOngletVentes(ss) {
  let sheet = ss.getSheetByName('VENTES');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('VENTES');
  sheet.setTabColor(CONFIG.COLORS.ACCENT);

  const colonnes = [
    { label: 'N° Commande', width: 110 },
    { label: 'N° Facture', width: 100 },
    { label: 'ID Dossier VO', width: 100 },
    { label: 'Client', width: 160 },
    { label: 'Date Vente', width: 100 },
    { label: 'Date Livraison', width: 100 },
    { label: 'Véhicule', width: 200 },
    { label: 'Immatriculation', width: 110 },
    { label: 'Prix Vente TTC', width: 120 },
    { label: 'Prix Vente HT', width: 120 },
    { label: 'Coût Revient', width: 110 },
    { label: 'Marge Brute €', width: 110 },
    { label: 'Marge Nette €', width: 110 },
    { label: 'TVA Marge', width: 100 },
    { label: 'Marge %', width: 80 },
    { label: 'Mode Paiement', width: 130 },
    { label: 'Reprise', width: 90 },
    { label: 'Valeur Reprise', width: 110 },
    { label: 'Financement', width: 110 },
    { label: 'Garantie Vendue', width: 115 },
    { label: 'Prix Garantie', width: 100 },
    { label: 'Statut', width: 110 },
    { label: 'Vendeur', width: 110 },
    { label: 'Notes', width: 200 },
  ];

  colonnes.forEach((col, i) => sheet.setColumnWidth(i + 1, col.width));

  sheet.setRowHeight(1, 40);
  const headerRange = sheet.getRange(1, 1, 1, colonnes.length);
  headerRange.setValues([colonnes.map(c => c.label)]);
  headerRange.setBackground(CONFIG.COLORS.ACCENT);
  headerRange.setFontColor(CONFIG.COLORS.WHITE);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrap(true);

  // Formules auto sur les 50 premières lignes
  for (let r = 2; r <= 51; r++) {
    // Prix HT = Prix TTC / 1.20 (pour vente avec TVA normale - à adapter si régime de marge)
    sheet.getRange(r, 10).setFormula(`=IF(I${r}="","",ROUND(I${r}/1.2,2))`);
    // Marge brute = Prix vente TTC - Coût revient
    sheet.getRange(r, 12).setFormula(`=IF(OR(I${r}="",K${r}=""),"",I${r}-K${r})`);
    // TVA sur marge = (Prix vente - Coût) / 6
    sheet.getRange(r, 14).setFormula(`=IF(L${r}="","",ROUND(L${r}/6,2))`);
    // Marge nette = Marge brute - TVA marge
    sheet.getRange(r, 13).setFormula(`=IF(L${r}="","",L${r}-N${r})`);
    // Marge %
    sheet.getRange(r, 15).setFormula(`=IF(OR(K${r}="",K${r}=0),"",M${r}/K${r})`);
  }

  // Formats
  sheet.getRange('E2:F5002').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('I2:N5002').setNumberFormat('#,##0 "€"');
  sheet.getRange('O2:O5002').setNumberFormat('0.0"%"');
  sheet.getRange('R2:R5002').setNumberFormat('#,##0 "€"');
  sheet.getRange('U2:U5002').setNumberFormat('#,##0 "€"');

  // Validations
  const ruleModePaiement = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Chèque', 'Virement', 'Espèces', 'Financement', 'Mixte'], true)
    .build();
  sheet.getRange('P2:P5002').setDataValidation(ruleModePaiement);

  const ruleStatut = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Bon de commande', 'Acompte reçu', 'Payé', 'Livré', 'Annulé'], true)
    .build();
  sheet.getRange('V2:V5002').setDataValidation(ruleStatut);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  sheet.getRange('A1:X1').createFilter();
  for (let r = 2; r <= 52; r++) sheet.setRowHeight(r, 28);
}

// ============================================================
// ONGLET ADMINISTRATIF (type AutoCERFA)
// ============================================================
function creerOngletAdministratif(ss) {
  let sheet = ss.getSheetByName('ADMINISTRATIF');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('ADMINISTRATIF');
  sheet.setTabColor(CONFIG.COLORS.WARNING);

  const colonnes = [
    { label: 'ID Dossier', width: 90 },
    { label: 'Immatriculation', width: 110 },
    { label: 'Marque / Modèle', width: 160 },
    { label: 'VIN', width: 160 },
    // ACHAT
    { label: 'Décl. Achat envoyée', width: 120 },
    { label: 'Date Décl. Achat', width: 110 },
    { label: 'CERFA Achat N°', width: 110 },
    // CESSION
    { label: 'Décl. Cession envoyée', width: 130 },
    // CT
    { label: 'CT Date Validité', width: 110 },
    { label: 'CT Résultat', width: 100 },
    { label: 'CT Avis', width: 90 },
    // CARTE GRISE
    { label: 'CG Reçue', width: 80 },
    { label: 'CG Type', width: 100 },
    { label: 'CG Transmise Client', width: 130 },
    { label: 'Date Transmission CG', width: 130 },
    // STATUT DOSSIER
    { label: 'STATUT DOSSIER', width: 130 },
    { label: 'Documents Manquants', width: 200 },
    // GARANTIE
    { label: 'Garantie Durée (mois)', width: 130 },
    { label: 'Garantie Fin', width: 110 },
    // DIVERS
    { label: 'Notes Admin', width: 220 },
  ];

  colonnes.forEach((col, i) => sheet.setColumnWidth(i + 1, col.width));

  sheet.setRowHeight(1, 40);
  const headerRange = sheet.getRange(1, 1, 1, colonnes.length);
  headerRange.setValues([colonnes.map(c => c.label)]);
  headerRange.setBackground('#92400E');
  headerRange.setFontColor(CONFIG.COLORS.WHITE);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrap(true);

  // Colonne STATUT DOSSIER
  sheet.getRange(1, 16).setBackground(CONFIG.COLORS.DANGER);

  // Validations
  const ruleOuiNon = SpreadsheetApp.newDataValidation()
    .requireValueInList(['✅ Oui', '❌ Non', '⏳ En attente'], true).build();
  ['E', 'H', 'L', 'N'].forEach(col =>
    sheet.getRange(`${col}2:${col}5002`).setDataValidation(ruleOuiNon));

  const ruleCTRes = SpreadsheetApp.newDataValidation()
    .requireValueInList(['✅ Favorable', '⚠️ Défavorable', '🔄 À faire', 'N/A'], true).build();
  sheet.getRange('J2:J5002').setDataValidation(ruleCTRes);

  const ruleStatutDossier = SpreadsheetApp.newDataValidation()
    .requireValueInList(['✅ COMPLET', '⚠️ INCOMPLET', '❌ BLOQUANT', '📋 EN COURS'], true).build();
  sheet.getRange('P2:P5002').setDataValidation(ruleStatutDossier);

  // Formats
  sheet.getRange('F2:F5002').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('I2:I5002').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('O2:O5002').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('S2:S5002').setNumberFormat('dd/mm/yyyy');

  // Formule Garantie Fin
  for (let r = 2; r <= 51; r++) {
    sheet.getRange(r, 19).setFormula(`=IF(OR(R${r}="",R${r}=0),"",EDATE(TODAY(),R${r}))`);
    sheet.getRange(r, 19).setNumberFormat('dd/mm/yyyy');
  }

  // Formatage conditionnel statut
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('✅ COMPLET').setBackground('#F0FDF4').setFontColor('#15803D')
      .setRanges([sheet.getRange('P2:P5002')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('⚠️ INCOMPLET').setBackground('#FFFBEB').setFontColor('#92400E')
      .setRanges([sheet.getRange('P2:P5002')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('❌ BLOQUANT').setBackground('#FEF2F2').setFontColor('#991B1B')
      .setRanges([sheet.getRange('P2:P5002')]).build(),
    // CT expirant dans 30j
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(I2<>"",I2<TODAY()+30,I2>=TODAY())')
      .setBackground('#FFFBEB').setFontColor('#92400E')
      .setRanges([sheet.getRange('I2:I5002')]).build(),
    // CT expiré
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(I2<>"",I2<TODAY())')
      .setBackground('#FEF2F2').setFontColor('#991B1B')
      .setRanges([sheet.getRange('I2:I5002')]).build(),
  ];
  sheet.setConditionalFormatRules(rules);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.getRange('A1:T1').createFilter();
  for (let r = 2; r <= 52; r++) sheet.setRowHeight(r, 28);
}

// ============================================================
// ONGLET FINANCE
// ============================================================
function creerOngletFinance(ss) {
  let sheet = ss.getSheetByName('FINANCE');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('FINANCE');
  sheet.setTabColor(CONFIG.COLORS.GOLD);

  // ---- Tableau de dépenses mensuelles ----
  const colsDep = [
    { label: 'Date', width: 100 },
    { label: 'Catégorie', width: 140 },
    { label: 'Sous-catégorie', width: 160 },
    { label: 'Description', width: 220 },
    { label: 'Montant HT', width: 110 },
    { label: 'TVA', width: 80 },
    { label: 'Montant TTC', width: 110 },
    { label: 'Mode Paiement', width: 130 },
    { label: 'Référence', width: 120 },
    { label: 'Mois', width: 70 },
    { label: 'Année', width: 70 },
    { label: 'Notes', width: 200 },
  ];

  colsDep.forEach((col, i) => sheet.setColumnWidth(i + 1, col.width));

  // Titre
  sheet.setRowHeight(1, 50);
  sheet.getRange('A1:L1').merge();
  sheet.getRange('A1').setValue('💰  SUIVI FINANCIER — Dépenses & Rentabilité');
  sheet.getRange('A1').setBackground(CONFIG.COLORS.PRIMARY);
  sheet.getRange('A1').setFontColor(CONFIG.COLORS.GOLD);
  sheet.getRange('A1').setFontSize(16);
  sheet.getRange('A1').setFontWeight('bold');
  sheet.getRange('A1').setHorizontalAlignment('center');
  sheet.getRange('A1').setVerticalAlignment('middle');

  // En-tête
  sheet.setRowHeight(2, 38);
  const headerRange = sheet.getRange(2, 1, 1, colsDep.length);
  headerRange.setValues([colsDep.map(c => c.label)]);
  headerRange.setBackground('#78350F');
  headerRange.setFontColor(CONFIG.COLORS.WHITE);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');

  // Validations
  const ruleCat = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'Publicité', 'Achats véhicules', 'Réparations', 'Préparation',
      'Transport', 'Carburant', 'Loyer', 'Salaires', 'Assurance',
      'Fournitures', 'Carte grise', 'Formation', 'Divers'
    ], true).build();
  sheet.getRange('B3:B5002').setDataValidation(ruleCat);

  // Formules auto
  for (let r = 3; r <= 52; r++) {
    sheet.getRange(r, 6).setFormula(`=IF(E${r}="","",ROUND(E${r}*${CONFIG.TVA_TAUX},2))`);
    sheet.getRange(r, 7).setFormula(`=IF(E${r}="","",E${r}+F${r})`);
    sheet.getRange(r, 10).setFormula(`=IF(A${r}="","",MONTH(A${r}))`);
    sheet.getRange(r, 11).setFormula(`=IF(A${r}="","",YEAR(A${r}))`);
  }

  sheet.getRange('A3:A5002').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('E3:G5002').setNumberFormat('#,##0.00 "€"');

  // ---- Résumé mensuel (à droite) ----
  sheet.setColumnWidth(14, 160);
  sheet.setColumnWidth(15, 130);
  sheet.setColumnWidth(16, 130);

  sheet.getRange('N1:P1').merge();
  sheet.getRange('N1').setValue('📊 RÉCAPITULATIF MENSUEL');
  sheet.getRange('N1').setBackground(CONFIG.COLORS.PRIMARY);
  sheet.getRange('N1').setFontColor(CONFIG.COLORS.GOLD);
  sheet.getRange('N1').setFontSize(13);
  sheet.getRange('N1').setFontWeight('bold');
  sheet.getRange('N1').setHorizontalAlignment('center');
  sheet.getRange('N1').setVerticalAlignment('middle');

  const resumeLabels = [
    'CA Mensuel', 'Achats Véhicules', 'Réparations',
    'Publicité', 'Frais fixes', 'Salaires',
    'Transport', 'Autres', 'TOTAL DÉPENSES', 'BÉNÉFICE NET'
  ];

  const resumeFormules = [
    `=SUMPRODUCT((MONTH(VENTES!E2:E5000)=MONTH(TODAY()))*(YEAR(VENTES!E2:E5000)=YEAR(TODAY()))*(VENTES!I2:I5000))`,
    `=SUMPRODUCT((FINANCE!J3:J5002=MONTH(TODAY()))*(FINANCE!K3:K5002=YEAR(TODAY()))*(FINANCE!B3:B5002="Achats véhicules")*(FINANCE!G3:G5002))`,
    `=SUMPRODUCT((FINANCE!J3:J5002=MONTH(TODAY()))*(FINANCE!K3:K5002=YEAR(TODAY()))*(FINANCE!B3:B5002="Réparations")*(FINANCE!G3:G5002))`,
    `=SUMPRODUCT((FINANCE!J3:J5002=MONTH(TODAY()))*(FINANCE!K3:K5002=YEAR(TODAY()))*(FINANCE!B3:B5002="Publicité")*(FINANCE!G3:G5002))`,
    `=SUMPRODUCT((FINANCE!J3:J5002=MONTH(TODAY()))*(FINANCE!K3:K5002=YEAR(TODAY()))*(FINANCE!B3:B5002="Loyer")*(FINANCE!G3:G5002))+SUMPRODUCT((FINANCE!J3:J5002=MONTH(TODAY()))*(FINANCE!K3:K5002=YEAR(TODAY()))*(FINANCE!B3:B5002="Assurance")*(FINANCE!G3:G5002))`,
    `=SUMPRODUCT((FINANCE!J3:J5002=MONTH(TODAY()))*(FINANCE!K3:K5002=YEAR(TODAY()))*(FINANCE!B3:B5002="Salaires")*(FINANCE!G3:G5002))`,
    `=SUMPRODUCT((FINANCE!J3:J5002=MONTH(TODAY()))*(FINANCE!K3:K5002=YEAR(TODAY()))*(FINANCE!B3:B5002="Transport")*(FINANCE!G3:G5002))`,
    `=SUMPRODUCT((FINANCE!J3:J5002=MONTH(TODAY()))*(FINANCE!K3:K5002=YEAR(TODAY()))*(FINANCE!G3:G5002))-O3-O4-O5-O6-O7-O8`,
    `=SUM(O3:O9)`,
    `=O2-O10`,
  ];

  resumeLabels.forEach((label, i) => {
    const row = i + 2;
    sheet.setRowHeight(row, 32);
    sheet.getRange(row, 14).setValue(label);
    sheet.getRange(row, 14).setBackground(i === 8 ? '#1F2937' : i === 9 ? '#14532D' : CONFIG.COLORS.LIGHT_BG);
    sheet.getRange(row, 14).setFontColor(i >= 8 ? CONFIG.COLORS.WHITE : CONFIG.COLORS.DARK_TEXT);
    sheet.getRange(row, 14).setFontWeight(i >= 8 ? 'bold' : 'normal');
    sheet.getRange(row, 14).setFontSize(10);
    sheet.getRange(row, 14).setVerticalAlignment('middle');

    sheet.getRange(row, 15).setFormula(resumeFormules[i]);
    sheet.getRange(row, 15).setNumberFormat('#,##0 "€"');
    sheet.getRange(row, 15).setBackground(i === 9 ? '#14532D' : i === 8 ? '#1F2937' : CONFIG.COLORS.WHITE);
    sheet.getRange(row, 15).setFontColor(i === 9 ? '#4ADE80' : i === 8 ? '#F87171' : CONFIG.COLORS.DARK_TEXT);
    sheet.getRange(row, 15).setFontWeight(i >= 8 ? 'bold' : 'normal');
    sheet.getRange(row, 15).setFontSize(i >= 8 ? 13 : 11);
    sheet.getRange(row, 15).setHorizontalAlignment('right');
    sheet.getRange(row, 15).setVerticalAlignment('middle');

    // Bordures
    sheet.getRange(row, 14, 1, 2).setBorder(true, true, true, true, false, false,
      CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  });

  sheet.setFrozenRows(2);
  sheet.getRange('A2:L2').createFilter();
  for (let r = 3; r <= 52; r++) sheet.setRowHeight(r, 26);
}

// ============================================================
// ONGLET PARAMÈTRES
// ============================================================
function creerOngletParametres(ss) {
  let sheet = ss.getSheetByName('PARAMÈTRES');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('PARAMÈTRES');
  sheet.setTabColor('#6B7280');
  sheet.setHiddenGridlines(true);

  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 30);

  const params = [
    { section: 'IDENTITÉ GARAGE', items: [
      ['Nom du garage', 'Mon Garage VO'],
      ['SIRET', ''],
      ['TVA intracommunautaire', ''],
      ['Adresse', ''],
      ['Code Postal', ''],
      ['Ville', ''],
      ['Téléphone', ''],
      ['Email', ''],
      ['Site web', ''],
    ]},
    { section: 'COMMERCIAUX', items: [
      ['Vendeur 1', ''],
      ['Vendeur 2', ''],
      ['Vendeur 3', ''],
    ]},
    { section: 'PARAMÈTRES SYSTÈME', items: [
      ['Marge minimum %', '8'],
      ['Seuil véhicule dormant (jours)', String(CONFIG.SEUIL_DORMANT)],
      ['Email notifications', ''],
      ['Devise', 'EUR'],
    ]},
  ];

  let currentRow = 2;
  sheet.setRowHeight(1, 20);

  // Titre
  sheet.getRange('B1:C1').merge();
  sheet.getRange('B1').setValue('⚙️  PARAMÈTRES SpiderVO Pro');
  sheet.getRange('B1').setBackground(CONFIG.COLORS.PRIMARY);
  sheet.getRange('B1').setFontColor(CONFIG.COLORS.WHITE);
  sheet.getRange('B1').setFontSize(14);
  sheet.getRange('B1').setFontWeight('bold');
  sheet.getRange('B1').setHorizontalAlignment('center');
  sheet.getRange('B1').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 45);

  params.forEach(section => {
    currentRow++;
    sheet.setRowHeight(currentRow, 32);
    sheet.getRange(`B${currentRow}:C${currentRow}`).merge();
    sheet.getRange(`B${currentRow}`).setValue(section.section);
    sheet.getRange(`B${currentRow}`).setBackground(CONFIG.COLORS.SECONDARY);
    sheet.getRange(`B${currentRow}`).setFontColor(CONFIG.COLORS.WHITE);
    sheet.getRange(`B${currentRow}`).setFontWeight('bold');
    sheet.getRange(`B${currentRow}`).setFontSize(11);
    sheet.getRange(`B${currentRow}`).setVerticalAlignment('middle');
    sheet.getRange(`B${currentRow}`).setHorizontalAlignment('left');
    sheet.getRange(`B${currentRow}`).setIndent(1);

    section.items.forEach(([label, valeur]) => {
      currentRow++;
      sheet.setRowHeight(currentRow, 28);
      sheet.getRange(`B${currentRow}`).setValue(label);
      sheet.getRange(`B${currentRow}`).setBackground(CONFIG.COLORS.LIGHT_BG);
      sheet.getRange(`B${currentRow}`).setFontColor(CONFIG.COLORS.DARK_TEXT);
      sheet.getRange(`B${currentRow}`).setVerticalAlignment('middle');
      sheet.getRange(`B${currentRow}`).setIndent(2);

      sheet.getRange(`C${currentRow}`).setValue(valeur);
      sheet.getRange(`C${currentRow}`).setBackground(CONFIG.COLORS.WHITE);
      sheet.getRange(`C${currentRow}`).setBorder(false, false, true, false, false, false,
        CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(`C${currentRow}`).setVerticalAlignment('middle');
    });
    currentRow++;
  });
}

// ============================================================
// MISE À JOUR DASHBOARD
// ============================================================
function majDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DASHBOARD');
  if (!sheet) return;

  // Mise à jour de la date
  const now = new Date();
  sheet.getRange('B4:G4').setValue(
    `Tableau de bord — Mis à jour le ${Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy à HH:mm')}`
  );
}

function allerAuDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  majDashboard();
  ss.setActiveSheet(ss.getSheetByName('DASHBOARD'));
}

function majStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName('STOCK VO'));
  SpreadsheetApp.getUi().alert('✅ Stock actualisé', 'Les formules ont été recalculées.\n\nVéhicules dormants et alertes mis à jour.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function recalculerTout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Forcer le recalcul en touchant une cellule
  majDashboard();
  SpreadsheetApp.getUi().alert('✅ Recalcul complet', 'Toutes les formules ont été recalculées.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
// NUMÉROTATION AUTOMATIQUE
// ============================================================
function genererNumeroDossier() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('STOCK VO');
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return CONFIG.DOSSIER_PREFIX + '0001';
  const lastID = sheet.getRange(lastRow, 1).getValue();
  if (!lastID || !String(lastID).startsWith(CONFIG.DOSSIER_PREFIX)) return CONFIG.DOSSIER_PREFIX + '0001';
  const num = parseInt(String(lastID).replace(CONFIG.DOSSIER_PREFIX, '')) + 1;
  return CONFIG.DOSSIER_PREFIX + String(num).padStart(4, '0');
}

function genererNumeroClient() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CRM CLIENTS');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return CONFIG.CLIENT_PREFIX + '0001';
  const lastID = sheet.getRange(lastRow, 1).getValue();
  if (!lastID || !String(lastID).startsWith(CONFIG.CLIENT_PREFIX)) return CONFIG.CLIENT_PREFIX + '0001';
  const num = parseInt(String(lastID).replace(CONFIG.CLIENT_PREFIX, '')) + 1;
  return CONFIG.CLIENT_PREFIX + String(num).padStart(4, '0');
}

function genererNumeroCommande() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('VENTES');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return CONFIG.COMMANDE_PREFIX + '0001';
  const lastID = sheet.getRange(lastRow, 1).getValue();
  if (!lastID || !String(lastID).startsWith(CONFIG.COMMANDE_PREFIX)) return CONFIG.COMMANDE_PREFIX + '0001';
  const num = parseInt(String(lastID).replace(CONFIG.COMMANDE_PREFIX, '')) + 1;
  return CONFIG.COMMANDE_PREFIX + String(num).padStart(4, '0');
}

// ============================================================
// FORMULAIRE AJOUT VÉHICULE (Sidebar HTML)
// ============================================================
function ouvrirFormulaireVehicule() {
  const html = HtmlService.createHtmlOutput(getHTMLFormulaireVehicule())
    .setTitle('➕ Nouveau Véhicule')
    .setWidth(480);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHTMLFormulaireVehicule() {
  const nextID = genererNumeroDossier();
  const carburants = JSON.stringify(CONFIG.CARBURANTS);
  const boites = JSON.stringify(CONFIG.BOITES);
  const statuts = JSON.stringify(CONFIG.STATUTS_STOCK);
  const provenances = JSON.stringify(CONFIG.PROVENANCES);

  return `<!DOCTYPE html>
<html>
<head>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', sans-serif; }
  body { background: #0A0F1E; color: #E2E8F0; padding: 16px; }
  h2 { color: #E63946; font-size: 18px; margin-bottom: 16px; border-bottom: 2px solid #E63946; padding-bottom: 8px; }
  .section { background: #1A2744; border-radius: 8px; padding: 12px; margin-bottom: 12px; }
  .section-title { font-size: 11px; font-weight: bold; color: #94A3B8; text-transform: uppercase; margin-bottom: 10px; letter-spacing: 1px; }
  label { display: block; font-size: 12px; color: #94A3B8; margin-bottom: 3px; margin-top: 8px; }
  input, select, textarea { width: 100%; padding: 8px 10px; background: #0A0F1E; border: 1px solid #334155; border-radius: 5px; color: #E2E8F0; font-size: 13px; }
  input:focus, select:focus { outline: none; border-color: #E63946; }
  .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .btn { width: 100%; padding: 12px; background: #E63946; color: white; border: none; border-radius: 6px; font-size: 15px; font-weight: bold; cursor: pointer; margin-top: 12px; }
  .btn:hover { background: #C62828; }
  .id-badge { background: #E63946; color: white; padding: 4px 10px; border-radius: 4px; font-weight: bold; font-size: 13px; }
  .success { background: #14532D; color: #4ADE80; padding: 10px; border-radius: 5px; text-align: center; display: none; margin-top: 10px; }
</style>
</head>
<body>
<h2>🚗 Nouveau Véhicule</h2>
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
  <span style="font-size:12px;color:#94A3B8;">ID Dossier</span>
  <span class="id-badge" id="idDossier">${nextID}</span>
</div>

<div class="section">
  <div class="section-title">🔍 Identification</div>
  <label>Immatriculation *</label>
  <input type="text" id="immat" placeholder="AB-123-CD" style="text-transform:uppercase;">
  <label>VIN</label>
  <input type="text" id="vin" placeholder="VF1AG0...">
  <div class="row2">
    <div><label>Marque *</label><input type="text" id="marque" placeholder="Renault"></div>
    <div><label>Modèle *</label><input type="text" id="modele" placeholder="Clio"></div>
  </div>
  <div class="row2">
    <div><label>Version</label><input type="text" id="version" placeholder="1.5 dCi 90ch"></div>
    <div><label>Finition</label><input type="text" id="finition" placeholder="Intens"></div>
  </div>
  <div class="row2">
    <div><label>Année</label><input type="number" id="annee" placeholder="2021"></div>
    <div><label>Kilométrage</label><input type="number" id="km" placeholder="45000"></div>
  </div>
  <div class="row2">
    <div><label>Carburant</label><select id="carburant">${CONFIG.CARBURANTS.map(c=>`<option>${c}</option>`).join('')}</select></div>
    <div><label>Boîte</label><select id="boite">${CONFIG.BOITES.map(b=>`<option>${b}</option>`).join('')}</select></div>
  </div>
  <div class="row2">
    <div><label>Couleur</label><input type="text" id="couleur" placeholder="Gris titanium"></div>
    <div><label>Puissance fiscale</label><input type="number" id="cv" placeholder="6"></div>
  </div>
</div>

<div class="section">
  <div class="section-title">🛒 Achat</div>
  <div class="row2">
    <div><label>Provenance</label><select id="provenance">${CONFIG.PROVENANCES.map(p=>`<option>${p}</option>`).join('')}</select></div>
    <div><label>Fournisseur</label><input type="text" id="fournisseur"></div>
  </div>
  <div class="row2">
    <div><label>Date achat</label><input type="date" id="dateAchat"></div>
    <div><label>Prix achat (€)</label><input type="number" id="prixAchat" placeholder="0"></div>
  </div>
  <div class="row2">
    <div><label>Transport (€)</label><input type="number" id="transport" placeholder="0"></div>
    <div><label>Réparations (€)</label><input type="number" id="reparations" placeholder="0"></div>
  </div>
  <div class="row2">
    <div><label>Nettoyage (€)</label><input type="number" id="nettoyage" placeholder="0"></div>
    <div><label>Carte grise (€)</label><input type="number" id="cg" placeholder="0"></div>
  </div>
</div>

<div class="section">
  <div class="section-title">💰 Commercial</div>
  <div class="row2">
    <div><label>Prix affiché (€)</label><input type="number" id="prixAffiche" placeholder="0"></div>
    <div><label>Statut</label><select id="statut">${CONFIG.STATUTS_STOCK.map(s=>`<option>${s}</option>`).join('')}</select></div>
  </div>
  <label>Notes</label>
  <textarea id="notes" rows="2" placeholder="Remarques..."></textarea>
</div>

<button class="btn" onclick="sauvegarder()">✅ Enregistrer le véhicule</button>
<div class="success" id="successMsg">✅ Véhicule ajouté avec succès !</div>

<script>
function sauvegarder() {
  const data = {
    idDossier: document.getElementById('idDossier').innerText,
    immat: document.getElementById('immat').value.toUpperCase(),
    vin: document.getElementById('vin').value,
    marque: document.getElementById('marque').value,
    modele: document.getElementById('modele').value,
    version: document.getElementById('version').value,
    finition: document.getElementById('finition').value,
    annee: document.getElementById('annee').value,
    km: document.getElementById('km').value,
    carburant: document.getElementById('carburant').value,
    boite: document.getElementById('boite').value,
    statut: document.getElementById('statut').value,
    cv: document.getElementById('cv').value,
    couleur: document.getElementById('couleur').value,
    dateAchat: document.getElementById('dateAchat').value,
    provenance: document.getElementById('provenance').value,
    fournisseur: document.getElementById('fournisseur').value,
    prixAchat: document.getElementById('prixAchat').value || 0,
    transport: document.getElementById('transport').value || 0,
    reparations: document.getElementById('reparations').value || 0,
    nettoyage: document.getElementById('nettoyage').value || 0,
    cg: document.getElementById('cg').value || 0,
    prixAffiche: document.getElementById('prixAffiche').value || 0,
    notes: document.getElementById('notes').value,
  };
  if (!data.immat || !data.marque || !data.modele) {
    alert('⚠️ Immatriculation, Marque et Modèle sont obligatoires.');
    return;
  }
  google.script.run.withSuccessHandler(function() {
    document.getElementById('successMsg').style.display = 'block';
    setTimeout(() => google.script.host.close(), 1800);
  }).ajouterVehiculeDepuisFormulaire(data);
}
</script>
</body>
</html>`;
}

// Callback depuis le formulaire HTML
function ajouterVehiculeDepuisFormulaire(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('STOCK VO');
  const lastRow = Math.max(sheet.getLastRow(), 2) + 1;

  const dateAchat = data.dateAchat ? new Date(data.dateAchat) : '';

  // A B C D E F G H I J K L M N O P Q R S T U V
  const values = [
    data.idDossier,   // A - ID
    data.immat,       // B - Immat
    data.vin,         // C - VIN
    data.marque,      // D - Marque
    data.modele,      // E - Modèle
    data.version,     // F - Version
    data.finition,    // G - Finition
    data.annee ? parseInt(data.annee) : '', // H - Année
    data.km ? parseInt(data.km) : '',       // I - Km
    data.carburant,   // J - Carburant
    data.boite,       // K - Boîte
    data.statut,      // L - Statut
    data.cv ? parseInt(data.cv) : '', // M - CV
    '',               // N - Date MEC
    data.couleur,     // O - Couleur
    dateAchat,        // P - Date achat
    '',               // Q - Nb clés
    '',               // R - Nb proprio
    '',               // S - Historique
    '',               // T - Crit'Air
    '',               // U - CT
    '',               // V - Garantie
    data.provenance,  // W - Provenance
    data.fournisseur, // X - Fournisseur
    data.prixAchat ? parseFloat(data.prixAchat) : '', // Y - Prix achat
    data.transport ? parseFloat(data.transport) : '',   // Z - Transport
    data.reparations ? parseFloat(data.reparations) : '', // AA - Réparations
    data.nettoyage ? parseFloat(data.nettoyage) : '',   // AB - Nettoyage
    data.cg ? parseFloat(data.cg) : '',                 // AC - CG
    '',               // AD - Préparation
    '',               // AE - Divers
    // AF à AS = formules automatiques
    '', '', '', '', '', '', '', '', '', // AF-AN formules
    '',               // AO - Date vente
    '',               // AP - Prix vente
    '',               // AQ - Client
    '',               // AR - Jours stock
    data.notes,       // AS - Notes
  ];

  sheet.getRange(lastRow, 1, 1, values.length).setValues([values]);
  sheet.getRange(lastRow, 16).setNumberFormat('dd/mm/yyyy'); // Date achat

  // Appliquer les formules sur cette ligne
  const r = lastRow;
  sheet.getRange(r, 32).setFormula(`=IF(A${r}="","",SUM(Y${r}:AE${r}))`);
  sheet.getRange(r, 32).setNumberFormat('#,##0 "€"');
  sheet.getRange(r, 33).setFormula(`=IF(AF${r}="","",ROUND(AF${r}*1.20/100,0)*100)`);
  sheet.getRange(r, 33).setNumberFormat('#,##0 "€"');
  if (data.prixAffiche && parseFloat(data.prixAffiche) > 0) {
    sheet.getRange(r, 34).setValue(parseFloat(data.prixAffiche));
  } else {
    sheet.getRange(r, 34).setFormula(`=AG${r}`);
  }
  sheet.getRange(r, 34).setNumberFormat('#,##0 "€"');
  sheet.getRange(r, 35).setFormula(`=IF(AF${r}="","",ROUND(AF${r}*1.05/100,0)*100)`);
  sheet.getRange(r, 35).setNumberFormat('#,##0 "€"');
  sheet.getRange(r, 36).setFormula(`=IF(OR(AH${r}="",AF${r}=""),"",AH${r}-AF${r})`);
  sheet.getRange(r, 36).setNumberFormat('#,##0 "€"');
  sheet.getRange(r, 37).setFormula(`=IF(OR(AF${r}="",AF${r}=0),"",AJ${r}/AF${r})`);
  sheet.getRange(r, 37).setNumberFormat('0.0"%"');
  sheet.getRange(r, 39).setFormula(`=IF(OR(AH${r}="",AF${r}=""),"",ROUND((AH${r}-AF${r})/1.20*0.20,2))`);
  sheet.getRange(r, 39).setNumberFormat('#,##0.00 "€"');
  sheet.getRange(r, 38).setFormula(`=IF(AJ${r}="","",AJ${r}-AM${r})`);
  sheet.getRange(r, 38).setNumberFormat('#,##0 "€"');
  sheet.getRange(r, 40).setFormula(`=IF(OR(AF${r}="",AF${r}=0),"",AL${r}/AF${r})`);
  sheet.getRange(r, 40).setNumberFormat('0.0"%"');
  sheet.getRange(r, 44).setFormula(`=IF(OR(A${r}="",P${r}=""),"",IF(AO${r}<>"",AO${r}-P${r},TODAY()-P${r}))`);
  sheet.getRange(r, 44).setNumberFormat('#,##0 "j"');

  // Copier dans ADMINISTRATIF automatiquement
  const sheetAdmin = ss.getSheetByName('ADMINISTRATIF');
  if (sheetAdmin) {
    const adminLastRow = Math.max(sheetAdmin.getLastRow(), 1) + 1;
    sheetAdmin.getRange(adminLastRow, 1, 1, 5).setValues([[
      data.idDossier,
      data.immat,
      `${data.marque} ${data.modele}`,
      data.vin,
      '❌ Non',
    ]]);
    sheetAdmin.getRange(adminLastRow, 16).setValue('📋 EN COURS');
  }

  majDashboard();
}

// ============================================================
// FORMULAIRE AJOUT CLIENT
// ============================================================
function ouvrirFormulaireClient() {
  const html = HtmlService.createHtmlOutput(getHTMLFormulaireClient())
    .setTitle('👥 Nouveau Client / Lead')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHTMLFormulaireClient() {
  const nextID = genererNumeroClient();
  return `<!DOCTYPE html>
<html>
<head>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', sans-serif; }
  body { background: #0A0F1E; color: #E2E8F0; padding: 16px; }
  h2 { color: #2DC653; font-size: 18px; margin-bottom: 16px; border-bottom: 2px solid #2DC653; padding-bottom: 8px; }
  .section { background: #1A2744; border-radius: 8px; padding: 12px; margin-bottom: 12px; }
  .section-title { font-size: 11px; font-weight: bold; color: #94A3B8; text-transform: uppercase; margin-bottom: 10px; letter-spacing: 1px; }
  label { display: block; font-size: 12px; color: #94A3B8; margin-bottom: 3px; margin-top: 8px; }
  input, select, textarea { width: 100%; padding: 8px 10px; background: #0A0F1E; border: 1px solid #334155; border-radius: 5px; color: #E2E8F0; font-size: 13px; }
  input:focus, select:focus { outline: none; border-color: #2DC653; }
  .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .btn { width: 100%; padding: 12px; background: #2DC653; color: white; border: none; border-radius: 6px; font-size: 15px; font-weight: bold; cursor: pointer; margin-top: 12px; }
  .btn:hover { background: #16A34A; }
  .id-badge { background: #2DC653; color: white; padding: 4px 10px; border-radius: 4px; font-weight: bold; font-size: 13px; }
</style>
</head>
<body>
<h2>👥 Nouveau Lead / Client</h2>
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
  <span style="font-size:12px;color:#94A3B8;">ID Client</span>
  <span class="id-badge">${nextID}</span>
</div>

<div class="section">
  <div class="section-title">👤 Identité</div>
  <label>Nom Prénom *</label>
  <input type="text" id="nom" placeholder="Jean Dupont">
  <div class="row2">
    <div><label>Téléphone *</label><input type="tel" id="tel" placeholder="06 12 34 56 78"></div>
    <div><label>Email</label><input type="email" id="email" placeholder="jean@email.com"></div>
  </div>
  <div class="row2">
    <div><label>Ville</label><input type="text" id="ville" placeholder="Paris"></div>
    <div><label>Source</label><select id="source">
      ${['LeBonCoin','Autosphere','LaArgus','Bouche à oreille','Réseaux sociaux','Site web','Walk-in','Tel entrant','Partenaire','Autre'].map(s=>`<option>${s}</option>`).join('')}
    </select></div>
  </div>
</div>

<div class="section">
  <div class="section-title">🚗 Recherche</div>
  <label>Véhicule recherché</label>
  <input type="text" id="vehicule" placeholder="Peugeot 308 diesel, max 150 000km">
  <div class="row2">
    <div><label>Budget max (€)</label><input type="number" id="budget" placeholder="15000"></div>
    <div><label>Priorité</label><select id="priorite">
      <option>🔴 Urgente</option><option>🟠 Haute</option><option selected>🟡 Normale</option><option>🟢 Basse</option>
    </select></div>
  </div>
  <div class="row2">
    <div><label>Statut pipeline</label><select id="statut">
      ${CONFIG.PIPELINE_ETAPES.map(e=>`<option>${e}</option>`).join('')}
    </select></div>
    <div><label>Score</label><select id="score">
      <option>🔥 Chaud</option><option selected>🌡️ Tiède</option><option>❄️ Froid</option>
    </select></div>
  </div>
  <label>Notes</label>
  <textarea id="notes" rows="2" placeholder="Besoin urgent, dossier financement prêt..."></textarea>
</div>

<button class="btn" onclick="sauvegarder()">✅ Ajouter le client</button>

<script>
function sauvegarder() {
  const data = {
    idClient: '${nextID}',
    nom: document.getElementById('nom').value,
    tel: document.getElementById('tel').value,
    email: document.getElementById('email').value,
    ville: document.getElementById('ville').value,
    source: document.getElementById('source').value,
    vehicule: document.getElementById('vehicule').value,
    statut: document.getElementById('statut').value,
    score: document.getElementById('score').value,
    budget: document.getElementById('budget').value || 0,
    priorite: document.getElementById('priorite').value,
    notes: document.getElementById('notes').value,
  };
  if (!data.nom || !data.tel) {
    alert('⚠️ Nom et téléphone sont obligatoires.');
    return;
  }
  google.script.run.withSuccessHandler(function() {
    alert('✅ Client ajouté !');
    google.script.host.close();
  }).ajouterClientDepuisFormulaire(data);
}
</script>
</body>
</html>`;
}

function ajouterClientDepuisFormulaire(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CRM CLIENTS');
  const lastRow = Math.max(sheet.getLastRow(), 1) + 1;

  sheet.getRange(lastRow, 1, 1, 19).setValues([[
    data.idClient,    // A - ID
    data.nom,         // B - Nom
    data.tel,         // C - Tel
    data.email,       // D - Email
    data.ville,       // E - Ville
    data.source,      // F - Source
    data.vehicule,    // G - Véhicule recherché
    data.statut,      // H - Statut pipeline
    data.score,       // I - Score
    data.budget ? parseFloat(data.budget) : '', // J - Budget
    new Date(),       // K - Date entrée
    new Date(),       // L - Dernier contact
    '',               // M - Prochain RDV
    1,                // N - Nb contacts
    '',               // O - Véhicule attribué
    '',               // P - Vendeur
    data.notes,       // Q - Notes
    '',               // R - Relance auto
    data.priorite,    // S - Priorité
  ]]);

  sheet.getRange(lastRow, 11).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(lastRow, 12).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(lastRow, 10).setNumberFormat('#,##0 "€"');
}

// ============================================================
// GÉNÉRATION BON DE COMMANDE (texte dans sheet dédié)
// ============================================================
function genererBonCommande() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const stockSheet = ss.getSheetByName('STOCK VO');
  const activeCell = stockSheet.getActiveCell();
  const row = activeCell.getRow();

  if (row < 3) {
    ui.alert('⚠️ Sélectionnez', 'Positionnez-vous sur la ligne du véhicule dans STOCK VO avant de générer le bon de commande.', ui.ButtonSet.OK);
    return;
  }

  const idDossier = stockSheet.getRange(row, 1).getValue();
  if (!idDossier) {
    ui.alert('⚠️ Ligne vide', 'La ligne sélectionnée est vide.', ui.ButtonSet.OK);
    return;
  }

  // Récupérer les données
  const data = {
    idDossier: stockSheet.getRange(row, 1).getValue(),
    immat: stockSheet.getRange(row, 2).getValue(),
    marque: stockSheet.getRange(row, 4).getValue(),
    modele: stockSheet.getRange(row, 5).getValue(),
    version: stockSheet.getRange(row, 6).getValue(),
    annee: stockSheet.getRange(row, 8).getValue(),
    km: stockSheet.getRange(row, 9).getValue(),
    vin: stockSheet.getRange(row, 3).getValue(),
    prixAffiche: stockSheet.getRange(row, 34).getValue(),
    coutRevient: stockSheet.getRange(row, 32).getValue(),
    margeNette: stockSheet.getRange(row, 38).getValue(),
    tvaMarge: stockSheet.getRange(row, 39).getValue(),
  };

  // Demander le nom du client
  const reponseClient = ui.prompt('👤 Client', 'Nom du client pour ce bon de commande :', ui.ButtonSet.OK_CANCEL);
  if (reponseClient.getSelectedButton() !== ui.Button.OK) return;
  const nomClient = reponseClient.getResponseText() || 'Client';

  const numBC = genererNumeroCommande();
  const dateAujourdhi = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  // Créer/accéder à une feuille temporaire BON DE COMMANDE
  let bcSheet = ss.getSheetByName('BON COMMANDE TEMP');
  if (bcSheet) ss.deleteSheet(bcSheet);
  bcSheet = ss.insertSheet('BON COMMANDE TEMP');
  bcSheet.setTabColor(CONFIG.COLORS.ACCENT2);
  bcSheet.setHiddenGridlines(true);

  // Récupérer infos garage depuis PARAMÈTRES
  const paramSheet = ss.getSheetByName('PARAMÈTRES');
  let garageNom = 'Mon Garage VO';
  let garageSiret = '';
  let garageAdresse = '';
  let garageVille = '';
  let garageTel = '';
  let garageEmail = '';
  if (paramSheet) {
    // Les paramètres sont en colonne C, labels en B
    const paramData = paramSheet.getDataRange().getValues();
    paramData.forEach(row => {
      if (row[1] === 'Nom du garage') garageNom = row[2] || garageNom;
      if (row[1] === 'SIRET') garageSiret = row[2] || '';
      if (row[1] === 'Adresse') garageAdresse = row[2] || '';
      if (row[1] === 'Ville') garageVille = row[2] || '';
      if (row[1] === 'Téléphone') garageTel = row[2] || '';
      if (row[1] === 'Email') garageEmail = row[2] || '';
    });
  }

  // Construire le bon de commande visuellement
  const col = (n) => String.fromCharCode(64 + n);

  // Largeurs colonnes
  bcSheet.setColumnWidth(1, 15);
  bcSheet.setColumnWidth(2, 180);
  bcSheet.setColumnWidth(3, 130);
  bcSheet.setColumnWidth(4, 130);
  bcSheet.setColumnWidth(5, 130);
  bcSheet.setColumnWidth(6, 130);
  bcSheet.setColumnWidth(7, 15);

  let r = 1;

  const setCell = (row, col, value, options = {}) => {
    const cell = bcSheet.getRange(row, col);
    cell.setValue(value);
    if (options.bg) cell.setBackground(options.bg);
    if (options.color) cell.setFontColor(options.color);
    if (options.bold) cell.setFontWeight('bold');
    if (options.size) cell.setFontSize(options.size);
    if (options.align) cell.setHorizontalAlignment(options.align);
    if (options.valign) cell.setVerticalAlignment(options.valign);
    if (options.format) cell.setNumberFormat(options.format);
    if (options.wrap) cell.setWrap(true);
    if (options.indent) cell.setIndent(options.indent);
  };

  const mergeCells = (r1, c1, r2, c2) => bcSheet.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1).merge();

  // HEADER
  bcSheet.setRowHeight(r, 15); r++;
  bcSheet.setRowHeight(r, 70);
  mergeCells(r, 2, r, 4);
  setCell(r, 2, garageNom, { bg: CONFIG.COLORS.PRIMARY, color: CONFIG.COLORS.WHITE, bold: true, size: 20, align: 'left', valign: 'middle', indent: 2 });
  mergeCells(r, 5, r, 6);
  setCell(r, 5, `N° ${numBC}\n${dateAujourdhi}`, { bg: CONFIG.COLORS.ACCENT, color: CONFIG.COLORS.WHITE, bold: true, size: 12, align: 'center', valign: 'middle', wrap: true });
  r++;

  bcSheet.setRowHeight(r, 35);
  mergeCells(r, 2, r, 4);
  setCell(r, 2, `${garageAdresse} — ${garageVille} | Tél: ${garageTel} | ${garageEmail}`, {
    bg: CONFIG.COLORS.SECONDARY, color: '#94A3B8', size: 10, align: 'left', valign: 'middle', indent: 2
  });
  mergeCells(r, 5, r, 6);
  setCell(r, 5, `SIRET: ${garageSiret}`, { bg: CONFIG.COLORS.SECONDARY, color: '#94A3B8', size: 10, align: 'center', valign: 'middle' });
  r++;

  bcSheet.setRowHeight(r, 20);
  mergeCells(r, 2, r, 6);
  setCell(r, 2, 'BON DE COMMANDE', { bg: CONFIG.COLORS.ACCENT, color: CONFIG.COLORS.WHITE, bold: true, size: 14, align: 'center', valign: 'middle' });
  r++;

  bcSheet.setRowHeight(r, 15); r++;

  // INFO CLIENT / VÉHICULE côte à côte
  bcSheet.setRowHeight(r, 25);
  mergeCells(r, 2, r, 3);
  setCell(r, 2, '👤 ACHETEUR', { bg: CONFIG.COLORS.SECONDARY, color: CONFIG.COLORS.WHITE, bold: true, size: 10, align: 'center', valign: 'middle' });
  mergeCells(r, 4, r, 6);
  setCell(r, 4, '🚗 VÉHICULE', { bg: CONFIG.COLORS.SECONDARY, color: CONFIG.COLORS.WHITE, bold: true, size: 10, align: 'center', valign: 'middle' });
  r++;

  const clientLines = [
    ['Nom', nomClient],
    ['Dossier VO', data.idDossier],
    ['Date commande', dateAujourdhi],
  ];
  const vehiculeLines = [
    ['Immatriculation', data.immat],
    ['Marque / Modèle', `${data.marque} ${data.modele}`],
    ['Version', data.version || '-'],
    ['Année', data.annee || '-'],
    ['Kilométrage', data.km ? `${Number(data.km).toLocaleString('fr-FR')} km` : '-'],
    ['VIN', data.vin || '-'],
  ];

  const maxLines = Math.max(clientLines.length, vehiculeLines.length);
  for (let i = 0; i < maxLines; i++) {
    bcSheet.setRowHeight(r, 26);
    if (clientLines[i]) {
      setCell(r, 2, clientLines[i][0], { bg: CONFIG.COLORS.LIGHT_BG, bold: true, size: 10, valign: 'middle', indent: 1 });
      setCell(r, 3, clientLines[i][1], { bg: CONFIG.COLORS.WHITE, size: 10, valign: 'middle', indent: 1 });
    }
    if (vehiculeLines[i]) {
      setCell(r, 4, vehiculeLines[i][0], { bg: CONFIG.COLORS.LIGHT_BG, bold: true, size: 10, valign: 'middle', indent: 1 });
      mergeCells(r, 5, r, 6);
      setCell(r, 5, vehiculeLines[i][1], { bg: CONFIG.COLORS.WHITE, size: 10, valign: 'middle', indent: 1 });
    }
    r++;
  }

  bcSheet.setRowHeight(r, 15); r++;

  // TABLEAU PRIX
  bcSheet.setRowHeight(r, 28);
  mergeCells(r, 2, r, 6);
  setCell(r, 2, '💰 CONDITIONS FINANCIÈRES', { bg: CONFIG.COLORS.PRIMARY, color: CONFIG.COLORS.WHITE, bold: true, size: 11, align: 'center', valign: 'middle' });
  r++;

  const lignesPrix = [
    ['Prix de vente TTC', data.prixAffiche, '#,##0 "€"', false],
    ['Dont TVA sur marge', data.tvaMarge, '#,##0.00 "€"', false],
    ['Acompte versé', 0, '#,##0 "€"', false],
    ['Solde à payer', data.prixAffiche, '#,##0 "€"', true],
  ];

  lignesPrix.forEach(([label, val, fmt, isSolde]) => {
    bcSheet.setRowHeight(r, 30);
    mergeCells(r, 2, r, 4);
    setCell(r, 2, label, {
      bg: isSolde ? CONFIG.COLORS.PRIMARY : CONFIG.COLORS.LIGHT_BG,
      color: isSolde ? CONFIG.COLORS.WHITE : CONFIG.COLORS.DARK_TEXT,
      bold: isSolde, size: isSolde ? 12 : 10, valign: 'middle', indent: 2
    });
    mergeCells(r, 5, r, 6);
    const cell = bcSheet.getRange(r, 5);
    cell.setValue(val);
    cell.setNumberFormat(fmt);
    cell.setBackground(isSolde ? CONFIG.COLORS.ACCENT : CONFIG.COLORS.WHITE);
    cell.setFontColor(isSolde ? CONFIG.COLORS.WHITE : CONFIG.COLORS.DARK_TEXT);
    cell.setFontWeight(isSolde ? 'bold' : 'normal');
    cell.setFontSize(isSolde ? 14 : 11);
    cell.setHorizontalAlignment('right');
    cell.setVerticalAlignment('middle');
    r++;
  });

  bcSheet.setRowHeight(r, 15); r++;

  // MENTIONS LÉGALES
  bcSheet.setRowHeight(r, 50);
  mergeCells(r, 2, r, 6);
  setCell(r, 2,
    'Le présent bon de commande est établi conformément aux dispositions légales en vigueur. ' +
    'TVA calculée sur la marge conformément à l\'article 297 A du CGI. ' +
    'Garantie légale de conformité et des vices cachés applicable.',
    { bg: '#FFFBEB', color: '#92400E', size: 9, valign: 'middle', wrap: true, indent: 2 }
  );
  r++;

  // SIGNATURES
  bcSheet.setRowHeight(r, 15); r++;
  bcSheet.setRowHeight(r, 60);
  mergeCells(r, 2, r, 3);
  setCell(r, 2, 'Signature Vendeur :\n\n\n_______________________', { bg: CONFIG.COLORS.LIGHT_BG, size: 10, valign: 'top', wrap: true, indent: 2 });
  mergeCells(r, 4, r, 6);
  setCell(r, 4, 'Signature Acheteur (Lu et approuvé) :\n\n\n_______________________', { bg: CONFIG.COLORS.LIGHT_BG, size: 10, valign: 'top', wrap: true, indent: 2 });
  r++;

  // Activer la feuille
  ss.setActiveSheet(bcSheet);

  ui.alert(
    '✅ Bon de commande créé',
    `Le bon ${numBC} a été créé dans l\'onglet "BON COMMANDE TEMP".\n\nPour imprimer : Fichier > Imprimer (ou Ctrl+P).\nPour PDF : Fichier > Télécharger > PDF.`,
    ui.ButtonSet.OK
  );
}

// ============================================================
// GÉNÉRATION FACTURE
// ============================================================
function genererFacture() {
  SpreadsheetApp.getUi().alert(
    '🧾 Générer Facture',
    'Fonctionnement identique au Bon de Commande.\n\nPositionnez-vous sur la ligne du véhicule vendu dans STOCK VO puis utilisez Bon de Commande.\n\nLa numérotation FA-XXXX est disponible via genererNumeroFacture() dans le script.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================
// RECHERCHE AVANCÉE
// ============================================================
function rechercheAvancee() {
  const html = HtmlService.createHtmlOutput(getHTMLRecherche())
    .setTitle('🔍 Recherche Stock VO')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getHTMLRecherche() {
  return `<!DOCTYPE html>
<html>
<head>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', sans-serif; }
  body { background: #0A0F1E; color: #E2E8F0; padding: 16px; }
  h2 { color: #4895EF; font-size: 16px; margin-bottom: 16px; }
  label { display: block; font-size: 12px; color: #94A3B8; margin-bottom: 3px; margin-top: 8px; }
  input, select { width: 100%; padding: 8px 10px; background: #1A2744; border: 1px solid #334155; border-radius: 5px; color: #E2E8F0; font-size: 13px; }
  .btn { width: 100%; padding: 11px; background: #4895EF; color: white; border: none; border-radius: 6px; font-size: 14px; font-weight: bold; cursor: pointer; margin-top: 14px; }
  .result { background: #1A2744; border-radius: 6px; padding: 10px; margin-top: 12px; font-size: 12px; min-height: 80px; }
  .item { background: #0A0F1E; border-radius: 5px; padding: 8px; margin-bottom: 6px; border-left: 3px solid #4895EF; }
  .item strong { color: #4895EF; font-size: 13px; }
</style>
</head>
<body>
<h2>🔍 Recherche Stock VO</h2>
<label>Marque</label>
<input type="text" id="marque" placeholder="Renault, Peugeot...">
<label>Carburant</label>
<select id="carburant">
  <option value="">Tous</option>
  ${CONFIG.CARBURANTS.map(c=>`<option>${c}</option>`).join('')}
</select>
<label>Statut</label>
<select id="statut">
  <option value="">Tous</option>
  ${CONFIG.STATUTS_STOCK.map(s=>`<option>${s}</option>`).join('')}
</select>
<label>Budget max (€)</label>
<input type="number" id="budgetMax" placeholder="20000">
<label>Km max</label>
<input type="number" id="kmMax" placeholder="100000">
<button class="btn" onclick="rechercher()">🔍 Rechercher</button>
<div class="result" id="results"><span style="color:#94A3B8">Les résultats apparaîtront ici...</span></div>

<script>
function rechercher() {
  const params = {
    marque: document.getElementById('marque').value,
    carburant: document.getElementById('carburant').value,
    statut: document.getElementById('statut').value,
    budgetMax: document.getElementById('budgetMax').value,
    kmMax: document.getElementById('kmMax').value,
  };
  document.getElementById('results').innerHTML = '<span style="color:#94A3B8">Recherche en cours...</span>';
  google.script.run.withSuccessHandler(afficherResultats).rechercherStock(params);
}
function afficherResultats(resultats) {
  const div = document.getElementById('results');
  if (!resultats || resultats.length === 0) {
    div.innerHTML = '<span style="color:#F87171">Aucun véhicule trouvé.</span>';
    return;
  }
  div.innerHTML = resultats.map(v =>
    '<div class="item"><strong>' + v.marque + ' ' + v.modele + '</strong><br>' +
    v.immat + ' — ' + v.annee + ' — ' + Number(v.km).toLocaleString('fr-FR') + ' km<br>' +
    '<span style="color:#4ADE80;font-weight:bold">' + Number(v.prix).toLocaleString('fr-FR') + ' €</span>' +
    ' — ' + v.statut + '</div>'
  ).join('');
}
</script>
</body>
</html>`;
}

function rechercherStock(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('STOCK VO');
  const data = sheet.getDataRange().getValues();
  const resultats = [];

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Ligne vide

    const marque = String(row[3] || '').toLowerCase();
    const modele = String(row[4] || '');
    const carburant = String(row[9] || '');
    const statut = String(row[11] || '');
    const km = parseFloat(row[8]) || 0;
    const prix = parseFloat(row[33]) || 0;
    const annee = row[7] || '';

    // Filtres
    if (params.marque && !marque.includes(params.marque.toLowerCase())) continue;
    if (params.carburant && carburant !== params.carburant) continue;
    if (params.statut && statut !== params.statut) continue;
    if (params.budgetMax && prix > parseFloat(params.budgetMax)) continue;
    if (params.kmMax && km > parseFloat(params.kmMax)) continue;

    resultats.push({
      idDossier: row[0],
      immat: row[1],
      marque: row[3],
      modele,
      annee,
      km,
      carburant,
      statut,
      prix,
    });

    if (resultats.length >= 20) break; // Limiter à 20 résultats
  }

  return resultats;
}

// ============================================================
// ALERTES
// ============================================================
function voirVehiculesDormants() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('STOCK VO');
  const data = sheet.getDataRange().getValues();
  const dormants = [];
  const today = new Date();

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const statut = row[11];
    if (statut === 'Vendu' || statut === 'Livré' || !statut) continue;
    const dateAchat = row[15];
    if (!dateAchat || !(dateAchat instanceof Date)) continue;
    const jours = Math.floor((today - dateAchat) / 86400000);
    if (jours >= CONFIG.SEUIL_DORMANT) {
      dormants.push(`${row[0]} — ${row[3]} ${row[4]} — ${row[1]} — ${jours} jours en stock (Statut: ${statut})`);
    }
  }

  if (dormants.length === 0) {
    SpreadsheetApp.getUi().alert('✅ Aucun véhicule dormant', `Tous les véhicules sont en stock depuis moins de ${CONFIG.SEUIL_DORMANT} jours.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert(
      `⚠️ ${dormants.length} véhicule(s) dormant(s)`,
      dormants.slice(0, 15).join('\n') + (dormants.length > 15 ? `\n... et ${dormants.length - 15} autres.` : ''),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function voirLeadsChauds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CRM CLIENTS');
  const data = sheet.getDataRange().getValues();
  const chauds = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (row[8] === '🔥 Chaud' && row[7] !== 'Vendu' && row[7] !== 'Perdu') {
      chauds.push(`${row[0]} — ${row[1]} — ${row[2]} — ${row[7]} — ${row[6]}`);
    }
  }

  SpreadsheetApp.getUi().alert(
    `🔥 ${chauds.length} lead(s) chaud(s)`,
    chauds.length > 0 ? chauds.join('\n') : 'Aucun lead chaud actuellement.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function voirDossiersIncomplets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ADMINISTRATIF');
  const data = sheet.getDataRange().getValues();
  const incomplets = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (row[15] === '⚠️ INCOMPLET' || row[15] === '❌ BLOQUANT') {
      incomplets.push(`${row[0]} — ${row[1]} — ${row[2]} — [${row[15]}]`);
    }
  }

  SpreadsheetApp.getUi().alert(
    `📋 ${incomplets.length} dossier(s) incomplet(s)`,
    incomplets.length > 0 ? incomplets.join('\n') : 'Tous les dossiers sont complets ✅',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function voirAlertesCT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ADMINISTRATIF');
  const data = sheet.getDataRange().getValues();
  const alertes = [];
  const today = new Date();
  const in30j = new Date(today.getTime() + 30 * 86400000);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || !row[8]) continue;
    const ctDate = row[8];
    if (!(ctDate instanceof Date)) continue;
    const jours = Math.floor((ctDate - today) / 86400000);
    if (jours < 30) {
      const label = jours < 0 ? `EXPIRÉ (${Math.abs(jours)} j)` : `Expire dans ${jours} j`;
      alertes.push(`${row[0]} — ${row[1]} — ${row[2]} — CT: ${Utilities.formatDate(ctDate, Session.getScriptTimeZone(), 'dd/MM/yyyy')} [${label}]`);
    }
  }

  SpreadsheetApp.getUi().alert(
    `⚠️ ${alertes.length} alerte(s) CT`,
    alertes.length > 0 ? alertes.join('\n') : 'Aucune alerte CT dans les 30 prochains jours ✅',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function voirCERFA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName('ADMINISTRATIF'));
}

// ============================================================
// GÉNÉRATION FICHE VÉHICULE
// ============================================================
function genererFicheVehicule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const stockSheet = ss.getSheetByName('STOCK VO');
  const row = stockSheet.getActiveCell().getRow();

  if (row < 3 || !stockSheet.getRange(row, 1).getValue()) {
    ui.alert('⚠️', 'Sélectionnez une ligne de véhicule dans STOCK VO.', ui.ButtonSet.OK);
    return;
  }

  const d = {
    id: stockSheet.getRange(row, 1).getValue(),
    immat: stockSheet.getRange(row, 2).getValue(),
    vin: stockSheet.getRange(row, 3).getValue(),
    marque: stockSheet.getRange(row, 4).getValue(),
    modele: stockSheet.getRange(row, 5).getValue(),
    version: stockSheet.getRange(row, 6).getValue(),
    finition: stockSheet.getRange(row, 7).getValue(),
    annee: stockSheet.getRange(row, 8).getValue(),
    km: stockSheet.getRange(row, 9).getValue(),
    carburant: stockSheet.getRange(row, 10).getValue(),
    boite: stockSheet.getRange(row, 11).getValue(),
    statut: stockSheet.getRange(row, 12).getValue(),
    cv: stockSheet.getRange(row, 13).getValue(),
    couleur: stockSheet.getRange(row, 15).getValue(),
    garantie: stockSheet.getRange(row, 22).getValue(),
    prixAffiche: stockSheet.getRange(row, 34).getValue(),
    coutRevient: stockSheet.getRange(row, 32).getValue(),
    margeNette: stockSheet.getRange(row, 38).getValue(),
    jours: stockSheet.getRange(row, 44).getValue(),
    notes: stockSheet.getRange(row, 45).getValue(),
  };

  let fiche = ss.getSheetByName('FICHE VÉHICULE TEMP');
  if (fiche) ss.deleteSheet(fiche);
  fiche = ss.insertSheet('FICHE VÉHICULE TEMP');
  fiche.setHiddenGridlines(true);
  fiche.setTabColor(CONFIG.COLORS.INFO);

  fiche.setColumnWidth(1, 20);
  fiche.setColumnWidth(2, 200);
  fiche.setColumnWidth(3, 200);
  fiche.setColumnWidth(4, 200);
  fiche.setColumnWidth(5, 20);

  let r = 1;
  fiche.setRowHeight(r, 10); r++;

  // Header
  fiche.setRowHeight(r, 70);
  fiche.getRange(r, 2, 1, 3).merge();
  fiche.getRange(r, 2).setValue(`${d.marque.toUpperCase()} ${d.modele.toUpperCase()}\n${d.version} — ${d.finition}`);
  fiche.getRange(r, 2).setBackground(CONFIG.COLORS.PRIMARY);
  fiche.getRange(r, 2).setFontColor(CONFIG.COLORS.WHITE);
  fiche.getRange(r, 2).setFontSize(20);
  fiche.getRange(r, 2).setFontWeight('bold');
  fiche.getRange(r, 2).setHorizontalAlignment('center');
  fiche.getRange(r, 2).setVerticalAlignment('middle');
  fiche.getRange(r, 2).setWrap(true);
  r++;

  // Immat & statut
  fiche.setRowHeight(r, 32);
  fiche.getRange(r, 2).setValue(`🔖 ${d.immat}`);
  fiche.getRange(r, 2).setBackground(CONFIG.COLORS.ACCENT);
  fiche.getRange(r, 2).setFontColor(CONFIG.COLORS.WHITE);
  fiche.getRange(r, 2).setFontSize(14);
  fiche.getRange(r, 2).setFontWeight('bold');
  fiche.getRange(r, 2).setHorizontalAlignment('center');
  fiche.getRange(r, 2).setVerticalAlignment('middle');

  fiche.getRange(r, 3).setValue(`📦 ${d.statut}`);
  fiche.getRange(r, 3).setBackground(CONFIG.COLORS.STATUTS_COULEURS ? '#2DC653' : '#2DC653');
  fiche.getRange(r, 3).setFontColor(CONFIG.COLORS.WHITE);
  fiche.getRange(r, 3).setFontSize(12);
  fiche.getRange(r, 3).setFontWeight('bold');
  fiche.getRange(r, 3).setHorizontalAlignment('center');
  fiche.getRange(r, 3).setVerticalAlignment('middle');

  fiche.getRange(r, 4).setValue(`💰 ${Number(d.prixAffiche).toLocaleString('fr-FR')} €`);
  fiche.getRange(r, 4).setBackground(CONFIG.COLORS.GOLD);
  fiche.getRange(r, 4).setFontColor(CONFIG.COLORS.DARK_TEXT);
  fiche.getRange(r, 4).setFontSize(14);
  fiche.getRange(r, 4).setFontWeight('bold');
  fiche.getRange(r, 4).setHorizontalAlignment('center');
  fiche.getRange(r, 4).setVerticalAlignment('middle');
  r++;

  fiche.setRowHeight(r, 12); r++;

  // Caractéristiques
  fiche.setRowHeight(r, 28);
  fiche.getRange(r, 2, 1, 3).merge();
  fiche.getRange(r, 2).setValue('📋 CARACTÉRISTIQUES');
  fiche.getRange(r, 2).setBackground(CONFIG.COLORS.SECONDARY);
  fiche.getRange(r, 2).setFontColor(CONFIG.COLORS.WHITE);
  fiche.getRange(r, 2).setFontWeight('bold');
  fiche.getRange(r, 2).setFontSize(11);
  fiche.getRange(r, 2).setHorizontalAlignment('center');
  fiche.getRange(r, 2).setVerticalAlignment('middle');
  r++;

  const specs = [
    ['Année', d.annee, 'Kilométrage', d.km ? `${Number(d.km).toLocaleString('fr-FR')} km` : '-'],
    ['Carburant', d.carburant, 'Boîte', d.boite],
    ['Couleur', d.couleur, 'Puissance fiscale', d.cv ? `${d.cv} CV` : '-'],
    ['Garantie', d.garantie ? `${d.garantie} mois` : 'Sans', 'Jours en stock', d.jours ? `${d.jours} j` : '-'],
  ];

  specs.forEach(([l1, v1, l2, v2]) => {
    fiche.setRowHeight(r, 28);
    fiche.getRange(r, 2).setValue(l1);
    fiche.getRange(r, 2).setBackground(CONFIG.COLORS.LIGHT_BG);
    fiche.getRange(r, 2).setFontWeight('bold');
    fiche.getRange(r, 2).setFontSize(10);
    fiche.getRange(r, 2).setVerticalAlignment('middle');
    fiche.getRange(r, 2).setIndent(1);

    fiche.getRange(r, 3).setValue(v1);
    fiche.getRange(r, 3).setBackground(CONFIG.COLORS.WHITE);
    fiche.getRange(r, 3).setFontSize(11);
    fiche.getRange(r, 3).setVerticalAlignment('middle');
    fiche.getRange(r, 3).setIndent(1);

    fiche.getRange(r, 4).setValue(`${l2}: ${v2}`);
    fiche.getRange(r, 4).setBackground(CONFIG.COLORS.LIGHT_BG);
    fiche.getRange(r, 4).setFontSize(10);
    fiche.getRange(r, 4).setVerticalAlignment('middle');
    fiche.getRange(r, 4).setIndent(1);
    r++;
  });

  // Notes
  if (d.notes) {
    fiche.setRowHeight(r, 12); r++;
    fiche.setRowHeight(r, 50);
    fiche.getRange(r, 2, 1, 3).merge();
    fiche.getRange(r, 2).setValue(`📝 Notes : ${d.notes}`);
    fiche.getRange(r, 2).setBackground('#FFFBEB');
    fiche.getRange(r, 2).setFontSize(10);
    fiche.getRange(r, 2).setVerticalAlignment('middle');
    fiche.getRange(r, 2).setWrap(true);
    fiche.getRange(r, 2).setIndent(1);
    r++;
  }

  // Infos confidentielles (internes)
  fiche.setRowHeight(r, 12); r++;
  fiche.setRowHeight(r, 28);
  fiche.getRange(r, 2, 1, 3).merge();
  fiche.getRange(r, 2).setValue('🔒 DONNÉES INTERNES (confidentiel)');
  fiche.getRange(r, 2).setBackground('#1F2937');
  fiche.getRange(r, 2).setFontColor('#9CA3AF');
  fiche.getRange(r, 2).setFontWeight('bold');
  fiche.getRange(r, 2).setFontSize(10);
  fiche.getRange(r, 2).setHorizontalAlignment('center');
  fiche.getRange(r, 2).setVerticalAlignment('middle');
  r++;

  [
    ['Coût de revient', `${Number(d.coutRevient).toLocaleString('fr-FR')} €`],
    ['Marge nette', `${Number(d.margeNette).toLocaleString('fr-FR')} €`],
    ['ID Dossier', d.id],
    ['VIN', d.vin],
  ].forEach(([l, v]) => {
    fiche.setRowHeight(r, 26);
    fiche.getRange(r, 2).setValue(l);
    fiche.getRange(r, 2).setBackground('#111827');
    fiche.getRange(r, 2).setFontColor('#6B7280');
    fiche.getRange(r, 2).setFontSize(10);
    fiche.getRange(r, 2).setVerticalAlignment('middle');
    fiche.getRange(r, 2).setIndent(1);
    fiche.getRange(r, 3, 1, 2).merge();
    fiche.getRange(r, 3).setValue(v);
    fiche.getRange(r, 3).setBackground('#111827');
    fiche.getRange(r, 3).setFontColor('#E5E7EB');
    fiche.getRange(r, 3).setFontSize(10);
    fiche.getRange(r, 3).setVerticalAlignment('middle');
    fiche.getRange(r, 3).setFontWeight('bold');
    fiche.getRange(r, 3).setIndent(1);
    r++;
  });

  ss.setActiveSheet(fiche);
  ui.alert('✅ Fiche générée', `Fiche du ${d.marque} ${d.modele} créée.\n\nImprimez via Fichier > Imprimer ou téléchargez en PDF.`, ui.ButtonSet.OK);
}

// ============================================================
// EMAILS & WHATSAPP
// ============================================================
function envoyerEmailRelance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('CRM CLIENTS');
  const row = sheet.getActiveCell().getRow();

  if (row < 2 || !sheet.getRange(row, 1).getValue()) {
    ui.alert('⚠️', 'Positionnez-vous sur la ligne du client dans CRM CLIENTS.', ui.ButtonSet.OK);
    return;
  }

  const nom = sheet.getRange(row, 2).getValue();
  const email = sheet.getRange(row, 4).getValue();
  const vehicule = sheet.getRange(row, 7).getValue();

  if (!email) {
    ui.alert('⚠️ Email manquant', `Le client ${nom} n'a pas d'email renseigné.`, ui.ButtonSet.OK);
    return;
  }

  const reponse = ui.alert(
    '📧 Confirmer l\'envoi',
    `Envoyer un email de relance à :\n${nom}\n${email}\n\nVéhicule recherché : ${vehicule || 'Non précisé'}`,
    ui.ButtonSet.YES_NO
  );

  if (reponse !== ui.Button.YES) return;

  try {
    // Récupérer nom du garage depuis PARAMÈTRES
    const paramSheet = ss.getSheetByName('PARAMÈTRES');
    let garageNom = 'Mon Garage VO';
    let garageTel = '';
    if (paramSheet) {
      paramSheet.getDataRange().getValues().forEach(r => {
        if (r[1] === 'Nom du garage') garageNom = r[2] || garageNom;
        if (r[1] === 'Téléphone') garageTel = r[2] || '';
      });
    }

    const sujet = `${garageNom} — Suite à votre recherche de véhicule`;
    const corps = `Bonjour ${nom.split(' ')[0] || nom},

Je me permets de vous recontacter suite à votre recherche de véhicule.

${vehicule ? `Vous êtes à la recherche de : ${vehicule}` : 'Vous souhaitez acquérir un véhicule d\'occasion.'}

Nous avons récemment fait évoluer notre stock et pourrions avoir le véhicule idéal pour vous.

N'hésitez pas à nous contacter ou à prendre rendez-vous pour un essai.

Cordialement,
${garageNom}
${garageTel ? `Tél : ${garageTel}` : ''}`;

    GmailApp.sendEmail(email, sujet, corps);

    // Mettre à jour le dernier contact
    sheet.getRange(row, 12).setValue(new Date());
    sheet.getRange(row, 14).setValue((sheet.getRange(row, 14).getValue() || 0) + 1);

    ui.alert('✅ Email envoyé', `Email de relance envoyé à ${nom} (${email}).`, ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Erreur', `Impossible d'envoyer l'email : ${e.message}`, ui.ButtonSet.OK);
  }
}

function genererMessageWhatsApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('CRM CLIENTS');
  const row = sheet.getActiveCell().getRow();

  if (row < 2) {
    ui.alert('⚠️', 'Positionnez-vous sur la ligne du client dans CRM CLIENTS.', ui.ButtonSet.OK);
    return;
  }

  const nom = sheet.getRange(row, 2).getValue();
  const tel = String(sheet.getRange(row, 5).getValue() || '').replace(/\s/g, '').replace(/^0/, '33');
  const vehicule = sheet.getRange(row, 7).getValue();

  const message = encodeURIComponent(
    `Bonjour ${nom.split(' ')[0] || nom} 👋\n\nJe me permets de vous recontacter concernant votre recherche : ${vehicule || 'véhicule d\'occasion'}.\n\nNous avons du nouveau dans notre stock, pouvons-nous en discuter ?\n\nCordialement 🚗`
  );

  const url = `https://wa.me/${tel}?text=${message}`;

  ui.alert(
    '💬 WhatsApp',
    `Lien WhatsApp généré pour ${nom} :\n\n${url}\n\n⚠️ Copiez ce lien dans votre navigateur pour ouvrir WhatsApp Web.`,
    ui.ButtonSet.OK
  );
}

function saisirReprise() {
  SpreadsheetApp.getUi().alert(
    '🔄 Saisir une reprise',
    'Pour saisir un véhicule en reprise :\n\n1. Allez dans STOCK VO\n2. Cliquez Menu > Ajouter un véhicule\n3. Sélectionnez Provenance = "Reprise"\n4. Renseignez les informations\n\nLa reprise sera automatiquement liée à la vente correspondante.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================
// RAPPORT MENSUEL
// ============================================================
function genererRapportMensuel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const ventes = ss.getSheetByName('VENTES');
  const stock = ss.getSheetByName('STOCK VO');

  const today = new Date();
  const moisActuel = today.getMonth() + 1;
  const anneeActuelle = today.getFullYear();
  const nomMois = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'][today.getMonth()];

  // Calculer les stats
  const ventesData = ventes.getDataRange().getValues();
  let caTotal = 0, nbVentes = 0, margeTotal = 0;

  for (let i = 1; i < ventesData.length; i++) {
    const row = ventesData[i];
    if (!row[0]) continue;
    const dateVente = row[4];
    if (!(dateVente instanceof Date)) continue;
    if (dateVente.getMonth() + 1 === moisActuel && dateVente.getFullYear() === anneeActuelle) {
      nbVentes++;
      caTotal += parseFloat(row[8]) || 0;
      margeTotal += parseFloat(row[12]) || 0;
    }
  }

  const stockData = stock.getDataRange().getValues();
  let nbStock = 0, valeurStock = 0;
  for (let i = 2; i < stockData.length; i++) {
    const row = stockData[i];
    if (!row[0]) continue;
    const statut = row[11];
    if (statut && statut !== 'Vendu' && statut !== 'Livré') {
      nbStock++;
      valeurStock += parseFloat(row[33]) || 0;
    }
  }

  const rapport = `📊 RAPPORT MENSUEL — ${nomMois} ${anneeActuelle}
${'═'.repeat(50)}

📦 STOCK
  • Véhicules en stock : ${nbStock}
  • Valeur stock total : ${caTotal > 0 ? valeurStock.toLocaleString('fr-FR') : '—'} €

💰 VENTES DU MOIS
  • Nombre de ventes : ${nbVentes}
  • Chiffre d'affaires : ${caTotal.toLocaleString('fr-FR')} €
  • Marge nette totale : ${margeTotal.toLocaleString('fr-FR')} €
  • Marge moyenne / véhicule : ${nbVentes > 0 ? Math.round(margeTotal / nbVentes).toLocaleString('fr-FR') : 0} €
  • Taux de marge moyen : ${caTotal > 0 ? (margeTotal / caTotal * 100).toFixed(1) : 0} %

📈 PERFORMANCE
  • CA objectif mensuel : [À définir dans PARAMÈTRES]
  • Rapport généré le : ${Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy à HH:mm')}`;

  ui.alert(`📊 Rapport ${nomMois}`, rapport, ui.ButtonSet.OK);
}

// ============================================================
// SAUVEGARDE DRIVE
// ============================================================
function sauvegarderDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    const fichierNom = `SpiderVO_Backup_${date}`;

    // Créer une copie
    const fichier = DriveApp.getFileById(ss.getId());
    const copie = fichier.makeCopy(fichierNom);

    // Déplacer dans un dossier SpiderVO si existant, sinon racine
    let dossier;
    try {
      const dossiers = DriveApp.getFoldersByName('SpiderVO Pro — Sauvegardes');
      if (dossiers.hasNext()) {
        dossier = dossiers.next();
      } else {
        dossier = DriveApp.createFolder('SpiderVO Pro — Sauvegardes');
      }
      dossier.addFile(copie);
      DriveApp.getRootFolder().removeFile(copie);
    } catch(e) {}

    ui.alert('✅ Sauvegarde réussie', `Fichier sauvegardé :\n${fichierNom}\n\nDossier : SpiderVO Pro — Sauvegardes`, ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Erreur sauvegarde', `${e.message}`, ui.ButtonSet.OK);
  }
}

// ============================================================
// DÉCLENCHEURS AUTOMATIQUES
// ============================================================
function configurerDeclencheurs() {
  // Supprimer les anciens déclencheurs
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'declencheurQuotidien' ||
        t.getHandlerFunction() === 'majDashboard') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Déclencheur quotidien à 8h
  ScriptApp.newTrigger('declencheurQuotidien')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // Mise à jour dashboard toutes les 30 minutes
  ScriptApp.newTrigger('majDashboard')
    .timeBased()
    .everyMinutes(30)
    .create();
}

function declencheurQuotidien() {
  // 1. Mise à jour dashboard
  majDashboard();

  // 2. Alertes email si configuré
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paramSheet = ss.getSheetByName('PARAMÈTRES');
  if (!paramSheet) return;

  let emailNotif = '';
  paramSheet.getDataRange().getValues().forEach(row => {
    if (row[1] === 'Email notifications') emailNotif = row[2] || '';
  });

  if (!emailNotif) return;

  // Véhicules dormants
  const stock = ss.getSheetByName('STOCK VO');
  const stockData = stock.getDataRange().getValues();
  const dormants = [];
  const today = new Date();

  for (let i = 2; i < stockData.length; i++) {
    const row = stockData[i];
    if (!row[0]) continue;
    const statut = row[11];
    if (statut === 'Vendu' || statut === 'Livré' || !statut) continue;
    const dateAchat = row[15];
    if (!(dateAchat instanceof Date)) continue;
    const jours = Math.floor((today - dateAchat) / 86400000);
    if (jours >= CONFIG.SEUIL_DORMANT) {
      dormants.push(`${row[0]} — ${row[3]} ${row[4]} — ${row[1]} — ${jours}j`);
    }
  }

  if (dormants.length > 0) {
    try {
      GmailApp.sendEmail(
        emailNotif,
        `⚠️ SpiderVO — ${dormants.length} véhicule(s) dormant(s)`,
        `Bonjour,\n\nVoici les véhicules en stock depuis plus de ${CONFIG.SEUIL_DORMANT} jours :\n\n${dormants.join('\n')}\n\nBonne journée !`
      );
    } catch(e) {
      Logger.log('Erreur envoi email : ' + e.message);
    }
  }
}

// ============================================================
// TEST EMAIL
// ============================================================
function testEmail() {
  const ui = SpreadsheetApp.getUi();
  try {
    const email = Session.getEffectiveUser().getEmail();
    GmailApp.sendEmail(email, '✅ Test SpiderVO Pro', 'Votre système SpiderVO Pro envoie bien les emails.\n\nBonne gestion !');
    ui.alert('✅ Email test envoyé', `Un email a été envoyé à : ${email}`, ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Erreur', `Impossible d'envoyer l'email : ${e.message}`, ui.ButtonSet.OK);
  }
}

// ============================================================
// À PROPOS
// ============================================================
function aPropos() {
  SpreadsheetApp.getUi().alert(
    `🚗 SpiderVO Pro — v${CONFIG.VERSION}`,
    `Système de Gestion de Véhicules d'Occasion\nInspired by SpiderVO, AutoCERFA, PlanetVO\n\nFonctionnalités :\n✅ Stock VO complet (5000+ véhicules)\n✅ CRM Clients & Pipeline commercial\n✅ Ventes, Factures, Bons de commande\n✅ Administratif type AutoCERFA\n✅ Finance & Rentabilité\n✅ Alertes automatiques\n✅ Formulaires intégrés\n✅ Sauvegarde Drive automatique\n✅ Notifications email\n\n© ${new Date().getFullYear()} SpiderVO Pro`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
