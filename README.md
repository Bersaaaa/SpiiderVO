# 🚗 SpiderVO Pro — Système de Gestion VO

> Remplaçant open-source de SpiderVO / PlanetVO / AutoCERFA — 100% Google Sheets + Excel + Apps Script

![Version](https://img.shields.io/badge/version-3.0-blue)
![Licence](https://img.shields.io/badge/licence-MIT-green)
![Sheets](https://img.shields.io/badge/Google%20Sheets-compatible-brightgreen?logo=google-sheets)
![Excel](https://img.shields.io/badge/Excel-compatible-217346?logo=microsoft-excel)
![Formules](https://img.shields.io/badge/formules-9%20522-orange)
![Erreurs](https://img.shields.io/badge/erreurs-0-success)

-----

## 📌 Présentation

**SpiderVO Pro** est un système complet de gestion de véhicules d’occasion conçu pour les **négociants automobiles français**. Il remplace les logiciels payants du marché (SpiderVO, PlanetVO, AutoCERFA) par une solution cloud gratuite basée sur **Google Sheets + Apps Script** et un fichier **Excel** standalone.

### Pourquoi SpiderVO Pro ?

|Fonctionnalité        |SpiderVO (payant)|SpiderVO Pro (gratuit)|
|----------------------|-----------------|----------------------|
|Gestion du stock VO   |✅                |✅                     |
|CRM clients & pipeline|✅                |✅                     |
|Calcul TVA sur marge  |✅                |✅                     |
|Suivi CERFA / admin   |✅                |✅                     |
|Bons de commande PDF  |✅                |✅                     |
|Alertes automatiques  |✅                |✅                     |
|Hébergement cloud     |Payant           |Google Drive (gratuit)|
|Personnalisable       |❌                |✅                     |
|Prix                  |~150€/mois       |**Gratuit**           |

-----

## 📁 Contenu du dépôt

```
spidervo-pro/
├── Code.gs                   # Script Google Apps Script complet
├── SpiderVO_Pro_v3.xlsx      # Fichier Excel standalone (offline)
└── README.md
```

-----

## 🚀 Installation — Google Sheets (recommandé)

### Prérequis

- Un compte Google (gratuit)
- Accès à Google Drive & Google Sheets

### Étapes

**1. Créer le fichier**

```
Google Drive → Nouveau → Google Sheets → Fichier vide
```

**2. Ouvrir l’éditeur de script**

```
Extensions → Apps Script
```

**3. Coller le code**

- Supprimer le contenu par défaut (`function myFunction() {}`)
- Copier-coller intégralement le contenu de `Code.gs`
- Enregistrer (`Ctrl+S` ou icône 💾)

**4. Lancer l’initialisation**

```
Retourner dans Google Sheets
Menu 🚗 SpiderVO Pro → ⚙️ Système → 🚀 Initialiser le système
```

**5. Autoriser les permissions**

Lors du premier lancement, Google demande les autorisations suivantes :

- Lire/modifier les feuilles de calcul
- Envoyer des emails (Gmail)
- Accéder à Google Drive (sauvegardes)

> ⚠️ Si l’avertissement “Application non vérifiée” apparaît → cliquer sur **Avancé → Accéder au projet**

**6. Configurer le garage**

```
Menu 🚗 SpiderVO Pro → Onglet PARAMÈTRES
```

Renseigner : nom, SIRET, adresse, téléphone, email.

-----

## 📊 Installation — Excel (offline)

Ouvrir `SpiderVO_Pro_v3.xlsx` directement dans **Microsoft Excel** (2016 ou supérieur recommandé) ou **LibreOffice Calc**.

> Les formules et validations fonctionnent nativement. Les automatisations Apps Script ne sont pas disponibles en mode Excel — utiliser la version Google Sheets pour les fonctionnalités avancées.

-----

## 🗂️ Architecture du système

### Onglets

|Onglet           |Description                          |Formules auto                         |
|-----------------|-------------------------------------|--------------------------------------|
|`DASHBOARD`      |KPIs en temps réel, navigation       |12 indicateurs live                   |
|`STOCK VO`       |Base véhicules complète (45 colonnes)|Coût revient, marges, ROI, jours stock|
|`CRM CLIENTS`    |Pipeline commercial, leads, relances |Alertes contact dépassé               |
|`VENTES`         |Bons de commande, factures, reprises |TVA marge, marge nette                |
|`ADMINISTRATIF`  |CERFA, carte grise, CT, garanties    |Alertes CT, dossiers incomplets       |
|`FINANCE`        |Dépenses, rentabilité mensuelle      |Récap automatique par catégorie       |
|`BON DE COMMANDE`|Template imprimable                  |Calculs financiers                    |
|`PARAMÈTRES`     |Configuration du garage              |—                                     |
|`📖 GUIDE`        |Mode d’emploi intégré                |—                                     |

-----

## 📦 Fonctionnalités détaillées

### Stock VO — 45 colonnes

**Identification**

- ID dossier auto-incrémenté (`VO-0001`, `VO-0002`…)
- Immatriculation, VIN, Marque, Modèle, Version, Finition
- Année, Kilométrage, Carburant, Boîte, Couleur, Puissance fiscale

**Caractéristiques**

- Date MEC, Nb clés, Nb propriétaires
- Historique entretien, Crit’Air, CT validité, Garantie

**Achat**

- Provenance, Fournisseur, Date achat
- Prix achat, Transport, Réparations, Nettoyage, Carte grise, Préparation, Divers

**Calculs automatiques**

```
Coût de Revient  = Prix achat + Transport + Réparations + Nettoyage + CG + Préparation + Divers
Prix Conseillé   = Coût de Revient × 1.20 (arrondi à 100€)
Prix Mini        = Coût de Revient × 1.05
Marge Brute €    = Prix Affiché − Coût de Revient
TVA sur Marge    = (Prix Affiché − Coût de Revient) / 1.20 × 0.20
Marge Nette €    = Marge Brute − TVA sur Marge
Marge Brute %    = Marge Brute / Coût de Revient
ROI %            = Marge Nette / Coût de Revient
Jours en Stock   = Aujourd'hui − Date Achat (ou Date Vente − Date Achat)
```

**Statuts** (avec code couleur automatique)

|Statut     |Couleur |
|-----------|--------|
|Disponible |🟢 Vert  |
|Préparation|🟠 Orange|
|Publié     |🔵 Bleu  |
|Réservé    |🟡 Jaune |
|Vendu      |🔴 Rouge |
|Livré      |🟣 Violet|
|SAV        |⚫ Gris  |

-----

### CRM Clients

**Pipeline commercial (8 étapes)**

```
Nouveau lead → Contacté → RDV planifié → Essai → Négociation → Réservé → Vendu → Perdu
```

**Score client**

- 🔥 Chaud — 🌡️ Tiède — ❄️ Froid — 💀 Perdu

**Alertes automatiques**

- Dernier contact > 7 jours → ligne en rouge
- Score 🔥 Chaud → fond orangé

-----

### Administratif (type AutoCERFA)

- Déclaration d’achat (CERFA 13750)
- Déclaration de cession (CERFA 15776)
- Suivi carte grise (reçue / transmise)
- CT : date de validité, résultat, alerte < 30 jours
- Statut dossier : ✅ COMPLET / ⚠️ INCOMPLET / ❌ BLOQUANT / 📋 EN COURS

-----

### Automatisations Apps Script (Google Sheets uniquement)

|Bouton / Fonction           |Description                                 |
|----------------------------|--------------------------------------------|
|`ouvrirFormulaireVehicule()`|Sidebar HTML pour ajouter un véhicule       |
|`ouvrirFormulaireClient()`  |Sidebar HTML pour ajouter un client/lead    |
|`genererBonCommande()`      |Génère un BC imprimable dans un onglet dédié|
|`rechercheAvancee()`        |Sidebar de recherche multi-critères         |
|`voirVehiculesDormants()`   |Liste les véhicules en stock > 45 jours     |
|`voirLeadsChauds()`         |Liste les leads score 🔥 actifs              |
|`voirDossiersIncomplets()`  |Dossiers admin incomplets ou bloquants      |
|`voirAlertesCT()`           |CT expirant dans les 30 prochains jours     |
|`envoyerEmailRelance()`     |Email automatique au client sélectionné     |
|`genererMessageWhatsApp()`  |Génère un lien WhatsApp pré-rempli          |
|`sauvegarderDrive()`        |Copie horodatée sur Google Drive            |
|`genererRapportMensuel()`   |Rapport synthèse du mois en cours           |
|`declencheurQuotidien()`    |Alerte email automatique chaque matin à 8h  |

-----

## 🎨 Design

- Palette **dark premium** (bleu nuit `#0A0F1E`, rouge accent `#E63946`)
- Alternance de couleurs sur les lignes de données
- En-têtes groupés par domaine (Identification / Achat / Calculs / Vente)
- Formatage conditionnel sur statuts, marges, CT, relances
- Dashboard type SaaS avec KPIs cards

-----

## 📐 Spécifications techniques

|Critère                  |Valeur          |
|-------------------------|----------------|
|Capacité stock           |5 000+ véhicules|
|Formules Excel           |9 522 (0 erreur)|
|Colonnes STOCK VO        |45              |
|Onglets                  |9               |
|Validations déroulantes  |15+ listes      |
|Règles cond. formatting  |25+             |
|Compatibilité Excel      |2016+           |
|Compatibilité LibreOffice|✅               |
|Apps Script              |ES2020 (V8)     |

-----

## ⚙️ Configuration avancée

### Modifier le seuil “véhicule dormant”

Dans `Code.gs`, modifier la constante :

```javascript
SEUIL_DORMANT: 45,  // Nombre de jours
```

### Modifier la marge de sécurité sur le prix conseillé

```javascript
// Ligne dans ajouterVehiculeDepuisFormulaire()
sheet.getRange(r, 33).setFormula(`=IF(AF${r}="","",ROUND(AF${r}*1.20/100,0)*100)`);
//                                                               ↑ 1.20 = +20%
```

### Activer les emails automatiques

Dans l’onglet `PARAMÈTRES`, renseigner la cellule **Email notifications**.  
Le script envoie chaque matin à 8h la liste des véhicules dormants.

### Changer la numérotation des dossiers

```javascript
DOSSIER_PREFIX: 'VO-',     // ex: 'SBR-' pour SBR AUTO
CLIENT_PREFIX:  'CL-',
COMMANDE_PREFIX:'BC-',
FACTURE_PREFIX: 'FA-',
```

-----

## 🔒 Sécurité recommandée

1. **Protéger les colonnes de formules** dans STOCK VO  
   `Sélectionner AF:AN → Format → Protéger la plage`
1. **Partager en lecture seule** avec les collaborateurs non-gestionnaires
1. **Sauvegarder hebdomadairement** via  
   `Menu 🚗 SpiderVO Pro → Système → Sauvegarder sur Drive`

-----

## 🗺️ Roadmap

- [ ] Intégration API LaCentrale / Autosphere (scraping prix marché)
- [ ] Export CERFA PDF automatique via Google Drive
- [ ] Module financement (calcul mensualités)
- [ ] Tableau de bord multi-site / multi-utilisateurs
- [ ] Application mobile (Google AppSheet)
- [ ] Intégration WhatsApp Business API
- [ ] Import CSV depuis fichiers BCA / SVI

-----

## 🤝 Contribution

Les contributions sont les bienvenues !

```bash
git clone https://github.com/[votre-compte]/spidervo-pro.git
cd spidervo-pro
# Modifier Code.gs ou le script Python de génération Excel
# Soumettre une Pull Request
```

-----

## 📄 Licence

MIT License — libre d’utilisation, modification et distribution.  
Voir <LICENSE> pour les détails.

-----

## 👤 Auteur

Développé pour **SBR AUTO** — Gennevilliers  
Inspiré de SpiderVO, AutoCERFA et PlanetVO

-----

> *“Un bon logiciel de gestion VO ne devrait pas coûter 150€/mois.”*