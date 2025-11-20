# ✅ Checklist Exécution - GenieFactory BP 14 Mois

**Objectif** : Restructurer BP 38 mois → 14 mois (Nov 2025 - Dec 2026)  
**Milestone critique** : ARR 800K€ à M14 (Dec 2026)  
**Durée estimée** : 6-8h

---

## 📋 Phase 0 : Setup Initial

### ✅ Environnement

- [ ] Python 3.9+ installé et accessible
- [ ] Créer environnement virtuel
  ```bash
  python -m venv venv
  source venv/bin/activate  # ou venv\Scripts\activate (Windows)
  ```
- [ ] Installer dépendances
  ```bash
  pip install -r requirements.txt
  ```
- [ ] Vérifier installation
  ```bash
  python -c "import openpyxl, docx, yaml; print('OK')"
  ```

### ✅ Structure Repo

- [ ] Créer arborescence complète
  ```bash
  mkdir -p data/{raw,structured,outputs}
  mkdir -p scripts tests templates logs
  ```
- [ ] Placer fichiers sources dans `data/raw/`
  - [ ] `BP_FABRIQ_PRODUCT-OCT2025.xlsx`
  - [ ] `Business_Plan_GenieFactory-SEPT2025.docx`
  - [ ] `GENIE_FACTORY_PACTE_AATL-v3.docx`
- [ ] Copier `assumptions_template.yaml` → `templates/`

### ✅ Validation Setup

- [ ] `ls data/raw/` affiche les 3 fichiers
- [ ] `python --version` affiche 3.9+
- [ ] `pip list | grep openpyxl` affiche version

**⏱️ Durée Phase 0 : 15 min**

---

## 📋 Phase 1 : Extraction Données

### Script : `scripts/1_extract.py`

**Objectif** : Parser BP Excel, BM Word, Pacte → JSON structuré

#### ✅ Sous-tâches

- [ ] **1.1 Parser BP Excel**
  - [ ] Charger workbook avec `openpyxl`
  - [ ] Identifier sheets critiques : Paramètres, P&L, Ventes
  - [ ] Extraire pricing (Paramètres rows 3-13)
  - [ ] Extraire CA mensuel (P&L rows 2-8, cols F-Q)
  - [ ] Extraire charges (P&L rows 10-13)
  - [ ] Parser formules Excel → JSON structure
  - [ ] Sauvegarder `data/structured/bp_extracted.json`

- [ ] **1.2 Parser BM Word**
  - [ ] Charger document avec `python-docx`
  - [ ] Extraire tableaux financiers (section 7.2, 7.3)
  - [ ] Extraire hypothèses croissance (patterns regex)
  - [ ] Sauvegarder `data/structured/bm_extracted.json`

- [ ] **1.3 Parser Pacte Word**
  - [ ] Extraire milestones ARR (regex `ARR\s*[≥>=]\s*(\d+)`)
  - [ ] Extraire : 800K€, 1.5M€
  - [ ] Sauvegarder `data/structured/pacte_extracted.json`

#### ✅ Validation Phase 1

- [ ] `bp_extracted.json` contient pricing 18K€ hackathon
- [ ] `bm_extracted.json` contient au moins 2 tableaux
- [ ] `pacte_extracted.json` contient ARR 800000 et 1500000
- [ ] Logs : "Extraction terminée : 3 fichiers parsés"

**⏱️ Durée Phase 1 : 1h**

---

## 📋 Phase 2 : Génération Assumptions

### Script : `scripts/2_generate_assumptions.py`

**Objectif** : Créer `assumptions.yaml` à partir données extraites + prompts interactifs

#### ✅ Sous-tâches

- [ ] **2.1 Charger données extraites**
  - [ ] Lire `bp_extracted.json`, `bm_extracted.json`, `pacte_extracted.json`
  - [ ] Merger dans structure unified

- [ ] **2.2 Générer sections YAML**
  - [ ] `meta` : sources, version, date
  - [ ] `timeline` : start_month, milestones (M1, M11, M14)
  - [ ] `pricing` : hackathon, factory, hub, services (avec évolution temporelle)
  - [ ] `sales_assumptions` : volumes mensuels par offre
  - [ ] `costs` : personnel, infra, marketing, admin
  - [ ] `financial_kpis` : targets ARR, revenue mix, marges
  - [ ] `validation_rules` : tolerances et seuils

- [ ] **2.3 Commentaires et sources**
  - [ ] Chaque valeur annotée avec `# Source: BP Oct 2025, sheet X`
  - [ ] Sections avec explications inline
  - [ ] Notes d'utilisation en footer

- [ ] **2.4 Prompts interactifs (si valeurs manquantes)**
  - [ ] Team evolution M1→M14 (si non extrait)
  - [ ] Volumes hackathons précis M1-M14
  - [ ] Hub launch month confirmation

- [ ] **2.5 Sauvegarder**
  - [ ] Écrire `data/structured/assumptions.yaml`
  - [ ] Validate YAML syntax
  - [ ] Print résumé : "Assumptions générées : 450 lignes, 12 sections"

#### ✅ Validation Phase 2

- [ ] `assumptions.yaml` existe et parsable (`yaml.safe_load()`)
- [ ] Contient section `pricing.hackathon.periods`
- [ ] Contient `timeline.milestones` avec 3 entries
- [ ] Tous les `m1` à `m14` définis dans `sales_assumptions.hackathon.volumes_monthly`
- [ ] **REVIEW MANUELLE** : Valider cohérence des hypothèses

**⏱️ Durée Phase 2 : 1h**

---

## 📋 Phase 3 : Calcul Projections

### Script : `scripts/3_calculate_projections.py`

**Objectif** : Calculer ARR, CA, charges, EBITDA pour chaque mois M1-M14

#### ✅ Sous-tâches

- [ ] **3.1 Charger assumptions**
  - [ ] Parse `assumptions.yaml`
  - [ ] Validate structure (jsonschema si dispo)

- [ ] **3.2 Calculer revenus mensuels**
  - [ ] Pour chaque mois M1-M14 :
    - [ ] **Hackathons** : `nb_hackathons(m) × price(m)`
    - [ ] **Factory** : Conversion hackathons M-2 × 30% × price
    - [ ] **Hub** : Si m≥8, calculer MRR (cumul clients - churn + upgrades)
    - [ ] **Services** : Basé sur hackathons + factory
    - [ ] **Total CA** : Sum des 4 lignes

- [ ] **3.3 Calculer coûts mensuels**
  - [ ] **Personnel** : `team_size(m) × avg_salary + freelance`
  - [ ] **Infra** : `base + (nb_clients × per_client_cost)`
  - [ ] **Marketing** : Base + events trimestriels
  - [ ] **Admin** : Fixe
  - [ ] **Total charges** : Sum

- [ ] **3.4 Calculer métriques**
  - [ ] **EBITDA** : `CA - Charges`
  - [ ] **Burn rate** : `max(0, -EBITDA)`
  - [ ] **MRR** : Hub uniquement
  - [ ] **ARR** : `MRR × 12`
  - [ ] **Cash** : Position cumulée avec fundings

- [ ] **3.5 Logs détaillés**
  - [ ] Pour chaque mois, logger :
    ```
    M5 : CA=360K€ (2.5 hackathons × 18K + 1 Factory × 75K + 10K services)
         Charges=48K€ (7 ETP × 6K + 5K freelance + 3K infra)
         EBITDA=312K€
         ARR=0€ (Hub pas encore lancé)
    ```

- [ ] **3.6 Sauvegarder**
  - [ ] Écrire `data/structured/projections.json`
  - [ ] Structure : array de 14 objets (un par mois)
  - [ ] Chaque objet : `{month, date, revenue, costs, metrics}`

#### ✅ Validation Phase 3

- [ ] `projections.json` contient 14 objets
- [ ] `projections[-1]['metrics']['arr']` (M14) entre 720K-880K€
- [ ] `projections[10]['metrics']['arr']` (M11) >= 400K€
- [ ] Aucun mois avec `cash < 0`
- [ ] Max burn rate < 60K€
- [ ] Logs affichent progression mois par mois

**⏱️ Durée Phase 3 : 2h**

---

## 📋 Phase 4 : Génération BP Excel

### Script : `scripts/4_generate_bp_excel.py`

**Objectif** : Créer `BP_14M_Nov2025-Dec2026.xlsx` avec formules Excel actives

#### ✅ Sous-tâches

- [ ] **4.1 Setup workbook**
  - [ ] Créer workbook vide avec `openpyxl`
  - [ ] Créer 8 sheets : Synthèse, P&L, Ventes, Paramètres, Financement, Charges Personnel, Infrastructure, Monitoring

- [ ] **4.2 Sheet P&L**
  - [ ] Row 1 : Headers (Période, M1, M2, ..., M14)
  - [ ] Row 2 : CA TOTAL avec formule `=SUM(F3:F8)` pour chaque mois
  - [ ] Rows 3-8 : CA par ligne (Hackathon, Factory, Hub, Services)
  - [ ] Row 9 : Charges TOTAL avec formule `=SUM(F10:F13)`
  - [ ] Rows 10-13 : Charges détail
  - [ ] Row 14 : EBITDA avec formule `=F2-F9`
  - [ ] Row 15 : Burn rate avec formule `=IF(F14<0,-F14,0)`
  - [ ] Row 16 : ARR avec formule `=F5*12` (Hub MRR × 12)
  - [ ] **Formules pour TOUTES les colonnes M1-M14**

- [ ] **4.3 Sheet Synthèse (Dashboard)**
  - [ ] KPIs clés : CA total 14M, ARR M14, EBITDA total, Burn max
  - [ ] Créer graphique ARR Growth (line chart)
  - [ ] Créer graphique Burn Rate (column chart)
  - [ ] Créer graphique Revenue Mix (stacked area)

- [ ] **4.4 Sheet Ventes**
  - [ ] Détail pipeline par offre
  - [ ] Nb hackathons par mois
  - [ ] Conversions Factory
  - [ ] Nouveaux clients Hub

- [ ] **4.5 Sheet Paramètres**
  - [ ] Grille pricing avec évolution M1-M6 vs M7-M14
  - [ ] Tableau recap des 4 offres

- [ ] **4.6 Sheet Financement**
  - [ ] Pre-seed M1 : 150K€ breakdown
  - [ ] Seed M11 : 500K€
  - [ ] Utilisation fonds

- [ ] **4.7 Sheet Charges Personnel**
  - [ ] Évolution ETP mensuelle
  - [ ] Salaires par rôle
  - [ ] Total charges personnel

- [ ] **4.8 Sheet Infrastructure**
  - [ ] Coûts base + scaling
  - [ ] Par client cost

- [ ] **4.9 Sheet Monitoring**
  - [ ] MRR mensuel
  - [ ] ARR tracking
  - [ ] Churn (si applicable)
  - [ ] LTV/CAC

- [ ] **4.10 Formatting**
  - [ ] Headers : Bold, background bleu
  - [ ] Totaux : Bold, background gris clair
  - [ ] Currency : # ##0 € (espace séparateur)
  - [ ] EBITDA négatif : Red text
  - [ ] ARR : Green bold
  - [ ] Conditional formatting : Burn >50K€ en rouge

- [ ] **4.11 Sauvegarder**
  - [ ] Écrire `data/outputs/BP_14M_Nov2025-Dec2026.xlsx`

#### ✅ Validation Phase 4

- [ ] Excel s'ouvre sans erreur dans MS Excel / LibreOffice
- [ ] Formules actives (pas valeurs hardcodées) : vérifier cell F2 contient `=SUM(F3:F8)`
- [ ] Graphiques affichent correctement
- [ ] ARR M14 (cell S16) affiche ~800K€
- [ ] Tous les mois (F-S) ont des valeurs
- [ ] Format currency avec espaces (exemple : 360 000 €)

**⏱️ Durée Phase 4 : 2h**

---

## 📋 Phase 5 : Update BM Word

### Script : `scripts/5_update_bm_word.py`

**Objectif** : Mettre à jour sections financières dans `BM_Updated_14M.docx`

#### ✅ Sous-tâches

- [ ] **5.1 Charger BM source**
  - [ ] Ouvrir `data/raw/Business_Plan_GenieFactory-SEPT2025.docx`
  - [ ] Identifier sections 7.2, 7.3, 7.4 (scan headings)

- [ ] **5.2 Mettre à jour Section 7.2 (P&L)**
  - [ ] Trouver tableau existant (4 colonnes)
  - [ ] Remplacer par nouveau tableau (14 colonnes M1-M14 + Total)
  - [ ] Données depuis `projections.json`
  - [ ] Lignes : CA Total, Hackathon, Factory, Hub, Services, Charges, EBITDA, ARR

- [ ] **5.3 Mettre à jour Section 7.3 (Financement)**
  - [ ] Tableau Pre-seed + Seed
  - [ ] Pre-seed M1 : 150K€ (breakdown)
  - [ ] Seed M11 : 500K€
  - [ ] Utilisation fonds par catégorie

- [ ] **5.4 Mettre à jour Section 7.4 (KPIs)**
  - [ ] Remplacer texte avec patterns regex :
    - `"ARR: 320K€ (2025)"` → `"ARR: 0€ (M1) → 800K€ (M14)"`
    - `"Break-even: Q1 2026"` → `"Break-even: Non attendu (croissance prioritaire)"`
    - `"Seed: 350K€"` → `"Seed: 500K€ (Sept 2026)"`
  - [ ] Update métriques : CA total, burn rate, équipe

- [ ] **5.5 Ajouter note méthodologique**
  - [ ] En fin de section 7 :
    ```
    Note méthodologique : Ces projections sont basées sur le fichier 
    assumptions.yaml (version 1.0) et sont reproductibles via le repo 
    GitHub geniefactory-bp-14m. Les hypothèses peuvent être ajustées 
    et les documents regénérés automatiquement.
    ```

- [ ] **5.6 Sauvegarder**
  - [ ] Écrire `data/outputs/BM_Updated_14M.docx`

#### ✅ Validation Phase 5

- [ ] Word s'ouvre sans erreur
- [ ] Tableau 7.2 a bien 14 colonnes mensuelles
- [ ] ARR M14 dans texte = 800K€
- [ ] Sections 7.2/7.3/7.4 cohérentes avec Excel
- [ ] Note méthodologique présente

**⏱️ Durée Phase 5 : 1h**

---

## 📋 Phase 6 : Validation Finale

### Script : `scripts/6_validate.py`

**Objectif** : Vérifier cohérence et targets

#### ✅ Sous-tâches

- [ ] **6.1 Checks financiers**
  - [ ] ARR M14 entre 720K-880K€
  - [ ] ARR M11 >= 400K€
  - [ ] Burn rate max < 60K€
  - [ ] Cash jamais négatif
  - [ ] Équipe M14 <= 15 ETP

- [ ] **6.2 Checks cohérence**
  - [ ] Extraire ARR M14 du Word
  - [ ] Extraire ARR M14 du Excel (cell S16)
  - [ ] Comparer : écart < 1K€
  - [ ] Extraire CA total Word vs Excel : écart < 5%

- [ ] **6.3 Checks formules Excel**
  - [ ] Ouvrir Excel avec openpyxl
  - [ ] Vérifier cell F2 contient formula (pas value)
  - [ ] Vérifier 10+ cellules formules actives

- [ ] **6.4 Génération rapport**
  - [ ] Créer rapport validation :
    ```
    ✅ FINANCIAL CHECKS
      ✓ ARR M14: 820,000€ (target 800,000€ ±10%)
      ✓ ARR M11: 460,000€ (>400,000€)
      ✓ Burn max: 48,000€ (<60,000€)
      ✓ Cash min: 85,000€ (>0€)
      ✓ Team M14: 12 ETP (<15)
    
    ⚠️ WARNINGS
      • Conversion Hack→Factory: 28% (target 30%)
    
    ✅ CONSISTENCY CHECKS
      ✓ Excel ↔ Word ARR: 0€ écart
      ✓ Excel ↔ Word CA: 0.2% écart
    
    STATUS: ✅ PASSED (1 warning)
    ```

- [ ] **6.5 Afficher rapport**
  - [ ] Print dans terminal avec `rich`
  - [ ] Sauvegarder `logs/validation_report_YYYYMMDD.txt`

#### ✅ Validation Phase 6

- [ ] Rapport affiche "STATUS: ✅ PASSED"
- [ ] Tous checks critiques passent
- [ ] Warnings (si présents) sont documentés
- [ ] Rapport sauvegardé

**⏱️ Durée Phase 6 : 1h**

---

## 📋 Phase 7 : Documentation & Tests

### ✅ Sous-tâches

- [ ] **7.1 Tests unitaires**
  - [ ] `tests/test_calculations.py` :
    - [ ] `test_arr_calculation()`
    - [ ] `test_factory_conversion()`
    - [ ] `test_hub_ramp()`
  - [ ] Lancer : `pytest tests/`
  - [ ] Coverage > 80%

- [ ] **7.2 Documentation**
  - [ ] README.md complet et à jour
  - [ ] CHANGELOG.md : Différences vs BP Oct 2025
  - [ ] Commentaires code : Docstrings Python
  - [ ] Logs : Tous scripts génèrent logs détaillés

- [ ] **7.3 Git**
  - [ ] Init repo : `git init`
  - [ ] Add : `git add .`
  - [ ] Commit : `git commit -m "Initial BP 14M generation"`
  - [ ] Tag : `git tag v1.0`

**⏱️ Durée Phase 7 : 1h**

---

## ✅ Checklist Finale

### 📦 Livrables

- [ ] `data/structured/assumptions.yaml` (450 lignes)
- [ ] `data/structured/projections.json`
- [ ] `data/outputs/BP_14M_Nov2025-Dec2026.xlsx` (8 sheets, formules)
- [ ] `data/outputs/BM_Updated_14M.docx`
- [ ] `logs/validation_report_YYYYMMDD.txt`
- [ ] `README.md`
- [ ] `CHANGELOG.md`
- [ ] Tests passing

### 🎯 Métriques Clés Validées

- [ ] ARR M14 = 800K€ ± 10%
- [ ] ARR M11 >= 400K€
- [ ] Cash position positive tout le temps
- [ ] Burn rate max < 60K€
- [ ] Équipe M14 = 12 ETP
- [ ] Cohérence Excel ↔ Word < 5%

### 📝 Prochaines Étapes

- [ ] Review manuelle assumptions.yaml
- [ ] Validation business avec équipe
- [ ] Ajustements si nécessaire
- [ ] Présentation aux investisseurs

---

## 📊 Résumé Effort

| Phase | Durée | Description |
|-------|-------|-------------|
| 0. Setup | 15min | Env + structure |
| 1. Extraction | 1h | Parse BP/BM/Pacte |
| 2. Assumptions | 1h | Génération YAML |
| 3. Projections | 2h | Calculs M1-M14 |
| 4. BP Excel | 2h | Génération Excel |
| 5. BM Word | 1h | Update Word |
| 6. Validation | 1h | Checks cohérence |
| 7. Doc & Tests | 1h | Finitions |
| **TOTAL** | **8h15** | **Complet** |

---

**Date de création** : 2025-01-15  
**Version** : 1.0  
**Auteur** : Claude Assistant (pour Claude Code)
