# GenieFactory - Business Plan 14 Mois (Nov 2025 - Dec 2026)

Génération automatisée du Business Plan et Business Model sur 14 mois à partir d'hypothèses centralisées.

## 🎯 Objectif

Restructurer le BP existant (38 mois) sur période focus **Nov 2025 → Dec 2026** pour :
- Seed round Sept 2026 (M11) : 500K€
- ARR Milestone Dec 2026 (M14) : 800K€
- Cohérence totale entre Excel et Word
- Traçabilité et reproductibilité complètes

## 📁 Structure Repo

```
geniefactory-bp-14m/
├── README.md                           ← Vous êtes ici
├── CLAUDE_CODE_PROMPT.md              ← Mission complète pour Claude Code
├── SPECIFICATIONS_FONCTIONNELLES.md   ← Specs détaillées
├── requirements.txt                   ← Dépendances Python
│
├── data/
│   ├── raw/                           ← Sources (read-only)
│   │   ├── BP_FABRIQ_PRODUCT-OCT2025.xlsx
│   │   ├── Business_Plan_GenieFactory-SEPT2025.docx
│   │   └── GENIE_FACTORY_PACTE_AATL-v3.docx
│   │
│   ├── structured/                    ← Données extraites
│   │   ├── assumptions.yaml           ← ⭐ SOURCE UNIQUE VÉRITÉ
│   │   ├── bp_extracted.json
│   │   ├── word_extracted.json
│   │   ├── projections.json
│   │   └── corrections_proposed.yaml  ← Corrections suggérées
│   │
│   ├── validation_rules.yaml          ← ⚠️ RÈGLES VALIDATION FINANCIÈRE
│   │
│   └── outputs/                       ← 📦 Livrables finaux
│       ├── BP_14M_Nov2025-Dec2026.xlsx
│       ├── BM_Updated_14M.docx
│       └── charts/                    ← Graphiques PNG
│           ├── arr_evolution.png
│           ├── ca_mensuel.png
│           ├── ebitda.png
│           ├── cash_position.png
│           ├── revenue_mix.png
│           └── team_evolution.png
│
├── scripts/
│   ├── 1_extract.py                   ← Extraction BP/BM/Pacte
│   ├── 2_generate_assumptions.py     ← Création assumptions.yaml
│   ├── 3_calculate_projections.py    ← Calculs financiers M1-M14
│   ├── 4_generate_bp_excel.py        ← Génération BP Excel + charts
│   ├── 5_update_bm_word.py           ← Update BM Word + visuals
│   ├── 6_validate.py                 ← Validation basique
│   ├── 7_validate_coherence.py       ← ⚠️ Validation cohérence avancée
│   ├── 8_fix_coherence.py            ← Correction automatique incohérences
│   └── generate_charts.py            ← Génération graphiques PNG
│
├── templates/
│   └── assumptions_template.yaml     ← Template avec commentaires
│
└── tests/
    └── test_calculations.py          ← Tests unitaires
```

## 🚀 Quickstart

### Installation

```bash
# Clone repo
git clone <repo_url>
cd geniefactory-bp-14m

# Install dependencies
pip install -r requirements.txt

# Vérifier structure
ls data/raw/  # Doit contenir les 3 fichiers sources
```

### Génération Complète

```bash
# 1. Extraction des données sources
python scripts/1_extract.py
# → Génère data/structured/bp_extracted.json + word_extracted.json

# 2. Création assumptions.yaml
python scripts/2_generate_assumptions.py
# → Génère data/structured/assumptions.yaml
# ⚠️ VALIDATION MANUELLE REQUISE : vérifier les hypothèses

# 3. Calcul projections
python scripts/3_calculate_projections.py
# → Génère data/structured/projections.json (ARR, CA, charges mensuels)

# 4. Génération BP Excel
python scripts/4_generate_bp_excel.py
# → Génère data/outputs/BP_14M_Nov2025-Dec2026.xlsx

# 5. Update BM Word
python scripts/5_update_bm_word.py
# → Génère data/outputs/BM_Updated_14M.docx

# 6. Validation basique
python scripts/6_validate.py
# → Checks ARR target, cohérence, formules Excel

# 7. Validation cohérence avancée ⚠️ CRITIQUE
python scripts/7_validate_coherence.py
# → Détecte incohérences valorisation, red flags investisseurs

# 8. Correction automatique (si nécessaire)
python scripts/8_fix_coherence.py
# → Corrige valorisations incohérentes, applique règles SaaS B2B

# 9. Re-validation
python scripts/7_validate_coherence.py
# → Vérifier Status: ✅ SUCCÈS
```

**OU** exécution d'un coup :
```bash
python run.py  # Enchaîne scripts 1-8 avec validation complète
```

## 📊 Métriques Clés

### Targets Financiers

| Métrique | M1 (Nov 25) | M11 (Sept 26) | M14 (Dec 26) |
|----------|-------------|---------------|--------------|
| **CA Total** | 36K€ | 120K€ | 140K€ |
| **ARR** | 0€ | 450K€ | 800K€ ✓ |
| **Équipe** | 5 ETP | 11 ETP | 12 ETP |
| **Cash** | 150K€ | 500K€ (seed) | 200K€ |
| **Burn Rate** | 35K€/mois | 45K€/mois | 40K€/mois |

### Hypothèses Principales

- **Hackathons** : 1.5-4/mois (progression)
- **Conversion Hack→Factory** : 30% (avec 2 mois délai)
- **Lancement Hub** : M8 (Juin 2026)
- **Churn Hub** : 10% annuel
- **Pre-seed** : 150K€ (M1)
- **Seed** : 500K€ (M11)

## 🔧 Ajuster les Hypothèses

### Modifier Volumes Hackathons

Éditer `data/structured/assumptions.yaml` :

```yaml
sales_assumptions:
  hackathon:
    volumes_monthly:
      m1: 2      # Au lieu de 1.5
      m2: 3      # Au lieu de 2
      # ... etc
```

Puis regénérer :
```bash
python scripts/3_calculate_projections.py
python scripts/4_generate_bp_excel.py
python scripts/6_validate.py
```

### Modifier Pricing

```yaml
pricing:
  hackathon:
    periods:
      - start_month: 1
        end_month: 6
        price_eur: 20000  # Au lieu de 18000
```

### Décaler Seed Round

```yaml
timeline:
  milestones:
    - month: 12  # Au lieu de 11
      name: "Seed Round"
      amount_eur: 500000
```

**→ Toujours relancer validation après modification !**

## ✅ Validation

### Validation Standard (6_validate.py)

Le script `6_validate.py` effectue les checks basiques suivants :

### Checks Financiers

- ✅ ARR M14 = 800K€ ± 10% (720K-880K€)
- ✅ ARR M11 ≥ 400K€ (attractivité seed)
- ✅ Burn rate max ≤ 60K€/mois
- ✅ Cash position jamais négative
- ✅ Équipe M14 ≤ 15 ETP

### Checks Cohérence

- ✅ Excel ↔ Word : ARR identique (<1K€ écart)
- ✅ Excel ↔ Word : CA total identique (<5% écart)
- ✅ Formules Excel fonctionnelles (pas hardcoded)

### Exemple Output

```
🔍 VALIDATION BP 14 MOIS - GenieFactory
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✅ FINANCIAL CHECKS
  ✓ ARR M14: 820,000€ (target 800,000€ ±10%)
  ✓ ARR M11: 460,000€ (>400,000€ minimum)
  ✓ Burn max: 48,000€/mois (<60,000€)
  ✓ Cash min: 85,000€ (>50,000€)
  ✓ Team M14: 12 ETP (<15)

⚠️ WARNINGS
  • M3-M4: CA flat 180K€ (vérifier saisonnalité)
  • Conversion Hack→Factory: 28% (target 30%)

✅ CONSISTENCY CHECKS
  ✓ Excel ↔ Word ARR: 820K€ ↔ 820K€ (Δ 0€)
  ✓ Excel ↔ Word CA: 1,050K€ ↔ 1,048K€ (Δ 0.2%)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STATUS: ✅ PASSED (2 warnings)
```

### Validation Cohérence Avancée (7_validate_coherence.py)

⚠️ **IMPORTANT** : Ce script détecte les incohérences **critiques** qui tuent la crédibilité investisseurs.

```bash
python scripts/7_validate_coherence.py
```

#### Règles de Validation Financière

Le script applique les règles du marché SaaS B2B définies dans `data/validation_rules.yaml` :

**1. Multiples de Valorisation**
```
Conservative (4-6x ARR)   : Croissance <30%/an
Realistic (7-10x ARR)     : Croissance 30-60%/an ✅ RECOMMANDÉ
Aggressive (11-15x ARR)   : Croissance >100%/an (justification requise)
ERREUR (>15x ARR)         : Incohérent sans hyper-croissance démontrée
```

**2. Red Flags Investisseurs**
- ❌ CAC > LTV/3
- ❌ Churn annuel > 15%
- ❌ Marge brute < 60% (SaaS)
- ❌ Break-even > 24 mois post-seed
- ❌ Valorisation/ARR > 12x sans justification
- ⚠️ NPS < 40
- ⚠️ Cycle de vente > 120 jours (PME/ETI)

**3. Cohérence Inter-Sections**

Le script vérifie automatiquement :

| Section Source | Section Cible | Formule |
|----------------|---------------|---------|
| 1.3 Vision | 7.2 Projections | Valorisation 2028 = ARR 2028 × 7-10 |
| 5.3 Recrutement | 7.2 Charges personnel | Charges = Effectifs × 65K€ |
| 4.1 Déploiement | 7.1 CA total | CA = Nb clients × ARPU |
| 7.2 Marketing | 4.1 Acquisitions | Charges marketing / acquisitions ≈ CAC |

#### Exemple Output

```
🔍 VALIDATION COHÉRENCE AVANCÉE
============================================================

💰 VALIDATION VALORISATION VS ARR
  ✗ Valorisation 200-300M€: Multiple 302.4x INCOHÉRENT
    ARR M14: 827K€
    Valorisation réaliste: 8M€ (10x)

✅ CORRECTIONS PROPOSÉES
╭────────────────────────────────────────────────╮
│ Section: 1.3 Vision                            │
│ Champ: valorisation_2028                       │
│                                                │
│ CONSERVATIVE: 5M€ (6x ARR)                     │
│   → Croissance <30%/an, marché mature          │
│                                                │
│ REALISTIC: 8M€ (10x ARR) ✅ RECOMMANDÉ         │
│   → Croissance 30-60%/an, marché stable        │
│                                                │
│ AGGRESSIVE: 12M€ (15x ARR)                     │
│   → Croissance >100%/an, hyper-croissance      │
╰────────────────────────────────────────────────╯

Statut: ❌ ÉCHEC - 3 erreurs critiques détectées
```

#### Correction Automatique

Si des incohérences sont détectées, utiliser le script de correction :

```bash
python scripts/8_fix_coherence.py
```

Ce script :
1. Lit les corrections proposées dans `data/structured/corrections_proposed.yaml`
2. Applique automatiquement les corrections recommandées
3. Sauvegarde le document Word corrigé
4. Génère un rapport de corrections

**⚠️ Workflow recommandé :**
```bash
# 1. Valider cohérence
python scripts/7_validate_coherence.py

# 2. Si erreurs, corriger automatiquement
python scripts/8_fix_coherence.py

# 3. Re-valider
python scripts/7_validate_coherence.py

# 4. Vérifier que Status: ✅ SUCCÈS
```

#### Règles de Valorisation - Exemples Concrets

**❌ INCORRECT** (Multiple 300x)
```
"Valorisation cible de 200-300M€ en 2028"
ARR 2028: 827K€
→ Multiple: 300x (INCOHÉRENT pour SaaS B2B)
```

**✅ CORRECT** (Multiple 10x)
```
"Valorisation cible de 8M€ en 2028"
ARR 2028: 827K€
→ Multiple: 10x (RÉALISTE pour croissance 30-60%/an)
```

**⚠️ AGRESSIF** (Multiple 15x)
```
"Valorisation cible de 12M€ en 2028"
ARR 2028: 827K€
→ Multiple: 15x (OK si croissance >100%/an démontrée)
```

#### Fichiers Générés

- `logs/coherence_report_YYYYMMDD_HHMMSS.txt` : Rapport détaillé
- `data/structured/corrections_proposed.yaml` : Corrections proposées

## 🧪 Tests

```bash
# Tests unitaires
pytest tests/

# Test calcul ARR
pytest tests/test_calculations.py::test_arr_calculation

# Test conversion hackathons
pytest tests/test_calculations.py::test_factory_conversion

# Coverage
pytest --cov=scripts tests/
```

## 📖 Documentation

### Pour Claude Code

Lire **CLAUDE_CODE_PROMPT.md** : prompt complet avec contexte, objectifs, contraintes techniques.

### Spécifications Fonctionnelles

Lire **SPECIFICATIONS_FONCTIONNELLES.md** : détail des 7 fonctionnalités attendues (F1 à F7).

### Assumptions Template

Voir **templates/assumptions_template.yaml** : structure complète commentée avec exemples.

## 🎨 Génération Excel : Détails

Le BP Excel généré (`BP_14M_Nov2025-Dec2026.xlsx`) contient :

### Sheets

1. **Synthèse** : Dashboard avec KPIs et graphiques
2. **P&L** : Détail mensuel CA/charges/EBITDA (14 colonnes)
3. **Ventes** : Pipeline détaillé par offre
4. **Paramètres** : Pricing reference
5. **Financement** : Pre-seed + Seed
6. **Charges Personnel** : Évolution ETP + salaires
7. **Infrastructure** : Coûts tech scaling
8. **Monitoring** : Métriques SaaS (MRR, ARR, churn)

### Formules Excel Actives

```excel
# CA Total mensuel (F2)
=SUM(F3:F8)

# ARR (F16) - uniquement Hub
=F5*12  # MRR Hub × 12

# Cash position (F20)
=E20+F2-F9+F_funding

# Validation ARR M14 (S16)
=IF(S16<720000,"⚠️ Sous target",IF(S16>880000,"⚠️ Sur-optimiste","✓ OK"))
```

### Charts

1. **ARR Growth** : Courbe évolution M1→M14
2. **Burn Rate** : Colonnes mensuelles (rouge si >50K€)
3. **Revenue Mix** : Stacked area (Hackathon/Factory/Hub/Services)

## 📝 Livrables Finaux

Après exécution complète :

✅ **assumptions.yaml** : 450 lignes commentées
✅ **projections.json** : Calculs mensuels M1-M14
✅ **BP_14M_Nov2025-Dec2026.xlsx** : 8 sheets, formules actives
✅ **BM_Updated_14M.docx** : Sections 7.2/7.3/7.4 actualisées
✅ **Validation report** : Tous checks passing

## 🚨 Troubleshooting

### Erreur : "ARR M14 hors target"

**Cause** : Volumes ou pricing trop conservateurs

**Solution** :
1. Vérifier `assumptions.yaml` → `sales_assumptions.hackathon.volumes_monthly`
2. Augmenter volumes M11-M14 (post-seed)
3. OU ajuster pricing Hub (starter/business/enterprise)
4. Regénérer

### Erreur : "Cash négatif M8"

**Cause** : Burn rate trop élevé ou pre-seed insuffisant

**Solution** :
1. Réduire charges personnel M1-M7
2. OU augmenter pre-seed M1 : 150K→200K€
3. OU lancer Hub plus tôt (M7 au lieu M8)

### Erreur : "Formules Excel cassées"

**Cause** : Génération Excel incorrecte

**Solution** :
```bash
# Regénérer avec verbose mode
python scripts/4_generate_bp_excel.py --verbose

# Vérifier logs
cat logs/generate_excel_YYYYMMDD.log
```

### Excel : Colonnes décalées

**Cause** : Erreur mapping colonnes

**Solution** :
Vérifier `scripts/4_generate_bp_excel.py` ligne ~150 :
```python
MONTH_COLS = ['F', 'G', 'H', ..., 'S']  # M1 à M14
```

## 🤝 Contribution

### Workflow Git

```bash
# Nouvelle feature
git checkout -b feature/adjust-hub-pricing

# Modifier assumptions
vim data/structured/assumptions.yaml

# Regénérer
python run.py

# Commit
git add data/structured/assumptions.yaml data/outputs/
git commit -m "Ajusté pricing Hub : starter 500→600€/mois"

# Push
git push origin feature/adjust-hub-pricing
```

### Versioning Assumptions

Chaque modification `assumptions.yaml` doit inclure :

```yaml
revision_history:
  - version: "1.1"
    date: "2025-01-XX"
    author: "Votre Nom"
    changes: "Augmenté volumes hackathons M4-M6 pour compenser lancement Hub retardé"
```

## 🔗 Ressources

- [Documentation openpyxl](https://openpyxl.readthedocs.io/)
- [Documentation python-docx](https://python-docx.readthedocs.io/)
- [YAML Spec](https://yaml.org/spec/1.2.2/)
- [GenieFactory - Pacte Actionnaires](data/raw/GENIE_FACTORY_PACTE_AATL-v3.docx)

## 📧 Support

Questions ? Ouvrir une issue GitHub ou contacter :
- Frédéric Ramet (CEO) : frederic@geniefactory.ai
- Repository maintainer : claude-code@anthropic.com

## 📜 License

Proprietary - GenieFactory SAS © 2025

---

**Version** : 1.0  
**Dernière mise à jour** : 2025-01-15  
**Auteur** : Claude Code (Anthropic)
