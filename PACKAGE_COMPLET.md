# 📦 Package Complet - GenieFactory BP 14 Mois

## Résumé

J'ai créé un **framework complet** pour que Claude Code restructure votre Business Plan de 38 mois → 14 mois (Nov 2025 - Dec 2026), avec traçabilité totale et reproductibilité.

---

## 🎯 Ce qui a été généré

### 1. **CLAUDE_CODE_PROMPT.md** (4 pages)
**Rôle** : Mission complète pour Claude Code

**Contenu** :
- Contexte détaillé (GenieFactory, objectifs seed, ARR 800K€)
- Structure repo attendue
- Workflow d'exécution (6 scripts séquentiels)
- Fonctionnalités clés détaillées (F1-F6)
- Hypothèses de travail (croissance, équipe, cash flow)
- Critères de succès
- Validation requise

**Usage** : Donner ce prompt à Claude Code en début de projet

---

### 2. **SPECIFICATIONS_FONCTIONNELLES.md** (15 pages)
**Rôle** : Spécifications techniques détaillées

**Contenu** :
- **F1** : Extraction données (Excel, Word parsing)
- **F2** : Génération assumptions.yaml
- **F3** : Calculs projections (ARR, CA, charges)
- **F4** : Génération BP Excel (formules actives)
- **F5** : Update BM Word
- **F6** : Validation automatique
- **F7** : Documentation

**Détails** :
- Algorithmes pseudo-code
- Exemples de formules Excel
- Validation rules
- Tests unitaires

**Usage** : Référence technique pour Claude Code

---

### 3. **assumptions_template.yaml** (450 lignes)
**Rôle** : Source unique de vérité, commentée et sourcée

**Sections** :
```yaml
meta                     # Version, sources
timeline                 # Milestones M1, M11, M14
pricing                  # Hackathon, Factory, Hub, Services
sales_assumptions        # Volumes mensuels par offre
costs                    # Personnel, infra, marketing
financial_kpis           # Targets ARR, marges
validation_rules         # Tolerances
scenarios                # Base/upside/downside
critical_assumptions     # Risques identifiés
```

**Highlights** :
- Chaque valeur sourcée (`# Source: BP Oct 2025, sheet X`)
- Explications inline
- Progression temporelle détaillée (M1-M14)
- Métriques SaaS (MRR, ARR, churn, LTV/CAC)

**Usage** : Template à valider/ajuster par Claude Code

---

### 4. **README.md** (10 pages)
**Rôle** : Guide complet d'utilisation du repo

**Contenu** :
- Quickstart (installation, génération)
- Structure repo expliquée
- Métriques clés et hypothèses principales
- Guide ajustement assumptions (volumes, pricing, timing)
- Validation détaillée (checks financiers, cohérence)
- Tests unitaires
- Troubleshooting (erreurs communes et solutions)
- Workflow Git et versioning

**Usage** : Documentation utilisateur finale

---

### 5. **requirements.txt**
**Rôle** : Dépendances Python

**Packages** :
```
openpyxl>=3.1.2         # Excel manipulation
python-docx>=1.1.0      # Word documents
pyyaml>=6.0.1           # YAML parsing
pandas>=2.0.3           # Data ops
pytest>=7.4.3           # Testing
rich>=13.7.0            # CLI output
```

**Usage** : `pip install -r requirements.txt`

---

### 6. **run.py**
**Rôle** : Orchestrateur principal (script Python)

**Fonctionnalités** :
- Check dépendances automatique
- Check fichiers sources
- Exécution séquentielle scripts 1-6
- Progress bars avec `rich`
- Gestion erreurs
- Rapport final avec timing

**Usage** :
```bash
python run.py                  # Exécution complète
python run.py --skip-extract   # Skip si déjà extrait
python run.py --validate-only  # Seulement validation
```

---

### 7. **CHECKLIST_EXECUTION.md** (8 pages)
**Rôle** : Checklist pas-à-pas pour Claude Code

**Structure** :
- **Phase 0** : Setup (15min)
- **Phase 1** : Extraction (1h)
- **Phase 2** : Assumptions (1h)
- **Phase 3** : Projections (2h)
- **Phase 4** : BP Excel (2h)
- **Phase 5** : BM Word (1h)
- **Phase 6** : Validation (1h)
- **Phase 7** : Doc & Tests (1h)

**Pour chaque phase** :
- Objectif clair
- Sous-tâches détaillées avec checkboxes
- Validation attendue
- Durée estimée

**Résumé** : 8h15 effort total

**Usage** : Suivi d'exécution pour Claude Code

---

## 🚀 Workflow Complet

### Pour toi (maintenant)

1. **Télécharger tous les fichiers** depuis `/mnt/user-data/outputs/`
   - CLAUDE_CODE_PROMPT.md
   - SPECIFICATIONS_FONCTIONNELLES.md
   - assumptions_template.yaml
   - README.md
   - requirements.txt
   - run.py
   - CHECKLIST_EXECUTION.md

2. **Créer repo GitHub**
   ```bash
   mkdir geniefactory-bp-14m
   cd geniefactory-bp-14m
   # Copier tous les fichiers
   ```

3. **Uploader les 3 docs sources**
   - `BP_FABRIQ_PRODUCT-OCT2025.xlsx`
   - `Business_Plan_GenieFactory-SEPT2025.docx`
   - `GENIE_FACTORY_PACTE_AATL-v3.docx`
   
   Dans `data/raw/`

### Pour Claude Code (ensuite)

4. **Démarrer Claude Code** avec le prompt :
   ```
   Voici le repo geniefactory-bp-14m.
   
   Mission : Restructurer le BP sur 14 mois (Nov 2025 - Dec 2026).
   
   Lire CLAUDE_CODE_PROMPT.md pour contexte complet.
   
   Suivre CHECKLIST_EXECUTION.md phase par phase.
   
   Utiliser assumptions_template.yaml comme base.
   
   Générer scripts 1-6 en suivant SPECIFICATIONS_FONCTIONNELLES.md.
   
   Objectif final : BP Excel + BM Word cohérents avec ARR 800K€ à M14.
   ```

5. **Claude Code exécute** :
   - Création des 6 scripts Python
   - Extraction données
   - Génération assumptions.yaml
   - Calculs projections
   - Génération BP Excel + BM Word
   - Validation complète

6. **Validation humaine** :
   - Review `data/structured/assumptions.yaml`
   - Vérifier `data/outputs/BP_14M_Nov2025-Dec2026.xlsx`
   - Valider `data/outputs/BM_Updated_14M.docx`
   - Ajuster si nécessaire et regénérer

---

## 🎯 Métriques Cibles Validées

| Métrique | Target | Validation |
|----------|--------|------------|
| **ARR M14** | 800K€ | ±10% (720K-880K€) |
| **ARR M11** | 450K€ | Minimum 400K€ |
| **Équipe M14** | 12 ETP | Maximum 15 ETP |
| **Burn max** | ~45K€/mois | Maximum 60K€ |
| **Cash** | Toujours >0 | Minimum 50K€ |
| **Seed** | 500K€ (M11) | Timing fixe |

---

## 📊 Hypothèses Principales

### Pricing
- **Hackathon** : 18K€ (M1-M6) → 20K€ (M7-M14)
- **Factory** : 75K€ → 82.5K€
- **Hub Starter** : 500€/mois
- **Hub Business** : 2000€/mois
- **Hub Enterprise** : 10000€/mois

### Volumes
- **Hackathons** : 1.5-4/mois (progression)
- **Factory** : 30% conversion avec 2 mois délai
- **Hub** : Lancement M8, 2-6 nouveaux clients/mois

### Équipe
- **M1** : 5 ETP (fondateurs + 1)
- **M3** : +2 dev
- **M7** : +2 commercial
- **M11** : +2 customer success (post-seed)
- **M14** : 12 ETP

---

## 🔧 Ajustements Possibles

Si les résultats ne conviennent pas, modifier `assumptions.yaml` :

### Augmenter ARR
```yaml
sales_assumptions:
  hackathon:
    volumes_monthly:
      m11: 5  # Au lieu de 4
      m12: 5
      m13: 5
      m14: 6
```

### Accélérer Hub
```yaml
enterprise_hub:
  launch_month: 7  # Au lieu de 8
  new_customers_monthly:
    m8: 3  # Au lieu de 2
```

### Réduire Burn
```yaml
costs:
  personnel:
    team_evolution:
      m7: 8  # Au lieu de 9
```

Puis regénérer :
```bash
python run.py
```

---

## ✅ Ce qui est Livré

### Documents Cadrage
- ✅ Prompt Claude Code (mission complète)
- ✅ Spécifications fonctionnelles (F1-F7)
- ✅ Checklist d'exécution (phases 0-7)

### Templates & Config
- ✅ assumptions.yaml template (450 lignes commentées)
- ✅ requirements.txt (dépendances)
- ✅ run.py (orchestrateur)

### Documentation
- ✅ README complet (setup, usage, troubleshooting)
- ✅ Structure repo définie

### Ce qui reste à faire (par Claude Code)
- ⏳ Création des 6 scripts Python (`1_extract.py` → `6_validate.py`)
- ⏳ Exécution extraction + génération
- ⏳ Génération BP Excel final
- ⏳ Update BM Word final
- ⏳ Tests unitaires

**Durée estimée Claude Code** : 6-8h

---

## 🎨 Différences vs BP Actuel

| Aspect | BP Oct 2025 (38 mois) | BP 14 Mois |
|--------|---------------------|------------|
| **Période** | Nov 2025 → 2028 | Nov 2025 → Dec 2026 |
| **Granularité** | Mensuelle puis annuelle | Mensuelle (M1-M14) |
| **Focus** | Croissance long terme | Seed + ARR 800K€ |
| **Hypothèses** | Implicites dans Excel | Explicites dans YAML |
| **Reproductibilité** | Manuelle | Automatisée (Python) |
| **Validation** | Ad-hoc | Checks automatiques |
| **Cohérence** | Risque dérive Excel/Word | Garantie (source unique) |

---

## 💡 Points d'Attention

### Critique (Must Fix)
⚠️ **ARR M14 avec Hub seul** : Le template actuel donne ~600K€ ARR Hub à M14, il manque 200K€ pour atteindre 800K€ total.

**Solutions** :
1. **Augmenter volumes Hub** : 6 → 10 nouveaux clients/mois M11-M14
2. **Accélérer upgrades** : 20% → 30% Starter→Business
3. **Lancer Hub M7** : 1 mois plus tôt
4. **Mix offres** : Considérer Factory comme ARR partiel ?

→ **Claude Code devra ajuster après premiers calculs**

### Important (Should Review)
- Conversion Hackathon→Factory : 30% ambitieux ? (industrie ~20-25%)
- Churn Hub : 10% optimiste pour early stage ?
- Équipe 12 ETP à M14 : suffisant pour 30+ clients Hub ?

### Nice to Have
- Scenarios (upside/downside) : calculés automatiquement
- Sensitivity analysis : impact ±10% volumes
- Export PDF du BP

---

## 🤝 Prochaines Étapes

1. **Télécharger les 7 fichiers** depuis `/mnt/user-data/outputs/`
2. **Créer repo** `geniefactory-bp-14m`
3. **Uploader les 3 docs sources** dans `data/raw/`
4. **Lancer Claude Code** avec CLAUDE_CODE_PROMPT.md
5. **Suivre CHECKLIST_EXECUTION.md** phase par phase
6. **Valider outputs** : BP Excel + BM Word
7. **Ajuster** si nécessaire (modifier assumptions.yaml)
8. **Présenter** aux investisseurs !

---

## 📞 Questions / Clarifications

Si besoin d'ajustements avant de lancer Claude Code :

1. **Timeline exacte** : Nov 2025 → Dec 2026 confirmé ?
2. **Seed timing** : Sept 2026 (M11) ou flexible ?
3. **ARR target strict** : 800K€ ou guideline ?
4. **Équipe** : 5 → 12 ETP réaliste ?
5. **Pre-seed** : 150K€ confirmé ?

**Tout est paramétrable dans assumptions.yaml !**

---

**Version** : 1.0  
**Date** : 2025-01-15  
**Auteur** : Claude Assistant (Anthropic)  
**Pour** : GenieFactory - Restructuration BP 14 Mois

**Durée génération framework** : ~2h  
**Durée exécution Claude Code estimée** : 6-8h  
**Durée totale projet** : ~10h
