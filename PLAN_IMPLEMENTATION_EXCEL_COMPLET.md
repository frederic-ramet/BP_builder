# Plan d'Implémentation - Excel Complet (50 mois, 15 sheets)

## 🎯 Objectif
Reproduire l'Excel source **exactement** avec les 15 sheets et la couverture temporelle complète (Nov 2025 - Dec 2029, ~50 mois).

---

## 📊 Structure Temporelle Cible

### Périodes Couvertes
```
2025-2026: Mois 11-12 (Nov-Dec) + Mois 1-12 (2026)  = 14 mois
2027:      Mois 1-12                                 = 12 mois
2028:      Mois 1-12                                 = 12 mois
2029:      Mois 1-12                                 = 12 mois
────────────────────────────────────────────────────
TOTAL:                                                 50 mois
```

### Colonnes P&L
- Colonne C: Total 2025-2026
- Colonnes D-Q: Nov 2025 - Dec 2026 (14 colonnes)
- Colonne R: Total 2027
- Colonnes S-AD: Jan-Dec 2027 (12 colonnes)
- Colonne AE: Total 2028
- Colonnes AF-AQ: Jan-Dec 2028 (12 colonnes)
- Colonne AR: Total 2029
- Colonnes AS-BD: Jan-Dec 2029 (12 colonnes)

**Total: ~122 colonnes** (incluant totaux et colonnes de calcul)

---

## 📋 Les 15 Sheets à Implémenter

### 1. ✅ Synthèse (Dashboard)
**Status**: Existe mais simplifié
**Besoins**:
- Ajouter vue annuelle (colonnes par année)
- Ajouter graphiques additionnels
- Formules référençant autres sheets
- KPIs par année (2025-2029)

**Structure**:
```
- CA TOTAL par année (lignes 3-8)
  - Hackathon
  - Services implémentation
  - Enterprise Hub
  - Factory
  - Marketplace
- DÉPENSES OPÉRATIONNELLES (lignes 10-14)
  - Sous traitance
  - Infrastructure
  - Charges personnel
- RÉSULTAT OPÉRATIONNEL
- Graphiques intégrés
```

### 2. ❌ Stratégie de vente (NOUVEAU)
**Status**: Manquant
**Description**: Détail pipeline par segment de marché

**Structure**:
```
Segments:
- Small Market (colonnes E-I)
  - Audit BPI
  - Implémentation
  - Licence Établi
  - Forge (Proto)
- Mid Market (colonnes J-L)
- Top Market (colonnes M-O)

Données par segment:
- Prix ÉTABLI
- Prix FORGE
- Crédit services ARR
- ARR moyen
- CA non-ARR
- CA total
- % Close
- Phases (Validation, etc.)
```

### 3. ✅ Financement
**Status**: Existe mais simplifié
**Besoins**: Ajouter détails levées

**Structure**:
```
Lignes par levée:
- Pre-seed Q4 2025 (150K€)
  - Autoposia (50K€)
  - F-Initiatives (40K€)
  - CIC (30K€)
  - BPI (30K€)
- Seed Q3 2026 (500K€)
- Series A Q4 2027
- Series B (optionnel)

Colonnes par année: 2025-2030
```

### 4. ✅ P&L (Compte de Résultat)
**Status**: Existe mais 14 mois seulement
**Besoins**: Étendre à 50 mois (122 colonnes)

**Structure**:
```
Lignes principales:
- CA TOTAL
  - GenieFactory Hackathon
  - Services d'implémentation
  - Enterprise Hub
  - GenieFactory Factory
  - Marketplace

- DÉPENSES OPÉRATIONNELLES
  - Sous traitance
  - Infrastructure technique
  - Charges de personnel
  - Licence ForgeAI

- RÉSULTAT OPÉRATIONNEL
- Cash Flow
- Position de trésorerie

Colonnes: Total année + détail mensuel pour 2025-2029
```

### 5. ✅ Paramètres (Pricing)
**Status**: Existe
**Besoins**: Vérifier cohérence avec assumptions.yaml

**Structure**:
```
Prix par année (2025-2029):
- Hackathon: 18K→20K→22K→25K→25K
- Services implémentation: 10K→25K→30K→35K→40K
- Formation: 5K→6K→7K→8K→9K
- Hub Starter: 500→550→605→665→732
- Hub Business: 2K→2.2K→2.4K→2.6K→2.9K
- Hub Enterprise: 10K→11K→12K→13K→14K
- Factory: 75K→90K→110K→125K→135K
- Academy: 60K→65K→70K→75K→80K
- Vertical Solutions: 30K→35K→40K
- AI Governance: 40K→45K→50K
```

### 6. ❌ GTMarket (Go-to-Market détaillé)
**Status**: Manquant
**Description**: Phases commerciales détaillées

**Structure**:
```
Timeline mensuelle 2025-2029
Phases:
- Validation Product-Market Fit (2025-Q4 à 2026-Q2)
- Accélération commerciale (2026-Q3+)
- Scale international (2027+)

Budget par phase
Équipes dédiées
Canaux d'acquisition
```

### 7. ✅ Ventes (Pipeline détaillé)
**Status**: Existe mais simplifié
**Besoins**: Étendre à 50 mois

**Structure détaillée**:
```
Par offre (lignes):
1. GenieFactory Hackathon
   - nb hackathons (ligne 2)
   - prix moyen (ligne 3)
   - CA Hackathons (ligne 4)

2. Services d'implémentation
   - nb implémentations (ligne 5)
   - prix moyen implémentation (ligne 6)
   - nb formations (ligne 7)
   - prix moyen formation (ligne 8)
   - CA Services (ligne 9)

3. Enterprise Hub
   - nb nouveaux Starter (ligne 10)
   - nb nouveaux Business (ligne 11)
   - nb nouveaux Enterprise (ligne 12)
   - abonnés Starter cumulés (ligne 13)
   - abonnés Business cumulés (ligne 14)
   - abonnés Enterprise cumulés (ligne 15)
   - Churn (ligne 16-18)
   - MRR (ligne 19-21)
   - ARR (ligne 22)

4. GenieFactory Factory
5. Marketplace

Colonnes: Total année + détail mensuel
```

### 8. ❌ Sous traitance (NOUVEAU)
**Status**: Manquant
**Description**: Coûts externes détaillés

**Structure**:
```
Catégories:
- Formation
  - PM (750€)
  - Besoin free
  - # interventions

- Audits
  - TJM (750€)
  - # interventions
  - charge % consulting
  - capa interne

- Accompagnement

Colonnes: Total + mensuel 2025-2029
```

### 9. ❌ Charges de personnel et FG (NOUVEAU)
**Status**: Manquant (crucial)
**Description**: Détail complet équipe + charges

**Structure**:
```
Rôles (lignes):
- Directeur (mini/intermédiaire/cible)
- Consultant
- Responsable Commercial
- Product Owner
- Tech Senior
- Tech Junior
- BD (junior)
- Stagiaire

Pour chaque rôle:
- ETP par mois
- Salaire brut
- Charges sociales (45%)
- Total

Ligne de synthèse:
- Total ETP
- Total salaires
- Total charges
- Total coût

Colonnes: Mensuel 2025-2029 (50 colonnes)
```

### 10. ❌ DIRECTION (NOUVEAU)
**Status**: Manquant
**Description**: Scénarios salaires direction

**Structure**:
```
Scénarios:
- Directeur (mini): 1 personne
- Directeur (intermédiaire): 2 personnes
- Directeur (cible): 3 personnes

Détail:
- Salaires (incluant stagiaires)
- Charges
- Total

Grille salariale:
- Directeur mini: 35K€ brut
- Directeur intermédiaire: 50K€ brut
- Directeur cible: 80K€+ brut
```

### 11. ❌ Infrastructure technique (NOUVEAU)
**Status**: Manquant (important)
**Description**: Coûts tech scaling

**Structure**:
```
Catégories:
- Cloud / AWS
  - Coût par utilisateur
  - Coût par client
  - Scaling automatique

- SaaS Tools
  - Notion, Slack, etc.
  - Outils dev (GitHub, etc.)
  - Outils analytics

- Licences logicielles
- R&D externe
- Infrastructure sécurité

Colonnes: Mensuel 2025-2029
Formules: Coût = f(nb clients, nb users)
```

### 12. ❌ Fundings (NOUVEAU)
**Status**: Manquant
**Description**: Détail complet levées

**Structure**:
```
Timeline levées:
- Pre-seed Q4 2025: 150K€
  - Détail par investisseur
  - Utilisation des fonds

- Seed Q3 2026: 500K€
  - VC / Business Angels
  - Valorisation pre-money
  - Dilution

- Series A Q4 2027
- Series B

Suivi trésorerie:
- Cash in
- Cash out
- Position nette
```

### 13. ❌ >> (Navigation)
**Status**: Bouton de navigation
**Description**: Liens rapides entre sheets

### 14. ❌ Positionnement (NOUVEAU)
**Status**: Manquant
**Description**: Analyse concurrentielle

**Structure**:
```
Matrice positionnement:
- Axe X: Prix
- Axe Y: Fonctionnalités

Concurrents:
- Solution A
- Solution B
- GenieFactory (positionnement)

Forces / Faiblesses
USP (Unique Selling Points)
```

### 15. ❌ Marketing (NOUVEAU)
**Status**: Manquant (important)
**Description**: Budget marketing détaillé

**Structure**:
```
Canaux:
- Pub digitale (Google, LinkedIn)
  - Budget mensuel
  - CPC / CPM
  - Conversions attendues

- Events / Conférences
  - Participation
  - Sponsoring
  - Networking

- Content Marketing
  - Blog
  - Newsletter
  - Webinars

- PR / Communication
- Partenariats

Métriques:
- CAC par canal
- ROI par canal
- Budget total

Colonnes: Mensuel 2025-2029
```

---

## 🔧 Plan d'Implémentation Technique

### Phase 1: Extension Assumptions (2-3h)
**Fichier**: `data/structured/assumptions.yaml`

**Modifications**:
```yaml
# Ajouter section long_term
long_term_projections:
  # Projections 2027-2029
  years:
    2027:
      arr_growth_rate: 0.70  # +70%/an
      team_growth: 5  # +5 ETP
      pricing_increase: 0.10  # +10%
      new_offerings:
        - vertical_solutions
        - ai_governance

    2028:
      arr_growth_rate: 0.60
      team_growth: 5
      pricing_increase: 0.10
      marketplace_launch: true

    2029:
      arr_growth_rate: 0.50
      team_growth: 4
      pricing_increase: 0.05

# Détail personnel par rôle
personnel_details:
  roles:
    directeur_mini:
      salary_brut: 35000
      charges_rate: 0.45
      fte_timeline:  # Par mois
        m1_m14: 1.0
        2027: 2.0
        2028: 3.0

    product_owner:
      salary_brut: 50000
      charges_rate: 0.45
      fte_timeline:
        m1_m6: 0.75
        m7_m14: 1.0
        2027: 2.0

    tech_senior:
      salary_brut: 65000
      charges_rate: 0.45
      fte_timeline:
        m1_m14: 0.0
        2027: 1.0
        2028: 2.0

    # ... autres rôles

# Infrastructure tech
infrastructure:
  cloud:
    cost_per_user: 5  # €/mois
    cost_per_client: 50  # €/mois
    base_cost: 1000  # €/mois

  saas_tools:
    - name: "Notion"
      cost_per_user: 10
    - name: "Slack"
      cost_per_user: 7
    # ...

# Marketing
marketing:
  channels:
    digital_ads:
      monthly_budget:
        2025: 2000
        2026: 5000
        2027: 10000
        2028: 15000
        2029: 20000
      expected_cac: 1500

    events:
      monthly_budget:
        2025: 1000
        2026: 3000
        2027: 5000

    content:
      monthly_budget:
        2025: 1000
        2026: 2000
```

### Phase 2: Extension Projections (3-4h)
**Fichier**: `scripts/3_calculate_projections.py`

**Modifications**:
```python
class ProjectionsCalculator:
    def calculate_projections(self, months_count=50):
        """Calculer projections M1-M50"""

        for month in range(1, months_count + 1):
            # Déterminer l'année
            if month <= 14:
                year = "2025-2026"
            elif month <= 26:
                year = "2027"
            elif month <= 38:
                year = "2028"
            else:
                year = "2029"

            # Appliquer paramètres de l'année
            year_params = self.assumptions['long_term_projections']['years'].get(year, {})

            # Calculer revenus avec croissance ajustée
            revenue = self.calculate_revenue_for_month(month, year_params)

            # Calculer charges avec team scaling
            costs = self.calculate_costs_for_month(month, year_params)

            # Personnel détaillé par rôle
            personnel_detail = self.calculate_personnel_detail(month, year_params)

            # Infrastructure (coûts fonction de clients)
            infrastructure = self.calculate_infrastructure(month, nb_clients)

            # Marketing (budget par canal)
            marketing = self.calculate_marketing(month, year_params)

            projections.append({
                'month': month,
                'year': year,
                'revenue': revenue,
                'costs': costs,
                'personnel_detail': personnel_detail,
                'infrastructure': infrastructure,
                'marketing': marketing,
                'metrics': self.calculate_metrics(...)
            })
```

### Phase 3: Refonte Complète Excel Generator (6-8h)
**Fichier**: `scripts/4_generate_bp_excel_full.py` (nouveau fichier)

**Structure**:
```python
class FullBPExcelGenerator:
    def __init__(self, projections_50m, assumptions):
        self.projections = projections_50m  # 50 mois
        self.assumptions = assumptions
        self.wb = Workbook()

        # Définir colonnes pour 50 mois
        # Col C: Total 2025-2026
        # Cols D-Q: Nov 2025 - Dec 2026 (14)
        # Col R: Total 2027
        # Cols S-AD: 2027 détail (12)
        # Col AE: Total 2028
        # Cols AF-AQ: 2028 détail (12)
        # Col AR: Total 2029
        # Cols AS-BD: 2029 détail (12)

    def create_sheet_synthese(self):
        """Sheet 1: Synthèse"""
        # Vue annuelle avec KPIs

    def create_sheet_strategie_vente(self):
        """Sheet 2: Stratégie de vente (NOUVEAU)"""
        # Pipeline par segment

    def create_sheet_financement(self):
        """Sheet 3: Financement détaillé"""
        # Levées détaillées

    def create_sheet_pl_full(self):
        """Sheet 4: P&L complet 50 mois"""
        # 122 colonnes

    def create_sheet_parametres(self):
        """Sheet 5: Paramètres"""
        # Pricing 2025-2029

    def create_sheet_gtmarket(self):
        """Sheet 6: GTMarket (NOUVEAU)"""

    def create_sheet_ventes_full(self):
        """Sheet 7: Ventes détaillées"""
        # Pipeline complet

    def create_sheet_sous_traitance(self):
        """Sheet 8: Sous traitance (NOUVEAU)"""

    def create_sheet_personnel(self):
        """Sheet 9: Charges personnel détail (NOUVEAU)"""
        # Le plus important - détail par rôle

    def create_sheet_direction(self):
        """Sheet 10: DIRECTION (NOUVEAU)"""

    def create_sheet_infrastructure(self):
        """Sheet 11: Infrastructure technique (NOUVEAU)"""

    def create_sheet_fundings(self):
        """Sheet 12: Fundings détaillé (NOUVEAU)"""

    def create_sheet_navigation(self):
        """Sheet 13: >> Navigation"""

    def create_sheet_positionnement(self):
        """Sheet 14: Positionnement (NOUVEAU)"""

    def create_sheet_marketing(self):
        """Sheet 15: Marketing (NOUVEAU)"""
```

### Phase 4: Mise à Jour Word (2-3h)
**Fichier**: `scripts/5_update_bm_word_full.py`

**Sections à mettre à jour**:
```python
def update_word_full():
    # Section 7.1 - Hypothèses clés
    update_section_7_1_hypotheses()

    # Section 7.2 - P&L prévisionnel
    update_section_7_2_pl()  # Vue 4 ans au lieu de 14 mois

    # Section 7.5 - KPIs financiers
    update_section_7_5_kpis()  # KPIs 2025-2029

    # Section 9.3 - Trajectoire valorisation
    update_section_9_3_valorisation()  # 8M€ (2026) → 50M€ (2029)

    # Section 10.2 - Jalons critiques
    update_section_10_2_jalons()  # 2025-2029
```

### Phase 5: Validation (1-2h)
**Validations à effectuer**:
```bash
# Vérifier cohérence Excel (formules, totaux)
python scripts/validate_excel_full.py

# Vérifier cohérence Excel ↔ Word
python scripts/7_validate_coherence.py

# Vérifier valorisations cohérentes 2026-2029
# ARR 2026: 827K€ → Valorisation 8M€
# ARR 2027: 1.4M€ → Valorisation 14M€
# ARR 2028: 2.4M€ → Valorisation 24M€
# ARR 2029: 3.6M€ → Valorisation 36M€
```

### Phase 6: Tests & Finition (1h)
- Tester ouverture Excel dans MS Excel / LibreOffice
- Vérifier formules actives
- Vérifier charts
- Tester modification assumptions → régénération complète

---

## ⏱️ Estimation Temps Total

| Phase | Tâches | Temps Estimé |
|-------|--------|--------------|
| 1 | Extension assumptions | 2-3h |
| 2 | Extension projections | 3-4h |
| 3 | Refonte Excel generator (15 sheets) | 6-8h |
| 4 | Mise à jour Word complète | 2-3h |
| 5 | Validation | 1-2h |
| 6 | Tests & finition | 1h |
| **TOTAL** | | **15-21h** |

**Estimation réaliste**: **2-3 jours de travail intensif**

---

## 🚨 Points d'Attention

### 1. Formules Excel Complexes
Les sheets source ont beaucoup de formules inter-sheets. Il faudra:
- Comprendre les dépendances
- Reproduire les formules exactement
- Tester que les calculs sont corrects

### 2. Données Manquantes
Certaines données du source ne sont pas dans nos assumptions actuelles:
- Détail sous-traitance
- Coûts infrastructure par service
- Budget marketing par canal
- Grilles salariales par rôle

**Solution**: Extraire du source ou définir valeurs raisonnables

### 3. Cohérence Temporelle
Le source couvre probablement Oct 2025 - Dec 2029.
Notre version commence Nov 2025.

**Solution**: Ajuster pour couvrir Nov 2025 - Dec 2029 (50 mois)

### 4. Maintenance
Un Excel de 15 sheets avec 50 mois de données sera complexe à maintenir.

**Solution**:
- Tout piloté par assumptions.yaml
- Scripts de régénération automatiques
- Validation automatique des formules

---

## ✅ Livrables Finaux

### Excel
- **Nom**: `BP_GenieFactory_2025-2029.xlsx`
- **Sheets**: 15
- **Couverture**: Nov 2025 - Dec 2029 (50 mois)
- **Colonnes P&L**: ~122
- **Formules**: Actives et fonctionnelles
- **Charts**: Intégrés
- **Taille estimée**: ~500 KB - 1 MB

### Word
- **Nom**: `BM_GenieFactory_2025-2029.docx`
- **Sections**: 10 (toutes mises à jour)
- **Couverture**: 2025-2029
- **Charts**: 6+ PNG insérés
- **Cohérence**: 100% avec Excel

### Validation
- ✅ Toutes formules Excel fonctionnelles
- ✅ Cohérence Excel ↔ Word (0% écart)
- ✅ Valorisations cohérentes (7-10x ARR)
- ✅ KPIs alignés sur toutes sections
- ✅ Timeline cohérente (2025-2029)

---

**Date**: 2025-11-20
**Auteur**: Claude Code
**Version**: 1.0 - Plan d'implémentation complet
