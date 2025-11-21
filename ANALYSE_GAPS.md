# Analyse des Écarts - Fichiers Source vs Générés

## 📊 EXCEL - Analyse Comparative

### Fichier Source: `BP FABRIQ_PRODUCT-OCT2025.xlsx`
**Caractéristiques:**
- **15 sheets** au total
- **P&L avec 122 colonnes** (couvre plusieurs années: 2025-2029 probablement)
- Dimensions massives (jusqu'à 1004 lignes × 122 colonnes)

**Sheets présents dans la source:**
1. ✅ Synthèse
2. ❌ **Stratégie de vente** (MANQUANT)
3. ✅ Financement (existe mais simplifié)
4. ✅ P&L (existe mais 14 colonnes vs 122)
5. ✅ Paramètres
6. ❌ **GTMarket** (MANQUANT)
7. ✅ Ventes
8. ❌ **Sous traitance** (MANQUANT)
9. ❌ **Charges de personnel et FG** (MANQUANT - détails complets)
10. ❌ **DIRECTION** (MANQUANT)
11. ❌ **Infrastructure technique** (MANQUANT)
12. ❌ **Fundings** (MANQUANT - détails complets)
13. ❌ **Positionnement** (MANQUANT)
14. ❌ **Marketing** (MANQUANT)
15. ✅ Monitoring (ajouté mais pas dans source)

### Fichier Généré: `BP_14M_Nov2025-Dec2026.xlsx`
**Caractéristiques:**
- **6 sheets** seulement
- **P&L avec 14 colonnes** (Nov 2025 - Dec 2026)
- Focus sur 14 mois uniquement

**Sheets générés:**
1. Synthèse ✅
2. P&L ✅ (simplifié)
3. Ventes ✅ (simplifié)
4. Paramètres ✅
5. Financement ✅
6. Monitoring ✅ (nouveau, pas dans source)

### 🔴 ÉCARTS CRITIQUES EXCEL:

#### 1. Couverture Temporelle
- **Source**: Plusieurs années (38 mois probablement, 2025-2029)
- **Généré**: 14 mois seulement (Nov 2025 - Dec 2026)
- **Impact**: Perte de vision long terme

#### 2. Sheets Manquants (9 sheets)
Sheets critiques absents:
- **Stratégie de vente**: Détail pipeline, conversion, canaux
- **GTMarket**: Go-to-market détaillé
- **Charges de personnel et FG**: Détail complet équipe + frais généraux
- **Infrastructure technique**: Coûts tech détaillés
- **Sous traitance**: Coûts externes
- **Fundings**: Détail levées de fonds
- **Marketing**: Budget marketing détaillé
- **DIRECTION**: Organigramme direction
- **Positionnement**: Analyse concurrentielle

#### 3. Détails Manquants dans Sheets Existants

**P&L:**
- Source: 122 colonnes (plusieurs années, détails mensuels)
- Généré: 14 colonnes (14 mois)
- Manque: Années 2027-2029, détails par ligne de revenus/coûts

**Ventes:**
- Source: 70 colonnes, 967 lignes (détail complet pipeline)
- Généré: 14 colonnes, 7 lignes (très simplifié)
- Manque: Pipeline détaillé, forecast par client, probabilités

**Charges de personnel:**
- Source: Sheet dédié 72 colonnes × 1010 lignes
- Généré: Rien (juste agrégé dans P&L)
- Manque: Détail par poste, salaires, charges sociales, primes

---

## 📄 WORD - Analyse Comparative

### Fichier Source: `Business Plan GenieFactory-SEPT2025.docx`
**Caractéristiques:**
- **423 paragraphes**
- **3 tableaux**
- **10 sections principales** complètes
- Document professionnel structuré

**Structure complète:**
1. Présentation de l'activité (1.1-1.3)
2. Analyse de marché (2.1-2.4)
3. Offre et modèle économique (3.1-3.4)
4. Stratégie Go-to-Market (4.1-4.5)
5. Équipe et organisation (5.1-5.5)
6. Roadmap produit et technologie (6.1-6.4)
7. **Projections financières 2025-2029** (7.1-7.5) ← Section cible
8. Analyse des risques et mitigations (8.1-8.5)
9. Ambition et stratégie de sortie (9.1-9.4)
10. Conclusion et prochaines étapes (10.1-10.5)

### Fichier Généré: `BM_Updated_14M.docx`
**Modifications apportées:**
- ✅ **Executive Summary ajouté** (nouveau, en début de document)
- ✅ **Tableau synthèse financière ajouté** (M1, M6, M11, M14)
- ✅ **Section "Demande de Financement" ajoutée**
- ✅ **6 graphiques PNG insérés**
- ✅ **Section 7.2 mise à jour** (tableau P&L 14 mois)
- ⚠️ **Reste du document préservé tel quel**

### 🔴 ÉCARTS CRITIQUES WORD:

#### 1. Sections Non Mises à Jour
Les sections suivantes du document original ne sont **PAS mises à jour** avec les nouvelles données:

**Section 7.1 - Hypothèses clés:**
- ❌ Pas mis à jour avec nos assumptions.yaml
- Contient probablement les anciennes hypothèses 38 mois

**Section 7.3 - Plan de financement:**
- ⚠️ Possiblement redondant avec notre nouvelle section
- Doit être vérifié et synchronisé

**Section 7.4 - Utilisation des fonds Seed:**
- ⚠️ Peut-être redondant avec notre section "Demande de Financement"
- Doit être cohérent

**Section 7.5 - KPIs financiers cibles:**
- ❌ Pas mis à jour avec nos métriques M14 (ARR 827K€)
- Contient probablement anciens KPIs

**Section 9.3 - Trajectoire de valorisation:**
- ❌ Pas mis à jour avec la valorisation cohérente (8M€)
- Peut contenir les anciennes valorisations incohérentes

**Section 10.2 - Jalons critiques 2025-2027:**
- ❌ Pas mis à jour avec notre timeline 14 mois
- Doit refléter M11 Seed, M14 ARR target

#### 2. Cohérence Temporelle
- Document original couvre 2025-2029 (plusieurs années)
- Nos données couvrent Nov 2025 - Dec 2026 (14 mois)
- **Incohérence**: Les sections non mises à jour référencent 2027-2029

#### 3. Tableaux Originaux
- Le document original contient **3 tableaux**
- Nous en ajoutons mais ne mettons pas à jour les existants
- Risque: Tableaux originaux avec données obsolètes

---

## 🎯 PLAN D'IMPLÉMENTATION

### PRIORITÉ 1: Excel - Enrichissement Critique

#### A. Étendre P&L sur 4 ans (2025-2029)
**Objectif**: Passer de 14 colonnes à ~50 colonnes (4 ans)
**Approche**:
1. Étendre assumptions.yaml avec hypothèses 2027-2029
2. Étendre calculate_projections.py pour calculer M15-M48
3. Adapter generate_bp_excel.py pour colonnes additionnelles
4. Ajouter sections "Projections LT" avec croissance post-14M

**Détails techniques**:
```python
# assumptions.yaml - ajouter
long_term_assumptions:
  years_2027_2029:
    arr_growth_rate: 0.80  # +80%/an
    team_growth: 5  # +5 ETP/an
    # ...
```

#### B. Ajouter Sheet "Charges Personnel Détail"
**Colonnes**: M1-M48 (4 ans)
**Lignes**:
- CEO / CTO / CPO / CMO / ...
- Développeurs (détail par seniority)
- Sales / Marketing
- Salaires bruts
- Charges sociales (45%)
- Primes / Bonus
- Total par mois

#### C. Ajouter Sheet "Infrastructure Technique"
**Lignes**:
- AWS / Cloud (scaling avec clients)
- SaaS tools (Notion, Slack, ...)
- Licences logicielles
- R&D externe
- Total mensuel

#### D. Ajouter Sheet "Marketing & Sales"
**Détail**:
- Budget pub digitale (Google, LinkedIn)
- Events / Conférences
- Content marketing
- Outils CRM/Marketing
- Coût d'acquisition client (CAC) calculé

### PRIORITÉ 2: Word - Mise à Jour Complète

#### A. Section 7.1 - Hypothèses clés
**Action**: Remplacer contenu par notre assumptions.yaml
**Contenu**:
- Pricing (Hackathon 18K→20K, Factory 75K→82K, Hub 500-10K€)
- Volumes mensuels (croissance détaillée)
- Churn Hub (10%)
- Conversion Hack→Factory (30%)
- Timeline milestones (Pre-seed M1, Seed M11)

#### B. Section 7.5 - KPIs financiers cibles
**Action**: Remplacer avec nos KPIs réels
**Contenu**:
```
KPIs M14 (Dec 2026):
- ARR: 827K€ ✓
- CA Total 14M: 2,180K€
- MRR: 69K€
- Clients Hub: 35
- Équipe: 12 ETP
- Cash: 2,096K€

KPIs M11 (Sept 2026) - Avant Seed:
- ARR: 343K€
- CA: 220K€
- Clients Hub: 15
- Équipe: 11 ETP
```

#### C. Section 9.3 - Trajectoire de valorisation
**Action**: Remplacer avec valorisations cohérentes
**Contenu**:
```
Valorisation cohérente SaaS B2B (multiples 7-10x ARR):

2026 (M14 - Dec):
- ARR: 827K€
- Valorisation: 5-8M€ (6-10x)
- Post-Seed valuation estimée

2027-2028:
- ARR projeté: 3-5M€ (croissance 80%/an)
- Valorisation: 21-50M€ (7-10x)

2029:
- ARR projeté: 8-10M€
- Valorisation: 56-100M€ (7-10x)
- Préparation Series A
```

#### D. Sections 2-6 - Vérification Cohérence
**Action**: Audit complet pour vérifier que les sections stratégiques sont cohérentes avec notre timeline 14 mois

**Vérifications**:
- Section 4.1 (Phases déploiement): Timeline cohérente avec M8 Hub launch?
- Section 5.3 (Plan recrutement): Effectifs cohérents avec notre projection 5→12 ETP?
- Section 6 (Roadmap): V1/V2/V3 cohérentes avec 14 mois?

---

## 📋 RÉSUMÉ EXÉCUTIF

### Excel - Écarts Critiques
| Élément | Source | Généré | Écart |
|---------|--------|---------|-------|
| Nombre sheets | 15 | 6 | -9 sheets (60%) |
| Colonnes P&L | 122 (4 ans) | 14 (14 mois) | -108 cols |
| Détail personnel | 1 sheet dédié | Agrégé | Perte détail |
| Marketing | 1 sheet dédié | Absent | Perte détail |
| Infrastructure | 1 sheet dédié | Absent | Perte détail |

### Word - Écarts Critiques
| Section | Source | Mise à Jour | Écart |
|---------|--------|-------------|-------|
| 7.1 Hypothèses | Anciennes | ❌ Non | Incohérence |
| 7.2 P&L | 2025-2029 | ✅ Oui (14M) | Partiel |
| 7.5 KPIs | Anciens | ❌ Non | Incohérence |
| 9.3 Valorisation | Ancienne | ❌ Non | Incohérence |
| Cohérence globale | 4 ans | 14 mois | Décalage |

---

## 🚨 RISQUES SI NON CORRIGÉ

1. **Incohérence investisseurs**: Document Word référence 2027-2029 mais Excel s'arrête en 2026
2. **Crédibilité**: Sections non mises à jour avec anciennes données
3. **Décisions**: Manque de visibilité long terme (pas de projection 2027-2029)
4. **Due diligence**: Absence de détails (personnel, infrastructure, marketing)

---

## ✅ PROCHAINES ACTIONS RECOMMANDÉES

### Phase 1: Excel Enrichissement (Priorité HAUTE)
1. Étendre projections à 4 ans (M1-M48)
2. Ajouter sheet "Charges Personnel Détail"
3. Ajouter sheet "Infrastructure Technique"
4. Ajouter sheet "Marketing & Sales Détail"

### Phase 2: Word Mise à Jour Complète (Priorité HAUTE)
1. Remplacer section 7.1 (Hypothèses)
2. Remplacer section 7.5 (KPIs)
3. Remplacer section 9.3 (Valorisation)
4. Audit cohérence sections 2-6

### Phase 3: Validation (Priorité CRITIQUE)
1. Vérifier cohérence Excel ↔ Word sur 4 ans
2. Valider valorisations cohérentes (7-10x ARR)
3. S'assurer aucune référence aux anciennes données

---

**Date d'analyse**: 2025-11-20
**Fichiers analysés**:
- Source Excel: `data/raw/BP FABRIQ_PRODUCT-OCT2025.xlsx`
- Source Word: `data/raw/Business Plan GenieFactory-SEPT2025.docx`
- Généré Excel: `data/outputs/BP_14M_Nov2025-Dec2026.xlsx`
- Généré Word: `data/outputs/BM_Updated_14M.docx`
