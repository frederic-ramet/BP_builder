# Spécifications Fonctionnelles - BP 14 Mois GenieFactory

## Vue d'Ensemble

Système de génération automatisée de Business Plan sur 14 mois (Nov 2025 - Dec 2026) à partir d'hypothèses centralisées, avec traçabilité complète et reproductibilité.

---

## F1 : Extraction de Données

### F1.1 : Parser BP Excel (openpyxl)

**Inputs** : `BP_FABRIQ_PRODUCT-OCT2025.xlsx`

**Outputs** : `data/structured/bp_extracted.json`

**Fonctionnalités** :
- Détecter structure automatiquement (header rows, data start)
- Extraire :
  * Pricing par offre (sheet "Paramètres")
  * CA mensuel par ligne métier (sheet "P&L")
  * Charges personnel (sheet "Charges de personnel et FG")
  * Infrastructure (sheet "Infrastructure technique")
  * Pipeline ventes (sheet "Ventes")
- Parser formules Excel → JSON structure
  ```json
  {
    "cell": "F2",
    "formula": "=SUM(F3:F8)",
    "dependencies": ["F3", "F4", "F5", "F6", "F7", "F8"],
    "logic": "sum_range"
  }
  ```

**Validation** :
- Vérifier présence sheets critiques
- Alert si formules circulaires détectées
- Log warning si cellules vides dans zones critiques

---

### F1.2 : Parser Documents Word (python-docx)

**Inputs** : 
- `Business_Plan_GenieFactory-SEPT2025.docx` (BM narratif)
- `GENIE_FACTORY_PACTE_AATL-v3.docx` (Pacte actionnaires)

**Outputs** : `data/structured/word_extracted.json`

**Fonctionnalités** :
- Extraire tableaux financiers du BM (section 7.2, 7.3)
- Extraire milestones ARR du pacte :
  * Pattern regex : `ARR\s*[≥>=]\s*(\d+)\s*[K€]`
  * Extraire : 800K€, 1.5M€
- Extraire hypothèses croissance :
  * Conversion rates, churn, CAC/LTV
  * Pattern : `taux de conversion.*?(\d+)%`

**Validation** :
- ARR milestones du pacte > 0
- Au moins 2 tableaux financiers extraits du BM

---

## F2 : Génération Assumptions

### F2.1 : Créer assumptions.yaml

**Inputs** : 
- `bp_extracted.json`
- `word_extracted.json`
- User validation (interactive prompts)

**Output** : `data/structured/assumptions.yaml`

**Structure YAML** :
```yaml
meta:
  version: "1.0"
  created_date: "2025-01-XX"
  sources:
    - "BP_FABRIQ_PRODUCT-OCT2025.xlsx"
    - "Business_Plan_GenieFactory-SEPT2025.docx"
    - "GENIE_FACTORY_PACTE_AATL-v3.docx"

timeline:
  start_month: "2025-11"
  duration_months: 14
  fiscal_year_start: 11  # Novembre
  
milestones:
  - month: 1
    event: "Pre-seed"
    amount_eur: 150000
    notes: "Prêts Autoposia (50K) + F-Initiatives (40K) + BPI (30K) + CIC (30K)"
  
  - month: 11
    event: "Seed Round"
    amount_eur: 500000
    valuation_pre_money: 2500000
    notes: "Target ARR ~450K€ avant seed"
    
  - month: 14
    event: "ARR Milestone (Pacte)"
    arr_target: 800000
    notes: "Milestone contractuel pacte actionnaires"

pricing:
  # Source: BP Oct 2025, sheet Paramètres
  hackathon:
    m1_m6: 18000      # Nov 2025 - Avril 2026
    m7_m14: 20000     # Mai 2026 - Dec 2026
    evolution_logic: "+10% annuel"
    
  factory:
    m1_m6: 75000
    m7_m14: 82500     # +10%
    
  services_implementation:
    m1_m6: 10000
    m7_m14: 17500     # Progression +75% (maturité offre)
    
  formation:
    m1_m6: 5000
    m7_m14: 5500
    
  enterprise_hub:
    starter_monthly: 500
    business_monthly: 2000
    enterprise_monthly: 10000
    start_month: 8    # Lancement Juin 2026
    ramp_months: 6    # 6 mois pour atteindre capacité

sales_assumptions:
  hackathon:
    m1_m3: 1.5        # Démarrage lent (1-2/mois)
    m4_m6: 2.5
    m7_m10: 3         # Post-product-market fit
    m11_m14: 4        # Post-seed acceleration
    
  factory:
    conversion_from_hackathon: 0.30
    delay_months: 2   # Signature 2 mois après hackathon
    
  enterprise_hub:
    # Lancement progressif M8-M14
    new_starter_monthly: [0,0,0,0,0,0,0, 2,2,3,4,4,5,6]  # Index 0 = M1
    starter_to_business_rate: 0.20  # 20% upgrade après 3 mois
    business_to_enterprise_rate: 0.10
    churn_annual: 0.10

costs:
  personnel:
    team_evolution:
      m1: 5    # Fondateurs (FRT, PCO, MAM, JBT) + 1 dev
      m3: 7    # +2 dev
      m7: 9    # +2 commercial
      m11: 11  # +2 customer success (post-seed)
      m14: 12  # +1 ops/admin
    
    avg_loaded_salary:
      founders: 0      # Pas de salaire année 1
      employee: 6000   # 6K€/mois chargé
      
    freelance_budget_monthly: 5000  # Dev + design freelance
  
  infrastructure:
    base_monthly: 2000    # Servers, SaaS tools
    per_client_monthly: 200  # Scaling cost
    
  marketing:
    base_monthly: 5000
    events_quarterly: 15000
    content_monthly: 2000
    
  office_admin:
    monthly: 3000

revenue_mix:
  # Répartition attendue sur 14 mois
  hackathon: 0.45       # 45%
  factory: 0.30         # 30%
  enterprise_hub: 0.15  # 15%
  services: 0.10        # 10%

financial_kpis:
  target_arr_dec_2026: 800000
  target_arr_sept_2026: 450000  # Avant seed
  acceptable_burn_rate_monthly: 50000
  acceptable_ebitda_negative: true  # OK pendant phase croissance
  cash_runway_months_min: 12

validation_rules:
  arr_tolerance_pct: 0.10  # ±10% vs target
  max_team_size: 15
  max_burn_monthly: 60000
  min_conversion_hackathon_factory: 0.25
```

**Fonctionnalités** :
- Interactive prompts pour valeurs non extraites automatiquement
- Validation ranges (ex: churn entre 0-30%)
- Export commenté (chaque section avec explication)

---

## F3 : Calculs Financiers

### F3.1 : Calculer Projections Mensuelles

**Input** : `assumptions.yaml`

**Output** : `data/structured/projections.json`

**Algorithme** :

```python
def calculate_monthly_projections(assumptions):
    months = []
    
    for m in range(1, 15):  # M1 à M14
        month_data = {
            'month': m,
            'date': get_date(m),
            'revenue': {},
            'costs': {},
            'metrics': {}
        }
        
        # REVENUS
        # 1. Hackathons
        nb_hackathons = get_hackathon_volume(m, assumptions)
        price_hackathon = get_hackathon_price(m, assumptions)
        rev_hackathon = nb_hackathons * price_hackathon
        
        # 2. Factory (delayed conversion from hackathons)
        nb_factory = get_factory_conversions(m, months, assumptions)
        price_factory = get_factory_price(m, assumptions)
        rev_factory = nb_factory * price_factory
        
        # 3. Enterprise Hub (SaaS MRR)
        if m >= assumptions['pricing']['enterprise_hub']['start_month']:
            mrr_hub = calculate_hub_mrr(m, months, assumptions)
            rev_hub = mrr_hub
        else:
            rev_hub = 0
        
        # 4. Services
        rev_services = calculate_services(m, nb_hackathons, nb_factory, assumptions)
        
        month_data['revenue'] = {
            'hackathon': rev_hackathon,
            'factory': rev_factory,
            'enterprise_hub': rev_hub,
            'services': rev_services,
            'total': rev_hackathon + rev_factory + rev_hub + rev_services
        }
        
        # COÛTS
        team_size = get_team_size(m, assumptions)
        month_data['costs'] = {
            'personnel': team_size * assumptions['costs']['personnel']['avg_loaded_salary']['employee'],
            'freelance': assumptions['costs']['personnel']['freelance_budget_monthly'],
            'infrastructure': calculate_infra_costs(m, months, assumptions),
            'marketing': calculate_marketing_costs(m, assumptions),
            'admin': assumptions['costs']['office_admin']['monthly']
        }
        month_data['costs']['total'] = sum(month_data['costs'].values()) - month_data['costs']['total']
        
        # MÉTRIQUES
        # ARR = MRR × 12 (uniquement Hub SaaS)
        month_data['metrics']['mrr'] = rev_hub
        month_data['metrics']['arr'] = calculate_arr(m, months)
        month_data['metrics']['ebitda'] = month_data['revenue']['total'] - month_data['costs']['total']
        month_data['metrics']['burn_rate'] = -month_data['metrics']['ebitda'] if month_data['metrics']['ebitda'] < 0 else 0
        month_data['metrics']['team_size'] = team_size
        month_data['metrics']['cash'] = calculate_cash_position(m, months, assumptions)
        
        months.append(month_data)
    
    return months
```

**Validation** :
- ARR M14 dans range [720K€ - 880K€]
- ARR M11 >= 400K€ (attractive pour seed)
- Burn rate max <= 60K€/mois
- Cash jamais négatif

---

### F3.2 : Calculer ARR (récurrent uniquement)

```python
def calculate_arr(month_index, historical_months):
    """
    ARR = Annualized Run Rate des revenus récurrents uniquement
    = MRR actuel × 12
    """
    
    # Seulement Enterprise Hub est récurrent
    current_mrr = historical_months[month_index]['revenue']['enterprise_hub']
    
    # ARR = MRR × 12
    arr = current_mrr * 12
    
    return arr
```

**Note** : Hackathons et Factory sont one-time revenue, pas ARR.

---

## F4 : Génération BP Excel

### F4.1 : Structure Excel Identique

**Input** : 
- `projections.json`
- `BP_FABRIQ_PRODUCT-OCT2025.xlsx` (template structure)

**Output** : `BP_14M_Nov2025-Dec2026.xlsx`

**Sheets à générer** :

1. **Synthèse** (dashboard)
   - KPIs clés : CA total, ARR, EBITDA, Burn rate
   - Graphiques : Courbe ARR, Burn mensuel
   
2. **P&L** (détail mensuel)
   ```
   Row 1: Période | M1 (Nov 25) | M2 | ... | M14 (Dec 26)
   Row 2: CA TOTAL | =SUM(3:8) | =SUM(3:8) | ...
   Row 3: Hackathon | 36000 | 36000 | ...
   Row 4: Factory | 0 | 0 | 75000 | ...  (delayed)
   Row 5: Enterprise Hub | 0 | ... | 10000 | ...  (start M8)
   Row 6: Services | 10000 | ...
   Row 9: CHARGES TOTAL | =SUM(10:13)
   Row 10: Personnel | ...
   Row 11: Infra | ...
   Row 14: EBITDA | =2-9
   Row 15: Burn rate | =IF(14<0, -14, 0)
   Row 16: ARR | =5*12  (Hub MRR × 12)
   ```

3. **Ventes** (pipeline détaillé)
   - Hackathons planifiés par mois
   - Factory conversions
   - Hub new customers
   
4. **Paramètres** (pricing reference)
   - Grille tarifaire période 1 vs période 2
   
5. **Financement**
   - Pre-seed M1 : 150K€
   - Seed M11 : 500K€
   - Utilisation fonds

**Formules Excel Critiques** :
```excel
# CA Total mensuel
=SUM(CA_Hackathon:CA_Services)

# ARR (uniquement Hub)
=Hub_MRR * 12

# Cash position
=Cash_précédent + CA_mois - Charges_mois + Funding_mois

# Validation ARR M14
=IF(ARR_M14 < 720000, "⚠️ ARR sous target", IF(ARR_M14 > 880000, "⚠️ ARR sur-optimiste", "✓ ARR OK"))
```

---

### F4.2 : Charts & Formatting

**Graphiques à créer** :

1. **ARR Growth Curve**
   - Type : Line chart
   - X-axis : M1 → M14
   - Y-axis : ARR (0 → 900K€)
   - Target line : 800K€ à M14
   
2. **Burn Rate Mensuel**
   - Type : Column chart
   - X-axis : M1 → M14
   - Y-axis : Burn (0 → 60K€)
   - Color : Rouge si >50K€

3. **Revenue Mix**
   - Type : Stacked area
   - Catégories : Hackathon, Factory, Hub, Services
   
**Formatting** :
- Headers : Bold, background #4472C4
- Totaux : Bold, background #D9E1F2
- EBITDA négatif : Red text
- ARR : Green text, bold
- Currency : # ##0 € (espace comme séparateur)

---

## F5 : Update BM Word

### F5.1 : Identifier Sections à Modifier

**Input** : `Business_Plan_GenieFactory-SEPT2025.docx`

**Sections cibles** :
- **7.2 Compte de résultat prévisionnel** (tableau)
- **7.3 Plan de financement** (tableau)
- **7.4 KPIs financiers cibles** (paragraphe)

**Approche** :
1. Identifier tableaux par titre section (heading scan)
2. Remplacer contenu tableau (garder structure)
3. Update paragraphes avec nouvelles métriques

---

### F5.2 : Générer Nouveau Tableau P&L

**Ancien format** (4 colonnes : 2025-2026, 2027, 2028, 2029)

**Nouveau format** (14 colonnes : M1-M14 + Total)

```
| Métrique (K€) | M1 | M2 | ... | M14 | TOTAL |
|---------------|----|----|-----|-----|-------|
| CA Total      | 36 | 56 | ... | 140 | 1050  |
| - Hackathon   | 36 | 36 | ... | 80  | 630   |
| - Factory     | 0  | 0  | ... | 75  | 270   |
| - Hub         | 0  | 0  | ... | 30  | 120   |
| - Services    | 0  | 20 | ... | 15  | 120   |
| Charges       | 45 | 48 | ... | 65  | 720   |
| EBITDA        |-9  | 8  | ... | 75  | 330   |
| ARR           | 0  | 0  | ... | 800 | -     |
```

---

### F5.3 : Update KPIs Textuels

**Pattern replacement** :

```python
replacements = {
    # Timeline
    r"2025-2027": "Nov 2025 - Dec 2026 (14 mois)",
    r"objectif 2029": "objectif Dec 2026",
    
    # ARR
    r"ARR:\s*320K€\s*\(2025\)": "ARR: 0€ (M1) → 800K€ (M14)",
    r"1\.4M€\s*\(2026\)": "800K€ (Dec 2026)",
    
    # Croissance
    r"\+330%.*2026": "+2000% (démarrage M1 → M14)",
    
    # Break-even
    r"Q1 2026": "Non attendu (croissance prioritaire)",
    
    # Seed
    r"350K€.*Seed": "500K€ (Seed Sept 2026)",
}
```

---

## F6 : Validation Automatique

### F6.1 : Checks Financiers

```python
def validate_projections(projections, assumptions):
    errors = []
    warnings = []
    
    # Check 1: ARR M14 dans target ±10%
    arr_m14 = projections[-1]['metrics']['arr']
    target = assumptions['financial_kpis']['target_arr_dec_2026']
    tolerance = target * assumptions['validation_rules']['arr_tolerance_pct']
    
    if arr_m14 < target - tolerance:
        errors.append(f"ARR M14 trop bas: {arr_m14}€ (target {target}€)")
    elif arr_m14 > target + tolerance:
        warnings.append(f"ARR M14 optimiste: {arr_m14}€ (target {target}€)")
    
    # Check 2: ARR M11 (avant seed) >= 400K€
    arr_m11 = projections[10]['metrics']['arr']
    if arr_m11 < 400000:
        warnings.append(f"ARR avant seed faible: {arr_m11}€ (min conseillé 400K€)")
    
    # Check 3: Burn rate max
    max_burn = max([m['metrics']['burn_rate'] for m in projections])
    if max_burn > assumptions['validation_rules']['max_burn_monthly']:
        errors.append(f"Burn rate max {max_burn}€ > limite {assumptions['validation_rules']['max_burn_monthly']}€")
    
    # Check 4: Cash jamais négatif
    for m in projections:
        if m['metrics']['cash'] < 0:
            errors.append(f"Cash négatif M{m['month']}: {m['metrics']['cash']}€")
    
    # Check 5: Équipe taille raisonnable
    team_m14 = projections[-1]['metrics']['team_size']
    if team_m14 > assumptions['validation_rules']['max_team_size']:
        warnings.append(f"Équipe large M14: {team_m14} ETP (max conseillé {assumptions['validation_rules']['max_team_size']})")
    
    # Check 6: Conversion hackathon → factory
    total_hackathons = sum([m['volume']['hackathon'] for m in projections])
    total_factory = sum([m['volume']['factory'] for m in projections])
    conversion_rate = total_factory / total_hackathons if total_hackathons > 0 else 0
    
    if conversion_rate < assumptions['validation_rules']['min_conversion_hackathon_factory']:
        warnings.append(f"Conversion Hackathon→Factory faible: {conversion_rate:.0%}")
    
    return {
        'errors': errors,
        'warnings': warnings,
        'status': 'FAILED' if errors else 'OK'
    }
```

### F6.2 : Checks Cohérence

```python
def validate_consistency(bp_excel, bm_word, projections):
    """Vérifier cohérence entre Excel et Word"""
    
    # Extraire ARR M14 du Word
    arr_word = extract_arr_from_word(bm_word)
    arr_excel = projections[-1]['metrics']['arr']
    
    if abs(arr_word - arr_excel) > 1000:  # Tolérance 1K€
        return {
            'error': f"ARR incohérent: Word={arr_word}€, Excel={arr_excel}€"
        }
    
    # Vérifier CA total
    ca_total_excel = sum([m['revenue']['total'] for m in projections])
    ca_total_word = extract_ca_from_word(bm_word)
    
    if abs(ca_total_excel - ca_total_word) / ca_total_excel > 0.05:  # 5% tolérance
        return {
            'error': f"CA total incohérent: Word={ca_total_word}€, Excel={ca_total_excel}€"
        }
    
    return {'status': 'OK'}
```

---

## F7 : Documentation & Logs

### F7.1 : Génération README.md

```markdown
# GenieFactory - Business Plan 14 Mois (Nov 2025 - Dec 2026)

## Quickstart

```bash
pip install -r requirements.txt
python run.py  # Génère BP Excel + BM Word
```

## Ajuster les Hypothèses

Éditer `data/structured/assumptions.yaml` :

- **Timeline** : Modifier dates milestones
- **Pricing** : Ajuster tarifs hackathon/factory/hub
- **Volumes** : Modifier nb hackathons/mois
- **Équipe** : Ajuster team_evolution

Puis regénérer :
```bash
python scripts/3_calculate_projections.py
python scripts/4_generate_bp_excel.py
python scripts/5_update_bm_word.py
```

## Structure

- `assumptions.yaml` : Source unique vérité
- `projections.json` : Calculs intermédiaires
- `BP_14M_*.xlsx` : Business Plan Excel
- `BM_Updated_*.docx` : Business Model Word

## Métriques Clés

- ARR M14 : 800K€ (±10%)
- CA Total 14M : ~1.05M€
- Équipe M14 : 12 ETP
- Seed M11 : 500K€

## Hypothèses Principales

- Conversion Hackathon→Factory : 30%
- Churn annuel Hub : 10%
- Démarrage Hub : M8 (Juin 2026)
- Burn rate moyen : 35K€/mois

## Validation

```bash
python scripts/6_validate.py
```

Checks :
- ✓ ARR M14 dans target
- ✓ Cash position positive
- ✓ Burn rate acceptable
- ✓ Cohérence Excel ↔ Word
```

### F7.2 : Logs Détaillés

À chaque exécution, générer `logs/run_YYYYMMDD_HHMMSS.log` :

```
[2025-01-15 10:23:45] START - BP Generation
[2025-01-15 10:23:46] EXTRACT - BP Excel parsed (14 sheets, 1004 rows)
[2025-01-15 10:23:47] EXTRACT - BM Word parsed (3 tables, 423 paragraphs)
[2025-01-15 10:23:48] ASSUMPTIONS - Generated assumptions.yaml (142 lines)
[2025-01-15 10:23:50] CALCULATE - M1: CA=36K€ (2 hackathons × 18K€)
[2025-01-15 10:23:50] CALCULATE - M2: CA=56K€ (2 hackathons + 1 services)
...
[2025-01-15 10:23:55] CALCULATE - M14: CA=140K€, ARR=800K€ ✓
[2025-01-15 10:23:58] GENERATE - BP Excel créé (14 colonnes, 8 sheets)
[2025-01-15 10:24:02] GENERATE - BM Word mis à jour (sections 7.2, 7.3, 7.4)
[2025-01-15 10:24:05] VALIDATE - ARR M14: 800K€ ✓ (target 800K€ ±10%)
[2025-01-15 10:24:05] VALIDATE - Burn max: 48K€ ✓ (<60K€)
[2025-01-15 10:24:05] VALIDATE - Cash min: 85K€ ✓ (>0)
[2025-01-15 10:24:05] VALIDATE - ⚠️ Warning: ARR M11 = 380K€ (recommandé >400K€)
[2025-01-15 10:24:05] END - SUCCESS (durée: 20s)
```

---

## Dépendances Python

```requirements.txt
openpyxl>=3.1.0
python-docx>=1.1.0
pyyaml>=6.0
pandas>=2.0.0
xlsxwriter>=3.1.0
pytest>=7.4.0
```

---

## Tests Unitaires

```python
# tests/test_calculations.py

def test_arr_calculation():
    """ARR = MRR × 12"""
    mrr = 66667  # 800K/12
    arr = calculate_arr_from_mrr(mrr)
    assert arr == 800004  # ~800K€

def test_factory_conversion():
    """30% hackathons → Factory avec 2 mois délai"""
    hackathons_m1 = 2
    hackathons_m2 = 2
    factory_m3 = calculate_factory(month=3, hackathons_history=[2,2,0,...])
    
    expected = (2+2) * 0.30  # 1.2 → arrondi 1
    assert factory_m3 == 1

def test_hub_ramp():
    """Lancement Hub M8, croissance progressive"""
    new_customers = [0,0,0,0,0,0,0, 2,2,3,4,4,5,6]
    
    # M8 : 2 nouveaux Starter
    customers_m8 = sum(new_customers[:8])
    assert customers_m8 == 2
    
    # M14 : cumul 26 customers
    customers_m14 = sum(new_customers)
    assert customers_m14 == 26
```

---

## Livrables Attendus

✅ **assumptions.yaml** : 150 lignes, commenté, sourcé
✅ **projections.json** : Calculs mensuels M1-M14
✅ **BP_14M_Nov2025-Dec2026.xlsx** : 8 sheets, formules actives
✅ **BM_Updated_14M.docx** : Sections 7.x actualisées
✅ **README.md** : Guide complet
✅ **6_validate.py** : Tous checks passing
✅ **Tests** : Coverage >80%

---

## Critères d'Acceptance

### Critique (Must Have)
- [ ] ARR M14 = 800K€ ± 10%
- [ ] ARR M11 >= 400K€ (attractif seed)
- [ ] Cash position jamais négative
- [ ] Formules Excel fonctionnelles (pas de hardcoding)
- [ ] Cohérence Excel ↔ Word (<5% écart)

### Important (Should Have)
- [ ] Burn rate max < 60K€/mois
- [ ] Équipe M14 <= 15 ETP
- [ ] Conversion Hackathon→Factory >= 25%
- [ ] Tests unitaires passing
- [ ] Documentation complète

### Nice to Have
- [ ] Charts Excel professionnels
- [ ] Logs détaillés avec timestamps
- [ ] Interactive CLI pour ajuster assumptions
- [ ] Export PDF du BP

---

**Estimation Effort Total** : 6-8h
- Extraction : 1h
- Assumptions : 1h
- Calculs : 2h
- Génération Excel : 2h
- Update Word : 1h
- Validation + Tests : 1h
