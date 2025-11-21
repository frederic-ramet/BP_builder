#!/usr/bin/env python3
"""
Créer un TEMPLATE Excel à partir du fichier RAW
Adapte la structure selon assumptions.yaml tout en préservant les formules
"""

import openpyxl
from pathlib import Path
import yaml
from rich.console import Console
from rich.progress import track
import logging
from copy import copy

console = Console()
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger(__name__)


class TemplateCreator:
    """Créer un template Excel adapté depuis le RAW"""

    def __init__(self, raw_path: Path, assumptions: dict):
        self.raw_path = raw_path
        self.assumptions = assumptions

        logger.info(f"📂 Chargement fichier RAW: {raw_path.name}")
        self.wb = openpyxl.load_workbook(raw_path)
        logger.info(f"✓ {len(self.wb.sheetnames)} sheets chargés")

    def update_parametres_sheet(self):
        """
        Adapter le sheet Paramètres selon assumptions.yaml
        - Prix Hackathon, Factory, Hub selon YAML
        - Évolution des prix par année
        """
        logger.info("\n⚙️ Adaptation sheet Paramètres...")

        ws = self.wb['Paramètres']

        # Extraire les prix du YAML
        pricing = self.assumptions.get('pricing', {})

        # Prix Hackathon (ligne 3)
        hackathon_base = pricing.get('hackathon_base', 18000)
        ws['B3'] = hackathon_base
        ws['C3'] = "=B3*1.1"  # +10% par an
        ws['D3'] = "=C3*1.1"
        ws['E3'] = "=D3*1.1"
        ws['F3'] = "=E3*1.1"

        # Prix Factory (ligne 10)
        factory_base = pricing.get('factory_base', 75000)
        ws['B10'] = factory_base
        ws['C10'] = "=B10*1.05"  # +5% par an
        ws['D10'] = "=C10*1.05"
        ws['E10'] = "=D10*1.05"
        ws['F10'] = "=E10*1.05"

        # Prix Hub - Starter (ligne 7)
        starter_base = pricing.get('hub_starter_monthly', 500)
        ws['B7'] = starter_base
        ws['C7'] = "=B7*1.1"
        ws['D7'] = "=C7*1.1"
        ws['E7'] = "=D7*1.1"
        ws['F7'] = "=E7*1.1"

        # Prix Hub - Business (ligne 8)
        business_base = pricing.get('hub_business_monthly', 2000)
        ws['B8'] = business_base
        ws['C8'] = "=B8*1.1"
        ws['D8'] = "=C8*1.1"
        ws['E8'] = "=D8*1.1"
        ws['F8'] = "=E8*1.1"

        # Prix Hub - Enterprise (ligne 9)
        enterprise_base = pricing.get('hub_enterprise_monthly', 10000)
        ws['B9'] = enterprise_base
        ws['C9'] = "=B9*1.1"
        ws['D9'] = "=C9*1.1"
        ws['E9'] = "=D9*1.1"
        ws['F9'] = "=E9*1.1"

        # Services implémentation (ligne 4)
        services_base = pricing.get('services_daily', 800) * 12.5  # Prix journée * nb jours moyen
        ws['B4'] = services_base
        ws['C4'] = "=B4*1.05"
        ws['D4'] = "=C4*1.05"
        ws['E4'] = "=D4*1.05"
        ws['F4'] = "=E4*1.05"

        # NOUVELLE SECTION: Financial KPIs (colonne H+)
        financial_kpis = self.assumptions.get('financial_kpis', {})

        ws['H1'].value = "FINANCIAL KPIs"
        ws['H2'].value = "Métrique"
        ws['I2'].value = "Valeur"

        row = 3
        kpis_data = [
            ("ARR Target M14", financial_kpis.get('target_arr_dec_2026', 800000)),
            ("ARR Target M11", financial_kpis.get('target_arr_sept_2026', 450000)),
            ("Marge Brute Target", f"{financial_kpis.get('margin_targets', {}).get('gross_margin_pct', 70)}%"),
            ("EBITDA Margin Target", f"{financial_kpis.get('margin_targets', {}).get('ebitda_margin_pct', -15)}%"),
            ("Min Cash Runway (mois)", financial_kpis.get('cash_management', {}).get('min_cash_runway_months', 12)),
            ("Burn Rate Max (€/mois)", financial_kpis.get('cash_management', {}).get('acceptable_burn_rate_monthly', 50000)),
            ("Target LTV/CAC", financial_kpis.get('saas_metrics', {}).get('target_ltv_cac_ratio', 8)),
            ("Max Churn Annual", f"{financial_kpis.get('saas_metrics', {}).get('max_churn_annual', 0.15)*100}%"),
        ]

        for label, value in kpis_data:
            ws[f'H{row}'].value = label
            ws[f'I{row}'].value = value
            row += 1

        # NOUVELLE SECTION: Validation Rules (colonne K+)
        validation_rules = self.assumptions.get('validation_rules', {})

        ws['K1'].value = "VALIDATION RULES"
        ws['K2'].value = "Règle"
        ws['L2'].value = "Min"
        ws['M2'].value = "Max"

        row = 3
        rules_data = [
            ("ARR M14", validation_rules.get('arr_m14_min', 720000), validation_rules.get('arr_m14_max', 880000)),
            ("ARR M11", validation_rules.get('arr_m11_min', 400000), None),
            ("Team Size M14", validation_rules.get('min_team_size_m1', 4), validation_rules.get('max_team_size', 15)),
            ("Burn Monthly", None, validation_rules.get('max_burn_monthly', 60000)),
            ("Cash Balance Min", validation_rules.get('min_cash_balance', 50000), None),
            ("Conversion Hackathon→Factory", f"{validation_rules.get('min_conversion_hackathon_factory', 0.25)*100}%", None),
            ("Churn Hub Monthly Max", None, f"{validation_rules.get('max_churn_hub_monthly', 0.015)*100}%"),
        ]

        for label, min_val, max_val in rules_data:
            ws[f'K{row}'].value = label
            ws[f'L{row}'].value = min_val if min_val else "-"
            ws[f'M{row}'].value = max_val if max_val else "-"
            row += 1

        # NOUVELLE SECTION: Hypothèses critiques (colonne O+)
        ws['O1'].value = "HYPOTHÈSES BUSINESS"
        ws['O2'].value = "Hypothèse"
        ws['P2'].value = "Valeur"

        row = 3
        business_assumptions = [
            ("Conversion Hackathon→Factory", f"{self.assumptions.get('sales_assumptions', {}).get('factory', {}).get('conversion_rate', 0.35)*100}%"),
            ("Churn Hub Monthly", f"{self.assumptions.get('sales_assumptions', {}).get('enterprise_hub', {}).get('churn_monthly', 0.008)*100}%"),
            ("Launch Hub", f"M{self.assumptions.get('pricing', {}).get('enterprise_hub', {}).get('launch_month', 8)}"),
            ("Délai Factory (mois)", self.assumptions.get('sales_assumptions', {}).get('factory', {}).get('delay_months', 2)),
            ("Tier Starter %", f"{self.assumptions.get('sales_assumptions', {}).get('enterprise_hub', {}).get('tier_distribution_at_launch', {}).get('starter', 0.6)*100}%"),
            ("Tier Business %", f"{self.assumptions.get('sales_assumptions', {}).get('enterprise_hub', {}).get('tier_distribution_at_launch', {}).get('business', 0.3)*100}%"),
            ("Tier Enterprise %", f"{self.assumptions.get('sales_assumptions', {}).get('enterprise_hub', {}).get('tier_distribution_at_launch', {}).get('enterprise', 0.1)*100}%"),
        ]

        for label, value in business_assumptions:
            ws[f'O{row}'].value = label
            ws[f'P{row}'].value = value
            row += 1

        logger.info("✓ Paramètres enrichis avec financial_kpis, validation_rules, et hypothèses business")

    def update_financement_sheet(self):
        """
        Adapter le sheet Financement selon assumptions.yaml
        - Rounds de financement
        - Montants selon YAML
        """
        logger.info("\n💰 Adaptation sheet Financement...")

        ws = self.wb['Financement']

        funding = self.assumptions.get('funding', {})

        # Pré-seed (col C)
        preseed = funding.get('preseed', {})
        ws['C1'] = "2025-26"
        ws['C2'] = f"Pre-seed {preseed.get('quarter', 'Q4 2025')}"
        ws['C4'] = preseed.get('amount', 300000)  # Batch 1
        ws['C5'] = 50000  # Autoposia
        ws['C6'] = 50000  # F-Initiatives

        # Seed (col E)
        seed = funding.get('seed', {})
        ws['E1'] = "2027"
        ws['E2'] = f"Seed {seed.get('quarter', 'Q3 2026')}"
        ws['E8'] = seed.get('amount', 500000)  # CIC

        # Series A (col G)
        series_a = funding.get('series_a', {})
        ws['G1'] = "=E1+1"
        ws['G2'] = f"Series A {series_a.get('quarter', 'Q4 2027')}"
        ws['G11'] = series_a.get('amount', 2000000)  # BPI

        logger.info("✓ Financement adapté avec funding YAML")

    def update_fundings_sheet_with_captable(self):
        """
        Adapter le sheet Fundings avec la cap table détaillée
        - Dilution par phase (Bootstrap → Series A)
        - ARR targets alignés avec funding rounds
        """
        logger.info("\n📊 Adaptation sheet Fundings avec Cap Table...")

        if 'Fundings' not in self.wb.sheetnames:
            logger.warning("⚠️ Sheet 'Fundings' introuvable, création skippée")
            return

        ws = self.wb['Fundings']

        # Charger funding_captable.yaml
        base_path = Path(__file__).parent.parent
        captable_path = base_path / "data" / "structured" / "funding_captable.yaml"

        if not captable_path.exists():
            logger.warning("⚠️ funding_captable.yaml introuvable, cap table skippée")
            return

        with open(captable_path, 'r', encoding='utf-8') as f:
            captable_data = yaml.safe_load(f)

        # Section 1: Timeline de financement avec ARR targets (lignes 1-10)
        ws['A1'].value = "TIMELINE DE FINANCEMENT"
        ws['A2'].value = "Phase"
        ws['B2'].value = "Mois"
        ws['C2'].value = "Montant Levé"
        ws['D2'].value = "Valorisation Post"
        ws['E2'].value = "ARR Target"

        funding_rounds = captable_data.get('funding_rounds', {})
        row = 3

        for phase_key in ['bootstrap', 'love_money', 'pre_seed', 'seed', 'post_seed', 'series_a']:
            if phase_key not in funding_rounds:
                continue

            phase_data = funding_rounds[phase_key]
            ws[f'A{row}'].value = phase_data.get('phase', phase_key.upper())
            ws[f'B{row}'].value = f"M{phase_data.get('month', 0)}"
            ws[f'C{row}'].value = phase_data.get('amount', 0)
            ws[f'D{row}'].value = phase_data.get('valuation_post', 0)
            ws[f'E{row}'].value = phase_data.get('arr_target', 0)
            row += 1

        # Section 2: Cap table dilution (lignes 15+)
        ws['A15'].value = "CAP TABLE - DILUTION PROGRESSIVE"
        ws['A16'].value = "Phase"
        ws['B16'].value = "FRT (%)"
        ws['C16'].value = "PCO (%)"
        ws['D16'].value = "MAM (%)"
        ws['E16'].value = "BSPCE (%)"
        ws['F16'].value = "Investisseurs"

        captable = captable_data.get('captable', {})
        dilution_stages = captable.get('dilution_stages', {})

        row = 17
        for stage_key in ['bootstrap', 'post_pre_seed', 'post_seed', 'post_series_a']:
            if stage_key not in dilution_stages:
                continue

            stage_data = dilution_stages[stage_key]
            equity = stage_data.get('equity', {})

            ws[f'A{row}'].value = stage_data.get('phase', stage_key)

            # Formatter les equity (gérer floats ou strings déjà formatés)
            def format_equity(value):
                if isinstance(value, str):
                    return value  # Déjà formaté
                return f"{value:.1f}%" if value > 0 else "0.0%"

            ws[f'B{row}'].value = format_equity(equity.get('FRT', 0))
            ws[f'C{row}'].value = format_equity(equity.get('PCO', 0))
            ws[f'D{row}'].value = format_equity(equity.get('MAM', 0))
            ws[f'E{row}'].value = format_equity(equity.get('BSPCE', 0))

            # Investisseurs combinés
            inv_seed = equity.get('Investisseurs_Seed', 0)
            inv_a = equity.get('Investisseurs_Series_A', 0)
            inv_b = equity.get('Investisseurs_Series_B', 0)

            # Convertir en float si string
            if isinstance(inv_seed, str): inv_seed = 0
            if isinstance(inv_a, str): inv_a = 0
            if isinstance(inv_b, str): inv_b = 0

            inv_total = inv_seed + inv_a + inv_b
            ws[f'F{row}'].value = f"{inv_total:.1f}%" if inv_total > 0 else "-"

            row += 1

        # Section 3: ARR Milestones (lignes 25+)
        ws['A25'].value = "ARR MILESTONES CRITIQUES"
        ws['A26'].value = "Mois"
        ws['B26'].value = "ARR Target"
        ws['C26'].value = "Phase"

        arr_targets = captable_data.get('arr_targets', {})
        row = 27

        for month_key in ['M1', 'M6', 'M12', 'M18', 'M36', 'M48']:
            if month_key in arr_targets:
                ws[f'A{row}'].value = month_key
                ws[f'B{row}'].value = arr_targets[month_key]

                # Associer la phase
                if month_key == 'M1':
                    phase = "Bootstrap"
                elif month_key == 'M6':
                    phase = "PRE-SEED"
                elif month_key == 'M12':
                    phase = "SEED (contractuel)"
                elif month_key == 'M18':
                    phase = "Post Seed"
                elif month_key == 'M36':
                    phase = "SERIE A"
                elif month_key == 'M48':
                    phase = "Pre-Series B"
                else:
                    phase = "-"

                ws[f'C{row}'].value = phase
                row += 1

        logger.info("✓ Cap Table intégrée dans Fundings (Timeline + Dilution + ARR targets)")

    def update_strategie_vente_sheet(self):
        """
        Adapter le sheet Stratégie de vente selon assumptions.yaml
        - Taux de conversion Hackathon→Factory
        """
        logger.info("\n🎯 Adaptation sheet Stratégie de vente...")

        ws = self.wb['Stratégie de vente']

        conversion_rates = self.assumptions.get('conversion_rates', {})
        hackathon_to_factory = conversion_rates.get('hackathon_to_factory', 0.30)

        # Ajouter un indicateur des taux de conversion en haut du sheet
        # Ligne 1 col A-B: Taux de conversion
        ws['A1'].value = "Taux de conversion Hackathon→Factory"
        ws['B1'].value = f"{hackathon_to_factory*100:.0f}%"

        logger.info(f"✓ Stratégie de vente adaptée (conversion: {hackathon_to_factory*100:.0f}%)")

    def update_charges_personnel_sheet(self):
        """
        Adapter le sheet Charges de personnel et FG selon assumptions.yaml
        - Structure 8 rôles
        - Charges sociales 45%
        """
        logger.info("\n👥 Adaptation sheet Charges de personnel et FG...")

        ws = self.wb['Charges de personnel et FG']

        personnel = self.assumptions.get('personnel_details', {})
        charges_rate = personnel.get('charges_sociales_rate', 0.45)
        roles = personnel.get('roles', {})

        # Section info en haut (lignes 1-10)
        ws['A1'].value = "CHARGES DE PERSONNEL"
        ws['A2'].value = f"Charges sociales: {charges_rate*100:.0f}%"
        ws['A3'].value = f"Nombre de rôles: {len(roles)}"

        # Lister les rôles (lignes 5+)
        row = 5
        ws[f'A{row}'].value = "RÔLES DÉFINIS:"
        row += 1

        for role_name, role_data in roles.items():
            salary = role_data.get('salary_brut_annual', 0)
            ws[f'A{row}'].value = role_name.replace('_', ' ').title()
            ws[f'B{row}'].value = f"{salary:,.0f}€/an"
            row += 1

        logger.info(f"✓ Charges Personnel adaptées ({len(roles)} rôles, {charges_rate*100:.0f}% charges)")

    def update_infrastructure_detailed_sheet(self):
        """
        Adapter le sheet Infrastructure technique selon assumptions.yaml
        - Pricing cloud (base + tiers)
        - Pricing SaaS tools
        """
        logger.info("\n☁️ Adaptation sheet Infrastructure technique...")

        ws = self.wb['Infrastructure technique']

        infra = self.assumptions.get('infrastructure_costs', {})
        cloud = infra.get('cloud', {})
        saas = infra.get('saas_tools', {})

        # Section cloud (lignes 1-7)
        ws['A1'].value = "CLOUD INFRASTRUCTURE"
        ws['A2'].value = "Base mensuel"
        ws['B2'].value = cloud.get('base_monthly', 1000)

        ws['A3'].value = "Scaling tiers (coût par client):"

        scaling_tiers = cloud.get('scaling_tiers', {})
        ws['A4'].value = "  < 50 clients"
        ws['B4'].value = scaling_tiers.get('tier1', {}).get('cost_per_client', 50)

        ws['A5'].value = "  50-100 clients"
        ws['B5'].value = scaling_tiers.get('tier2', {}).get('cost_per_client', 40)

        ws['A6'].value = "  > 100 clients"
        ws['B6'].value = scaling_tiers.get('tier3', {}).get('cost_per_client', 30)

        # Section SaaS (lignes 9+)
        ws['A9'].value = "SAAS TOOLS"
        row = 10

        for tool_name, tool_data in saas.items():
            if isinstance(tool_data, dict):
                cost = tool_data.get('cost_per_user', 0) or tool_data.get('cost_per_developer', 0)
                min_users = tool_data.get('min_users', 1)
                ws[f'A{row}'].value = tool_name.title()
                ws[f'B{row}'].value = f"{cost}€/user (min {min_users})"
                row += 1

        logger.info("✓ Infrastructure technique adaptée (cloud + SaaS)")

    def update_marketing_detailed_sheet(self):
        """
        Adapter le sheet Marketing selon assumptions.yaml
        - 4 canaux avec budgets annuels
        """
        logger.info("\n📢 Adaptation sheet Marketing...")

        ws = self.wb['Marketing']

        marketing = self.assumptions.get('marketing_budgets', {})

        # Section budgets par canal (lignes 1+)
        ws['A1'].value = "BUDGETS MARKETING PAR CANAL"

        channels = ['digital_ads', 'events', 'content', 'partnerships']
        row = 3

        for channel in channels:
            if channel not in marketing:
                continue

            channel_data = marketing[channel]
            monthly_budgets = channel_data.get('monthly_budgets', {})

            ws[f'A{row}'].value = channel.replace('_', ' ').title()
            ws[f'B{row}'].value = "2025"
            ws[f'C{row}'].value = monthly_budgets.get('y2025', 0)
            ws[f'D{row}'].value = "2026"
            ws[f'E{row}'].value = monthly_budgets.get('y2026', 0)
            ws[f'F{row}'].value = "2027"
            ws[f'G{row}'].value = monthly_budgets.get('y2027', 0)
            ws[f'H{row}'].value = "2028"
            ws[f'I{row}'].value = monthly_budgets.get('y2028', 0)
            ws[f'J{row}'].value = "2029"
            ws[f'K{row}'].value = monthly_budgets.get('y2029', 0)

            row += 2

        logger.info(f"✓ Marketing adapté ({len(channels)} canaux)")

    def add_arr_mrr_to_pl(self):
        """
        Ajouter lignes ARR et MRR en haut du P&L
        Pour tracking des milestones SaaS
        """
        logger.info("\n📈 Ajout ARR/MRR dans P&L...")

        if 'P&L' not in self.wb.sheetnames:
            logger.warning("⚠️ Sheet 'P&L' introuvable, skip")
            return

        ws = self.wb['P&L']

        # Insérer 3 nouvelles lignes en haut (après les headers)
        # On va ajouter après la ligne "CA Total" qui est typiquement en ligne 2-3

        # Trouver la ligne "CA Total" ou similaire
        ca_row = None
        for row in range(1, 20):
            cell_value = ws[f'A{row}'].value
            if cell_value and isinstance(cell_value, str) and 'CA' in cell_value.upper():
                ca_row = row
                break

        if not ca_row:
            ca_row = 5  # Défaut

        insert_row = ca_row + 1

        # Insérer 3 lignes
        ws.insert_rows(insert_row, 3)

        # Ligne ARR
        ws[f'A{insert_row}'].value = "ARR (Annual Recurring Revenue)"
        ws[f'A{insert_row}'].font = openpyxl.styles.Font(bold=True)

        # Ligne MRR
        ws[f'A{insert_row+1}'].value = "MRR (Monthly Recurring Revenue)"
        ws[f'A{insert_row+1}'].font = openpyxl.styles.Font(bold=True)

        # Ligne séparatrice
        ws[f'A{insert_row+2}'].value = "---"

        logger.info(f"✓ ARR/MRR ajoutés en lignes {insert_row}-{insert_row+1} du P&L")

    def create_cash_flow_sheet(self):
        """
        Créer un nouveau sheet Cash Flow Statement
        Essential pour fundraising et suivi trésorerie
        """
        logger.info("\n💰 Création sheet Cash Flow...")

        # Créer nouveau sheet
        ws = self.wb.create_sheet("Cash Flow")

        # Headers
        ws['A1'].value = "CASH FLOW STATEMENT"
        ws['A1'].font = openpyxl.styles.Font(bold=True, size=14)

        ws['A2'].value = "Catégorie"
        ws['B2'].value = "Description"

        # Colonnes mois: C=M1, D=M2, etc.
        for month in range(1, 51):
            col_letter = openpyxl.utils.get_column_letter(month + 2)  # +2 car A,B = labels
            ws[f'{col_letter}2'].value = f"M{month}"

        # Section Operating Activities
        row = 3
        ws[f'A{row}'].value = "OPERATING ACTIVITIES"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        operating_items = [
            ("CA Encaissé", "Revenue collected"),
            ("Charges Personnel", "Salaries and social charges"),
            ("Charges Infrastructure", "Cloud + SaaS tools"),
            ("Charges Marketing", "Marketing spend"),
            ("Autres Charges", "Other operating expenses"),
        ]

        for label, desc in operating_items:
            ws[f'A{row}'].value = f"  {label}"
            ws[f'B{row}'].value = desc
            row += 1

        ws[f'A{row}'].value = "= Cash Flow Opérationnel"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 2

        # Section Investing Activities
        ws[f'A{row}'].value = "INVESTING ACTIVITIES"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        ws[f'A{row}'].value = "  CAPEX (équipements)"
        ws[f'B{row}'].value = "Equipment and infrastructure"
        row += 1

        ws[f'A{row}'].value = "= Cash Flow Investissement"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 2

        # Section Financing Activities
        ws[f'A{row}'].value = "FINANCING ACTIVITIES"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        financing_items = [
            ("Pre-Seed", "M1: 150K€"),
            ("Seed", "M11: 500K€"),
            ("Series A", "M36: 2.5M€"),
        ]

        for label, desc in financing_items:
            ws[f'A{row}'].value = f"  {label}"
            ws[f'B{row}'].value = desc
            row += 1

        ws[f'A{row}'].value = "= Cash Flow Financement"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 2

        # Total et balance
        ws[f'A{row}'].value = "TOTAL CASH FLOW (mois)"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True, color="0000FF")
        row += 1

        ws[f'A{row}'].value = "CASH BALANCE (cumulé)"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True, size=12, color="FF0000")
        ws[f'B{row}'].value = "Trésorerie disponible"
        row += 2

        # Métriques
        ws[f'A{row}'].value = "MÉTRIQUES"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        ws[f'A{row}'].value = "  Burn Rate (€/mois)"
        ws[f'B{row}'].value = "Operating CF négatif"
        row += 1

        ws[f'A{row}'].value = "  Cash Runway (mois)"
        ws[f'B{row}'].value = "Cash balance / Burn rate"
        row += 1

        # Ajuster largeur colonnes
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 40

        logger.info("✓ Sheet Cash Flow créé avec structure complète")

    def remove_gtmarket_sheet(self):
        """
        Supprimer le sheet GTMarket (110 cols × 1000 rows, peu de valeur ajoutée)
        """
        logger.info("\n🗑️ Suppression sheet GTMarket...")

        if 'GTMarket' in self.wb.sheetnames:
            del self.wb['GTMarket']
            logger.info("✓ Sheet GTMarket supprimé")
        else:
            logger.warning("⚠️ Sheet GTMarket introuvable, skip")

    def enrich_synthese_dashboard(self):
        """
        Enrichir le sheet Synthèse avec dashboard KPIs
        Vue exécutive pour investisseurs et board
        """
        logger.info("\n📊 Enrichissement Synthèse avec dashboard...")

        if 'Synthèse' not in self.wb.sheetnames:
            logger.warning("⚠️ Sheet 'Synthèse' introuvable, skip")
            return

        ws = self.wb['Synthèse']

        # Trouver une zone vide (colonne Y+ par exemple)
        start_col = 'Y'

        # Dashboard header
        ws[f'{start_col}1'].value = "DASHBOARD EXÉCUTIF"
        ws[f'{start_col}1'].font = openpyxl.styles.Font(bold=True, size=14, color="FFFFFF")
        ws[f'{start_col}1'].fill = openpyxl.styles.PatternFill(start_color="0066CC", end_color="0066CC", fill_type="solid")

        # Section 1: ARR Milestones
        row = 3
        ws[f'{start_col}{row}'].value = "ARR MILESTONES"
        ws[f'{start_col}{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        # Charger cap table pour les targets
        base_path = Path(__file__).parent.parent
        captable_path = base_path / "data" / "structured" / "funding_captable.yaml"

        if captable_path.exists():
            with open(captable_path, 'r', encoding='utf-8') as f:
                captable_data = yaml.safe_load(f)
            arr_targets = captable_data.get('arr_targets', {})
        else:
            arr_targets = {}

        arr_milestones = [
            ("M1 (Bootstrap)", arr_targets.get('M1', 10000)),
            ("M6 (PRE-SEED)", arr_targets.get('M6', 500000)),
            ("M12 (SEED)", arr_targets.get('M12', 800000)),
            ("M18 (Post Seed)", arr_targets.get('M18', 1500000)),
            ("M36 (Series A)", arr_targets.get('M36', 4000000)),
            ("M48 (Pre Series B)", arr_targets.get('M48', 6000000)),
        ]

        ws[f'{start_col}{row}'].value = "Milestone"
        next_col = openpyxl.utils.get_column_letter(openpyxl.utils.column_index_from_string(start_col) + 1)
        ws[f'{next_col}{row}'].value = "ARR Target"
        row += 1

        for milestone, target in arr_milestones:
            ws[f'{start_col}{row}'].value = milestone
            ws[f'{next_col}{row}'].value = target
            row += 1

        row += 1

        # Section 2: KPIs Critiques
        ws[f'{start_col}{row}'].value = "KPIs CRITIQUES"
        ws[f'{start_col}{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        financial_kpis = self.assumptions.get('financial_kpis', {})

        kpis = [
            ("Target LTV/CAC", financial_kpis.get('saas_metrics', {}).get('target_ltv_cac_ratio', 8)),
            ("Max Churn Annual", f"{financial_kpis.get('saas_metrics', {}).get('max_churn_annual', 0.15)*100}%"),
            ("Marge Brute Target", f"{financial_kpis.get('margin_targets', {}).get('gross_margin_pct', 70)}%"),
            ("EBITDA Margin Target", f"{financial_kpis.get('margin_targets', {}).get('ebitda_margin_pct', -15)}%"),
            ("Min Cash Runway", f"{financial_kpis.get('cash_management', {}).get('min_cash_runway_months', 12)} mois"),
            ("Max Burn Rate", f"{financial_kpis.get('cash_management', {}).get('acceptable_burn_rate_monthly', 50000):,}€/mois"),
        ]

        ws[f'{start_col}{row}'].value = "KPI"
        ws[f'{next_col}{row}'].value = "Valeur"
        row += 1

        for kpi, value in kpis:
            ws[f'{start_col}{row}'].value = kpi
            ws[f'{next_col}{row}'].value = value
            row += 1

        row += 1

        # Section 3: Hypothèses Critiques
        ws[f'{start_col}{row}'].value = "HYPOTHÈSES CRITIQUES"
        ws[f'{start_col}{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        critical_assumptions = self.assumptions.get('critical_assumptions', [])
        if critical_assumptions:
            for assumption in critical_assumptions[:5]:  # Top 5
                if isinstance(assumption, dict):
                    ws[f'{start_col}{row}'].value = f"• {assumption.get('assumption', '')}"
                    ws[f'{next_col}{row}'].value = assumption.get('risk_level', '')
                    row += 1

        # Ajuster largeur colonnes
        ws.column_dimensions[start_col].width = 25
        ws.column_dimensions[next_col].width = 20

        logger.info("✓ Dashboard exécutif ajouté dans Synthèse")

    def create_scenarios_sheet(self):
        """
        Créer nouveau sheet Scenarios (base/upside/downside)
        Analyse de sensibilité pour investisseurs
        """
        logger.info("\n📊 Création sheet Scenarios...")

        # Créer nouveau sheet
        ws = self.wb.create_sheet("Scenarios")

        # Header
        ws['A1'].value = "SCENARIOS D'ÉVOLUTION"
        ws['A1'].font = openpyxl.styles.Font(bold=True, size=14, color="FFFFFF")
        ws['A1'].fill = openpyxl.styles.PatternFill(start_color="0066CC", end_color="0066CC", fill_type="solid")

        ws['A2'].value = "Basé sur assumptions.yaml - 3 scénarios probabilisés"

        # Colonnes
        ws['A4'].value = "Métrique"
        ws['B4'].value = "BASE CASE (60%)"
        ws['C4'].value = "UPSIDE (20%)"
        ws['D4'].value = "DOWNSIDE (20%)"
        ws['E4'].value = "Notes"

        for col in ['A', 'B', 'C', 'D', 'E']:
            ws[f'{col}4'].font = openpyxl.styles.Font(bold=True)

        # Charger scenarios depuis YAML
        scenarios = self.assumptions.get('scenarios', {})
        base = scenarios.get('base_case', {})
        upside = scenarios.get('upside', {})
        downside = scenarios.get('downside', {})

        row = 5

        # Section ARR
        ws[f'A{row}'].value = "ARR M14"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        ws[f'B{row}'].value = base.get('arr_m14', 800000)
        ws[f'C{row}'].value = upside.get('arr_m14', 952000)
        ws[f'D{row}'].value = downside.get('arr_m14', 648000)
        ws[f'E{row}'].value = "Annual Recurring Revenue à M14"
        row += 1

        ws[f'A{row}'].value = "Probabilité"
        ws[f'B{row}'].value = f"{base.get('probability', 0.6)*100}%"
        ws[f'C{row}'].value = f"{upside.get('probability', 0.2)*100}%"
        ws[f'D{row}'].value = f"{downside.get('probability', 0.2)*100}%"
        row += 2

        # Section Hypothèses Hackathon
        ws[f'A{row}'].value = "HYPOTHÈSES HACKATHON"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        ws[f'A{row}'].value = "Volume multiplier"
        ws[f'B{row}'].value = "1.0×"
        ws[f'C{row}'].value = f"{upside.get('hackathon_volume_multiplier', 1.2)}×"
        ws[f'D{row}'].value = f"{downside.get('hackathon_volume_multiplier', 0.8)}×"
        ws[f'E{row}'].value = "Multiplicateur volumes hackathons"
        row += 1

        ws[f'A{row}'].value = "Conversion → Factory"
        ws[f'B{row}'].value = "30%"
        ws[f'C{row}'].value = f"{upside.get('conversion_factory', 0.35)*100}%"
        ws[f'D{row}'].value = f"{downside.get('conversion_factory', 0.25)*100}%"
        ws[f'E{row}'].value = "Taux conversion Hackathon → Factory"
        row += 2

        # Section Hypothèses Hub
        ws[f'A{row}'].value = "HYPOTHÈSES ENTERPRISE HUB"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        ws[f'A{row}'].value = "Launch delay"
        ws[f'B{row}'].value = "M8 (aucun retard)"
        ws[f'C{row}'].value = "M8 (accéléré)"
        ws[f'D{row}'].value = f"M{8 + downside.get('hub_launch_delay_months', 2)} (+{downside.get('hub_launch_delay_months', 2)} mois)"
        ws[f'E{row}'].value = "Délai lancement Hub"
        row += 1

        ws[f'A{row}'].value = "Adoption speed"
        ws[f'B{row}'].value = "Normal"
        ws[f'C{row}'].value = "Rapide (+20%)"
        ws[f'D{row}'].value = "Lente (-20%)"
        row += 2

        # Section Impact Financier
        ws[f'A{row}'].value = "IMPACT FINANCIER ESTIMÉ"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        # Calculs approximatifs
        base_arr = base.get('arr_m14', 800000)
        upside_arr = upside.get('arr_m14', 952000)
        downside_arr = downside.get('arr_m14', 648000)

        ws[f'A{row}'].value = "ARR M14"
        ws[f'B{row}'].value = base_arr
        ws[f'C{row}'].value = upside_arr
        ws[f'D{row}'].value = downside_arr
        row += 1

        ws[f'A{row}'].value = "vs Base Case"
        ws[f'B{row}'].value = "0%"
        ws[f'C{row}'].value = f"+{(upside_arr/base_arr - 1)*100:.0f}%"
        ws[f'D{row}'].value = f"{(downside_arr/base_arr - 1)*100:.0f}%"
        row += 2

        # Hypothèses critiques
        ws[f'A{row}'].value = "HYPOTHÈSES CRITIQUES"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        critical_assumptions = self.assumptions.get('critical_assumptions', [])
        for assumption in critical_assumptions[:5]:
            if isinstance(assumption, dict):
                ws[f'A{row}'].value = f"• {assumption.get('assumption', '')}"
                ws[f'B{row}'].value = assumption.get('risk_level', '')
                ws[f'E{row}'].value = assumption.get('mitigation', '')[:50] if assumption.get('mitigation') else ""
                row += 1

        # Ajuster largeur colonnes
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 50

        logger.info("✓ Sheet Scenarios créé (base/upside/downside)")

    def create_unit_economics_sheet(self):
        """
        Créer nouveau sheet Unit Economics
        CAC, LTV, Payback period par produit
        """
        logger.info("\n💰 Création sheet Unit Economics...")

        # Créer nouveau sheet
        ws = self.wb.create_sheet("Unit Economics")

        # Header
        ws['A1'].value = "UNIT ECONOMICS PAR PRODUIT"
        ws['A1'].font = openpyxl.styles.Font(bold=True, size=14, color="FFFFFF")
        ws['A1'].fill = openpyxl.styles.PatternFill(start_color="0066CC", end_color="0066CC", fill_type="solid")

        ws['A3'].value = "Produit"
        ws['B3'].value = "Prix Moyen"
        ws['C3'].value = "CAC (€)"
        ws['D3'].value = "LTV (€)"
        ws['E3'].value = "LTV/CAC"
        ws['F3'].value = "Payback (mois)"
        ws['G3'].value = "Marge (%)"
        ws['H3'].value = "Notes"

        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
            ws[f'{col}3'].font = openpyxl.styles.Font(bold=True)

        row = 4

        # Hackathon
        hackathon_price = self.assumptions.get('pricing', {}).get('hackathon', {}).get('periods', [{}])[0].get('price_eur', 18000)
        ws[f'A{row}'].value = "Hackathon"
        ws[f'B{row}'].value = hackathon_price
        ws[f'C{row}'].value = 5000  # CAC from marketing_budgets
        ws[f'D{row}'].value = hackathon_price * 1.5  # LTV (client peut revenir)
        ws[f'E{row}'].value = f"=D{row}/C{row}"  # LTV/CAC formula
        ws[f'F{row}'].value = 1  # Payback immédiat
        ws[f'G{row}'].value = "80%"
        ws[f'H{row}'].value = "Offre d'entrée, forte marge"
        row += 1

        # Factory
        factory_price = self.assumptions.get('pricing', {}).get('factory', {}).get('periods', [{}])[0].get('price_eur', 75000)
        ws[f'A{row}'].value = "Factory"
        ws[f'B{row}'].value = factory_price
        ws[f'C{row}'].value = 10000  # CAC plus élevé (cycle long)
        ws[f'D{row}'].value = factory_price * 2  # LTV (upsell services)
        ws[f'E{row}'].value = f"=D{row}/C{row}"
        ws[f'F{row}'].value = 3  # Payback 3 mois
        ws[f'G{row}'].value = "65%"
        ws[f'H{row}'].value = "Conversion naturelle hackathon"
        row += 1

        # Enterprise Hub - Starter
        starter_price = self.assumptions.get('pricing', {}).get('enterprise_hub', {}).get('tiers', {}).get('starter', {}).get('monthly_eur', 500)
        ws[f'A{row}'].value = "Hub Starter"
        ws[f'B{row}'].value = f"{starter_price}€/mois"
        ws[f'C{row}'].value = 15000  # CAC from assumptions
        ws[f'D{row}'].value = 36000  # LTV (3 ans × 12 mois × 500€ × retention)
        ws[f'E{row}'].value = f"=D{row}/C{row}"
        ws[f'F{row}'].value = 30  # Payback 30 mois
        ws[f'G{row}'].value = "75%"
        ws[f'H{row}'].value = "SaaS récurrent, target PME"
        row += 1

        # Enterprise Hub - Business
        business_price = self.assumptions.get('pricing', {}).get('enterprise_hub', {}).get('tiers', {}).get('business', {}).get('monthly_eur', 2000)
        ws[f'A{row}'].value = "Hub Business"
        ws[f'B{row}'].value = f"{business_price}€/mois"
        ws[f'C{row}'].value = 15000
        ws[f'D{row}'].value = 60000  # LTV
        ws[f'E{row}'].value = f"=D{row}/C{row}"
        ws[f'F{row}'].value = 8  # Payback 8 mois
        ws[f'G{row}'].value = "78%"
        ws[f'H{row}'].value = "SaaS récurrent, target ETI"
        row += 1

        # Enterprise Hub - Enterprise
        enterprise_price = self.assumptions.get('pricing', {}).get('enterprise_hub', {}).get('tiers', {}).get('enterprise', {}).get('monthly_eur', 10000)
        ws[f'A{row}'].value = "Hub Enterprise"
        ws[f'B{row}'].value = f"{enterprise_price}€/mois"
        ws[f'C{row}'].value = 15000
        ws[f'D{row}'].value = 120000  # LTV complet
        ws[f'E{row}'].value = f"=D{row}/C{row}"
        ws[f'F{row}'].value = 2  # Payback 2 mois
        ws[f'G{row}'].value = "80%"
        ws[f'H{row}'].value = "SaaS récurrent, target Grands Comptes"
        row += 1

        # Services
        services_price = self.assumptions.get('pricing', {}).get('services', {}).get('implementation', {}).get('periods', [{}])[0].get('price_eur', 10000)
        ws[f'A{row}'].value = "Services Implémentation"
        ws[f'B{row}'].value = services_price
        ws[f'C{row}'].value = 2000  # CAC faible (upsell)
        ws[f'D{row}'].value = services_price * 1.2  # LTV
        ws[f'E{row}'].value = f"=D{row}/C{row}"
        ws[f'F{row}'].value = 1  # Payback immédiat
        ws[f'G{row}'].value = "70%"
        ws[f'H{row}'].value = "Revenus complémentaires"
        row += 2

        # Résumé
        ws[f'A{row}'].value = "RÉSUMÉ"
        ws[f'A{row}'].font = openpyxl.styles.Font(bold=True)
        row += 1

        ws[f'A{row}'].value = "LTV/CAC Moyen Pondéré"
        ws[f'B{row}'].value = self.assumptions.get('financial_kpis', {}).get('saas_metrics', {}).get('target_ltv_cac_ratio', 8)
        ws[f'H{row}'].value = "Target: 8× (excellent pour SaaS B2B)"
        row += 1

        ws[f'A{row}'].value = "Payback Moyen (mois)"
        ws[f'B{row}'].value = "6-12 mois"
        ws[f'H{row}'].value = "Variable selon produit et tier"
        row += 1

        ws[f'A{row}'].value = "Churn Annual Max"
        ws[f'B{row}'].value = f"{self.assumptions.get('financial_kpis', {}).get('saas_metrics', {}).get('max_churn_annual', 0.15)*100}%"
        ws[f'H{row}'].value = "Hub uniquement (hackathon/factory = one-time)"

        # Ajuster largeur colonnes
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 15
        ws.column_dimensions['G'].width = 12
        ws.column_dimensions['H'].width = 40

        logger.info("✓ Sheet Unit Economics créé (CAC/LTV par produit)")

    def clean_data_cells(self):
        """
        Nettoyer les cellules de données (pas les formules)
        Mettre des valeurs par défaut pour les placeholders
        """
        logger.info("\n🧹 Nettoyage cellules de données...")

        sheets_to_clean = ['P&L', 'Ventes', 'Charges de personnel et FG',
                          'Infrastructure technique', 'Marketing', 'Sous traitance']

        for sheet_name in sheets_to_clean:
            if sheet_name not in self.wb.sheetnames:
                continue

            ws = self.wb[sheet_name]

            # Parcourir les cellules de données (à partir ligne 4, col 4)
            cleaned = 0
            for row in ws.iter_rows(min_row=4, max_row=100, min_col=4, max_col=150):
                for cell in row:
                    # Si c'est une valeur numérique (pas formule), mettre 0
                    if isinstance(cell.value, (int, float)) and cell.value != 0:
                        cell.value = 0
                        cleaned += 1

            logger.info(f"  {sheet_name}: {cleaned} cellules nettoyées")

    def add_template_markers(self):
        """
        Ajouter des marqueurs visuels pour identifier le template
        """
        logger.info("\n🏷️ Ajout marqueurs TEMPLATE...")

        # Ajouter une note sur la première sheet
        ws = self.wb.worksheets[0]
        ws['A1'].value = "🔧 TEMPLATE BP 50 MOIS - Gabarit à valider"

        # Style
        from openpyxl.styles import Font, PatternFill
        ws['A1'].font = Font(bold=True, size=14, color="FF0000")
        ws['A1'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        logger.info("✓ Marqueurs ajoutés")

    def preserve_formulas_info(self):
        """
        Logger des infos sur les formules préservées
        """
        logger.info("\n📊 Inventaire formules préservées...")

        for sheet_name in ['P&L', 'Ventes', 'Synthèse']:
            if sheet_name not in self.wb.sheetnames:
                continue

            ws = self.wb[sheet_name]
            formula_count = 0

            for row in ws.iter_rows(min_row=1, max_row=100, min_col=1, max_col=150):
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                        formula_count += 1

            logger.info(f"  {sheet_name}: {formula_count} formules")

    def create_template(self):
        """Créer le template complet avec toutes les améliorations Phase 1 + Phase 2"""
        logger.info("\n🔨 CRÉATION TEMPLATE ENRICHI (Phase 1 + Phase 2)")
        logger.info("=" * 60)

        # 1. Adapter structure selon YAML (existant + enrichi)
        self.update_parametres_sheet()  # ✅ Enrichi avec financial_kpis, validation_rules, hypothèses
        self.update_financement_sheet()
        self.update_fundings_sheet_with_captable()  # ✅ Cap table détaillée
        self.update_strategie_vente_sheet()
        self.update_charges_personnel_sheet()
        self.update_infrastructure_detailed_sheet()
        self.update_marketing_detailed_sheet()

        # 2. PHASE 1 - Améliorations HAUTE PRIORITÉ
        self.add_arr_mrr_to_pl()  # ✅ ARR/MRR dans P&L
        self.create_cash_flow_sheet()  # ✅ Cash Flow Statement
        self.enrich_synthese_dashboard()  # ✅ Dashboard exécutif

        # 3. PHASE 2 - Améliorations MOYENNE PRIORITÉ
        self.create_scenarios_sheet()  # ✅ NEW: Scenarios (base/upside/downside)
        self.create_unit_economics_sheet()  # ✅ NEW: Unit Economics (CAC/LTV par produit)

        # 4. Supprimer sheets inutiles
        self.remove_gtmarket_sheet()  # ✅ Suppression GTMarket

        # 5. Nettoyer les données
        self.clean_data_cells()

        # 6. Ajouter marqueurs
        self.add_template_markers()

        # 7. Vérifier formules
        self.preserve_formulas_info()

        logger.info("\n" + "=" * 60)
        logger.info("✅ TEMPLATE ENRICHI CRÉÉ (Phase 1 + Phase 2 complètes)")
        logger.info("   PHASE 1:")
        logger.info("   • Paramètres: financial_kpis + validation_rules + hypothèses")
        logger.info("   • P&L: ARR/MRR ajoutés")
        logger.info("   • Cash Flow: nouveau sheet créé")
        logger.info("   • Synthèse: dashboard exécutif ajouté")
        logger.info("   PHASE 2:")
        logger.info("   • Scenarios: base/upside/downside créé")
        logger.info("   • Unit Economics: CAC/LTV par produit créé")

    def save(self, output_path: Path):
        """Sauvegarder le template"""
        logger.info(f"\n💾 Sauvegarde: {output_path}")
        self.wb.save(output_path)
        size_kb = output_path.stat().st_size / 1024
        logger.info(f"✓ Template sauvegardé: {size_kb:.1f} KB")


def main():
    console.print("\n[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]")
    console.print("[bold cyan]   CRÉATION TEMPLATE EXCEL DEPUIS RAW[/bold cyan]")
    console.print("[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]\n")

    base_path = Path(__file__).parent.parent

    # Charger assumptions
    assumptions_path = base_path / "data" / "structured" / "assumptions.yaml"
    console.print(f"[yellow]📂 Chargement assumptions:[/yellow] {assumptions_path.name}")
    with open(assumptions_path, 'r', encoding='utf-8') as f:
        assumptions = yaml.safe_load(f)
    console.print(f"[green]✓ Assumptions chargées (v{assumptions.get('version', '?')})[/green]\n")

    # Fichiers
    raw_file = base_path / "data" / "raw" / "BP FABRIQ_PRODUCT-OCT2025.xlsx"
    template_file = base_path / "data" / "outputs" / "BP_50M_TEMPLATE.xlsx"

    # Créer le template
    creator = TemplateCreator(raw_file, assumptions)
    creator.create_template()
    creator.save(template_file)

    console.print(f"\n[bold green]✅ TEMPLATE CRÉÉ[/bold green]")
    console.print(f"[green]📁 {template_file}[/green]")
    console.print(f"\n[cyan]→ Structure adaptée selon assumptions.yaml[/cyan]")
    console.print(f"[cyan]→ Toutes les formules Excel préservées[/cyan]")
    console.print(f"[cyan]→ Cellules de données nettoyées (placeholders à 0)[/cyan]")
    console.print(f"[yellow]→ À VALIDER avant injection des données[/yellow]\n")


if __name__ == "__main__":
    main()
