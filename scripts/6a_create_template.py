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

        logger.info("✓ Paramètres adaptés avec pricing YAML")

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
        """Créer le template complet"""
        logger.info("\n🔨 CRÉATION TEMPLATE")
        logger.info("=" * 60)

        # 1. Adapter structure selon YAML
        self.update_parametres_sheet()
        self.update_financement_sheet()
        self.update_fundings_sheet_with_captable()  # NEW: Cap table détaillée
        self.update_strategie_vente_sheet()
        self.update_charges_personnel_sheet()
        self.update_infrastructure_detailed_sheet()
        self.update_marketing_detailed_sheet()

        # 2. Supprimer sheets inutiles
        self.remove_gtmarket_sheet()  # NEW: Suppression GTMarket

        # 3. Nettoyer les données
        self.clean_data_cells()

        # 4. Ajouter marqueurs
        self.add_template_markers()

        # 5. Vérifier formules
        self.preserve_formulas_info()

        logger.info("\n" + "=" * 60)
        logger.info("✅ TEMPLATE CRÉÉ")

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
