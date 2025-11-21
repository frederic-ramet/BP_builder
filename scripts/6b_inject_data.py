#!/usr/bin/env python3
"""
Injecter les projections 50M dans le TEMPLATE Excel
pour générer le fichier FINAL
"""

import openpyxl
from pathlib import Path
import json
from rich.console import Console
from rich.progress import track
import logging

console = Console()
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger(__name__)


class DataInjector:
    """Injecter les données dans le template"""

    def __init__(self, template_path: Path, projections: list):
        self.template_path = template_path
        self.projections = projections

        logger.info(f"📂 Chargement TEMPLATE: {template_path.name}")
        self.wb = openpyxl.load_workbook(template_path)
        logger.info(f"✓ {len(self.wb.sheetnames)} sheets chargés")

        # Mapper les colonnes
        self.setup_month_mapping()

    def setup_month_mapping(self):
        """Mapper les 50 mois aux colonnes Excel"""
        self.month_to_col = {}

        pl_sheet = self.wb['P&L']

        current_month = 1
        for col_idx in range(4, pl_sheet.max_column + 1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)

            year_cell = pl_sheet.cell(1, col_idx).value
            month_cell = pl_sheet.cell(2, col_idx).value

            # Skip colonnes totaux annuels
            if isinstance(year_cell, (int, float)) and month_cell is None:
                continue

            # C'est un mois
            if isinstance(month_cell, (int, float)) or (isinstance(month_cell, str) and month_cell.startswith('=')):
                self.month_to_col[current_month] = col_letter
                current_month += 1

            if current_month > 50:
                break

        logger.info(f"✓ Mapping: {len(self.month_to_col)} mois (M1→{self.month_to_col.get(1)}, M50→{self.month_to_col.get(50)})")

    def inject_pl_data(self):
        """Injecter données P&L"""
        logger.info("\n📊 Injection P&L...")

        ws = self.wb['P&L']

        row_map = {
            'ca_total': 2,
            'hackathons': 3,
            'factory': 4,
            'hub_mrr': 5,
            'services': 6,
            'sous_traitance': 10,
            'infrastructure': 11,
            'charges_personnel_ops': 12,
            'marketing': 18,
            'charges_personnel_fonc': 19,
            'frais_generaux': 20,
        }

        injected = 0
        for month in range(1, 51):
            if month not in self.month_to_col:
                continue

            col = self.month_to_col[month]
            proj = self.projections[month - 1]

            # Injecter seulement si pas de formule
            data = {
                'ca_total': proj['revenue']['total'],
                'hackathons': proj['revenue']['hackathon']['revenue'],
                'factory': proj['revenue']['factory']['revenue'],
                'hub_mrr': proj['revenue']['enterprise_hub']['mrr'],
                'services': proj['revenue']['services']['revenue'],
                'sous_traitance': proj['costs']['personnel'].get('freelance', 0),
                'infrastructure': proj['costs']['infrastructure']['total'],
                'charges_personnel_ops': proj['costs']['personnel']['total'],
                'marketing': proj['costs']['marketing']['total'],
                'charges_personnel_fonc': proj['costs']['personnel']['total'],
                'frais_generaux': proj['costs'].get('admin', 0),
            }

            for key, value in data.items():
                if key not in row_map:
                    continue

                cell = ws[f'{col}{row_map[key]}']
                if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                    cell.value = value
                    injected += 1

        logger.info(f"✓ P&L: {injected} cellules injectées")

    def inject_ventes_data(self):
        """Injecter données Ventes"""
        logger.info("\n💼 Injection Ventes...")

        ws = self.wb['Ventes']

        row_map = {
            'nb_hackathons': 3,
            'ca_hackathons': 4,
            'nb_factory': 6,
            'ca_factory': 9,
            'nb_clients_hub': 11,
            'mrr_hub': 13,
            'arr': 15,
        }

        injected = 0
        for month in range(1, 51):
            if month not in self.month_to_col:
                continue

            col = self.month_to_col[month]
            proj = self.projections[month - 1]

            data = {
                'nb_hackathons': proj['revenue']['hackathon'].get('volume', 0),
                'ca_hackathons': proj['revenue']['hackathon']['revenue'],
                'nb_factory': proj['revenue']['factory'].get('volume', 0),
                'ca_factory': proj['revenue']['factory']['revenue'],
                'nb_clients_hub': proj['revenue']['enterprise_hub']['customers']['total'],
                'mrr_hub': proj['revenue']['enterprise_hub']['mrr'],
                'arr': proj['revenue']['enterprise_hub']['arr'],
            }

            for key, value in data.items():
                if key not in row_map:
                    continue

                cell = ws[f'{col}{row_map[key]}']
                if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                    cell.value = value
                    injected += 1

        logger.info(f"✓ Ventes: {injected} cellules injectées")

    def inject_personnel_data(self):
        """Injecter données Personnel"""
        logger.info("\n👥 Injection Charges Personnel...")

        ws = self.wb['Charges de personnel et FG']

        injected = 0
        for month in range(1, 51):
            if month not in self.month_to_col:
                continue

            col = self.month_to_col[month]
            proj = self.projections[month - 1]

            # Total personnel (ligne approximative 5)
            cell = ws[f'{col}5']
            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                cell.value = proj['costs']['personnel']['total']
                injected += 1

        logger.info(f"✓ Personnel: {injected} cellules injectées")

    def inject_infrastructure_data(self):
        """Injecter données Infrastructure"""
        logger.info("\n☁️ Injection Infrastructure...")

        ws = self.wb['Infrastructure technique']

        injected = 0
        for month in range(1, 51):
            if month not in self.month_to_col:
                continue

            col = self.month_to_col[month]
            proj = self.projections[month - 1]

            # Cloud (ligne 3)
            cell = ws[f'{col}3']
            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                cell.value = proj['costs']['infrastructure'].get('cloud', 0)
                injected += 1

            # SaaS (ligne 5)
            cell = ws[f'{col}5']
            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                cell.value = proj['costs']['infrastructure'].get('saas_tools', 0)
                injected += 1

        logger.info(f"✓ Infrastructure: {injected} cellules injectées")

    def inject_marketing_data(self):
        """Injecter données Marketing"""
        logger.info("\n📢 Injection Marketing...")

        ws = self.wb['Marketing']

        injected = 0
        for month in range(1, 51):
            if month not in self.month_to_col:
                continue

            col = self.month_to_col[month]
            proj = self.projections[month - 1]

            # Total marketing (ligne 3)
            cell = ws[f'{col}3']
            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                cell.value = proj['costs']['marketing']['total']
                injected += 1

        logger.info(f"✓ Marketing: {injected} cellules injectées")

    def inject_sous_traitance_data(self):
        """Injecter données Sous-traitance"""
        logger.info("\n🔧 Injection Sous-traitance...")

        ws = self.wb['Sous traitance']

        injected = 0
        for month in range(1, 51):
            if month not in self.month_to_col:
                continue

            col = self.month_to_col[month]
            proj = self.projections[month - 1]

            # Freelance (ligne 3)
            cell = ws[f'{col}3']
            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                cell.value = proj['costs']['personnel'].get('freelance', 0)
                injected += 1

        logger.info(f"✓ Sous-traitance: {injected} cellules injectées")

    def remove_template_marker(self):
        """Retirer le marqueur TEMPLATE"""
        logger.info("\n🏷️ Retrait marqueur TEMPLATE...")

        ws = self.wb.worksheets[0]
        if ws['A1'].value and 'TEMPLATE' in str(ws['A1'].value):
            ws['A1'].value = "Business Plan GenieFactory - 50 Mois (Nov 2025 - Dec 2029)"
            from openpyxl.styles import Font, PatternFill
            ws['A1'].font = Font(bold=True, size=14, color="000000")
            ws['A1'].fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            logger.info("✓ Marqueur retiré")

    def inject_all(self):
        """Injecter toutes les données"""
        logger.info("\n🔨 INJECTION DONNÉES")
        logger.info("=" * 60)

        self.inject_pl_data()
        self.inject_ventes_data()
        self.inject_personnel_data()
        self.inject_infrastructure_data()
        self.inject_marketing_data()
        self.inject_sous_traitance_data()
        self.remove_template_marker()

        logger.info("\n" + "=" * 60)
        logger.info("✅ INJECTION TERMINÉE")

    def save(self, output_path: Path):
        """Sauvegarder le fichier final"""
        logger.info(f"\n💾 Sauvegarde: {output_path}")
        self.wb.save(output_path)
        size_kb = output_path.stat().st_size / 1024
        logger.info(f"✓ Fichier FINAL sauvegardé: {size_kb:.1f} KB")


def main():
    console.print("\n[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]")
    console.print("[bold cyan]   INJECTION DONNÉES DANS TEMPLATE[/bold cyan]")
    console.print("[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]\n")

    base_path = Path(__file__).parent.parent

    # Charger projections
    projections_path = base_path / "data" / "structured" / "projections_50m.json"
    console.print(f"[yellow]📂 Chargement projections:[/yellow] {projections_path.name}")
    with open(projections_path, 'r', encoding='utf-8') as f:
        projections = json.load(f)
    console.print(f"[green]✓ {len(projections)} mois chargés[/green]\n")

    # Fichiers
    template_file = base_path / "data" / "outputs" / "BP_50M_TEMPLATE.xlsx"
    final_file = base_path / "data" / "outputs" / "BP_50M_FINAL_Nov2025-Dec2029.xlsx"

    # Injecter
    injector = DataInjector(template_file, projections)
    injector.inject_all()
    injector.save(final_file)

    console.print(f"\n[bold green]✅ FICHIER FINAL GÉNÉRÉ[/bold green]")
    console.print(f"[green]📁 {final_file}[/green]")
    console.print(f"\n[cyan]→ Données Python injectées depuis projections_50m.json[/cyan]")
    console.print(f"[cyan]→ Toutes les formules Excel préservées[/cyan]")
    console.print(f"[cyan]→ Prêt pour validation finale[/cyan]\n")


if __name__ == "__main__":
    main()
