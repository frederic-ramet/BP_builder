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

        # 2. Nettoyer les données
        self.clean_data_cells()

        # 3. Ajouter marqueurs
        self.add_template_markers()

        # 4. Vérifier formules
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
