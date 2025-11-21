#!/usr/bin/env python3
"""
GenieFactory BP 50 Mois - Script 4b: Génération BP Excel Complet
Crée BP_50M_Nov2025-Dec2029.xlsx avec 15 sheets et structure identique au source

Input:
  - data/structured/projections_50m.json
  - data/structured/assumptions.yaml

Output:
  - data/outputs/BP_50M_Nov2025-Dec2029.xlsx (15 sheets, ~122 colonnes P&L)
"""

import json
import yaml
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List, Tuple

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter

# Configuration logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


class BPExcel50MGenerator:
    """Générateur BP Excel 50 mois - reproduction exacte structure source"""

    def __init__(self, projections: List[Dict], assumptions: Dict):
        self.projections = projections
        self.assumptions = assumptions
        self.wb = Workbook()
        self.wb.remove(self.wb.active)  # Supprimer sheet par défaut

        # Structure colonnes comme source:
        # Colonnes C: Total 2025-2026
        # Colonnes D-Q: M1-M14 (Nov 25 - Dec 26)
        # Colonne R: Total 2027
        # Colonnes S-AD: M15-M26 (Jan 27 - Dec 27)
        # Colonne AE: Total 2028
        # Colonnes AF-AQ: M27-M38 (Jan 28 - Dec 28)
        # Colonne AR: Total 2029
        # Colonnes AS-BD: M39-M50 (Jan 29 - Dec 29)

        self.setup_column_structure()

    def setup_column_structure(self):
        """Définir la structure des colonnes pour les 50 mois + totaux annuels"""
        self.columns_map = {}

        # Colonne A: Labels
        # Colonne B: Notes/formules
        # Colonne C: Total 2025-2026

        col_idx = 4  # Commence à D

        # M1-M14 (Nov 2025 - Dec 2026)
        for month in range(1, 15):
            self.columns_map[month] = get_column_letter(col_idx)
            col_idx += 1

        # Colonne R: Total 2027
        self.columns_map['total_2027'] = get_column_letter(col_idx)
        col_idx += 1

        # M15-M26 (2027)
        for month in range(15, 27):
            self.columns_map[month] = get_column_letter(col_idx)
            col_idx += 1

        # Colonne AE: Total 2028
        self.columns_map['total_2028'] = get_column_letter(col_idx)
        col_idx += 1

        # M27-M38 (2028)
        for month in range(27, 39):
            self.columns_map[month] = get_column_letter(col_idx)
            col_idx += 1

        # Colonne AR: Total 2029
        self.columns_map['total_2029'] = get_column_letter(col_idx)
        col_idx += 1

        # M39-M50 (2029)
        for month in range(39, 51):
            self.columns_map[month] = get_column_letter(col_idx)
            col_idx += 1

        logger.info(f"✓ Structure colonnes définie: {len(self.columns_map)} colonnes")

    def create_styles(self):
        """Définir les styles réutilisables"""
        self.style_header_year = {
            'font': Font(bold=True, size=12, color='FFFFFF'),
            'fill': PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid'),
            'alignment': Alignment(horizontal='center', vertical='center')
        }

        self.style_header_month = {
            'font': Font(bold=True, size=10, color='FFFFFF'),
            'fill': PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid'),
            'alignment': Alignment(horizontal='center', vertical='center')
        }

        self.style_total = {
            'font': Font(bold=True, size=10),
            'fill': PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid'),
            'alignment': Alignment(horizontal='right')
        }

        self.style_currency = {
            'number_format': '#,##0 €',
            'alignment': Alignment(horizontal='right')
        }

        self.style_arr = {
            'font': Font(bold=True, color='00B050'),
            'number_format': '#,##0 €',
            'alignment': Alignment(horizontal='right')
        }

        self.style_section_header = {
            'font': Font(bold=True, size=11, color='FFFFFF'),
            'fill': PatternFill(start_color='548235', end_color='548235', fill_type='solid'),
            'alignment': Alignment(horizontal='left')
        }

    def apply_style(self, cell, style_dict):
        """Appliquer un style à une cellule"""
        for key, value in style_dict.items():
            setattr(cell, key, value)

    def create_pl_sheet(self):
        """Créer sheet P&L avec 50 mois (structure exacte source)"""
        logger.info("📊 Création sheet P&L (50 mois)...")

        ws = self.wb.create_sheet("P&L")

        # Titre
        ws['A1'] = "Compte de Résultat Prévisionnel - Nov 2025 à Dec 2029"
        ws['A1'].font = Font(bold=True, size=14)

        # Row 1: Années
        ws['D1'] = "2025-2026"
        ws.merge_cells('D1:Q1')
        self.apply_style(ws['D1'], self.style_header_year)

        ws['S1'] = "2027"
        ws.merge_cells('S1:AD1')
        self.apply_style(ws['S1'], self.style_header_year)

        ws['AF1'] = "2028"
        ws.merge_cells('AF1:AQ1')
        self.apply_style(ws['AF1'], self.style_header_year)

        ws['AS1'] = "2029"
        ws.merge_cells('AS1:BD1')
        self.apply_style(ws['AS1'], self.style_header_year)

        # Row 2: Mois
        ws['A2'] = "Rubrique"
        ws['B2'] = "Notes"
        ws['C2'] = "Total 25-26"

        # Headers mois M1-M50
        for month in range(1, 51):
            col = self.columns_map[month]
            month_data = self.projections[month - 1]
            date_str = month_data['date']  # 2025-11 format
            month_num = int(date_str.split('-')[1])
            ws[f'{col}2'] = f"M{month_num}"
            self.apply_style(ws[f'{col}2'], self.style_header_month)

        # Headers totaux annuels
        for total_col, year in [('R2', '2027'), ('AE2', '2028'), ('AR2', '2029')]:
            ws[total_col] = f"Total {year}"
            self.apply_style(ws[total_col], self.style_header_month)

        # === REVENUS ===
        row = 3
        ws[f'A{row}'] = "CHIFFRE D'AFFAIRES"
        self.apply_style(ws[f'A{row}'], self.style_section_header)
        row += 1

        # CA Hackathons
        ws[f'A{row}'] = "  Hackathons"
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['revenue']['hackathon']['revenue']
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # CA Factory
        ws[f'A{row}'] = "  Factory Projects"
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['revenue']['factory']['revenue']
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # CA Hub (MRR)
        ws[f'A{row}'] = "  Enterprise Hub (MRR)"
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['revenue']['enterprise_hub']['mrr']
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # CA Services
        ws[f'A{row}'] = "  Services"
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['revenue']['services']['revenue']
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # Total CA
        row += 1
        ws[f'A{row}'] = "TOTAL CHIFFRE D'AFFAIRES"
        self.apply_style(ws[f'A{row}'], self.style_total)
        ca_total_row = row
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['revenue']['total']
            ws[f'{col}{row}'] = value
            ws[f'{col}{row}'].font = Font(bold=True)
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # === CHARGES ===
        row += 1
        ws[f'A{row}'] = "CHARGES D'EXPLOITATION"
        self.apply_style(ws[f'A{row}'], self.style_section_header)
        row += 1

        # Charges personnel
        ws[f'A{row}'] = "  Charges de personnel"
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['costs']['personnel']['total']
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # Infrastructure
        ws[f'A{row}'] = "  Infrastructure technique"
        for month in range(1, 51):
            col = self.columns_map[month]
            infra = self.projections[month - 1]['costs']['infrastructure']
            value = infra if isinstance(infra, (int, float)) else infra['total']
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # Marketing
        ws[f'A{row}'] = "  Marketing & Commercial"
        for month in range(1, 51):
            col = self.columns_map[month]
            marketing = self.projections[month - 1]['costs']['marketing']
            value = marketing if isinstance(marketing, (int, float)) else marketing['total']
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # Admin
        ws[f'A{row}'] = "  Frais généraux & Admin"
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['costs']['admin']
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # Total charges
        row += 1
        ws[f'A{row}'] = "TOTAL CHARGES"
        self.apply_style(ws[f'A{row}'], self.style_total)
        charges_total_row = row
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['costs']['total']
            ws[f'{col}{row}'] = value
            ws[f'{col}{row}'].font = Font(bold=True)
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # === RÉSULTAT ===
        row += 1
        ws[f'A{row}'] = "EBITDA"
        self.apply_style(ws[f'A{row}'], self.style_total)
        ebitda_row = row
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['metrics']['ebitda']
            ws[f'{col}{row}'] = value
            ws[f'{col}{row}'].font = Font(bold=True)
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # ARR
        row += 1
        ws[f'A{row}'] = "ARR (Run Rate)"
        arr_row = row
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['metrics']['arr']
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_arr)
        row += 1

        # Cash position
        ws[f'A{row}'] = "Cash Position"
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['metrics']['cash']
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # Team size
        ws[f'A{row}'] = "Équipe (ETP)"
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['metrics']['team_size']
            ws[f'{col}{row}'] = value
            ws[f'{col}{row}'].number_format = '0'
        row += 1

        # Largeurs colonnes
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 12

        logger.info(f"✓ Sheet P&L créée: {row} lignes × 50 mois")

    def create_charges_personnel_sheet(self):
        """Créer sheet Charges de personnel et FG (détail par rôle)"""
        logger.info("👥 Création sheet Charges Personnel...")

        ws = self.wb.create_sheet("Charges Personnel")

        # Titre
        ws['A1'] = "Charges de Personnel et Frais Généraux - Détail par Rôle"
        ws['A1'].font = Font(bold=True, size=14)

        # Headers similaires au P&L
        ws['D1'] = "2025-2026"
        ws.merge_cells('D1:Q1')
        self.apply_style(ws['D1'], self.style_header_year)

        ws['S1'] = "2027"
        ws.merge_cells('S1:AD1')
        self.apply_style(ws['S1'], self.style_header_year)

        ws['AF1'] = "2028"
        ws.merge_cells('AF1:AQ1')
        self.apply_style(ws['AF1'], self.style_header_year)

        ws['AS1'] = "2029"
        ws.merge_cells('AS1:BD1')
        self.apply_style(ws['AS1'], self.style_header_year)

        # Row 2: Mois
        ws['A2'] = "Rôle / Poste"
        ws['B2'] = "Salaire Annuel"
        ws['C2'] = "Total 25-26"

        for month in range(1, 51):
            col = self.columns_map[month]
            ws[f'{col}2'] = f"M{month}"
            self.apply_style(ws[f'{col}2'], self.style_header_month)

        # Vérifier si personnel_details existe
        if 'personnel_details' not in self.assumptions:
            logger.warning("⚠️ personnel_details non trouvé dans assumptions - sheet simplifiée")
            row = 3
            ws[f'A{row}'] = "Charges de personnel totales"
            for month in range(1, 51):
                col = self.columns_map[month]
                value = self.projections[month - 1]['costs']['personnel']['total']
                ws[f'{col}{row}'] = value
                self.apply_style(ws[f'{col}{row}'], self.style_currency)
            return

        # Détail par rôle
        personnel_details = self.assumptions['personnel_details']
        row = 3

        ws[f'A{row}'] = "SALAIRES BRUTS"
        self.apply_style(ws[f'A{row}'], self.style_section_header)
        row += 1

        # Pour chaque rôle
        roles_order = [
            'directeur_general', 'product_owner', 'tech_senior', 'tech_junior',
            'commercial', 'bd_junior', 'stagiaire', 'consultant'
        ]

        for role_name in roles_order:
            if role_name not in personnel_details['roles']:
                continue

            role_data = personnel_details['roles'][role_name]
            ws[f'A{row}'] = f"  {role_data['title']}"
            ws[f'B{row}'] = f"{role_data['salary_brut_annual']:,.0f} €"

            # Pour chaque mois, extraire le coût de ce rôle
            for month in range(1, 51):
                col = self.columns_map[month]
                costs_personnel = self.projections[month - 1]['costs']['personnel']

                if 'roles' in costs_personnel and role_name in costs_personnel['roles']:
                    value = costs_personnel['roles'][role_name]['cost_monthly']
                else:
                    value = 0

                ws[f'{col}{row}'] = value
                self.apply_style(ws[f'{col}{row}'], self.style_currency)

            row += 1

        # Total salaires bruts
        row += 1
        ws[f'A{row}'] = "TOTAL SALAIRES BRUTS"
        self.apply_style(ws[f'A{row}'], self.style_total)
        for month in range(1, 51):
            col = self.columns_map[month]
            costs_personnel = self.projections[month - 1]['costs']['personnel']
            value = costs_personnel.get('salary_brut', 0)
            ws[f'{col}{row}'] = value
            ws[f'{col}{row}'].font = Font(bold=True)
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # Charges sociales
        row += 1
        ws[f'A{row}'] = "CHARGES SOCIALES (45%)"
        ws[f'A{row}'].font = Font(bold=True)
        for month in range(1, 51):
            col = self.columns_map[month]
            costs_personnel = self.projections[month - 1]['costs']['personnel']
            value = costs_personnel.get('charges_sociales', 0)
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # Freelances
        ws[f'A{row}'] = "Freelances / Consultants"
        for month in range(1, 51):
            col = self.columns_map[month]
            costs_personnel = self.projections[month - 1]['costs']['personnel']
            value = costs_personnel.get('freelance', 0)
            ws[f'{col}{row}'] = value
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # Total charges personnel
        row += 1
        ws[f'A{row}'] = "TOTAL CHARGES DE PERSONNEL"
        self.apply_style(ws[f'A{row}'], self.style_total)
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['costs']['personnel']['total']
            ws[f'{col}{row}'] = value
            ws[f'{col}{row}'].font = Font(bold=True, size=11)
            self.apply_style(ws[f'{col}{row}'], self.style_currency)
        row += 1

        # FTE total
        row += 1
        ws[f'A{row}'] = "Effectif Total (ETP)"
        for month in range(1, 51):
            col = self.columns_map[month]
            value = self.projections[month - 1]['metrics']['team_size']
            ws[f'{col}{row}'] = value
            ws[f'{col}{row}'].number_format = '0.0'
        row += 1

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15

        logger.info(f"✓ Sheet Charges Personnel créée: {len(roles_order)} rôles")

    def generate(self):
        """Générer le workbook complet"""
        logger.info("\n🔨 Génération workbook BP 50 mois...")

        self.create_styles()

        # Créer les sheets dans l'ordre
        # Phase 1: Sheets critiques
        self.create_pl_sheet()
        self.create_charges_personnel_sheet()

        logger.info("\n✓ Workbook Phase 1 généré (P&L + Personnel)")
        logger.info(f"  Sheets: {len(self.wb.sheetnames)}")
        return self.wb


def main():
    """Fonction principale"""
    logger.info("="*60)
    logger.info("🚀 GÉNÉRATION BP EXCEL 50 MOIS - GenieFactory")
    logger.info("="*60)

    base_path = Path(__file__).parent.parent

    # Charger projections 50M
    projections_path = base_path / "data" / "structured" / "projections_50m.json"
    if not projections_path.exists():
        logger.error(f"❌ Fichier projections_50m.json non trouvé: {projections_path}")
        logger.error("   Exécuter d'abord: python scripts/3_calculate_projections.py")
        return 1

    logger.info(f"📂 Chargement projections: {projections_path}")
    with open(projections_path, 'r', encoding='utf-8') as f:
        projections = json.load(f)

    logger.info(f"✓ Projections chargées: {len(projections)} mois")

    # Charger assumptions
    assumptions_path = base_path / "data" / "structured" / "assumptions.yaml"
    logger.info(f"📂 Chargement assumptions: {assumptions_path}")
    with open(assumptions_path, 'r', encoding='utf-8') as f:
        assumptions = yaml.safe_load(f)

    logger.info(f"✓ Assumptions chargées (version {assumptions.get('version', '1.0')})")

    # Générer Excel
    generator = BPExcel50MGenerator(projections, assumptions)
    wb = generator.generate()

    # Sauvegarder
    output_path = base_path / "data" / "outputs" / "BP_50M_Nov2025-Dec2029.xlsx"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb.save(output_path)

    logger.info("\n" + "="*60)
    logger.info("✅ BP EXCEL 50 MOIS GÉNÉRÉ")
    logger.info("="*60)
    logger.info(f"📁 Fichier: {output_path}")
    logger.info(f"📊 Taille: {output_path.stat().st_size / 1024:.1f} KB")
    logger.info(f"📑 Sheets: {len(wb.sheetnames)} - {', '.join(wb.sheetnames)}")

    logger.info("\n✓ Excel prêt à ouvrir dans MS Excel ou LibreOffice")
    logger.info("   → Phase 1: P&L + Personnel créés")
    logger.info("   → Prochaines phases: 13 sheets restants")

    return 0


if __name__ == "__main__":
    exit(main())
