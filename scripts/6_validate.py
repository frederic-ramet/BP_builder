#!/usr/bin/env python3
"""
GenieFactory BP 14 Mois - Script 6: Validation
Vérifier cohérence et targets

Inputs:
  - data/structured/projections.json
  - data/structured/assumptions.yaml
  - data/outputs/BP_14M_Nov2025-Dec2026.xlsx
  - data/outputs/BM_Updated_14M.docx

Output:
  - Rapport validation (console + logs/validation_report_YYYYMMDD.txt)
"""

import json
import yaml
import re
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List, Tuple

import openpyxl
from docx import Document
from rich.console import Console
from rich.table import Table
from rich.panel import Panel

# Configuration logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

console = Console()


class Validator:
    """Validateur de cohérence BP"""

    def __init__(self, projections: List[Dict], assumptions: Dict):
        self.projections = projections
        self.assumptions = assumptions
        self.errors = []
        self.warnings = []
        self.checks_passed = []

    def check_arr_targets(self) -> bool:
        """Vérifier ARR targets"""
        console.print("\n[cyan]📊 CHECKS ARR TARGETS[/]")

        target_m14 = self.assumptions['financial_kpis']['target_arr_dec_2026']
        arr_m14 = self.projections[13]['metrics']['arr']
        arr_m11 = self.projections[10]['metrics']['arr']

        tolerance = self.assumptions['validation_rules']['arr_tolerance_pct']
        min_arr = target_m14 * (1 - tolerance)
        max_arr = target_m14 * (1 + tolerance)

        # Check M14
        if min_arr <= arr_m14 <= max_arr:
            self.checks_passed.append(
                f"ARR M14: {arr_m14:,.0f}€ (target {target_m14:,.0f}€ ±{tolerance:.0%})"
            )
            console.print(f"  ✓ ARR M14: {arr_m14:,.0f}€ [green](target {target_m14:,.0f}€ ±{tolerance:.0%})[/]")
        elif arr_m14 < min_arr:
            self.errors.append(
                f"ARR M14 trop bas: {arr_m14:,.0f}€ (min {min_arr:,.0f}€)"
            )
            console.print(f"  ✗ ARR M14: {arr_m14:,.0f}€ [red](< {min_arr:,.0f}€)[/]")
        else:
            self.warnings.append(
                f"ARR M14 optimiste: {arr_m14:,.0f}€ (max {max_arr:,.0f}€)"
            )
            console.print(f"  ⚠ ARR M14: {arr_m14:,.0f}€ [yellow](> {max_arr:,.0f}€)[/]")

        # Check M11 (avant seed)
        min_arr_m11 = self.assumptions['validation_rules']['arr_m11_min']
        if arr_m11 >= min_arr_m11:
            self.checks_passed.append(
                f"ARR M11: {arr_m11:,.0f}€ (>= {min_arr_m11:,.0f}€)"
            )
            console.print(f"  ✓ ARR M11: {arr_m11:,.0f}€ [green](>= {min_arr_m11:,.0f}€)[/]")
        else:
            self.warnings.append(
                f"ARR M11 faible: {arr_m11:,.0f}€ (min conseillé {min_arr_m11:,.0f}€)"
            )
            console.print(f"  ⚠ ARR M11: {arr_m11:,.0f}€ [yellow](< {min_arr_m11:,.0f}€)[/]")

        return len(self.errors) == 0

    def check_cash_position(self) -> bool:
        """Vérifier cash jamais négatif"""
        console.print("\n[cyan]💰 CHECK CASH POSITION[/]")

        min_cash_balance = self.assumptions['validation_rules']['min_cash_balance']
        negative_months = []

        for month_data in self.projections:
            cash = month_data['metrics']['cash']
            if cash < 0:
                negative_months.append((month_data['month'], cash))
            elif cash < min_cash_balance:
                self.warnings.append(
                    f"Cash M{month_data['month']} bas: {cash:,.0f}€ (< {min_cash_balance:,.0f}€)"
                )

        if negative_months:
            for month, cash in negative_months:
                self.errors.append(f"Cash négatif M{month}: {cash:,.0f}€")
                console.print(f"  ✗ Cash M{month}: {cash:,.0f}€ [red](NÉGATIF!)[/]")
            return False
        else:
            min_cash = min(m['metrics']['cash'] for m in self.projections)
            min_cash_month = next(m['month'] for m in self.projections if m['metrics']['cash'] == min_cash)
            self.checks_passed.append(
                f"Cash min: {min_cash:,.0f}€ (M{min_cash_month})"
            )
            console.print(f"  ✓ Cash toujours positif [green](min: {min_cash:,.0f}€ à M{min_cash_month})[/]")
            return True

    def check_burn_rate(self) -> bool:
        """Vérifier burn rate acceptable"""
        console.print("\n[cyan]🔥 CHECK BURN RATE[/]")

        max_burn_allowed = self.assumptions['validation_rules']['max_burn_monthly']
        max_burn = max(m['metrics']['burn_rate'] for m in self.projections)
        max_burn_month = next(m['month'] for m in self.projections if m['metrics']['burn_rate'] == max_burn)

        avg_burn = sum(m['metrics']['burn_rate'] for m in self.projections) / len(self.projections)

        if max_burn <= max_burn_allowed:
            self.checks_passed.append(
                f"Burn rate max: {max_burn:,.0f}€/mois (M{max_burn_month}, limite {max_burn_allowed:,.0f}€)"
            )
            console.print(
                f"  ✓ Burn max: {max_burn:,.0f}€/mois [green](M{max_burn_month}, limite {max_burn_allowed:,.0f}€)[/]"
            )
            console.print(f"  ✓ Burn moyen: {avg_burn:,.0f}€/mois")
            return True
        else:
            self.errors.append(
                f"Burn rate trop élevé: {max_burn:,.0f}€/mois (max {max_burn_allowed:,.0f}€)"
            )
            console.print(
                f"  ✗ Burn max: {max_burn:,.0f}€/mois [red](> {max_burn_allowed:,.0f}€)[/]"
            )
            return False

    def check_team_size(self) -> bool:
        """Vérifier taille équipe raisonnable"""
        console.print("\n[cyan]👥 CHECK ÉQUIPE[/]")

        max_team = self.assumptions['validation_rules']['max_team_size']
        team_m14 = self.projections[13]['metrics']['team_size']
        team_m1 = self.projections[0]['metrics']['team_size']

        if team_m14 <= max_team:
            self.checks_passed.append(
                f"Équipe M14: {team_m14} ETP (max {max_team})"
            )
            console.print(
                f"  ✓ Équipe M1→M14: {team_m1} → {team_m14} ETP [green](max {max_team})[/]"
            )
            return True
        else:
            self.warnings.append(
                f"Équipe large M14: {team_m14} ETP (max conseillé {max_team})"
            )
            console.print(
                f"  ⚠ Équipe M14: {team_m14} ETP [yellow](> {max_team})[/]"
            )
            return True

    def check_conversion_rates(self) -> bool:
        """Vérifier taux de conversion"""
        console.print("\n[cyan]📈 CHECK TAUX CONVERSION[/]")

        # Calculer conversion réelle hackathon → factory
        total_hackathons = sum(m['revenue']['hackathon']['volume'] for m in self.projections)
        total_factory = sum(m['revenue']['factory']['volume'] for m in self.projections)

        if total_hackathons > 0:
            actual_conversion = total_factory / total_hackathons
            target_conversion = self.assumptions['sales_assumptions']['factory']['conversion_rate']
            min_conversion = self.assumptions['validation_rules']['min_conversion_hackathon_factory']

            if actual_conversion >= min_conversion:
                self.checks_passed.append(
                    f"Conversion Hack→Factory: {actual_conversion:.1%} (target {target_conversion:.0%})"
                )
                console.print(
                    f"  ✓ Conversion Hack→Factory: {actual_conversion:.1%} [green](target {target_conversion:.0%})[/]"
                )
            else:
                self.warnings.append(
                    f"Conversion faible: {actual_conversion:.1%} (min {min_conversion:.0%})"
                )
                console.print(
                    f"  ⚠ Conversion Hack→Factory: {actual_conversion:.1%} [yellow](< {min_conversion:.0%})[/]"
                )

        return True

    def check_excel_formulas(self, excel_path: Path) -> bool:
        """Vérifier formules Excel actives"""
        console.print("\n[cyan]📊 CHECK FORMULES EXCEL[/]")

        try:
            wb = openpyxl.load_workbook(excel_path, data_only=False)
            pl_sheet = wb['P&L']

            formulas_found = 0
            formulas_checked = [
                ('F4', 'SUM'),  # CA Total M1
                ('F16', '*'),   # ARR M1
            ]

            for cell_ref, expected_pattern in formulas_checked:
                cell = pl_sheet[cell_ref]
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    if expected_pattern in cell.value:
                        formulas_found += 1

            if formulas_found >= len(formulas_checked) * 0.5:  # Au moins 50%
                self.checks_passed.append(
                    f"Formules Excel actives: {formulas_found}/{len(formulas_checked)} vérifiées"
                )
                console.print(
                    f"  ✓ Formules Excel actives [green]({formulas_found} vérifiées)[/]"
                )
                return True
            else:
                self.warnings.append(
                    f"Peu de formules détectées: {formulas_found}/{len(formulas_checked)}"
                )
                console.print(
                    f"  ⚠ Formules Excel: {formulas_found}/{len(formulas_checked)} [yellow](hardcoded?)[/]"
                )
                return True

        except Exception as e:
            self.warnings.append(f"Erreur lecture Excel: {str(e)}")
            console.print(f"  ⚠ Impossible vérifier formules Excel: {str(e)}")
            return True

    def check_excel_word_consistency(self, excel_path: Path, word_path: Path) -> bool:
        """Vérifier cohérence Excel ↔ Word"""
        console.print("\n[cyan]🔗 CHECK COHÉRENCE EXCEL ↔ WORD[/]")

        try:
            # ARR M14 depuis projections
            arr_proj = self.projections[13]['metrics']['arr']

            # ARR depuis Excel
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            pl_sheet = wb['P&L']
            arr_excel = pl_sheet['S19'].value  # Dernière colonne (M14), ligne ARR (row 19)

            if arr_excel is None:
                arr_excel = 0

            # ARR depuis Word (extraction pattern)
            doc = Document(word_path)
            full_text = '\n'.join([p.text for p in doc.paragraphs])
            arr_matches = re.findall(r'ARR.*?(\d[\d\s,\.]+)\s*[K€]', full_text, re.IGNORECASE)

            arr_word = 0
            if arr_matches:
                # Prendre la plus grande valeur (probablement M14)
                for match in arr_matches:
                    value_str = match.replace(' ', '').replace(',', '').replace('.', '')
                    try:
                        value = int(value_str)
                        if 'K' in full_text[full_text.find(match):full_text.find(match)+20]:
                            value *= 1000
                        if value > arr_word:
                            arr_word = value
                    except:
                        pass

            # Comparaison
            max_deviation = self.assumptions['validation_rules']['max_deviation_excel_word_pct']
            deviation_excel = abs(arr_excel - arr_proj) / arr_proj if arr_proj > 0 else 0
            deviation_word = abs(arr_word - arr_proj) / arr_proj if arr_proj > 0 and arr_word > 0 else 1

            console.print(f"  ARR Projections: {arr_proj:,.0f}€")
            console.print(f"  ARR Excel: {arr_excel:,.0f}€")
            console.print(f"  ARR Word: {arr_word:,.0f}€" if arr_word > 0 else "  ARR Word: Non détecté")

            if deviation_excel <= max_deviation:
                self.checks_passed.append(
                    f"Cohérence Excel: {deviation_excel:.1%} écart"
                )
                console.print(f"  ✓ Excel ↔ Projections: [green]{deviation_excel:.1%} écart[/]")
            else:
                self.errors.append(
                    f"Incohérence Excel: {deviation_excel:.1%} écart (max {max_deviation:.0%})"
                )
                console.print(f"  ✗ Excel ↔ Projections: [red]{deviation_excel:.1%} écart[/]")

            if arr_word > 0:
                if deviation_word <= max_deviation:
                    self.checks_passed.append(
                        f"Cohérence Word: {deviation_word:.1%} écart"
                    )
                    console.print(f"  ✓ Word ↔ Projections: [green]{deviation_word:.1%} écart[/]")
                else:
                    self.warnings.append(
                        f"Incohérence Word: {deviation_word:.1%} écart"
                    )
                    console.print(f"  ⚠ Word ↔ Projections: [yellow]{deviation_word:.1%} écart[/]")

            return deviation_excel <= max_deviation

        except Exception as e:
            self.warnings.append(f"Erreur check cohérence: {str(e)}")
            console.print(f"  ⚠ Erreur: {str(e)}")
            return True

    def generate_report(self) -> str:
        """Générer rapport de validation"""
        status = "✅ PASSED" if len(self.errors) == 0 else "❌ FAILED"
        if len(self.warnings) > 0 and len(self.errors) == 0:
            status = f"✅ PASSED ({len(self.warnings)} warnings)"

        report = []
        report.append("="*60)
        report.append("🔍 RAPPORT VALIDATION - GenieFactory BP 14 Mois")
        report.append("="*60)
        report.append(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append(f"Status: {status}")
        report.append("")

        if self.checks_passed:
            report.append("✅ CHECKS PASSED:")
            for check in self.checks_passed:
                report.append(f"  ✓ {check}")
            report.append("")

        if self.warnings:
            report.append("⚠️ WARNINGS:")
            for warning in self.warnings:
                report.append(f"  • {warning}")
            report.append("")

        if self.errors:
            report.append("❌ ERRORS:")
            for error in self.errors:
                report.append(f"  ✗ {error}")
            report.append("")

        report.append("="*60)

        return '\n'.join(report)


def main():
    """Fonction principale"""
    console.print(Panel.fit(
        "[bold cyan]🔍 VALIDATION FINALE[/]\n"
        "[dim]GenieFactory BP 14 Mois[/]",
        border_style="cyan"
    ))

    base_path = Path(__file__).parent.parent

    # Charger données
    projections_path = base_path / "data" / "structured" / "projections.json"
    assumptions_path = base_path / "data" / "structured" / "assumptions.yaml"
    excel_path = base_path / "data" / "outputs" / "BP_14M_Nov2025-Dec2026.xlsx"
    word_path = base_path / "data" / "outputs" / "BM_Updated_14M.docx"

    # Vérifier existence fichiers
    missing = []
    for path in [projections_path, assumptions_path, excel_path, word_path]:
        if not path.exists():
            missing.append(path.name)

    if missing:
        console.print(f"\n[red]❌ Fichiers manquants: {', '.join(missing)}[/]")
        console.print("[yellow]Exécuter les scripts précédents d'abord[/]")
        return 1

    # Charger
    console.print("\n[cyan]📂 Chargement données...[/]")
    with open(projections_path, 'r', encoding='utf-8') as f:
        projections = json.load(f)

    with open(assumptions_path, 'r', encoding='utf-8') as f:
        assumptions = yaml.safe_load(f)

    # Validation
    validator = Validator(projections, assumptions)

    validator.check_arr_targets()
    validator.check_cash_position()
    validator.check_burn_rate()
    validator.check_team_size()
    validator.check_conversion_rates()
    validator.check_excel_formulas(excel_path)
    validator.check_excel_word_consistency(excel_path, word_path)

    # Rapport
    report = validator.generate_report()
    console.print(f"\n{report}")

    # Sauvegarder rapport
    logs_dir = base_path / "logs"
    logs_dir.mkdir(exist_ok=True)
    report_path = logs_dir / f"validation_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report)

    console.print(f"\n📄 Rapport sauvegardé: {report_path}")

    # Status final
    if len(validator.errors) == 0:
        console.print(Panel.fit(
            "[bold green]✅ VALIDATION RÉUSSIE[/]\n"
            f"[dim]{len(validator.checks_passed)} checks passed, "
            f"{len(validator.warnings)} warnings[/]",
            border_style="green"
        ))
        return 0
    else:
        console.print(Panel.fit(
            "[bold red]❌ VALIDATION ÉCHOUÉE[/]\n"
            f"[dim]{len(validator.errors)} errors, "
            f"{len(validator.warnings)} warnings[/]",
            border_style="red"
        ))
        return 1


if __name__ == "__main__":
    exit(main())
