#!/usr/bin/env python3
"""
Analyser TOUTES les sections de assumptions.yaml et funding_captable.yaml
Identifier ce qui n'est PAS encore mappé dans Excel
"""

import yaml
from pathlib import Path
from rich.console import Console
from rich.table import Table
from rich import box

console = Console()


def analyze_yaml_structure(yaml_path: Path, name: str):
    """Analyser la structure d'un fichier YAML"""
    with open(yaml_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)

    sections = []

    def extract_sections(d, prefix=""):
        if isinstance(d, dict):
            for key, value in d.items():
                full_key = f"{prefix}.{key}" if prefix else key
                sections.append({
                    'section': full_key,
                    'type': type(value).__name__,
                    'has_subsections': isinstance(value, dict)
                })
                if isinstance(value, dict):
                    extract_sections(value, full_key)

    extract_sections(data)
    return sections


def main():
    console.print("\n[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]")
    console.print("[bold cyan]   ANALYSE COMPLÈTE MAPPING YAML → EXCEL[/bold cyan]")
    console.print("[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]\n")

    base_path = Path(__file__).parent.parent
    assumptions_path = base_path / "data" / "structured" / "assumptions.yaml"
    captable_path = base_path / "data" / "structured" / "funding_captable.yaml"

    # Analyser assumptions.yaml
    console.print("[yellow]📂 Analyse assumptions.yaml...[/yellow]")
    assumptions_sections = analyze_yaml_structure(assumptions_path, "assumptions.yaml")

    console.print(f"  ✓ {len(assumptions_sections)} sections trouvées\n")

    # Analyser funding_captable.yaml
    console.print("[yellow]📂 Analyse funding_captable.yaml...[/yellow]")
    captable_sections = analyze_yaml_structure(captable_path, "funding_captable.yaml")

    console.print(f"  ✓ {len(captable_sections)} sections trouvées\n")

    # Mapping actuel (ce qui est déjà fait)
    current_mappings = {
        'pricing.hackathon': {'sheet': 'Paramètres', 'cell': 'B3', 'mapped': True},
        'pricing.factory': {'sheet': 'Paramètres', 'cell': 'D3', 'mapped': True},
        'pricing.enterprise_hub': {'sheet': 'Paramètres', 'cell': 'F3-H3', 'mapped': True},
        'pricing.services': {'sheet': 'Paramètres', 'cell': 'J3', 'mapped': True},

        'timeline.milestones': {'sheet': 'Financement', 'cell': 'C2-G2', 'mapped': True},

        'sales_assumptions.hackathon.volumes_monthly': {'sheet': 'Ventes', 'cell': 'F-BC (via JSON)', 'mapped': True},
        'sales_assumptions.factory': {'sheet': 'Ventes', 'cell': 'Calculé via conversion', 'mapped': True},
        'sales_assumptions.enterprise_hub': {'sheet': 'Ventes', 'cell': 'F-BC (via JSON)', 'mapped': True},
        'sales_assumptions.long_term_sales': {'sheet': 'Ventes', 'cell': 'F-BC (via JSON)', 'mapped': True},

        'costs.personnel': {'sheet': 'Charges de personnel et FG', 'cell': 'A1-B20', 'mapped': True},
        'personnel_details': {'sheet': 'Charges de personnel et FG', 'cell': 'A1-B20', 'mapped': True},

        'costs.infrastructure': {'sheet': 'Infrastructure technique', 'cell': 'A1-B15', 'mapped': True},
        'infrastructure_costs': {'sheet': 'Infrastructure technique', 'cell': 'A1-B15', 'mapped': True},

        'costs.marketing': {'sheet': 'Marketing', 'cell': 'A1-K10', 'mapped': True},
        'marketing_budgets': {'sheet': 'Marketing', 'cell': 'A1-K10', 'mapped': True},

        'funding_captable.funding_rounds': {'sheet': 'Fundings', 'cell': 'A1-E10', 'mapped': True},
        'funding_captable.captable': {'sheet': 'Fundings', 'cell': 'A15-F22', 'mapped': True},
        'funding_captable.arr_targets': {'sheet': 'Fundings', 'cell': 'A25-C32', 'mapped': True},
    }

    # Identifier les sections principales de assumptions.yaml
    main_sections = [
        'meta', 'timeline', 'pricing', 'sales_assumptions', 'costs',
        'financial_kpis', 'validation_rules', 'scenarios', 'critical_assumptions',
        'long_term_projections', 'personnel_details', 'infrastructure_costs',
        'marketing_budgets', 'revision_history', 'usage_notes'
    ]

    console.print("[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]")
    console.print("[bold]SECTIONS PRINCIPALES - MAPPING STATUS[/bold]")
    console.print("[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]\n")

    table = Table(box=box.ROUNDED, title="Mapping Status par Section")
    table.add_column("Section YAML", style="cyan", width=35)
    table.add_column("Sheet Excel", style="yellow", width=25)
    table.add_column("Status", justify="center", width=10)
    table.add_column("Priorité", justify="center", width=10)

    not_mapped = []

    for section in main_sections:
        if section in current_mappings:
            mapping = current_mappings[section]
            table.add_row(
                section,
                mapping['sheet'],
                "✅ Mappé",
                "-"
            )
        else:
            # Déterminer priorité et sheet suggéré
            if section in ['financial_kpis', 'validation_rules']:
                priority = "🔴 HAUTE"
                suggested_sheet = "Paramètres / P&L"
                not_mapped.append({'section': section, 'priority': 'HAUTE', 'sheet': suggested_sheet})
            elif section in ['scenarios', 'critical_assumptions']:
                priority = "🟡 MOYENNE"
                suggested_sheet = "NOUVEAU: Scenarios"
                not_mapped.append({'section': section, 'priority': 'MOYENNE', 'sheet': suggested_sheet})
            elif section in ['meta', 'revision_history', 'usage_notes']:
                priority = "🟢 BASSE"
                suggested_sheet = "NOUVEAU: Documentation"
                not_mapped.append({'section': section, 'priority': 'BASSE', 'sheet': suggested_sheet})
            elif section == 'long_term_projections':
                priority = "🟡 MOYENNE"
                suggested_sheet = "P&L / Ventes (années)"
                not_mapped.append({'section': section, 'priority': 'MOYENNE', 'sheet': suggested_sheet})
            else:
                priority = "🟢 BASSE"
                suggested_sheet = "À définir"
                not_mapped.append({'section': section, 'priority': 'BASSE', 'sheet': suggested_sheet})

            table.add_row(
                section,
                suggested_sheet,
                "❌ Non mappé",
                priority
            )

    console.print(table)
    console.print()

    # Résumé
    mapped_count = len([s for s in main_sections if s in current_mappings])
    not_mapped_count = len(not_mapped)

    console.print(f"[bold]Résumé:[/bold]")
    console.print(f"  • Total sections principales: {len(main_sections)}")
    console.print(f"  • Mappées: {mapped_count} ({mapped_count/len(main_sections)*100:.0f}%)")
    console.print(f"  • Non mappées: {not_mapped_count} ({not_mapped_count/len(main_sections)*100:.0f}%)")
    console.print()

    # Sections non mappées par priorité
    high_priority = [s for s in not_mapped if s['priority'] == 'HAUTE']
    medium_priority = [s for s in not_mapped if s['priority'] == 'MOYENNE']
    low_priority = [s for s in not_mapped if s['priority'] == 'BASSE']

    if high_priority:
        console.print(f"[bold red]🔴 HAUTE PRIORITÉ ({len(high_priority)} sections):[/bold red]")
        for item in high_priority:
            console.print(f"  • {item['section']} → {item['sheet']}")
        console.print()

    if medium_priority:
        console.print(f"[bold yellow]🟡 MOYENNE PRIORITÉ ({len(medium_priority)} sections):[/bold yellow]")
        for item in medium_priority:
            console.print(f"  • {item['section']} → {item['sheet']}")
        console.print()

    if low_priority:
        console.print(f"[bold green]🟢 BASSE PRIORITÉ ({len(low_priority)} sections):[/bold green]")
        for item in low_priority:
            console.print(f"  • {item['section']} → {item['sheet']}")
        console.print()

    # Recommandations
    console.print("[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]")
    console.print("[bold]RECOMMANDATIONS MAPPING COMPLET[/bold]")
    console.print("[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]\n")

    console.print("[bold]Pour atteindre 100% de mapping, ajouter:[/bold]\n")

    console.print("1. [bold]Dans Paramètres (enrichir):[/bold]")
    console.print("   • financial_kpis (ARR targets, marges, burn rate)")
    console.print("   • validation_rules (min/max checks)")
    console.print()

    console.print("2. [bold]Créer nouveau sheet 'Scenarios':[/bold]")
    console.print("   • scenarios (base/upside/downside)")
    console.print("   • critical_assumptions")
    console.print()

    console.print("3. [bold]Créer nouveau sheet 'Documentation':[/bold]")
    console.print("   • meta (version, auteur, sources)")
    console.print("   • revision_history")
    console.print("   • usage_notes")
    console.print()

    console.print("4. [bold]Dans P&L (ajouter lignes annuelles):[/bold]")
    console.print("   • long_term_projections (growth rates par année)")
    console.print()

    return not_mapped


if __name__ == "__main__":
    not_mapped = main()
