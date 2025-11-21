#!/usr/bin/env python3
"""
Extraire les hypothèses de croissance du fichier RAW
pour comprendre comment il atteint ses CA
"""

import openpyxl
from pathlib import Path
from rich.console import Console
from rich.table import Table
from rich import box

console = Console()

def extract_revenue_assumptions(raw_file):
    """Extraire les hypothèses de revenus du RAW"""

    wb = openpyxl.load_workbook(raw_file, data_only=True)

    # Analyser le sheet Ventes pour les volumes
    ventes = wb['Ventes']

    console.print("\n[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]")
    console.print("[bold cyan]   EXTRACTION HYPOTHÈSES CROISSANCE - RAW[/bold cyan]")
    console.print("[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]\n")

    # Trouver les lignes clés
    results = {}

    # M1, M14, M26 (colonnes F, S, AE)
    months = {
        'M1': ('F', 1),
        'M14': ('S', 14),
        'M26': ('AE', 26)
    }

    for month_label, (col, month_num) in months.items():
        console.print(f"[yellow]📊 {month_label} (Col {col}):[/yellow]")

        # Lire les différentes lignes de revenus
        revenues = {}

        # Essayer de trouver les lignes automatiquement
        for row in range(2, 30):  # Lignes 2-30
            label = ventes[f'A{row}'].value
            value = ventes[f'{col}{row}'].value

            if label and value and isinstance(value, (int, float)) and value > 0:
                revenues[str(label)] = value
                console.print(f"  L{row} {label}: {value:,.0f}€")

        results[month_label] = revenues
        console.print()

    # Analyser la croissance
    console.print("[bold cyan]═══════════════════════════════════════════════════════[/bold cyan]")
    console.print("[bold]ANALYSE CROISSANCE:[/bold]\n")

    if results.get('M1') and results.get('M14'):
        m1_total = sum(results['M1'].values())
        m14_total = sum(results['M14'].values())

        if m1_total > 0:
            growth_m1_m14 = ((m14_total / m1_total) - 1) * 100
            console.print(f"  M1 → M14: {m1_total:,.0f}€ → {m14_total:,.0f}€ (+{growth_m1_m14:.0f}%)")

    if results.get('M14') and results.get('M26'):
        m14_total = sum(results['M14'].values())
        m26_total = sum(results['M26'].values())

        if m14_total > 0:
            growth_m14_m26 = ((m26_total / m14_total) - 1) * 100
            console.print(f"  M14 → M26: {m14_total:,.0f}€ → {m26_total:,.0f}€ (+{growth_m14_m26:.0f}%)")

    return results

def main():
    base_path = Path(__file__).parent.parent
    raw_file = base_path / "data" / "raw" / "BP FABRIQ_PRODUCT-OCT2025.xlsx"

    results = extract_revenue_assumptions(raw_file)

    console.print("\n[green]✓ Analyse terminée[/green]\n")

if __name__ == "__main__":
    main()
