#!/usr/bin/env python3
"""
GenieFactory BP 14 Mois - Génération Graphiques PNG
Crée des graphiques PNG pour insertion dans Word

Input:
  - data/structured/projections.json

Output:
  - data/outputs/charts/*.png
"""

import json
import logging
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from typing import List, Dict

# Configuration logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# Style français
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 10


def create_arr_evolution_chart(projections: List[Dict], output_path: Path):
    """Créer graphique évolution ARR"""
    logger.info("📈 Création graphique ARR...")

    months = [f"M{p['month']}" for p in projections]
    arr_values = [p['metrics']['arr'] / 1000 for p in projections]  # en K€

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(months, arr_values, marker='o', linewidth=2, color='#00B050', markersize=6)
    ax.axhline(y=800, color='red', linestyle='--', label='Target 800K€', alpha=0.7)

    ax.set_title('Évolution ARR - Nov 2025 à Dec 2026', fontsize=14, fontweight='bold')
    ax.set_xlabel('Mois', fontsize=11)
    ax.set_ylabel('ARR (K€)', fontsize=11)
    ax.grid(True, alpha=0.3)
    ax.legend()

    # Annotations importantes
    ax.annotate(f'{arr_values[10]:.0f}K€\n(Seed)',
                xy=(10, arr_values[10]), xytext=(8, arr_values[10] + 100),
                arrowprops=dict(arrowstyle='->', color='blue'),
                fontsize=9, ha='center')

    ax.annotate(f'{arr_values[13]:.0f}K€\n(Target)',
                xy=(13, arr_values[13]), xytext=(11, arr_values[13] + 100),
                arrowprops=dict(arrowstyle='->', color='green'),
                fontsize=9, ha='center')

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    logger.info(f"✓ Graphique ARR sauvegardé: {output_path}")


def create_ca_mensuel_chart(projections: List[Dict], output_path: Path):
    """Créer graphique CA mensuel"""
    logger.info("📊 Création graphique CA mensuel...")

    months = [f"M{p['month']}" for p in projections]
    ca_values = [p['revenue']['total'] / 1000 for p in projections]  # en K€

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.bar(months, ca_values, color='#4472C4', alpha=0.8, edgecolor='black', linewidth=0.5)

    ax.set_title('CA Mensuel - Nov 2025 à Dec 2026', fontsize=14, fontweight='bold')
    ax.set_xlabel('Mois', fontsize=11)
    ax.set_ylabel('CA (K€)', fontsize=11)
    ax.grid(True, alpha=0.3, axis='y')

    # Ligne de tendance
    from numpy import polyfit, poly1d
    import numpy as np
    x_vals = np.arange(len(ca_values))
    z = polyfit(x_vals, ca_values, 2)
    p = poly1d(z)
    ax.plot(months, p(x_vals), "r--", alpha=0.5, linewidth=2, label='Tendance')
    ax.legend()

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    logger.info(f"✓ Graphique CA sauvegardé: {output_path}")


def create_revenue_mix_chart(projections: List[Dict], output_path: Path):
    """Créer camembert répartition revenus"""
    logger.info("🥧 Création camembert revenue mix...")

    # Calculer totaux sur 14 mois
    total_hackathon = sum(p['revenue']['hackathon']['revenue'] for p in projections)
    total_factory = sum(p['revenue']['factory']['revenue'] for p in projections)
    total_hub = sum(p['revenue']['enterprise_hub']['mrr'] for p in projections)
    total_services = sum(p['revenue']['services']['revenue'] for p in projections)

    labels = ['Hackathon', 'Factory', 'Enterprise Hub', 'Services']
    sizes = [total_hackathon, total_factory, total_hub, total_services]
    colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000']
    explode = (0.05, 0.05, 0.1, 0)  # Séparer Hub

    fig, ax = plt.subplots(figsize=(8, 6))
    wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, colors=colors,
                                        autopct='%1.1f%%', startangle=90, textprops={'fontsize': 11})

    # Mise en forme
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')

    ax.set_title('Répartition Revenus 14 Mois\n(Total: {:.0f}K€)'.format(sum(sizes)/1000),
                 fontsize=14, fontweight='bold')

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    logger.info(f"✓ Camembert revenue mix sauvegardé: {output_path}")


def create_ebitda_chart(projections: List[Dict], output_path: Path):
    """Créer graphique EBITDA mensuel"""
    logger.info("💰 Création graphique EBITDA...")

    months = [f"M{p['month']}" for p in projections]
    ebitda_values = [p['metrics']['ebitda'] / 1000 for p in projections]  # en K€

    # Couleurs: rouge si négatif, vert si positif
    colors = ['red' if v < 0 else 'green' for v in ebitda_values]

    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(months, ebitda_values, color=colors, alpha=0.7, edgecolor='black', linewidth=0.5)

    ax.axhline(y=0, color='black', linestyle='-', linewidth=1)
    ax.set_title('EBITDA Mensuel - Nov 2025 à Dec 2026', fontsize=14, fontweight='bold')
    ax.set_xlabel('Mois', fontsize=11)
    ax.set_ylabel('EBITDA (K€)', fontsize=11)
    ax.grid(True, alpha=0.3, axis='y')

    # Légende
    red_patch = mpatches.Patch(color='red', alpha=0.7, label='Négatif')
    green_patch = mpatches.Patch(color='green', alpha=0.7, label='Positif')
    ax.legend(handles=[red_patch, green_patch])

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    logger.info(f"✓ Graphique EBITDA sauvegardé: {output_path}")


def create_cash_chart(projections: List[Dict], output_path: Path):
    """Créer graphique cash position"""
    logger.info("💵 Création graphique cash position...")

    months = [f"M{p['month']}" for p in projections]
    cash_values = [p['metrics']['cash'] / 1000 for p in projections]  # en K€

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.fill_between(range(len(months)), 0, cash_values, alpha=0.3, color='#4472C4')
    ax.plot(months, cash_values, marker='o', linewidth=2, color='#4472C4', markersize=5)

    # Marqueurs funding
    ax.axvline(x=0, color='green', linestyle='--', alpha=0.5, label='Pre-seed 150K€')
    ax.axvline(x=10, color='orange', linestyle='--', alpha=0.5, label='Seed 500K€')

    ax.set_title('Position Cash - Nov 2025 à Dec 2026', fontsize=14, fontweight='bold')
    ax.set_xlabel('Mois', fontsize=11)
    ax.set_ylabel('Cash (K€)', fontsize=11)
    ax.grid(True, alpha=0.3)
    ax.legend()

    # Annotations fundings
    ax.annotate('Pre-seed\n150K€', xy=(0, cash_values[0]), xytext=(1, cash_values[0] + 200),
                arrowprops=dict(arrowstyle='->', color='green'),
                fontsize=9, color='green', fontweight='bold')

    ax.annotate('Seed\n500K€', xy=(10, cash_values[10]), xytext=(8, cash_values[10] - 300),
                arrowprops=dict(arrowstyle='->', color='orange'),
                fontsize=9, color='orange', fontweight='bold')

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    logger.info(f"✓ Graphique cash sauvegardé: {output_path}")


def create_team_evolution_chart(projections: List[Dict], output_path: Path):
    """Créer graphique évolution équipe"""
    logger.info("👥 Création graphique évolution équipe...")

    months = [f"M{p['month']}" for p in projections]
    team_values = [p['metrics']['team_size'] for p in projections]

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.step(months, team_values, where='post', linewidth=2, color='#ED7D31', marker='o', markersize=6)
    ax.fill_between(range(len(months)), 0, team_values, step='post', alpha=0.2, color='#ED7D31')

    ax.set_title('Évolution Équipe - Nov 2025 à Dec 2026', fontsize=14, fontweight='bold')
    ax.set_xlabel('Mois', fontsize=11)
    ax.set_ylabel('Effectif (ETP)', fontsize=11)
    ax.set_ylim(0, max(team_values) + 2)
    ax.grid(True, alpha=0.3, axis='y')

    # Annotations paliers
    for i in range(len(team_values) - 1):
        if team_values[i+1] > team_values[i]:
            ax.annotate(f'+{int(team_values[i+1] - team_values[i])}',
                       xy=(i+0.5, team_values[i+1]), xytext=(i+0.5, team_values[i+1] + 0.5),
                       fontsize=8, ha='center', color='red', fontweight='bold')

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    logger.info(f"✓ Graphique équipe sauvegardé: {output_path}")


def main():
    """Fonction principale"""
    logger.info("="*60)
    logger.info("🎨 GÉNÉRATION GRAPHIQUES PNG")
    logger.info("="*60)

    base_path = Path(__file__).parent.parent

    # Charger projections
    projections_path = base_path / "data" / "structured" / "projections.json"
    logger.info(f"📂 Chargement projections: {projections_path}")
    with open(projections_path, 'r', encoding='utf-8') as f:
        projections = json.load(f)

    # Créer dossier charts
    charts_dir = base_path / "data" / "outputs" / "charts"
    charts_dir.mkdir(parents=True, exist_ok=True)

    # Générer tous les graphiques
    create_arr_evolution_chart(projections, charts_dir / "arr_evolution.png")
    create_ca_mensuel_chart(projections, charts_dir / "ca_mensuel.png")
    create_revenue_mix_chart(projections, charts_dir / "revenue_mix.png")
    create_ebitda_chart(projections, charts_dir / "ebitda.png")
    create_cash_chart(projections, charts_dir / "cash_position.png")
    create_team_evolution_chart(projections, charts_dir / "team_evolution.png")

    logger.info("\n" + "="*60)
    logger.info("✅ GRAPHIQUES GÉNÉRÉS")
    logger.info("="*60)
    logger.info(f"📁 Dossier: {charts_dir}")
    logger.info("📊 Fichiers créés:")
    for chart_file in charts_dir.glob("*.png"):
        logger.info(f"  • {chart_file.name}")

    return 0


if __name__ == "__main__":
    exit(main())
