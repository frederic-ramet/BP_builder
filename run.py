#!/usr/bin/env python3
"""
GenieFactory BP 14 Mois - Orchestrateur Principal
Exécute séquentiellement tous les scripts de génération du Business Plan

Usage:
    python run.py                    # Exécution complète
    python run.py --skip-extract     # Skip extraction (si déjà fait)
    python run.py --validate-only    # Seulement validation
"""

import sys
import subprocess
import argparse
from pathlib import Path
from datetime import datetime
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn
from rich.panel import Panel
from rich import print as rprint

# Setup console
console = Console()

# Scripts à exécuter (ordre important)
SCRIPTS = [
    {
        "name": "1. Extraction",
        "script": "scripts/1_extract.py",
        "description": "Parse BP Excel, BM Word, Pacte",
        "skip_flag": "--skip-extract"
    },
    {
        "name": "2. Assumptions",
        "script": "scripts/2_generate_assumptions.py",
        "description": "Génère assumptions.yaml (validation manuelle requise)",
        "skip_flag": "--skip-assumptions"
    },
    {
        "name": "3. Projections",
        "script": "scripts/3_calculate_projections.py",
        "description": "Calcule ARR, CA, charges M1-M14",
        "skip_flag": None  # Jamais skip
    },
    {
        "name": "4. BP Excel",
        "script": "scripts/4_generate_bp_excel.py",
        "description": "Génère BP_14M_Nov2025-Dec2026.xlsx",
        "skip_flag": None
    },
    {
        "name": "5. BM Word",
        "script": "scripts/5_update_bm_word.py",
        "description": "Update BM_Updated_14M.docx",
        "skip_flag": None
    },
    {
        "name": "6. Validation",
        "script": "scripts/6_validate.py",
        "description": "Checks cohérence et targets",
        "skip_flag": None
    }
]


def check_dependencies():
    """Vérifier que toutes les dépendances sont installées"""
    console.print("\n[bold cyan]🔍 Vérification des dépendances...[/]")
    
    required = ['openpyxl', 'docx', 'yaml', 'pandas', 'rich']
    missing = []
    
    for module in required:
        try:
            __import__(module)
            console.print(f"  ✓ {module}")
        except ImportError:
            missing.append(module)
            console.print(f"  ✗ {module} [red](manquant)[/]")
    
    if missing:
        console.print(f"\n[red]❌ Dépendances manquantes : {', '.join(missing)}[/]")
        console.print("[yellow]Installer avec : pip install -r requirements.txt[/]")
        return False
    
    console.print("[green]✅ Toutes les dépendances sont installées[/]")
    return True


def check_source_files():
    """Vérifier présence des fichiers sources"""
    console.print("\n[bold cyan]📂 Vérification des fichiers sources...[/]")
    
    required_files = [
        "data/raw/BP_FABRIQ_PRODUCT-OCT2025.xlsx",
        "data/raw/Business_Plan_GenieFactory-SEPT2025.docx",
        "data/raw/GENIE_FACTORY_PACTE_AATL-v3.docx"
    ]
    
    missing = []
    for filepath in required_files:
        path = Path(filepath)
        if path.exists():
            console.print(f"  ✓ {filepath}")
        else:
            missing.append(filepath)
            console.print(f"  ✗ {filepath} [red](manquant)[/]")
    
    if missing:
        console.print(f"\n[red]❌ Fichiers sources manquants[/]")
        console.print("[yellow]Placer les fichiers dans data/raw/[/]")
        return False
    
    console.print("[green]✅ Tous les fichiers sources présents[/]")
    return True


def run_script(script_info, args):
    """Exécuter un script Python"""
    script_path = Path(script_info['script'])
    
    # Check si skip demandé
    skip_flag = script_info.get('skip_flag')
    if skip_flag and getattr(args, skip_flag.replace('--skip-', ''), False):
        console.print(f"[yellow]⏭️  Skipping {script_info['name']}[/]")
        return True
    
    console.print(f"\n[bold cyan]▶️  {script_info['name']}[/]")
    console.print(f"[dim]{script_info['description']}[/]")
    
    if not script_path.exists():
        console.print(f"[red]❌ Script non trouvé : {script_path}[/]")
        return False
    
    try:
        # Exécuter le script
        result = subprocess.run(
            [sys.executable, str(script_path)],
            check=True,
            capture_output=True,
            text=True
        )
        
        # Afficher output si verbose
        if args.verbose and result.stdout:
            console.print(result.stdout)
        
        console.print(f"[green]✅ {script_info['name']} terminé avec succès[/]")
        return True
        
    except subprocess.CalledProcessError as e:
        console.print(f"[red]❌ Erreur dans {script_info['name']}[/]")
        console.print(f"[red]{e.stderr}[/]")
        return False


def main():
    """Fonction principale"""
    parser = argparse.ArgumentParser(
        description="GenieFactory BP 14 Mois - Génération complète"
    )
    parser.add_argument(
        '--skip-extract',
        action='store_true',
        help="Skip extraction (si déjà exécutée)"
    )
    parser.add_argument(
        '--skip-assumptions',
        action='store_true',
        help="Skip génération assumptions (si déjà validé)"
    )
    parser.add_argument(
        '--validate-only',
        action='store_true',
        help="Exécuter seulement la validation"
    )
    parser.add_argument(
        '--verbose',
        '-v',
        action='store_true',
        help="Afficher output détaillé"
    )
    
    args = parser.parse_args()
    
    # Header
    console.print(Panel.fit(
        "[bold cyan]GenieFactory - Business Plan 14 Mois[/]\n"
        "[dim]Génération automatisée Nov 2025 → Dec 2026[/]",
        border_style="cyan"
    ))
    
    start_time = datetime.now()
    
    # Checks préliminaires
    if not check_dependencies():
        sys.exit(1)
    
    if not check_source_files():
        sys.exit(1)
    
    # Mode validation uniquement
    if args.validate_only:
        console.print("\n[bold yellow]⚡ Mode validation uniquement[/]")
        validation_script = next(s for s in SCRIPTS if '6.' in s['name'])
        success = run_script(validation_script, args)
        sys.exit(0 if success else 1)
    
    # Exécution séquentielle
    console.print("\n[bold cyan]🚀 Démarrage génération BP...[/]")
    
    success_count = 0
    for script_info in SCRIPTS:
        if run_script(script_info, args):
            success_count += 1
        else:
            console.print(f"\n[red]❌ Échec à l'étape {script_info['name']}[/]")
            console.print("[yellow]Vérifier les logs pour détails[/]")
            sys.exit(1)
    
    # Résumé final
    elapsed = datetime.now() - start_time
    
    console.print("\n" + "="*60)
    console.print(Panel.fit(
        f"[bold green]✅ Génération BP terminée avec succès ![/]\n\n"
        f"[cyan]📊 Livrables générés :[/]\n"
        f"  • data/structured/assumptions.yaml\n"
        f"  • data/structured/projections.json\n"
        f"  • data/outputs/BP_14M_Nov2025-Dec2026.xlsx\n"
        f"  • data/outputs/BM_Updated_14M.docx\n\n"
        f"[cyan]⏱️  Durée totale : {elapsed.total_seconds():.1f}s[/]\n"
        f"[cyan]✓ Scripts exécutés : {success_count}/{len(SCRIPTS)}[/]",
        border_style="green"
    ))
    
    # Prochaines étapes
    console.print("\n[bold cyan]📋 Prochaines étapes :[/]")
    console.print("  1. Vérifier data/outputs/BP_14M_Nov2025-Dec2026.xlsx")
    console.print("  2. Review data/outputs/BM_Updated_14M.docx")
    console.print("  3. Ajuster assumptions.yaml si nécessaire")
    console.print("  4. Regénérer : python run.py")
    console.print("\n[dim]Logs détaillés : logs/run_YYYYMMDD_HHMMSS.log[/]")


if __name__ == "__main__":
    main()
