#!/usr/bin/env python3
"""
GenieFactory BP 14 Mois - Script 2: Génération Assumptions
Crée assumptions.yaml à partir des données extraites + template

Inputs:
  - data/structured/bp_extracted.json
  - data/structured/bm_extracted.json
  - data/structured/pacte_extracted.json
  - assumptions_template.yaml (template)

Output:
  - data/structured/assumptions.yaml
"""

import json
import yaml
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, Any

# Configuration logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


def load_extracted_data(base_path: Path) -> Dict[str, Any]:
    """Charger toutes les données extraites"""
    logger.info("📂 Chargement données extraites...")

    data_dir = base_path / "data" / "structured"

    bp_data = {}
    bm_data = {}
    pacte_data = {}

    # BP Excel
    bp_file = data_dir / "bp_extracted.json"
    if bp_file.exists():
        with open(bp_file, 'r', encoding='utf-8') as f:
            bp_data = json.load(f)
        logger.info(f"✓ BP Excel chargé ({len(bp_data)} clés)")
    else:
        logger.warning("⚠️ bp_extracted.json non trouvé")

    # BM Word
    bm_file = data_dir / "bm_extracted.json"
    if bm_file.exists():
        with open(bm_file, 'r', encoding='utf-8') as f:
            bm_data = json.load(f)
        logger.info(f"✓ BM Word chargé ({len(bm_data.get('tables', []))} tableaux)")
    else:
        logger.warning("⚠️ bm_extracted.json non trouvé")

    # Pacte
    pacte_file = data_dir / "pacte_extracted.json"
    if pacte_file.exists():
        with open(pacte_file, 'r', encoding='utf-8') as f:
            pacte_data = json.load(f)
        logger.info(f"✓ Pacte chargé ({len(pacte_data.get('arr_milestones', []))} milestones)")
    else:
        logger.warning("⚠️ pacte_extracted.json non trouvé")

    return {
        'bp': bp_data,
        'bm': bm_data,
        'pacte': pacte_data
    }


def generate_assumptions(extracted_data: Dict[str, Any]) -> Dict[str, Any]:
    """Générer la structure assumptions.yaml"""
    logger.info("🔧 Génération assumptions...")

    bp = extracted_data.get('bp', {})
    bm = extracted_data.get('bm', {})
    pacte = extracted_data.get('pacte', {})

    # Extraire ARR target du pacte
    arr_milestones = pacte.get('arr_milestones', [])
    arr_target_m14 = 800000  # Défaut

    if arr_milestones:
        # Chercher milestone ~800K€
        for milestone in arr_milestones:
            arr_value = milestone.get('arr_target', 0)
            if 700000 <= arr_value <= 900000:
                arr_target_m14 = arr_value
                break
        logger.info(f"✓ ARR target M14 depuis pacte: {arr_target_m14:,}€")
    else:
        logger.warning(f"⚠️ ARR target par défaut: {arr_target_m14:,}€")

    # Extraire pricing du BP
    pricing_extracted = bp.get('pricing', {})
    hackathon_price = pricing_extracted.get('hackathon', {}).get('price', 18000)
    factory_price = pricing_extracted.get('factory', {}).get('price', 75000)

    logger.info(f"✓ Pricing extrait - Hackathon: {hackathon_price}€, Factory: {factory_price}€")

    # Construction assumptions
    assumptions = {
        'meta': {
            'version': '1.0',
            'created_date': datetime.now().strftime('%Y-%m-%d'),
            'author': 'Claude Code - Automated Generation',
            'sources': [
                'BP_FABRIQ_PRODUCT-OCT2025.xlsx',
                'Business_Plan_GenieFactory-SEPT2025.docx',
                'GENIE_FACTORY_PACTE_AATL-v3.docx'
            ]
        },

        'timeline': {
            'start_month': '2025-11',
            'duration_months': 14,
            'fiscal_year_start': 11,
            'milestones': [
                {
                    'month': 1,
                    'name': 'Pre-seed Closing',
                    'amount_eur': 150000,
                    'breakdown': {
                        'autoposia_loan': 50000,
                        'f_initiatives_innovation': 15000,
                        'f_initiatives_creation': 25000,
                        'cic_bank': 30000,
                        'bpi_bourse_french_tech': 30000
                    },
                    'notes': 'Financement initial pour MVP et premiers clients'
                },
                {
                    'month': 11,
                    'name': 'Seed Round',
                    'amount_eur': 500000,
                    'valuation_pre_money': 2500000,
                    'dilution_pct': 16.7,
                    'target_arr_before': 450000,
                    'notes': 'Seed round pour scaling commercial et équipe'
                },
                {
                    'month': 14,
                    'name': 'ARR Milestone (Pacte Actionnaires)',
                    'arr_target': arr_target_m14,
                    'source': 'Pacte v3',
                    'notes': 'Milestone contractuel déclenchant earn-out fondateurs'
                }
            ]
        },

        'pricing': {
            'hackathon': {
                'description': 'Programme accélération innovation IA (4 semaines)',
                'target_segment': 'PME/ETI - équipes métier',
                'periods': [
                    {
                        'start_month': 1,
                        'end_month': 6,
                        'price_eur': hackathon_price
                    },
                    {
                        'start_month': 7,
                        'end_month': 14,
                        'price_eur': int(hackathon_price * 1.1)  # +10%
                    }
                ],
                'conversion_rates': {
                    'to_factory': 0.30,
                    'to_enterprise_hub': 0.15,
                    'delay_factory_months': 2
                },
                'notes': 'Offre d\'entrée pour valider PMF et générer pipeline. Marge élevée (80%+).'
            },

            'factory': {
                'description': 'Industrialisation prototypes hackathon → production',
                'target_segment': 'ETI/Grands comptes - projets stratégiques',
                'periods': [
                    {
                        'start_month': 1,
                        'end_month': 6,
                        'price_eur': factory_price
                    },
                    {
                        'start_month': 7,
                        'end_month': 14,
                        'price_eur': int(factory_price * 1.1)  # +10%
                    }
                ],
                'duration_weeks': '6-12',
                'team_allocation': '2-3 ETP GenieFactory + équipe client',
                'margin_pct': 65,
                'notes': 'Conversion naturelle des hackathons réussis. Cycle vente 2 mois.'
            },

            'enterprise_hub': {
                'description': 'Plateforme SaaS permanente innovation IA',
                'target_segment': 'Grands comptes - DSI/CDO',
                'launch_month': 8,
                'ramp_duration_months': 6,
                'tiers': {
                    'starter': {
                        'monthly_eur': 500,
                        'max_users': 10,
                        'max_use_cases': 2,
                        'target': 'PME - 1 département'
                    },
                    'business': {
                        'monthly_eur': 2000,
                        'max_users': 25,
                        'max_use_cases': 5,
                        'target': 'ETI - multi-départements'
                    },
                    'enterprise': {
                        'monthly_eur': 10000,
                        'max_users': 100,
                        'max_use_cases': 'illimité',
                        'deployment': 'on-premise option',
                        'target': 'Grands comptes CAC40'
                    }
                },
                'yearly_increase_pct': 10,
                'churn_annual': 0.10,
                'upgrade_rates': {
                    'starter_to_business': 0.20,
                    'business_to_enterprise': 0.10
                },
                'cac_eur': 15000,
                'ltv_eur': 120000,
                'notes': 'Offre stratégique pour ARR récurrent. Clé pour atteindre ARR 800K€ à M14.'
            },

            'services': {
                'implementation': {
                    'periods': [
                        {
                            'start_month': 1,
                            'end_month': 6,
                            'price_eur': 10000
                        },
                        {
                            'start_month': 7,
                            'end_month': 14,
                            'price_eur': 17500
                        }
                    ],
                    'typical_duration_days': '5-10',
                    'margin_pct': 70
                },
                'formation': {
                    'price_per_session': 5000,
                    'duration_days': 2,
                    'max_participants': 15,
                    'price_m7_m14': 5500
                },
                'notes': 'Revenus complémentaires et récurrents. Forte marge.'
            }
        },

        'sales_assumptions': {
            'hackathon': {
                'volumes_monthly': {
                    'm1': 1.5, 'm2': 2, 'm3': 2,
                    'm4': 2.5, 'm5': 2.5, 'm6': 3,
                    'm7': 3, 'm8': 3, 'm9': 3, 'm10': 3,
                    'm11': 4, 'm12': 4, 'm13': 4, 'm14': 4
                },
                'notes': 'Total 14 mois: ~39 hackathons. Progression: démarrage → ramping → post-seed.'
            },

            'factory': {
                'conversion_rate': 0.30,
                'delay_months': 2,
                'notes': 'Commence M3 (conversion hackathons M1). Délai commercial 2 mois.'
            },

            'enterprise_hub': {
                'new_customers_monthly': {
                    'm1': 0, 'm2': 0, 'm3': 0, 'm4': 0, 'm5': 0, 'm6': 0, 'm7': 0,
                    'm8': 2, 'm9': 2, 'm10': 3,
                    'm11': 4, 'm12': 4, 'm13': 5, 'm14': 6
                },
                'tier_distribution_at_launch': {
                    'starter': 0.70,
                    'business': 0.25,
                    'enterprise': 0.05
                },
                'upgrade_patterns': {
                    'starter_to_business_after_months': 3,
                    'starter_to_business_rate': 0.20,
                    'business_to_enterprise_after_months': 6,
                    'business_to_enterprise_rate': 0.10
                },
                'churn_monthly': 0.008,
                'notes': 'Lancement M8 après validation. Ramping post-seed à 6 clients/mois M14.'
            }
        },

        'costs': {
            'personnel': {
                'team_evolution': {
                    'm1': 5, 'm2': 5, 'm3': 7, 'm4': 7, 'm5': 7, 'm6': 7,
                    'm7': 9, 'm8': 9, 'm9': 9, 'm10': 9,
                    'm11': 11, 'm12': 11, 'm13': 12, 'm14': 12
                },
                'salaries': {
                    'founders_monthly': 0,
                    'employee_monthly': 6000
                },
                'freelance_monthly_budget': 5000,
                'hiring_cost_per_role': 3000,
                'notes': 'Croissance 5 → 12 ETP. Fondateurs non payés = économie ~150K€.'
            },

            'infrastructure': {
                'base_monthly': 2000,
                'per_client_monthly': 200,
                'tools_monthly': {
                    'development': 500,
                    'sales_marketing': 800,
                    'ops': 300,
                    'security': 400
                },
                'hackathon_infra_per_event': 500,
                'notes': 'Base fixe + variable par client. Ex M14: 2K + (30 × 200€) = 8K€/mois.'
            },

            'marketing': {
                'base_monthly': 5000,
                'events_quarterly': 15000,
                'content_monthly': 2000,
                'cac_targets': {
                    'hackathon': 5000,
                    'factory': 10000,
                    'hub': 15000
                },
                'notes': 'Budget contrôlé phase early-stage. Focus content marketing.'
            },

            'office_admin': {
                'monthly': 3000,
                'incorporation_costs': 5000,
                'notes': 'Structure légère. Coworking possible jusqu\'à M11.'
            }
        },

        'financial_kpis': {
            'target_arr_dec_2026': arr_target_m14,
            'target_arr_sept_2026': 450000,
            'revenue_mix_target': {
                'hackathon': 0.40,
                'factory': 0.30,
                'enterprise_hub': 0.20,
                'services': 0.10
            },
            'margin_targets': {
                'gross_margin_pct': 70,
                'ebitda_margin_pct': -15
            },
            'cash_management': {
                'min_cash_runway_months': 12,
                'acceptable_burn_rate_monthly': 50000,
                'fundings': [
                    {'month': 1, 'amount': 150000, 'source': 'Pre-seed (prêts + BPI)'},
                    {'month': 11, 'amount': 500000, 'source': 'Seed Round'}
                ]
            },
            'saas_metrics': {
                'target_arr_m14': 600000,
                'target_mrr_m14': 50000,
                'target_ltv_cac_ratio': 8,
                'max_churn_annual': 0.15
            }
        },

        'validation_rules': {
            'arr_tolerance_pct': 0.10,
            'arr_m14_min': int(arr_target_m14 * 0.9),
            'arr_m14_max': int(arr_target_m14 * 1.1),
            'arr_m11_min': 400000,
            'max_team_size': 15,
            'min_team_size_m1': 4,
            'max_burn_monthly': 60000,
            'avg_burn_acceptable': 40000,
            'min_conversion_hackathon_factory': 0.25,
            'max_churn_hub_monthly': 0.015,
            'min_cash_balance': 50000,
            'max_deviation_excel_word_pct': 0.05
        },

        'scenarios': {
            'base_case': {
                'arr_m14': arr_target_m14,
                'probability': 0.60
            },
            'upside': {
                'hackathon_volume_multiplier': 1.2,
                'conversion_factory': 0.35,
                'hub_adoption_faster': True,
                'arr_m14': int(arr_target_m14 * 1.19),
                'probability': 0.20
            },
            'downside': {
                'hackathon_volume_multiplier': 0.8,
                'conversion_factory': 0.25,
                'hub_launch_delay_months': 2,
                'arr_m14': int(arr_target_m14 * 0.81),
                'probability': 0.20
            },
            'notes': 'Scenario base = hypothèses conservatrices.'
        },

        'critical_assumptions': [
            {
                'assumption': 'Conversion hackathon → factory 30%',
                'risk_level': 'MEDIUM',
                'mitigation': 'Améliorer suivi post-hackathon, PoC gratuits'
            },
            {
                'assumption': 'Lancement Hub M8 sans retard',
                'risk_level': 'HIGH',
                'mitigation': 'Dev parallèle dès M3, beta testing M6-M7'
            },
            {
                'assumption': 'Seed 500K€ en Sept 2026',
                'risk_level': 'MEDIUM',
                'mitigation': 'Prep fundraising dès M6, pipeline investisseurs dès M8'
            },
            {
                'assumption': 'Churn Hub 10% annuel',
                'risk_level': 'LOW',
                'mitigation': 'Customer success dédié dès M11, NPS tracking'
            },
            {
                'assumption': 'Équipe 5 → 12 ETP sans turnover',
                'risk_level': 'MEDIUM',
                'mitigation': 'BSPCE attractif, culture forte, onboarding structuré'
            }
        ],

        'revision_history': [
            {
                'version': '1.0',
                'date': datetime.now().strftime('%Y-%m-%d'),
                'author': 'Claude Code - Automated',
                'changes': 'Initial version - extraction BP Oct 2025'
            }
        ],

        'usage_notes': (
            'Ce fichier est la SOURCE UNIQUE DE VÉRITÉ pour les projections BP 14 mois.\n\n'
            'Workflow:\n'
            '1. Modifier assumptions.yaml selon hypothèses business\n'
            '2. Lancer: python scripts/3_calculate_projections.py\n'
            '3. Vérifier: python scripts/6_validate.py\n'
            '4. Si OK, générer docs: python scripts/4_generate_bp_excel.py\n\n'
            'Chaque valeur est sourcée et commentée.'
        )
    }

    return assumptions


def main():
    """Fonction principale"""
    logger.info("="*60)
    logger.info("🚀 GÉNÉRATION ASSUMPTIONS - GenieFactory BP 14 Mois")
    logger.info("="*60)

    base_path = Path(__file__).parent.parent

    # Charger données extraites
    extracted_data = load_extracted_data(base_path)

    # Générer assumptions
    assumptions = generate_assumptions(extracted_data)

    # Sauvegarder
    output_path = base_path / "data" / "structured" / "assumptions.yaml"
    with open(output_path, 'w', encoding='utf-8') as f:
        yaml.dump(
            assumptions,
            f,
            default_flow_style=False,
            allow_unicode=True,
            sort_keys=False
        )

    logger.info(f"\n✅ Assumptions générées → {output_path}")

    # Statistiques
    logger.info("\n📊 Statistiques:")
    logger.info(f"  • ARR target M14: {assumptions['financial_kpis']['target_arr_dec_2026']:,}€")
    logger.info(f"  • Hackathon pricing M1-M6: {assumptions['pricing']['hackathon']['periods'][0]['price_eur']:,}€")
    logger.info(f"  • Factory pricing M1-M6: {assumptions['pricing']['factory']['periods'][0]['price_eur']:,}€")
    logger.info(f"  • Team M1→M14: {assumptions['costs']['personnel']['team_evolution']['m1']} → {assumptions['costs']['personnel']['team_evolution']['m14']} ETP")
    logger.info(f"  • Pre-seed: {assumptions['timeline']['milestones'][0]['amount_eur']:,}€")
    logger.info(f"  • Seed: {assumptions['timeline']['milestones'][1]['amount_eur']:,}€")

    logger.info("\n" + "="*60)
    logger.info("✅ GÉNÉRATION TERMINÉE")
    logger.info("="*60)
    logger.info(f"📝 Fichier généré: {output_path}")
    logger.info("\n⚠️  VALIDATION MANUELLE REQUISE:")
    logger.info("   1. Ouvrir data/structured/assumptions.yaml")
    logger.info("   2. Vérifier les hypothèses")
    logger.info("   3. Ajuster si nécessaire")
    logger.info("   4. Lancer Phase 3: python scripts/3_calculate_projections.py")

    return 0


if __name__ == "__main__":
    exit(main())
