#!/usr/bin/env python3
"""
GenieFactory BP - Script 3: Calcul Projections (Étendu 50 mois)
Calcule ARR, CA, charges, EBITDA pour chaque mois M1-M50 (Nov 2025 - Dec 2029)

Input:
  - data/structured/assumptions.yaml

Output:
  - data/structured/projections_50m.json
"""

import json
import yaml
import logging
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Any, List
from dateutil.relativedelta import relativedelta

# Configuration logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


class ProjectionCalculator:
    """Calculateur de projections financières (50 mois)"""

    def __init__(self, assumptions: Dict[str, Any], months_count: int = 50):
        self.assumptions = assumptions
        self.months_data = []
        self.months_count = months_count

    def get_month_date(self, month_index: int) -> str:
        """Calculer la date d'un mois (M1 = Nov 2025)"""
        start_date = datetime(2025, 11, 1)  # Nov 2025
        month_date = start_date + relativedelta(months=month_index - 1)
        return month_date.strftime('%Y-%m')

    def get_year_for_month(self, month: int) -> int:
        """Déterminer l'année pour un mois donné
        M1-M14: 2025-2026
        M15-M26: 2027
        M27-M38: 2028
        M39-M50: 2029
        """
        if month <= 14:
            return 2025  # 2025-2026
        elif month <= 26:
            return 2027
        elif month <= 38:
            return 2028
        else:
            return 2029

    def get_hackathon_price(self, month: int) -> float:
        """Obtenir le prix hackathon pour un mois donné"""
        periods = self.assumptions['pricing']['hackathon']['periods']
        for period in periods:
            if period['start_month'] <= month <= period['end_month']:
                return period['price_eur']
        return periods[0]['price_eur']  # Défaut

    def get_factory_price(self, month: int) -> float:
        """Obtenir le prix factory pour un mois donné"""
        periods = self.assumptions['pricing']['factory']['periods']
        for period in periods:
            if period['start_month'] <= month <= period['end_month']:
                return period['price_eur']
        return periods[0]['price_eur']

    def get_services_price(self, month: int) -> float:
        """Obtenir le prix services pour un mois donné"""
        periods = self.assumptions['pricing']['services']['implementation']['periods']
        for period in periods:
            if period['start_month'] <= month <= period['end_month']:
                return period['price_eur']
        return periods[0]['price_eur']

    def get_team_size(self, month: int) -> int:
        """Obtenir la taille de l'équipe pour un mois donné"""
        evolution = self.assumptions['costs']['personnel']['team_evolution']
        # Trouver la dernière valeur <= month
        team_size = 5  # Défaut
        for m in range(1, month + 1):
            key = f'm{m}'
            if key in evolution:
                team_size = evolution[key]
        return team_size

    def calculate_hackathon_revenue(self, month: int) -> Dict[str, Any]:
        """Calculer revenus hackathons"""
        volumes = self.assumptions['sales_assumptions']['hackathon']['volumes_monthly']
        nb_hackathons = volumes.get(f'm{month}', 0)
        price = self.get_hackathon_price(month)

        revenue = nb_hackathons * price

        return {
            'volume': nb_hackathons,
            'price_unit': price,
            'revenue': revenue
        }

    def calculate_factory_revenue(self, month: int) -> Dict[str, Any]:
        """Calculer revenus factory (conversion hackathons avec délai)"""
        conversion_rate = self.assumptions['sales_assumptions']['factory']['conversion_rate']
        delay = self.assumptions['sales_assumptions']['factory']['delay_months']

        # Factory M = conversion des hackathons (M - delay)
        source_month = month - delay
        if source_month < 1:
            return {'volume': 0, 'price_unit': 0, 'revenue': 0}

        # Hackathons source
        hackathon_volumes = self.assumptions['sales_assumptions']['hackathon']['volumes_monthly']
        source_hackathons = hackathon_volumes.get(f'm{source_month}', 0)

        # Conversion
        nb_factory = source_hackathons * conversion_rate
        price = self.get_factory_price(month)
        revenue = nb_factory * price

        return {
            'volume': nb_factory,
            'price_unit': price,
            'revenue': revenue,
            'source_month': source_month,
            'source_hackathons': source_hackathons
        }

    def calculate_hub_revenue(self, month: int) -> Dict[str, Any]:
        """Calculer revenus Enterprise Hub (MRR avec churn et upgrades)"""
        hub_config = self.assumptions['sales_assumptions']['enterprise_hub']
        launch_month = self.assumptions['pricing']['enterprise_hub']['launch_month']

        if month < launch_month:
            return {
                'customers': {'starter': 0, 'business': 0, 'enterprise': 0, 'total': 0},
                'mrr': 0,
                'arr': 0
            }

        # Récupérer état précédent
        customers_starter = 0
        customers_business = 0
        customers_enterprise = 0

        if month > launch_month and len(self.months_data) > 0:
            prev_hub = self.months_data[-1]['revenue']['enterprise_hub']
            customers_starter = prev_hub['customers']['starter']
            customers_business = prev_hub['customers']['business']
            customers_enterprise = prev_hub['customers']['enterprise']

        # Nouveaux clients ce mois
        # Pour M1-M14: utiliser volumes définis
        # Pour M15+: utiliser long_term_projections
        new_customers = hub_config['new_customers_monthly'].get(f'm{month}', None)

        if new_customers is None and month > 14:
            # Extension long terme
            if 'long_term_projections' in self.assumptions:
                year = self.get_year_for_month(month)
                lt_proj = self.assumptions['long_term_projections']['years'].get(str(year), {})
                new_customers = lt_proj.get('new_customers_hub_monthly', 8)  # Défaut 8/mois
            else:
                # Fallback: croissance linéaire basée sur M14
                new_customers = 8  # Conservateur
        elif new_customers is None:
            new_customers = 0

        tier_dist = hub_config['tier_distribution_at_launch']

        new_starter = new_customers * tier_dist['starter']
        new_business = new_customers * tier_dist['business']
        new_enterprise = new_customers * tier_dist['enterprise']

        # Churn
        churn_rate = hub_config['churn_monthly']
        customers_starter = customers_starter * (1 - churn_rate) + new_starter
        customers_business = customers_business * (1 - churn_rate) + new_business
        customers_enterprise = customers_enterprise * (1 - churn_rate) + new_enterprise

        # Upgrades (simplifié - seulement après 3 mois)
        if month >= launch_month + 3:
            upgrade_rate = hub_config['upgrade_patterns']['starter_to_business_rate']
            upgrades = customers_starter * upgrade_rate * 0.1  # 10% par mois des eligibles
            customers_starter -= upgrades
            customers_business += upgrades

        # MRR
        tiers = self.assumptions['pricing']['enterprise_hub']['tiers']
        mrr = (
            customers_starter * tiers['starter']['monthly_eur'] +
            customers_business * tiers['business']['monthly_eur'] +
            customers_enterprise * tiers['enterprise']['monthly_eur']
        )

        arr = mrr * 12

        return {
            'customers': {
                'starter': round(customers_starter, 2),
                'business': round(customers_business, 2),
                'enterprise': round(customers_enterprise, 2),
                'total': round(customers_starter + customers_business + customers_enterprise, 2)
            },
            'new_customers': new_customers,
            'mrr': mrr,
            'arr': arr
        }

    def calculate_services_revenue(self, month: int, nb_hackathons: float, nb_factory: float) -> Dict[str, Any]:
        """Calculer revenus services (proportionnels aux hackathons/factory)"""
        # Services implementation liés aux projets
        # Simplifié: 50% des hackathons + 20% des factory génèrent services

        services_from_hack = nb_hackathons * 0.5
        services_from_factory = nb_factory * 0.2

        total_services = services_from_hack + services_from_factory
        price = self.get_services_price(month)

        revenue = total_services * price

        return {
            'volume': total_services,
            'price_unit': price,
            'revenue': revenue
        }

    def calculate_personnel_costs(self, month: int) -> Dict[str, Any]:
        """Calculer coûts personnel (détaillé par rôle)"""
        team_size = self.get_team_size(month)
        year = self.get_year_for_month(month)

        # Vérifier si personnel_details existe
        if 'personnel_details' not in self.assumptions:
            # Fallback sur ancien mode
            salaries_config = self.assumptions['costs']['personnel']['salaries']
            nb_employees = team_size - 4  # 4 fondateurs
            salary_cost = max(0, nb_employees) * salaries_config['employee_monthly']
            freelance = self.assumptions['costs']['personnel']['freelance_monthly_budget']
            total = salary_cost + freelance

            return {
                'team_size': team_size,
                'employees': max(0, nb_employees),
                'salary_cost': salary_cost,
                'freelance': freelance,
                'total': total
            }

        # Mode détaillé avec personnel_details
        personnel_details = self.assumptions['personnel_details']
        charges_sociales_rate = personnel_details['charges_sociales_rate']

        # Déterminer la clé de timeline selon l'année
        if year == 2025:
            timeline_key = 'm1_m14'
        else:
            timeline_key = f'y{year}'

        total_brut = 0
        total_fte = 0
        roles_detail = {}

        for role_name, role_data in personnel_details['roles'].items():
            fte = role_data['fte_timeline'].get(timeline_key, 0)
            if fte > 0:
                salary_annual = role_data['salary_brut_annual']
                salary_monthly = salary_annual / 12
                cost_monthly = salary_monthly * fte

                total_brut += cost_monthly
                total_fte += fte

                roles_detail[role_name] = {
                    'fte': fte,
                    'salary_monthly': salary_monthly,
                    'cost_monthly': cost_monthly
                }

        # Charges sociales
        charges_sociales = total_brut * charges_sociales_rate

        # Freelance
        freelance = self.assumptions['costs']['personnel'].get('freelance_monthly_budget', 0)

        total = total_brut + charges_sociales + freelance

        return {
            'team_size': round(total_fte),
            'fte_total': total_fte,
            'salary_brut': total_brut,
            'charges_sociales': charges_sociales,
            'freelance': freelance,
            'total': total,
            'roles': roles_detail
        }

    def calculate_infrastructure_costs(self, month: int, nb_clients: float, team_size: int) -> Dict[str, Any]:
        """Calculer coûts infrastructure (détaillé avec scaling)"""
        # Vérifier si infrastructure_costs existe (nouveau format)
        if 'infrastructure_costs' in self.assumptions:
            infra_new = self.assumptions['infrastructure_costs']

            # Cloud avec tiers de scaling
            cloud_base = infra_new['cloud']['base_monthly']
            cost_per_client = infra_new['cloud']['cost_per_client']

            # Tiers de scaling: moins cher avec plus de clients
            scaling_tiers = infra_new['cloud'].get('scaling_tiers', {})
            if nb_clients > 100:
                cost_per_client = scaling_tiers.get('tier3', {}).get('cost_per_client', 30)  # 30€
            elif nb_clients > 50:
                cost_per_client = scaling_tiers.get('tier2', {}).get('cost_per_client', 40)  # 40€
            else:
                cost_per_client = scaling_tiers.get('tier1', {}).get('cost_per_client', 50)  # 50€

            cloud_total = cloud_base + (nb_clients * cost_per_client)

            # SaaS tools par user/developer
            saas_tools = infra_new['saas_tools']
            notion_cost = max(team_size, saas_tools['notion'].get('min_users', 5)) * saas_tools['notion']['cost_per_user']
            slack_cost = max(team_size, saas_tools['slack'].get('min_users', 5)) * saas_tools['slack']['cost_per_user']
            nb_devs = max(team_size // 2, saas_tools['github'].get('min_users', 2))
            github_cost = nb_devs * saas_tools['github']['cost_per_developer']
            analytics_cost = saas_tools.get('analytics', {}).get('monthly_flat', 0)
            crm_users = max(team_size // 4, saas_tools['crm'].get('min_users', 2))  # ~25% uses CRM
            crm_cost = crm_users * saas_tools['crm']['cost_per_user']

            saas_total = notion_cost + slack_cost + github_cost + analytics_cost + crm_cost

            # R&D externe (optionnel)
            rd_external = infra_new.get('rd_external', {}).get('monthly_budget', 0)

            total = cloud_total + saas_total + rd_external

            return {
                'cloud': cloud_total,
                'saas_tools': saas_total,
                'rd_external': rd_external,
                'total': total
            }

        # Fallback sur ancien format
        infra_config = self.assumptions['costs']['infrastructure']
        base = infra_config['base_monthly']
        per_client = nb_clients * infra_config['per_client_monthly']
        tools = sum(infra_config['tools_monthly'].values())
        total = base + per_client + tools

        return {
            'total': total
        }

    def calculate_marketing_costs(self, month: int) -> Dict[str, Any]:
        """Calculer coûts marketing (détaillé par canal)"""
        year = self.get_year_for_month(month)

        # Vérifier si marketing_budgets existe (nouveau format)
        if 'marketing_budgets' in self.assumptions:
            marketing_new = self.assumptions['marketing_budgets']

            # Digital ads selon l'année
            digital_budget = marketing_new['digital_ads']['monthly_budgets'].get(f'y{year}', 2000)

            # Events selon l'année (pas tous les mois - trimestriels)
            events_monthly = marketing_new['events']['monthly_budgets'].get(f'y{year}', 1000)
            events_budget = 0
            if month % 3 == 0:  # Trimestres (M3, M6, M9, M12, etc.)
                events_budget = events_monthly * 3  # Budget trimestriel

            # Content selon l'année
            content_budget = marketing_new['content']['monthly_budgets'].get(f'y{year}', 1000)

            # Partnerships selon l'année
            partnerships_budget = marketing_new['partnerships']['monthly_budgets'].get(f'y{year}', 500)

            total = digital_budget + events_budget + content_budget + partnerships_budget

            return {
                'digital_ads': digital_budget,
                'events': events_budget,
                'content': content_budget,
                'partnerships': partnerships_budget,
                'total': total
            }

        # Fallback sur ancien format
        marketing_config = self.assumptions['costs']['marketing']
        base = marketing_config['base_monthly']
        content = marketing_config['content_monthly']
        events = 0
        if month in [3, 6, 9, 12]:
            events = marketing_config['events_quarterly']
        total = base + content + events

        return {
            'total': total
        }

    def calculate_month(self, month: int) -> Dict[str, Any]:
        """Calculer toutes les métriques pour un mois"""
        logger.info(f"📊 Calcul M{month} ({self.get_month_date(month)})...")

        # REVENUS
        hackathon_data = self.calculate_hackathon_revenue(month)
        factory_data = self.calculate_factory_revenue(month)
        hub_data = self.calculate_hub_revenue(month)
        services_data = self.calculate_services_revenue(
            month,
            hackathon_data['volume'],
            factory_data['volume']
        )

        total_revenue = (
            hackathon_data['revenue'] +
            factory_data['revenue'] +
            hub_data['mrr'] +
            services_data['revenue']
        )

        # COÛTS
        personnel_data = self.calculate_personnel_costs(month)
        nb_clients = hub_data['customers']['total']
        team_size = personnel_data['team_size']
        infrastructure_data = self.calculate_infrastructure_costs(month, nb_clients, team_size)
        marketing_data = self.calculate_marketing_costs(month)
        admin_cost = self.assumptions['costs']['office_admin']['monthly']

        # Extraire total (compatible ancien et nouveau format)
        infrastructure_cost = infrastructure_data if isinstance(infrastructure_data, (int, float)) else infrastructure_data['total']
        marketing_cost = marketing_data if isinstance(marketing_data, (int, float)) else marketing_data['total']

        total_costs = (
            personnel_data['total'] +
            infrastructure_cost +
            marketing_cost +
            admin_cost
        )

        # MÉTRIQUES
        ebitda = total_revenue - total_costs
        burn_rate = -ebitda if ebitda < 0 else 0

        # Cash position (avec fundings)
        fundings = self.assumptions['financial_kpis']['cash_management']['fundings']
        funding_this_month = 0
        for funding in fundings:
            if funding['month'] == month:
                funding_this_month = funding['amount']

        # Cash cumulé
        prev_cash = 0
        if len(self.months_data) > 0:
            prev_cash = self.months_data[-1]['metrics']['cash']

        cash = prev_cash + total_revenue - total_costs + funding_this_month

        month_data = {
            'month': month,
            'date': self.get_month_date(month),
            'year': self.get_year_for_month(month),
            'revenue': {
                'hackathon': hackathon_data,
                'factory': factory_data,
                'enterprise_hub': hub_data,
                'services': services_data,
                'total': total_revenue
            },
            'costs': {
                'personnel': personnel_data,
                'infrastructure': infrastructure_data if isinstance(infrastructure_data, dict) else {'total': infrastructure_data},
                'marketing': marketing_data if isinstance(marketing_data, dict) else {'total': marketing_data},
                'admin': admin_cost,
                'total': total_costs
            },
            'metrics': {
                'ebitda': ebitda,
                'burn_rate': burn_rate,
                'arr': hub_data['arr'],
                'mrr': hub_data['mrr'],
                'cash': cash,
                'funding': funding_this_month,
                'team_size': personnel_data['team_size']
            }
        }

        # Log résumé
        logger.info(
            f"  CA: {total_revenue:,.0f}€ | "
            f"Charges: {total_costs:,.0f}€ | "
            f"EBITDA: {ebitda:,.0f}€ | "
            f"ARR: {hub_data['arr']:,.0f}€ | "
            f"Cash: {cash:,.0f}€"
        )

        return month_data

    def calculate_all_months(self) -> List[Dict[str, Any]]:
        """Calculer projections pour tous les mois (M1-M50 par défaut)"""
        logger.info(f"\n🔢 CALCUL PROJECTIONS M1-M{self.months_count}")
        logger.info("="*60)

        for month in range(1, self.months_count + 1):
            month_data = self.calculate_month(month)
            self.months_data.append(month_data)

            # Log milestones
            if month in [14, 26, 38, 50]:
                year_label = {14: "2026", 26: "2027", 38: "2028", 50: "2029"}[month]
                logger.info(f"\n  🎯 MILESTONE M{month} (Dec {year_label}):")
                logger.info(f"    ARR: {month_data['metrics']['arr']:,.0f}€ | Cash: {month_data['metrics']['cash']:,.0f}€")

        return self.months_data


def main():
    """Fonction principale"""
    logger.info("="*60)
    logger.info("🚀 CALCUL PROJECTIONS - GenieFactory BP 50 Mois (Nov 2025 - Dec 2029)")
    logger.info("="*60)

    base_path = Path(__file__).parent.parent

    # Charger assumptions
    assumptions_path = base_path / "data" / "structured" / "assumptions.yaml"
    if not assumptions_path.exists():
        logger.error(f"❌ Fichier assumptions.yaml non trouvé: {assumptions_path}")
        logger.error("   Exécuter d'abord: python scripts/2_generate_assumptions.py")
        return 1

    logger.info(f"📂 Chargement assumptions: {assumptions_path}")
    with open(assumptions_path, 'r', encoding='utf-8') as f:
        assumptions = yaml.safe_load(f)

    logger.info(f"✓ Assumptions chargées (version {assumptions.get('version', '1.0')})")
    logger.info(f"  • ARR target M14: {assumptions['financial_kpis']['target_arr_dec_2026']:,}€")
    logger.info(f"  • Période: Nov 2025 - Dec 2029 (50 mois)")
    if 'long_term_projections' in assumptions:
        logger.info(f"  • Extensions long terme détectées (2027-2029)")

    # Calculer projections 50 mois
    calculator = ProjectionCalculator(assumptions, months_count=50)
    projections = calculator.calculate_all_months()

    # Sauvegarder
    output_path = base_path / "data" / "structured" / "projections_50m.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(projections, f, indent=2, ensure_ascii=False)

    logger.info("\n" + "="*60)
    logger.info("✅ PROJECTIONS 50 MOIS CALCULÉES")
    logger.info("="*60)
    logger.info(f"📁 Fichier généré: {output_path}")
    logger.info(f"📊 {len(projections)} mois de projections")

    # Résumé - Milestones clés
    logger.info("\n📊 RÉSUMÉ MILESTONES:")

    m1 = projections[0]
    m11 = projections[10]
    m14 = projections[13]
    m26 = projections[25]
    m38 = projections[37]
    m50 = projections[49]

    logger.info(f"\n  M1 (Nov 2025) - Lancement:")
    logger.info(f"    • CA: {m1['revenue']['total']:,.0f}€")
    logger.info(f"    • ARR: {m1['metrics']['arr']:,.0f}€")
    logger.info(f"    • Équipe: {m1['metrics']['team_size']} ETP")

    logger.info(f"\n  M11 (Sept 2026) - Avant Seed:")
    logger.info(f"    • CA: {m11['revenue']['total']:,.0f}€")
    logger.info(f"    • ARR: {m11['metrics']['arr']:,.0f}€")
    logger.info(f"    • Équipe: {m11['metrics']['team_size']} ETP")
    logger.info(f"    • Cash: {m11['metrics']['cash']:,.0f}€")

    logger.info(f"\n  M14 (Dec 2026) - Fin année 1:")
    logger.info(f"    • CA: {m14['revenue']['total']:,.0f}€")
    logger.info(f"    • ARR: {m14['metrics']['arr']:,.0f}€")
    logger.info(f"    • Équipe: {m14['metrics']['team_size']} ETP")
    logger.info(f"    • Cash: {m14['metrics']['cash']:,.0f}€")

    logger.info(f"\n  M26 (Dec 2027) - Fin année 2:")
    logger.info(f"    • CA: {m26['revenue']['total']:,.0f}€")
    logger.info(f"    • ARR: {m26['metrics']['arr']:,.0f}€")
    logger.info(f"    • Équipe: {m26['metrics']['team_size']} ETP")
    logger.info(f"    • Cash: {m26['metrics']['cash']:,.0f}€")

    logger.info(f"\n  M38 (Dec 2028) - Fin année 3:")
    logger.info(f"    • CA: {m38['revenue']['total']:,.0f}€")
    logger.info(f"    • ARR: {m38['metrics']['arr']:,.0f}€")
    logger.info(f"    • Équipe: {m38['metrics']['team_size']} ETP")
    logger.info(f"    • Cash: {m38['metrics']['cash']:,.0f}€")

    logger.info(f"\n  M50 (Dec 2029) - Fin année 4:")
    logger.info(f"    • CA: {m50['revenue']['total']:,.0f}€")
    logger.info(f"    • ARR: {m50['metrics']['arr']:,.0f}€")
    logger.info(f"    • Équipe: {m50['metrics']['team_size']} ETP")
    logger.info(f"    • Cash: {m50['metrics']['cash']:,.0f}€")

    # Vérifications rapides
    logger.info("\n🔍 CHECKS RAPIDES:")

    arr_m14 = m14['metrics']['arr']
    target = assumptions['financial_kpis']['target_arr_dec_2026']
    arr_ok = (target * 0.9) <= arr_m14 <= (target * 1.1)

    logger.info(f"  {'✓' if arr_ok else '✗'} ARR M14: {arr_m14:,.0f}€ (target {target:,.0f}€ ±10%)")

    cash_ok = all(m['metrics']['cash'] >= 0 for m in projections)
    logger.info(f"  {'✓' if cash_ok else '✗'} Cash position: {'Toujours positive' if cash_ok else 'NÉGATIF détecté!'}")

    max_burn = max(m['metrics']['burn_rate'] for m in projections)
    logger.info(f"  ℹ️  Burn rate max: {max_burn:,.0f}€/mois")

    ca_total_14 = sum(m['revenue']['total'] for m in projections[:14])
    ca_total_50 = sum(m['revenue']['total'] for m in projections)
    logger.info(f"\n  📈 CA cumulé:")
    logger.info(f"    • 14 mois (2025-2026): {ca_total_14:,.0f}€")
    logger.info(f"    • 50 mois (2025-2029): {ca_total_50:,.0f}€")

    if arr_ok and cash_ok:
        logger.info("\n✅ Tous les checks essentiels passent!")
        logger.info("   → Prêt pour Phase suivante: génération Excel 50M")
    else:
        logger.warning("\n⚠️ Certains checks échouent - ajuster assumptions.yaml")

    return 0


if __name__ == "__main__":
    exit(main())
