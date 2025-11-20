#!/usr/bin/env python3
"""
GenieFactory BP 14 Mois - Script 3: Calcul Projections
Calcule ARR, CA, charges, EBITDA pour chaque mois M1-M14

Input:
  - data/structured/assumptions.yaml

Output:
  - data/structured/projections.json
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
    """Calculateur de projections financières"""

    def __init__(self, assumptions: Dict[str, Any]):
        self.assumptions = assumptions
        self.months_data = []

    def get_month_date(self, month_index: int) -> str:
        """Calculer la date d'un mois (M1 = Nov 2025)"""
        start_date = datetime(2025, 11, 1)  # Nov 2025
        month_date = start_date + relativedelta(months=month_index - 1)
        return month_date.strftime('%Y-%m')

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
        new_customers = hub_config['new_customers_monthly'].get(f'm{month}', 0)
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
        """Calculer coûts personnel"""
        team_size = self.get_team_size(month)
        salaries_config = self.assumptions['costs']['personnel']['salaries']

        # Salaires (fondateurs = 0 pour l'instant)
        nb_employees = team_size - 4  # 4 fondateurs
        salary_cost = max(0, nb_employees) * salaries_config['employee_monthly']

        # Freelance
        freelance = self.assumptions['costs']['personnel']['freelance_monthly_budget']

        total = salary_cost + freelance

        return {
            'team_size': team_size,
            'employees': max(0, nb_employees),
            'salary_cost': salary_cost,
            'freelance': freelance,
            'total': total
        }

    def calculate_infrastructure_costs(self, month: int, nb_clients: float) -> float:
        """Calculer coûts infrastructure"""
        infra_config = self.assumptions['costs']['infrastructure']

        base = infra_config['base_monthly']
        per_client = nb_clients * infra_config['per_client_monthly']

        # Tools
        tools = sum(infra_config['tools_monthly'].values())

        total = base + per_client + tools

        return total

    def calculate_marketing_costs(self, month: int) -> float:
        """Calculer coûts marketing"""
        marketing_config = self.assumptions['costs']['marketing']

        base = marketing_config['base_monthly']
        content = marketing_config['content_monthly']

        # Events trimestriels (M3, M6, M9, M12)
        events = 0
        if month in [3, 6, 9, 12]:
            events = marketing_config['events_quarterly']

        total = base + content + events

        return total

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
        infrastructure_cost = self.calculate_infrastructure_costs(month, nb_clients)
        marketing_cost = self.calculate_marketing_costs(month)
        admin_cost = self.assumptions['costs']['office_admin']['monthly']

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
            'revenue': {
                'hackathon': hackathon_data,
                'factory': factory_data,
                'enterprise_hub': hub_data,
                'services': services_data,
                'total': total_revenue
            },
            'costs': {
                'personnel': personnel_data,
                'infrastructure': infrastructure_cost,
                'marketing': marketing_cost,
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
        """Calculer projections pour tous les mois M1-M14"""
        logger.info("\n🔢 CALCUL PROJECTIONS M1-M14")
        logger.info("="*60)

        for month in range(1, 15):
            month_data = self.calculate_month(month)
            self.months_data.append(month_data)

        return self.months_data


def main():
    """Fonction principale"""
    logger.info("="*60)
    logger.info("🚀 CALCUL PROJECTIONS - GenieFactory BP 14 Mois")
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

    logger.info(f"✓ Assumptions chargées")
    logger.info(f"  • ARR target M14: {assumptions['financial_kpis']['target_arr_dec_2026']:,}€")
    logger.info(f"  • Durée: {assumptions['timeline']['duration_months']} mois")

    # Calculer projections
    calculator = ProjectionCalculator(assumptions)
    projections = calculator.calculate_all_months()

    # Sauvegarder
    output_path = base_path / "data" / "structured" / "projections.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(projections, f, indent=2, ensure_ascii=False)

    logger.info("\n" + "="*60)
    logger.info("✅ PROJECTIONS CALCULÉES")
    logger.info("="*60)
    logger.info(f"📁 Fichier généré: {output_path}")

    # Résumé
    logger.info("\n📊 RÉSUMÉ 14 MOIS:")

    m1 = projections[0]
    m11 = projections[10]
    m14 = projections[13]

    logger.info(f"\n  M1 (Nov 2025):")
    logger.info(f"    • CA: {m1['revenue']['total']:,.0f}€")
    logger.info(f"    • ARR: {m1['metrics']['arr']:,.0f}€")
    logger.info(f"    • Équipe: {m1['metrics']['team_size']} ETP")

    logger.info(f"\n  M11 (Sept 2026) - Avant Seed:")
    logger.info(f"    • CA: {m11['revenue']['total']:,.0f}€")
    logger.info(f"    • ARR: {m11['metrics']['arr']:,.0f}€ {'✓' if m11['metrics']['arr'] >= 400000 else '⚠️'}")
    logger.info(f"    • Équipe: {m11['metrics']['team_size']} ETP")
    logger.info(f"    • Cash: {m11['metrics']['cash']:,.0f}€")

    logger.info(f"\n  M14 (Dec 2026) - TARGET:")
    logger.info(f"    • CA: {m14['revenue']['total']:,.0f}€")
    logger.info(f"    • ARR: {m14['metrics']['arr']:,.0f}€")
    logger.info(f"    • Équipe: {m14['metrics']['team_size']} ETP")
    logger.info(f"    • Cash: {m14['metrics']['cash']:,.0f}€")

    # Vérifications rapides
    logger.info("\n🔍 CHECKS RAPIDES:")

    arr_m14 = m14['metrics']['arr']
    target = assumptions['financial_kpis']['target_arr_dec_2026']
    arr_ok = (target * 0.9) <= arr_m14 <= (target * 1.1)

    logger.info(f"  {'✓' if arr_ok else '✗'} ARR M14: {arr_m14:,.0f}€ (target {target:,.0f}€ ±10%)")

    arr_m11 = m11['metrics']['arr']
    arr_m11_ok = arr_m11 >= 400000
    logger.info(f"  {'✓' if arr_m11_ok else '✗'} ARR M11: {arr_m11:,.0f}€ (min 400,000€)")

    cash_ok = all(m['metrics']['cash'] >= 0 for m in projections)
    logger.info(f"  {'✓' if cash_ok else '✗'} Cash position: {'Toujours positive' if cash_ok else 'NÉGATIF détecté!'}")

    max_burn = max(m['metrics']['burn_rate'] for m in projections)
    burn_ok = max_burn <= 60000
    logger.info(f"  {'✓' if burn_ok else '✗'} Burn rate max: {max_burn:,.0f}€/mois (limite 60,000€)")

    ca_total = sum(m['revenue']['total'] for m in projections)
    logger.info(f"\n  📈 CA total 14 mois: {ca_total:,.0f}€")

    if arr_ok and arr_m11_ok and cash_ok and burn_ok:
        logger.info("\n✅ Tous les checks rapides passent!")
        logger.info("   → Prêt pour Phase 4: python scripts/4_generate_bp_excel.py")
    else:
        logger.warning("\n⚠️ Certains checks échouent - ajuster assumptions.yaml")

    return 0


if __name__ == "__main__":
    exit(main())
