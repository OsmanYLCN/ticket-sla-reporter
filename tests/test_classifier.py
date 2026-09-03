"""
Unit Tests for src.classifier
=============================
Tests deterministic ticket categorization, scope boundary conditions, and plant extraction.
"""

import unittest
from src.classifier import classify_ticket, get_plant_location


class TestClassifier(unittest.TestCase):

    def _base_ticket(self):
        """Returns a valid baseline In-Scope ticket template."""
        return {
            'BusinessLine': 'Infrastructure Services',
            'Severity': '1 - Molto alta',
            'Servizio': 'MANAGED NETWORK SERVICE',
            'Tipologia Calcolata': 'Incidente',
            'Oggetto': 'Core switch offline',
            'Configuration Item': 'SW-CORE-01',
            'Descrizione Articolo': 'Switch unreachability',
            'LocalitaCI': 'Kocaeli Plant',
            'Ubicazione sede Utente': 'Turkey Kocaeli'
        }

    # -------------------------------------------------------------------------
    # Rule 1 & Rule 2: BusinessLine & Severity Boundaries
    # -------------------------------------------------------------------------
    def test_out_of_scope_business_line(self):
        """Tickets with BusinessLine other than 'Infrastructure Services' must be rejected."""
        t = self._base_ticket()
        t['BusinessLine'] = 'Digital Workplace'
        in_scope, category = classify_ticket(t)
        self.assertFalse(in_scope)
        self.assertEqual(category, 'Out of Scope - BusinessLine')

        t['BusinessLine'] = 'Application Services'
        in_scope, category = classify_ticket(t)
        self.assertFalse(in_scope)
        self.assertEqual(category, 'Out of Scope - BusinessLine')

    def test_out_of_scope_severity(self):
        """Only Severity 1 (Molto alta) and Severity 2 (Alta) are permitted in scope."""
        t = self._base_ticket()
        for excluded_sev in ['3 - Media', '4 - Bassa', '5 - Pianificata', '', None]:
            t['Severity'] = excluded_sev
            in_scope, category = classify_ticket(t)
            self.assertFalse(in_scope, f"Severity '{excluded_sev}' should be excluded")
            self.assertEqual(category, 'Out of Scope - Severity')

    def test_valid_severities(self):
        """Both Severity 1 and Severity 2 must be accepted when criteria are met."""
        t = self._base_ticket()
        t['Severity'] = '1 - Molto alta'
        in_scope, cat = classify_ticket(t)
        self.assertTrue(in_scope)

        t['Severity'] = '2 - Alta'
        in_scope, cat = classify_ticket(t)
        self.assertTrue(in_scope)

    # -------------------------------------------------------------------------
    # Rule 3: Network Categorization
    # -------------------------------------------------------------------------
    def test_network_by_servizio_or_tipologia(self):
        """Managed Network Service and Networking categories are automatically flagged."""
        t = self._base_ticket()
        t['Servizio'] = 'MANAGED NETWORK SERVICE'
        t['Oggetto'] = 'General task'
        in_scope, cat = classify_ticket(t)
        self.assertTrue(in_scope)
        self.assertEqual(cat, 'Network')

        t['Servizio'] = 'Other Service'
        t['Tipologia Calcolata'] = 'NETWORKING INCIDENT'
        in_scope, cat = classify_ticket(t)
        self.assertTrue(in_scope)
        self.assertEqual(cat, 'Network')

    def test_network_by_keywords(self):
        """Keywords such as cato, router, firewall, wifi in text trigger Network scope."""
        t = self._base_ticket()
        t['Servizio'] = 'GENERIC IT SUPPORT'
        t['Tipologia Calcolata'] = 'INCIDENT'

        keywords = ['cato socket alarm', 'firewall packet dropped', 'vpn tunnel closed', 'wifi signal down', 'proxy error']
        for kw in keywords:
            t['Oggetto'] = f"Alert: {kw}"
            in_scope, cat = classify_ticket(t)
            self.assertTrue(in_scope, f"Failed to classify network keyword: {kw}")
            self.assertEqual(cat, 'Network')

    # -------------------------------------------------------------------------
    # Rule 4: Industrial OT & Plant Infrastructure
    # -------------------------------------------------------------------------
    def test_industrial_ot_accepted_in_plants(self):
        """Plant location + server/virtualization technology qualifies as Industrial OT."""
        t = self._base_ticket()
        t['Servizio'] = 'INFRASTRUCTURE COMPUTE'
        t['Tipologia Calcolata'] = 'INCIDENT'
        t['Oggetto'] = 'SimpliVity node hardware failure on production cluster'
        t['Configuration Item'] = 'SRV-ESX-01'
        t['Descrizione Articolo'] = 'Host compute virtualization node'
        t['LocalitaCI'] = 'Kocaeli Plant'
        t['Ubicazione sede Utente'] = 'Kocaeli'

        in_scope, cat = classify_ticket(t)
        self.assertTrue(in_scope)
        self.assertEqual(cat, 'Industrial OT')

    def test_industrial_ot_rejected_for_hq_and_cloud(self):
        """Centralized HQ (Cesano, Milano, TIM) and Cloud (Azure, GCP) must NOT qualify as Industrial OT."""
        t = self._base_ticket()
        t['Servizio'] = 'INFRASTRUCTURE COMPUTE'
        t['Tipologia Calcolata'] = 'INCIDENT'
        t['Oggetto'] = 'ESXi host crashed in server farm'
        t['Configuration Item'] = 'SRV-ESX-01'
        t['Descrizione Articolo'] = 'Host compute virtualization node'

        hq_locations = [
            'Cesano Maderno HQ',
            'Milano Head Office',
            'Cinisello Balsamo',
            'TIM Data Center Rozzano',
            'Azure Cloud WestEurope'
        ]
        for hq in hq_locations:
            t['LocalitaCI'] = hq
            t['Ubicazione sede Utente'] = hq
            in_scope, cat = classify_ticket(t)
            self.assertFalse(in_scope, f"HQ/Cloud location '{hq}' should not be Industrial OT")
            self.assertEqual(cat, 'Out of Scope - Other Infra')

    # -------------------------------------------------------------------------
    # Rule 5: SAP Exception Handling
    # -------------------------------------------------------------------------
    def test_sap_standard_application_excluded(self):
        """Standard SAP business application errors are strictly excluded."""
        t = self._base_ticket()
        t['Servizio'] = 'ENTERPRISE APPLICATIONS'
        t['Tipologia Calcolata'] = 'INCIDENT'
        t['Oggetto'] = 'SAP transaction VA02 short dump'
        t['Descrizione Articolo'] = 'SAP job cancelled with memory error'

        in_scope, cat = classify_ticket(t)
        self.assertFalse(in_scope)
        self.assertEqual(cat, 'Excluded - SAP Standard App')

    def test_sap_network_connectivity_included(self):
        """SAP tickets with network connectivity issues are classified as Network."""
        t = self._base_ticket()
        t['Servizio'] = 'ENTERPRISE APPLICATIONS'
        t['Tipologia Calcolata'] = 'INCIDENT'
        t['Oggetto'] = 'Cannot reach SAP server via Cato network link'

        in_scope, cat = classify_ticket(t)
        self.assertTrue(in_scope)
        self.assertEqual(cat, 'Network')

    # -------------------------------------------------------------------------
    # Plant Location Extraction Tests
    # -------------------------------------------------------------------------
    def test_get_plant_location_kocaeli(self):
        self.assertEqual(get_plant_location({'LocalitaCI': 'Kocaeli Plant'}), 'Kocaeli')
        self.assertEqual(get_plant_location({'Oggetto': 'Switch down in Izmit'}), 'Kocaeli')
        self.assertEqual(get_plant_location({'Configuration Item': 'TR-TPI-SW01'}), 'Kocaeli')

    def test_get_plant_location_alexandria(self):
        self.assertEqual(get_plant_location({'LocalitaCI': 'Alexandria Egypt Plant'}), 'Alexandria')
        self.assertEqual(get_plant_location({'Oggetto': 'Printer server in Alessandria'}), 'Alexandria')
        self.assertEqual(get_plant_location({'Configuration Item': 'EG-TPI-SRV01'}), 'Alexandria')

    def test_get_plant_location_santo_andre(self):
        self.assertEqual(get_plant_location({'LocalitaCI': 'Santo Andre Factory'}), 'Santo André')
        self.assertEqual(get_plant_location({'Configuration Item': 'BRSA-SW-01'}), 'Santo André')

    def test_get_plant_location_gravatai(self):
        self.assertEqual(get_plant_location({'LocalitaCI': 'Gravatai Tire Plant'}), 'Gravataí')
        self.assertEqual(get_plant_location({'Configuration Item': 'BRGV-ESX-01'}), 'Gravataí')

    def test_get_plant_location_other_sites(self):
        self.assertEqual(get_plant_location({'LocalitaCI': 'Cesano Maderno HQ'}), 'Other Sites')
        self.assertEqual(get_plant_location({}), 'Other Sites')


if __name__ == '__main__':
    unittest.main()
