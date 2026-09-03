"""
Unit Tests for src.config
=========================
Validates configuration constants, regex patterns, site mappings, and visual styling.
"""

import unittest
from src.config import (
    TARGET_BUSINESS_LINE,
    TARGET_SEVERITIES,
    NET_KW_PATTERN,
    HQ_CLOUD_PATTERN,
    PLANT_PATTERN,
    OT_TECH_PATTERN,
    SAP_PATTERN,
    SAP_CONN_PATTERN,
    PLANT_MAPPINGS,
    FONT_MAIN_TITLE,
    FILL_NAVY,
    CELL_BORDER,
    TOTAL_BORDER
)


class TestConfig(unittest.TestCase):

    def test_target_constants(self):
        """Validates that target BusinessLine and Severity levels match IT Infrastructure SLA rules."""
        self.assertEqual(TARGET_BUSINESS_LINE, 'Infrastructure Services')
        self.assertIn('1 - Molto alta', TARGET_SEVERITIES)
        self.assertIn('2 - Alta', TARGET_SEVERITIES)
        self.assertEqual(len(TARGET_SEVERITIES), 2)

    def test_network_keyword_pattern(self):
        """Validates that network keywords match properly with case insensitivity."""
        valid_samples = [
            "Network switch down in building B",
            "Cato SD-WAN socket unreachable",
            "Core router packet loss",
            "VPN tunnel disconnection issue",
            "Wi-Fi authentication failure",
            "Firewall rule configuration",
            "Proxy server latency",
            "Cisco ISE client posture error",
            "Zscaler ZPA authentication timeout",
            "Perdita di connessione alla rete",
            "Verifica rotte di routing",
            "DNS resolution failure on primary port"
        ]
        for sample in valid_samples:
            with self.subTest(sample=sample):
                self.assertIsNotNone(
                    NET_KW_PATTERN.search(sample),
                    f"Expected regex match for network text: {sample}"
                )

        invalid_samples = [
            "Printer jammed on 3rd floor",
            "Laptop monitor flickering",
            "Request for new keyboard and mouse",
            "Office 365 license renewal"
        ]
        for sample in invalid_samples:
            with self.subTest(sample=sample):
                self.assertIsNone(
                    NET_KW_PATTERN.search(sample),
                    f"Unexpected regex match for non-network text: {sample}"
                )

    def test_hq_cloud_pattern(self):
        """Validates exclusion patterns for Corporate Headquarters and Cloud tenants."""
        hq_samples = [
            "Cesano Maderno Headquarters",
            "Milano IT operations room",
            "TIM Data Center room 4",
            "Cinisello Balsamo office",
            "Microsoft Azure cloud region Europe",
            "GCP Kubernetes node pool"
        ]
        for sample in hq_samples:
            with self.subTest(sample=sample):
                self.assertIsNotNone(
                    HQ_CLOUD_PATTERN.search(sample),
                    f"Expected regex match for HQ/Cloud text: {sample}"
                )

    def test_plant_pattern(self):
        """Validates plant recognition patterns including accent variants."""
        plant_samples = [
            "Kocaeli plant tire curing sector",
            "Izmit manufacturing facility",
            "Alexandria mixing department",
            "Santo Andre plant server room",
            "Santo André plant maintenance",
            "Gravatai vulcanization facility",
            "Gravataí plant network cabinet",
            "Turkey plant OT environment",
            "Egypt plant factory line",
            "Brazil manufacturing site"
        ]
        for sample in plant_samples:
            with self.subTest(sample=sample):
                self.assertIsNotNone(
                    PLANT_PATTERN.search(sample),
                    f"Expected regex match for plant text: {sample}"
                )

    def test_ot_tech_pattern(self):
        """Validates Industrial OT technology and server compute keywords."""
        ot_samples = [
            "HPE SimpliVity hyperconverged node error",
            "VMware ESXi host hardware alarm",
            "Production line virtual machine frozen",
            "MES application server shutdown",
            "Veeam backup copy job failure",
            "Storage NAS share inaccessible",
            "SAN switch zone configuration"
        ]
        for sample in ot_samples:
            with self.subTest(sample=sample):
                self.assertIsNotNone(
                    OT_TECH_PATTERN.search(sample),
                    f"Expected regex match for OT tech text: {sample}"
                )

    def test_sap_patterns(self):
        """Validates differentiation between standard SAP app errors and network transport."""
        std_sap_samples = [
            "SAP transaction VA01 abap dump",
            "SAP Sprint background job cancelled",
            "User locked in SAP ERP system"
        ]
        for sample in std_sap_samples:
            with self.subTest(sample=sample):
                self.assertTrue(bool(SAP_PATTERN.search(sample)))
                self.assertFalse(bool(SAP_CONN_PATTERN.search(sample)))

        conn_sap_samples = [
            "SAP connection timeout via Cato tunnel",
            "Cannot connect to SAP server network error",
            "VPN link down SAP inaccessible"
        ]
        for sample in conn_sap_samples:
            with self.subTest(sample=sample):
                self.assertTrue(bool(SAP_PATTERN.search(sample)))
                self.assertTrue(bool(SAP_CONN_PATTERN.search(sample)))

    def test_plant_mappings_structure(self):
        """Validates structure and uniqueness of manufacturing plant mappings."""
        self.assertEqual(len(PLANT_MAPPINGS), 5)
        codes = [code for _, code in PLANT_MAPPINGS]
        self.assertEqual(len(codes), len(set(codes)), "Plant codes must be unique")
        self.assertIn("Kocaeli", codes)
        self.assertIn("Alexandria", codes)
        self.assertIn("Santo André", codes)
        self.assertIn("Gravataí", codes)
        self.assertIn("Other Sites", codes)

    def test_openpyxl_styles(self):
        """Validates that styling objects are initialized properly."""
        self.assertEqual(FONT_MAIN_TITLE.name, 'Segoe UI')
        self.assertTrue(FONT_MAIN_TITLE.bold)
        self.assertEqual(FILL_NAVY.fill_type, 'solid')
        self.assertIsNotNone(CELL_BORDER)
        self.assertIsNotNone(TOTAL_BORDER)


if __name__ == '__main__':
    unittest.main()
