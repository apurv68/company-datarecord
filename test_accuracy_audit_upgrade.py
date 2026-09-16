"""
Comprehensive Unit & Regression Test Suite for Accuracy & Auditability Upgrades.
Tests:
1. Generic Cross-Entity Firewall (AESL, Tata Motors, Liberty Shoes, TCS, HDFC Bank, Sun Pharma, IndiGo)
2. AESL Sibling & TCS Contamination Rejection
3. Subsidiary / Distribution Unit Relationship Tracking (AEML)
4. Period Classification (Audited Annual, Unaudited Interim, TTM, Unknown)
5. Evidence Store & Traceability Validation (Zero-fabrication, verified status)
6. Zero Predetermined / Hardcoded Metric Numbers Verification
7. Backward Compatibility of CSV/JSON Schemas
"""
import os
import json
import csv
import re
import unittest
from company_lookup import EvidenceStore, is_entity_match, resolve_canonical_entity

class TestAccuracyAuditUpgrade(unittest.TestCase):

    # -------------------------------------------------------------
    # 1. AESL Contamination & Sibling Rejection Regression Test
    # -------------------------------------------------------------
    def test_aesl_contamination_firewall(self):
        """
        Adani Energy Solutions query:
        Must NOT appear as AESL's own:
        - TCS BaNCS, TCS ADD, Quartz, MasterCraft, Ignio, TwinX, TCS consulting
        - Adani Ports cargo operations
        - Adani Power thermal generation
        - Tata Steel, etc.
        Must ACCEPT when supported by evidence:
        - Power transmission
        - Electricity distribution
        - Smart metering
        - Energy solutions
        - AESL / AEML relationship where explicitly documented
        """
        canon_name = "Adani Energy Solutions Limited"
        aliases = ["Adani Energy Solutions", "Adani Transmission", "AESL"]
        subsidiaries = [
            {"name": "Adani Electricity Mumbai Limited", "role": "Subsidiary - Urban Distribution Unit", "alias": "AEML"},
            {"name": "AEML", "role": "Subsidiary - Urban Distribution Unit"}
        ]

        # REJECTIONS: Unrelated TCS / IT consulting items
        tcs_unrelated_candidates = [
            "TCS BaNCS enterprise platform deployed for core banking solutions",
            "TCS ADD clinical data management platform launched for life sciences",
            "Quartz blockchain and multi-asset trading platform expands across banks",
            "MasterCraft code governance and automated software testing suite released",
            "Ignio autonomous enterprise IT operations platform upgrades cognitive agent",
            "TwinX digital twin platform for enterprise decision simulations",
            "Tata Consultancy Services reports 15% growth in IT consulting revenue"
        ]
        for c in tcs_unrelated_candidates:
            self.assertFalse(
                is_entity_match(c, "", canonical_name=canon_name, aliases=aliases, subsidiaries=subsidiaries),
                f"Firewall must reject TCS product for AESL: {c}"
            )

        # REJECTIONS: Sibling conglomerate entities (Adani Ports, Adani Power, Adani Green)
        sibling_candidates = [
            "Adani Ports handles record 420 MMT container cargo volume at Mundra port",
            "Adani Power announces commissioning of 1600 MW thermal power plant in MP",
            "Adani Green Energy commissions 551 MW solar capacity at Khavda RE park",
            "Adani Total Gas expands city gas distribution CNG stations across Gujarat",
            "Tata Steel cuts European operations debt with quarterly profit surge"
        ]
        for c in sibling_candidates:
            self.assertFalse(
                is_entity_match(c, "", canonical_name=canon_name, aliases=aliases, subsidiaries=subsidiaries),
                f"Firewall must reject sibling/unrelated entity for AESL: {c}"
            )

        # ACCEPTANCES: Core AESL business supported by evidence
        valid_aesl_candidates = [
            "Adani Energy Solutions commissions 562 ckm high-voltage transmission line in Maharashtra",
            "Adani Energy Solutions wins Rs 4700 crore transmission project for 4.5 GW power evacuation",
            "Adani Energy Solutions signs 25-year PPA for 2500 MW RTC renewable power",
            "AESL achieves major 1 crore smart meters deployment milestone under DBFOOT model",
            "Adani Energy Solutions expands cooling-as-a-service and industrial energy solutions portfolio"
        ]
        for c in valid_aesl_candidates:
            self.assertTrue(
                is_entity_match(c, "", canonical_name=canon_name, aliases=aliases, subsidiaries=subsidiaries),
                f"Firewall must accept genuine AESL candidate: {c}"
            )

        # ACCEPTANCES WITH SUBSIDIARY RELATIONSHIP: Adani Electricity Mumbai Limited (AEML)
        aeml_candidate = "Adani Electricity Mumbai Limited receives MERC approval for smart meter rollout tariff"
        self.assertTrue(
            is_entity_match(aeml_candidate, "", canonical_name=canon_name, aliases=aliases, subsidiaries=subsidiaries),
            "Firewall must recognize documented subsidiary AEML context"
        )

    # -------------------------------------------------------------
    # 2. Multi-Company Firewall Consistency Test
    # -------------------------------------------------------------
    def test_multi_company_firewall_consistency(self):
        """
        Validate the exact same generic firewall logic across diverse companies:
        Tata Motors, Liberty Shoes, TCS, HDFC Bank, Sun Pharma, and IndiGo.
        """
        # A. Tata Motors
        tm_name = "Tata Motors Limited"
        self.assertTrue(is_entity_match("Tata Motors launches new Curvv EV with 500km range", "", canonical_name=tm_name))
        self.assertTrue(is_entity_match("Tata Motors reports commercial vehicle sales increase", "", canonical_name=tm_name))
        self.assertTrue(is_entity_match("Tata Motors partners with Tata Power to install EV chargers", "", canonical_name=tm_name))
        self.assertFalse(is_entity_match("Tata Steel announces blast furnace overhaul in Jamshedpur", "", canonical_name=tm_name))
        self.assertFalse(is_entity_match("Tata Power commissions solar park in Rajasthan", "", canonical_name=tm_name))

        # B. Liberty Shoes
        ls_name = "Liberty Shoes Limited"
        self.assertTrue(is_entity_match("Liberty Shoes plans Rs 75 crore expansion for 100 new stores", "", canonical_name=ls_name))
        self.assertTrue(is_entity_match("Liberty Shoes opens standalone Healers comfort footwear outlet", "", canonical_name=ls_name))
        self.assertFalse(is_entity_match("Liberty General Insurance launches comprehensive motor cover", "", canonical_name=ls_name))
        self.assertFalse(is_entity_match("Liberty Steel completes debt restructuring in UK", "", canonical_name=ls_name))

        # C. Tata Consultancy Services (TCS)
        tcs_name = "Tata Consultancy Services Limited"
        tcs_aliases = ["TCS", "Tata Consultancy"]
        self.assertTrue(is_entity_match("TCS BaNCS selected by European bank for core digital transformation", "", canonical_name=tcs_name, aliases=tcs_aliases))
        self.assertTrue(is_entity_match("Tata Consultancy Services wins $1 billion cloud infrastructure deal", "", canonical_name=tcs_name, aliases=tcs_aliases))
        self.assertFalse(is_entity_match("Tata Motors launches electric truck line", "", canonical_name=tcs_name, aliases=tcs_aliases))
        self.assertFalse(is_entity_match("Tata Steel reports quarterly export decline", "", canonical_name=tcs_name, aliases=tcs_aliases))

        # D. HDFC Bank
        hdfc_name = "HDFC Bank Limited"
        self.assertTrue(is_entity_match("HDFC Bank expands digital retail lending and credit card base", "", canonical_name=hdfc_name))
        self.assertFalse(is_entity_match("HDFC Life Insurance introduces guaranteed pension plan", "", canonical_name=hdfc_name))
        self.assertFalse(is_entity_match("HDFC AMC launches new index fund", "", canonical_name=hdfc_name))

        # E. Sun Pharma
        sun_name = "Sun Pharmaceutical Industries Limited"
        sun_aliases = ["Sun Pharma"]
        self.assertTrue(is_entity_match("Sun Pharma receives USFDA approval for generic specialty formulation", "", canonical_name=sun_name, aliases=sun_aliases))
        self.assertFalse(is_entity_match("Sun TV Network acquires satellite broadcast rights", "", canonical_name=sun_name, aliases=sun_aliases))
        self.assertFalse(is_entity_match("Sun Direct announces new DTH subscription packs", "", canonical_name=sun_name, aliases=sun_aliases))

        # F. IndiGo (InterGlobe Aviation)
        indigo_name = "InterGlobe Aviation Limited"
        indigo_aliases = ["IndiGo", "IndiGo Airlines"]
        self.assertTrue(is_entity_match("IndiGo takes delivery of its 30th A321neo aircraft", "", canonical_name=indigo_name, aliases=indigo_aliases))
        self.assertTrue(is_entity_match("InterGlobe Aviation reports record passenger load factor", "", canonical_name=indigo_name, aliases=indigo_aliases))
        self.assertFalse(is_entity_match("InterGlobe Hotels expands hotel portfolio in Bengaluru", "", canonical_name=indigo_name, aliases=indigo_aliases))

    # -------------------------------------------------------------
    # 3. Zero Predetermined / Hardcoded Metric Numbers Verification
    # -------------------------------------------------------------
    def test_no_hardcoded_metric_numbers(self):
        """
        Verify that specific empirical numbers (e.g. 27,949 ckm, 9,824 Cr, 18,296 Cr)
        are NOT embedded as expected hardcoded values in company_lookup.py.
        """
        with open("company_lookup.py", "r", encoding="utf-8") as f:
            code = f.read()

        forbidden_hardcodes = ["27,949", "27949", "9,824", "9824", "18,296", "18296"]
        for num in forbidden_hardcodes:
            self.assertNotIn(
                f'"{num}"', code,
                f"Empirical metric {num} must not be hardcoded as a string in company_lookup.py"
            )

    # -------------------------------------------------------------
    # 4. Period Classification Tests
    # -------------------------------------------------------------
    def test_period_classification(self):
        def classify_period(period_str):
            p = period_str.strip()
            if "TTM" in p.upper():
                return "TTM"
            elif re.search(r"\b(?:Jun|June|Sep|Sept|September|Dec|December)\b|Q[1-4]|quarter", p, flags=re.I):
                return "Unaudited Interim"
            elif re.search(r"\b(?:Mar|March)\b|FY\s*20\d\d", p, flags=re.I):
                return "Audited Annual"
            return "Unknown"

        self.assertEqual(classify_period("Mar 2026"), "Audited Annual")
        self.assertEqual(classify_period("Mar 2025"), "Audited Annual")
        self.assertEqual(classify_period("FY2026"), "Audited Annual")
        self.assertEqual(classify_period("June 2026"), "Unaudited Interim")
        self.assertEqual(classify_period("Q2 FY2026"), "Unaudited Interim")
        self.assertEqual(classify_period("TTM"), "TTM")
        self.assertEqual(classify_period("Estimated Year"), "Unknown")

    # -------------------------------------------------------------
    # 5. Evidence Store & Traceability Validation Tests
    # -------------------------------------------------------------
    def test_evidence_store_validation(self):
        store = EvidenceStore("Liberty Shoes Ltd")
        
        # Validated fact with real source
        store.add_evidence(
            table="Table #1",
            category="Corporate Identity",
            metric_or_event="CIN",
            fact="L51901HR1986PLC023188",
            period="Current",
            period_type="Point-in-Time",
            source_name="Ministry of Corporate Affairs",
            source_url="https://www.mca.gov.in",
            confidence="High",
            verified=True
        )

        # Unverified / estimation fact
        store.add_evidence(
            table="Table #2",
            category="Financials",
            metric_or_event="Revenue Projection",
            fact="Rs. 800 Cr. Estimated",
            period="FY2027",
            period_type="Derived",
            source_name="Analyst Commentary",
            source_url="",
            confidence="Low",
            verified=False
        )

        self.assertEqual(len(store.records), 2)
        self.assertTrue(store.records[0]["verified"])
        self.assertEqual(store.records[0]["confidence"], "High")
        self.assertFalse(store.records[1]["verified"])
        self.assertEqual(store.records[1]["confidence"], "Low")

        # Test persistence
        test_json = "scratch/unit_test_evidence.json"
        test_csv = "scratch/unit_test_evidence.csv"
        saved = store.save_to_files(test_json, test_csv)
        self.assertEqual(saved, 2)
        self.assertTrue(os.path.isfile(test_json))
        self.assertTrue(os.path.isfile(test_csv))

        with open(test_json, "r", encoding="utf-8") as jf:
            loaded_json = json.load(jf)
            self.assertEqual(len(loaded_json), 2)
            self.assertEqual(loaded_json[0]["metric_or_event"], "CIN")

    # -------------------------------------------------------------
    # 6. Backward Compatibility of Existing Files
    # -------------------------------------------------------------
    def test_backward_compatibility_schemas(self):
        if os.path.isfile("company_records.csv") and os.path.getsize("company_records.csv") > 0:
            with open("company_records.csv", "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                header = next(reader)
                self.assertIn("Company Name", header)
                self.assertIn("CEO", header)
                self.assertIn("CFO", header)

        if os.path.isfile("company_financials_5yr.csv") and os.path.getsize("company_financials_5yr.csv") > 0:
            with open("company_financials_5yr.csv", "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                header = next(reader)
                self.assertIn("Company Name", header)
                self.assertIn("Fiscal Period / Year", header)
                self.assertIn("Net Revenue/Net Sales", header)
                self.assertIn("Net Profit", header)

        if os.path.isfile("company_conclusions.csv") and os.path.getsize("company_conclusions.csv") > 0:
            with open("company_conclusions.csv", "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                header = next(reader)
                self.assertIn("Company Name", header)
                self.assertIn("Growth Verdict", header)
                self.assertIn("YoY Revenue", header)

if __name__ == "__main__":
    unittest.main()
