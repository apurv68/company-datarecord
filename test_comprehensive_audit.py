"""
Comprehensive Regression Test Suite covering all outstanding bug patterns:
  1. Table 1 cross-field consistency (listed company / unlisted wording)
  2. Leadership signal → Table 5 mapping (Marketing Head, CEO appointment)
  3. Signal classification (financial performance vs M&A separation)
  4. M&A semantic mapping (stake acquisition, rights issue, demerger, financial performance)
  5. Marketing leadership detection (Priyanka Vishnoi / Liberty Shoes case)
  6. Evidence grounding (no invented claims)
  7. Airtel / Bharti Airtel entity resolution
"""
import os
import re
import unittest
from company_lookup import (
    EvidenceStore,
    is_entity_match,
    resolve_canonical_entity,
    validate_table1_cross_field_consistency,
    validate_table5_claims,
    identify_granular_event_type,
    is_marketing_leadership_claim,
    is_ceo_transition_claim,
    is_financial_performance_claim,
    is_acquisition_claim,
    is_capital_raise_claim,
    is_merger_claim,
    is_demerger_claim,
    is_caio_claim,
    is_store_expansion_claim,
    clean_insight_text,
)


class TestTable1CrossFieldConsistency(unittest.TestCase):
    """
    Fix 1: Table 1 field consistency.
    If Is Listed Company = Yes:
      - CTO and other exec fields must NEVER contain 'N/A (Unlisted / Not Publicly Disclosed)'
      - Business Type must be 'Public Limited', not 'Private Limited'
      - Stock Ticker should not say 'Unlisted'
    If Is Listed Company = No:
      - Stock Ticker must say 'N/A (Unlisted)'
    """

    def test_listed_company_no_unlisted_wording(self):
        """Listed company should never use 'Unlisted / Not Publicly Disclosed' in any field."""
        data = {
            "Company Name": "Liberty Shoes Limited",
            "CEO": "N/A",
            "CFO": "Rajesh Gupta",
            "CTO": "N/A",
            "Is Listed Company": "Yes",
            "Business Type (Private Limited/Public Limited)": "Private Limited",
            "Stock Ticker": "N/A (Unlisted)",
            "Current Market Cap (Market Value/Mcap)": "N/A (Privately Held)",
            "Share Price": "N/A (Privately Held)",
        }
        result = validate_table1_cross_field_consistency(data)
        # CTO should NOT say 'Unlisted / Not Publicly Disclosed'
        self.assertNotIn("Unlisted", result["CTO"])
        self.assertNotIn("Privately Held", result["CTO"])
        # Instead it should say 'N/A — Not publicly disclosed'
        self.assertIn("Not publicly disclosed", result["CTO"])
        # CEO similarly
        self.assertIn("Not publicly disclosed", result["CEO"])
        # Business Type should be corrected to Public Limited
        self.assertEqual(result["Business Type (Private Limited/Public Limited)"], "Public Limited")
        # Stock Ticker should not say 'Unlisted' for a listed company
        self.assertNotIn("Unlisted", result["Stock Ticker"])
        # Market cap and share price should not say 'Privately Held' for a listed company
        self.assertNotIn("Privately Held", str(result.get("Current Market Cap (Market Value/Mcap)", "")))
        self.assertNotIn("Privately Held", str(result.get("Share Price", "")))

    def test_listed_company_existing_exec_names_preserved(self):
        """Listed company with known exec names should preserve them."""
        data = {
            "Company Name": "Infosys Limited",
            "CEO": "Salil Parekh",
            "CFO": "Jayesh Sanghrajka",
            "CTO": "Mohammed Rafee Tarafdar",
            "Is Listed Company": "Yes",
            "Business Type (Private Limited/Public Limited)": "Public Limited",
            "Stock Ticker": "NSE/BSE: INFY",
        }
        result = validate_table1_cross_field_consistency(data)
        self.assertEqual(result["CEO"], "Salil Parekh")
        self.assertEqual(result["CFO"], "Jayesh Sanghrajka")
        self.assertEqual(result["CTO"], "Mohammed Rafee Tarafdar")

    def test_unlisted_company_ticker_wording(self):
        """Unlisted company should have 'N/A (Unlisted)' for Stock Ticker."""
        data = {
            "Company Name": "Bikanervala Foods Pvt Ltd",
            "CEO": "N/A",
            "CFO": "N/A",
            "CTO": "N/A",
            "Is Listed Company": "No",
            "Business Type (Private Limited/Public Limited)": "Private Limited",
            "Stock Ticker": "N/A",
        }
        result = validate_table1_cross_field_consistency(data)
        self.assertEqual(result["Stock Ticker"], "N/A (Unlisted)")
        # Exec N/A is fine for unlisted companies — no forced 'Not publicly disclosed'
        self.assertEqual(result["CEO"], "N/A")

    def test_listed_company_already_consistent(self):
        """Listed company with correct values should not be mutated."""
        data = {
            "Company Name": "TCS Limited",
            "CEO": "K. Krithivasan",
            "CFO": "Samir Seksaria",
            "CTO": "N/A",
            "Is Listed Company": "Yes",
            "Business Type (Private Limited/Public Limited)": "Public Limited",
            "Stock Ticker": "NSE: TCS",
            "Current Market Cap (Market Value/Mcap)": "₹14,50,000 Cr",
            "Share Price": "₹3,950",
        }
        result = validate_table1_cross_field_consistency(data)
        self.assertEqual(result["CEO"], "K. Krithivasan")
        self.assertEqual(result["Business Type (Private Limited/Public Limited)"], "Public Limited")
        self.assertIn("Not publicly disclosed", result["CTO"])


class TestMarketingLeadershipDetection(unittest.TestCase):
    """
    Fix 2: Leadership signal → Table 5 mapping.
    The Priyanka Vishnoi / Liberty Shoes case must be detected.
    """

    def test_priyanka_vishnoi_marketing_head(self):
        """Priyanka Vishnoi appointed as Marketing Head at Liberty Shoes must be detected."""
        texts = [
            "Priyanka Vishnoi, Marketing Head at Liberty Shoes, drives brand campaigns",
            "Priyanka Vishnoi appointed as Marketing Head of Liberty Shoes",
            "Liberty Shoes appoints Priyanka Vishnoi as Marketing Director",
            "Priyanka Vishnoi is the new Marketing Head at Liberty Shoes Limited",
        ]
        for text in texts:
            self.assertTrue(
                is_marketing_leadership_claim(text),
                f"Must detect marketing leadership in: '{text}'"
            )

    def test_marketing_head_generic_detection(self):
        """Generic marketing leadership claims must be detected."""
        positive_cases = [
            "Ananya Sharma appointed as CMO of Nykaa",
            "Rajiv Kumar joins as Chief Marketing Officer at Flipkart",
            "The company hired John as Head of Marketing",
            "Sarah was promoted to Marketing Director",
            "Marketing leadership change at ABC Corp announced today",
            "Kumar Verma serves as Marketing Head",
        ]
        for text in positive_cases:
            self.assertTrue(
                is_marketing_leadership_claim(text),
                f"Must detect marketing leadership in: '{text}'"
            )

    def test_non_marketing_claims_rejected(self):
        """Non-marketing claims must NOT be detected as marketing leadership."""
        negative_cases = [
            "Liberty Shoes reports 15% revenue growth in Q2",
            "Marketing strategy focuses on digital channels",  # strategy, not appointment
            "New product launch at Liberty Shoes retail outlets",
        ]
        for text in negative_cases:
            self.assertFalse(
                is_marketing_leadership_claim(text),
                f"Must NOT detect marketing leadership in: '{text}'"
            )


class TestCEOTransitionDetection(unittest.TestCase):
    """CEO/CXO appointment, resignation, and succession detection."""

    def test_ceo_appointment(self):
        """CEO appointments must be detected as transitions."""
        positive_cases = [
            "Adesh Gupta steps down as Managing Director of Liberty Shoes",
            "CEO Rajesh Kumar resigns from the company",
            "New CEO appointed to succeed retiring MD at Infosys",
            "Anish Shah takes charge as MD & CEO of Mahindra Group",
            "Board names Srinivas as next CEO replacing outgoing leader",
        ]
        for text in positive_cases:
            self.assertTrue(
                is_ceo_transition_claim(text),
                f"Must detect CEO transition in: '{text}'"
            )

    def test_non_ceo_claims_rejected(self):
        """Non-CEO leadership claims must NOT trigger CEO transition."""
        negative_cases = [
            "Liberty Shoes opens 10 new stores",
            "Revenue growth reported for the quarter",
            "Marketing Head appointed at the company",
        ]
        for text in negative_cases:
            self.assertFalse(
                is_ceo_transition_claim(text),
                f"Must NOT detect CEO transition in: '{text}'"
            )


class TestSignalClassification(unittest.TestCase):
    """
    Fix 3: Signal classification.
    Financial results MUST be classified as FINANCIAL_PERFORMANCE, never as M&A.
    Stake acquisition / rights issue MUST be classified as ACQUISITION / CAPITAL_RAISE.
    """

    def test_financial_results_never_mna(self):
        """Financial performance claims must map to FINANCIAL_PERFORMANCE, never to M&A types."""
        financial_texts = [
            "Liberty Shoes net profit declined 45% to Rs 3.5 crore in Q2 FY2025",
            "Revenue growth of 12% YoY driven by festive season demand",
            "EBITDA margin expanded from 8.5% to 10.2% in FY2025",
            "Net sales surged 18% to Rs 425 crore in the latest quarter",
            "Quarterly profit of Rs 15 crore, down from Rs 22 crore in the same quarter last year",
        ]
        for text in financial_texts:
            evt = identify_granular_event_type(text)
            self.assertEqual(
                evt, "FINANCIAL_PERFORMANCE",
                f"Financial text must classify as FINANCIAL_PERFORMANCE, got '{evt}' for: '{text}'"
            )
            # Also confirm these are NOT acquisitions
            self.assertFalse(
                is_acquisition_claim(text),
                f"Financial text must NOT be classified as acquisition: '{text}'"
            )

    def test_stake_acquisition_classification(self):
        """Stake acquisition must classify as ACQUISITION."""
        acq_texts = [
            "Adani Group acquired 51% stake in NDTV through indirect route",
            "Company hikes stake in subsidiary to 75% from 60%",
            "Tata Sons buys additional 10% stake in AirAsia India",
        ]
        for text in acq_texts:
            evt = identify_granular_event_type(text)
            self.assertEqual(
                evt, "ACQUISITION",
                f"Stake acquisition must classify as ACQUISITION, got '{evt}' for: '{text}'"
            )

    def test_rights_issue_classification(self):
        """Rights issue / capital raise must classify as CAPITAL_RAISE."""
        capital_texts = [
            "Company announces rights issue of Rs 500 crore to fund expansion",
            "Liberty Shoes raised Rs 75 crore through preferential allotment",
            "QIP approved by board to raise Rs 1000 crore equity",
        ]
        for text in capital_texts:
            evt = identify_granular_event_type(text)
            self.assertEqual(
                evt, "CAPITAL_RAISE",
                f"Capital raise must classify as CAPITAL_RAISE, got '{evt}' for: '{text}'"
            )

    def test_historical_demerger_rejected(self):
        """Historical demerger (10+ years ago) must not be treated as active demerger."""
        old_demerger = "The company was demerged in 2008 from the parent holding entity"
        self.assertFalse(
            is_demerger_claim(old_demerger),
            "Historical demerger from 2008 must be rejected"
        )

    def test_recent_demerger_accepted(self):
        """Recent demerger must be accepted."""
        new_demerger = "Demerger of cement business into a separate listed entity approved by NCLT in 2024"
        self.assertTrue(
            is_demerger_claim(new_demerger),
            "Recent demerger must be detected"
        )

    def test_marketing_leadership_event_type(self):
        """Marketing leadership appointment must classify as MARKETING_LEADERSHIP."""
        text = "Priyanka Vishnoi appointed as Marketing Head at Liberty Shoes"
        evt = identify_granular_event_type(text)
        self.assertEqual(evt, "MARKETING_LEADERSHIP")

    def test_ceo_exit_event_type(self):
        """CEO stepping down must classify as CEO_EXIT."""
        text = "CEO Rajesh Kumar steps down from the company after 10 years"
        evt = identify_granular_event_type(text)
        self.assertEqual(evt, "CEO_EXIT")


class TestTable5LeadershipMapping(unittest.TestCase):
    """
    Fix 2 (continued): validate_table5_claims Rule I must map leadership evidence.
    When Table 3 evidence contains a verified marketing appointment,
    Table 5's 'Growth & Marketing Leader' must NOT remain N/A.
    """

    def test_marketing_signal_maps_to_growth_leader(self):
        """Marketing leadership signal must map to Growth & Marketing Leader."""
        evidence_store = EvidenceStore("Liberty Shoes Limited")
        conclusions = {
            "Leadership Dynamics": {
                "Growth & Marketing Leader": "N/A — No verified marketing executive appointment found.",
                "CEO / CXO Hiring or Exit": "N/A — No verified leadership change found.",
                "CEO Transition / Stepping Down": "N/A — No verified leadership transition found.",
                "Chief AI Officer": "N/A — No verified Chief AI Officer appointment found.",
            },
        }
        signals = [
            {"text": "Priyanka Vishnoi, Marketing Head at Liberty Shoes, drives brand campaigns", "cat": "LEADERSHIP"},
        ]
        result, meta = validate_table5_claims(
            conclusions, evidence_store=evidence_store, all_signals=signals
        )
        growth_leader = result["Leadership Dynamics"]["Growth & Marketing Leader"]
        self.assertNotIn("N/A", growth_leader, "Growth & Marketing Leader should be populated from marketing signal")
        self.assertIn("Marketing Leadership", growth_leader)

    def test_ceo_signal_maps_to_cxo_field(self):
        """CEO appointment signal must map to CEO / CXO Hiring or Exit."""
        evidence_store = EvidenceStore("Liberty Shoes Limited")
        conclusions = {
            "Leadership Dynamics": {
                "Growth & Marketing Leader": "N/A",
                "CEO / CXO Hiring or Exit": "N/A — No verified leadership change found.",
                "CEO Transition / Stepping Down": "N/A — No verified leadership transition found.",
                "Chief AI Officer": "N/A",
            },
        }
        signals = [
            {"text": "Rajesh Kumar appointed as CEO of Liberty Shoes after board approval", "cat": "LEADERSHIP"},
        ]
        result, meta = validate_table5_claims(
            conclusions, evidence_store=evidence_store, all_signals=signals
        )
        cxo = result["Leadership Dynamics"]["CEO / CXO Hiring or Exit"]
        self.assertNotIn("N/A", cxo, "CEO/CXO field should be populated from CEO appointment signal")
        self.assertIn("Reported Leadership Movement", cxo)

    def test_evidence_store_fallback_for_marketing(self):
        """Evidence store Table #3 records should also populate Growth & Marketing Leader."""
        evidence_store = EvidenceStore("Liberty Shoes Limited")
        evidence_store.add_evidence(
            table="Table #3",
            category="Leadership",
            metric_or_event="Marketing Head Appointment",
            fact="Priyanka Vishnoi appointed as Marketing Head at Liberty Shoes",
            period="Current",
            period_type="Point-in-Time",
            source_name="Business Standard",
            source_url="https://www.business-standard.com",
            confidence="High",
            verified=True,
        )
        conclusions = {
            "Leadership Dynamics": {
                "Growth & Marketing Leader": "N/A — No verified marketing executive appointment found.",
                "CEO / CXO Hiring or Exit": "N/A",
                "CEO Transition / Stepping Down": "N/A",
                "Chief AI Officer": "N/A",
            },
        }
        # No all_signals provided — should fall back to evidence store
        result, meta = validate_table5_claims(
            conclusions, evidence_store=evidence_store, all_signals=[]
        )
        growth_leader = result["Leadership Dynamics"]["Growth & Marketing Leader"]
        self.assertNotIn("N/A", growth_leader, "Should populate from evidence store Table #3 records")
        self.assertIn("Marketing Leadership", growth_leader)


class TestMnASemanticMapping(unittest.TestCase):
    """
    Fix 4: M&A semantic mapping rules.
    - Financial performance → never M&A
    - Stake acquisition → ACQUISITION
    - Rights issue / capital raise → CAPITAL_RAISE
    - Historical demerger → rejected
    """

    def test_financial_performance_never_in_mna(self):
        """Financial performance text must be moved OUT of M&A section by Rule D."""
        evidence_store = EvidenceStore("Liberty Shoes Limited")
        conclusions = {
            "Mergers, Acquisitions & Capital Actions": {
                "Buying Company / Startup": "Net profit declined 45% to Rs 3.5 crore in Q2 FY2025",
                "Merged with Company": "N/A",
                "Demerger": "N/A",
                "New Funding / IPO Launch": "N/A",
            },
            "Financial Health": {
                "Latest Interim Performance": "N/A",
            },
        }
        result, meta = validate_table5_claims(
            conclusions, evidence_store=evidence_store
        )
        buying = result["Mergers, Acquisitions & Capital Actions"]["Buying Company / Startup"]
        self.assertTrue(buying.startswith("N/A"), f"Financial performance must be removed from M&A: got '{buying}'")

    def test_acquisition_signal_maps_correctly(self):
        """Acquisition signal should map to Buying Company field via Rule M."""
        evidence_store = EvidenceStore("Test Company")
        conclusions = {
            "Mergers, Acquisitions & Capital Actions": {
                "Buying Company / Startup": "N/A — No verified evidence available.",
                "Merged with Company": "N/A",
                "Demerger": "N/A",
                "New Funding / IPO Launch": "N/A",
            },
        }
        signals = [
            {"text": "Test Company acquired 100% stake in Startup XYZ for $50 million", "cat": "M&A"},
        ]
        result, meta = validate_table5_claims(
            conclusions, evidence_store=evidence_store, all_signals=signals
        )
        buying = result["Mergers, Acquisitions & Capital Actions"]["Buying Company / Startup"]
        self.assertNotIn("N/A", buying, "Acquisition signal should populate Buying Company field")
        self.assertIn("Acquisition Recorded", buying)

    def test_capital_raise_signal_maps_correctly(self):
        """Capital raise signal should map to New Funding field via Rule M."""
        evidence_store = EvidenceStore("Test Company")
        conclusions = {
            "Mergers, Acquisitions & Capital Actions": {
                "Buying Company / Startup": "N/A",
                "Merged with Company": "N/A",
                "Demerger": "N/A",
                "New Funding / IPO Launch": "N/A — No verified evidence available.",
            },
        }
        signals = [
            {"text": "Test Company announces rights issue of Rs 200 crore to fund expansion", "cat": "M&A"},
        ]
        result, meta = validate_table5_claims(
            conclusions, evidence_store=evidence_store, all_signals=signals
        )
        funding = result["Mergers, Acquisitions & Capital Actions"]["New Funding / IPO Launch"]
        self.assertNotIn("N/A", funding, "Capital raise signal should populate New Funding field")
        self.assertIn("Capital Action", funding)


class TestEvidenceGrounding(unittest.TestCase):
    """
    Fix 5: Evidence grounding — no invented claims.
    Unverified negative patterns must be blocked.
    N/A cannot co-exist with matching evidence.
    """

    def test_unverified_negative_blocked(self):
        """Claims like 'No event occurred' must be sanitized to N/A."""
        evidence_store = EvidenceStore("Test Company")
        conclusions = {
            "Contraction & Shutdown Signals": {
                "Closing Factory / Plant": "No plant shutdown occurred during the reporting period",
                "Closing Stores": "No mass store closures reported",
            },
        }
        result, meta = validate_table5_claims(conclusions, evidence_store=evidence_store)
        for field_name in ["Closing Factory / Plant", "Closing Stores"]:
            val = result["Contraction & Shutdown Signals"][field_name]
            self.assertTrue(val.startswith("N/A"), f"Unverified negative must become N/A: '{val}'")


class TestAirtelEntityResolution(unittest.TestCase):
    """
    Fix 6: Airtel / Bharti Airtel entity resolution.
    Ambiguous 'airtel' query must NOT automatically select 'Airtel India'
    when Bharti Airtel is the relevant corporate entity.
    """

    def test_airtel_entity_match_accepts_bharti_airtel(self):
        """Bharti Airtel content must be accepted for Bharti Airtel canonical entity."""
        canon = "Bharti Airtel Limited"
        aliases = ["Bharti Airtel", "Airtel", "Airtel India"]
        self.assertTrue(
            is_entity_match("Bharti Airtel reports quarterly revenue growth", "", canonical_name=canon, aliases=aliases)
        )
        self.assertTrue(
            is_entity_match("Airtel announces 5G network expansion in India", "", canonical_name=canon, aliases=aliases)
        )

    def test_airtel_rejects_unrelated_entities(self):
        """Unrelated entities (Jio, Vodafone) must be rejected for Bharti Airtel."""
        canon = "Bharti Airtel Limited"
        aliases = ["Bharti Airtel", "Airtel"]
        self.assertFalse(
            is_entity_match("Reliance Jio announces new tariff plan for broadband", "", canonical_name=canon, aliases=aliases)
        )
        self.assertFalse(
            is_entity_match("Vodafone Idea reports subscriber loss in Q3", "", canonical_name=canon, aliases=aliases)
        )


class TestIdentifyGranularEventType(unittest.TestCase):
    """Verify granular event type classification edge cases."""

    def test_financial_not_acquisition(self):
        """Financial performance text must be FINANCIAL_PERFORMANCE, not ACQUISITION."""
        text = "Net profit declined 28% as EBITDA margin contracted 120 bps"
        self.assertEqual(identify_granular_event_type(text), "FINANCIAL_PERFORMANCE")

    def test_store_expansion(self):
        """Store opening must be STORE_EXPANSION."""
        text = "Company plans to open 100 new stores across tier-2 cities"
        self.assertEqual(identify_granular_event_type(text), "STORE_EXPANSION")

    def test_caio_detection(self):
        """Chief AI Officer appointment must be AI_LEADERSHIP."""
        text = "Company appointed Dr. Kumar as Chief AI Officer to lead AI strategy"
        self.assertEqual(identify_granular_event_type(text), "AI_LEADERSHIP")


class TestTable2Display(unittest.TestCase):
    """Verify Table #2 renders correctly and handles various financial row formats."""

    def test_display_table2_renders_without_error(self):
        from company_lookup import display_table2
        sample_data = {
            "Company Name": "Liberty Shoes Limited",
            "periods": ["Mar 2024", "Mar 2023"],
            "rows": [
                {
                    "Fiscal Period / Year": "Mar 2024",
                    "Market Cap": "Rs 540 Cr",
                    "Net Revenue/Net Sales": "Rs 650 Cr",
                    "Net Profit": "Rs 14 Cr",
                    "EBITDA": "Rs 45 Cr",
                    "Employee Headcount": "2,100"
                },
                {
                    "Fiscal Period / Year": "Mar 2023",
                    "Market Cap": "Rs 480 Cr",
                    "Net Revenue/Net Sales": "Rs 600 Cr",
                    "Net Profit": "Rs 12 Cr",
                    "EBITDA": "Rs 40 Cr",
                    "Employee Headcount": "2,050"
                }
            ]
        }
        sources = [{"name": "Screener.in Audited Financials", "url": "https://www.screener.in/company/LIBERTSHOE/"}]
        # Must execute cleanly without exception
        try:
            display_table2(sample_data, sources)
        except Exception as e:
            self.fail(f"display_table2 raised an exception: {e}")


class TestTable3NewsQuality(unittest.TestCase):
    """Verify Table #3 news headline cleaning and junk elimination."""

    def test_clean_news_headline_strips_show_prefixes(self):
        from company_lookup import clean_news_headline
        raw = "afaqs! Pause - Anupam Bansal, Executive Director, Liberty Shoes"
        cleaned = clean_news_headline(raw)
        self.assertNotIn("afaqs! Pause", cleaned)
        self.assertEqual(cleaned, "Anupam Bansal, Executive Director, Liberty Shoes")

    def test_clean_news_headline_strips_quote_attribution(self):
        from company_lookup import clean_news_headline
        raw = "Clear execution needs humility and focus, says Anupam Bansal"
        cleaned = clean_news_headline(raw)
        self.assertNotIn("says Anupam Bansal", cleaned)
        self.assertEqual(cleaned, "Clear execution needs humility and focus")

    def test_is_news_junk_filters_clickbait_and_gossip(self):
        from company_lookup import is_news_junk
        self.assertTrue(is_news_junk("Mercs and 'mic drops': What's behind the high drama at Liberty Shoes?"))
        self.assertTrue(is_news_junk("afaqs! Pause - Interview series with CXOs"))
        self.assertTrue(is_news_junk("Execution needs humility and focus"))
        self.assertFalse(is_news_junk("Liberty Shoes appoints Priyanka Vishnoi as head of marketing"))
        self.assertFalse(is_news_junk("Liberty Shoes reports Q3 FY25 net profit of Rs 5.2 crore"))


class TestSaveCompanyDocx(unittest.TestCase):
    """Verify DOCX generation for all 5 tables."""

    def test_save_company_docx_creates_valid_document(self):
        from company_lookup import save_company_docx
        import docx

        data1 = {
            "Company Name": "Liberty Shoes Limited",
            "Business Type": "Footwear Manufacturing & Retail",
            "Is Listed Company": "Yes",
            "Stock Ticker": "LIBERTSHOE",
            "Official Domain": "libertyshoes.com",
            "Registered Address": "Karnal, Haryana, India",
            "Managing Director / CEO": "Adesh Kumar Gupta",
            "Chief Financial Officer (CFO)": "Sunil Bansal",
            "Chief Technology Officer (CTO)": "N/A — Not publicly disclosed"
        }
        sources1 = [{"name": "BSE India", "url": "https://www.bseindia.com"}]

        data2 = {
            "Company Name": "Liberty Shoes Limited",
            "periods": ["Mar 2024"],
            "rows": [
                {
                    "Fiscal Period / Year": "Mar 2024",
                    "Market Cap": "Rs 540 Cr",
                    "Net Revenue/Net Sales": "Rs 650 Cr",
                    "Net Profit": "Rs 14 Cr",
                    "EBITDA": "Rs 45 Cr",
                    "Employee Headcount": "2,100"
                }
            ]
        }
        sources2 = [{"name": "Screener.in", "url": "https://www.screener.in"}]

        data3 = {
            "2026 Developments & Strategic Milestones": [
                "[LEADERSHIP SIGNAL] [Liberty Shoes Directly] [CONFIRMED] Liberty Shoes appoints Priyanka Vishnoi as Head of Marketing (Year: 2026)\n    ↳ Intelligence Brief: Leads national retail branding and product portfolio expansion.\n↳ Accelerates omnichannel growth."
            ]
        }
        sources3 = [{"name": "Google Gemini Intelligence", "url": "https://generativelanguage.googleapis.com"}]

        data4 = {
            "Core Business Profile": "Manufacturer and retailer of footwear, accessories, and leather goods.",
            "Brands & Trademarks": ["Gliders", "Force 10", "Senorita", "Fortune"],
            "Industry / Sector": "Consumer Discretionary - Footwear"
        }
        sources4 = [{"name": "Company Website", "url": "https://libertyshoes.com"}]

        data5 = {
            "Growth Assessment": {
                "Verdict": "Steady Growth",
                "Summary & Drivers": "Expanding retail footprints across Tier-2 cities."
            },
            "Expansion Vectors": {
                "New Markets": "Tier-2/3 Retail Expansion"
            },
            "Leadership Dynamics": {
                "Growth & Marketing Leader": "Priyanka Vishnoi (Head of Marketing)"
            }
        }
        sources5 = [{"name": "Regulatory Filings", "url": "https://www.nseindia.com"}]

        test_out = "scratch_test_liberty.docx"
        try:
            saved_path = save_company_docx(
                data1, sources1, data2, sources2, data3, sources3, data4, sources4, data5, sources5,
                output_filepath=test_out
            )
            self.assertIsNotNone(saved_path)
            self.assertTrue(os.path.exists(test_out))
            self.assertGreater(os.path.getsize(test_out), 1000)

            # Reopen document and inspect contents
            doc = docx.Document(test_out)
            text_blob = "\n".join([p.text for p in doc.paragraphs] + [c.text for tbl in doc.tables for r in tbl.rows for c in r.cells])
            self.assertIn("Liberty Shoes Limited", text_blob)
            self.assertIn("1.0 Corporate Identity & Leadership Governance", text_blob)
            self.assertIn("2.0 5-Year Historical & Present Financial Disclosures", text_blob)
            self.assertIn("3.0 Latest Corporate Developments & Strategic Milestones", text_blob)
            self.assertIn("4.0 Business Activities, Operational Footprint & Brand Matrix", text_blob)
            self.assertIn("5.0 Strategic Conclusions & Growth Assessment", text_blob)
            self.assertIn("Priyanka Vishnoi", text_blob)
            # Tables exist in document
            self.assertGreaterEqual(len(doc.tables), 3)
        finally:
            if os.path.exists(test_out):
                try:
                    os.remove(test_out)
                except Exception:
                    pass


if __name__ == "__main__":
    unittest.main()


