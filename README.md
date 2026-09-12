# pilot-business-intelligence-system

This pilot project will be built across 3 phases as:

**Phase 1: Company Profile**  
Accurately build the company profile.

**Phase 2: Signal Engine**  
Continuously discover and classify meaningful changes.

**Phase 3: Intelligence Engine**  
Explain what those changes mean and what someone should do about them.

---

### Signals Taxonomy
- **Type**: Leadership / Expansion / Technology / Financial / Legal / Operations & Contracts / Strategic Development
- **Date**
- **Source**
- **Confidence**
- **Recency**
- **Importance**
- **What changed?**
- **Why does it matter?**
- **Possible business implication**
- **Related technologies/services**

### Information to be Collected

| Bucket                 | Examples                                                                        |
| ---------------------- | ------------------------------------------------------------------------------- |
| **Identity**           | Name, website, HQ, industry, founded, founders, CEO                             |
| **Business**           | Products, categories, markets, business model, locations                        |
| **Growth Signals**     | New stores, warehouses, plants, offices, hiring, expansion                      |
| **Change Signals**     | M&A, funding, leadership changes, launches, partnerships, lawsuits              |
| **Technology Signals** | Website stack, ecommerce platform, mobile app, AI usage, CRM, cloud, automation |

---

# Corporate Intelligence & Signal Engine (`company_lookup.py`)

A Python tool that takes any company name as input, performs canonical entity resolution to prevent cross-company contamination, and automatically retrieves corporate intelligence and live market signals grouped into structured tables with year-wise timeline records.

---

## Key Features

1. **Canonical Entity Resolution & Negative Shields:**
   - Pre-resolution candidate selection disambiguates conglomerate peers and subsidiaries (e.g., `Adani Energy Solutions` vs `Adani Power` vs `TCS`).
   - Dynamic negative regex filters prevent cross-company noise or brand leakage.
   - 4-Tier news evidence classification (Tier 1 Primary Subject, Tier 2 Joint Venture/Partnership, Tier 3 Group Company, Tier 4 Unrelated).

2. **Corporate Intelligence Signal Taxonomy:**
   - Automatically parses corporate developments into granular signal categories:
     - `[PROCUREMENT / CONTRACT SIGNAL]`
     - `[FINANCIAL & M&A SIGNAL]`
     - `[STRATEGIC DEVELOPMENT]`
     - `[OPERATIONS & PROJECTS]`
     - `[EXPANSION & INFRASTRUCTURE]`
     - `[REGULATORY / LEGAL SIGNAL]`
     - `[INVESTOR & CAPITAL MARKETS SIGNAL]`
     - `[TECHNOLOGY / CYBERSECURITY SIGNAL]`
     - `[LEADERSHIP / WORKFORCE SIGNAL]`
     - `[AWARDS / RECOGNITION]`
     - `[PRODUCT & FLEET SIGNAL]`
     - `[CORPORATE UPDATE]`

3. **📅 Year-Wise Timeline Categorization:**
   - Every category automatically parses dates and fiscal years (`FY26`, `FY25`, `2026`, `2025`, `2024`, etc.).
   - Findings are chronologically sorted and tagged:
     - `📅 [2026 (Current Year)]`
     - `📅 [2025 (Previous Year)]`
     - `📅 [2024 / Historical]`
     - `📅 [Recent / Active Records]`

4. **Multi-Table Intelligence Architecture:**
   - **Table #1:** Corporate Identity, Legal Type, Leadership (CEO, CFO, CTO) & Market Standing.
   - **Table #2:** 5-Year Historical and Present Financials (Revenue, Net Profit, Market Cap, Headcount).
   - **Table #3:** Signal Intelligence, Live Corporate Signals, and News Attribution.
   - **Table #4:** Operational Profile, Brands, Products, Retail vs. Customer Service distinction, and Archetype Classifiers.
   - **Table #5:** Strategic Business Intelligence Conclusions & Growth Assessment (Growth trajectory verdict, Expansion vectors, Operational shutdowns/discontinuations, Leadership dynamics & AI transformation, Property/real estate acquisitions vs sales, M&A/demerger/capital actions, YoY revenue & margin health).

5. **Multi-Format Export with Automatic Sync:**
   - **`company_records.csv`**: Table #1 corporate profiles and leadership directory.
   - **`company_financials_5yr.csv`**: Table #2 audited 5-year historical and present financials.
   - **`company_conclusions.csv`**: Table #5 strategic conclusions and growth intelligence across all 7 dimensions.
   - **`company_records.json`**: Complete hierarchical JSON format integrating all 5 tables and verified source links.

---

## How to Run

### Interactive Mode
```powershell
python company_lookup.py
```
1. Type a company name (e.g., `Adani Energy Solutions`, `Tata Consultancy Services`, `Reliance Jio`, `Amul`).
2. If similar/matching companies exist, select the number from the list or press enter to proceed with the exact entered text.
3. The engine scans multi-source search backends, classifies signals, renders formatted tables, and exports records to CSV and JSON.

### Direct Command-Line Mode
```powershell
python company_lookup.py "Adani Energy Solutions"
python company_lookup.py "Tata Consultancy Services"
python company_lookup.py "Tata Motors"
```
