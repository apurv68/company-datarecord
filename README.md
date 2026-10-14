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

5. **Structured Evidence Store & Auditability Layer:**
   - **Full Source Traceability:** Every extracted claim, metric, and signal is anchored with its snippet, extraction method, confidence level, timestamp, and source URL.
   - **Generic Entity Match Firewall (`is_entity_match`):** Disambiguates canonical company entities against candidate content without brittle company-specific hardcoding, filtering out homonyms, unrelated peers, and sister subsidiaries.
   - **`evidence_store.json`**: Complete structured evidence repository with unique UUID-backed evidence items mapping back to entity records.
   - **`evidence_store.csv`**: Tabular export of the entire evidence audit trail for downstream compliance and verification.

6. **Multi-Format Export with Automatic Sync:**
   - **`company_records.csv`**: Table #1 corporate profiles and leadership directory.
   - **`company_financials_5yr.csv`**: Table #2 audited 5-year historical and present financials.
   - **`company_conclusions.csv`**: Table #5 strategic conclusions and growth intelligence across all 8 strategic pillars.
   - **`company_records.json`**: Complete hierarchical JSON format integrating all 5 tables and verified source links.
   - **`evidence_store.json`** & **`evidence_store.csv`**: Traceable structured evidence records with source provenance.

---

---

## Executive Web Portal (Browser Interface)

The platform provides a browser-based Executive Workstation powered by FastAPI and modern UI architecture for conducting real-time statutory due diligence, corporate intelligence synthesis, and institutional document export.

### 🌐 Option 1: Live Cloud Deployment
Access the pre-deployed production portal directly without any local installation:  
👉 **[corporate-intelligence-portal.onrender.com](https://corporate-intelligence-portal.onrender.com)**

### 💻 Option 2: Running the Web Portal Locally

1. **Clone & Install Dependencies:**
   ```powershell
   git clone https://github.com/adeeshi-solutions/pilot-business-intelligence-system.git
   cd pilot-business-intelligence-system
   pip install -r requirements.txt
   ```

2. **Configure API Keys (Optional but Recommended):**
   Create a `.env` file or export your environment keys for enhanced web intelligence and LLM synthesis:
   ```ini
   GEMINI_API_KEY=your_google_gemini_api_key
   TAVILY_API_KEY=your_tavily_api_key
   PORT=8000
   ```
   *(Note: The system contains built-in statutory fallback heuristics and operates even without third-party API keys).*

3. **Start the Web Server:**
   ```powershell
   python web_server.py
   ```

4. **Launch the Portal:**
   Open your browser and navigate to:
   ```
   http://localhost:8000
   ```

---

## 🖥️ Web Portal Usage: Step-by-Step Guide

Follow these steps to research any Indian listed or unlisted private enterprise:

### Step 1: Verify System Pipeline Health
- In the top-right header, observe the **Systems Active (5/5)** status pill.
- Click the pill to view live telemetry for all 5 integrated data pipelines:
  - **MCA & RoC Registries:** Statutory government filings for private companies.
  - **BSE & NSE Exchange APIs:** Audited public market metrics and ticker feeds.
  - **Wikidata Knowledge Graph:** Semantic entity disambiguation.
  - **Google Gemini AI Layer:** Synthesis and executive narrative structuring.
  - **Tavily Web Intelligence:** Real-time web and news evidence scraping.

### Step 2: Search Entity or Choose Quick Presets
- **Direct Search:** In the command bar, type the company name, CIN, or exchange ticker (e.g., `Tata Motors`, `Adeeshi Solutions`, `Haldiram Snacks`, `Liberty Shoes`).
- **Live Autocomplete:** The engine displays matching candidate entities with corporate descriptions. Click your target entity from the dropdown to ensure disambiguation.
- **Quick Presets:** Click the **Quick Presets** dropdown to instantly run benchmark queries:
  - *Public Conglomerates:* `Tata Motors Limited` (BSE: 500570), `Liberty Shoes Limited` (NSE: LIBERTSHOE)
  - *Unlisted & FMCG / Retail:* `Haldiram Snacks Private Limited`, `Amul (GCMMF)`
  - *Private Tech & Edge Cases:* `AgreeYa Solutions`
- Click **Analyze** (or press `Enter`).

### Step 3: Monitor Live Pipeline Execution
The visual progress stepper tracks the pipeline in real-time across 5 stages:
1. `Query Normalizer`: Cleans search terms and detects corporate suffixes.
2. `MCA & Exchange Harvesting`: Cross-references ZaubaCorp, IndiaFilings, Screener, and official registries.
3. `Entity Scope Firewall`: Rejects sister subsidiaries, homonyms, or cross-company noise.
4. `5-Yr Financial Standardization`: Harmonizes multi-year balance sheets and P&L statements.
5. `Synthesis & Peer Matrix`: Executes strategic conclusions and benchmark comparisons.

### Step 4: Explore the Intelligence Workstation
Once analysis completes, the top KPI grid displays valuation/market cap, multi-year revenue, net profit margin, and ground-truth confidence rating. Navigate through the left navigation rail:

- 📋 **Master Identity & Leadership:**
  - Corporate Identification Number (CIN), RoC jurisdiction (e.g., RoC-Kanpur), Registration Date, and Company Status.
  - Complete Registered Office Address.
  - Verified Official Website, Email, and Telephone contacts.
  - Board of Directors, Managing Directors, and executive leadership roster.
- 📊 **5-Year Financials & Trends:**
  - Audited balance sheet, revenue, EBITDA, net margin, and growth trajectories.
  - Distinguishes audited exchange figures vs. statutory unlisted filings.
- 📰 **Market Signals & News:**
  - Chronological timeline tagged by fiscal year (`📅 [2026]`, `📅 [2025]`, `📅 [2024 / Historical]`).
  - Categorized under institutional tags (M&A, Capex, Infrastructure, Regulatory, Leadership).
- 🏢 **Operations & Brands:**
  - Business archetypes, retail presence, B2B/B2C activities, and product portfolio.
- 🎯 **Strategic Synthesis & Moats:**
  - Deep-dive analysis across 8 strategic vectors: growth trajectory, expansion, shutdowns, AI & technology transformation, and risk indicators.
- 👥 **Peer Benchmarking Matrix:**
  - Side-by-side financial and market standing comparisons against industry competitors.
- 🛡️ **Audit Evidence Store:**
  - Complete inspection of every fact's source URL, verification badge, timestamp, and extraction method.

### Step 5: Export Institutional Word Report (.docx)
- In the top-right header, click the **Export Brief (.docx)** button.
- The system dynamically generates an official executive corporate dossier formatted with typography, financial tables, statutory registrations, and source bibliographies.

---

## Command-Line Interface (CLI Mode)

### Interactive Mode
```powershell
python company_lookup.py
```
1. Type a company name (e.g., `Adani Energy Solutions`, `Tata Consultancy Services`, `Reliance Jio`, `Amul`, `Adeeshi Solutions`).
2. If similar/matching companies exist, select the number from the list or press enter to proceed with the exact entered text.
3. The engine scans multi-source search backends, classifies signals, renders formatted tables, and exports records to CSV and JSON.

### Direct Command-Line Mode
```powershell
python company_lookup.py "Adani Energy Solutions"
python company_lookup.py "Tata Consultancy Services"
python company_lookup.py "Tata Motors"
python company_lookup.py "Adeeshi Solutions"
```

