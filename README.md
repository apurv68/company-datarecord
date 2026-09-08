# Comprehensive Categorized Company Intelligence Record (via DuckDuckGo)

A Python tool that takes any company name as input, discovers similar/matching company candidates for ambiguous searches, and automatically retrieves corporate intelligence grouped into **6 distinct categories organized by Year-Wise Timeline**.

---

## Key Features

1. **Smart Company Disambiguation & Candidate Selection:**
   - If you enter an ambiguous or short name (e.g. `Tata`, `Apple`, `Reliance`, `Apex`), the script presents a **numbered list of similar/matching companies** with brief descriptions.
   - You can choose the exact entity you want (e.g., `Tata Motors` vs `Tata Steel` vs `TCS`) or proceed with the exact entered text.

2. **📅 Year-Wise Timeline Categorization:**
   - Every category automatically parses dates and fiscal years (`FY26`, `FY25`, `2026`, `2025`, `2024`, etc.).
   - Findings are chronologically sorted and tagged:
     - `📅 [2026 (Current Year)]`
     - `📅 [2025 (Previous Year)]`
     - `📅 [2024 / Historical]`
     - `📅 [Recent / Active Records]`

3. **The 6 Corporate Intelligence Categories:**
   - **🏢 Core Corporate Profile:** Official Legal Name, Industry, Ownership Type, Stock Ticker, Founding Date & Location, Founders, Headquarters, Official Website, and Summary Description.
   - **🏞️ Land, Properties & Real Estate (Year-Wise):** Land acquired, parcels sold, corporate campuses, office leases, manufacturing plants, and real estate investments.
   - **📈 Revenue & Financial Performance (YoY Timeline):** Current year annual revenue, previous financial year comparison, quarterly results, profit margins, and YoY growth reports.
   - **👥 Key Employees, Roles & Leadership (Current Head):** Chief Executive Officer (CEO), Chief Financial Officer (CFO), Chief Technology Officer (CTO), Board of Directors, and active management designations (filtering out obsolete/deceased executives).
   - **💼 New Job Openings & Workforce Changes (Year-Wise):** Active career portals, ongoing hiring campaigns, job openings by department, workforce expansions, restructuring, or layoff reports.
   - **⚖️ Legal Information & Lawsuits (Year-Wise):** Pending/settled court disputes, class-action lawsuits, regulatory actions (SEC, CCI, SEBI, FTC), compliance audits, and litigation summaries.

3. **🌐 Multi-Source Web Intelligence & Flexible Discovery:**
   - Automatically draws relevant data from anywhere on the public web (DuckDuckGo, Yahoo, Brave, Google indexing, Wikipedia encyclopedic records, MCA / financial databases like Tofler & Tracxn, LinkedIn, job portals).
   - Intelligently handles company names with apostrophes and punctuation (e.g. `Haldiram's`, `McDonald's`, `Wendy's`) by searching both possessive, root, and corporate entity variations.
   - Automatic query cascading: if strict quoted terms yield zero results for private unlisted companies, it automatically runs natural relaxed queries across search engines to never miss public records.

4. **Strict Proximity & Noise Filtering:**
   - Eliminates false positives and 3rd-party noise (e.g. rejects "apartments near company" or "competitors of company").
   - Filters out past/deceased executives to ensure only current leadership is displayed.

5. **Dual Export with Automatic Sync:**
   - **`company_records.csv`**: Updates existing company rows or appends new ones (ready for Excel/Sheets with clean year tags).
   - **`company_records.json`**: Structured JSON format with nested year-wise timelines and source URLs.

---

## How to Run

### Interactive Menu Mode
```powershell
python company_lookup.py
```
1. Type a company name (e.g., `Tata`, `Infosys`, `Tesla`, `Razorpay`).
2. If similar/matching companies exist, select the number from the list.
3. The script scans DuckDuckGo across all 6 categories, organizes results by **Year**, renders color-coded cards, and saves the data.

### Direct Command-Line Mode
```powershell
python company_lookup.py "Tata Motors"
python company_lookup.py "Tesla"
python company_lookup.py "Infosys"
```
