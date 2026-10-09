import sys
import os
import json
import re
from typing import Optional, Dict, Any, List
from fastapi import FastAPI, Query, HTTPException, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

# Ensure parent directory is on sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from company_lookup import (
    find_company_candidates,
    fetch_table1_data,
    resolve_canonical_entity,
    apply_direct_evidence_filter,
    fetch_table2_data,
    fetch_latest_news,
    fetch_business_activities,
    fetch_strategic_conclusions,
    fetch_peer_comparison,
    save_company_docx,
    EvidenceStore,
    filter_evidence_store_direct_only
)

app = FastAPI(
    title="Corporate Business Intelligence Engine",
    description="Executive Multi-Tier Autonomous Corporate Intelligence API",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.middleware("http")
async def add_no_cache_headers(request: Request, call_next):
    response = await call_next(request)
    if request.url.path.startswith("/static") or request.url.path == "/":
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response

# In-memory store for the most recently analyzed company dossiers (for docx exports)
_CACHE_STORE: Dict[str, Any] = {}

@app.get("/api/status")
def get_system_status():
    """Return live status of AI and statutory integration pipelines."""
    has_gemini = bool(os.environ.get("GEMINI_API_KEY", "").strip())
    has_tavily = bool(os.environ.get("TAVILY_API_KEY", "").strip())
    return {
        "status": "online",
        "integrations": {
            "gemini_ai": {"connected": has_gemini, "mode": "Synthesis & JSON Extraction" if has_gemini else "Fallback Heuristic"},
            "tavily_search": {"connected": has_tavily, "mode": "Anti-Bot Web Intelligence" if has_tavily else "DuckDuckGo Fallback"},
            "screener_exchange": {"connected": True, "mode": "Audited BSE/NSE Public Tables"},
            "wikidata_sparql": {"connected": True, "mode": "Semantic Entity Disambiguation (P31)"},
            "mca_roc_registry": {"connected": True, "mode": "Statutory Indian Corporate Filings"}
        }
    }

@app.get("/api/candidates")
def get_candidates(q: str = Query(..., min_length=1)):
    """Fetch matching company candidates for live autocomplete & disambiguation."""
    try:
        candidates = find_company_candidates(q.strip())
        return {"query": q, "count": len(candidates), "candidates": candidates}
    except Exception as e:
        return {"query": q, "count": 0, "candidates": [], "error": str(e)}

@app.post("/api/lookup")
async def execute_lookup(payload: Dict[str, Any]):
    """
    Execute full 6-table multi-tier intelligence pipeline for target company.
    """
    query = str(payload.get("query", "")).strip()
    selected_cand = payload.get("candidate")
    
    if not query and not selected_cand:
        raise HTTPException(status_code=400, detail="Missing company name or candidate object")

    try:
        # Step 1: Candidate resolution
        if not selected_cand:
            cands = find_company_candidates(query)
            if cands:
                selected_cand = cands[0]
            else:
                selected_cand = {"name": query, "desc": "Direct Query Entity", "is_indian": True}

        company_name = selected_cand.get("name", query)
        evidence_store = EvidenceStore(company_name)

        # Step 2: Table #1 Master Corporate Identity
        data1, sources1 = fetch_table1_data(selected_cand, evidence_store=evidence_store)
        canonical_entity = resolve_canonical_entity(selected_cand, data1, sources1)
        canon_disp = canonical_entity.get("canonical_name", company_name)
        evidence_store.canonical_name = canon_disp
        data1 = apply_direct_evidence_filter(data1, canonical_entity, "Table #1")

        # Step 3: Table #2 5-Year Financials
        ticker = canonical_entity.get("ticker", "N/A")
        data2, sources2 = fetch_table2_data(canonical_entity, ticker, evidence_store=evidence_store)
        data2 = apply_direct_evidence_filter(data2, canonical_entity, "Table #2")

        # Step 4: Table #3 Strategic News Milestones
        data3, sources3 = fetch_latest_news(canonical_entity, evidence_store=evidence_store)
        data3 = apply_direct_evidence_filter(data3, canonical_entity, "Table #3")

        # Step 5: Table #4 Business Activities & Operations
        data4, sources4 = fetch_business_activities(canonical_entity, evidence_store=evidence_store)
        data4 = apply_direct_evidence_filter(data4, canonical_entity, "Table #4")

        # Step 6: Table #5 Strategic Conclusions
        data5, sources5 = fetch_strategic_conclusions(canonical_entity, data1, data2, data3, data4, evidence_store=evidence_store)
        data5 = apply_direct_evidence_filter(data5, canonical_entity, "Table #5")

        # Step 7: Filter Evidence Store
        removed_ev = filter_evidence_store_direct_only(evidence_store, canonical_entity)

        # Step 8: Table #6 Peer Comparison
        data6, sources6 = fetch_peer_comparison(canonical_entity, data2=data2, evidence_store=evidence_store)

        # Cache complete result for docx export
        cache_id = re.sub(r'[^a-zA-Z0-9_]', '_', canon_disp.lower())
        _CACHE_STORE[cache_id] = {
            "canonical_entity": canonical_entity,
            "data1": data1, "sources1": sources1,
            "data2": data2, "sources2": sources2,
            "data3": data3, "sources3": sources3,
            "data4": data4, "sources4": sources4,
            "data5": data5, "sources5": sources5,
            "data6": data6, "sources6": sources6,
            "evidence_store": evidence_store
        }

        # Clean string serializable format for frontend
        t2_rows = data2.get("rows", [])
        t2_sources = data2.get("data_sources", {})
        serializable_sources = {f"{k[0]}|{k[1]}": v for k, v in t2_sources.items()} if isinstance(t2_sources, dict) else {}

        # Convert EvidenceStore records
        ev_records = [
            {
                "table": r.get("table"),
                "category": r.get("category"),
                "metric": r.get("metric_or_event"),
                "fact": r.get("fact"),
                "period": r.get("period"),
                "source_name": r.get("source_name"),
                "source_url": r.get("source_url"),
                "confidence": r.get("confidence"),
                "verified": r.get("verified", False)
            }
            for r in evidence_store.records[:50]
        ]

        response_payload = {
            "success": True,
            "cache_id": cache_id,
            "canonical_entity": canonical_entity,
            "table1": {"data": data1, "sources": sources1},
            "table2": {
                "rows": t2_rows,
                "data_sources": serializable_sources,
                "is_synthetic": data2.get("is_synthetic", {}),
                "sources": sources2
            },
            "table3": {"news": data3, "sources": sources3},
            "table4": {"activities": data4, "sources": sources4},
            "table5": {"conclusions": data5, "sources": sources5},
            "table6": {"peers": data6, "sources": sources6},
            "evidence": ev_records,
            "audit_pruned_count": removed_ev
        }
        return response_payload

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Pipeline processing error: {str(e)}")

@app.get("/api/export/docx/{cache_id}")
def export_docx(cache_id: str):
    """Generate and return official executive Word Brief (.docx)."""
    item = _CACHE_STORE.get(cache_id)
    if not item:
        raise HTTPException(status_code=404, detail="Analysis report expired or not found. Please re-run lookup.")
    
    try:
        doc_filename = f"{cache_id}_Executive_Brief.docx"
        doc_path = os.path.join(BASE_DIR, "scratch", doc_filename)
        os.makedirs(os.path.dirname(doc_path), exist_ok=True)
        
        saved_path = save_company_docx(
            data1=item["data1"], sources1=item["sources1"],
            data2=item["data2"], sources2=item["sources2"],
            data3=item["data3"], sources3=item["sources3"],
            data4=item["data4"], sources4=item["sources4"],
            data5=item["data5"], sources5=item["sources5"],
            data6=item["data6"], sources6=item["sources6"],
            evidence_store=item["evidence_store"],
            output_filepath=doc_path,
            auto_open=False
        )
        
        if saved_path and os.path.isfile(saved_path):
            return FileResponse(
                path=saved_path,
                filename=f"{item['canonical_entity'].get('clean_name', 'Company')}_Executive_Report.docx",
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        raise HTTPException(status_code=500, detail="Failed to compile Word document.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Word export error: {str(e)}")

# Mount static frontend assets
web_static_dir = os.path.join(BASE_DIR, "web", "static")
if os.path.isdir(web_static_dir):
    app.mount("/static", StaticFiles(directory=web_static_dir), name="static")

@app.get("/")
def serve_index():
    index_file = os.path.join(BASE_DIR, "web", "static", "index.html")
    if os.path.isfile(index_file):
        return FileResponse(index_file)
    return {"message": "Web frontend loading. Please build static assets."}

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "0.0.0.0")
    print("\n" + "="*70)
    print("  VANTAGE INTELLIGENCE SYSTEM")
    print(f"  Executive Web Portal running on: http://{host}:{port}")
    print("="*70 + "\n")
    uvicorn.run("web_server:app", host=host, port=port, reload=False)
