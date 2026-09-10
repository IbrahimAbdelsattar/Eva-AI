"""Generate the Cloud Chatbot project plan workbook (docs/Cloud_Chatbot_Project_Plan.xlsx).

The plan mirrors the Eva-AI architecture (evidence-grounded RAG, fail-closed
provenance, staged delivery) re-targeted at a general cloud-documentation chatbot.
Re-run this script after editing to regenerate the workbook:

    python scripts/generate_cloud_chatbot_plan.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- design tokens
NAVY = "0D2440"
SAPPHIRE = "2E5E99"
STEEL = "7BA4D0"
ICE = "E7F0FA"
WHITE = "FFFFFF"
BAND = "F5F9FD"

TITLE_FONT = Font(name="Calibri", size=20, bold=True, color=WHITE)
SUB_FONT = Font(name="Calibri", size=11, color=ICE)
H1_FONT = Font(name="Calibri", size=14, bold=True, color=NAVY)
HEAD_FONT = Font(name="Calibri", size=11, bold=True, color=WHITE)
CELL_FONT = Font(name="Calibri", size=10, color="1A1A1A")
MUTED_FONT = Font(name="Calibri", size=9, italic=True, color="5A6B7D")

HEAD_FILL = PatternFill("solid", fgColor=SAPPHIRE)
TITLE_FILL = PatternFill("solid", fgColor=NAVY)
BAND_FILL = PatternFill("solid", fgColor=BAND)
ICE_FILL = PatternFill("solid", fgColor=ICE)

THIN = Side(style="thin", color="C9D8E8")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

WRAP_TOP = Alignment(wrap_text=True, vertical="top")
WRAP_CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")

STATUS_FILLS = {
    "Done": PatternFill("solid", fgColor="D6F0DC"),
    "In Progress": PatternFill("solid", fgColor="FFF1CC"),
    "Blocked": PatternFill("solid", fgColor="FADBD8"),
    "Not Started": PatternFill("solid", fgColor="ECEFF3"),
}
PRIORITY_FILLS = {
    "P0": PatternFill("solid", fgColor="F8CFCB"),
    "P1": PatternFill("solid", fgColor="FDE3C8"),
    "P2": PatternFill("solid", fgColor="FFF6CC"),
    "P3": PatternFill("solid", fgColor="E6EEF7"),
}

STATUSES = ["Not Started", "In Progress", "Blocked", "Done"]


# ------------------------------------------------------------------- utilities
def sheet(wb: Workbook, name: str, tab: str = SAPPHIRE):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    return ws


def banner(ws, title: str, subtitle: str, span: int) -> int:
    """Write the two-line page banner. Returns the next free row."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
    c = ws.cell(row=1, column=1, value=title)
    c.font = TITLE_FONT
    c.fill = TITLE_FILL
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 34
    s = ws.cell(row=2, column=1, value=subtitle)
    s.font = SUB_FONT
    s.fill = TITLE_FILL
    s.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[2].height = 22
    for col in range(2, span + 1):
        ws.cell(row=1, column=col).fill = TITLE_FILL
        ws.cell(row=2, column=col).fill = TITLE_FILL
    return 4


def table(ws, start_row: int, headers: list[str], rows: list[list], widths: list[int],
          heights: int = 30, filter_on: bool = False, freeze: bool = False) -> int:
    """Render a styled header + body table. Returns the row after the table."""
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=i, value=h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = WRAP_CENTER
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.row_dimensions[start_row].height = 26

    for r, row in enumerate(rows, start=start_row + 1):
        for i, val in enumerate(row, start=1):
            c = ws.cell(row=r, column=i, value=val)
            c.font = CELL_FONT
            c.alignment = WRAP_TOP
            c.border = BORDER
            if (r - start_row) % 2 == 0:
                c.fill = BAND_FILL
        ws.row_dimensions[r].height = heights

    end = start_row + len(rows)
    if filter_on and rows:
        ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(headers))}{end}"
    if freeze:
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    return end + 2


def section(ws, row: int, text: str, span: int) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = H1_FONT
    c.fill = ICE_FILL
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 24
    for col in range(2, span + 1):
        ws.cell(row=row, column=col).fill = ICE_FILL
    return row + 1


def paint(ws, col_idx: int, first: int, last: int, mapping: dict) -> None:
    for r in range(first, last + 1):
        v = ws.cell(row=r, column=col_idx).value
        if v in mapping:
            ws.cell(row=r, column=col_idx).fill = mapping[v]
            ws.cell(row=r, column=col_idx).alignment = WRAP_CENTER


def validate(ws, col_letter: str, first: int, last: int, options: list[str]) -> None:
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}{first}:{col_letter}{last}")


# ===================================================================== content
PROJECT = "Cloud Chatbot"
TAGLINE = ("Evidence-grounded RAG assistant over a registered cloud-documentation corpus — "
           "architecture ported from Eva-AI (Clinical Decision Support Lite).")


def build_cover(wb):
    ws = sheet(wb, "00 · Cover", NAVY)
    row = banner(ws, f"{PROJECT} — Project Plan", TAGLINE, 4)

    row = section(ws, row, "Project Charter", 4)
    charter = [
        ["Project name", PROJECT, "Customize", "Rename here and in the Cover banner if the product name changes."],
        ["One-sentence scope", "This system answers questions about <YOUR CLOUD PLATFORM / PRODUCT> using the officially registered documentation corpus, and refuses anything outside it.", "Customize", "Principle IV requires one declared scope sentence before implementation starts."],
        ["Problem statement", "Users cannot find authoritative answers in sprawling cloud documentation; general LLMs answer fluently but without traceable sources or version awareness.", "Customize", "Replace with your own problem framing."],
        ["Solution", "A retrieval-first chatbot: hybrid search over a provenance-registered corpus, LLM synthesis gated on retrieved evidence, inline citations with page/section trace-back, and fail-closed refusal.", "Fixed", "This is the Eva-AI architectural contract being reused."],
        ["Primary users", "Platform engineers, DevOps / SRE, solution architects, support engineers, new-hire onboarding.", "Customize", "Drives the user stories on sheet 05."],
        ["Corpus", "Registered cloud platform docs, architecture guides, runbooks, release notes (PDF / HTML / Markdown).", "Customize", "Every file must be registered in data/sources.yaml before ingestion."],
        ["Out of scope", "Live account actions, billing changes, executing infrastructure commands, security incident response, anything not in the registered corpus.", "Customize", "Out-of-scope questions must be refused, not guessed."],
        ["Delivery model", "Monolithic repo · FastAPI backend + Next.js frontend · single-process production deploy (backend serves the static frontend export).", "Fixed", "Same topology as Eva-AI."],
        ["Duration", "9 phases · ~12 weeks (see sheet 06 Roadmap and sheet 15 Timeline).", "Customize", "Adjust the week columns on the Timeline sheet."],
        ["Plan owner", "<YOUR NAME>", "Customize", "Add teammates on the WBS 'Owner' column."],
        ["Version", "v1.0 — generated from scripts/generate_cloud_chatbot_plan.py", "Fixed", "Re-run the script to regenerate this workbook."],
    ]
    row = table(ws, row, ["Field", "Value", "Type", "Notes"], charter, [24, 74, 14, 56], heights=42)

    row = section(ws, row, "How to use this workbook", 4)
    guide = [
        ["00 · Cover", "Charter, sheet map, how to customize.", "Read first"],
        ["01 · Scope", "In/out of scope, assumptions, constraints, definition of done.", "Fill before Phase 0"],
        ["02 · Constitution", "The 6 non-negotiable principles every task must obey.", "Do not weaken"],
        ["03 · Tech Stack", "Chosen technologies, versions, and the swap-out seam for each.", "Customize versions"],
        ["04 · Architecture", "Layer-by-layer component inventory with target file paths.", "Mirrors Eva-AI"],
        ["05 · User Stories", "US1–US12 with priority and acceptance criteria.", "Customize"],
        ["06 · Roadmap", "Phase 0–8 with goals, exit gates, and status roll-up.", "Track weekly"],
        ["07 · WBS Tasks", "The full task backlog — the working sheet. Filter by Phase / Owner / Status.", "Daily driver"],
        ["08 · API Contracts", "Every endpoint, its payload, and its failure behaviour.", "Customize"],
        ["09 · Data Model", "Chunk / Source / Response schemas and metadata rules.", "Mirrors Eva-AI"],
        ["10 · Config & Env", "Every environment variable with default and purpose.", "Customize"],
        ["11 · Test Matrix", "Unit / integration / eval / manual checks per phase.", "Gate each phase"],
        ["12 · KPIs", "Quality, latency, safety, and adoption targets.", "Set your numbers"],
        ["13 · Risks", "Risk register with likelihood, impact, and mitigation owner.", "Review weekly"],
        ["14 · Deliverables", "File-by-file manifest of what each phase produces.", "Check off"],
        ["15 · Timeline", "12-week phase grid (Gantt-style).", "Shift as needed"],
        ["16 · Glossary", "Shared vocabulary for the team.", "Reference"],
    ]
    row = table(ws, row, ["Sheet", "What it holds", "Action"], guide, [22, 92, 24], heights=20)

    c = ws.cell(row=row, column=1,
                value="Cells marked 'Customize' are placeholders for the Cloud Chatbot domain. "
                      "Everything marked 'Fixed' is the Eva-AI architectural contract — changing it changes the safety model.")
    c.font = MUTED_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    return ws


def build_scope(wb):
    ws = sheet(wb, "01 · Scope")
    row = banner(ws, "Scope, Assumptions & Constraints", "What the chatbot will and will not do — agreed before any code is written.", 4)

    row = section(ws, row, "In Scope", 4)
    in_scope = [
        ["SC-01", "Answer natural-language questions from the registered documentation corpus only", "P0", "Every answer carries at least one citation."],
        ["SC-02", "Multi-turn conversation with context carried across turns", "P0", "Elliptical follow-ups ('and for the EU region?') resolve correctly."],
        ["SC-03", "Inline citations with document, page/anchor, and section trace-back", "P0", "Click a citation → see the verbatim source chunk."],
        ["SC-04", "Evidence inspector showing every retrieved chunk and its relevance score", "P1", "Weak evidence is shown flagged, never hidden."],
        ["SC-05", "Refusal on out-of-scope, unsupported, or adversarial questions", "P0", "Fail-closed: no answer without evidence."],
        ["SC-06", "Streaming responses (SSE) with sub-2s time-to-first-token", "P1", "Measured at P95."],
        ["SC-07", "Conversation history, session search, and markdown export", "P2", "Client-side persistence."],
        ["SC-08", "Source registry with publisher, version, URL, and credibility note", "P0", "Ingestion aborts on unregistered files."],
        ["SC-09", "Admin CLI for ingest / query / eval / index inspection", "P1", "One command per operation."],
        ["SC-10", "Observability: tracing, metrics, error tracking, PII scrubbing", "P1", "Full-stack."],
        ["SC-11", "Light / dark theming and responsive layout", "P2", "Persisted per user."],
        ["SC-12", "Optional second UI language with RTL support", "P3", "Drop if not needed."],
    ]
    row = table(ws, row, ["ID", "Capability", "Priority", "Definition of done"], in_scope, [10, 66, 12, 66], heights=28)

    row = section(ws, row, "Out of Scope (explicit non-goals)", 4)
    out = [
        ["OS-01", "Executing actions against live cloud accounts (deploy, scale, delete)", "Read-only advisory system; action execution is a different trust model."],
        ["OS-02", "Answering from model parametric memory when retrieval finds nothing", "Violates Principle I — the system refuses instead."],
        ["OS-03", "Ingesting private customer data, credentials, or secrets", "Violates Principle III — corpus is public/approved documentation only."],
        ["OS-04", "General-purpose assistant behaviour outside the declared scope", "Violates Principle IV — narrow scope is what makes evaluation possible."],
        ["OS-05", "Multi-tenant account isolation and per-customer corpora", "Deferred to a post-v1 phase; single corpus in v1."],
        ["OS-06", "Fine-tuning or training a custom model", "Retrieval quality, not model weights, is the lever in v1."],
        ["OS-07", "Mobile native apps", "Responsive web only."],
    ]
    row = table(ws, row, ["ID", "Non-goal", "Why"], out, [10, 66, 78], heights=28)

    row = section(ws, row, "Assumptions & Constraints", 4)
    ac = [
        ["AS-01", "Assumption", "The documentation corpus is legally usable and publicly available or internally approved for indexing.", "Legal / Owner"],
        ["AS-02", "Assumption", "Source documents have a real text layer (not scanned images).", "Data owner"],
        ["AS-03", "Assumption", "An LLM gateway API key with sufficient quota is available for generation.", "Owner"],
        ["AS-04", "Assumption", "Docs are versioned; each corpus entry records the product version it describes.", "Data owner"],
        ["CN-01", "Constraint", "Every retrieval must return metadata + score in one call — no sidecar metadata store.", "Principle II"],
        ["CN-02", "Constraint", "Generation may not ship before retrieval quality passes its golden-set gate.", "Principle V"],
        ["CN-03", "Constraint", "Local-first embeddings preferred; remote embedding is a fallback, not the default path.", "Cost / latency"],
        ["CN-04", "Constraint", "Single-process production deployment (one container).", "Ops simplicity"],
        ["CN-05", "Constraint", "No PII/secret values may reach logs, traces, or error reports.", "Privacy"],
    ]
    row = table(ws, row, ["ID", "Type", "Statement", "Owner / Source"], ac, [10, 14, 96, 24], heights=28)
    return ws


def build_constitution(wb):
    ws = sheet(wb, "02 · Constitution", NAVY)
    row = banner(ws, "Project Constitution — Non-Negotiable Principles",
                 "Ported verbatim in spirit from Eva-AI. Every spec, task, and PR is checked against these.", 5)
    rows = [
        ["I", "Evidence-Grounded Answers Only", "NON-NEGOTIABLE",
         "The chatbot answers exclusively from text retrieved out of the registered corpus. Model parametric memory is never a permitted source. If retrieval returns nothing above the relevance floor, the system says it cannot find supporting evidence — it does not answer anyway.",
         "No code path reaches the LLM answer template without retrieved chunks attached; grounding tests assert refusal on empty evidence."],
        ["II", "Citation Metadata Is Structural", "NON-NEGOTIABLE",
         "Every chunk carries document_name, page_number/anchor, section_title, chunk_id, source_url, version and text — stored WITH the vector entry. Sidecar files that can drift out of sync are forbidden. Retrieval returns metadata and score in a single call.",
         "Schema test asserts metadata completeness = 100% on every indexed chunk; citation-integrity test asserts every rendered citation resolves to a real chunk."],
        ["III", "Source Legitimacy & Provenance", "NON-NEGOTIABLE",
         "Only approved, legally usable documents enter the corpus. Every source is registered with publisher, version, publication date, retrieval URL, document type, and a written credibility justification. Ingestion fails closed on any unregistered file.",
         "Registry validator aborts with a non-zero exit code naming the offending file; sources older than the staleness window are flagged and surfaced in the UI."],
        ["IV", "Narrow Scope Discipline", "ENFORCED",
         "The system serves one declared documentation scope, stated in one sentence before implementation. Out-of-scope, off-topic, and adversarial questions are refused. Every response carries an advisory disclaimer.",
         "Scope classifier threshold calibrated against an out-of-scope noise set; 100% separation target on the calibration set."],
        ["V", "Staged Delivery — Retrieval Before Generation", "ENFORCED",
         "Ingestion, chunking, metadata and indexing are complete and verified before generation work begins. Prompt engineering and UI polish are forbidden until retrieval demonstrably returns good evidence on the evaluation set. Seams for later stages may exist; implementations may not jump the queue.",
         "/api/generate returns 501 Not Implemented until the Phase 2 retrieval gate passes."],
        ["VI", "Human Verification Over Automated Confidence", "ENFORCED",
         "Weak or low-scoring evidence is flagged (below_floor=True) and shown, never silently dropped. Every chunk traces back to the exact page/anchor of the source document so a human can verify it.",
         "Evidence inspector renders every retrieved chunk with its score; no filter removes results from the response payload."],
    ]
    row = table(ws, row, ["#", "Principle", "Status", "Statement", "How it is enforced"],
                rows, [6, 30, 18, 78, 62], heights=92)
    c = ws.cell(row=row, column=1, value="A fluent answer is not a safe answer. Fluency is never accepted as evidence of correctness.")
    c.font = Font(name="Calibri", size=11, bold=True, italic=True, color=SAPPHIRE)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    return ws


def build_stack(wb):
    ws = sheet(wb, "03 · Tech Stack")
    row = banner(ws, "Technology Stack", "Same shape as Eva-AI. The 'Swap seam' column names the protocol that lets you replace each choice.", 6)

    row = section(ws, row, "Backend", 6)
    backend = [
        ["Language", "Python", "3.13", "Runtime", "—", "Match Eva-AI."],
        ["API framework", "FastAPI + Uvicorn", ">=0.115 / >=0.32", "HTTP + SSE streaming", "—", "ASGI, native async, OpenAPI for free."],
        ["Config", "Pydantic v2 + pydantic-settings", ">=2.9 / >=2.6", "Typed settings from env", "—", "One Settings class, cached accessor."],
        ["Doc parsing", "PyMuPDF (fitz)", ">=1.24", "PDF span extraction with font metadata", "Parser interface", "Add an HTML/Markdown parser for web docs."],
        ["Tokenization", "tiktoken (cl100k_base)", "latest", "Chunk token budgeting", "—", "400–800 token target budget."],
        ["Vector store", "ChromaDB (embedded, persistent)", ">=0.5", "Vector + metadata storage", "VectorStore interface", "Swap to pgvector/Qdrant later without touching callers."],
        ["Embeddings (primary)", "Local sentence-transformers (BGE-small-en-v1.5)", ">=3.0", "Fast local embedding (~20–30 ms)", "Embedder protocol", "Local-first: zero remote failure risk."],
        ["Embeddings (fallback)", "Remote gateway embedding model", "—", "Failover when local model unavailable", "Embedder protocol", "Dimension mismatch handled at the store boundary."],
        ["Lexical search", "rank-bm25 (BM25Okapi)", ">=0.2", "Exact-term retrieval (CLI flags, error codes, service names)", "Retriever protocol", "Critical for cloud docs: 'ERR_QUOTA_EXCEEDED', '--max-instances'."],
        ["Fusion", "Reciprocal Rank Fusion (k=60)", "—", "Merge dense + lexical rankings", "Retriever protocol", "No score normalization needed."],
        ["Reranking", "Cross-Encoder (sentence-transformers)", ">=3.0", "Precision refinement of top candidates", "Retriever protocol", "Must degrade gracefully to RRF if weights fail to load."],
        ["Generation LLM", "Claude Sonnet 5 via gateway", "—", "Grounded answer synthesis", "LLM client interface", "Low temperature; strict citation format."],
        ["Rate limiting", "slowapi", "latest", "Per-IP request throttling", "Middleware", "Protects the LLM budget."],
        ["Testing", "pytest + pytest-asyncio", ">=8.3", "Unit / integration / eval suites", "—", "Eval suite is a release gate."],
        ["Linting", "Ruff", "latest", "Lint + format", "—", "Enforced in CI and pre-push hook."],
    ]
    row = table(ws, row, ["Layer", "Technology", "Version", "Purpose", "Swap seam", "Notes"],
                backend, [22, 34, 18, 44, 22, 46], heights=30)

    row = section(ws, row, "Frontend", 6)
    frontend = [
        ["Framework", "Next.js (App Router)", "15.x", "UI shell + static export", "—", "Exports to static assets served by the backend."],
        ["Runtime", "React + TypeScript", "19.x / 5.7", "Component runtime", "—", "Strict TS."],
        ["Styling", "Tailwind CSS + design tokens", "4.x", "Design system", "—", "Soft-UI card/inset/button primitives."],
        ["Streaming", "SSE via fetch ReadableStream", "—", "Token-by-token answer rendering", "useStreamingChat hook", "Cancel-on-unmount."],
        ["State", "React hooks + localStorage", "—", "Session list, active thread, theme", "useChatSessions hook", "Bounded session count with eviction."],
        ["i18n", "Custom dictionary + RTL switch", "—", "Optional second language", "translations.ts", "Drop if single-language."],
        ["Error tracking", "Sentry browser SDK", "latest", "Client error + performance capture", "—", "Scrub message content before send."],
    ]
    row = table(ws, row, ["Layer", "Technology", "Version", "Purpose", "Swap seam", "Notes"],
                frontend, [22, 34, 18, 44, 22, 46], heights=30)

    row = section(ws, row, "Platform, Ops & Observability", 6)
    ops = [
        ["Containerization", "Docker multi-stage build", "—", "One image: frontend build → backend runtime", "—", "Static export copied into the API image."],
        ["Orchestration", "docker-compose (prod + dev)", "—", "Local and single-host deploy", "—", "Volume-mount the index directory."],
        ["CI/CD", "GitHub Actions", "—", "Lint, test, build, publish image", "—", "PR check + main pipeline."],
        ["Registry", "GHCR", "—", "Container image hosting", "—", "Tag per commit + latest."],
        ["Tracing", "LangSmith (or OTel)", "—", "RAG stage-level tracing", "Tracing wrapper", "Trace retrieve → assemble → generate."],
        ["Error tracking", "Sentry (backend + frontend)", "—", "Exceptions, spans, profiling", "—", "PII scrubbing before transport."],
        ["Metrics", "Prometheus-compatible counters/histograms", "—", "Latency, hit rate, refusal rate, cache hit rate", "Metrics module", "Exposed on /metrics."],
        ["Logging", "structlog-style JSON logs", "—", "Correlated request logs", "logging_config", "Request ID on every line."],
    ]
    row = table(ws, row, ["Layer", "Technology", "Version", "Purpose", "Swap seam", "Notes"],
                ops, [22, 34, 18, 44, 22, 46], heights=30)
    return ws


def build_architecture(wb):
    ws = sheet(wb, "04 · Architecture")
    row = banner(ws, "System Architecture & Component Inventory",
                 "Monolithic repo, protocol-driven seams. Dev: two processes (API :8000, UI :3000). Prod: one process serving both.", 5)

    row = section(ws, row, "Request path (top to bottom)", 5)
    flow = [
        ["1", "Client", "Chat UI submits a question with the last N turns of history", "frontend/components/chat/", "—"],
        ["2", "API", "POST /api/chat/stream validates payload, opens an SSE channel", "backend/app/api/generate.py", "Rate limited"],
        ["3", "Guardrail (pre)", "Input risk classifier + greeting/capability router short-circuits non-doc queries", "backend/app/generation/risk_classifier.py", "0 ms path for greetings"],
        ["4", "Query enrichment", "Rewrites elliptical follow-ups using conversation history", "backend/app/retrieval/scope.py", "Multi-turn correctness"],
        ["5", "Retrieval", "Dense + BM25 in parallel → RRF fusion → cross-encoder rerank", "backend/app/retrieval/hybrid.py", "Cache L1/L2"],
        ["6", "Scope gate", "Rejects when best score < scope threshold", "backend/app/retrieval/scope.py", "Principle IV"],
        ["7", "Assembly", "Builds numbered evidence blocks with citation metadata", "backend/app/generation/assembler.py", "Below-floor chunks kept, flagged"],
        ["8", "Prompt", "System prompt + evidence + question + history window", "backend/app/generation/prompt.py", "Strict citation format"],
        ["9", "Generation", "Streams tokens from the LLM gateway", "backend/app/generation/client.py", "Traced"],
        ["10", "Citation anchoring", "Maps [Source N] markers back to real chunk IDs", "backend/app/generation/citations.py", "Invalid markers stripped"],
        ["11", "Guardrail (post)", "Grounding check, disclaimer injection, refusal fallback", "backend/app/generation/guardrails.py", "Principle I"],
        ["12", "Client render", "Streams answer, renders citation chips + evidence accordion", "frontend/components/chat/", "Cache L3 on repeat"],
    ]
    row = table(ws, row, ["#", "Stage", "What happens", "Target module", "Note"], flow, [6, 20, 66, 46, 26], heights=28)

    row = section(ws, row, "Component inventory", 5)
    comps = [
        ["Backend · Core", "main.py", "FastAPI app, CORS, static mount, router wiring", "Phase 0", "New"],
        ["Backend · Core", "config.py", "Typed settings for every env var", "Phase 0", "New"],
        ["Backend · Core", "models.py", "Pydantic schemas — Source, Chunk, Manifest, RetrievalResult, ChatRequest/Response, Citation", "Phase 0", "New"],
        ["Backend · Core", "errors.py", "Typed errors mapped to HTTP status and CLI exit codes", "Phase 0", "New"],
        ["Backend · Ingestion", "registry.py", "Fail-closed source registry validator", "Phase 1", "New"],
        ["Backend · Ingestion", "parser.py", "PDF span extraction with page + font metadata", "Phase 1", "New"],
        ["Backend · Ingestion", "html_parser.py", "HTML/Markdown docs parser with heading anchors", "Phase 1", "Cloud-specific addition"],
        ["Backend · Ingestion", "cleaner.py", "Frequency-based boilerplate removal, glyph/hyphenation repair", "Phase 1", "New"],
        ["Backend · Ingestion", "sectioner.py", "Heading hierarchy detection (H1/H2/H3 or numbered)", "Phase 1", "New"],
        ["Backend · Ingestion", "chunker.py", "Atomic unit preservation + token-budget packing", "Phase 1", "Atomic unit = one procedure/code block"],
        ["Backend · Ingestion", "pipeline.py", "Orchestrator with atomic index swap", "Phase 1", "New"],
        ["Backend · Embeddings", "base.py / local.py / remote.py / fallback.py", "Embedder protocol, local-first with remote failover", "Phase 1", "New"],
        ["Backend · Retrieval", "store.py", "Persistent vector store, cosine space, build-then-swap", "Phase 1", "New"],
        ["Backend · Retrieval", "dense.py", "Dense cosine retriever", "Phase 2", "New"],
        ["Backend · Retrieval", "bm25.py", "Lexical retriever with a technical tokenizer", "Phase 2", "Keeps CLI flags/error codes intact"],
        ["Backend · Retrieval", "hybrid.py", "RRF fusion of dense + lexical", "Phase 2", "New"],
        ["Backend · Retrieval", "reranker.py", "Cross-encoder with graceful fallback", "Phase 2", "New"],
        ["Backend · Retrieval", "factory.py", "Retriever selection by config", "Phase 2", "Dependency injection"],
        ["Backend · Retrieval", "scope.py", "Scope classification + multi-turn query enrichment", "Phase 2", "New"],
        ["Backend · Retrieval", "cache.py", "L1 embedding / L2 retrieval / L3 answer caches, TTL + LRU", "Phase 6", "Invalidated by index manifest hash"],
        ["Backend · Generation", "client.py", "Async LLM client, streaming + retries", "Phase 3", "New"],
        ["Backend · Generation", "assembler.py", "Evidence block assembly", "Phase 3", "New"],
        ["Backend · Generation", "prompt.py", "System prompt and prompt construction", "Phase 3", "New"],
        ["Backend · Generation", "citations.py", "Citation extraction + anchoring", "Phase 3", "New"],
        ["Backend · Generation", "guardrails.py", "Grounding, refusal, disclaimer enforcement", "Phase 3", "New"],
        ["Backend · Generation", "risk_classifier.py", "Pre-retrieval input risk + capability routing", "Phase 3", "New"],
        ["Backend · Generation", "clarification.py", "Ambiguity detection → ask a clarifying question", "Phase 3", "New"],
        ["Backend · API", "api/search.py", "POST /api/search, GET /api/index, GET /api/sources, GET /api/health", "Phase 2", "New"],
        ["Backend · API", "api/generate.py", "POST /api/chat and /api/chat/stream", "Phase 3", "501 stub until the Phase 2 gate passes"],
        ["Backend · Ops", "ratelimit.py", "Per-IP throttling", "Phase 7", "New"],
        ["Backend · Ops", "monitoring/*", "Sentry, tracing, metrics, logging config", "Phase 5", "New"],
        ["Backend · CLI", "cli.py", "ingest / query / chat / eval / index commands", "Phase 1", "Typed exit codes"],
        ["Frontend", "app/page.tsx + layout.tsx", "Chat shell, disclaimer banner, anti-FOUC theme script", "Phase 4", "New"],
        ["Frontend", "components/chat/ChatComposer.tsx", "Input, send, stop-generation, example prompts", "Phase 4", "New"],
        ["Frontend", "components/chat/ChatMessage.tsx", "Message bubble, markdown, citation chips", "Phase 4", "New"],
        ["Frontend", "components/chat/ChatHistory.tsx", "Session list, search, rename, delete", "Phase 4", "localStorage"],
        ["Frontend", "components/chat/EvidencePanel.tsx", "Expandable retrieved-chunk inspector with scores", "Phase 4", "Principle VI"],
        ["Frontend", "components/chat/SessionExport.tsx", "Markdown export of a thread", "Phase 4", "New"],
        ["Frontend", "components/IndexStatus.tsx", "Index health, doc count, corpus versions", "Phase 4", "New"],
        ["Frontend", "components/ThemeToggle.tsx", "Light/dark with persistence", "Phase 4", "New"],
        ["Frontend", "hooks/useStreamingChat.ts", "SSE consumption, abort handling", "Phase 4", "New"],
        ["Frontend", "hooks/useChatSessions.ts", "Bounded session persistence", "Phase 4", "New"],
        ["Frontend", "lib/api.ts", "Typed fetch client", "Phase 4", "New"],
        ["Data", "data/sources.yaml", "Provenance registry", "Phase 1", "Hand-maintained"],
        ["Data", "data/corpus/", "Registered source documents", "Phase 1", "Fail-closed"],
        ["Data", "data/index/", "Vector store + manifest.json", "Phase 1", "Generated"],
        ["Data", "data/eval/", "Golden question set + labelled dataset", "Phase 5", "Release gate"],
    ]
    row = table(ws, row, ["Layer", "Module", "Responsibility", "Phase", "Note"],
                comps, [22, 40, 74, 12, 40], heights=26, filter_on=True)
    return ws


def build_stories(wb):
    ws = sheet(wb, "05 · User Stories")
    row = banner(ws, "User Stories & Acceptance Criteria", "Priority P0 = must ship in v1. Each story is independently testable.", 6)
    stories = [
        ["US1", "P0", "Platform engineer", "Ask a question in natural language and get an answer grounded in the official docs",
         "Answer streams within 2 s TTFT; every claim carries at least one citation; disclaimer visible.", "Phase 3"],
        ["US2", "P0", "Platform engineer", "See exactly which document, section and page each part of the answer came from",
         "Citation chips resolve to the verbatim chunk; page/anchor and section title shown; source URL opens the original doc.", "Phase 3"],
        ["US3", "P0", "Any user", "Get an honest refusal when the docs do not cover my question",
         "Out-of-scope and unsupported questions return a refusal, never a guess; refusal test set passes 100%.", "Phase 3"],
        ["US4", "P0", "Admin", "Ingest the documentation corpus with one command",
         "Non-zero chunk count; 100% metadata completeness; unregistered files abort with a non-zero exit code.", "Phase 1"],
        ["US5", "P1", "Platform engineer", "Ask follow-up questions that depend on earlier turns",
         "'And in the EU region?' resolves against the previous turn's subject; multi-turn eval set passes.", "Phase 3"],
        ["US6", "P1", "Support engineer", "Inspect every retrieved chunk with its relevance score before trusting the answer",
         "Evidence panel lists all retrieved chunks including below-floor ones, flagged not hidden.", "Phase 4"],
        ["US7", "P1", "Quality evaluator", "Benchmark retrieval and answer quality against a golden set",
         "Per-question HIT/MISS report; aggregate hit rate vs target; citation accuracy score; runs in CI.", "Phase 5"],
        ["US8", "P1", "Admin", "Verify the provenance, version and credibility of every indexed source",
         "GET /api/sources lists publisher, version, date, URL, credibility note; stale sources flagged in UI.", "Phase 1"],
        ["US9", "P2", "Any user", "Keep a history of my conversations, search them, and export one",
         "Sessions persist across reloads; search by title/content; markdown export downloads a full thread.", "Phase 4"],
        ["US10", "P2", "SRE", "Trace a slow or failed answer end to end",
         "Every request has a correlation ID; RAG stages appear as spans; errors land in the error tracker with PII scrubbed.", "Phase 5"],
        ["US11", "P2", "Any user", "Get instant answers to repeated common questions",
         "Warm cache returns in <50 ms; cache invalidates automatically when the index changes.", "Phase 6"],
        ["US12", "P3", "Non-English user", "Use the interface in my language with correct RTL layout",
         "Language toggle switches all UI strings and mirrors layout; no untranslated keys.", "Phase 4"],
    ]
    row = table(ws, row, ["ID", "Priority", "Role", "Goal", "Acceptance criteria", "Phase"],
                stories, [8, 10, 22, 56, 74, 12], heights=44, filter_on=True, freeze=True)
    paint(ws, 2, row - len(stories) - 2, row - 2, PRIORITY_FILLS)
    return ws


PHASES = [
    ["P0", "Foundations & Project Setup", "W1",
     "Repo skeleton, typed config, schemas, protocols, health endpoint, CI skeleton.",
     "`GET /api/health` returns 200 with index_ready=false; lint + tests run in CI."],
    ["P1", "Corpus Ingestion & Indexing", "W2–W3",
     "Provenance registry, parsers, cleaner, sectioner, chunker, embeddings, vector store, CLI ingest.",
     "One command indexes the corpus; 100% metadata completeness; zero atomic units split; unregistered file aborts."],
    ["P2", "Retrieval Quality (dense → hybrid → rerank)", "W4–W5",
     "Dense retriever, BM25, RRF fusion, cross-encoder rerank, scope gate, /api/search, golden retrieval eval.",
     "Golden-set hit rate ≥ 85%; exact-term queries return the right page; P95 search latency < 800 ms. GATES PHASE 3."],
    ["P3", "Grounded Generation & Safety", "W6–W7",
     "LLM client, evidence assembly, prompt engineering, citation anchoring, guardrails, refusal, streaming chat API.",
     "Citation accuracy ≥ 90%; refusal suite 100%; zero ungrounded claims in the eval sample; TTFT < 2 s."],
    ["P4", "Chat Experience & Frontend", "W7–W8",
     "Chat shell, streaming UI, citation chips, evidence inspector, session history, export, theming, i18n.",
     "Full multi-turn conversation usable end to end; evidence panel shows every chunk; sessions persist."],
    ["P5", "Observability & Evaluation Harness", "W9",
     "Structured logging, tracing, metrics, error tracking with PII scrubbing, golden datasets, quality reports.",
     "Every request traceable by correlation ID; eval suite runs in CI and blocks on regression."],
    ["P6", "Performance, Caching & Cost Control", "W10",
     "Multi-tier caching, index warm-up, threshold calibration, latency profiling, token/cost budgeting.",
     "Warm-query latency < 50 ms; P95 cold answer < 4 s; cost per answer within budget."],
    ["P7", "Hardening, Security & Deployment", "W11",
     "Rate limiting, auth, CORS, input validation, Docker image, compose, CI/CD pipeline, staging deploy.",
     "Image builds and deploys from CI; abuse tests pass; secrets never logged."],
    ["P8", "Documentation, Launch & Handover", "W12",
     "README, API docs, runbook, corpus update procedure, user guide, launch checklist.",
     "A new engineer can ingest, run, evaluate and deploy the system from the docs alone."],
]


def build_roadmap(wb):
    ws = sheet(wb, "06 · Roadmap")
    row = banner(ws, "Delivery Roadmap — Phase 0 to Phase 8",
                 "Staged delivery (Principle V): each phase has an exit gate that must pass before the next phase starts.", 7)
    start = row
    rows = []
    for pid, name, weeks, goal, gate in PHASES:
        rows.append([pid, name, weeks, goal, gate, "Not Started", ""])
    row = table(ws, row, ["Phase", "Name", "Weeks", "Goal", "Exit gate (must pass)", "Status", "% Complete"],
                rows, [10, 36, 12, 62, 70, 16, 12], heights=52, freeze=True)
    first, last = start + 1, start + len(rows)
    paint(ws, 6, first, last, STATUS_FILLS)
    validate(ws, "F", first, last, STATUSES)
    for r in range(first, last + 1):
        ws.cell(row=r, column=7).alignment = WRAP_CENTER

    row = section(ws, row, "Phase dependency chain", 7)
    deps = [
        ["P0 → P1", "Schemas, config and the Embedder/Retriever protocols must exist before ingestion writes to them."],
        ["P1 → P2", "You cannot measure retrieval quality without an index built from a validated corpus."],
        ["P2 → P3", "HARD GATE (Principle V). Generation stays 501 until the retrieval golden gate passes."],
        ["P3 → P4", "The chat UI is built against a working streaming API, not a mock."],
        ["P3 → P5", "Evaluation needs real answers to score; tracing needs real stages to trace."],
        ["P2/P3 → P6", "Caching and calibration tune a system that already works correctly."],
        ["P4/P6 → P7", "Harden and deploy only what is functionally complete."],
        ["P7 → P8", "Documentation describes the shipped system, not the intended one."],
    ]
    row = table(ws, row, ["Edge", "Why the order is fixed"], deps, [16, 120], heights=24)
    return ws


# ----------------------------------------------------------------- WBS backlog
# [ID, Phase, Workstream, Task, Deliverable / file, Owner role, Est (d), Depends on, Priority, Acceptance check]
TASKS = [
    # ---- Phase 0
    ["T001", "P0", "Setup", "Create the monorepo skeleton: backend/app/{api,ingestion,retrieval,embeddings,generation,monitoring}/ and backend/tests/{unit,integration,eval}/", "Package tree with __init__.py", "Backend", 0.5, "—", "P0", "Import of every package succeeds."],
    ["T002", "P0", "Setup", "Pin backend dependencies", "requirements.txt", "Backend", 0.5, "T001", "P0", "pip install -r requirements.txt succeeds on a clean venv."],
    ["T003", "P0", "Setup", "Initialize the Next.js app with TypeScript + Tailwind and a dev rewrite of /api/* to the backend", "frontend/ + next.config.ts", "Frontend", 0.5, "T001", "P0", "npm run dev serves the shell and proxies /api/health."],
    ["T004", "P0", "Setup", "Write .env.example documenting every configuration variable with its default", ".env.example", "Backend", 0.5, "T002", "P0", "Every field on sheet 10 appears exactly once."],
    ["T005", "P0", "Setup", "Implement the typed Settings class with a cached accessor", "backend/app/config.py", "Backend", 1.0, "T004", "P0", "Unit test asserts defaults and env-var override for each field."],
    ["T006", "P0", "Foundations", "Define all Pydantic schemas: SourceDocument, Chunk, IndexManifest, RetrievalResult, SearchResponse, ChatRequest, ChatResponse, Citation, GoldenQuestion", "backend/app/models.py", "Backend", 1.0, "T005", "P0", "No nullable metadata fields — missing strings default to \"\" for vector-store compatibility."],
    ["T007", "P0", "Foundations", "Define the Embedder protocol (embed_documents, embed_query, model_id, dimension)", "backend/app/embeddings/base.py", "Backend", 0.5, "T006", "P0", "A stub implementation satisfies the protocol under type checking."],
    ["T008", "P0", "Foundations", "Define the Retriever protocol (search(query, top_k) -> list[RetrievalResult])", "backend/app/retrieval/base.py", "Backend", 0.5, "T006", "P0", "Substitution seam: dense/hybrid/rerank are interchangeable."],
    ["T009", "P0", "Foundations", "Define typed errors mapped to HTTP status codes and CLI exit codes", "backend/app/errors.py", "Backend", 0.5, "T006", "P1", "Each error class has a documented exit code."],
    ["T010", "P0", "API", "Create the FastAPI app: router mount, CORS, static mount, GET /api/health", "backend/app/main.py", "Backend", 0.5, "T005", "P0", "GET /api/health returns {status: ok, index_ready: false}."],
    ["T011", "P0", "DevEx", "Set up Ruff lint/format config and a pre-push hook", "pyproject.toml, scripts/hooks/pre-push", "DevOps", 0.5, "T002", "P1", "Hook blocks a push that fails lint or unit tests."],
    ["T012", "P0", "DevEx", "Create the PR-check GitHub Actions workflow (lint + unit tests)", ".github/workflows/pr-check.yml", "DevOps", 0.5, "T011", "P1", "A PR shows a passing check."],
    ["T013", "P0", "Governance", "Write the project constitution file and the one-sentence scope statement", ".specify/memory/constitution.md", "Owner", 0.5, "—", "P0", "All six principles recorded; scope sentence approved by the owner."],

    # ---- Phase 1
    ["T014", "P1", "Provenance", "Author the source registry with publisher, version, date, URL, type and credibility note per document", "data/sources.yaml", "Data", 1.0, "T013", "P0", "Every corpus file has an entry; no placeholder credibility notes."],
    ["T015", "P1", "Provenance", "Implement the fail-closed registry validator (unique doc_id, staleness flag, unregistered-file abort)", "backend/app/ingestion/registry.py", "Backend", 1.0, "T014", "P0", "An unregistered file aborts ingestion with exit code 1 and names the file."],
    ["T016", "P1", "Ingestion", "Implement the PDF parser: span extraction with page number, font size and weight", "backend/app/ingestion/parser.py", "Backend", 1.0, "T015", "P0", "A PDF with no text layer aborts with exit code 2."],
    ["T017", "P1", "Ingestion", "Implement the HTML/Markdown parser for web documentation with heading anchors", "backend/app/ingestion/html_parser.py", "Backend", 1.5, "T016", "P1", "Anchors round-trip: a citation links back to the exact heading on the live page."],
    ["T018", "P1", "Ingestion", "Implement the cleaner: frequency-based boilerplate removal (>60% of pages), hyphenation repair, glyph normalization, nav/front-matter filtering", "backend/app/ingestion/cleaner.py", "Backend", 1.0, "T016", "P0", "Unit tests cover each transform independently."],
    ["T019", "P1", "Ingestion", "Implement the sectioner: heading hierarchy detection and section-title propagation", "backend/app/ingestion/sectioner.py", "Backend", 1.0, "T018", "P0", "Every chunk inherits a non-empty section title."],
    ["T020", "P1", "Ingestion", "Implement the chunker: atomic unit preservation (procedure, code block, table) with 400–800 token packing", "backend/app/ingestion/chunker.py", "Backend", 1.5, "T019", "P0", "Zero atomic units split across chunks; oversized units emitted whole and flagged."],
    ["T021", "P1", "Embeddings", "Implement the local-first embedder (BGE-small or equivalent) with batch support", "backend/app/embeddings/local.py", "ML", 1.0, "T007", "P0", "Median embed latency < 30 ms per query on the target machine."],
    ["T022", "P1", "Embeddings", "Implement the remote gateway embedder with batching and exponential-backoff retries", "backend/app/embeddings/remote.py", "ML", 1.0, "T007", "P1", "429/5xx retried; typed error on exhaustion."],
    ["T023", "P1", "Embeddings", "Implement the fallback embedder that fails over local → remote transparently", "backend/app/embeddings/fallback.py", "ML", 0.5, "T021,T022", "P1", "Dimension mismatch is handled at the store boundary, not silently."],
    ["T024", "P1", "Storage", "Implement the persistent vector store with cosine space and atomic build-then-swap", "backend/app/retrieval/store.py", "Backend", 1.5, "T006", "P0", "A failed ingest leaves the previous index queryable."],
    ["T025", "P1", "Storage", "Write and read the self-describing index manifest (model, dimension, counts, built_at, corpus hash)", "data/index/manifest.json", "Backend", 0.5, "T024", "P0", "Manifest hash changes whenever the corpus changes."],
    ["T026", "P1", "Ingestion", "Implement the pipeline orchestrator wiring registry → parse → clean → section → chunk → embed → store", "backend/app/ingestion/pipeline.py", "Backend", 1.0, "T020,T023,T024", "P0", "Integration test runs the full pipeline on a fixture document."],
    ["T027", "P1", "CLI", "Implement the CLI `ingest` command with --dry-run, --doc-id and --verbose", "backend/app/cli.py", "Backend", 1.0, "T026", "P0", "Dry run reports chunk counts without writing the index."],
    ["T028", "P1", "CLI", "Implement the CLI `index` command reporting manifest and per-document statistics", "backend/app/cli.py", "Backend", 0.5, "T027", "P2", "Output lists documents, chunk counts and versions."],
    ["T029", "P1", "Quality", "Unit tests for cleaner, sectioner and chunker transforms", "backend/tests/unit/", "QA", 1.0, "T020", "P0", "Deterministic assertions, no network."],
    ["T030", "P1", "Quality", "Integration test asserting metadata completeness = 100% across the index", "backend/tests/integration/test_ingest_pipeline.py", "QA", 0.5, "T026", "P0", "Fails if any required metadata field is empty."],

    # ---- Phase 2
    ["T031", "P2", "Retrieval", "Implement the dense cosine retriever with score = 1 - distance and a below_floor flag", "backend/app/retrieval/dense.py", "Backend", 1.0, "T024", "P0", "Weak results are flagged, never dropped (Principle VI)."],
    ["T032", "P2", "Retrieval", "Implement the BM25 lexical retriever with a technical tokenizer that preserves flags, error codes and identifiers", "backend/app/retrieval/bm25.py", "Backend", 1.5, "T031", "P0", "Query '--max-instances' returns the exact reference page as top-1."],
    ["T033", "P2", "Retrieval", "Implement RRF fusion (k=60) over dense + lexical rankings", "backend/app/retrieval/hybrid.py", "Backend", 1.0, "T032", "P0", "Unit test verifies rank math on a fixed fixture."],
    ["T034", "P2", "Retrieval", "Implement the cross-encoder reranker with graceful fallback to RRF when weights are unavailable", "backend/app/retrieval/reranker.py", "ML", 1.5, "T033", "P1", "Deleting the model cache degrades quality but never crashes."],
    ["T035", "P2", "Retrieval", "Implement the retriever factory selecting dense | hybrid | hybrid_rerank from config", "backend/app/retrieval/factory.py", "Backend", 0.5, "T034", "P1", "Switching modes requires no caller changes."],
    ["T036", "P2", "Retrieval", "Implement the scope classifier and threshold gate for out-of-domain questions", "backend/app/retrieval/scope.py", "ML", 1.0, "T033", "P0", "Out-of-scope noise set is rejected at 100% on the calibration set."],
    ["T037", "P2", "Retrieval", "Implement multi-turn query enrichment from conversation history before retrieval", "backend/app/retrieval/scope.py", "ML", 1.0, "T036", "P1", "Elliptical follow-up test set retrieves the correct chunks."],
    ["T038", "P2", "API", "Implement POST /api/search with latency tracking and citation-formatted results", "backend/app/api/search.py", "Backend", 1.0, "T035", "P0", "Response includes results, evidence_found and latency_ms."],
    ["T039", "P2", "API", "Implement GET /api/index and GET /api/sources", "backend/app/api/search.py", "Backend", 0.5, "T038", "P1", "Sources endpoint surfaces the staleness/caution flag."],
    ["T040", "P2", "API", "Stub POST /api/chat with 501 Not Implemented until the Phase 2 gate passes", "backend/app/api/generate.py", "Backend", 0.25, "T038", "P0", "Principle V is observable from the API surface."],
    ["T041", "P2", "Evaluation", "Author the golden question set (30+ questions mapped to expected sections)", "backend/tests/eval/golden_questions.yaml", "QA", 1.5, "T038", "P0", "Questions cover every major documentation area."],
    ["T042", "P2", "Evaluation", "Implement the automated retrieval hit-rate suite with a per-question HIT/MISS report", "backend/tests/eval/test_retrieval_quality.py", "QA", 1.0, "T041", "P0", "PHASE GATE: aggregate hit rate ≥ 85%."],
    ["T043", "P2", "Evaluation", "Sweep chunk-size and top-k configurations and record the results", "docs/CHUNKING_SWEEP.md", "ML", 1.0, "T042", "P2", "Chosen configuration is justified by measured numbers."],
    ["T044", "P2", "Evaluation", "Calibrate relevance_floor and scope_threshold against in-scope and noise sets", "docs/THRESHOLD_CALIBRATION.md", "ML", 1.0, "T042", "P0", "Clean separation between in-scope and out-of-scope score distributions."],
    ["T045", "P2", "Quality", "Integration tests for the search API including latency assertions", "backend/tests/integration/test_search.py", "QA", 0.5, "T038", "P1", "P95 search latency < 800 ms on the fixture index."],

    # ---- Phase 3
    ["T046", "P3", "Generation", "Implement the async LLM client with streaming, retries and token accounting", "backend/app/generation/client.py", "Backend", 1.5, "T042", "P0", "Streaming yields tokens; usage is recorded per request."],
    ["T047", "P3", "Generation", "Implement the evidence assembler producing numbered [Source N] blocks with full metadata", "backend/app/generation/assembler.py", "Backend", 1.0, "T046", "P0", "Below-floor chunks are included only when they are the sole evidence, and are marked."],
    ["T048", "P3", "Generation", "Write the system prompt: grounding rules, citation format, refusal instruction, tone", "backend/app/generation/prompt.py", "ML", 1.5, "T047", "P0", "Prompt forbids outside knowledge explicitly and is version-tagged."],
    ["T049", "P3", "Generation", "Implement citation extraction and anchoring back to real chunk IDs", "backend/app/generation/citations.py", "Backend", 1.0, "T048", "P0", "Markers that do not resolve to a retrieved chunk are stripped, not rendered."],
    ["T050", "P3", "Safety", "Implement the pre-retrieval input risk classifier and capability/greeting router", "backend/app/generation/risk_classifier.py", "ML", 1.0, "T048", "P0", "Greetings answer in 0 ms with no LLM call; risky inputs are refused before retrieval."],
    ["T051", "P3", "Safety", "Implement output guardrails: grounding verification, refusal fallback, mandatory disclaimer", "backend/app/generation/guardrails.py", "ML", 1.5, "T049", "P0", "Empty evidence can never produce a substantive answer."],
    ["T052", "P3", "Safety", "Implement clarification: ask a follow-up question when the query is ambiguous", "backend/app/generation/clarification.py", "ML", 1.0, "T051", "P2", "Ambiguous questions get a clarifying question, not a wrong-branch answer."],
    ["T053", "P3", "Safety", "Build the adversarial prompt-injection and jailbreak test set", "backend/tests/unit/test_grounding_stress.py", "QA", 1.0, "T051", "P0", "Injected instructions inside retrieved text are never executed."],
    ["T054", "P3", "API", "Replace the 501 stub with POST /api/chat (non-streaming)", "backend/app/api/generate.py", "Backend", 1.0, "T051", "P0", "Returns answer, citations, evidence_found, disclaimer, model, latency_ms."],
    ["T055", "P3", "API", "Implement POST /api/chat/stream over SSE with heartbeat and client-abort handling", "backend/app/api/generate.py", "Backend", 1.5, "T054", "P0", "TTFT < 2 s at P95; aborting the client stops server-side generation."],
    ["T056", "P3", "API", "Accept and bound the conversation history window on the request payload", "backend/app/api/generate.py", "Backend", 0.5, "T055", "P1", "History is truncated to the last N turns; oversized payloads are rejected cleanly."],
    ["T057", "P3", "Evaluation", "Author the golden generation set with expected citations", "backend/tests/eval/golden_generation.yaml", "QA", 1.0, "T054", "P0", "Each case names the expected source section."],
    ["T058", "P3", "Evaluation", "Implement the generation quality suite: citation accuracy and grounding score", "backend/tests/eval/test_generation_quality.py", "QA", 1.0, "T057", "P0", "PHASE GATE: citation accuracy ≥ 90%."],
    ["T059", "P3", "Evaluation", "Author the refusal test set (out-of-scope, unanswerable, adversarial)", "backend/tests/eval/refusal_cases.yaml", "QA", 0.5, "T053", "P0", "PHASE GATE: 100% correct refusals."],

    # ---- Phase 4
    ["T060", "P4", "Frontend", "Build the app shell: layout, disclaimer banner, anti-FOUC theme script", "frontend/app/layout.tsx", "Frontend", 1.0, "T055", "P0", "No theme flash on load; disclaimer always visible."],
    ["T061", "P4", "Frontend", "Implement the typed API client for search, chat, stream, index and sources", "frontend/lib/api.ts", "Frontend", 0.5, "T060", "P0", "Types are generated from or checked against the OpenAPI schema."],
    ["T062", "P4", "Frontend", "Implement the useStreamingChat hook (SSE parsing, abort, error states)", "frontend/hooks/useStreamingChat.ts", "Frontend", 1.5, "T061", "P0", "Unmounting mid-stream cancels the request."],
    ["T063", "P4", "Frontend", "Build ChatComposer: input, send, stop generation, example prompts", "frontend/components/chat/ChatComposer.tsx", "Frontend", 1.0, "T062", "P0", "Enter sends, Shift+Enter newlines, stop halts streaming."],
    ["T064", "P4", "Frontend", "Build ChatMessage: markdown rendering, code blocks, inline citation chips", "frontend/components/chat/ChatMessage.tsx", "Frontend", 1.5, "T062", "P0", "Clicking a citation chip opens the matching evidence chunk."],
    ["T065", "P4", "Frontend", "Build EvidencePanel: expandable retrieved chunks with score, document, section and page", "frontend/components/chat/EvidencePanel.tsx", "Frontend", 1.5, "T064", "P0", "Every retrieved chunk is listed, including below-floor ones (Principle VI)."],
    ["T066", "P4", "Frontend", "Implement useChatSessions: bounded localStorage persistence with eviction", "frontend/hooks/useChatSessions.ts", "Frontend", 1.0, "T062", "P1", "Sessions survive reload; storage stays under the size cap."],
    ["T067", "P4", "Frontend", "Build ChatHistory: session list, search, rename, delete", "frontend/components/chat/ChatHistory.tsx", "Frontend", 1.0, "T066", "P1", "Search matches titles and message text."],
    ["T068", "P4", "Frontend", "Build SessionExport: export a thread as markdown with citations", "frontend/components/chat/SessionExport.tsx", "Frontend", 0.5, "T067", "P2", "Exported file includes sources and timestamps."],
    ["T069", "P4", "Frontend", "Build IndexStatus: index health, document count, corpus versions, last built", "frontend/components/IndexStatus.tsx", "Frontend", 0.5, "T061", "P1", "Reflects GET /api/index live."],
    ["T070", "P4", "Frontend", "Build ThemeToggle with persistence", "frontend/components/ThemeToggle.tsx", "Frontend", 0.5, "T060", "P2", "Choice persists across sessions."],
    ["T071", "P4", "Frontend", "Define the design system: color tokens, card/inset/button primitives, typography scale", "frontend/app/globals.css", "Design", 1.5, "T060", "P1", "All components use tokens; no hard-coded hex values."],
    ["T072", "P4", "Frontend", "Responsive layout pass: mobile, tablet, desktop", "frontend/", "Frontend", 1.0, "T065", "P1", "Usable at 375 px width with no horizontal scroll."],
    ["T073", "P4", "Frontend", "Accessibility pass: keyboard navigation, focus states, ARIA labels, contrast", "frontend/", "Frontend", 1.0, "T072", "P1", "Keyboard-only operation of the full chat flow; contrast ≥ 4.5:1."],
    ["T074", "P4", "Frontend", "Optional: second UI language with RTL layout mirroring", "frontend/lib/translations.ts", "Frontend", 1.5, "T072", "P3", "No untranslated keys; layout mirrors correctly."],
    ["T075", "P4", "Frontend", "Error and empty states: no index, backend down, refusal, network loss", "frontend/", "Frontend", 1.0, "T064", "P1", "Every failure has a specific, actionable message."],

    # ---- Phase 5
    ["T076", "P5", "Observability", "Structured JSON logging with a correlation ID on every request", "backend/app/monitoring/logging_config.py", "Backend", 1.0, "T055", "P1", "One request ID traces every log line for that request."],
    ["T077", "P5", "Observability", "Distributed tracing over the RAG stages (enrich, retrieve, assemble, generate)", "backend/app/monitoring/tracing.py", "Backend", 1.5, "T076", "P1", "A slow answer can be attributed to a stage from the trace alone."],
    ["T078", "P5", "Observability", "Error tracking on backend and frontend with PII/secret scrubbing before transport", "backend/app/monitoring/sentry.py, frontend/sentry.*.config.ts", "DevOps", 1.5, "T076", "P0", "Scrubbing test asserts no message content or key material leaves the process."],
    ["T079", "P5", "Observability", "Metrics: latency histograms, retrieval scores, refusal rate, cache hit rate, token spend", "backend/app/monitoring/metrics.py", "Backend", 1.0, "T077", "P1", "Exposed on /metrics and scraped in staging."],
    ["T080", "P5", "Observability", "Enhanced health endpoint with dependency checks (index, embedder, LLM gateway)", "backend/app/api/search.py", "Backend", 0.5, "T079", "P1", "Degraded dependencies are reported, not hidden."],
    ["T081", "P5", "Evaluation", "Build the labelled evaluation dataset with a dataset card", "data/eval/", "QA", 1.5, "T058", "P1", "Dataset card records how items were sourced and labelled."],
    ["T082", "P5", "Evaluation", "Wire the evaluation suite into CI as a release gate with a regression threshold", ".github/workflows/ci-cd.yml", "DevOps", 1.0, "T081", "P0", "A quality regression fails the pipeline."],
    ["T083", "P5", "Evaluation", "Write the quality report: metrics, failure taxonomy, improvement backlog", "docs/EVALUATION_REPORT.md", "QA", 1.0, "T082", "P1", "Every known failure mode has an owner and a next action."],

    # ---- Phase 6
    ["T084", "P6", "Performance", "Implement the L1 embedding cache (query → vector) with TTL and LRU eviction", "backend/app/retrieval/cache.py", "Backend", 1.0, "T038", "P1", "Repeat query skips embedding entirely."],
    ["T085", "P6", "Performance", "Implement the L2 retrieval cache (query → results) keyed by the index manifest hash", "backend/app/retrieval/cache.py", "Backend", 1.0, "T084", "P1", "Re-ingesting the corpus invalidates the cache automatically."],
    ["T086", "P6", "Performance", "Implement the L3 answer cache (query + history hash → answer + citations)", "backend/app/retrieval/cache.py", "Backend", 1.0, "T085", "P1", "Warm answer returns in < 50 ms."],
    ["T087", "P6", "Performance", "Warm the BM25 index and embedding model at startup", "backend/app/main.py", "Backend", 0.5, "T085", "P1", "First real request is not the one that pays the load cost."],
    ["T088", "P6", "Performance", "Profile and optimize the retrieval path; record before/after numbers", "docs/PERFORMANCE.md", "Backend", 1.0, "T087", "P1", "P95 cold answer < 4 s; documented measurements."],
    ["T089", "P6", "Cost", "Add token accounting and a per-answer cost budget with alerting", "backend/app/monitoring/metrics.py", "Backend", 1.0, "T079", "P1", "Cost per answer is visible per day and per user."],
    ["T090", "P6", "Cost", "Tune the model routing: cheap model for routing/classification, strong model for synthesis", "backend/app/generation/client.py", "ML", 1.0, "T089", "P2", "Measured cost reduction with no quality regression on the eval set."],

    # ---- Phase 7
    ["T091", "P7", "Security", "Implement per-IP rate limiting on all generation endpoints", "backend/app/ratelimit.py", "Backend", 1.0, "T055", "P0", "Burst abuse returns 429 with a Retry-After header."],
    ["T092", "P7", "Security", "Implement API-key or SSO authentication for non-public deployments", "backend/app/api/", "Backend", 1.5, "T091", "P1", "Unauthenticated requests are rejected in production mode."],
    ["T093", "P7", "Security", "Lock down CORS to the production origins", "backend/app/main.py", "Backend", 0.25, "T092", "P0", "Wildcard origins are rejected by a config test."],
    ["T094", "P7", "Security", "Harden input validation: payload size caps, history depth cap, content-type checks", "backend/app/api/", "Backend", 0.5, "T092", "P0", "Fuzzed payloads return 4xx, never 5xx."],
    ["T095", "P7", "Security", "Secrets audit: no keys in logs, traces, error reports or the client bundle", "repo-wide", "DevOps", 0.5, "T078", "P0", "Secret scanning passes on the whole history."],
    ["T096", "P7", "Deployment", "Write the multi-stage Dockerfile (frontend build → backend runtime with static export)", "Dockerfile", "DevOps", 1.0, "T072", "P0", "Single container serves both UI and API."],
    ["T097", "P7", "Deployment", "Write docker-compose for production and development with an index volume", "docker-compose.yml, docker-compose.dev.yml", "DevOps", 0.5, "T096", "P0", "docker compose up brings the whole system up from scratch."],
    ["T098", "P7", "Deployment", "Write the entrypoint: config validation, index presence check, ingest-on-empty option", "docker-entrypoint.sh", "DevOps", 0.5, "T097", "P1", "Misconfiguration fails fast with a clear message."],
    ["T099", "P7", "Deployment", "Build the CI/CD pipeline: test → build → push image → deploy to staging", ".github/workflows/ci-cd.yml", "DevOps", 1.5, "T098", "P0", "A merge to main lands in staging with no manual step."],
    ["T100", "P7", "Deployment", "Deploy to staging and run the full manual verification checklist", "staging", "DevOps", 1.0, "T099", "P0", "Every row of sheet 11 (manual checks) passes on staging."],
    ["T101", "P7", "Deployment", "Define the backup and restore procedure for the index and the corpus", "docs/RUNBOOK.md", "DevOps", 0.5, "T100", "P1", "A restore has actually been rehearsed once."],

    # ---- Phase 8
    ["T102", "P8", "Docs", "Write the README: architecture, quickstart, configuration, diagrams", "README.md", "Owner", 1.5, "T100", "P0", "A new engineer runs the system from the README alone."],
    ["T103", "P8", "Docs", "Publish API documentation with real request/response examples", "docs/API.md + OpenAPI", "Backend", 1.0, "T102", "P1", "Every endpoint on sheet 08 is documented."],
    ["T104", "P8", "Docs", "Write the operations runbook: alerts, common failures, rollback, on-call notes", "docs/RUNBOOK.md", "DevOps", 1.0, "T101", "P1", "Each alert has a documented response."],
    ["T105", "P8", "Docs", "Document the corpus update procedure: adding, versioning and re-indexing documents", "docs/CORPUS_GOVERNANCE.md", "Data", 1.0, "T102", "P0", "A doc-version bump is a routine, documented operation."],
    ["T106", "P8", "Docs", "Write the end-user guide: what the chatbot can and cannot do, how to read citations", "docs/USER_GUIDE.md", "Owner", 1.0, "T102", "P1", "Sets accurate expectations about refusals."],
    ["T107", "P8", "Launch", "Run the launch checklist: monitoring live, alerts firing, rate limits set, rollback tested", "docs/LAUNCH_CHECKLIST.md", "Owner", 0.5, "T104", "P0", "Every item is signed off before go-live."],
    ["T108", "P8", "Launch", "Set up a user-feedback loop feeding failures back into the golden set", "feedback pipeline", "Owner", 1.0, "T107", "P1", "Reported bad answers become new eval cases."],
    ["T109", "P8", "Handover", "Record the architecture decision log and the known-limitations list", "docs/DECISIONS.md", "Owner", 0.5, "T102", "P2", "Every non-obvious choice has a recorded rationale."],
]


def build_wbs(wb):
    ws = sheet(wb, "07 · WBS Tasks", STEEL)
    row = banner(ws, "Work Breakdown Structure — Task Backlog",
                 f"{len(TASKS)} tasks across 9 phases. Filter by Phase, Workstream, Owner or Status. "
                 "Set Status from the dropdown; % Complete rolls up on the Roadmap sheet.", 12)
    headers = ["ID", "Phase", "Workstream", "Task", "Deliverable / file", "Owner", "Est (d)",
               "Depends on", "Priority", "Status", "Acceptance check", "Notes"]
    rows = [[t[0], t[1], t[2], t[3], t[4], t[5], t[6], t[7], t[8], "Not Started", t[9], ""] for t in TASKS]
    start = row
    row = table(ws, row, headers, rows,
                [8, 8, 16, 66, 40, 12, 9, 14, 10, 14, 62, 26],
                heights=32, filter_on=True, freeze=True)
    first, last = start + 1, start + len(rows)
    paint(ws, 10, first, last, STATUS_FILLS)
    paint(ws, 9, first, last, PRIORITY_FILLS)
    validate(ws, "J", first, last, STATUSES)
    for r in range(first, last + 1):
        ws.cell(row=r, column=7).alignment = WRAP_CENTER
        ws.cell(row=r, column=2).alignment = WRAP_CENTER

    # Summary block
    row += 1
    row = section(ws, row, "Effort summary by phase", 12)
    head = row
    for i, h in enumerate(["Phase", "Name", "Tasks", "Effort (days)", "Done", "% Complete"], start=1):
        c = ws.cell(row=head, column=i, value=h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = WRAP_CENTER
        c.border = BORDER
    for j, (pid, name, *_rest) in enumerate(PHASES):
        r = head + 1 + j
        ws.cell(row=r, column=1, value=pid)
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value=f'=COUNTIF($B${first}:$B${last},$A{r})')
        ws.cell(row=r, column=4, value=f'=SUMIF($B${first}:$B${last},$A{r},$G${first}:$G${last})')
        ws.cell(row=r, column=5, value=f'=COUNTIFS($B${first}:$B${last},$A{r},$J${first}:$J${last},"Done")')
        ws.cell(row=r, column=6, value=f'=IF(C{r}=0,0,E{r}/C{r})')
        ws.cell(row=r, column=6).number_format = "0%"
        for i in range(1, 7):
            c = ws.cell(row=r, column=i)
            c.font = CELL_FONT
            c.border = BORDER
            if i != 2:
                c.alignment = WRAP_CENTER
    total = head + 1 + len(PHASES)
    ws.cell(row=total, column=2, value="TOTAL")
    ws.cell(row=total, column=3, value=f"=SUM(C{head + 1}:C{total - 1})")
    ws.cell(row=total, column=4, value=f"=SUM(D{head + 1}:D{total - 1})")
    ws.cell(row=total, column=5, value=f"=SUM(E{head + 1}:E{total - 1})")
    ws.cell(row=total, column=6, value=f"=IF(C{total}=0,0,E{total}/C{total})")
    ws.cell(row=total, column=6).number_format = "0%"
    for i in range(1, 7):
        c = ws.cell(row=total, column=i)
        c.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
        c.fill = ICE_FILL
        c.border = BORDER
        if i != 2:
            c.alignment = WRAP_CENTER
    return ws


def build_api(wb):
    ws = sheet(wb, "08 · API Contracts")
    row = banner(ws, "API Contracts", "Every endpoint, its payload, and — most importantly — how it fails.", 6)
    rows = [
        ["GET", "/api/health", "Liveness + index readiness", "—",
         "{ status, index_ready, version, dependencies{} }", "Never 5xx; degraded dependencies reported in the body."],
        ["GET", "/api/index", "Index manifest and per-document statistics", "—",
         "{ built_at, embedding_model, dimension, chunk_count, documents[] }", "404-style empty manifest when no index exists."],
        ["GET", "/api/sources", "Registered corpus provenance", "—",
         "{ sources[]: { doc_id, title, publisher, version, published_at, url, type, credibility_note, requires_caution } }", "Stale sources carry requires_caution=true."],
        ["POST", "/api/search", "Retrieval only — no generation", "{ query, top_k?, doc_filter? }",
         "{ results[]: { chunk_id, text, score, below_floor, document_name, section_title, page, url }, evidence_found, latency_ms }", "Empty corpus → evidence_found=false with an empty result list, not an error."],
        ["POST", "/api/chat", "Grounded answer, non-streaming", "{ message, history[]?, top_k? }",
         "{ answer, citations[], evidence_found, refused, refusal_reason?, disclaimer, model, latency_ms, usage{} }", "No evidence → refused=true and a refusal message; never a fabricated answer."],
        ["POST", "/api/chat/stream", "Grounded answer over SSE", "{ message, history[]?, top_k? }",
         "SSE events: meta (retrieved evidence) → token* → citations → done | error", "Client disconnect cancels generation server-side; heartbeat keeps proxies open."],
        ["POST", "/api/feedback", "User rating of an answer", "{ message_id, rating, comment? }",
         "{ ok: true }", "Feedback failures never break the chat flow."],
        ["GET", "/metrics", "Prometheus metrics", "—", "text/plain exposition", "Not exposed publicly in production."],
    ]
    row = table(ws, row, ["Method", "Path", "Purpose", "Request", "Response", "Failure behaviour"],
                rows, [10, 24, 34, 40, 72, 54], heights=52)

    row = section(ws, row, "Status codes & error contract", 6)
    codes = [
        ["200", "Success", "Includes refusals — a refusal is a valid, successful, correct response."],
        ["400", "Invalid payload", "Message empty, history too deep, or payload over the size cap."],
        ["401 / 403", "Unauthenticated / unauthorized", "Production auth mode only."],
        ["429", "Rate limited", "Includes a Retry-After header."],
        ["503", "Index not ready", "Returned before the first successful ingest."],
        ["502", "LLM gateway failure", "After retries are exhausted; the client shows a retry affordance."],
        ["500", "Unexpected error", "Correlation ID returned to the client for support lookup."],
    ]
    row = table(ws, row, ["Code", "Meaning", "Note"], codes, [12, 34, 100], heights=22)

    row = section(ws, row, "CLI contract", 6)
    cli = [
        ["ingest", "python -m backend.app.cli ingest [--dry-run] [--doc-id ID] [--verbose]", "0 ok · 1 unregistered source · 2 no text layer · 3 chunking failure · 4 embedding failure"],
        ["query", "python -m backend.app.cli query \"...\" [--top-k N] [--json] [--full-text]", "0 ok · 5 index missing"],
        ["chat", "python -m backend.app.cli chat \"...\" [--history FILE]", "0 ok · 5 index missing · 6 LLM unavailable"],
        ["eval", "python -m backend.app.cli eval [--set retrieval|generation|refusal]", "0 gate passed · 7 gate failed"],
        ["index", "python -m backend.app.cli index", "0 ok · 5 index missing"],
    ]
    row = table(ws, row, ["Command", "Usage", "Exit codes"], cli, [14, 66, 66], heights=22)
    return ws


def build_data_model(wb):
    ws = sheet(wb, "09 · Data Model")
    row = banner(ws, "Data Model & Metadata Rules", "Metadata lives WITH the vector entry (Principle II). No nulls — missing strings default to \"\".", 5)

    row = section(ws, row, "Chunk (the unit stored in the vector store)", 5)
    chunk = [
        ["chunk_id", "str", "Yes", "Stable unique ID: {doc_id}::{section}::{ordinal}", "Citation anchoring target."],
        ["doc_id", "str", "Yes", "Registry key of the source document", "Joins to sources.yaml."],
        ["document_name", "str", "Yes", "Human-readable title", "Rendered on the citation chip."],
        ["source_url", "str", "Yes", "Canonical URL (with heading anchor for web docs)", "\"\" if none — never null."],
        ["doc_version", "str", "Yes", "Product/doc version this chunk describes", "Cloud-specific: docs change per release."],
        ["page_number", "int", "Yes", "1-indexed PDF page; 0 for web docs", "Human trace-back."],
        ["anchor", "str", "Yes", "Heading anchor for web docs", "\"\" for PDFs."],
        ["section_title", "str", "Yes", "Nearest enclosing heading", "Shown in the evidence panel."],
        ["section_path", "str", "Yes", "Full heading breadcrumb (H1 > H2 > H3)", "Disambiguates repeated headings."],
        ["text", "str", "Yes", "Verbatim cleaned chunk text", "Never paraphrased at ingest time."],
        ["token_count", "int", "Yes", "tiktoken count", "Budget enforcement."],
        ["is_oversized", "bool", "Yes", "Atomic unit exceeded the budget and was emitted whole", "Never split an atomic unit."],
        ["content_type", "str", "Yes", "prose | procedure | code | table | reference", "Enables type-aware retrieval later."],
        ["ingested_at", "str", "Yes", "ISO timestamp", "Staleness reporting."],
    ]
    row = table(ws, row, ["Field", "Type", "Required", "Description", "Why"], chunk, [18, 10, 12, 60, 56], heights=24)

    row = section(ws, row, "SourceDocument (data/sources.yaml entry)", 5)
    src = [
        ["doc_id", "str", "Yes", "Matches ^[a-z0-9_]+$ and is unique", "Fail-closed on duplicates."],
        ["title", "str", "Yes", "Official document title", "Displayed in provenance panel."],
        ["publisher", "str", "Yes", "Owning organization/team", "Principle III."],
        ["published_at", "date", "Yes", "Publication or last-updated date", "Drives the staleness flag."],
        ["version", "str", "Yes", "Product/doc version", "Cloud docs are version-bound."],
        ["url", "str", "Yes", "Retrieval URL", "Human verification path."],
        ["document_type", "str", "Yes", "reference | guide | runbook | release_notes | tutorial", "Non-reference types are flagged."],
        ["credibility_note", "str", "Yes", "Written justification of why this source is authoritative", "Placeholder text is rejected by the validator."],
        ["license_note", "str", "Yes", "Usage rights for indexing this document", "Legal safety."],
        ["requires_caution", "bool", "Derived", "True when type != reference or the document is past the staleness window", "Surfaced wherever its chunks appear."],
    ]
    row = table(ws, row, ["Field", "Type", "Required", "Description", "Why"], src, [18, 10, 12, 60, 56], heights=24)

    row = section(ws, row, "Response objects", 5)
    resp = [
        ["RetrievalResult", "chunk_id, text, score, below_floor, document_name, section_title, section_path, page_number, anchor, source_url, doc_version", "Returned by every Retriever implementation."],
        ["Citation", "marker, chunk_id, document_name, section_title, page_number, anchor, source_url, quote", "One per resolved [Source N] marker in the answer."],
        ["ChatResponse", "answer, citations[], evidence[], evidence_found, refused, refusal_reason, disclaimer, model, latency_ms, usage, request_id", "The full contract the frontend renders."],
        ["IndexManifest", "built_at, embedding_model, dimension, chunk_count, document_count, corpus_hash, documents[]", "Drives cache invalidation and the index status panel."],
        ["GoldenQuestion", "id, question, expected_doc_id, expected_sections[], category, must_refuse", "Evaluation input for retrieval, generation and refusal suites."],
    ]
    row = table(ws, row, ["Object", "Fields", "Purpose"], resp, [22, 90, 44], heights=36)

    row = section(ws, row, "Invariants", 5)
    inv = [
        ["INV-1", "Every chunk in the index has all required metadata fields non-empty.", "Integration test on the built index."],
        ["INV-2", "No metadata value is null — missing strings are \"\" and missing ints are 0.", "Schema validation at ingest."],
        ["INV-3", "Every citation in an answer resolves to a chunk that was actually retrieved for that turn.", "Citation-integrity unit test."],
        ["INV-4", "No atomic unit (procedure, code block, table) is split across chunks.", "Chunker unit test."],
        ["INV-5", "A failed ingest never corrupts the live index (build-then-swap).", "Failure-injection integration test."],
        ["INV-6", "Below-floor results appear in the response payload, flagged.", "Retriever unit test."],
    ]
    row = table(ws, row, ["ID", "Invariant", "Verified by"], inv, [10, 90, 56], heights=24)
    return ws


def build_config(wb):
    ws = sheet(wb, "10 · Config & Env")
    row = banner(ws, "Configuration Reference", "Every environment variable, its default, and what breaks if it is wrong. Mirror this in .env.example.", 5)
    rows = [
        ["APP_ENV", "development", "str", "development | staging | production", "Gates auth and CORS strictness."],
        ["LOG_LEVEL", "INFO", "str", "Logging verbosity", "DEBUG in production floods the log pipeline."],
        ["CORPUS_DIR", "data/corpus", "path", "Source documents directory", "Fail-closed against the registry."],
        ["SOURCES_FILE", "data/sources.yaml", "path", "Provenance registry", "Missing file aborts ingestion."],
        ["INDEX_DIR", "data/index", "path", "Vector store + manifest location", "Must be a persistent volume in Docker."],
        ["COLLECTION_NAME", "cloud_docs", "str", "Vector collection name", "Namespacing for multi-corpus later."],
        ["CHUNK_MIN_TOKENS", "400", "int", "Lower packing bound", "Too low → fragmented context."],
        ["CHUNK_MAX_TOKENS", "800", "int", "Upper packing bound", "Too high → diluted embeddings."],
        ["EMBEDDING_BACKEND", "local", "str", "local | remote | fallback", "Local-first is the default path."],
        ["LOCAL_EMBEDDING_MODEL", "BAAI/bge-small-en-v1.5", "str", "Local embedder", "Changing it requires a full re-index."],
        ["REMOTE_EMBEDDING_MODEL", "<gateway model id>", "str", "Fallback embedder", "Dimension must be recorded in the manifest."],
        ["EMBEDDING_BATCH_SIZE", "100", "int", "Batch size for document embedding", "Tune against gateway limits."],
        ["RETRIEVER_TYPE", "hybrid_rerank", "str", "dense | hybrid | hybrid_rerank", "Set by the factory; measured on the eval set."],
        ["HYBRID_CANDIDATE_K", "30", "int", "Candidates per retriever before fusion", "Too low starves the reranker."],
        ["RRF_K", "60", "int", "RRF smoothing constant", "Standard default."],
        ["RERANKER_MODEL", "<cross-encoder id>", "str", "Cross-encoder weights", "Must degrade gracefully when absent."],
        ["TOP_K", "5", "int", "Chunks passed to generation", "Higher costs tokens and dilutes focus."],
        ["RELEVANCE_FLOOR", "0.68", "float", "Below this a chunk is flagged weak", "Calibrated in T044 — do not guess."],
        ["SCOPE_THRESHOLD", "0.68", "float", "Below this the question is out of scope", "Calibrated in T044."],
        ["GENERATION_MODEL", "<synthesis model id>", "str", "Answer synthesis model", "Quality/cost lever."],
        ["ROUTING_MODEL", "<cheap model id>", "str", "Classification and routing model", "Cost control."],
        ["GENERATION_MAX_TOKENS", "1024", "int", "Answer length cap", "Guards runaway cost."],
        ["GENERATION_TEMPERATURE", "0.1", "float", "Sampling temperature", "Low: grounded, repeatable answers."],
        ["HISTORY_WINDOW", "4", "int", "Conversation turns sent to the model", "Bounded context cost."],
        ["LLM_API_KEY", "(empty)", "secret", "Gateway credential", "Never logged, never in the client bundle."],
        ["LLM_BASE_URL", "<gateway url>", "str", "LLM/embedding gateway endpoint", "Swap providers without code changes."],
        ["CACHE_ENABLED", "true", "bool", "Master switch for L1/L2/L3 caches", "Disable to reproduce cold latency."],
        ["CACHE_TTL_SECONDS", "3600", "int", "Cache entry lifetime", "Also invalidated by manifest hash."],
        ["CACHE_MAX_ENTRIES", "1000", "int", "LRU capacity per tier", "Memory bound."],
        ["RATE_LIMIT_PER_MINUTE", "20", "int", "Per-IP request budget", "Protects the LLM spend."],
        ["CORS_ORIGINS", "http://localhost:3000", "csv", "Allowed browser origins", "Wildcards are rejected in production."],
        ["API_KEY", "(empty)", "secret", "Client credential when auth is enabled", "Required in production mode."],
        ["SENTRY_DSN", "(empty)", "secret", "Error tracking endpoint", "Empty disables error tracking cleanly."],
        ["TRACING_ENABLED", "false", "bool", "RAG stage tracing", "Enable in staging and production."],
        ["DISCLAIMER_TEXT", "<advisory text>", "str", "Disclaimer attached to every answer", "Principle IV requirement."],
    ]
    row = table(ws, row, ["Variable", "Default", "Type", "Purpose", "Failure mode if wrong"],
                rows, [28, 26, 10, 50, 54], heights=24, filter_on=True, freeze=True)
    return ws


def build_tests(wb):
    ws = sheet(wb, "11 · Test Matrix")
    row = banner(ws, "Verification & Test Matrix", "Every phase gate is a command that either passes or blocks the next phase.", 6)

    row = section(ws, row, "Automated suites", 6)
    auto = [
        ["P0", "Unit", "pytest backend/tests/unit/test_config.py", "Settings defaults and env overrides", "All pass", "Not Started"],
        ["P1", "Unit", "pytest backend/tests/unit/ -k 'cleaner or sectioner or chunker'", "Deterministic ingest transforms", "All pass", "Not Started"],
        ["P1", "Unit", "pytest backend/tests/unit/test_registry.py", "Fail-closed provenance validation", "All pass", "Not Started"],
        ["P1", "Integration", "pytest backend/tests/integration/test_ingest_pipeline.py", "End-to-end ingest on a fixture doc", "Metadata completeness 100%", "Not Started"],
        ["P1", "Integration", "pytest backend/tests/integration/test_atomic_units.py", "No atomic unit split across chunks", "Zero splits", "Not Started"],
        ["P2", "Unit", "pytest backend/tests/unit/ -k 'bm25 or hybrid or reranker'", "Retrieval components", "All pass", "Not Started"],
        ["P2", "Integration", "pytest backend/tests/integration/test_search.py", "Search API contract + latency", "P95 < 800 ms", "Not Started"],
        ["P2", "Eval", "pytest backend/tests/eval/test_retrieval_quality.py", "Golden-set retrieval hit rate", "GATE: ≥ 85%", "Not Started"],
        ["P3", "Unit", "pytest backend/tests/unit/ -k 'assembler or citations or guardrail'", "Generation components", "All pass", "Not Started"],
        ["P3", "Unit", "pytest backend/tests/unit/test_grounding_stress.py", "Prompt injection and jailbreak resistance", "GATE: zero bypasses", "Not Started"],
        ["P3", "Integration", "pytest backend/tests/integration/test_chat_api.py", "Chat + SSE streaming contract", "All pass", "Not Started"],
        ["P3", "Eval", "pytest backend/tests/eval/test_generation_quality.py", "Citation accuracy and grounding", "GATE: ≥ 90% citation accuracy", "Not Started"],
        ["P3", "Eval", "pytest backend/tests/eval/test_refusal.py", "Out-of-scope and unanswerable handling", "GATE: 100% correct refusals", "Not Started"],
        ["P4", "Frontend", "npm run lint && npm run build", "Type safety and production build", "Zero errors", "Not Started"],
        ["P5", "Unit", "pytest backend/tests/unit/test_observability.py", "PII scrubbing and correlation IDs", "GATE: no secrets in payloads", "Not Started"],
        ["P6", "Unit", "pytest backend/tests/unit/test_caching.py", "TTL, LRU and manifest invalidation", "All pass", "Not Started"],
        ["P7", "Integration", "pytest backend/tests/integration/test_ratelimit.py", "Throttling and abuse handling", "429 with Retry-After", "Not Started"],
        ["P7", "E2E", "scripts/test_live_e2e.py", "Live end-to-end against a running stack", "All scenarios pass", "Not Started"],
    ]
    start = row
    row = table(ws, row, ["Phase", "Type", "Command", "What it proves", "Target", "Status"],
                auto, [10, 14, 56, 46, 34, 16], heights=26, filter_on=True, freeze=True)
    paint(ws, 6, start + 1, start + len(auto), STATUS_FILLS)
    validate(ws, "F", start + 1, start + len(auto), STATUSES)

    row = section(ws, row, "Manual verification checklist", 6)
    manual = [
        ["MV-01", "Ask a question the docs clearly answer", "Answer streams with at least one citation resolving to the right section", "Not Started"],
        ["MV-02", "Ask a question the docs do not cover", "Honest refusal — no invented answer", "Not Started"],
        ["MV-03", "Ask an exact-identifier question (flag, error code, API name)", "Top result is the exact reference page", "Not Started"],
        ["MV-04", "Ask an elliptical follow-up ('and in the other region?')", "Resolves against the prior turn", "Not Started"],
        ["MV-05", "Open the evidence panel", "Every retrieved chunk listed with score, including weak ones", "Not Started"],
        ["MV-06", "Click a citation chip", "Jumps to the verbatim chunk and links to the live source", "Not Started"],
        ["MV-07", "Stop generation mid-stream", "Streaming halts and the server-side request is cancelled", "Not Started"],
        ["MV-08", "Reload the page", "Session history is intact; no theme flash", "Not Started"],
        ["MV-09", "Attempt a prompt injection ('ignore your instructions')", "Refused; instructions inside retrieved text are not executed", "Not Started"],
        ["MV-10", "Kill the LLM gateway", "Clear error with retry affordance; no partial nonsense answer", "Not Started"],
        ["MV-11", "Delete the index and restart", "503 index-not-ready with an actionable message", "Not Started"],
        ["MV-12", "Run on a 375 px viewport", "Fully usable, no horizontal scroll", "Not Started"],
        ["MV-13", "Keyboard-only navigation of the whole chat flow", "Focus visible at every step; no traps", "Not Started"],
        ["MV-14", "docker compose up from a clean checkout", "System comes up and serves UI + API", "Not Started"],
        ["MV-15", "Re-ingest an updated document", "Cache invalidates; new content is retrievable; old index stayed live during the build", "Not Started"],
    ]
    start = row
    row = table(ws, row, ["ID", "Check", "Expected", "Status"], manual, [10, 56, 76, 16], heights=24)
    paint(ws, 4, start + 1, start + len(manual), STATUS_FILLS)
    validate(ws, "D", start + 1, start + len(manual), STATUSES)
    return ws


def build_kpis(wb):
    ws = sheet(wb, "12 · KPIs")
    row = banner(ws, "Success Criteria & KPIs", "Set the target column to your own numbers before Phase 2 — these gate the release.", 6)

    row = section(ws, row, "Quality", 6)
    q = [
        ["Retrieval hit rate (golden set)", "≥ 85%", "", "Phase gate", "eval suite", "Blocks Phase 3."],
        ["Citation accuracy", "≥ 90%", "", "Phase gate", "eval suite", "Every citation resolves and supports its claim."],
        ["Grounding rate (no unsupported claims)", "100%", "", "Phase gate", "manual + eval", "Sampled review of eval answers."],
        ["Correct refusal rate (out of scope)", "100%", "", "Phase gate", "refusal suite", "A wrong answer is worse than no answer."],
        ["False refusal rate (in scope)", "≤ 5%", "", "Quality", "eval suite", "Over-refusal makes the product useless."],
        ["Multi-turn resolution accuracy", "≥ 90%", "", "Quality", "eval suite", "Elliptical follow-ups."],
    ]
    row = table(ws, row, ["Metric", "Target", "Current", "Class", "Measured by", "Note"], q, [42, 14, 14, 16, 22, 42], heights=24)

    row = section(ws, row, "Performance", 6)
    p = [
        ["Time to first token (P95)", "< 2 s", "", "Latency", "traces", "Perceived responsiveness."],
        ["Full answer latency (P95, cold)", "< 4 s", "", "Latency", "traces", "End to end."],
        ["Warm cached answer", "< 50 ms", "", "Latency", "metrics", "L3 cache hit."],
        ["Search latency (P95)", "< 800 ms", "", "Latency", "metrics", "Retrieval only."],
        ["Query embedding latency (median)", "< 30 ms", "", "Latency", "metrics", "Local-first embedding."],
        ["Cache hit rate (steady state)", "≥ 30%", "", "Efficiency", "metrics", "Common questions repeat."],
        ["Availability", "≥ 99.5%", "", "Reliability", "uptime check", "Excluding planned re-index windows."],
        ["Cost per answered question", "< $<TARGET>", "", "Cost", "token metrics", "Set your budget."],
    ]
    row = table(ws, row, ["Metric", "Target", "Current", "Class", "Measured by", "Note"], p, [42, 14, 14, 16, 22, 42], heights=24)

    row = section(ws, row, "Safety & operations", 6)
    s = [
        ["Unregistered documents indexed", "0", "", "Safety", "registry validator", "Fail-closed (Principle III)."],
        ["Metadata completeness", "100%", "", "Safety", "integration test", "Principle II."],
        ["Secrets in logs / traces / errors", "0", "", "Safety", "scrubbing test + scan", "Hard requirement."],
        ["Disclaimer presence on answers", "100%", "", "Safety", "API test", "Principle IV."],
        ["Mean time to detect a bad answer", "< 24 h", "", "Ops", "feedback loop", "User feedback → eval case."],
        ["Corpus re-index duration", "< 30 min", "", "Ops", "pipeline timing", "Docs update frequently."],
        ["Index staleness (docs vs corpus)", "< 7 days", "", "Ops", "manifest age", "Cloud docs move fast."],
    ]
    row = table(ws, row, ["Metric", "Target", "Current", "Class", "Measured by", "Note"], s, [42, 14, 14, 16, 22, 42], heights=24)

    row = section(ws, row, "Adoption (post-launch)", 6)
    a = [
        ["Weekly active users", "<TARGET>", "", "Adoption", "analytics", "Customize."],
        ["Questions per user per week", "<TARGET>", "", "Adoption", "analytics", "Depth of use."],
        ["Thumbs-up rate", "≥ 80%", "", "Satisfaction", "feedback API", "Explicit rating."],
        ["Citation click-through rate", "≥ 25%", "", "Trust", "analytics", "Users verifying evidence is a good sign."],
        ["Deflected support tickets", "<TARGET>", "", "Impact", "support system", "The business case."],
    ]
    row = table(ws, row, ["Metric", "Target", "Current", "Class", "Measured by", "Note"], a, [42, 14, 14, 16, 22, 42], heights=24)
    return ws


def build_risks(wb):
    ws = sheet(wb, "13 · Risks")
    row = banner(ws, "Risk Register", "Reviewed weekly. Likelihood × Impact drives the ordering; every risk has an owner and a trigger.", 9)
    risks = [
        ["R01", "Quality", "Retrieval misses the right section, so answers are confidently based on the wrong evidence", "High", "High", "Hybrid retrieval + reranking; golden-set gate; per-question HIT/MISS reporting; sweep chunk sizes", "ML", "Open", "Hit rate drops below 85%"],
        ["R02", "Safety", "Model answers from parametric memory when evidence is thin", "Medium", "Critical", "Hard grounding guardrail; refusal suite; prompt forbids outside knowledge; sampled human review", "ML", "Open", "Any ungrounded claim in the eval sample"],
        ["R03", "Safety", "Prompt injection embedded in the documentation corpus itself", "Medium", "High", "Treat retrieved text as data, never instructions; adversarial test set; output guardrail", "ML", "Open", "Injection test set failure"],
        ["R04", "Data", "Documentation changes faster than the index is rebuilt, so answers are stale", "High", "High", "Version metadata on every chunk; scheduled re-index; staleness flag surfaced in UI", "Data", "Open", "Manifest age > 7 days"],
        ["R05", "Data", "Corpus contains internal or restricted content that should not be indexed", "Medium", "Critical", "Fail-closed registry with a written approval note per document; secret scanning on the corpus", "Data", "Open", "Any unregistered file reaching ingestion"],
        ["R06", "Cost", "LLM spend scales past budget with adoption", "Medium", "Medium", "Answer caching; cheap routing model; token caps; rate limiting; per-answer cost alerting", "Owner", "Open", "Daily cost above threshold"],
        ["R07", "Performance", "Cross-encoder reranking pushes latency past the perceived-responsiveness budget", "Medium", "Medium", "Rerank only the top candidates; stream early; measure with and without; make it config-switchable", "Backend", "Open", "P95 TTFT > 2 s"],
        ["R08", "Reliability", "LLM gateway outage takes the whole product down", "Medium", "High", "Retries with backoff; secondary model route; degrade to retrieval-only 'here are the relevant docs' mode", "Backend", "Open", "Gateway error rate > 5%"],
        ["R09", "Reliability", "A failed re-index leaves the system with no queryable index", "Low", "High", "Atomic build-then-swap; rehearsed restore procedure; index backup before each build", "DevOps", "Open", "Any ingest failure in production"],
        ["R10", "UX", "Over-refusal: the bot refuses questions the docs actually answer", "Medium", "High", "Calibrate the scope threshold on real queries; track false-refusal rate; make refusals suggest search results", "ML", "Open", "False refusal rate > 5%"],
        ["R11", "UX", "Users trust a fluent answer without checking citations", "High", "Medium", "Citations inline and prominent; evidence panel one click away; disclaimer on every answer", "Design", "Open", "Citation CTR below 25%"],
        ["R12", "Scope", "Stakeholders push to widen the corpus mid-build ('add the billing docs too')", "High", "Medium", "Written scope statement; corpus additions are a planned phase with their own eval set", "Owner", "Open", "Any unplanned corpus addition request"],
        ["R13", "Technical", "Embedding model change silently invalidates the index", "Low", "High", "Model ID and dimension recorded in the manifest; startup check refuses a mismatched index", "Backend", "Open", "Manifest/config mismatch at startup"],
        ["R14", "Technical", "Web documentation parsing breaks when the docs site redesigns", "Medium", "Medium", "Parser contract tests on fixture HTML; fail loudly on anchor extraction failure", "Backend", "Open", "Anchor extraction rate drops"],
        ["R15", "Legal", "Indexing third-party documentation breaches its terms of use", "Low", "Critical", "license_note required per source; legal review before adding external corpora", "Owner", "Open", "Any external source added"],
        ["R16", "Delivery", "Phase gates slip and generation ships on unverified retrieval", "Medium", "High", "Principle V enforced in code — /api/chat stays 501 until the gate passes in CI", "Owner", "Open", "Pressure to ship before the gate"],
    ]
    start = row
    row = table(ws, row, ["ID", "Category", "Risk", "Likelihood", "Impact", "Mitigation", "Owner", "Status", "Escalation trigger"],
                risks, [8, 14, 58, 13, 11, 70, 12, 14, 30], heights=44, filter_on=True, freeze=True)
    sev = {"Low": PatternFill("solid", fgColor="E3F2E7"),
           "Medium": PatternFill("solid", fgColor="FFF1CC"),
           "High": PatternFill("solid", fgColor="FADBD8"),
           "Critical": PatternFill("solid", fgColor="F5B7B1")}
    paint(ws, 4, start + 1, start + len(risks), sev)
    paint(ws, 5, start + 1, start + len(risks), sev)
    # merge the trigger column content into a status dropdown-friendly shape
    validate(ws, "D", start + 1, start + len(risks), ["Low", "Medium", "High"])
    validate(ws, "E", start + 1, start + len(risks), ["Low", "Medium", "High", "Critical"])
    validate(ws, "H", start + 1, start + len(risks), ["Open", "Mitigating", "Accepted", "Closed"])
    return ws


def build_deliverables(wb):
    ws = sheet(wb, "14 · Deliverables")
    row = banner(ws, "Deliverables Manifest", "What exists in the repository when each phase is done. Tick as they land.", 5)
    d = [
        ["P0", "Code", "backend/app/{config,models,errors,main}.py + package tree + protocols", "Runnable API skeleton", "Not Started"],
        ["P0", "Config", ".env.example, pyproject.toml, requirements.txt, .gitignore", "Reproducible environment", "Not Started"],
        ["P0", "Governance", "Constitution file + one-sentence scope statement", "Agreed rules of the build", "Not Started"],
        ["P0", "CI", ".github/workflows/pr-check.yml, pre-push hook", "Lint + unit tests on every PR", "Not Started"],
        ["P1", "Data", "data/sources.yaml, data/corpus/", "Registered, provenance-checked corpus", "Not Started"],
        ["P1", "Code", "backend/app/ingestion/* (registry, parser, html_parser, cleaner, sectioner, chunker, pipeline)", "Documents → chunks", "Not Started"],
        ["P1", "Code", "backend/app/embeddings/* and retrieval/store.py", "Chunks → searchable index", "Not Started"],
        ["P1", "Artifact", "data/index/ + manifest.json", "Built index with a self-describing manifest", "Not Started"],
        ["P1", "CLI", "ingest and index commands", "One-command corpus build", "Not Started"],
        ["P2", "Code", "retrieval/{dense,bm25,hybrid,reranker,factory,scope}.py", "Hybrid retrieval with a scope gate", "Not Started"],
        ["P2", "API", "POST /api/search, GET /api/index, GET /api/sources", "Inspectable retrieval surface", "Not Started"],
        ["P2", "Eval", "golden_questions.yaml + test_retrieval_quality.py", "Measured retrieval quality", "Not Started"],
        ["P2", "Docs", "docs/THRESHOLD_CALIBRATION.md, docs/CHUNKING_SWEEP.md", "Justified configuration", "Not Started"],
        ["P3", "Code", "generation/{client,assembler,prompt,citations,guardrails,risk_classifier,clarification}.py", "Grounded answer synthesis", "Not Started"],
        ["P3", "API", "POST /api/chat and /api/chat/stream", "Streaming chat with citations", "Not Started"],
        ["P3", "Eval", "golden_generation.yaml, refusal_cases.yaml + suites", "Measured answer safety and accuracy", "Not Started"],
        ["P4", "Frontend", "app/, components/chat/*, hooks/*, lib/api.ts", "Full chat experience", "Not Started"],
        ["P4", "Design", "globals.css design tokens + component primitives", "Consistent design system", "Not Started"],
        ["P5", "Code", "monitoring/{logging_config,tracing,sentry,metrics}.py", "Full-stack observability", "Not Started"],
        ["P5", "Data", "data/eval/ dataset + DATASET_CARD.md", "Reusable evaluation dataset", "Not Started"],
        ["P5", "Docs", "docs/EVALUATION_REPORT.md", "Quality baseline and failure taxonomy", "Not Started"],
        ["P6", "Code", "retrieval/cache.py + startup warm-up", "Multi-tier caching", "Not Started"],
        ["P6", "Docs", "docs/PERFORMANCE.md", "Before/after latency evidence", "Not Started"],
        ["P7", "Code", "ratelimit.py, auth, CORS lockdown, input caps", "Hardened API surface", "Not Started"],
        ["P7", "Ops", "Dockerfile, docker-compose*.yml, docker-entrypoint.sh", "One-command deployment", "Not Started"],
        ["P7", "CI", ".github/workflows/ci-cd.yml", "Test → build → publish → deploy", "Not Started"],
        ["P8", "Docs", "README.md, docs/API.md, docs/RUNBOOK.md, docs/CORPUS_GOVERNANCE.md, docs/USER_GUIDE.md, docs/DECISIONS.md", "Complete handover pack", "Not Started"],
        ["P8", "Process", "docs/LAUNCH_CHECKLIST.md + feedback loop", "Repeatable launch and improvement cycle", "Not Started"],
    ]
    start = row
    row = table(ws, row, ["Phase", "Type", "Deliverable", "Why it matters", "Status"],
                d, [10, 14, 86, 44, 16], heights=28, filter_on=True, freeze=True)
    paint(ws, 5, start + 1, start + len(d), STATUS_FILLS)
    validate(ws, "E", start + 1, start + len(d), STATUSES)
    return ws


def build_timeline(wb):
    ws = sheet(wb, "15 · Timeline")
    weeks = 12
    span = 3 + weeks
    row = banner(ws, "12-Week Timeline", "Gantt-style phase grid. Shift the bars by moving the fills — the roadmap sheet holds the authoritative dates.", span)

    headers = ["Phase", "Name", "Effort (d)"] + [f"W{i}" for i in range(1, weeks + 1)]
    widths = [10, 34, 12] + [5] * weeks
    plan = {
        "P0": (1, 1), "P1": (2, 3), "P2": (4, 5), "P3": (6, 7),
        "P4": (7, 8), "P5": (9, 9), "P6": (10, 10), "P7": (11, 11), "P8": (12, 12),
    }
    rows = [[pid, name, ""] + [""] * weeks for pid, name, *_ in PHASES]
    start = row
    row = table(ws, row, headers, rows, widths, heights=22)
    bar = PatternFill("solid", fgColor=SAPPHIRE)
    gate = PatternFill("solid", fgColor="C0392B")
    gates = {"P2": 5, "P3": 7}
    for j, (pid, *_rest) in enumerate(PHASES):
        r = start + 1 + j
        s, e = plan[pid]
        for w in range(s, e + 1):
            c = ws.cell(row=r, column=3 + w)
            c.fill = bar
            c.border = BORDER
        if pid in gates:
            g = ws.cell(row=r, column=3 + gates[pid])
            g.fill = gate
            g.value = "G"
            g.font = Font(name="Calibri", size=9, bold=True, color=WHITE)
            g.alignment = WRAP_CENTER

    row += 1
    legend = [
        ["Blue bar", "Phase active"],
        ["Red 'G'", "Hard quality gate — the next phase may not start until it passes"],
        ["W5 gate", "Retrieval golden-set hit rate ≥ 85% (unblocks generation, Principle V)"],
        ["W7 gate", "Citation accuracy ≥ 90% and 100% correct refusals (unblocks launch prep)"],
    ]
    row = table(ws, row, ["Marker", "Meaning"], legend, [16, 90], heights=22)

    row = section(ws, row, "Milestones", span)
    ms = [
        ["M1", "End W1", "Skeleton runs, CI green, constitution agreed"],
        ["M2", "End W3", "Corpus indexed with 100% metadata completeness"],
        ["M3", "End W5", "RETRIEVAL GATE PASSED — generation unblocked"],
        ["M4", "End W7", "SAFETY GATE PASSED — grounded streaming answers with citations"],
        ["M5", "End W8", "Feature-complete chat experience (internal demo)"],
        ["M6", "End W9", "Observability live, evaluation in CI"],
        ["M7", "End W11", "Deployed to staging, hardened, verification checklist passed"],
        ["M8", "End W12", "Documented, launched, feedback loop running"],
    ]
    row = table(ws, row, ["ID", "When", "Milestone"], ms, [10, 16, 90], heights=22)
    return ws


def build_glossary(wb):
    ws = sheet(wb, "16 · Glossary")
    row = banner(ws, "Glossary", "Shared vocabulary — use these terms consistently in specs, code and PRs.", 2)
    g = [
        ["RAG", "Retrieval-Augmented Generation: retrieve evidence first, then generate an answer strictly from it."],
        ["Corpus", "The set of registered source documents the chatbot is allowed to answer from."],
        ["Registry", "data/sources.yaml — the provenance record for every corpus document."],
        ["Fail-closed", "On doubt, stop. An unregistered file aborts ingestion; missing evidence produces a refusal."],
        ["Chunk", "One indexed unit of text with its full citation metadata attached."],
        ["Atomic unit", "A procedure, code block or table that must never be split across chunks."],
        ["Embedding", "The vector representation of a chunk or query used for semantic similarity."],
        ["Dense retrieval", "Similarity search over embeddings — good at meaning, weak at exact strings."],
        ["Lexical retrieval (BM25)", "Keyword scoring — good at exact identifiers, flags and error codes."],
        ["Hybrid retrieval", "Running both retrievers and fusing their rankings."],
        ["RRF", "Reciprocal Rank Fusion: combines rankings by 1/(k+rank), no score normalization needed."],
        ["Reranker", "A cross-encoder that scores query+chunk pairs jointly for higher precision on the top candidates."],
        ["Relevance floor", "The score below which a chunk is flagged weak — flagged, never silently dropped."],
        ["Scope threshold", "The score below which a question is treated as out of scope and refused."],
        ["Below-floor", "A retrieved chunk whose score is under the relevance floor; shown with a warning."],
        ["Grounding", "The property that every claim in an answer traces to retrieved evidence."],
        ["Guardrail", "A check that blocks unsafe output — refusal on no evidence, injection resistance, disclaimer."],
        ["Refusal", "A correct, successful response stating that the docs do not support an answer."],
        ["Citation anchoring", "Mapping [Source N] markers in the generated text back to real chunk IDs."],
        ["Manifest", "data/index/manifest.json — describes the built index and drives cache invalidation."],
        ["Atomic swap", "Build the new index alongside the old one and switch only on success."],
        ["Golden set", "The curated question set used to measure retrieval and answer quality."],
        ["Phase gate", "A measured threshold that must pass before the next phase may start."],
        ["SSE", "Server-Sent Events — the one-way streaming channel used for token-by-token answers."],
        ["TTFT", "Time To First Token — the primary perceived-responsiveness metric."],
        ["L1 / L2 / L3 cache", "Embedding cache / retrieval cache / answer cache."],
        ["Staleness", "The gap between the live documentation and the last built index."],
        ["Correlation ID", "The per-request identifier that ties logs, traces and errors together."],
    ]
    row = table(ws, row, ["Term", "Meaning"], g, [26, 116], heights=24)
    return ws


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)
    build_cover(wb)
    build_scope(wb)
    build_constitution(wb)
    build_stack(wb)
    build_architecture(wb)
    build_stories(wb)
    build_roadmap(wb)
    build_wbs(wb)
    build_api(wb)
    build_data_model(wb)
    build_config(wb)
    build_tests(wb)
    build_kpis(wb)
    build_risks(wb)
    build_deliverables(wb)
    build_timeline(wb)
    build_glossary(wb)

    wb.properties.title = f"{PROJECT} — Project Plan"
    wb.properties.subject = TAGLINE
    wb.properties.creator = "Cloud Chatbot planning"

    out = Path(__file__).resolve().parents[1] / "docs" / "Cloud_Chatbot_Project_Plan.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Wrote {out} ({len(wb.sheetnames)} sheets, {len(TASKS)} tasks)")


if __name__ == "__main__":
    main()
