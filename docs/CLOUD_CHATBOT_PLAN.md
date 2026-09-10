# Cloud Chatbot — Project Plan

The full plan lives in **[`docs/Cloud_Chatbot_Project_Plan.xlsx`](Cloud_Chatbot_Project_Plan.xlsx)**.

It ports the Eva-AI architecture — evidence-grounded RAG, fail-closed provenance, staged
delivery — onto a general cloud-documentation chatbot, and leaves every domain-specific
value as a marked placeholder for you to customize.

## What is in the workbook

| Sheet | Contents |
|---|---|
| `00 · Cover` | Project charter, sheet map, what is customizable vs. fixed |
| `01 · Scope` | In scope (SC-01…12), non-goals (OS-01…07), assumptions and constraints |
| `02 · Constitution` | The six non-negotiable principles and how each is enforced |
| `03 · Tech Stack` | Backend, frontend and ops choices with the swap seam for each |
| `04 · Architecture` | 12-stage request path and the full component inventory with target file paths |
| `05 · User Stories` | US1–US12 with priority and acceptance criteria |
| `06 · Roadmap` | Phases P0–P8 with goals, exit gates and the dependency chain |
| `07 · WBS Tasks` | 109 tasks: owner, estimate, dependencies, priority, status, acceptance check |
| `08 · API Contracts` | Every endpoint, payload, failure behaviour, status codes and the CLI contract |
| `09 · Data Model` | Chunk / SourceDocument / response schemas plus six index invariants |
| `10 · Config & Env` | 35 environment variables with defaults and failure modes |
| `11 · Test Matrix` | Automated suites per phase plus a 15-item manual verification checklist |
| `12 · KPIs` | Quality, performance, safety and adoption targets |
| `13 · Risks` | 16-risk register with likelihood, impact, mitigation, owner and escalation trigger |
| `14 · Deliverables` | File-by-file manifest of what each phase produces |
| `15 · Timeline` | 12-week Gantt-style grid with the two hard quality gates marked |
| `16 · Glossary` | Shared vocabulary |

## Architecture inherited from Eva-AI

1. **Evidence-grounded answers only** — no path to an answer that bypasses retrieval.
2. **Citation metadata is structural** — stored with the vector entry, never in a sidecar.
3. **Fail-closed provenance** — unregistered documents abort ingestion.
4. **Narrow scope discipline** — out-of-scope questions are refused, not guessed at.
5. **Staged delivery** — `/api/chat` stays `501` until the retrieval golden gate passes.
6. **Human verification over automated confidence** — weak evidence is flagged, never hidden.

The two hard gates: **retrieval hit rate ≥ 85%** (end of week 5, unblocks generation) and
**citation accuracy ≥ 90% with 100% correct refusals** (end of week 7, unblocks launch prep).

## Customizing it

Everything marked `Customize` on the Cover sheet, plus every `<PLACEHOLDER>` value, is
yours to fill in: the scope sentence, the corpus, the user roles, the model IDs, the KPI
targets and the week numbers. Everything marked `Fixed` is the safety architecture —
changing it changes what the system is allowed to do.

## Regenerating

The workbook is generated, not hand-edited, so structural changes belong in the script:

```bash
pip install openpyxl
python scripts/generate_cloud_chatbot_plan.py
```

Day-to-day tracking (status, owners, percentages) is meant to be edited directly in the
`.xlsx`; re-running the script overwrites those edits.
