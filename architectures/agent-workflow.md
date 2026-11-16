# Architecture

## Stage ownership
| Stage | Owns | Durable output |
|---|---|---|
| ppt-task-design | Audience, communication objective, required claims, uncertainty, output contract | brief.json |
| ppt-mapping | Source inventory, claim evidence, datasets, relationships, page assignments and visual encodings | sources.json, claims.json, data/, deck-map.json |
| ppt-dev | Canonical page source, charts, editable package, render outputs | svg_output/, exports/, renders/ |
| ppt-validation | Independent checks, contradictions, decisions, revision routing | validation/round-N/report.json |

Requirement R -> claim C -> source S/data D -> page P -> element E -> check V is the traceability chain. A requirement may be structural rather than factual, and a page may intentionally make no factual claim. Mark this explicitly; do not invent claims to fill a schema.

The evidence ledger is authoritative for values and source identities. The deck map references IDs and owns presentation decisions. Page sources own geometry, not independently maintained copies of factual truth. Rendered PPTX is the delivery artifact whose appearance must be checked.

## Gates
G0: brief contains deliverables, required coverage, source policy, audience and unresolved critical choices.
G1: every publishable factual claim has inspected evidence or is visibly attributed as an unverified supplied assertion; important contradictions are resolved or disclosed. Quantitative figures have usable data and units.
G2: every required item maps to a page, element and planned check, or an explicit blocker. Page titles reflect the strength of evidence. Determine canvas and reading context before typography.
G3: page sources pass the selected backend's checks, and export succeeds.
G4: inspect each rendered output page, verify numerical/semantic integrity and editability; resolve all blocking findings. Missing rendering leaves visual quality unverified and prevents full-completion status.

Do not claim automatic enforcement: SKILL instructions are behavioral constraints. Backend checks enforce their own narrow contracts. A future validate.py can enforce schema, IDs, hashes and pending work, but cannot certify truth or aesthetics. No general semantic or aesthetic validator is implemented in this design package. The topic-specific example verifier performs only package/text/notes checks; manual semantic and visual gates remain required.

## Revisions
Evidence/intent errors -> task-design or mapping. Incorrect visual encoding or omitted content -> mapping. Overlap, clipping, export defects -> dev. Revalidate affected pages and all downstream calculations; reuse unaffected checks only with matching identities. Preserve failures and contrary evidence. Stop only at complete or an externally blocked state after independent checks are exhausted.

## Implemented example and drawing policy
Read [drawing-workflow.md](drawing-workflow.md). The current diagram policy is owned by [diagram-rich.zh-CN.md](../references/diagram-rich.zh-CN.md); the four skills reference it at their owning stages. Canonical sources may be Python drawing code for standalone python-pptx or SVG for the separate ppt-master backend. The current AI-compute example has manually authored coordinates, not an automatic layout solver.
