---
name: ppt-validation
description: Assess a presentation against source evidence and reference design, inspect rendered pages and direct corrections before delivery.
---
# Presentation validation

The same task agent reviews the original obligations and evidence independently of its production rationale. Read the original request, brief, ledger, map, exported PPTX and current render outputs. Check input/output identities before reusing results.

For every requirement retain: ID/source -> claim/data/page/element -> expected result and basis -> check executed -> observation/artifact -> disposition.

## Four checks
1. Facts: verify cited passages support wording, forecast labels and qualifications. Recalculate displayed derived values from raw data; compare every quantitative chart with normalized values, units and axis settings. Inspect contradictory evidence. A citation footer alone is insufficient.
2. Coverage and logic: every required source item appears or has an authorized omission; arrows have supported meaning; titles do not overstate correlation, forecasts or uncertain findings. Compare intended reading order and hierarchy with the reference's transferable characteristics.
3. Visual: inspect every rendered page both at overview scale and the declared reading scale. Check CJK glyphs, substitutions, clipping, overlap, label collisions, citation readability, contrast, alignment, density and chart legends. Keep screenshots and concrete findings. Do not judge appearance from XML alone. Do not assert PowerPoint fidelity from an alternate renderer without stating the limit.
4. Package: verify openable PPTX, page count, expected text, editable object types, relationships, embedded assets and requested notes/links. A text extraction test cannot establish visual readability.

## Evidence modes
Source-guided verification checks against inspected source content and actual datasets. Reference-guided design review compares hierarchy, grouping, density and readability against the user's visual intent; it is not pixel matching. A rubric is a reasoned review aid, never a universal aesthetic score. Label synthetic diagnostic content separately from factual output.

## Decisions
supported: criterion met under recorded conditions.
contradicted: capture failure, identify cause and revise the owning stage.
unchecked: execute the available check; do not relabel as passed.
blocked: identify external missing evidence/tool/decision and finish independent checks.

Use investigate/revise/complete/blocked in the round report. Complete requires all mandatory checks supported, no unresolved critical findings and evidence matching current artifacts. Partial artifacts may be delivered as drafts with limits. Rendering unavailable means visual checks blocked; source/render discrepancy means revise. Changed page/data invalidates dependent checks. Preserve prior rounds. Do not set an arbitrary retry cap or silently waive requirements to finish.

## Diagram-rich implementation

Paths in this section are relative to the package root. Apply the semantic checks in references/diagram-rich.zh-CN.md: input provenance, stage order, shared versus request-local state, synchronization, update destinations and outputs. Inspect overview and reading-size actual PPTX renders. The example verifier checks package/text/notes only and deliberately leaves semantic and visual review unchecked; do not treat it as a completion gate on its own.

## Physical-machine composition
Apply the checklist in references/physical-machine.zh-CN.md. Verify that the left assembly reads as one machine, the right tiers match it, component counts and arithmetic agree, leaders land on correct objects, and unknown geometry/counts remain visibly qualified. Check generated chip counts and passive/active board distinctions explicitly. Inspect final PPTX renders and static-media packaging; report file sizes. A plausible illustration is not evidence of the target machine’s real construction.

## Output root
Keep validation reports and preview artifacts under the same `output/<project-name>/` as the deliverables. Verify artifact links after relocation; do not relabel previous manual checks as newly executed or overwrite them with a package-only rerun.

## Optional mechanism-illustrated style
Read `references/styles/mechanism-illustrated.zh-CN.md` relative to the repository root. When selected, apply the style checklist: conceptual relationships, problem solved, same-example continuity, data and feedback destinations, evidence boundaries, final takeaway, actual renders, editability and static-media file size. Compare the actual complete composition against visual-v3; reject sparse device-cutout/card substitutes. Inspect all raster labels and arrows visually against the approved semantic map; text extraction cannot verify embedded image text. Device recognition alone is not explanatory success.
