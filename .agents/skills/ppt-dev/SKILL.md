---
name: ppt-dev
description: Produce editable research slides from an evidence-backed page map using the installed PPT backend and preserve traceability.
---
# Presentation production

Read the brief, evidence ledger and deck map. Resolve missing semantics through task-design and mapping before dependent production. Load .agents/tools/ppt-backend.md; inspect the actual installed backend, version and supported features.

1. Create an isolated project output directory, not a Git worktree. Keep original inputs unchanged. Prepare all referenced resources and data transforms before the dependent page.
2. Implement the mapped content and encodings. Keep native editable text, shapes and supported chart/table objects. A necessary raster figure must have source attribution and disclosed editability limits; do not turn entire slides into images and call them editable.
3. Produce one representative dense page to test typography, information hierarchy, citations and export behavior. Correct systemic layout issues before applying the pattern elsewhere. Do not require a user approval stop unless the selected backend explicitly requires one.
4. Preserve IDs or an explicit element correspondence map when the exporter rewrites identifiers. Keep data values linked to source datasets. Do not fix a chart by manually changing plotted values.
5. Run backend checks and export. Render the actual exported PPTX with a capable renderer when available, recording renderer/version and font substitutions. SVG previews are supplemental; they do not establish final PowerPoint rendering.
6. Save command status, outputs, identities and render paths. Continue ppt-validation without another approval. Do not treat a zero exit code as proof of content correctness or visual quality.

## Diagram-rich implementation

Paths in this section are relative to the package root. Follow architectures/drawing-workflow.md and the selected backend adapter. For standalone python-pptx, use native shapes/text/connectors; document manual coordinates, units, font and z-order. Route edges before labels, preserve state ownership and use explicit junctions. Inspect a representative dense page before completing the deck. The AI-compute sample demonstrates the pattern but does not select layouts for other topics.

## Physical-machine composition
Apply references/physical-machine.zh-CN.md. Produce a coherent left whole-machine assembly and corresponding right detail tiers. Use source-backed static technical illustrations or original hardware imagery as appropriate; native mechanism shapes are not the only allowed implementation. Keep quantity labels/source notes editable where practical and disclose raster limits. Do not invent chassis geometry without authorized, visibly labeled illustration. Avoid dynamic assets unless requested.

## Output root
The isolated project directory is repository-root `output/<project-name>/`, not a sibling workspace folder or legacy `outputs/`. Keep sources, exports and renders together; resolve script paths from the project directory. Exclude local renderer profiles/caches from Git. User-specified destinations override this default.

## Selectable mechanism styles
Read `references/styles/README.md` relative to the repository root, then the selected style contract. Both mechanism-illustrated and mechanism-editable are independent options, not stages. For mechanism-illustrated follow its preserved integrated-static-illustration contract: complete static artwork with disclosed raster-main-image limits. For mechanism-editable use native PowerPoint text, geometry, groups and paths; no flattened main diagram. Preserve explanatory detail and obey the chosen backend. Do not silently substitute one branch for the other.
