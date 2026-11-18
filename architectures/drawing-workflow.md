# Diagram-rich drawing workflow

The A presentation uses standalone Python drawing code and python-pptx. It does not use image generation, Mermaid, Graphviz, a browser screenshot, or ppt-master's SVG exporter. Its composition and coordinates are authored by the task agent. The same agent carries out four stages; the stages are not four autonomous agents.

## 1. Intent and evidence

Read the user's title, supplied content and visual reference. Retain authoritative source excerpts, scope and provenance. Use the reference for hierarchy, grouping, internal detail and reading order, not as numerical evidence. Save the brief and explanatory questions. The current policy is [diagram-rich.zh-CN.md](../references/diagram-rich.zh-CN.md).

## 2. Semantic diagram before geometry

Map inputs, internal subprocesses, state owners, dependencies, outputs and applicable feedback/exception paths. Mark a relationship as data flow, read/write, synchronization, update, containment or analytical association. Preserve fact IDs and qualifications. Do not invent loops to make the diagram look complete.

Example training decomposition: separate batches → model replicas → forward/loss/backward → local gradients → collective synchronization → update each replica. Example inference decomposition: request identities → scheduling → mixed prefill/decode execution ↔ request-local KV state → per-request output.

Select a page composition from these relationships: lanes for distinct owners; nesting for state ownership; adjacent comparable constructions for contrasting mechanisms; explicit feedback for real updates. Enrichment means explaining a previously hidden part of the mechanism, not adding shapes or hardware specifications.

## 3. Native drawing implementation

The reference example uses a 1280 × 720 design coordinate space. Coordinates are converted to inches with `px / 96`; font sizes convert to points with `px × 0.75`. The PPTX is 16:9.

| Helper | Native construction | Purpose |
| --- | --- | --- |
| `rect`, `txt` | PowerPoint shapes and editable text boxes | Fields, owners and labels |
| `ln`, `wire` | Native connector segments; DrawingML line-end metadata | Directed edges and orthogonal routes |
| `node`, `band` | Shape/text compositions | Named process or ownership region |
| `chip` | Outline, pins and label | Device boundary; internal content must explain its role |
| `matrix`, `tokens` | Repeated native rectangles and labels | Qualitative state or token identity; not measured geometry |
| `phase`, `page` | Native title, markers and frame | Reading sequence and page hierarchy |

These are drawing helpers, not an automatic diagram layout library. `wire` uses authored points; it does not route around obstacles or attach endpoints automatically when a user moves a node. Text uses authored boxes with wrapping disabled; this requires checking actual renders for overflow. Diagrams consist of editable components, but the Python helpers do not automatically create movable semantic groups in PowerPoint.

Use the paint order intentionally: fields → connections → nodes/state → labels. Where a feedback path is drawn later, inspect whether it crosses labels or needs rerouting. Preserve a consistent meaning for color. Equal block size and repeated counts are qualitative unless supported by real data. Detail beyond the current backend's capabilities requires a deliberate new implementation, not a false capability claim.

## 4. Export and rendering

python-pptx writes the native PPTX package directly, including notes and fact references. Render that actual file using LibreOffice or Microsoft PowerPoint. In the inspected cloud run, LibreOffice 25.2 was extracted into a task-local runtime; the A deck uses Noto Sans CJK SC, while B needed fallback for Microsoft YaHei. The local runtime is environment setup, not a bundled or portable dependency of this design package.

PyMuPDF opens the rendered PDF, extracts text and creates PNG previews. Pillow was used in the original delivery for contact sheets; it is optional for the portable example's single-page PNG exports.

## 5. Checks and revision ownership

- Package checks: six slides for this example, 16:9, valid ZIP, notes, expected native object types.
- Text checks: normalized PPTX text appears in the corresponding rendered PDF page. This does not detect every wrap, overlap or bad connection.
- Semantic review: stages, source scope, state identity and arrow destinations are correct.
- Visual review: overview hierarchy and reading-size labels, intersections, clipping and font substitution.

A semantic error returns to mapping. A position or wrapping defect returns to drawing. Re-export and rerender changed pages; do not reuse verification from an old package hash.

The portable verifier leaves manual checks `unchecked`. It is intentionally not a universal quality validator or a replacement for the four-stage completion contract. See [the executable example](../examples/ai-compute/README.md).

## 6. Physical-machine composition branch

For real machines, cabinets and hardware composition, apply [physical-machine.zh-CN.md](../references/physical-machine.zh-CN.md). The native drawing implementation above describes the AI-compute example, not a restriction on all future diagrams. This branch permits source-backed static hardware imagery, model renders or image generation/editing, embedded through the selected PPT backend with disclosed editability.

Research → component/quantity ledger → one-page whole-machine-left / layered-details-right mapping → static illustration → actual PPTX rendering → factual, assembly, visual and package checks. A whole-machine view must remain a coherent assembly, not scattered components. Do not split overview and details across pages as a substitute for the requested same-page correspondence. Keep quantities, placement evidence, illustrative permissions and reference-platform boundaries explicit.

## 7. Project output location
Create every new presentation project in repository-root `output/<project-name>/`. Keep its brief, sources, claims, page map, drawing code, `research/`, `exports/`, `renders/` and `validation/` together. Retain revisions in separate named subdirectories. Scripts resolve paths from their own location; do not default to a sibling workspace directory or the legacy `outputs/` spelling. Keep local renderer profiles/caches out of version control. See [output convention](../output/README.md).

## 8. Selectable mechanism-illustrated style
For static illustrated devices with explanatory data/feedback paths, apply [mechanism-illustrated](../references/styles/mechanism-illustrated.zh-CN.md). Sequence: concept relationships → one running example → focused mechanism changes → takeaway comparison, as applicable. Default to a complete static technical illustration per page with integrated cutaways, data blocks, queues and feedback, matching visual-v3. Raster main images are an explicit style choice and are not independently editable; native footers and notes are allowed. Do not replace this with device cutouts plus native boxes/cards. This extends drawing choices without changing backend or evidence requirements.
