# Backend selection

Choose and record one backend before production. A standalone python-pptx project does not use ppt-master routing, SVG checks or its exporter. A project explicitly using ppt-master must follow that engine’s own workflow; standalone code is not a bypass for its gates.

## Standalone python-pptx (implemented A example)

`examples/ai-compute/build_example.py` draws native PowerPoint text, shapes and lines directly into DrawingML using python-pptx. Page coordinates and diagram composition are manually authored. It records notes and source fact IDs. Run with an explicit output directory; bundled facts are the default input. See the example README for rendering and verification.

The example verifier checks only page count, aspect, ZIP integrity, notes, editable object types and normalized PDF text coverage. It cannot prove visual quality or source truth. Render the actual PPTX and review it manually.

## Separate ppt-master integration

This package is a design proposal, not an installed replacement for ppt-master. Its four stages express product obligations; the installed PPT engine owns supported SVG grammar, file formats and export commands.

For a separately installed ppt-master engine, resolve its actual repository root and `skills/ppt-master` directory from the current environment; do not assume a workspace path. Read that repository's AGENTS.md and SKILL.md, run its attribution guard, and follow its route authority. Never bypass source integrity, canonical SVG validation or export gates.

Integration must resolve ownership BEFORE installing this package into the existing repository:
- Standard Generate already owns design_spec/spec_lock and confirmation. Map the brief and deck map into those authoritative artifacts, or retain only fields the engine does not own; do not maintain conflicting parallel plans.
- Quick Generate explicitly forbids durable strategy/spec/lock artifacts and requires explicit quick/direct-export intent. This persistent evidence/mapping design must not be layered onto unmodified Quick. Do not force --quick-generate merely to avoid interaction.
- This proposal's automatic stage handoff does not silently remove the standard engine's explicit confirmation contract. A future integration must deliberately revise that contract if the user wants unattended operation.
- Edit Native PPTX is a separate backend route for preserving existing slides; new deck creation must not masquerade as an edit to evade generation rules.

Known backend capabilities checked in this session include source_to_md.py, svg_quality_checker.py, svg_to_pptx.py and pptx_to_svg.py. Exact flags and prerequisites come from current backend documentation. The previous two-page smoke test demonstrates only simple text/shape export and reimport, not research-deck quality.

A future tools/validate.py should validate referential IDs, required fields, data/slide hashes, required check dispositions, and pending work. It must fail a complete claim on stale artifacts or unchecked mandatory obligations. It must not claim to assess evidence truth or aesthetics. This file specifies that tool; no such custom command is implemented here.
