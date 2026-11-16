# AI-compute layered-diagram example

This is the actual third-round A drawing implementation, made relocatable. It recreates six fixed-topic slides with editable native components and source-backed notes. It is not a generic title-to-PPT program, automatic layout solver, or a slide-design quality model.

## Contents

- `build_example.py`: native drawing helpers and six authored compositions; Python 3.10+.
- `research.md`, `research.facts.json`: bundled source evidence and scope.
- `brief.json`, `diagram-map.json`: the brief and semantic page mapping retained from the example.
- `verify_example.py`: narrow package/text/notes checks and optional PNG renders from PDF.
- `requirements.txt`: Python dependencies. The inspected run used python-pptx 1.0.2 and PyMuPDF 2.1.

## Run

From this example directory, with dependencies installed in a virtual environment:

```bash
python -m pip install -r requirements.txt
python build_example.py --output-dir /tmp/ai-compute-example
soffice --headless --convert-to pdf --outdir /tmp/ai-compute-example/A /tmp/ai-compute-example/A/A_分层机制图.pptx
python verify_example.py --pptx /tmp/ai-compute-example/A/A_分层机制图.pptx --pdf /tmp/ai-compute-example/A/A_分层机制图.pdf --report /tmp/ai-compute-example/validation.json --preview-dir /tmp/ai-compute-example/previews
```

LibreOffice (`soffice`) and Noto Sans CJK SC must be installed separately. A missing renderer or font must be reported; do not substitute a source-image preview and call it a PPTX render. The task-local extracted LibreOffice path used in the cloud session is not portable and is deliberately not hard-coded here.

The output contains `A/A_分层机制图.pptx`, `A/notes.json`, and `A/fact-map.json`. Repeated runs in the same output directory overwrite those example outputs; use a new directory to retain a round. Supplying `--facts-file` changes the evidence file only; it does not adapt hard-coded slide content to a new topic. Adapt the brief, map, evidence and drawing code together for a new presentation.

The verifier must exit successfully and every page must then receive semantic and visual review. A successful verifier deliberately keeps `complete=false`. Editable components are not automatic semantic groups, and connector segments do not auto-reroute when a user moves objects.

Read [the workflow](../../architectures/drawing-workflow.md) and [the current diagram contract](../../references/diagram-rich.zh-CN.md) before adapting this example.
