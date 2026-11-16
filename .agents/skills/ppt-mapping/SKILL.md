---
name: ppt-mapping
description: Map presentation requirements to inspected evidence, reproducible data, page structure and appropriate visual encodings.
---
# Evidence and page mapping

Read brief.json and actual sources. Reconcile every R ID with the original request.

## Evidence
Create sources.json with S ID, title, publisher, publication/access dates, URL or local path, precise locator, version/hash when captured, and access status. Search snippets are discovery aids, not evidence of reading the source.
Create claims.json with C ID, claim text, kind (fact/derived/forecast/interpretation/supplied-unverified/illustrative), supporting and contradicting S IDs, locators, applicable period/geography/population, uncertainty, and publication disposition. Source existence does not prove claim support; inspect the cited passage.
For every D dataset, retain raw input separately from normalized data, transformations, units, denominator, period, assumptions and rounding. Derivations reference inputs and formula. Distinguish quoted forecasts from agent scenarios. Conflicting definitions must not be averaged away.

## Map
Create deck-map.json with P ID, R/C IDs, page job, evidence-calibrated title, reading order, module bounds, element IDs, encoding, dataset IDs, source footer, typography roles and planned checks. Required obligations with no page remain blockers.
Choose encoding from the relation: sequence -> process; hierarchy -> tree; time -> line; comparable amounts -> bars; conserved flows -> Sankey; qualitative relationships -> labeled diagram; supported headline magnitude -> large number with unit, period and source.
Never manufacture causal arrows from correlation, use decorative curves as data, use dual axes without distinct labels, truncate bar baselines without explicit justification, or turn unrelated totals into a Sankey. Missing data permits a labeled qualitative diagram only if it preserves the required message.

## Research-report visual contract
Use conclusion -> mechanism -> evidence -> implication as a default narrative, adapting to the material. Pages can carry multiple related modules, but each page has one dominant message. Preserve logical grouping, clear reading order, aligned chart plots and visible citations. Vary page composition with information structure, not arbitrary visual novelty. Avoid repeated equal cards when modules differ in importance.
For a reference like the supplied tall graphic, map scenario taxonomy, production/value-chain logic, quantitative evidence and scale estimates into distinct sections. Do not require its exact counts, company names or numbers for other topics.
Declare layout budgets and font roles for the delivery medium. For 16:9 projection, start with title 30–36 pt, body 18–22 pt, sources 10–12 pt; these are design defaults, not evidence of readability. Never lower them silently to meet a page count. For long-form reading, declare the actual display/print scale separately.

Before dev, check every data-bearing element has real data and every required claim has a planned factual check. Save the map and continue ppt-dev.

## Diagram-rich implementation

Paths in this section are relative to the package root. Read references/diagram-rich.zh-CN.md. Map each page to inputs, subprocesses, state owners, dependency types, outputs and applicable feedback/exception paths. Record the purpose and source of each relation before coordinates. Expand the main black box one level; no node/edge count can substitute for explanatory coverage.

## Physical-machine composition
Apply references/physical-machine.zh-CN.md. Map parent/component IDs, per-unit and system quantities, derivations, placement/orientation evidence, connection types, source/model scope and unknowns. Map each left whole-machine element to a numbered right detail tier on the same page. Separate actual target hardware, reported design, reference platform and authorized illustrative geometry.

## Output root
Store source, claim, dataset and page-map artifacts in the task’s repository-root `output/<project-name>/`. Use relative artifact paths and keep the same root as brief.json.
