# NEXARIONT QC Data Mining — Handover Checkpoint

**Purpose:** Working-state handover for continuation in a new chat without repeating resolved discovery.

## 1. Governing Working Principle

- Anti-MNPBA applies strictly: do not invent, infer, reconstruct, silently cleanse, merge, discard, or redefine unestablished data/rules.
- CURRENT implementation is governed by actual source/config/code/runtime/Git evidence.
- User-established decisions below are treated as established working facts and must not be reopened without contradictory evidence or explicit user instruction.
- Unknown/unestablished matters remain **TBD / belum terverifikasi**.
- Do not modify code merely because a case looks unusual. Require evidence of an implementation defect.
- Preserve source traceability. Parser output remains candidate output, not automatic QC validation/approval.

## 2. Current Workflow Position

- Phase A — Core Parser Baseline: **COMPLETE**.
- Phase B — Multi-Worksheet Capability: **COMPLETE**.
- Phase C — Multi-File Regression & Exception Discovery: **ACTIVE**.
- NPG regression: completed to the evidence recorded in this checkpoint.
- **NEXT: continue Phase C with Octanol regression.**
- Subsequent domains: Syngas → Utility → Utility NPG → WWT, subject to established routing/evidence.
- Utility NPG is distinct from Utility. Source routing for Utility NPG remains **TBD unless subsequently established**.
- After applicable cross-domain regression/reconciliation: proceed to Phase D — Acceptance & Stabilization. Do not invent acceptance population or thresholds.

## 3. Current Git / Implementation Checkpoint

**Branch:** `feature/multi-worksheet`
**HEAD:** `6b96e75`
**Commit:** `Add master-driven sampling point recognition`

Committed implementation files only:

- `main.py`
- `src/config.py`
- `src/parser.py`
- `src/service.py`
- `src/sampling_point_master.py`

Diagnostic scripts and other untracked working files were intentionally **not committed**.

### Implemented behavior

- CLI accepts `--master` and `--domain` together.
- Sampling Point Master worksheet configured as `QA Review`.
- Parser can use `SamplingPointMaster` for controlled SP recognition.
- Matching is deliberately conservative; no general fuzzy inference.
- Primary SSP alone is accepted only when unique in the selected domain.
- Where the same primary SSP occurs with different secondary descriptions, B alone is ambiguous and B+C is required.
- Controlled B+C textual forms are recognized.
- Explicit trailing `(HH.MM)` / `(HH:MM)` stripping is supported only when the stripped base is itself an exact unique Master SSP in that domain.

## 4. Sampling Point Master Working Governance

Current working master used during the completed NPG regression:

`NEXARIONT_Sampling_Point_Master_Candidate-sam01.xlsx`

Relevant columns:

- Column B — `Source Sampling Identity`: **primary SSP identity**.
- Column C — `Sample / Stream Description`: secondary discriminator/context.

Established rules:

- Same primary SSP with different secondary description may represent distinct Master records; B+C is required where B is not unique.
- Human typing variations such as spacing, hyphenation and case can represent the same primary SSP when controlled matching establishes it.
- `Utility` **is not the same domain as** `Utility NPG`.
- `Nat Gas = Natural Gas` was established as duplicate in the discussed Master cleanup.
- Do not silently resolve remaining B+C collisions. Example previously observed: duplicate `Evonik | Non Permeate STG 1` entries remained in the Master at that stage; ambiguity must not be guessed away.

## 5. User-Established NPG Sampling Points

The following were explicitly established by the user as Sampling Points during the work. Do **not** reclassify them as parameters/metadata merely from heuristic appearance:

- `NPG PRODUK (FLAKE)`
- `V-2701`
- `NPG PRODUK`
- `CA-2502`
- `CA 2307`
- `CA-2301 (Hydrogen)`
- `X-2101 (TMA)`
- `NPG Aqueous Solution QA`
- `NPG FLAKER`
- `TOTANK I`
- `TOTANK II`
- `V-2402 (T-2401 REFLUX LIQUID)`
- `P-728`
- `P-730`
- `P-733`
- `P-734`

For `P-728/P-730/P-733/P-734`, source forms containing sampling-time suffixes such as `(14.00)` or `(17.00)` are treated as the corresponding SP with the time suffix removed, provided the base SP exists uniquely in the controlled Master.

Previously confirmed `NPG PRODUK (FLAKE)` and `V-2701` were discovered in source evidence; they must not be sent back through generic candidate review merely because they were absent from an earlier Master revision.

## 6. NPG August 2026 Regression — Confirmed Runtime Evidence

Source population processed: **28 NPG workbooks** across Weeks I–IV of August 2026.

Regression result:

- Workbooks: **28**
- Parser errors: **0**
- Output records: **3,460**
- Unique output SP: **24**
- Output records whose SP was not in NPG Master: **0**
- Cases where output Parameter itself matched a Master SSP: **0**

The 24 output SP were:

- `CA 2307`
- `CA-1101`
- `CA-1102`
- `CA-1103`
- `CA-1104`
- `CA-2101`
- `CA-2102`
- `CA-2201`
- `CA-2202`
- `CA-2203`
- `CA-2301 (Hydrogen)`
- `CA-2302`
- `CA-2303`
- `CA-2306`
- `CA-2401`
- `CA-2404`
- `CA-2501`
- `CA-2502`
- `CA-2503`
- `CA-4101`
- `CA-4115`
- `NPG Aqueous Solution QA`
- `TOTANK I`
- `X-2101 (TMA)`

### Structural verification

At the time of verification, NPG Master contained **36 records**:

- Master SSP observed in output: **24**
- Master SSP not observed in output: **12**
- Parameter = Master SSP: **0**

The 12 Master SSP not observed in output were:

- `CA-2402` — `T-2401 Reflux Liquid`
- `Flaker` — `Produk NPG Flake`
- `CA-2403` — `T-2402 Bottom Liquid`
- `CA-4105` — `V-4102 (HPN Tank)`
- `NPG PRODUK`
- `NPG FLAKER`
- `TOTANK II`
- `V-2402 (T-2401 REFLUX LIQUID)`
- `P-728`
- `P-730`
- `P-733`
- `P-734`

**Interpretation boundary:** `NOT OBSERVED` means only that the Master SSP did not produce output in this regression population. It is **not automatically a defect**, not proof that the SP is absent from all source data, and not a basis for deleting it from Master.

## 7. Representative NPG Verification Evidence

For `2-SenHin-NPG.xls`, controlled-master parsing produced:

- Records: **118**
- Worksheets processed: **4**
- Success: **4**
- Excluded: **0**
- Error: **0**

Observed output structure demonstrated SP → Parameter separation, including examples such as:

- `CA-1103` → `Formaldehyde`, `MeOH`, `H2O`, `Formic Acid`
- `CA-1101` → `MeOH`, `H2O`
- `CA-2102` → multiple parameters beneath the SP
- `CA-2202` → `pH`
- `CA-2303` → `FORMALDHYDE`

A complete audit of that workbook found **0 output SP outside Master**.

## 8. Important Negative / Boundary Findings

Do not repeat broad "explode every text cell" candidate discovery as if every text value might be an SP. Prior generic discovery produced obvious non-SP content such as units, values, report metadata, dates, section headings and parameters. That approach was intentionally superseded by controlled Master-driven recognition and structural evidence.

Examples of text that must not become SP merely because they appear frequently include units/values/report metadata/section headings such as `wt %`, `<0.01`, `Unit`, `Spv. Production`, company/report titles, dates, and area headings. Parameter names such as `H2`, `O2`, `N2`, `CH4`, etc. must be interpreted from source structure rather than promoted to SP without controlled evidence.

Do not reopen already user-established SP decisions listed in Section 5.

## 9. Immediate Next Action

**Continue Phase C with Octanol regression using commit `6b96e75` as the implementation baseline.**

Working method:

1. Run current implementation against the agreed Octanol source population.
2. Preserve workbook/worksheet/source traceability.
3. Compare output SP against the controlled Master for the `Octanol` domain.
4. Identify implementation exceptions only from actual evidence.
5. Do not change code before a reproducible parser/layout defect is demonstrated.
6. Do not infer missing SSP/description/equivalence.
7. If an unresolved source/master ambiguity is encountered, report it as **TBD / review required**, not as an inferred mapping.
8. Once Octanol regression evidence is stable, continue Phase C to the next established domain rather than returning to resolved NPG discovery.

## 10. Instruction to the Next Chat

Start from this checkpoint. **Do not restart NPG discovery. Do not ask the user to reconfirm decisions already recorded here. Do not propose a new architecture or rewrite the parser without runtime evidence.**

Use this operating sequence:

**evidence → CONFIRMED → finding → minimum action → verification**.

Keep CURRENT separate from REQUIRED/INTENDED. Keep unknowns unknown. When the user provides new runtime output, treat that output as CURRENT evidence and advance from it rather than repeating prior tests.

---

**Handover state:** Phase C ACTIVE — NPG regression checkpoint established at Git `6b96e75`; **NEXT = Octanol regression**.
