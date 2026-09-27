# NEXARIONT QC Data Mining --- Phase D Handover Checkpoint

**Handover purpose:** Continue the same QC Data Mining work in a new
chat without restarting discovery, changing the goal, or reopening
resolved decisions.

**Handover date:** 2026-09-26\
**Branch:** `feature/multi-worksheet`\
**CURRENT Git HEAD:** `6b96e75` ---
`Add master-driven sampling point recognition`\
**CURRENT workflow position:** **Phase D --- Acceptance &
Stabilization**

## 1. Goal --- DO NOT REORIENT

The goal is to build the **NEXARIONT QC Data Mining tool** that extracts
QC data from legacy/operational `.xls/.xlsx` workbooks in a
**controlled, repeatable, traceable** manner and produces **structured
QC candidate data** for later verification/integration.

The goal is **not**: - to build only an NPG parser; - to build only
Sampling Point recognition; - to repeat domain discovery; - to make all
legacy workbooks uniform; - to treat parser output as automatically
validated QC fact, QC approval, product release decision, or System of
Record.

## 2. Governing Workflow

Use `qc_multi_worksheet_workflow_reference_Rev02.md`.

Finite workflow:

`Phase A Core Parser Baseline -> Phase B Multi-Worksheet Capability -> Phase C Multi-File Regression & Exception Discovery -> Phase D Acceptance & Stabilization -> Phase E Release Baseline -> CLOSED`

CURRENT is now **Phase D** because regression evidence exists and
Phase-C work has been reconciled sufficiently to proceed with
stabilization. The workflow reference still contains an older CURRENT
checkpoint (`e874eef`) because it predates later implementation. For
CURRENT implementation, actual Git/code/runtime evidence governs.

Do **not** change the workflow, add arbitrary CP5/CP6/etc., or send the
work back to Phase C without new evidence requiring it.

## 3. CURRENT Implementation Evidence

Verified in the user's local environment:

-   Branch: `feature/multi-worksheet`
-   HEAD: `6b96e75`
-   Commit: `Add master-driven sampling point recognition`
-   Commit scope:
    -   `main.py`
    -   `src/config.py`
    -   `src/parser.py`
    -   `src/service.py`
    -   `src/sampling_point_master.py`
-   `python -m py_compile` on the changed implementation files: **PASS**
-   `git diff --check`: **PASS**
-   No tracked implementation modification remained after the commit.
-   Untracked diagnostic/support files existed and were intentionally
    not part of the implementation commit.

Implemented behavior includes: - CLI `--master` + `--domain`; - Sampling
Point Master worksheet `QA Review`; - conservative controlled
Master-driven SP recognition; - no general fuzzy inference; - primary B
identity accepted only when unique in domain; - B+C required where B is
ambiguous; - controlled B+C textual forms; - trailing `(HH.MM)` /
`(HH:MM)` stripping only when the stripped base is an exact unique
Master SP.

## 4. Phase-C Evidence Already Completed --- DO NOT REPEAT

### Multi-domain source discovery/readability

Previously established routed source domains: - Octanol -\> Octanol -
NPG -\> NPG - Syngas -\> Syn Gas - Utility -\> Utility - wwt -\> WWT

`Utility NPG` is a distinct domain from `Utility`; its source routing
remained **TBD** in the available evidence.

Do not claim that other domains were unread or unfinished merely because
NPG later required deeper SP/Parameter work.

### NPG August 2026 regression

Completed against **28 NPG workbooks**: - parser errors: **0** - output
records: **3,460** - unique output SP: **24** - output SP not in NPG
Master: **0** - Parameter that matched Master SSP: **0**

NPG regression was the major deeper regression because NPG exposed
SP-vs-Parameter classification issues that led to Master-driven
recognition.

### Targeted post-`6b96e75` smoke evidence

A later targeted smoke run was performed only to check for obvious
regression after the shared parser change: - Octanol: 125 records; 8
worksheets; 7 SUCCESS; 1 EXCLUDED; 0 ERROR - Syngas: 2 records; 10
worksheets; 8 SUCCESS; 2 EXCLUDED; 0 ERROR - Utility: 44 records; 7
worksheets; 7 SUCCESS; 0 EXCLUDED; 0 ERROR - wwt: 4 records; 2
worksheets; 2 SUCCESS; 0 EXCLUDED; 0 ERROR

This is already done. **Do not ask to rerun it without a specific
evidence-based reason.**

## 5. Sampling Point Master Governance

Working Master: `NEXARIONT_Sampling_Point_Master_Candidate-sam01.xlsx`

Key columns: - B = `Source Sampling Identity` - C =
`Sample / Stream Description`

Established: - same primary SSP with different secondary description can
be distinct; B+C is needed when B is not unique; -
`Utility != Utility NPG`; - `Nat Gas = Natural Gas` was established as
duplicate in the discussed Master cleanup; - do not silently resolve
remaining B+C collisions; - no fuzzy inference beyond established
controlled matching.

User-established NPG SP decisions must not be reopened merely because
heuristic appearance suggests otherwise. These include:
`NPG PRODUK (FLAKE)`, `V-2701`, `NPG PRODUK`, `CA-2502`, `CA 2307`,
`CA-2301 (Hydrogen)`, `X-2101 (TMA)`, `NPG Aqueous Solution QA`,
`NPG FLAKER`, `TOTANK I`, `TOTANK II`, `V-2402 (T-2401 REFLUX LIQUID)`,
`P-728`, `P-730`, `P-733`, `P-734`.

For `P-728/P-730/P-733/P-734`, time suffixes such as `(14.00)` /
`(17.00)` are removed only under the already implemented controlled rule
when the base SP is exact and unique in Master.

## 6. Anti-MNPBA / Working Discipline

Strictly: - do not invent or infer unestablished data, rules, mappings,
equivalence, acceptance thresholds, authority, SoR, or validation
status; - do not restart broad "explode every text cell" SP discovery; -
do not reopen resolved user-established SP classifications; - do not
change code merely because source data looks unusual; - require
reproducible evidence of an implementation defect before code
modification; - preserve source/workbook/worksheet traceability; -
parser output remains candidate data; - unknowns remain
`TBD / belum terverifikasi`; - distinguish CURRENT from
REQUIRED/INTENDED; - use
`evidence -> CONFIRMED -> finding -> minimum action -> verification`.

## 7. Phase D --- What To Do Now

Phase D from the workflow is **Acceptance & Stabilization**.

The governing logic is:

`Regression finding -> no defect/expected behavior: retain implementation`

or

`confirmed implementation defect -> minimum scoped correction -> syntax/unit verification as applicable -> affected regression re-test -> reconciliation evidence`

Important: - do not modify implementation merely to make legacy sources
uniform; - formal acceptance criteria, approver/authority, required
regression population, and release approval are still **TBD unless
established separately**; - therefore do not invent a formal PASS
threshold; - Phase D should first reconcile existing evidence and
identify only actual outstanding evidence-based defects/gaps; - if no
confirmed implementation defect exists, retain `6b96e75`; - do not rerun
all domains by default; - do not proceed to merge/tag Phase E until
acceptance evidence supports it and the unresolved acceptance
authority/criteria issue is handled explicitly.

## 8. Files To Attach To The New Chat

### Mandatory working set

1.  `NEXARIONT_QC_Data_Mining_Phase_D_Handover_Rev01.md` --- this
    handover; read first.
2.  `qc_multi_worksheet_workflow_reference_Rev02.md` --- governing QC
    Data Mining workflow.
3.  `NEXARIONT_QC_Data_Mining_MasterSSP_Integrated.zip` ---
    implementation snapshot corresponding to the Master-SP integrated
    code; actual local Git/runtime remains authority for CURRENT.
4.  `NEXARIONT_Sampling_Point_Master_Candidate-sam01.xlsx` --- working
    Sampling Point Master.
5.  `00_NEXARIONT_Project_Master_Context_and_Document_Governance_Rev0.5.md`
    --- project state/governance/routing authority.
6.  `04_NEXARIONT_Quality_Control_Information_Integration_Concept_Rev0.2.docx`
    --- controlled QC integration concept.

### Optional evidence files

Attach only if needed for a specific Phase-D finding: - representative
source workbook(s) involved in that finding; - corresponding parser
output workbook; - exact terminal/runtime output.

Do **not** upload all August source workbooks merely to start the new
chat. The user's local PC remains the execution environment for bulk
regression; terminal output can be supplied as runtime evidence.

## 9. Source Authority / Conflict Rule

For project facts, first use `00` for project state/governance/routing,
then the most specific controlled domain source.

For QC, use `04` for established QC integration concepts.

For this QC Data Mining implementation: - actual
source/config/code/runtime/Git evidence governs CURRENT behavior; - the
workflow reference governs the agreed phase sequence and closure path; -
if the older workflow document's recorded HEAD/state conflicts with
CURRENT Git/runtime evidence, record the discrepancy and use CURRENT
Git/runtime evidence; do not silently rewrite history.

## 10. Exact Starting Instruction For The New Chat

The next chat must: 1. read this handover first; 2. read the workflow
reference; 3. inspect the attached implementation snapshot and Master
only as needed; 4. use Project `00` and `04` for applicable
governance/QC boundaries; 5. start directly at **Phase D --- Acceptance
& Stabilization**; 6. not restart Phase C, NPG discovery, Octanol
regression, domain discovery, or generic SP candidate mining; 7. not ask
the user to reconfirm facts already established here; 8. not create a
new architecture or rewrite the parser without evidence; 9. keep
responses concise and action-oriented; 10. ask for/run only the minimum
evidence required for the next Phase-D decision.

## 11. Current Handover State

**Phase A:** COMPLETE\
**Phase B:** COMPLETE\
**Phase C:** COMPLETED for the available agreed evidence; no confirmed
new parser defect remains from the latest reconciliation.\
**Phase D:** START HERE --- Acceptance & Stabilization\
**Phase E:** NOT YET --- Release Baseline\
**CURRENT implementation:** `6b96e75` on `feature/multi-worksheet`

Do not regress the workflow without new evidence.
