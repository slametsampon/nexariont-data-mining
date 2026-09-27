# NEXARIONT QC Data Mining Tool — User Guide

**Document Type:** User Guide (UG)  
**Revision:** Rev0.2  
**Prepared:** 2026-09-27  
**Application:** NEXARIONT  
**Classification:** Internal  
**Approval / Effective Date:** TBD  
**CURRENT Development Branch:** `feature/multi-worksheet`  
**CURRENT Implementation Checkpoint:** `6b96e75` — `Add master-driven sampling point recognition`  
**Reference Workflow:** `qc_multi_worksheet_workflow_reference_Rev02.md`  
**Project Governance:** `00_NEXARIONT_Project_Master_Context_and_Document_Governance_Rev0.5.md`  
**QC Domain Basis:** `04_NEXARIONT_Quality_Control_Information_Integration_Concept_Rev0.2.docx`

> **CURRENT authority rule:** actual source/config/code/schema/runtime/Git evidence governs CURRENT implementation.  
> Workflow Rev02 governs the agreed workflow, phase sequence, checkpoints, and completion path.  
> Where the workflow or older UG still records checkpoint `e874eef`, that entry is historical/provenance; it does not replace verified CURRENT Git/runtime evidence at `6b96e75`.

---

## Document Control

### Revision History

| Revision | Date | Description |
|---|---|---|
| Rev0.1 | Previous issue | Initial User Guide through verified multi-worksheet implementation at `e874eef`; scope baseline through Phase B and introduction of Phase C-D-E workflow. |
| **Rev0.2** | **2026-09-27** | Material update to CURRENT implementation and operating evidence: report-level `Date/Day` fallback; configured worksheet exclusions; Sampling Point Master integration; domain routing; NPG SP-vs-Parameter investigation; Master-driven recognition; NPG 28-workbook regression; cross-domain evidence; Phase D working guidance; output/evidence interpretation; repository and handover controls. |

### Supersession / Traceability

Rev0.2 updates Rev0.1 for CURRENT use. Rev0.1 remains historical/provenance evidence for the implementation state at `e874eef`.

This revision does **not** silently convert unresolved QC governance items into approved requirements. Items not established by controlled sources remain **TBD**.

---

# 1. Purpose

User Guide ini menjelaskan penggunaan, interpretation, verification, troubleshooting, recovery, dan evidence handling untuk **NEXARIONT QC Data Mining Tool**.

Tool dikembangkan untuk membaca actual QC workbook legacy/operasional `.xls` / `.xlsx`, memproses worksheet yang eligible, mempertahankan provenance, mengenali Sampling Point (SP) secara terkontrol, menghubungkan Parameter dengan SP dan Sampling Time, lalu menghasilkan **structured QC candidate data** beserta execution trace.

Rev0.2 mencakup perkembangan dari:

```text
Single-Worksheet Baseline
        ↓
Multi-Worksheet Capability
        ↓
Multi-File Regression / Exception Discovery
        ↓
NPG SP-vs-Parameter Investigation
        ↓
Sampling Point Master Integration
        ↓
CURRENT implementation @ 6b96e75
        ↓
Phase D — Acceptance & Stabilization
```

Output tool **bukan otomatis**:

- validated QC fact;
- approved QC result;
- product release decision;
- Quality Control specification compliance decision;
- formal System of Record;
- Asset Readiness input;
- evidence of causality terhadap Production atau equipment condition.

---

# 2. Goal Utama

Goal tool **bukan sekadar membuat parser worksheet atau mengenali Sampling Point**.

Goal adalah membangun **QC Data Mining tool untuk NEXARIONT** yang dapat mengambil data QC dari workbook legacy/operasional secara:

- controlled;
- repeatable;
- traceable;
- tanpa mengubah original source workbook;
- tanpa merekonstruksi fakta yang tidak tersedia;
- tanpa silently cleansing / merging / discarding ambiguous source;
- dengan source workbook dan source worksheet provenance;
- dengan processing trace per worksheet;
- dengan controlled Sampling Point recognition ketika Master SSP digunakan;
- dengan structured output yang dapat digunakan pada verification / integration stage berikutnya.

```text
Actual QC Workbook
        ↓
Worksheet discovery / eligibility
        ↓
Sampling date / layout handling
        ↓
Sampling Point recognition
        ↓
Parameter + Sampling Time association
        ↓
Candidate normalized records
        ↓
Provenance + Processing_Log
        ↓
Verification / reconciliation
        ↓
Future controlled integration
```

Authority/source/specification/validation/acceptance/downstream integration yang belum ditetapkan tetap **TBD**.

---

# 3. Governance and Boundary

## 3.1 Source Authority

Untuk project facts:

```text
00 Project Master Context
        ↓
identify current state / routing / applicable source
        ↓
04 Quality Control Information Integration Concept
        ↓
QC domain concept and boundary
        ↓
actual implementation/runtime/Git evidence
        ↓
CURRENT data-mining behavior
```

Document 00 adalah routing/governance authority, tetapi tidak menggantikan substantive authority dokumen domain.

## 3.2 Quality Control Boundary

Controlled QC concept menetapkan capability:

- Quality Control information integration;
- current information;
- historical information;
- trending;
- data validation principle;
- source traceability;
- timestamp/freshness principle;
- data quality;
- integration logging.

Namun item berikut belum ditetapkan secara umum dan **tidak boleh diisi oleh Data Mining Tool**:

- authoritative QC source system;
- complete QC parameter catalogue;
- KPI;
- specification limits;
- acceptance/release criteria;
- sample/batch/lot relationship;
- refresh interval;
- reconciliation method/threshold;
- statistical/aggregation method;
- alert/deviation rule;
- retention;
- business/data/source-system owner;
- approval authority;
- Production-QC detailed relationship;
- QC-to-Asset-Readiness relationship;
- formal System of Record.

## 3.3 Candidate Data Boundary

Parser output adalah:

```text
candidate normalized QC record
```

bukan otomatis:

```text
validated QC fact
approved QC result
product release decision
formal System-of-Record record
Asset Readiness evidence
```

---

# 4. CURRENT Implementation State

## 4.1 Git State

Verified local Git evidence pada checkpoint yang digunakan untuk Rev0.2:

```text
Branch : feature/multi-worksheet
HEAD   : 6b96e75
Commit : Add master-driven sampling point recognition
```

Commit `6b96e75` mencakup perubahan implementation pada:

```text
main.py
src/config.py
src/parser.py
src/sampling_point_master.py
src/service.py
```

Verification yang telah dijalankan:

```text
python -m py_compile ...   PASS
git diff --check           PASS
```

## 4.2 Workflow Position

```text
Phase A — Core Parser Baseline                 COMPLETE
Phase B — Multi-Worksheet Capability           COMPLETE
Phase C — Regression / Exception Discovery     COMPLETE for available/agreed evidence
Phase D — Acceptance & Stabilization           CURRENT
Phase E — Release Baseline                     NOT YET
```

**Catatan:** Workflow Rev02 sendiri masih merekam `e874eef` sebagai checkpoint pada saat dokumen workflow dibuat. Sesuai governance workflow, actual runtime/code/Git evidence menguasai CURRENT implementation bila berbeda.

## 4.3 No Automatic Phase Expansion

Finite completion path tetap:

```text
Implementation
    → Regression
    → Acceptance / Stabilization
    → Release Baseline
    → CLOSED
```

Tidak ada automatic CP5/CP6/CP7 atau phase baru tanpa scoped need yang explicitly established.

---

# 5. Development Evolution / Recovery Checkpoints

Verified development history yang tetap relevan:

```text
5afe138  Baseline: direct XLS shift-pagi parser verified 74 records
0486509  Add workbook worksheet discovery
1f7c992  Add worksheet visibility filtering
5fc2936  Add multi-worksheet diagnostic processing
3e7150d  Add worksheet provenance model
75261a9  Add multi-worksheet service orchestration
69deea8  Add multi-worksheet Excel export
375e9a3  Add multi-worksheet processing log
e874eef  Integrate multi-worksheet main execution path
3984a65  Add report-level Date/Day fallback
d7257ce  Add configured worksheet exclusions
6b96e75  Add master-driven sampling point recognition
```

Gunakan Git sebagai recovery mechanism. Jangan membuat reconstruction implementation berdasarkan memory jika local Git source dapat diperiksa.

---

# 6. Execution Architecture

CURRENT functional architecture secara konseptual:

```text
QC Source Workbook (.xls / .xlsx)
        |
        v
main.py
        |
        +-- optional controlled Sampling Point Master
        |       |
        |       +-- master workbook
        |       +-- domain
        |
        v
WorkbookReader
        |
        |-- .xls reader
        |-- .xlsx reader
        |-- worksheet discovery
        |-- visibility metadata
        |
        v
QCDataMiningService.process_all_worksheets()
        |
        |-- exclude non-visible worksheet
        |-- exclude configured worksheet names
        |-- process eligible worksheet
        |-- isolate error per worksheet
        |-- retain source worksheet
        |
        v
ShiftReportParser
        |
        |-- detect/read source layout
        |-- sampling date
        |-- Sampling Point
        |-- Parameter
        |-- Sampling Time
        |-- candidate SamplingRecord
        |
        v
SourcedSamplingRecord
        |
        v
ExcelExporter.export_multi_worksheet()
        |
        +--> Data_Mining
        +--> Control
        +--> Processing_Log
```

Tool saat ini memproses **satu source workbook per CLI invocation**. Batch multi-file dilakukan melalui external shell loop / orchestration evidence script, bukan dengan menganggap satu output workbook otomatis merupakan multi-domain consolidated file.

---

# 7. Repository / Working Directory

Verified working repository:

```text
C:\Users\sam294\Documents\DATA\Staff-ahli\nexariont\toll\qc-data-mining
```

Typical working structure:

```text
qc-data-mining/
├── .venv/
├── doc/
├── output/
├── src/
│   ├── __init__.py
│   ├── config.py
│   ├── exporter.py
│   ├── models.py
│   ├── parser.py
│   ├── sampling_point_master.py
│   ├── service.py
│   └── workbook_reader.py
├── main.py
├── requirements.txt
├── .gitignore
└── diagnostic / investigation scripts
```

## 7.1 File Classification

**CURRENT implementation:**

- `main.py`
- `src/*.py`
- applicable dependency/config files.

**Diagnostic/investigation artifacts:**

- `diagnose_*.py`
- `triage_*.py`
- `phase_c_reconcile.py`
- temporary/reference artifacts such as `parser_current.txt`.

Diagnostic artifacts dapat penting sebagai historical investigation evidence, tetapi bukan otomatis production implementation.

**Generated output:**

- `output/`

**Do not package as implementation source:**

- `.venv/`
- `__pycache__/`
- generated cache.

## 7.2 Git Safety

Banyak diagnostic files pernah berada sebagai untracked files. Karena itu:

```text
DO NOT automatically use:
git add .
```

tanpa review.

Gunakan:

```powershell
git status --short
git diff --check
git log -1 --oneline
```

dan stage hanya file yang memang berada dalam intended change scope.

---

# 8. Prerequisites

Verified development environment:

```text
Python virtual environment : .venv
Command shell              : PowerShell
Repository                 : qc-data-mining
```

Established source formats:

```text
.xls
.xlsx
```

Libraries used by implementation include:

- `xlrd` for `.xls`;
- `openpyxl` for `.xlsx`, Sampling Point Master reading, and output Excel.

Aktifkan virtual environment sesuai workstation configuration:

```powershell
& .\.venv\Scripts\Activate.ps1
```

---

# 9. Source and Master Paths Used During Development

Verified August 2026 source root:

```text
C:\Users\sam294\Documents\DATA\Staff-ahli\nexariont\data\qca\8-AGUSTUS 2026
```

Verified working Sampling Point Master:

```text
C:\Users\sam294\Documents\DATA\Staff-ahli\nexariont\data\qca\NEXARIONT_Sampling_Point_Master_Candidate-sam01.xlsx
```

Verified August week folders:

```text
MInggu I ( 02 AGU - 08 AGU  )
MInggu II ( 09 AGU - 15 AGU  )
MInggu III ( 16 AGU - 22 AGU )
MInggu IV ( 23 AGU - 29 AGU  )
```

These absolute paths are **verified workstation examples**, not universal installation requirements.

---

# 10. Domain Routing

Routing established during development:

| Data Mining Domain | Source Folder / Domain Context |
|---|---|
| `Octanol` | `Octanol` |
| `NPG` | `NPG` |
| `Syngas` | `Syn Gas` |
| `Utility` | `Utility` |
| `wwt` | `WWT` |
| `Utility NPG` | **UNROUTED / TBD in available evidence** |

Important:

```text
Utility ≠ Utility NPG
```

Do not silently route `Utility NPG` to `Utility`.

Domain routing digunakan untuk memilih Master SSP context. Ia bukan QC business ownership atau System-of-Record decision.

---

# 11. Sampling Point Master

## 11.1 Purpose

Sampling Point Master digunakan sebagai **controlled parser recognition reference** untuk membedakan Sampling Point dari Parameter/structural text pada source yang human-typed dan tidak selalu konsisten.

Master SSP adalah supporting mechanism; **bukan goal utama Data Mining Tool**.

Working Master:

```text
NEXARIONT_Sampling_Point_Master_Candidate-sam01.xlsx
```

Master ini adalah working candidate reference. Keberadaan record di Master tidak dengan sendirinya membuatnya approved QC fact atau authoritative QC master.

## 11.2 Worksheet and Columns

Established worksheet:

```text
QA Review
```

Established relevant columns:

| Column | Meaning |
|---|---|
| A | Domain / Worksheet |
| B | Source Sampling Identity |
| C | Sample / Stream Description |

Column B adalah primary source Sampling Point identity.

Column C adalah secondary description yang dapat diperlukan untuk membedakan record bila primary B tidak unik.

## 11.3 Controlled Recognition Principle

Recognition dilakukan secara konservatif.

Established working principles:

- no general fuzzy edit-distance inference;
- no semantic guessing;
- case / spacing / punctuation normalization dapat dipakai sesuai CURRENT implementation;
- unique primary identity dapat dikenali;
- bila primary identity ambigu, secondary description diperlukan;
- similar text **tidak otomatis** berarti same Sampling Point;
- duplicated/ambiguous Master entries tidak boleh silently resolved;
- missing description tidak boleh diisi dengan asumsi;
- literal `(blank)` pada working Master diperlakukan sebagai tidak adanya secondary description pada current implementation evidence.

## 11.4 Important NPG Identity Decisions

User-established Sampling Point decisions yang tidak boleh dibuka ulang tanpa new evidence antara lain:

```text
NPG PRODUK (FLAKE)
V-2701
NPG PRODUK
CA-2502
CA 2307
CA-2301 (Hydrogen)
X-2101 (TMA)
NPG Aqueous Solution QA
NPG FLAKER
TOTANK I
TOTANK II
V-2402 (T-2401 REFLUX LIQUID)
P-728
P-730
P-733
P-734
```

Important distinctions:

```text
NPG FLAKER ≠ automatically NPG PRODUK (FLAKE)
```

`CA 2307` follows current Master spelling; do not silently rename to `CA-2307` unless Master is deliberately revised.

## 11.5 Trailing Sampling-Time Annotation

Established examples:

```text
P-728 (14.00) → P-728
P-730 (14.00) → P-730
P-733 (17.00) → P-733
P-734 (17.00) → P-734
```

For these established cases, the trailing time is not part of Sampling Point identity.

Do **not** generalize this decision to arbitrary parenthetical descriptions. Arbitrary text in parentheses may be part of Sampling Point identity or description and must remain evidence-driven.

## 11.6 Duplicate / Collision Handling

A duplicate or collision in Master is not permission to choose one silently.

Example discovered during Master review included duplicated `B+C` identity for an Evonik entry. Such conditions remain ambiguous unless explicitly resolved by controlled Master maintenance.

---

# 12. Worksheet Discovery, Visibility, and Explicit Exclusions

## 12.1 Worksheet Discovery

WorkbookReader discovers worksheet names and visibility metadata.

```text
Worksheet discovered
        ≠
Worksheet automatically parsed
```

## 12.2 Visibility

Implementation representation:

```text
0 = visible
1 = hidden
2 = very hidden
```

Non-visible worksheet:

```text
HIDDEN
  ↓
EXCLUDED
```

Visibility filtering is an implementation eligibility rule, not a QC validity decision.

## 12.3 Configured Worksheet Exclusions

In addition to hidden/non-visible filtering, configured visible worksheet exclusions were established:

```text
CMKS
QA Rekap
QA Recap
```

These worksheets are logged as `EXCLUDED`, not parsed as candidate QC data worksheets.

Do not add new exclusion names without evidence and controlled implementation change.

---

# 13. Sampling Date Handling

## 13.1 Primary Date Pattern

Existing parser supports source dates such as:

```text
dd/mm/yyyy
```

within the expected source/header area.

## 13.2 Report-Level `Date/Day` Fallback

NPG source variation established a report-level date such as:

```text
Date/Day : Senin, 24 Agustus 2026
```

A controlled fallback was added to parse Indonesian textual month names from the report-level `Date/Day` field.

Established principle:

- use date explicitly present in source;
- do not infer date from filename;
- do not infer date from worksheet name;
- do not infer date from file timestamp;
- do not reconstruct missing date from surrounding files.

If source date remains unavailable/ambiguous, treat it as an exception; do not invent it.

---

# 14. Sampling Point → Parameter Structure

The parser treats Sampling Point as parent/context and analytical rows below it as Parameters within the source block.

Conceptual example:

```text
Sampling Point
    ├── Parameter A
    ├── Parameter B
    ├── Parameter C
    └── ...
```

A verified NPG structural example showed:

```text
CA-4115 (NPG Aqueous Solution)
    ↓
MeOH
TMA
...
```

and:

```text
CA-2203 ...
    ↓
pH
```

This parent/parameter relationship is based on source layout evidence. Parameter rows must not be promoted to Sampling Point merely because they contain text.

Similarly, structural headings, units, values, signoff metadata, and section labels must not be treated as Sampling Points without evidence.

---

# 15. Running the Tool

## 15.1 CURRENT Recommended Master-Controlled Run

When a controlled Master and domain routing are available:

```powershell
python .\main.py `
  --input "<FULL_PATH_SOURCE_WORKBOOK>" `
  --output "<OUTPUT_XLSX_PATH>" `
  --master "<FULL_PATH_MASTER_SSP_XLSX>" `
  --domain "<DOMAIN>"
```

`--master` and `--domain` are a pair in the CURRENT implementation path. Do not supply only one of them.

Example pattern for NPG:

```powershell
python .\main.py `
  --input "C:\...\NPG\2-SenHin-NPG.xls" `
  --output ".\output\masterssp_npg_2_senin_verify.xlsx" `
  --master "C:\...\NEXARIONT_Sampling_Point_Master_Candidate-sam01.xlsx" `
  --domain "NPG"
```

`masterssp_npg_2_senin_verify.xlsx` is a **single-input NPG verification artifact**. It is not expected to contain Octanol, Syngas, Utility, or WWT.

## 15.2 Baseline / Structural Mode

The historical CLI without Master:

```powershell
python .\main.py `
  --input "<FULL_PATH_SOURCE_WORKBOOK>" `
  --output "<OUTPUT_XLSX_PATH>"
```

remains relevant as baseline/recovery behavior where supported by CURRENT code.

For CURRENT acceptance/stabilization work on domains using Master SSP, use the agreed Master-driven path unless a specific comparison requires otherwise.

## 15.3 One Invocation = One Source Workbook

The tool does not establish that a single invocation processes an entire folder/domain.

For multiple workbooks:

```text
PowerShell loop
    ↓
main.py once per source workbook
    ↓
one output workbook per source workbook
```

Do not interpret a single verification file as a cross-domain consolidated output.

---

# 16. Console Result

Typical console summary:

```text
QC DATA MINING COMPLETED
--------------------------------------------------
Input     : <source path>
Output    : <output path>
Records   : <candidate record count>
Worksheets: <worksheet result count>
Success   : <success count>
Excluded  : <excluded count>
Error     : <error count>
--------------------------------------------------
```

Interpretation:

- `Records` = candidate normalized records produced;
- `Worksheets` = worksheet processing results;
- `Success` = parser completed for worksheet;
- `Excluded` = worksheet intentionally not processed;
- `Error` = worksheet processing exception.

None of these counters constitutes QC business acceptance.

---

# 17. Processing Status

## 17.1 SUCCESS

Parser completes without exception.

```text
VISIBLE | SUCCESS | n records
```

`SUCCESS` does not mean QC validation/approval.

## 17.2 SUCCESS + 0 Records

Parser completes without exception but produces no candidate record.

```text
VISIBLE | SUCCESS | 0
```

Do not automatically classify as:

- invalid worksheet;
- missing data;
- software defect;
- no QC information.

Inspect source and expected layout first.

## 17.3 EXCLUDED

Worksheet is not processed because:

- non-visible; or
- explicitly configured exclusion.

Processing_Log message differentiates the reason.

## 17.4 ERROR

Worksheet processing throws an exception.

Treat as a finding:

```text
source file
worksheet
error message
CURRENT code
actual source/layout
```

Do not immediately modify parser or source.

---

# 18. Output Workbook

Verified multi-worksheet output structure:

```text
Data_Mining
Control
Processing_Log
```

## 18.1 Data_Mining

Current logical headers:

```text
Source Worksheet
Tanggal
Jam
Sampling Point (SP)
Parameter
Unit
Min
Max
Nilai
```

`Source Worksheet` retains record provenance.

When Master SSP canonical recognition is used, output Sampling Point may use the Master sampling identity for recognized source text.

This canonicalization is parser identity handling; it does not constitute QC approval or business reclassification.

## 18.2 Control

`Control` records minimum execution/output context, including source file and candidate record status.

It does not establish QC business authority, validation, acceptance rule, or release decision.

## 18.3 Processing_Log

Header:

```text
Worksheet
Visibility
Status
Record Count
Message
```

Processing_Log is execution evidence, not QC approval evidence.

---

# 19. Source Value Handling

Reader/parser handling remains conservative.

Established behavior includes:

```text
Blank   → no invented value
Text    → retained as source text unless a controlled parser transformation applies
Numeric → retained as numeric
Boolean → retained as boolean
```

Source text such as:

```text
<0.01
-
OFF
```

must not be silently converted into new QC business meaning.

Established parser handling includes:

- `OFF` sampling time does not create a measurement record for that sampling slot;
- blank measurement values do not create measurement records;
- `Total` is used as a source structural boundary where applicable;
- `STD '-'` is preserved as source content and is not converted to zero.

---

# 20. Traceability

Minimum record-level path:

```text
Source Workbook
      |
      v
Source Worksheet
      |
      v
Sampling Point / Parameter / Time
      |
      v
Candidate SamplingRecord
      |
      +--> Data_Mining
      |
      +--> Processing_Log
```

Traceability supports verification, troubleshooting, and future controlled integration.

It does not establish source authority or System-of-Record status.

---

# 21. Minimum Post-Run Verification

After every controlled run:

1. review console `Records`, `Worksheets`, `Success`, `Excluded`, and `Error`;
2. confirm output workbook exists;
3. confirm `Data_Mining`, `Control`, and `Processing_Log`;
4. review all `Processing_Log` rows;
5. investigate every `ERROR`;
6. do not treat `SUCCESS + 0` as defect without source evidence;
7. confirm `Source Worksheet` is retained;
8. sample source-to-output reconciliation where verification requires it;
9. where Master SSP is used, investigate unexpected SP identity rather than silently editing output;
10. preserve source and output artifact names for traceability.

Optional technical verification pattern:

```powershell
python -c "from openpyxl import load_workbook; p=r'<OUTPUT.xlsx>'; wb=load_workbook(p,data_only=True,read_only=True); print(wb.sheetnames); dm=wb['Data_Mining']; pl=wb['Processing_Log']; print('DATA_MINING=',dm.max_row-1); print('PROCESSING_LOG=',pl.max_row-1); wb.close()"
```

---

# 22. Phase A and Phase B Verified Baseline

## 22.1 Phase A — Core Parser

Verified original baseline:

```text
Worksheet : shift-pagi
Records   : 74
```

Git baseline:

```text
5afe138
```

## 22.2 Phase B — Multi-Worksheet Capability

Reference source `3-Selasa_oct.xls` established:

```text
Worksheets discovered : 9
Visible candidates     : 8
Hidden/excluded        : 1
Success                : 8
Error                  : 0
Candidate records      : 232
```

Reference hidden worksheet:

```text
dbase → HIDDEN → EXCLUDED
```

This remains historical/reference evidence and not a universal acceptance threshold.

---

# 23. Phase C — Multi-File Regression & Exception Discovery

## 23.1 Purpose

Phase C verifies that implementation behavior is not dependent on a single reference workbook.

Workflow:

```text
CURRENT implementation
        ↓
actual source workbook
        ↓
run
        ↓
capture evidence
        ↓
review Data_Mining / Processing_Log
        ↓
source-to-output reconciliation
        ↓
classify finding
        ├── expected behavior
        ├── source/layout exception
        └── suspected implementation defect
```

Minimum evidence:

```text
Source file
Worksheets discovered
Visible candidates
Excluded worksheets
Success count
Error count
Candidate record count
Processing_Log result
Relevant source/output reconciliation finding
```

Complete-population acceptance threshold remains TBD unless explicitly established.

## 23.2 Important Interpretation Learned During Phase C

```text
Domain/workbook can be read
        ≠
all worksheets clean
        ≠
candidate records validated
        ≠
formal acceptance
```

Historical regression evidence showed that some earlier Utility and WWT workbooks had worksheet errors, while Octanol/Syngas examples could complete cleanly. Therefore the statement “all domains have been read” must not be converted into “all complete domain populations are fully accepted.”

---

# 24. NPG Investigation and Stabilization Evidence

NPG became the main exception-discovery case because its source structure exposed a clearer **Sampling Point vs Parameter** distinction and source/layout variation.

## 24.1 Structural Discovery

At an earlier Master state:

```text
NPG workbooks scanned    : 28
Worksheets scanned       : 166
Recognized SP cells      : 743
Rows captured below SP   : 10126
Ambiguous SP cells       : 0
Scan errors              : 0
```

These diagnostic results were investigation evidence, not final QC validation.

## 24.2 Verified Missing-Header Cases

A key workbook:

```text
2-SenHin-NPG.xls
```

included worksheets previously failing on structural assumptions.

After the final current implementation work, verified runtime result became:

```text
Records   : 118
Worksheets: 4
Success   : 4
Excluded  : 0
Error     : 0
```

including successful processing of previously problematic:

```text
Senin MALAM
AWR Senin MALAM
```

The operational lesson is that CURRENT behavior must be judged from runtime evidence at local HEAD, not from older parser copies or diagnostic snapshots.

## 24.3 Full August NPG Regression

Verified against 28 NPG workbooks:

```text
WORKBOOKS             : 28
OUTPUT RECORDS         : 3460
UNIQUE OUTPUT SP       : 24
OUTPUT SP NOT IN MASTER: 0
PARAMETER = MASTER SSP : 0
PARSER ERRORS          : 0
```

At the recorded Master checkpoint:

```text
MASTER NPG RECORDS     : 36
MASTER SSP OBSERVED    : 24
MASTER SSP NOT OBSERVED: 12
```

“Not observed” means not observed in that regression population; it does **not** mean invalid Master record.

## 24.4 NPG Output Interpretation

NPG full regression established useful structural invariants for the tested population, but it does not by itself establish:

- QC value correctness for every record;
- source authority;
- specification compliance;
- QC approval;
- complete business acceptance.

---

# 25. Cross-Domain Evidence

## 25.1 Earlier Cross-Domain Discovery

Cross-domain diagnostic work demonstrated source coverage across routed domains and Master observations.

This work was used for investigation and Master alignment; it was not a formal all-domain acceptance test.

## 25.2 Post-`6b96e75` Targeted Smoke Evidence

After Master-driven recognition integration, targeted current-head smoke runs produced:

| Domain | Records | Worksheets | Success | Excluded | Error |
|---|---:|---:|---:|---:|---:|
| Octanol | 125 | 8 | 7 | 1 | 0 |
| Syngas | 2 | 10 | 8 | 2 | 0 |
| Utility | 44 | 7 | 7 | 0 | 0 |
| wwt | 4 | 2 | 2 | 0 | 0 |

Interpretation:

- no parser exception appeared in these specific smoke workbooks;
- low/high record count alone is not acceptance evidence;
- this does not prove every workbook in every domain is clean;
- no automatic code change is justified only because record counts differ from older runs.

## 25.3 Utility NPG

`Utility NPG` remains separate from `Utility`.

Routing remains TBD on available evidence and must not be silently assigned.

---

# 26. Phase D — Acceptance & Stabilization

## 26.1 CURRENT Position

Phase D is CURRENT because regression evidence exists.

Workflow Rev02:

```text
Regression finding
      |
      +-- no defect / expected behavior
      |       ↓
      |   RETAIN IMPLEMENTATION
      |
      +-- confirmed implementation defect
              ↓
        minimum scoped correction
              ↓
        syntax/unit verification
              ↓
        affected regression re-test
              ↓
        reconciliation evidence
```

## 26.2 What Phase D Is Not

Phase D does **not** automatically mean:

- rerun every workbook;
- reopen Sampling Point discovery;
- normalize every legacy layout;
- change code until all outputs look identical;
- invent acceptance thresholds;
- declare business approval.

## 26.3 Evidence Packaging

Phase D may include retained CURRENT artifacts per domain when explicitly needed for acceptance/stabilization evidence.

A domain-specific output file is exactly that: **domain/workbook-specific evidence**.

Example:

```text
masterssp_npg_2_senin_verify.xlsx
```

contains NPG because it was generated from one NPG input. The absence of other domains in that file is expected.

Other domains are represented by their own run/output artifacts unless a separately designed and verified cross-workbook consolidation capability is established.

## 26.4 CURRENT User-Selected Phase D Priority

At the end of the development session captured by this revision, the user selected:

```text
Octanol first
Syngas second
```

for Phase D evidence work before continuing with other domain cases.

This is a working evidence priority, not a universal workflow requirement.

## 26.5 Outstanding Acceptance Governance

The following remain TBD unless established separately:

- formal acceptance criteria;
- required regression population;
- acceptance threshold;
- approver/authority;
- release approval;
- exact Phase E version/tag.

---

# 27. Handling a Suspected Defect

Do not modify code based only on a surprising number or unusual source layout.

Use:

```text
Evidence
   ↓
CONFIRMED observation
   ↓
Finding
   ↓
actual source + CURRENT code inspection
   ↓
confirm defect?
   ├── NO  → retain implementation / record source behavior
   └── YES → minimum scoped correction
                ↓
             syntax/unit verification
                ↓
             affected re-test only
                ↓
             reconciliation evidence
                ↓
             Git recovery checkpoint
```

A source difference is not automatically a software defect.

---

# 28. Troubleshooting Decision Guide

```text
Command fails
   |
   +-- file/path problem?
   |      → verify actual path/file
   |
   +-- Master path/sheet problem?
   |      → verify actual Master file and QA Review
   |
   +-- --master without --domain or vice versa?
   |      → provide both together
   |
   +-- unsupported input format?
   |      → .xls / .xlsx only unless new capability is established
   |
   +-- worksheet EXCLUDED?
   |      → inspect visibility or configured exclusion
   |
   +-- worksheet ERROR?
   |      → inspect source + exact Processing_Log message
   |
   +-- SUCCESS but 0?
   |      → inspect source; do not assume defect
   |
   +-- unexpected Sampling Point?
   |      → compare source text + applicable Master entry
   |
   +-- ambiguous Master entry?
   |      → do not silently select one
   |
   +-- unexpected record?
          → source/output reconciliation
          → preserve provenance
          → classify before code change
```

---

# 29. Prohibited / Unsafe Data Handling

Do not:

- modify original workbook to make it fit the parser;
- unhide a worksheet and automatically treat it as eligible;
- bypass configured exclusions without evidence;
- invent missing/ambiguous value;
- infer a Sampling Point only because text “looks like” a tag;
- silently merge similar Sampling Points;
- silently deduplicate Master collisions;
- silently rewrite Master spelling;
- turn a Parameter into Sampling Point without evidence;
- reconstruct missing date from filename;
- convert `<0.01`, `-`, `OFF`, or other source text into unapproved business meaning;
- treat parser `SUCCESS` as QC approval;
- treat Master recognition as QC validation;
- treat candidate record as System-of-Record record;
- create QC specification/threshold/acceptance rule that is not established;
- infer Production-QC causality;
- automatically use QC output as Asset Readiness input.

---

# 30. Local Execution Model for Regression / Acceptance Work

Actual QC workbooks reside on the user's Windows workstation.

Operating model:

```text
Chat / engineering review
        ↓
precise PowerShell/Python command
        ↓
user executes in local .venv
        ↓
stdout + output workbook
        ↓
runtime evidence
        ↓
engineering evaluation
```

Absence of local QC workbook inside a ChatGPT runtime is **not** automatically a project blocker.

Bulk source data does not need to be uploaded merely to continue controlled local verification.

Upload/share only the specific source/output artifact required for a finding when direct inspection is necessary.

---

# 31. Batch Execution Pattern

Generic controlled batch pattern:

```powershell
$master = "<FULL_PATH_MASTER>"
$files = Get-ChildItem "<DOMAIN_FOLDER>" -File |
    Where-Object { $_.Extension -in ".xls", ".xlsx" }

foreach ($f in $files) {
    $out = Join-Path ".\output\<RUN_FOLDER>" ($f.BaseName + "_output.xlsx")

    python .\main.py `
        --input $f.FullName `
        --output $out `
        --master $master `
        --domain "<DOMAIN>"
}
```

Before using a recursive batch:

- verify source root;
- verify domain folder;
- avoid output folders becoming input candidates;
- preserve unique output names;
- record the exact run population;
- do not silently include `Utility NPG` under `Utility`.

---

# 32. Evidence Artifact Management

Recommended evidence categories:

```text
output/
├── baseline / reference outputs
├── phase_c diagnostic outputs
├── phase_c regression outputs
├── phase_d current acceptance/stabilization outputs
└── finding-specific reconciliation artifacts
```

Evidence file names should identify:

- phase/run purpose;
- domain;
- source workbook identity where practical;
- implementation checkpoint if required for controlled retention.

Exact naming standard is not yet established and therefore remains a working convention unless formally baselined.

---

# 33. Recovery and Verification Commands

## 33.1 Current Git

```powershell
git log -1 --oneline
git status --short
git show --stat --oneline 6b96e75
```

Expected CURRENT checkpoint:

```text
6b96e75 Add master-driven sampling point recognition
```

## 33.2 Syntax Verification

For current changed implementation scope:

```powershell
python -m py_compile `
  .\src\config.py `
  .\src\sampling_point_master.py `
  .\src\parser.py `
  .\src\service.py `
  .\main.py
```

No stdout normally indicates compile success.

## 33.3 Whitespace / Patch Check

```powershell
git diff --check
```

## 33.4 Inspect Historical Commit

```powershell
git show <commit>
```

## 33.5 Compare Commits

```powershell
git diff <commit-a>..<commit-b>
```

Do not use destructive recovery commands until the intended state is verified.

---

# 34. Phase E — Release Baseline

Phase E occurs only after acceptance evidence supports proceeding.

Expected workflow:

```text
Accepted feature branch
        ↓
verify clean working tree
        ↓
merge feature/multi-worksheet → master
        ↓
post-merge verification
        ↓
version tag / release baseline
        ↓
CLOSED
```

Exact release decision, approver, and version/tag remain TBD until explicitly established.

---

# 35. Operational Quick Checklist

## Before Run

- [ ] correct repository;
- [ ] `.venv` active;
- [ ] source workbook exists;
- [ ] source format `.xls` / `.xlsx`;
- [ ] correct output path;
- [ ] original source will not be modified;
- [ ] if Master-controlled run: correct Master file;
- [ ] if Master-controlled run: correct domain;
- [ ] `--master` and `--domain` provided together;
- [ ] `Utility NPG` not silently routed to `Utility`.

## During Run

- [ ] retain complete console summary;
- [ ] note non-zero process exit;
- [ ] do not edit source to force success.

## After Run

- [ ] review `Data_Mining`;
- [ ] review `Control`;
- [ ] review every `Processing_Log` row;
- [ ] investigate each `ERROR`;
- [ ] do not treat `SUCCESS + 0` as defect without evidence;
- [ ] confirm `Source Worksheet`;
- [ ] reconcile representative source/output;
- [ ] check unexpected SP against Master;
- [ ] preserve evidence artifact.

---

# 36. User / Maintainer Decision Rules

Use the following concise decision logic:

```text
Can actual source support the conclusion?
    |
    +-- NO → TBD / unresolved
    |
    +-- YES
          ↓
Does CURRENT implementation reproduce source structure as intended?
          |
          +-- YES → retain behavior
          |
          +-- NO / exception
                 ↓
             investigate
                 ↓
         confirmed software defect?
                 |
                 +-- NO → source/layout finding
                 |
                 +-- YES → minimum scoped correction + affected re-test
```

No assumption is introduced merely to make the dataset appear complete.

---

# 37. Handover / Continuity Guidance

When work is transferred to another chat/session or maintainer, minimum continuity package should contain:

1. current handover/checkpoint document;
2. `qc_multi_worksheet_workflow_reference_Rev02.md`;
3. exact CURRENT local Git source snapshot or Git reference for `6b96e75`;
4. CURRENT `NEXARIONT_Sampling_Point_Master_Candidate-sam01.xlsx`;
5. access to Project `00` governance source;
6. access to Project `04` QC concept;
7. only the specific source/output evidence required for the CURRENT finding.

Do not use conversation summary alone as replacement for controlled sources or CURRENT Git evidence.

A packaged source ZIP must be verified against local Git HEAD before it is treated as CURRENT implementation authority.

---

# 38. Known Current Boundaries / TBD

At Rev0.2, the following remain intentionally unresolved unless separately established:

```text
QC authoritative source                         TBD
QC System of Record                             TBD
QC parameter catalogue                          TBD
QC KPI                                          TBD
QC specification logic                          TBD
QC acceptance/release criteria                  TBD
QC sample/batch/lot model                       TBD
QC refresh rule                                 TBD
QC reconciliation threshold                     TBD
QC aggregation/statistical rule                 TBD
QC retention                                    TBD
Production-QC detailed join/correlation rule    TBD
Production-QC causality rule                    NOT ESTABLISHED
QC-to-Asset-Readiness rule                      NOT ESTABLISHED
Utility NPG source routing                      TBD
Phase D formal acceptance threshold             TBD
Phase D formal required population              TBD
Phase D approver/authority                      TBD
Phase E release approval                        TBD
Phase E version/tag                             TBD
```

These gaps are not parser defects.

---

# 39. Definition of Completion

The QC Data Mining implementation scope follows the finite workflow:

```text
Core Implementation
        ↓
Multi-Worksheet Capability
        ↓
Multi-File Regression / Exception Discovery
        ↓
Acceptance & Stabilization
        ↓
Release Baseline
        ↓
CLOSED FOR AGREED SCOPE
```

A new capability after closure is a new scoped work item/feature, not an automatic continuation of the current checkpoint series.

---

# Appendix A — Reference Multi-Worksheet Case

Historical verified reference:

```text
Source                  : 3-Selasa_oct.xls
Worksheets discovered   : 9
Visible candidates      : 8
Hidden/excluded         : 1
Success                 : 8
Errors                  : 0
Candidate records       : 232
Hidden worksheet        : dbase
```

Output:

```text
Data_Mining     : 232 candidate records
Control         : present
Processing_Log  : 9 worksheet results
```

This remains regression history, not universal acceptance threshold.

---

# Appendix B — NPG Current-Head Evidence Summary

Verified full August regression:

```text
Workbooks       : 28
Records         : 3460
Unique output SP: 24
Parser errors   : 0
SP outside Master: 0
Parameter=Master SP: 0
```

Example one-workbook CURRENT verification:

```text
Source     : 2-SenHin-NPG.xls
Records    : 118
Worksheets : 4
Success    : 4
Excluded   : 0
Error      : 0
```

---

# Appendix C — Post-Master Cross-Domain Smoke Summary

```text
Octanol  : 125 records | 8 WS  | 7 success | 1 excluded | 0 error
Syngas   :   2 records | 10 WS | 8 success | 2 excluded | 0 error
Utility  :  44 records | 7 WS  | 7 success | 0 excluded | 0 error
wwt      :   4 records | 2 WS  | 2 success | 0 excluded | 0 error
```

These values are evidence from specific smoke sources only.

---

# Appendix D — Phase D Working Position at Rev0.2

```text
Phase A  COMPLETE
Phase B  COMPLETE
Phase C  COMPLETE for available/agreed evidence
Phase D  CURRENT
Phase E  NOT YET

CURRENT local Git:
feature/multi-worksheet @ 6b96e75
```

Working priority explicitly selected at the end of the referenced development session:

```text
1. Octanol
2. Syngas
```

for Phase D current-evidence packaging/stabilization before proceeding to later affected cases.

This priority does not establish a formal required regression population.

---

# Appendix E — Source Basis for Rev0.2

Rev0.2 is based on:

1. `NEXARIONT_QC_Data_Mining_User_Guide_Rev0.1.md`;
2. `qc_multi_worksheet_workflow_reference_Rev02.md`;
3. `00_NEXARIONT_Project_Master_Context_and_Document_Governance_Rev0.5.md`;
4. `04_NEXARIONT_Quality_Control_Information_Integration_Concept_Rev0.2.docx`;
5. actual development/runtime/Git evidence recorded through the referenced QC Data Mining working session;
6. CURRENT local Git checkpoint `6b96e75` as reported and verified in the local execution environment.

Where an older document records `e874eef` as CURRENT, Rev0.2 preserves that value as historical provenance while using verified runtime/Git evidence for CURRENT implementation, consistent with Workflow Rev02 governance.

---

**End of User Guide — Rev0.2**
