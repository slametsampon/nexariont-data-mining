# NEXARIONT QC Data Mining Tool — User Guide

**Document Type:** User Guide (UG)
**Revision:** Rev0.4
**Prepared:** 2026-09-27
**Application:** NEXARIONT
**Classification:** Internal
**Approval / Effective Date:** TBD
**CURRENT Repository:** `nexariont-data-mining`
**CURRENT Branch:** `master`
**CURRENT Repository Checkpoint:** `fa86702` — `Restructure data mining repository for multi-domain development`
**QC Core Release Baseline:** `v0.2-qc-data-mining` → `de44b72` — `Fix ambiguous B+C sampling point output identity`
**Remote Repository:** `https://github.com/slametsampon/nexariont-data-mining.git`
**Reference Workflow:** `qc_multi_worksheet_workflow_reference_Rev02.md`
**Project Governance:** `00_NEXARIONT_Project_Master_Context_and_Document_Governance_Rev0.5.md`
**QC Domain Basis:** `04_NEXARIONT_Quality_Control_Information_Integration_Concept_Rev0.2.docx`

> **CURRENT authority rule:** actual source/config/code/schema/runtime/Git evidence governs CURRENT implementation.
> Workflow Rev02 governs the agreed finite workflow, checkpoint discipline, verification path, and closure principle.
> Rev0.1, Rev0.2, dan Rev0.3 tetap dipertahankan sebagai historical/provenance evidence. Rev0.4 mempertahankan CURRENT operating state Rev0.3 dan memperbarui Bab 33 agar deployment pada mesin lain dapat mengikuti dua path: **dengan Git** atau **tanpa Git**, tanpa mengubah QC mining behavior yang telah diverifikasi.

---

## Document Control

### Revision History

| Revision | Date | Description |
|---|---|---|
| Rev0.1 | Previous issue | Initial User Guide sampai verified multi-worksheet implementation pada `e874eef`; baseline Phase A-B dan introduction Phase C-D-E. |
| Rev0.2 | 2026-09-27 | Material update sampai Master-driven Sampling Point recognition pada `6b96e75`, NPG investigation/regression, cross-domain smoke evidence, Phase D guidance, dan governance boundary. |
| **Rev0.3** | **2026-09-27** | Material update dari keseluruhan working session berikutnya: complete multi-domain regression/reconciliation; fix ambiguous B+C identity; Phase D dan Phase E completion untuk agreed QC core scope; release tag `v0.2-qc-data-mining`; `Master-Data.xlsx`; monthly multi-file runner; August 2026 consolidated run; repository restructuring menjadi multi-domain root; move QC implementation ke `domains/qc`; GitHub publication; fresh clone; new `.venv`; dependency installation; dan end-to-end deployment verification. |
| **Rev0.4** | **2026-09-27** | Deployment guidance update: Bab 33 dipisahkan menjadi **Path A — Git Available** dan **Path B — Git Not Available**; ditambahkan prolog pemilihan path, manual ZIP/package acquisition untuk mesin tanpa Git, traceability boundary untuk non-Git deployment, dan common environment/setup steps setelah source tersedia. Tidak ada perubahan QC mining logic. |

### Supersession / Traceability

Rev0.4 supersedes Rev0.3 untuk **CURRENT operating use**.

Rev0.1 dan Rev0.2 tetap disimpan sebagai historical/provenance evidence karena keduanya merekam implementation state dan workflow position yang valid pada saat diterbitkan.

Rev0.4 tidak mengubah unresolved QC governance items menjadi approved requirements dan tidak mengubah QC mining logic. Item yang tidak ditetapkan oleh controlled source tetap **TBD / belum terverifikasi**.

---

# 1. Purpose

User Guide ini menjelaskan penggunaan, operating model, interpretation, verification, troubleshooting, recovery, deployment, dan evidence handling untuk **NEXARIONT QC Data Mining Tool**.

Tool membaca actual QC workbook legacy/operasional `.xls` / `.xlsx`, memproses worksheet yang eligible, mempertahankan provenance, mengenali Sampling Point (SP) secara terkontrol, menghubungkan Parameter dengan SP dan Sampling Time, dan menghasilkan **structured QC candidate data** beserta execution trace.

CURRENT operating capability juga mencakup monthly orchestration untuk memproses kumpulan workbook per minggu/domain dan menghasilkan satu consolidated monthly workbook.

Rev0.4 mempertahankan evolution implementation yang dicatat pada Rev0.3:

```text
Single-Worksheet Baseline
        ↓
Multi-Worksheet Capability
        ↓
Multi-File Regression / Exception Discovery
        ↓
Sampling Point Master Integration
        ↓
B+C Ambiguous Identity Correction
        ↓
Phase D — Acceptance & Stabilization
        ↓
Phase E — Release Baseline
        ↓
v0.2-qc-data-mining @ de44b72
        ↓
CLOSED FOR AGREED QC CORE SCOPE
        ↓
New Scoped Work:
Master-Data + Monthly Runner
        ↓
Multi-Domain Repository Restructuring
        ↓
master @ fa86702
        ↓
GitHub Publish
        ↓
Fresh Clone + New .venv + End-to-End Runtime Verification
```

Output tool **bukan otomatis**:

- validated QC fact;
- approved QC result;
- product release decision;
- QC specification compliance decision;
- formal System of Record;
- Asset Readiness input;
- evidence of causality terhadap Production atau equipment condition.

---

# 2. Goal Utama

Goal adalah membangun **QC Data Mining tool untuk NEXARIONT** yang dapat mengambil data QC dari workbook legacy/operasional secara:

- controlled;
- repeatable;
- traceable;
- tanpa mengubah original source workbook;
- tanpa merekonstruksi fakta yang tidak tersedia;
- tanpa silently cleansing / merging / discarding ambiguous source;
- dengan source workbook dan source worksheet provenance;
- dengan processing trace per worksheet;
- dengan controlled Sampling Point recognition;
- dengan structured output untuk verification / integration stage berikutnya;
- dengan batch/monthly orchestration yang tetap mempertahankan asal source.

Conceptual flow:

```text
Actual QC Workbook(s)
        ↓
Source discovery
        ↓
Worksheet discovery / eligibility
        ↓
Sampling date / layout handling
        ↓
Controlled Sampling Point recognition
        ↓
Parameter + Sampling Time association
        ↓
Candidate normalized records
        ↓
Workbook/worksheet provenance
        ↓
Processing_Log / Run_Summary
        ↓
Verification / reconciliation
        ↓
Future controlled integration
```

Authority, SoR, specification, QC approval, release criteria, KPI, validation authority, dan downstream integration yang belum ditetapkan tetap **TBD**.

---

# 3. Governance and Boundary

## 3.1 Source Authority

Untuk project facts:

```text
00 Project Master Context
        ↓
state / routing / governance / TBD
        ↓
04 Quality Control Information Integration Concept
        ↓
QC domain concept and boundary
        ↓
actual source/config/code/runtime/Git evidence
        ↓
CURRENT Data Mining behavior
```

Document 00 adalah project-level routing/governance authority dan tidak menggantikan substantive authority dokumen QC.

## 3.2 QC Boundary

Controlled QC concept menetapkan capability seperti:

- current QC information;
- historical QC information;
- trending;
- validation principle;
- source traceability;
- timestamp/freshness principle;
- data quality;
- integration logging.

Namun Data Mining Tool tidak menetapkan sendiri:

- authoritative QC source;
- formal QC System of Record;
- complete QC parameter catalogue;
- KPI;
- specification limits;
- acceptance/release criteria;
- sample/batch/lot model;
- refresh rule;
- reconciliation threshold;
- statistical/aggregation method;
- alert/deviation rule;
- retention;
- owner/approval authority;
- Production-QC detailed join key;
- QC-to-Asset-Readiness rule.

## 3.3 Candidate Data Boundary

Parser/miner output adalah:

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

## 3.4 Cross-Domain Boundary

Repository sekarang menyediakan folder untuk:

```text
QC
Production
Maintenance
SHE
```

Tetapi keberadaan folder **tidak berarti** Production, Maintenance, atau SHE Data Mining capability sudah diimplementasikan.

Tidak boleh ada automatic propagation QC rule ke domain lain.

`shared/` hanya untuk reusable technical utility yang sudah diverifikasi domain-neutral. Business rule domain tidak otomatis dipindahkan ke `shared/`.

---

# 4. CURRENT Implementation and Release State

## 4.1 QC Core Release Baseline

QC core finite workflow telah mencapai release baseline dan CLOSED untuk agreed scope.

Verified core release:

```text
Commit : de44b72
Message: Fix ambiguous B+C sampling point output identity
Tag    : v0.2-qc-data-mining
```

Tag `v0.2-qc-data-mining` tetap menunjuk ke `de44b72`.

Tag tersebut **tidak dipindahkan** ketika repository kemudian direstrukturisasi.

## 4.2 CURRENT Repository Baseline

CURRENT repository state setelah restructuring:

```text
Branch : master
HEAD   : fa86702
Commit : Restructure data mining repository for multi-domain development
Remote : origin/master
Status : clean
```

Repository remote:

```text
https://github.com/slametsampon/nexariont-data-mining.git
```

CURRENT repository checkpoint `fa86702` mencakup repository restructuring dan menambahkan current monthly-operating artifacts yang belum berada pada tag `v0.2-qc-data-mining`.

Tidak ada new formal version tag untuk `fa86702` yang ditetapkan pada working session ini.

## 4.3 Workflow Position

Workflow Rev02 memiliki finite path:

```text
Phase A — Core Parser Baseline
Phase B — Multi-Worksheet Capability
Phase C — Multi-File Regression & Exception Discovery
Phase D — Acceptance & Stabilization
Phase E — Release Baseline
CLOSED
```

CURRENT status untuk **agreed QC core scope**:

```text
Phase A   COMPLETE
Phase B   COMPLETE
Phase C   COMPLETE for agreed regression evidence
Phase D   COMPLETE
Phase E   COMPLETE
Release   v0.2-qc-data-mining @ de44b72
Status    CLOSED FOR AGREED QC CORE SCOPE
```

Capability setelah closure diperlakukan sebagai **new scoped work item / feature**, bukan automatic CP5/CP6 atau Phase F.

Repository restructuring, monthly orchestration, Master-Data consolidation, GitHub publication, dan deployment verification adalah post-closure scoped work; pekerjaan tersebut tidak membuka kembali Phase A-E.

---

# 5. Development and Recovery Checkpoints

Verified development/recovery history yang relevan:

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
d7257ce  Exclude non-candidate QC worksheets
6b96e75  Add master-driven sampling point recognition
de44b72  Fix ambiguous B+C sampling point output identity
fa86702  Restructure data mining repository for multi-domain development
```

Important release points:

```text
v0.1-baseline       historical local baseline
v0.2-qc-data-mining QC core release @ de44b72
```

Gunakan Git sebagai recovery mechanism. Jangan merekonstruksi implementation dari memory jika actual Git source dapat diperiksa.

---

# 6. CURRENT Repository Structure

CURRENT repository name:

```text
nexariont-data-mining
```

Current verified structure:

```text
nexariont-data-mining/
├── .gitignore
├── docs/
│   ├── deployment/
│   │   └── README.md
│   └── governance/
│       └── README.md
├── domains/
│   ├── maintenance/
│   │   └── README.md
│   ├── production/
│   │   └── README.md
│   ├── qc/
│   │   ├── main.py
│   │   ├── requirements.txt
│   │   ├── run_monthly_qc_mining.py
│   │   ├── doc/
│   │   │   ├── Master-Data.xlsx
│   │   │   ├── NEXARIONT_QC_Data_Mining_Handover_Checkpoint.md
│   │   │   ├── NEXARIONT_QC_Data_Mining_Phase_D_Handover_Rev01.md
│   │   │   ├── NEXARIONT_QC_Data_Mining_User_Guide_Rev0.1.md
│   │   │   ├── NEXARIONT_Sampling_Point_Master_Candidate-sam01.xlsx
│   │   │   ├── qc_multi_worksheet_workflow_reference.md
│   │   │   └── qc_multi_worksheet_workflow_reference_Rev02.md
│   │   └── src/
│   │       ├── __init__.py
│   │       ├── config.py
│   │       ├── exporter.py
│   │       ├── models.py
│   │       ├── parser.py
│   │       ├── sampling_point_master.py
│   │       ├── service.py
│   │       └── workbook_reader.py
│   └── she/
│       └── README.md
└── shared/
    └── README.md
```

Local-only/runtime locations may include:

```text
.venv/
output/
scratch/
__pycache__/
```

dan tidak dimaksudkan sebagai implementation source di Git.

`diagnostic.py` telah dipindahkan ke `scratch/` setelah trace menunjukkan tidak dipakai pada CURRENT runtime path. `scratch/` di-ignore oleh Git.

---

# 7. Execution Architecture

## 7.1 Single-Workbook Core Path

```text
QC Source Workbook (.xls / .xlsx)
        |
        v
domains/qc/main.py
        |
        +-- optional controlled Sampling Point Master
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
        |-- exclude non-visible
        |-- exclude configured non-candidate worksheet
        |-- process eligible worksheet
        |-- isolate worksheet error
        |-- retain source worksheet
        |
        v
ShiftReportParser
        |
        |-- source layout
        |-- Sampling Date
        |-- Sampling Point
        |-- Parameter
        |-- Sampling Time
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

## 7.2 Monthly Orchestration Path

CURRENT monthly operating path:

```text
Monthly Source Root
        |
        +--> Minggu I
        +--> Minggu II
        +--> Minggu III
        +--> Minggu IV
        +--> Minggu V (jika ada)
                 |
                 +--> NPG
                 +--> Octanol
                 +--> Syn Gas
                 +--> Utility
                 +--> WWT
        |
        v
domains/qc/run_monthly_qc_mining.py
        |
        |-- discover week/domain folders
        |-- discover .xls / .xlsx
        |-- invoke QC core per workbook
        |-- collect Data_Mining / Control / Processing_Log
        |-- retain month/week/domain/workbook provenance
        |-- isolate execution/output-read errors
        |
        v
Single Consolidated Monthly Workbook
        |
        +--> Data_Mining
        +--> Control
        +--> Processing_Log
        +--> Run_Summary
```

Monthly runner tidak mengubah source workbook.

---

# 8. Prerequisites

## 8.1 Verified Environment

Verified workstation environment selama final deployment test:

```text
Operating shell : Windows PowerShell
Python          : 3.14
Virtual env     : .venv
Git             : used for source/recovery/deployment
```

Python 3.14 adalah **verified environment evidence**, bukan formal minimum Python specification untuk semua future machines.

## 8.2 CURRENT QC Requirements

Fresh-clone installation verified:

```text
et_xmlfile==2.0.0
openpyxl==3.1.5
xlrd==2.0.2
```

Source:

```text
domains/qc/requirements.txt
```

Fresh clone test also completed:

```text
python -m pip check
No broken requirements found.
```

`pip` version itself is not a QC business/technical requirement unless separately baselined.

## 8.3 Supported Source Workbook Formats

Established:

```text
.xls
.xlsx
```

`.xls` is read using `xlrd`.

`.xlsx` is read using `openpyxl`.

Other formats remain unsupported/TBD unless new capability is established and verified.

---

# 9. Source Folder Model

Verified August 2026 source root example:

```text
C:\Users\sam294\Documents\DATA\Staff-ahli\nexariont\data\qca\8-AGUSTUS 2026
```

Verified weeks:

```text
MInggu I ( 02 AGU - 08 AGU  )
MInggu II ( 09 AGU - 15 AGU  )
MInggu III ( 16 AGU - 22 AGU )
MInggu IV ( 23 AGU - 29 AGU  )
```

Monthly runner is designed to detect week folders matching the controlled week naming approach, including Roman I-V / supported numeric handling in CURRENT implementation.

Source root may be local or UNC/server path.

Absolute example path above is workstation evidence, not universal installation requirement.

---

# 10. Domain Routing

CURRENT QC monthly routing:

| Source Folder | Parser Domain |
|---|---|
| `NPG` | `NPG` |
| `Octanol` | `Octanol` |
| `Syn Gas` | `Syngas` |
| `Utility` | `Utility` |
| `WWT` | `wwt` |

Important:

```text
Utility ≠ Utility NPG
```

`Utility NPG` routing remains TBD unless separately established.

Do not silently route `Utility NPG` as `Utility`.

Domain routing is parser/master context; it does not establish data ownership, QC authority, or SoR.

---

# 11. CURRENT Master Data

## 11.1 Canonical Working File

CURRENT monthly working master:

```text
domains/qc/doc/Master-Data.xlsx
```

Established worksheets:

```text
Sampling-Point
Parameter
```

## 11.2 Sampling-Point Worksheet

`Sampling-Point` replaces the older operational dependency on:

```text
NEXARIONT_Sampling_Point_Master_Candidate-sam01.xlsx
worksheet: QA Review
```

for the CURRENT monthly runner.

The historical candidate file remains in `domains/qc/doc/` as reference/provenance and is not the default monthly master.

## 11.3 Core Compatibility Adapter

QC core release was built to read Master worksheet:

```text
QA Review
```

CURRENT monthly runner therefore:

1. opens canonical `Master-Data.xlsx`;
2. requires worksheet `Sampling-Point`;
3. creates a temporary adapter workbook;
4. copies the `Sampling-Point` rows into temporary worksheet `QA Review`;
5. passes the adapter to the existing core;
6. does not modify canonical `Master-Data.xlsx`.

This preserves current core behavior without silently rewriting the released parser.

## 11.4 Parameter Worksheet Boundary

`Master-Data.xlsx` also contains:

```text
Parameter
```

The `Parameter` worksheet is maintained as working master/reference data.

It is **not automatically a parser rule** in the CURRENT core.

Do not wire Parameter Master into recognition/validation logic without a separately defined requirement, implementation change, and verification.

## 11.5 Parameter Worksheet Creation Evidence

The working `Parameter` worksheet was generated from the verified consolidated August mining output:

```text
QC_Data_Mining_8-AGUSTUS_2026.xlsx
worksheet: Data_Mining
```

The source range inspected contained:

```text
1 header row
12633 data rows
```

Working extraction result:

```text
126 unique Parameter identities
```

The extraction used Parameter identity as the de-duplication key after case-insensitive / whitespace-normalized comparison. Where the same Parameter appeared with more than one Domain or Unit, the observed source values were preserved rather than silently selecting one. Blank Unit evidence was preserved explicitly in the working extraction rather than inferred.

This extraction is a **working Master preparation result**. It does not by itself establish an approved complete QC parameter catalogue, parameter authority, or QC specification rule.

---

# 12. Controlled Sampling Point Recognition

Sampling Point recognition remains conservative.

Principles:

- no general fuzzy edit-distance inference;
- no semantic guessing;
- unique primary Sampling Point identity may be recognized;
- if primary identity is ambiguous, controlled B+C identity is used;
- similar text is not automatically same Sampling Point;
- missing description is not invented;
- duplicate/collision is not silently resolved;
- trailing sampling-time stripping is only applied under established controlled logic.

The correction at:

```text
de44b72  Fix ambiguous B+C sampling point output identity
```

was verified to prevent collapsing distinct ambiguous B+C identities.

Full cross-domain regression after the fix preserved:

```text
172 B+C rows
0 collapsed/collisions
```

for the verified regression evidence.

---

# 13. Sampling Date Handling

Parser uses dates explicitly available in source.

Established handling includes:

- expected `dd/mm/yyyy` date form;
- report-level `Date/Day` fallback for verified source variation;
- Indonesian textual month recognition where implemented.

Mandatory boundary:

```text
Do not infer Sampling Date from:
- filename
- folder name
- worksheet name
- file timestamp
- neighboring workbook
```

If usable Sampling Date evidence is unavailable, leave the condition as exception/error evidence.

---

# 14. Worksheet Discovery and Eligibility

## 14.1 Visibility

Representation:

```text
0 = visible
1 = hidden
2 = very hidden
```

Rule:

```text
Worksheet discovered
        ≠
Worksheet automatically parsed
```

Non-visible worksheet:

```text
HIDDEN / VERY HIDDEN
        ↓
EXCLUDED
```

## 14.2 Configured Non-Candidate Exclusions

Established visible worksheet exclusions include:

```text
CMKS
QA Rekap
QA Recap
```

They are logged as `EXCLUDED` rather than parsed as QC candidate worksheets.

Do not add new exclusion names without evidence and controlled change.

---

# 15. Processing Status

## 15.1 SUCCESS

Parser completed without exception.

```text
VISIBLE | SUCCESS | n records
```

`SUCCESS` is execution status, not QC approval.

## 15.2 SUCCESS + 0 Records

```text
VISIBLE | SUCCESS | 0
```

Means parser completed without exception but produced no candidate record.

It does not automatically mean:

- invalid worksheet;
- no QC information;
- software defect;
- source error.

Inspect actual source before classification.

## 15.3 EXCLUDED

Worksheet intentionally not processed because of visibility or configured exclusion.

## 15.4 ERROR

Worksheet processing exception.

Treat as:

```text
source file
+ worksheet
+ exact message
+ current code
+ actual source layout
→ finding
```

Do not immediately change parser.

---

# 16. Single-Workbook Running

For targeted/reconciliation work, core path remains available from:

```text
domains/qc/
```

Example controlled pattern:

```powershell
python .\domains\qc\main.py `
  --input "<FULL_PATH_SOURCE_WORKBOOK>" `
  --output "<OUTPUT_XLSX_PATH>" `
  --master "<MASTER_COMPATIBLE_XLSX>" `
  --domain "<DOMAIN>"
```

Use current core CLI behavior as implemented; do not assume new options beyond actual `--help`/source evidence.

For normal monthly operating use, prefer `run_monthly_qc_mining.py`.

---

# 17. Monthly Running — Recommended Operating Path

## 17.1 CLI

CURRENT verified CLI:

```text
usage: run_monthly_qc_mining.py
       --input-root INPUT_ROOT
       --output-dir OUTPUT_DIR
       [--master MASTER]
       [--output-name OUTPUT_NAME]
       [--overwrite]
```

Arguments:

- `--input-root`: monthly source root; local or UNC/server.
- `--output-dir`: destination for consolidated workbook; may be different location from source.
- `--master`: optional canonical master path; default resolves to QC `doc/Master-Data.xlsx`.
- `--output-name`: optional custom final workbook name.
- `--overwrite`: explicitly allows replacement of existing final output.

## 17.2 Local Example

```powershell
python .\domains\qc\run_monthly_qc_mining.py `
  --input-root "C:\Users\sam294\Documents\DATA\Staff-ahli\nexariont\data\qca\8-AGUSTUS 2026" `
  --output-dir "D:\NEXARIONT\QC_Output\8-AGUSTUS 2026"
```

## 17.3 Server / UNC Example Pattern

```powershell
python .\domains\qc\run_monthly_qc_mining.py `
  --input-root "\\SERVER\QC_DATA\8-AGUSTUS 2026" `
  --output-dir "D:\NEXARIONT\QC_Output\8-AGUSTUS 2026"
```

`\\SERVER\QC_DATA\...` adalah placeholder pattern. Actual server path harus menggunakan approved/actual environment evidence.

## 17.4 Default Output Name

Pattern:

```text
QC_Data_Mining_<monthly-folder-name>.xlsx
```

Example:

```text
QC_Data_Mining_8-AGUSTUS_2026.xlsx
```

---

# 18. Monthly Output Workbook

Verified monthly output sheets:

```text
Data_Mining
Control
Processing_Log
Run_Summary
```

## 18.1 Data_Mining

Retains candidate normalized records and provenance.

Monthly consolidation adds source context such as:

```text
Source Month
Source Week
Domain
Source Workbook
```

together with existing QC record context.

## 18.2 Control

Retains consolidated control/execution context from processed workbook outputs.

It does not create QC approval authority.

## 18.3 Processing_Log

Retains worksheet-level processing evidence across monthly population.

Do not remove error rows merely to make output appear clean.

## 18.4 Run_Summary

Provides workbook-level monthly orchestration evidence, including applicable context such as:

```text
Source Month
Source Week
Domain
Source Workbook
Source Path
Execution
Candidate Records
Worksheets
Success
Excluded
Error
Message
```

`Run_Summary` is execution/reconciliation evidence, not business approval.

---

# 19. Verified August 2026 Monthly Runtime Baseline

The final repository path and fresh-clone environment were tested against the same August 2026 source population.

Verified result:

```text
Month folder          : 8-AGUSTUS 2026
Weeks detected        : 4
Workbooks attempted   : 145
Candidate records     : 12633
Worksheets            : 1012
Success               : 833
Excluded              : 142
Error                 : 37
Execution failures    : 0
Output read errors    : 0
Missing domain folder : 0
Empty domain folder   : 0
```

Verified workbook:

```text
SHEETS =
['Data_Mining', 'Control', 'Processing_Log', 'Run_Summary']

DATA_RECORDS = 12633
```

The same result was reproduced:

1. after moving QC implementation to `domains/qc`;
2. after renaming root repository to `nexariont-data-mining`;
3. after publishing to GitHub;
4. from a fresh GitHub clone;
5. after creating a new `.venv`;
6. after installing `domains/qc/requirements.txt`.

This is strong deployment/relocation evidence for the tested environment.

It is not a universal business acceptance threshold for all future months.

---

# 20. Cross-Domain August Regression Evidence

The August monthly total is consistent with previously reconciled per-domain evidence:

| Domain | Workbooks | Candidate Records | Known Processing Evidence |
|---|---:|---:|---|
| Syngas | 28 | 2,987 | 308 worksheets; 232 SUCCESS; 76 EXCLUDED; no parser/execution errors in reconciled run |
| Octanol | 28 | 4,796 | 232 worksheets; 204 SUCCESS; 28 EXCLUDED; no parser/execution errors |
| NPG | 28 | 3,460 | 0 parser error; 24 unique output SP; 0 output SP outside Master |
| Utility | 33 | 1,278 | 17 worksheet ERROR; 0 execution failures |
| WWT | 28 | 112 | 20 worksheet ERROR; 0 execution failures |
| **Total** | **145** | **12,633** | **37 worksheet ERROR; 0 execution failures** |

For the 37 reconciled worksheet errors, the recurring error evidence was:

```text
ValueError: Sampling Date tidak ditemukan untuk block column 1.
```

Source inspection established missing usable Sampling Date evidence for those cases; the working session did not confirm a parser defect for this population.

Do not reconstruct dates from filename/folder/worksheet name to eliminate these errors.

---

# 21. Output Interpretation

Counters mean:

- `Candidate records` = structured candidate records produced.
- `Worksheets` = worksheet processing results.
- `Success` = parser completed for worksheet.
- `Excluded` = worksheet intentionally excluded.
- `Error` = worksheet-level processing exception.
- `Execution failures` = runner could not successfully execute a workbook-level core process.
- `Output read errors` = runner could not read expected generated workbook output.

None of these counters alone establishes QC validation or business acceptance.

---

# 22. Minimum Post-Run Verification

After each controlled monthly run:

1. review complete terminal summary;
2. confirm expected source root;
3. confirm Master path and `Sampling-Point` worksheet;
4. confirm week count;
5. confirm workbook population;
6. review `Execution failures`;
7. review `Output read errors`;
8. confirm final workbook exists;
9. confirm `Data_Mining`, `Control`, `Processing_Log`, `Run_Summary`;
10. review every `ERROR`;
11. do not treat `SUCCESS + 0` as defect without source evidence;
12. verify representative source-to-output traceability;
13. preserve output file and terminal evidence;
14. do not silently edit output to improve apparent result.

Technical check pattern:

```powershell
python -c "from openpyxl import load_workbook; from pathlib import Path; p=Path(r'<OUTPUT_XLSX>'); print('EXISTS =', p.exists()); w=load_workbook(p, read_only=True, data_only=True); print('SHEETS =', w.sheetnames); print('DATA_RECORDS =', w['Data_Mining'].max_row - 1); w.close()"
```

---

# 23. Source Value Handling

Reader/parser handling remains conservative.

Established behavior includes:

```text
Blank   → no invented value
Text    → retained unless controlled transformation applies
Numeric → retained as numeric
Boolean → retained as boolean
```

Source text such as:

```text
<0.01
-
OFF
```

must not be silently assigned new QC business meaning.

Established implementation handling includes:

- `OFF` sampling time does not create a measurement record for that slot;
- blank measurement values do not create measurement records;
- applicable structural `Total` boundary is respected;
- `STD '-'` is not silently converted to zero.

---

# 24. Traceability

Minimum single-record path:

```text
Source Month
      ↓
Source Week
      ↓
Domain
      ↓
Source Workbook
      ↓
Source Worksheet
      ↓
Sampling Point / Parameter / Time
      ↓
Candidate Record
      ↓
Data_Mining / Processing_Log / Run_Summary
```

Traceability supports verification, troubleshooting, and future controlled integration.

It does not establish source authority or System-of-Record status.

---

# 25. Evidence → Finding → Action Method

Use:

```text
EVIDENCE
   ↓
CONFIRMED
   ↓
FINDING
   ↓
MINIMUM ACTION
   ↓
VERIFICATION
```

For suspected defect:

```text
Unexpected result
      ↓
inspect actual source
      +
inspect CURRENT implementation
      ↓
confirmed defect?
   ├── NO  → retain implementation / record source behavior
   └── YES → minimum scoped correction
                ↓
             syntax/unit verification
                ↓
             affected re-test
                ↓
             reconciliation evidence
                ↓
             Git checkpoint
```

Do not modify code merely because legacy sources differ.

---

# 26. Workflow Rev02 Completion Principle

Workflow Rev02 established:

```text
Implementation
    → Regression
    → Acceptance/Stabilization
    → Release Baseline
    → CLOSED
```

Phase E expected path:

```text
Accepted feature branch
        ↓
clean working tree
        ↓
merge to master
        ↓
post-merge verification
        ↓
version tag / release baseline
        ↓
CLOSED
```

This was completed for the agreed QC core scope with:

```text
v0.2-qc-data-mining @ de44b72
```

No CP5/CP6/Phase F is created automatically.

Monthly orchestration and repository restructuring are separately scoped post-closure work.

---

# 27. Post-Closure Repository Restructuring

## 27.1 Purpose

Repository root was changed from QC-only orientation:

```text
qc-data-mining
```

to:

```text
nexariont-data-mining
```

to allow future domain-specific Data Mining capability without forcing common business rules.

## 27.2 Branch and Commit

Restructuring branch:

```text
feature/multi-domain-repository
```

Verified commit:

```text
fa86702
Restructure data mining repository for multi-domain development
```

The branch was fast-forward merged to `master`.

Final local evidence:

```text
master @ fa86702
working tree clean
```

## 27.3 Restructuring Verification

Before commit, the relocated QC runtime was tested against the full August monthly population.

After commit/merge and root rename, runtime was tested again.

Results remained:

```text
145 workbooks
12633 candidate records
1012 worksheets
833 success
142 excluded
37 error
0 execution failures
0 output read errors
```

Therefore no behavior change was identified from repository relocation in the tested environment.

---

# 28. Git Working Discipline

Before change:

```powershell
git status --short
git log -3 --oneline --decorate
```

Use scoped branch for material work.

Before commit:

```powershell
git diff --check
git status --short
```

After staging:

```powershell
git diff --cached --check
git diff --cached --stat
git diff --cached --name-status
```

Do not treat `git add -A` as sufficient review by itself; inspect staged content.

Line-ending warnings such as LF→CRLF are not automatically errors; actual policy should follow repository/site standards if later established.

---

# 29. GitHub Publication

Configured remote:

```text
origin
https://github.com/slametsampon/nexariont-data-mining.git
```

Publication commands used:

```powershell
git remote add origin https://github.com/slametsampon/nexariont-data-mining.git
git remote -v
git push -u origin master
git push origin v0.2-qc-data-mining
```

Verified:

```text
master → origin/master
local master up to date with origin/master
v0.2-qc-data-mining available on remote
```

Remote annotated tag resolution verified that:

```text
v0.2-qc-data-mining^{} → de44b72
```

Repository hosting/visibility and corporate authorization are governance matters. Technical presence on GitHub is not evidence that a hosting policy/approval has been established.

Do not store credentials, passwords, PAT tokens, raw monthly QC source, or generated production output in Git.

---

# 30. Fresh Clone Verification

A fresh clone was performed into:

```text
nexariont-data-mining-clone-test
```

Verified after clone:

```text
Branch       : master
HEAD         : fa86702
origin/master: fa86702
working tree : clean
tag          : v0.2-qc-data-mining
```

Critical files verified present:

```text
domains/qc/main.py                        True
domains/qc/run_monthly_qc_mining.py       True
domains/qc/requirements.txt                True
domains/qc/doc/Master-Data.xlsx            True
```

This established that the GitHub repository itself contains the deployment-critical QC implementation and current canonical Master file.

---

# 31. Fresh Environment Setup Verification

From fresh clone:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r .\domains\qc\requirements.txt
python -m pip check
```

Verified installation:

```text
et_xmlfile-2.0.0
openpyxl-3.1.5
xlrd-2.0.2
No broken requirements found.
```

CLI verification:

```powershell
python .\domains\qc\run_monthly_qc_mining.py --help
```

Verified options:

```text
--input-root
--output-dir
--master
--output-name
--overwrite
```

---

# 32. Fresh-Clone End-to-End Runtime Test

Fresh clone + fresh `.venv` was then run against August 2026 actual source.

Verified Master used:

```text
...\nexariont-data-mining-clone-test\domains\qc\doc\Master-Data.xlsx
```

Verified worksheet:

```text
Sampling-Point
```

Verified result:

```text
Weeks detected        : 4
Workbooks attempted   : 145
Candidate records     : 12633
Worksheets            : 1012
Success               : 833
Excluded              : 142
Error                 : 37
Execution failures    : 0
Output read errors    : 0
Missing domain folder : 0
Empty domain folder   : 0
```

Verified final workbook:

```text
EXISTS = True
SHEETS = ['Data_Mining', 'Control', 'Processing_Log', 'Run_Summary']
DATA_RECORDS = 12633
```

Git remained:

```text
On branch master
up to date with origin/master
nothing to commit, working tree clean
```

This is the current strongest deployment verification evidence from the working session.

---

# 33. Setup on Another Machine

Bab ini menyediakan **dua deployment path** sesuai kondisi mesin target:

```text
Path A — Git Available
    → source diperoleh dengan git clone
    → Git dapat digunakan untuk branch/commit/status/remote traceability

Path B — Git Not Available
    → source diperoleh sebagai ZIP / deployment package
    → runtime tetap dapat dijalankan tanpa Git
    → Git command tidak tersedia pada mesin target
```

Kedua path tersebut **bertemu pada langkah yang sama setelah source repository tersedia**:

```text
Source tersedia pada target machine
        ↓
verify Python
        ↓
create .venv
        ↓
install domains/qc/requirements.txt
        ↓
verify deployment files
        ↓
verify CLI
        ↓
verify source access
        ↓
run QC monthly mining
```

Git merupakan **deployment/source-control path**, bukan runtime dependency dari QC Data Mining Tool. Mesin target yang tidak mempunyai Git tetap dapat menjalankan tool selama source package yang benar telah tersedia dan Python environment dapat dibuat.

Ketersediaan Python **tetap harus diverifikasi** pada mesin target dengan `python --version`; UG ini tidak mengasumsikan Python tersedia pada setiap instalasi Windows.

CURRENT verified development/deployment environment menggunakan Python 3.14. Formal minimum supported Python version masih **TBD** sampai dibaseline secara terpisah.

---

## 33.1 Path A — Git Available

Gunakan path ini apabila Git tersedia dan penggunaan Git diperbolehkan pada mesin target.

### 33.1.1 Verify Prerequisite

```powershell
git --version
python --version
```

### 33.1.2 Clone Repository

```powershell
git clone https://github.com/slametsampon/nexariont-data-mining.git
cd .\nexariont-data-mining
```

### 33.1.3 Verify Git Baseline

```powershell
git status
git branch --show-current
git log -3 --oneline --decorate
```

Untuk baseline yang diverifikasi pada Rev0.4, expected reference adalah:

```text
Branch : master
HEAD   : fa86702
```

Jika repository telah berubah setelah Rev0.4, gunakan actual current Git evidence dan lakukan verification sesuai change scope; jangan menganggap `fa86702` otomatis tetap CURRENT.

Path A memberikan Git traceability langsung pada target machine.

---

## 33.2 Path B — Git Not Available

Gunakan path ini apabila Git **tidak terpasang** atau tidak dapat digunakan pada mesin target.

Git tidak perlu dipasang hanya untuk menjalankan QC Data Mining Tool jika approved/controlled source package sudah dapat dipindahkan atau di-download ke mesin target.

### 33.2.1 Verify Python

```powershell
python --version
```

Jika command tersebut tidak tersedia, Python environment harus disiapkan melalui mekanisme software installation yang berlaku pada site/IT. Exact installation mechanism/authority berada di luar scope UG ini dan tetap mengikuti kebijakan target environment.

### 33.2.2 Obtain Repository as ZIP / Deployment Package

CURRENT repository source:

```text
https://github.com/slametsampon/nexariont-data-mining
```

Untuk manual GitHub download:

```text
1. Buka repository GitHub pada browser.
2. Pastikan source/reference yang akan digunakan telah ditentukan.
3. Pilih Code.
4. Pilih Download ZIP.
5. Simpan ZIP pada target machine.
6. Extract ZIP ke folder kerja yang disetujui.
7. Masuk ke extracted repository folder.
```

Contoh:

```powershell
cd "<EXTRACTED_REPOSITORY_PATH>"
```

Nama folder hasil extraction dapat berbeda dari `nexariont-data-mining`; gunakan **actual extracted path** dan jangan mengandalkan nama folder yang diasumsikan.

### 33.2.3 Non-Git Traceability Boundary

Pada Path B tidak tersedia local Git metadata untuk menjalankan:

```text
git status
git branch
git log
git tag
git remote
```

Karena itu, source package harus diperlakukan sebagai deployment snapshot.

Untuk controlled traceability, minimum provenance yang perlu dipertahankan bersama package adalah:

```text
Repository/source identity
Reference branch / commit / package baseline jika tersedia
Download/copy date
Package filename/location
```

Untuk Rev0.4, repository baseline yang telah diverifikasi selama working session adalah:

```text
master @ fa86702
```

Namun ZIP yang di-download pada waktu lain **tidak boleh otomatis dianggap identik dengan `fa86702`** jika repository telah berubah.

Jika exact commit/reference package tidak dapat dibuktikan, status tersebut harus dicatat sebagai **belum terverifikasi**, bukan diasumsikan.

Exact formal deployment-package record format/approval authority masih TBD kecuali ditetapkan separately.

---

## 33.3 Create Python Environment — Common Path

Setelah source tersedia melalui Path A atau Path B:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r .\domains\qc\requirements.txt
python -m pip check
```

CURRENT verified requirements:

```text
et_xmlfile==2.0.0
openpyxl==3.1.5
xlrd==2.0.2
```

Expected dependency check:

```text
No broken requirements found.
```

Jika PowerShell activation dibatasi oleh target environment, ikuti applicable IT/site policy; UG ini tidak menetapkan bypass policy.

---

## 33.4 Verify Deployment Files — Common Path

Jalankan dari repository/extracted package root:

```powershell
Test-Path .\domains\qc\main.py
Test-Path .\domains\qc\run_monthly_qc_mining.py
Test-Path .\domains\qc\requirements.txt
Test-Path .\domains\qc\doc\Master-Data.xlsx
```

Expected:

```text
True
True
True
True
```

Jika salah satu `False`, jangan lanjut ke production run sampai source/package completeness diverifikasi.

---

## 33.5 Verify CLI — Common Path

```powershell
python .\domains\qc\run_monthly_qc_mining.py --help
```

CURRENT verified options:

```text
--input-root
--output-dir
--master
--output-name
--overwrite
```

CLI verification membuktikan command path dan import dasar dapat dimuat; ini bukan pengganti actual runtime test jika deployment membutuhkan functional verification.

---

## 33.6 Verify Source Access — Common Path

Untuk local path:

```powershell
Test-Path "<MONTH_SOURCE_ROOT>"
```

Untuk UNC/server path:

```powershell
Test-Path "\\SERVER\SHARE\<MONTH_FOLDER>"
Get-ChildItem "\\SERVER\SHARE\<MONTH_FOLDER>"
```

Actual server path, credential, permission, dan access authorization harus diverifikasi pada target environment. Jangan menginventarisasi atau hardcode server/share path yang belum ditetapkan.

---

## 33.7 Run QC Monthly Mining — Common Path

```powershell
python .\domains\qc\run_monthly_qc_mining.py `
  --input-root "<ACTUAL_MONTH_SOURCE_ROOT>" `
  --output-dir "<ACTUAL_OUTPUT_DIR>"
```

Optional arguments hanya digunakan sesuai actual need:

```text
--master
--output-name
--overwrite
```

Default master path tetap diarahkan ke:

```text
domains/qc/doc/Master-Data.xlsx
```

dengan Sampling Point source dari worksheet:

```text
Sampling-Point
```

Do not hardcode workstation-specific path into source code.

---

## 33.8 Post-Setup / First-Run Verification — Common Path

Setelah first controlled run pada target machine:

1. review terminal summary;
2. confirm `Execution failures`;
3. confirm `Output read errors`;
4. confirm final workbook exists;
5. confirm sheets:
   - `Data_Mining`
   - `Control`
   - `Processing_Log`
   - `Run_Summary`
6. review each `ERROR`;
7. do not treat `SUCCESS + 0` as automatic defect;
8. preserve source/output traceability;
9. compare with an applicable known baseline only when source population dan implementation reference memang comparable.

Untuk technical workbook check:

```powershell
python -c "from openpyxl import load_workbook; from pathlib import Path; p=Path(r'<OUTPUT_XLSX>'); print('EXISTS =', p.exists()); w=load_workbook(p, read_only=True, data_only=True); print('SHEETS =', w.sheetnames); print('DATA_RECORDS =', w['Data_Mining'].max_row - 1); w.close()"
```

Jika target machine menggunakan Path B tanpa Git, successful runtime **tidak membuktikan** exact Git commit identity; runtime evidence dan package provenance tetap harus dibedakan.

---

# 34. Troubleshooting Decision Guide

```text
Command fails
   |
   +-- Python/venv?
   |      → verify active interpreter and installed requirements
   |
   +-- source root?
   |      → Test-Path / permissions / actual folder
   |
   +-- no week folder?
   |      → inspect actual source structure
   |
   +-- missing domain folder?
   |      → treat as run evidence; do not invent data
   |
   +-- Master missing?
   |      → verify domains/qc/doc/Master-Data.xlsx
   |
   +-- Sampling-Point missing?
   |      → canonical Master structure issue; do not silently select another sheet
   |
   +-- worksheet EXCLUDED?
   |      → inspect visibility / configured exclusion
   |
   +-- worksheet ERROR?
   |      → inspect exact source + Processing_Log message
   |
   +-- SUCCESS + 0?
   |      → inspect source; do not assume defect
   |
   +-- unexpected SP?
   |      → compare source identity + applicable Master entry
   |
   +-- ambiguous B+C?
   |      → do not collapse or silently choose
   |
   +-- execution failure?
   |      → inspect Run_Summary + subprocess message
   |
   +-- output read error?
          → inspect generated workbook path/format and exact error
```

---

# 35. Prohibited / Unsafe Data Handling

Do not:

- modify original source workbook to fit parser;
- infer missing date from filename/folder/worksheet;
- silently route `Utility NPG` as `Utility`;
- unhide worksheet and automatically treat it as eligible;
- bypass configured exclusions without evidence;
- invent missing/ambiguous values;
- fuzzy-match SP without controlled rule;
- silently merge similar SP;
- silently resolve Master collisions;
- silently rewrite source spelling;
- promote Parameter to SP without evidence;
- convert `<0.01`, `-`, `OFF`, or other source text into unapproved business meaning;
- treat `SUCCESS` as QC approval;
- treat Master match as QC validation;
- treat candidate output as SoR;
- invent KPI/specification/acceptance threshold;
- infer Production-QC causality;
- automatically use QC output as Asset Readiness input;
- delete ERROR rows merely to obtain an apparently clean monthly result.

---

# 36. Repository Security and Data Separation

Git repository should contain controlled code/reference artifacts only.

Local/generated locations:

```text
.venv/
output/
scratch/
__pycache__/
```

are excluded from implementation tracking as applicable.

Raw monthly source workbooks are not to be copied into Git merely for execution.

Do not commit:

- credentials;
- passwords;
- GitHub PAT/token;
- network credentials;
- raw operational QC source unless separately authorized and intentionally controlled;
- generated monthly output unless an approved evidence-retention decision requires it.

---

# 37. Change Management for Future Data Mining Domains

Current root is intentionally multi-domain:

```text
domains/qc
domains/production
domains/maintenance
domains/she
shared
```

Future domain implementation must be a separately scoped capability.

Before adding Production/Maintenance/SHE:

1. identify applicable controlled source;
2. define actual source evidence;
3. keep unresolved rule TBD;
4. establish domain-specific parsing/mapping rules;
5. do not reuse QC business logic merely for convenience;
6. verify implementation against actual source;
7. use Git checkpoints;
8. preserve source traceability;
9. define domain acceptance evidence without inventing thresholds.

Production, Maintenance, and SHE current README placeholders do not establish technical or business requirements.

---

# 38. `shared/` Governance

`shared/` may contain only technical utilities demonstrated to be domain-neutral, for example potentially:

- path/file utility;
- generic logging utility;
- generic Excel helper;
- generic CLI/helper utility.

Do not automatically put into `shared/`:

- Sampling Point rules;
- QC Parameter rules;
- Production calculations;
- Maintenance event logic;
- SHE classification;
- KPI/threshold;
- domain Master data;
- validation/approval rules.

Cross-domain commonality must be proven/established, not assumed.

---

# 39. Recovery Commands

Current state:

```powershell
git status
git log --oneline --decorate
git remote -v
git tag --list
```

Inspect commit:

```powershell
git show <commit>
```

Compare:

```powershell
git diff <commit-a>..<commit-b>
```

Check patch whitespace:

```powershell
git diff --check
```

Check staged patch:

```powershell
git diff --cached --check
git diff --cached --stat
git diff --cached --name-status
```

Do not use destructive recovery command until intended state is confirmed.

---

# 40. CURRENT Known Boundaries / TBD

The following remain unresolved unless established separately:

```text
QC authoritative source                         TBD
QC formal System of Record                     TBD
QC complete parameter catalogue                TBD
QC KPI                                         TBD
QC specification logic                         TBD
QC acceptance/release criteria                 TBD
QC sample/batch/lot model                      TBD
QC refresh rule                                TBD
QC reconciliation threshold                    TBD
QC aggregation/statistical rule                TBD
QC retention                                   TBD
QC ownership / approval authority              TBD
Production-QC detailed join/correlation rule   TBD
Production-QC causality rule                   NOT ESTABLISHED
QC-to-Asset-Readiness rule                     NOT ESTABLISHED
Utility NPG source routing                     TBD
Formal minimum supported Python version        TBD
Formal production deployment host standard     TBD
Formal retention of generated mining output    TBD
Formal version tag for fa86702                 NOT ESTABLISHED
Production Data Mining implementation          NOT YET ESTABLISHED
Maintenance Data Mining implementation         NOT YET ESTABLISHED
SHE Data Mining implementation                 NOT YET ESTABLISHED
```

These gaps are not automatically software defects.

---

# 41. Operational Quick Checklist

## Before Run

- [ ] correct repository `nexariont-data-mining`;
- [ ] branch/state verified if controlled run requires it;
- [ ] `.venv` active;
- [ ] `pip check` acceptable;
- [ ] source root exists;
- [ ] actual source population understood;
- [ ] source format `.xls` / `.xlsx`;
- [ ] output directory exists/is writable;
- [ ] `domains/qc/doc/Master-Data.xlsx` exists;
- [ ] `Sampling-Point` worksheet exists;
- [ ] original source will not be modified;
- [ ] `Utility NPG` not silently routed to `Utility`.

## During Run

- [ ] preserve complete terminal summary;
- [ ] note execution failure;
- [ ] note output read error;
- [ ] do not edit source to force success.

## After Run

- [ ] final workbook exists;
- [ ] `Data_Mining` present;
- [ ] `Control` present;
- [ ] `Processing_Log` present;
- [ ] `Run_Summary` present;
- [ ] review every worksheet ERROR;
- [ ] review every execution failure;
- [ ] review every output read error;
- [ ] do not treat `SUCCESS + 0` as defect without evidence;
- [ ] reconcile candidate record count when applicable;
- [ ] verify provenance fields;
- [ ] preserve output/evidence.

---

# 42. Maintainer Decision Rules

```text
Can actual source support the conclusion?
    |
    +-- NO → TBD / unresolved
    |
    +-- YES
          ↓
Does CURRENT implementation reproduce source as intended?
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
                 +-- YES → minimum scoped correction
                              +
                         affected re-test
                              +
                       reconciliation evidence
```

Do not introduce assumptions to make data appear complete or implementation appear cleaner.

---

# 43. Handover / Continuity Package

For CURRENT work, minimum continuity package should identify:

1. `00_NEXARIONT_Project_Master_Context_and_Document_Governance_Rev0.5.md`;
2. `04_NEXARIONT_Quality_Control_Information_Integration_Concept_Rev0.2.docx`;
3. `qc_multi_worksheet_workflow_reference_Rev02.md`;
4. this User Guide Rev0.4;
5. Git repository/commit:
   - `master @ fa86702`;
   - core tag `v0.2-qc-data-mining @ de44b72`;
6. current canonical:
   - `domains/qc/doc/Master-Data.xlsx`;
7. applicable source/output artifacts for the finding under review;
8. terminal/runtime evidence where relevant.

Do not use chat summary alone as replacement for controlled sources or actual Git/runtime evidence.

---

# 44. Definition of Completion / Current Status

## 44.1 QC Core

```text
Core Implementation
        ↓
Multi-Worksheet Capability
        ↓
Multi-File Regression
        ↓
Acceptance & Stabilization
        ↓
Release Baseline
        ↓
v0.2-qc-data-mining
        ↓
CLOSED FOR AGREED QC CORE SCOPE
```

## 44.2 Current Repository/Deployment Work

```text
QC Core Closed
        ↓
Master-Data consolidation
        ↓
Monthly runner
        ↓
Multi-domain repository restructuring
        ↓
master @ fa86702
        ↓
GitHub publish
        ↓
Fresh clone
        ↓
New venv + requirements
        ↓
End-to-end August runtime test
        ↓
VERIFIED FOR TESTED DEPLOYMENT PATH
```

This does not establish universal production acceptance beyond the tested evidence.

---

# Appendix A — Historical Multi-Worksheet Reference

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

This is regression history, not universal threshold.

---

# Appendix B — NPG Regression Reference

Verified full August NPG regression:

```text
Workbooks              : 28
Output records         : 3460
Unique output SP       : 24
Output SP not in Master: 0
Parameter = Master SP  : 0
Parser errors          : 0
```

“Master SP not observed” in a population is not evidence that the Master entry is invalid.

---

# Appendix C — Post-Master Smoke Reference

Historical targeted smoke evidence:

```text
Octanol  : 125 records | 8 WS  | 7 success | 1 excluded | 0 error
Syngas   :   2 records | 10 WS | 8 success | 2 excluded | 0 error
Utility  :  44 records | 7 WS  | 7 success | 0 excluded | 0 error
wwt      :   4 records | 2 WS  | 2 success | 0 excluded | 0 error
```

These are specific smoke artifacts, not all-domain acceptance thresholds.

---

# Appendix D — August 2026 Consolidated Verification

Final verified monthly result:

```text
Input root:
C:\Users\sam294\Documents\DATA\Staff-ahli\nexariont\data\qca\8-AGUSTUS 2026

Month folder          : 8-AGUSTUS 2026
Weeks detected        : 4
Workbooks attempted   : 145
Candidate records     : 12633
Worksheets            : 1012
Success               : 833
Excluded              : 142
Error                 : 37
Execution failures    : 0
Output read errors    : 0
Missing domain folder : 0
Empty domain folder   : 0
```

Verified output:

```text
QC_Data_Mining_8-AGUSTUS_2026.xlsx
```

Sheets:

```text
Data_Mining
Control
Processing_Log
Run_Summary
```

Data records:

```text
12633
```

---

# Appendix E — Git / Release / Deployment Reference

Current core release:

```text
v0.2-qc-data-mining → de44b72
```

Current repository:

```text
master → fa86702
origin/master → fa86702
```

Remote:

```text
https://github.com/slametsampon/nexariont-data-mining.git
```

Fresh clone verification:

```text
HEAD          : fa86702
Branch        : master
origin/master : fa86702
working tree  : clean
tag available : v0.2-qc-data-mining
```

---

# Appendix F — Source Basis for Rev0.4

Rev0.4 is based on:

1. `NEXARIONT_QC_Data_Mining_User_Guide_Rev0.1.md`;
2. `NEXARIONT_QC_Data_Mining_User_Guide_Rev0.2.md`;
3. `NEXARIONT_QC_Data_Mining_User_Guide_Rev0.3.md`;
4. `qc_multi_worksheet_workflow_reference_Rev02.md`;
5. `00_NEXARIONT_Project_Master_Context_and_Document_Governance_Rev0.5.md`;
6. `04_NEXARIONT_Quality_Control_Information_Integration_Concept_Rev0.2.docx`;
7. actual Git/source/runtime evidence captured throughout the working session;
8. release baseline `v0.2-qc-data-mining @ de44b72`;
9. CURRENT repository baseline `master @ fa86702`;
10. verified August 2026 monthly execution evidence;
11. verified GitHub publish and fresh-clone deployment evidence;
12. Rev0.4 deployment-path update for target machines with and without Git.

Where Rev0.1, Rev0.2, Rev0.3, or Workflow Rev02 contains an older CURRENT checkpoint, it is retained as historical provenance. For CURRENT implementation behavior, actual current Git/runtime evidence governs.

---

**End of User Guide — Rev0.4**
