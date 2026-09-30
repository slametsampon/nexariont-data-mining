# NEXARIONT QC Data Mining — README

**Application:** NEXARIONT  
**Domain:** Quality Control (QC)  
**Repository:** `nexariont-data-mining`  
**Current repository branch:** `master`  
**Current repository checkpoint:** `64e394d`  
**QC core release baseline:** `v0.2-qc-data-mining @ de44b72`  
**Classification:** Internal

> README ini merupakan operational/technical quick reference untuk QC Data Mining. Untuk penjelasan lengkap, governance, historical evidence, dan acceptance history, gunakan User Guide dan controlled project documents yang berlaku.

## 1. Purpose

NEXARIONT QC Data Mining digunakan untuk membaca workbook QC legacy/operasional `.xls` dan `.xlsx`, memproses worksheet yang eligible, mempertahankan provenance, dan menghasilkan structured QC candidate data beserta execution trace.

Tool menyediakan dua operating path:

1. **Single-workbook processing** melalui `main.py`.
2. **Monthly multi-workbook processing** melalui `run_monthly_qc_mining.py`.

Output tool adalah **candidate normalized QC data** dan execution evidence. Output tersebut **bukan otomatis** validated QC fact, approved QC result, product release decision, QC specification compliance decision, formal System of Record, Asset Readiness evidence, atau evidence of causality terhadap Production/equipment condition.

## 2. Current Architecture

### 2.1 Single-Workbook Path

```text
QC Source Workbook
        ↓
domains/qc/main.py
   (thin CLI)
        ↓
QCWorkbookProcessor
        ↓
Existing QC Core
        ├── MiningConfig
        ├── SamplingPointMaster
        ├── WorkbookReader
        ├── QCDataMiningService
        ├── ShiftReportParser
        └── ExcelExporter
        ↓
Single Workbook Output
        ├── Data_Mining
        ├── Control
        └── Processing_Log
```

`main.py` hanya menangani CLI/presentation dan mendelegasikan actual workbook processing ke `QCWorkbookProcessor`.

### 2.2 Monthly Path

```text
Monthly Source Root
        ↓
domains/qc/run_monthly_qc_mining.py
   (thin CLI)
        ↓
MonthlyMiningOrchestrator
        ├── MonthlySourcePlanner
        ├── QCWorkbookProcessor
        └── MonthlyWorkbookConsolidator
        ↓
Single Consolidated Monthly Workbook
        ├── Data_Mining
        ├── Control
        ├── Processing_Log
        └── Run_Summary
```

Responsibility:

```text
MonthlySourcePlanner
    → week/domain/workbook discovery
    → deterministic work plan

QCWorkbookProcessor
    → process satu workbook end-to-end

MonthlyMiningOrchestrator
    → coordinate monthly execution
    → temporary workspace lifecycle
    → master adapter coordination
    → workbook failure isolation
    → EXEC_FAILED
    → continue next workbook

MonthlyWorkbookConsolidator
    → read temporary workbook output
    → consolidate monthly result
    → provenance
    → totals
    → Run_Summary
    → OUTPUT_READ_ERROR
```

Active monthly path **tidak bergantung pada runtime execution `main.py`** dan tidak menggunakan `subprocess` untuk memanggil single-workbook CLI.

## 3. Current Repository Structure

```text
nexariont-data-mining/
└── domains/
    └── qc/
        ├── main.py
        ├── run_monthly_qc_mining.py
        ├── requirements.txt
        ├── doc/
        │   ├── Master-Data.xlsx
        │   ├── NEXARIONT_QC_Data_Mining_User_Guide_Rev0.5.docx
        │   ├── NEXARIONT_QC_Mining_Refactor_Concept_Workflow_Gates_Rev0.2.md
        │   └── ...
        └── src/
            ├── config.py
            ├── exporter.py
            ├── models.py
            ├── parser.py
            ├── sampling_point_master.py
            ├── service.py
            ├── workbook_reader.py
            ├── workbook_processor.py
            ├── monthly_source_planner.py
            ├── monthly_mining_orchestrator.py
            └── monthly_workbook_consolidator.py
```

## 4. Environment

### 4.1 Verified Environment

```text
Operating shell : Windows PowerShell
Python          : 3.14
Virtual env     : .venv
```

Python 3.14 adalah **verified environment evidence**, bukan formal minimum supported Python version. Formal minimum supported Python version masih **TBD**.

### 4.2 Python Requirements

CURRENT QC requirements:

```text
et_xmlfile==2.0.0
openpyxl==3.1.5
xlrd==2.0.2
```

Install:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
python -m pip install -r .\domains\qc\requirements.txt
python -m pip check
```

Expected dependency check:

```text
No broken requirements found.
```

## 5. Supported Input Format

Established source workbook formats:

```text
.xls
.xlsx
```

Reader:

```text
.xls   → xlrd
.xlsx  → openpyxl
```

Other source formats remain unsupported/TBD unless separately established and verified.

## 6. Master Data

CURRENT canonical monthly working master:

```text
domains/qc/doc/Master-Data.xlsx
```

Established worksheets:

```text
Sampling-Point
Parameter
```

### 6.1 Sampling-Point

`Sampling-Point` digunakan sebagai canonical working Sampling Point source untuk monthly flow.

Monthly application membuat temporary compatibility adapter:

```text
Master-Data.xlsx
    ↓
Sampling-Point
    ↓
temporary adapter workbook
    ↓
QA Review
    ↓
QCWorkbookProcessor / existing QC core
```

Canonical `Master-Data.xlsx` tidak dimodifikasi.

### 6.2 Parameter Worksheet

`Parameter` saat ini merupakan working master/reference data. `Parameter` **bukan otomatis parser rule** dan tidak boleh digunakan sebagai recognition/validation logic tanpa separately defined requirement, implementation change, dan verification.

## 7. Single-Workbook Processing — `main.py`

### 7.1 CLI

```text
usage: main.py [-h]
               --input INPUT
               --output OUTPUT
               [--master MASTER]
               [--domain DOMAIN]
```

Arguments:

```text
--input     Path source XLS/XLSX
--output    Path output XLSX
--master    Path Sampling Point Master XLSX
--domain    Domain master context
```

`--master` dan `--domain` harus diberikan bersama-sama.

### 7.2 Basic Example

Tanpa master/domain:

```powershell
python .\domains\qc\main.py `
  --input "D:\QC\source.xls" `
  --output "D:\QC\output.xlsx"
```

Dengan controlled master/domain:

```powershell
python .\domains\qc\main.py `
  --input "D:\QC\source.xls" `
  --output "D:\QC\output.xlsx" `
  --master "D:\QC\Sampling_Point_Master.xlsx" `
  --domain "Octanol"
```

Gunakan actual domain/master evidence. Jangan menginventarisasi domain atau master path.

### 7.3 Single-Workbook Output

Expected workbook sheets:

```text
Data_Mining
Control
Processing_Log
```

- `Data_Mining`: candidate normalized QC records.
- `Control`: control/execution context dari current exporter.
- `Processing_Log`: worksheet-level processing evidence (`SUCCESS`, `EXCLUDED`, `ERROR`).

## 8. Monthly Processing — `run_monthly_qc_mining.py`

Untuk normal monthly operating use, gunakan monthly runner.

### 8.1 CLI

```text
usage: run_monthly_qc_mining.py [-h]
                                --input-root INPUT_ROOT
                                --output-dir OUTPUT_DIR
                                [--master MASTER]
                                [--output-name OUTPUT_NAME]
                                [--overwrite]
```

Arguments:

```text
--input-root
    Monthly source root; local atau UNC/server path.

--output-dir
    Destination folder untuk consolidated monthly XLSX.
    Boleh berbeda dari source location.

--master
    Optional canonical QC master.
    Default: domains/qc/doc/Master-Data.xlsx

--output-name
    Optional custom final workbook filename.

--overwrite
    Mengizinkan replacement final output yang sudah ada.
```

### 8.2 Local Example

```powershell
python .\domains\qc\run_monthly_qc_mining.py `
  --input-root "C:\QC_DATA\8-AGUSTUS 2026" `
  --output-dir "D:\NEXARIONT\QC_Output\8-AGUSTUS 2026"
```

### 8.3 UNC / Server Example Pattern

```powershell
python .\domains\qc\run_monthly_qc_mining.py `
  --input-root "\\SERVER\QC_DATA\8-AGUSTUS 2026" `
  --output-dir "D:\NEXARIONT\QC_Output\8-AGUSTUS 2026"
```

`\\SERVER\QC_DATA\...` adalah pattern contoh. Gunakan actual approved/available network path dan access authorization pada target environment.

### 8.4 Custom Output Name

```powershell
python .\domains\qc\run_monthly_qc_mining.py `
  --input-root "C:\QC_DATA\8-AGUSTUS 2026" `
  --output-dir "D:\NEXARIONT\QC_Output" `
  --output-name "QC_Data_Mining_8-AGUSTUS_2026.xlsx"
```

### 8.5 Overwrite Existing Output

```powershell
python .\domains\qc\run_monthly_qc_mining.py `
  --input-root "C:\QC_DATA\8-AGUSTUS 2026" `
  --output-dir "D:\NEXARIONT\QC_Output" `
  --output-name "QC_Data_Mining_8-AGUSTUS_2026.xlsx" `
  --overwrite
```

Gunakan `--overwrite` hanya jika replacement memang disengaja.

## 9. Monthly Source Folder Model

Verified source pattern:

```text
<MONTH ROOT>
├── MInggu I (...)
│   ├── NPG
│   ├── Octanol
│   ├── Syn Gas
│   ├── Utility
│   └── WWT
├── MInggu II (...)
├── MInggu III (...)
├── MInggu IV (...)
└── MInggu V (...)   # jika ada
```

Monthly runner mempertahankan deterministic week discovery/order, domain routing/folder matching, `.xls/.xlsx` discovery, temporary Excel exclusion sesuai CURRENT implementation, workbook ordering, missing/no-workbook evidence, dan month/week/domain/workbook provenance.

## 10. CURRENT Domain Routing

```text
Source Folder   Parser Domain
-------------   -------------
NPG             NPG
Octanol         Octanol
Syn Gas         Syngas
Utility         Utility
WWT             wwt
```

Important:

```text
Utility ≠ Utility NPG
```

Utility NPG routing tetap **TBD** kecuali separately established. Jangan silently route Utility NPG sebagai Utility.

## 11. Monthly Runtime Progress

Monthly runner menampilkan progress selama execution, misalnya:

```text
=== MInggu I (...) ===
[NPG] 7 workbook(s)
  workbook-1.xls
  workbook-2.xls

[Octanol] 7 workbook(s)
  workbook-1.xls
  ...
```

Presentation dilakukan oleh CLI. Execution state dikoordinasikan oleh `MonthlyMiningOrchestrator`.

## 12. Monthly Output

Final consolidated monthly workbook mempunyai sheets:

```text
Data_Mining
Control
Processing_Log
Run_Summary
```

### 12.1 Data_Mining

Berisi consolidated candidate QC records dan provenance seperti `Source Month`, `Source Week`, `Domain`, `Source Workbook`, dan `Source Worksheet` bersama current QC record fields.

### 12.2 Control

Berisi consolidated control/execution context. Control sheet tidak menetapkan QC approval authority.

### 12.3 Processing_Log

Berisi worksheet-level processing evidence dari seluruh source workbooks. Jangan menghapus ERROR hanya untuk membuat output terlihat bersih.

### 12.4 Run_Summary

Berisi workbook-level monthly orchestration evidence seperti:

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

`Run_Summary` adalah execution/reconciliation evidence, bukan business approval.

## 13. Processing Status

### SUCCESS

Parser completed tanpa exception.

```text
VISIBLE | SUCCESS | n records
```

SUCCESS adalah execution status, bukan QC approval.

### SUCCESS + 0 Records

```text
VISIBLE | SUCCESS | 0
```

Artinya parser completed tanpa exception tetapi menghasilkan 0 candidate records. Tidak otomatis berarti invalid worksheet, tidak ada QC information, software defect, atau source error. Inspect actual source sebelum classification.

### EXCLUDED

Worksheet tidak diproses karena visibility atau configured exclusion.

Established visible worksheet exclusions mencakup:

```text
CMKS
QA Rekap
QA Recap
```

Jangan menambah exclusion baru tanpa evidence dan controlled change.

### ERROR

Worksheet-level processing exception. Investigasi menggunakan source workbook + worksheet + exact message + current implementation + actual source layout. Jangan langsung mengubah parser.

### EXEC_FAILED

Workbook-level execution failure yang diisolasi oleh `MonthlyMiningOrchestrator`. Monthly execution melanjutkan work item berikutnya.

### OUTPUT_READ_ERROR

Temporary workbook output tidak dapat dibaca/di-consolidate oleh `MonthlyWorkbookConsolidator`. Inspect exact output path, workbook condition, dan error message.

## 14. Verified August 2026 Regression Reference

Verified comparable source population:

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

Verified output sheets:

```text
Data_Mining
Control
Processing_Log
Run_Summary
```

Refactor regression verification juga memperoleh:

```text
Data_Mining     PASS
Control         PASS
Processing_Log  PASS
Run_Summary     PASS

CONTENT PARITY  PASS
```

Angka tersebut adalah **regression evidence untuk source population yang comparable**, bukan universal acceptance threshold untuk future months.

## 15. Minimum Post-Run Verification

Setelah controlled monthly run:

1. Review terminal summary.
2. Confirm actual source root.
3. Confirm actual Master path.
4. Confirm `Sampling-Point` worksheet.
5. Confirm detected week count.
6. Confirm workbook population.
7. Review `Execution failures`.
8. Review `Output read errors`.
9. Confirm final workbook exists.
10. Confirm `Data_Mining`, `Control`, `Processing_Log`, `Run_Summary`.
11. Review every worksheet `ERROR`.
12. Jangan menganggap `SUCCESS + 0` sebagai defect tanpa source evidence.
13. Verify representative source-to-output provenance.
14. Preserve output/runtime evidence sesuai applicable control.

Technical check:

```powershell
python -c "from openpyxl import load_workbook; from pathlib import Path; p=Path(r'<OUTPUT_XLSX>'); print('EXISTS =', p.exists()); w=load_workbook(p, read_only=True, data_only=True); print('SHEETS =', w.sheetnames); print('DATA_RECORDS =', w['Data_Mining'].max_row - 1); w.close()"
```

## 16. Troubleshooting Quick Guide

```text
Command fails
   |
   +-- Python / .venv?
   |      → verify interpreter and requirements
   |
   +-- source root?
   |      → Test-Path / permissions / actual path
   |
   +-- no week folder?
   |      → inspect actual source structure
   |
   +-- missing domain folder?
   |      → treat as run evidence
   |
   +-- Master missing?
   |      → verify domains/qc/doc/Master-Data.xlsx
   |
   +-- Sampling-Point missing?
   |      → inspect canonical Master structure
   |
   +-- worksheet EXCLUDED?
   |      → inspect visibility / configured exclusion
   |
   +-- worksheet ERROR?
   |      → inspect source + Processing_Log
   |
   +-- SUCCESS + 0?
   |      → inspect actual source
   |
   +-- unexpected Sampling Point?
   |      → compare source identity + applicable Master entry
   |
   +-- EXEC_FAILED?
   |      → inspect Run_Summary + exact execution message
   |
   +-- OUTPUT_READ_ERROR?
          → inspect temporary/output workbook + exact error
```

## 17. Important Data Handling Rules

Do **not**:

- modify original source workbook agar sesuai parser;
- infer missing date dari filename/folder/worksheet/file timestamp;
- silently route Utility NPG sebagai Utility;
- unhide worksheet lalu otomatis menganggapnya eligible;
- bypass configured exclusions tanpa evidence;
- invent missing/ambiguous value;
- fuzzy-match Sampling Point tanpa controlled rule;
- silently merge similar Sampling Point;
- silently resolve Master collision;
- silently rewrite source spelling;
- promote Parameter menjadi Sampling Point tanpa evidence;
- mengubah `<0.01`, `-`, `OFF`, atau text lain menjadi business meaning yang belum disetujui;
- memperlakukan SUCCESS sebagai QC approval;
- memperlakukan Master match sebagai QC validation;
- memperlakukan candidate output sebagai SoR;
- invent KPI/specification/acceptance threshold;
- infer Production-QC causality;
- menggunakan QC output otomatis sebagai Asset Readiness input;
- menghapus ERROR hanya untuk membuat hasil terlihat clean.

## 18. Git / Recovery Quick Reference

Current accepted repository state:

```text
master          @ 64e394d
origin/master   @ 64e394d
working tree    clean
```

Relevant refactor checkpoints:

```text
0523cd5  docs(qc): add mining refactor concept and gate workflow
c2b94d4  refactor(qc): add shared workbook processor
d04de6e  refactor(qc): route main through shared workbook processor
ed026c6  refactor(qc): modularize monthly mining flow
8816899  fix(qc): preserve monthly progress reporting
64e394d  docs(qc): update user guide for mining flow refactor
```

Useful commands:

```powershell
git status
git log --oneline --decorate
git remote -v
git tag --list

git show <commit>
git diff <commit-a>..<commit-b>
```

Gunakan Git sebagai recovery mechanism. Jangan merekonstruksi implementation dari memory jika actual source/Git evidence dapat diperiksa.

## 19. Security / Repository Data Separation

Git repository ditujukan untuk controlled code/reference artifacts.

Runtime/local locations seperti:

```text
.venv/
output/
scratch/
__pycache__/
```

tidak dimaksudkan sebagai implementation source.

Jangan commit credentials, password, GitHub PAT/token, network credentials, raw operational QC source kecuali separately authorized, atau generated monthly output kecuali separately controlled/approved.

## 20. Governance Boundary

QC Data Mining Tool tidak menetapkan sendiri authoritative QC source, formal QC System of Record, complete QC parameter catalogue, KPI, specification limits, release/acceptance criteria, sample/batch/lot model, refresh rule, reconciliation threshold, aggregation/statistical method, alert/deviation rule, retention, owner/approval authority, Production-QC detailed join/causality, atau QC-to-Asset-Readiness rule.

Item tersebut tetap **TBD / NOT ESTABLISHED** kecuali controlled source menetapkannya.

## 21. Reference Documents

Gunakan applicable/current documents:

```text
00_NEXARIONT_Project_Master_Context_and_Document_Governance_Rev0.5.md

04_NEXARIONT_Quality_Control_Information_Integration_Concept_Rev0.2.docx

domains/qc/doc/
    NEXARIONT_QC_Data_Mining_User_Guide_Rev0.5.docx
    NEXARIONT_QC_Mining_Refactor_Concept_Workflow_Gates_Rev0.2.md
    qc_multi_worksheet_workflow_reference_Rev02.md
    Master-Data.xlsx
```

Untuk CURRENT implementation behavior, actual source + config + runtime evidence + Git evidence mengendalikan CURRENT technical behavior.

## 22. Quick Operating Summary

### Single Workbook

```powershell
python .\domains\qc\main.py `
  --input "<SOURCE_WORKBOOK>" `
  --output "<OUTPUT_XLSX>"
```

atau controlled master/domain:

```powershell
python .\domains\qc\main.py `
  --input "<SOURCE_WORKBOOK>" `
  --output "<OUTPUT_XLSX>" `
  --master "<MASTER_XLSX>" `
  --domain "<DOMAIN>"
```

### Monthly

```powershell
python .\domains\qc\run_monthly_qc_mining.py `
  --input-root "<MONTH_SOURCE_ROOT>" `
  --output-dir "<OUTPUT_DIR>"
```

Optional:

```text
--master
--output-name
--overwrite
```

### Verify

```powershell
python .\domains\qc\main.py --help
python .\domains\qc\run_monthly_qc_mining.py --help

git status
```

---

**End of README — NEXARIONT QC Data Mining**
