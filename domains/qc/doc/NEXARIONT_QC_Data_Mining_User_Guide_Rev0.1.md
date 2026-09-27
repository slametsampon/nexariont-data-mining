# NEXARIONT QC Data Mining Tool --- User Guide

**Document Type:** User Guide (UG)\
**Revision:** Rev0.1\
**Reference Implementation:** `feature/multi-worksheet` @ `e874eef`\
**Reference Workflow:**
`qc_multi_worksheet_workflow_reference_Rev02.md`\
**Scope:** Baseline sampai verified Multi-Worksheet Capability

------------------------------------------------------------------------

## 1. Purpose

User Guide ini menjelaskan penggunaan dan interpretasi **NEXARIONT QC
Data Mining Tool** yang telah dikembangkan dari baseline
single-worksheet sampai multi-worksheet execution path.

Tool membaca workbook QC `.xls` / `.xlsx`, menemukan worksheet,
menerapkan visibility filtering, memproses worksheet eligible,
mempertahankan source worksheet sebagai provenance, dan menghasilkan
structured QC candidate data beserta processing trace.

Output tool **bukan otomatis** validated QC fact, QC approval, product
release decision, atau System of Record.

------------------------------------------------------------------------

## 2. Goal

Goal tool adalah mengambil data QC dari workbook legacy/operasional
secara:

-   controlled;
-   repeatable;
-   traceable;
-   tanpa mengubah source workbook;
-   tanpa merekonstruksi fakta yang tidak tersedia;
-   dengan source worksheet provenance;
-   dengan execution trace per worksheet.

Authority/source/specification/validation/downstream integration yang
belum ditetapkan tetap **TBD**.

------------------------------------------------------------------------

## 3. Verified Implementation State

``` text
Branch : feature/multi-worksheet
HEAD   : e874eef
Commit : Integrate multi-worksheet main execution path
```

Development state:

``` text
Phase A — Core Parser Baseline                 COMPLETE
Phase B — Multi-Worksheet Capability           COMPLETE

CP1    Workbook Worksheet Discovery            COMPLETE
CP2    Visibility Filter                       COMPLETE
CP3    Multi-Worksheet Diagnostic Processing   COMPLETE
CP4.1  Worksheet Provenance Model              COMPLETE
CP4.2  Service Orchestration                   COMPLETE
CP4.3  Multi-Worksheet Excel Export            COMPLETE
CP4.4  Processing Log                          COMPLETE
CP4.5  Main Execution Path Integration         COMPLETE
```

Completion workflow selanjutnya:

``` text
Phase C — Multi-File Regression & Exception Discovery
Phase D — Acceptance & Stabilization
Phase E — Release Baseline
CLOSED
```

------------------------------------------------------------------------

## 4. Execution Architecture

``` text
QC Source Workbook (.xls / .xlsx)
        |
        v
main.py
        |
        v
WorkbookReader
        |
        |-- worksheet discovery
        |-- visibility metadata
        |
        v
QCDataMiningService.process_all_worksheets()
        |
        |-- exclude non-visible worksheet
        |-- parse visible worksheet
        |-- isolate per-worksheet error
        |-- retain source worksheet
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

------------------------------------------------------------------------

## 5. Prerequisites

Environment yang digunakan selama development:

``` text
Python virtual environment : .venv
Command shell              : PowerShell
Repository                 : qc-data-mining
```

Implementation telah menggunakan:

-   `xlrd` untuk `.xls`;
-   `openpyxl` untuk `.xlsx` dan output Excel.

Aktifkan environment sesuai konfigurasi workstation, contoh:

``` powershell
& .\.venv\Scripts\Activate.ps1
```

Supported source format yang telah established:

``` text
.xls
.xlsx
```

------------------------------------------------------------------------

## 6. Running the Tool

Command utama:

``` powershell
python main.py `
  --input "<FULL_PATH_SOURCE_WORKBOOK>" `
  --output "<OUTPUT_XLSX_PATH>"
```

Verified example:

``` powershell
python main.py `
  --input "C:\Users\sam294\Documents\DATA\Staff-ahli\nexariont\data\qca\2026\cf_11982\1790115243\3-Selasa_oct.xls" `
  --output ".\output\3-Selasa_oct_mining.xlsx"
```

Verified console result untuk reference workbook:

``` text
QC DATA MINING COMPLETED
--------------------------------------------------
Input     : <source path>
Output    : <output path>
Records   : 232
Worksheets: 9
Success   : 8
Excluded  : 1
Error     : 0
--------------------------------------------------
```

Angka tersebut hanya verified result untuk `3-Selasa_oct.xls`, bukan
universal acceptance threshold.

------------------------------------------------------------------------

## 7. Worksheet Discovery and Visibility

Reference workbook mempunyai:

``` text
shift-pagi
shift-pagi(2)
AWR - Shift-Pagi
shift-sore
shift-sore(2)
AWR - Shift-Sore
shift-malam
AWR - Shift-malam
dbase
```

Implementation visibility representation:

``` text
0 = visible
1 = hidden
2 = very hidden
```

Rule penting:

``` text
Worksheet discovered
        !=
Worksheet automatically parsed
```

Hanya worksheet visible yang menjadi candidate parsing pada
implementation yang telah diverifikasi.

Reference evidence:

``` text
dbase visibility = 1
dbase -> HIDDEN -> EXCLUDED -> 0 records
```

Ini adalah filtering behavior, bukan penilaian QC validity terhadap isi
`dbase`.

------------------------------------------------------------------------

## 8. Processing Status

### SUCCESS

Parser selesai tanpa exception.

``` text
shift-pagi  VISIBLE  SUCCESS  74
```

### SUCCESS + 0 Records

``` text
AWR - Shift-Pagi  VISIBLE  SUCCESS  0
```

Artinya parser selesai tanpa exception tetapi tidak menghasilkan
candidate record.

Jangan otomatis menyimpulkan worksheet invalid atau software defect.

### EXCLUDED

Worksheet tidak diproses karena non-visible.

``` text
dbase  HIDDEN  EXCLUDED  0  Non-visible worksheet
```

### ERROR

Worksheet menghasilkan exception saat processing.

Jika `ERROR` muncul, perlakukan sebagai finding untuk investigation.
Jangan langsung mengubah source atau parser sebelum penyebab
terverifikasi.

------------------------------------------------------------------------

## 9. Output Workbook

Verified multi-worksheet output:

``` text
Data_Mining
Control
Processing_Log
```

### 9.1 Data_Mining

Header:

``` text
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

`Source Worksheet` mempertahankan provenance record.

Record harus dipahami sebagai:

``` text
candidate normalized QC record
```

bukan otomatis:

``` text
validated QC fact
approved QC result
product release decision
formal System-of-Record record
```

### 9.2 Control

`Control` telah terverifikasi hadir pada output multi-worksheet.

UG ini tidak menambahkan business authority, acceptance rule, atau
validation rule untuk isi `Control` yang belum established.

### 9.3 Processing_Log

Header:

``` text
Worksheet
Visibility
Status
Record Count
Message
```

`Processing_Log` adalah execution trace, bukan QC approval record.

Verified reference:

``` text
shift-pagi             VISIBLE  SUCCESS   74
shift-pagi(2)          VISIBLE  SUCCESS    8
AWR - Shift-Pagi       VISIBLE  SUCCESS    0
shift-sore             VISIBLE  SUCCESS   71
shift-sore(2)          VISIBLE  SUCCESS    0
AWR - Shift-Sore       VISIBLE  SUCCESS    8
shift-malam            VISIBLE  SUCCESS   71
AWR - Shift-malam      VISIBLE  SUCCESS    0
dbase                  HIDDEN   EXCLUDED   0  Non-visible worksheet
```

Candidate total:

``` text
74 + 8 + 0 + 71 + 0 + 8 + 71 + 0 = 232
```

------------------------------------------------------------------------

## 10. Minimum Post-Run Verification

Setelah mining:

1.  Periksa console `Records`, `Worksheets`, `Success`, `Excluded`, dan
    `Error`.
2.  Buka output workbook.
3.  Pastikan `Data_Mining`, `Control`, dan `Processing_Log` tersedia.
4.  Review seluruh `Processing_Log`.
5.  Investigate setiap `ERROR`.
6.  Jangan treat `SUCCESS + 0` sebagai defect tanpa source evidence.
7.  Pastikan `Source Worksheet` tersedia pada `Data_Mining`.
8.  Untuk regression, lakukan sample source-to-output reconciliation.

Optional technical check:

``` powershell
python -c "from openpyxl import load_workbook; p=r'.\output\3-Selasa_oct_mining.xlsx'; wb=load_workbook(p,data_only=True); print('SHEETS=',wb.sheetnames); dm=wb['Data_Mining']; pl=wb['Processing_Log']; print('DATA_MINING RECORDS=',dm.max_row-1); print('PROCESSING_LOG ROWS=',pl.max_row-1); print('PROCESSING_LOG=',[[c.value for c in row] for row in pl.iter_rows(min_row=2)]); wb.close()"
```

Verified reference result:

``` text
SHEETS               = ['Data_Mining', 'Control', 'Processing_Log']
DATA_MINING RECORDS  = 232
PROCESSING_LOG ROWS  = 9
```

------------------------------------------------------------------------

## 11. Source Value Handling

Reader melakukan conversion secara konservatif.

Established behavior:

``` text
Blank   -> None
Text    -> retained as text
Numeric -> numeric
Boolean -> boolean
```

Text seperti:

``` text
<0.01
-
OFF
```

dipertahankan sebagai text. Jangan mengubahnya menjadi angka/business
state baru tanpa controlled rule.

------------------------------------------------------------------------

## 12. Error Handling

### Input file tidak ditemukan

Periksa path, filename, extension, dan availability.

### Unsupported format

Established input format hanya:

``` text
.xls
.xlsx
```

Format lain diperlakukan sebagai unsupported sampai capability
ditetapkan.

### Worksheet ERROR

Catat:

``` text
source file
worksheet
status
message
```

Kemudian inspeksi actual source/layout dan CURRENT implementation.

### SUCCESS tetapi 0 records

Inspect source terlebih dahulu. Jangan mengubah parser berdasarkan angka
`0` saja.

### Unexpected extracted record

Bandingkan source dengan output. Pertahankan evidence dan provenance
sebelum menentukan defect.

------------------------------------------------------------------------

## 13. Prohibited / Unsafe Data Handling

Untuk menjaga traceability:

-   jangan mengubah original workbook agar sesuai parser;
-   jangan unhide worksheet lalu menganggapnya otomatis eligible;
-   jangan memasukkan hidden worksheet secara manual tanpa established
    rule;
-   jangan mengisi missing/ambiguous value berdasarkan tebakan;
-   jangan silently cleanse, merge, discard, atau reconstruct record;
-   jangan menghapus `Source Worksheet` provenance;
-   jangan menganggap processing `SUCCESS` sebagai QC approval;
-   jangan menetapkan QC specification/limit/acceptance threshold yang
    belum established.

------------------------------------------------------------------------

## 14. Traceability

``` text
Source Workbook
      |
      v
Worksheet
      |
      v
Source Worksheet provenance
      |
      v
Candidate SamplingRecord
      |
      +--> Data_Mining
      |
      +--> Processing_Log
```

Traceability tersebut membantu verification/investigation tetapi tidak
menetapkan data authority.

------------------------------------------------------------------------

## 15. Phase C --- Multi-File Regression & Exception Discovery

Setelah verified implementation `e874eef`, next workflow phase adalah
Phase C.

Purpose:

> memastikan behavior yang terbukti pada `3-Selasa_oct.xls` tidak
> bergantung hanya pada satu workbook.

Workflow:

``` text
CURRENT implementation
        |
        v
Select actual QC source workbook
        |
        v
Run main.py
        |
        v
Capture execution evidence
        |
        v
Inspect Data_Mining / Processing_Log
        |
        v
Source-to-output reconciliation
        |
        v
Classify finding
        |
        +-- expected behavior
        +-- source/layout exception
        +-- suspected software defect
        |
        v
Regression evidence / exception inventory
```

Minimum evidence per workbook:

``` text
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

Acceptance threshold untuk complete regression population masih **TBD**
sampai explicitly established.

------------------------------------------------------------------------

## 16. Jika Regression Menemukan Defect

``` text
Regression finding
      |
      v
inspect actual source + CURRENT code
      |
      v
confirm defect
      |
      v
minimum scoped correction
      |
      v
final replacement file
      |
      v
syntax verification
      |
      v
affected regression re-test
      |
      v
reconciliation evidence
      |
      v
Git recovery checkpoint
```

Jangan membuat modification hanya untuk menyeragamkan source workbook
yang memang berbeda.

------------------------------------------------------------------------

## 17. Phase D --- Acceptance & Stabilization

Phase D dilakukan setelah regression evidence tersedia.

Scope:

-   reconcile regression evidence;
-   retain expected behavior;
-   resolve evidence-based confirmed defects;
-   re-test affected cases;
-   establish acceptance evidence.

Formal acceptance criteria, required regression population,
approver/authority, dan release approval tetap **TBD unless established
separately**.

------------------------------------------------------------------------

## 18. Phase E --- Release Baseline

Expected release workflow:

``` text
Accepted feature branch
        |
        v
verify clean working tree
        |
        v
merge feature/multi-worksheet -> master
        |
        v
post-merge verification
        |
        v
version tag / release baseline
        |
        v
CLOSED
```

Exact version/tag masih TBD.

------------------------------------------------------------------------

## 19. Developer / Maintainer Recovery Reference

Verified history:

``` text
e874eef  Integrate multi-worksheet main execution path
375e9a3  Add multi-worksheet processing log
69deea8  Add multi-worksheet Excel export
75261a9  Add multi-worksheet service orchestration
3e7150d  Add worksheet provenance model
5fc2936  Add multi-worksheet diagnostic processing
1f7c992  Add worksheet visibility filtering
0486509  Add workbook worksheet discovery
5afe138  Baseline: direct XLS shift-pagi parser verified 74 records
```

Important recovery points:

``` text
5afe138  verified single-worksheet baseline
e874eef  verified integrated multi-worksheet implementation
```

Check repository state:

``` powershell
git status
git log --oneline --decorate
```

------------------------------------------------------------------------

## 20. Reference Verified Case

``` text
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

``` text
Data_Mining     : 232 candidate records
Control         : present
Processing_Log  : 9 worksheet results
```

Nilai ini adalah reference regression evidence, bukan universal business
acceptance threshold.

------------------------------------------------------------------------

## 21. Operational Quick Checklist

### Before

-   source workbook tersedia;
-   format `.xls` / `.xlsx`;
-   output path ditentukan;
-   Python environment aktif;
-   original source tidak dimodifikasi untuk menyesuaikan parser.

### Run

``` text
python main.py --input <source> --output <output>
```

### After

-   review console summary;
-   review `Processing_Log`;
-   inspect `ERROR`;
-   jangan treat `SUCCESS + 0` sebagai defect tanpa evidence;
-   verify `Source Worksheet`;
-   reconcile source/output sesuai verification need.

------------------------------------------------------------------------

## 22. Troubleshooting Decision Guide

``` text
Command gagal
   |
   +-- file/path problem?
   |      -> verify actual path/file
   |
   +-- unsupported format?
   |      -> record unsupported input finding
   |
   +-- worksheet ERROR?
   |      -> inspect source layout + message
   |
   +-- SUCCESS but 0?
   |      -> inspect source; do not assume defect
   |
   +-- unexpected record?
          -> compare source vs output
          -> preserve provenance
          -> classify before code change
```

------------------------------------------------------------------------

## 23. Governance

``` text
CURRENT implementation fact
    -> actual source/config/code/schema/runtime/Git evidence

Workflow
    -> qc_multi_worksheet_workflow_reference_Rev02.md

Unknown
    -> TBD / belum terverifikasi
```

Jika documentation berbeda dengan actual runtime/code behavior, jangan
silently reconcile. Catat discrepancy dan verifikasi CURRENT
implementation.

------------------------------------------------------------------------

## 24. Definition of Completion

Finite workflow:

``` text
Implementation
    -> Multi-File Regression
    -> Acceptance & Stabilization
    -> Release Baseline
    -> CLOSED
```

Tidak ada automatic CP5, CP6, CP7, dan seterusnya tanpa scoped need yang
explicitly established.

Capability baru setelah closure adalah scoped work item/feature baru.

------------------------------------------------------------------------

# Appendix A --- Development Evolution

``` text
Single-Worksheet Baseline
        |
        v
Worksheet Discovery
        |
        v
Visibility Filter
        |
        v
Diagnostic Multi-Worksheet
        |
        v
Worksheet Provenance
        |
        v
Service Orchestration
        |
        v
Consolidated Excel Export
        |
        v
Processing_Log
        |
        v
main.py Integration
        |
        v
Verified Multi-Worksheet Implementation
```

------------------------------------------------------------------------

# Appendix B --- Verified Processing Detail

``` text
Worksheet              Visibility  Status    Records
----------------------------------------------------
shift-pagi             VISIBLE     SUCCESS       74
shift-pagi(2)          VISIBLE     SUCCESS        8
AWR - Shift-Pagi       VISIBLE     SUCCESS        0
shift-sore             VISIBLE     SUCCESS       71
shift-sore(2)          VISIBLE     SUCCESS        0
AWR - Shift-Sore       VISIBLE     SUCCESS        8
shift-malam            VISIBLE     SUCCESS       71
AWR - Shift-malam      VISIBLE     SUCCESS        0
dbase                  HIDDEN      EXCLUDED       0
----------------------------------------------------
Candidate Records                              232
```

------------------------------------------------------------------------

# Appendix C --- Document Basis

UG ini disusun dari:

1.  implementation/verification evidence selama development dari
    baseline sampai Git checkpoint `e874eef`; dan
2.  `qc_multi_worksheet_workflow_reference_Rev02.md`.

Jika implementation berubah setelah `e874eef`, UG harus direview
terhadap CURRENT implementation sebelum diperlakukan sebagai description
of current behavior.

------------------------------------------------------------------------

**End of User Guide**
