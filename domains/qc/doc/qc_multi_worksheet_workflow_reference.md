# QC Data Mining --- Multi-Worksheet Implementation Workflow

**Project:** NEXARIONT QC Data Mining\
**Branch:** `feature/multi-worksheet`\
**Purpose:** Referensi workflow pengembangan multi-worksheet dengan
checkpoint, verification, dan recovery menggunakan Git.

> Dokumen ini merekonstruksi workflow yang telah dijalankan dalam sesi
> pengembangan, termasuk **Checkpoint 2 --- Visibility Filter**, dan
> diperbarui dengan checkpoint Git yang sudah terverifikasi sampai
> integrasi `main.py`.

------------------------------------------------------------------------

## 1. Prinsip Kerja

Workflow pengembangan menggunakan prinsip berikut:

1.  Source `.xls` / `.xlsx` dibaca tanpa mengubah source workbook.
2.  Perubahan dilakukan pada branch `feature/multi-worksheet`.
3.  Setiap capability utama diverifikasi sebelum commit.
4.  Git commit digunakan sebagai recovery checkpoint.
5.  Existing behavior dipertahankan sampai replacement path sudah
    terverifikasi.
6.  Hidden/non-visible worksheet tidak otomatis menjadi candidate
    parsing.
7.  Source worksheet dipertahankan sebagai provenance/traceability.
8.  Candidate parsing record bukan pernyataan QC validation atau
    approval.
9.  Tidak melakukan inference terhadap isi worksheet yang tidak
    menghasilkan record.
10. Perubahan berikutnya dimulai dari working tree yang clean.

------------------------------------------------------------------------

## 2. Baseline / Recovery Point Awal

Baseline single-worksheet telah diverifikasi dengan:

``` text
Worksheet : shift-pagi
Records   : 74
```

Git baseline:

``` text
5afe138  Baseline: direct XLS shift-pagi parser verified 74 records
tag      v0.1-baseline
```

Branch pengembangan:

``` bash
git switch -c feature/multi-worksheet
```

Baseline harus tetap dapat dipulihkan melalui Git.

------------------------------------------------------------------------

# CHECKPOINT 1 --- Workbook Worksheet Discovery

## Tujuan

Menambahkan capability untuk menemukan seluruh worksheet dalam workbook
tanpa mengubah parser existing.

## Scope

`src/workbook_reader.py`

Capability:

``` python
WorkbookReader.get_worksheet_names()
```

Supported:

-   `.xls` melalui `xlrd`
-   `.xlsx` melalui `openpyxl`

## Verification

Untuk source:

``` text
3-Selasa_oct.xls
```

hasil discovery:

``` text
[
  'shift-pagi',
  'shift-pagi(2)',
  'AWR - Shift-Pagi',
  'shift-sore',
  'shift-sore(2)',
  'AWR - Shift-Sore',
  'shift-malam',
  'AWR - Shift-malam',
  'dbase'
]
```

## Git Checkpoint

``` text
0486509  Add workbook worksheet discovery
```

------------------------------------------------------------------------

# CHECKPOINT 2 --- Visibility Filter

## Tujuan

Membedakan worksheet yang terlihat oleh user dari worksheet yang
disimpan sebagai hidden/non-visible dalam workbook.

Discovery worksheet **tidak sama** dengan keputusan bahwa semua
worksheet harus diparse.

## Evidence Source Workbook

Verification menggunakan `xlrd` menunjukkan:

``` text
0 'shift-pagi'        visibility=0
1 'shift-pagi(2)'     visibility=0
2 'AWR - Shift-Pagi'  visibility=0
3 'shift-sore'        visibility=0
4 'shift-sore(2)'     visibility=0
5 'AWR - Shift-Sore'  visibility=0
6 'shift-malam'       visibility=0
7 'AWR - Shift-malam' visibility=0
8 'dbase'             visibility=1
```

Interpretasi implementation:

``` text
0 = visible
1 = hidden
2 = very hidden
```

Untuk source yang diuji:

``` text
dbase = hidden
```

Karena itu `dbase` **tidak menjadi candidate parsing**.

## Data Structure

Metadata worksheet direpresentasikan melalui:

``` python
@dataclass(frozen=True)
class WorksheetInfo:
    name: str
    visibility: int

    @property
    def is_visible(self) -> bool:
        return self.visibility == 0
```

## Capability

`WorkbookReader` menyediakan:

``` python
get_worksheet_info()
get_visible_worksheet_names()
```

## Verified Visible Worksheets

``` text
[
  'shift-pagi',
  'shift-pagi(2)',
  'AWR - Shift-Pagi',
  'shift-sore',
  'shift-sore(2)',
  'AWR - Shift-Sore',
  'shift-malam',
  'AWR - Shift-malam'
]
```

`dbase` tidak termasuk.

## Important Rule

``` text
Worksheet discovered
        ≠
Worksheet automatically parsed
```

dan:

``` text
Non-visible worksheet
        ↓
EXCLUDED from candidate parsing
```

Visibility filter hanya menentukan candidate worksheet berdasarkan
visibility metadata. Filter ini tidak menyatakan validitas atau approval
QC.

## Regression Verification

Existing single-worksheet execution tetap menghasilkan:

``` text
Worksheet : shift-pagi
Records   : 74
```

## Git Checkpoint

``` text
1f7c992  Add worksheet visibility filtering
```

------------------------------------------------------------------------

# CHECKPOINT 3 --- Multi-Worksheet Diagnostic Processing

## Tujuan

Menjalankan existing parser terhadap setiap visible worksheet dalam mode
diagnostic sebelum mengubah production execution path.

Hidden worksheet dicatat sebagai excluded.

## Verified Diagnostic Result

``` text
shift-pagi             VISIBLE   SUCCESS   74
shift-pagi(2)          VISIBLE   SUCCESS    8
AWR - Shift-Pagi       VISIBLE   SUCCESS    0
shift-sore             VISIBLE   SUCCESS   71
shift-sore(2)          VISIBLE   SUCCESS    0
AWR - Shift-Sore       VISIBLE   SUCCESS    8
shift-malam            VISIBLE   SUCCESS   71
AWR - Shift-malam      VISIBLE   SUCCESS    0
dbase                  HIDDEN    EXCLUDED   0
```

Summary:

``` text
Worksheets discovered : 9
Visible candidates    : 8
Excluded              : 1
Success               : 8
Error                 : 0
Candidate records     : 232
```

Important:

``` text
SUCCESS + 0 records
```

berarti parser selesai tanpa exception tetapi tidak menghasilkan
candidate record. Itu tidak otomatis berarti worksheet invalid atau
tidak berguna.

## Git Checkpoint

``` text
5fc2936  Add multi-worksheet diagnostic processing
```

------------------------------------------------------------------------

# CHECKPOINT 4 --- Multi-Worksheet Production Capability

Checkpoint 4 dikembangkan bertahap untuk menjaga recovery, tetapi
keseluruhan scope-nya adalah membangun production-capable
multi-worksheet pipeline.

------------------------------------------------------------------------

## CP4.1 --- Worksheet Provenance Model

### Tujuan

Mempertahankan asal worksheet setiap normalized sampling record.

Model:

``` python
SourcedSamplingRecord(
    source_worksheet=...,
    record=SamplingRecord(...)
)
```

`SamplingRecord` existing tidak diubah menjadi model source-specific.

Output row multi-worksheet menambahkan:

``` text
Source Worksheet
```

sebagai provenance.

### Git Checkpoint

``` text
3e7150d  Add worksheet provenance model
```

------------------------------------------------------------------------

## CP4.2 --- Multi-Worksheet Service Orchestration

### Tujuan

Menambahkan orchestration untuk:

-   discovery seluruh worksheet;
-   exclude non-visible worksheet;
-   parsing visible worksheet;
-   error isolation per worksheet;
-   aggregate candidate records;
-   mempertahankan source worksheet.

Capability:

``` python
QCDataMiningService.process_all_worksheets()
```

Verified:

``` text
TOTAL    = 232
SUCCESS  = 8
ERROR    = 0
EXCLUDED = 1
```

Record provenance distribution:

``` text
shift-pagi       74
shift-pagi(2)     8
shift-sore       71
AWR - Shift-Sore  8
shift-malam      71
```

Total:

``` text
232
```

### Git Checkpoint

``` text
75261a9  Add multi-worksheet service orchestration
```

------------------------------------------------------------------------

## CP4.3 --- Multi-Worksheet Excel Export

### Tujuan

Mengekspor consolidated candidate records ke workbook output sambil
mempertahankan provenance.

Verified output:

``` text
Sheets:
- Data_Mining
- Control
```

`Data_Mining`:

``` text
Rows = 233
```

terdiri dari:

``` text
1 header + 232 candidate records
```

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

### Git Checkpoint

``` text
69deea8  Add multi-worksheet Excel export
```

------------------------------------------------------------------------

## CP4.4 --- Processing Log

### Tujuan

Menambahkan execution trace per worksheet ke output workbook.

Output:

``` text
Processing_Log
```

Header:

``` text
Worksheet
Visibility
Status
Record Count
Message
```

Verified result:

``` text
shift-pagi             VISIBLE SUCCESS  74
shift-pagi(2)          VISIBLE SUCCESS   8
AWR - Shift-Pagi       VISIBLE SUCCESS   0
shift-sore             VISIBLE SUCCESS  71
shift-sore(2)          VISIBLE SUCCESS   0
AWR - Shift-Sore       VISIBLE SUCCESS   8
shift-malam            VISIBLE SUCCESS  71
AWR - Shift-malam      VISIBLE SUCCESS   0
dbase                  HIDDEN  EXCLUDED  0  Non-visible worksheet
```

Verified workbook:

``` text
SHEETS               = ['Data_Mining', 'Control', 'Processing_Log']
DATA_MINING RECORDS  = 232
PROCESSING_LOG ROWS  = 9
```

Existing single-worksheet regression setelah implementasi CP4.4:

``` text
Worksheet : shift-pagi
Records   : 74
```

### Git Checkpoint

``` text
375e9a3  Add multi-worksheet processing log
```

------------------------------------------------------------------------

## CP4.5 --- Main Execution Path Integration

### Tujuan

Mengaktifkan multi-worksheet processing melalui normal CLI execution:

``` bash
python main.py --input ... --output ...
```

Existing single-worksheet `process_file()` tidak lagi menjadi main
execution path.

Main path:

``` text
main.py
   ↓
process_all_worksheets()
   ↓
MultiWorksheetProcessingResult
   ↓
export_multi_worksheet()
   ↓
Data_Mining + Control + Processing_Log
```

Verified execution:

``` text
QC DATA MINING COMPLETED
--------------------------------------------------
Records   : 232
Worksheets: 9
Success   : 8
Excluded  : 1
Error     : 0
--------------------------------------------------
```

Verified workbook:

``` text
SHEETS               = ['Data_Mining', 'Control', 'Processing_Log']
DATA_MINING RECORDS  = 232
PROCESSING_LOG ROWS  = 9
```

`dbase` tetap:

``` text
HIDDEN / EXCLUDED / 0 records
```

### Git Checkpoint

``` text
e874eef  Integrate multi-worksheet main execution path
```

------------------------------------------------------------------------

# 5. Current Verified Git History

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

Branch:

``` text
feature/multi-worksheet
```

Current recovery point:

``` text
e874eef
```

------------------------------------------------------------------------

# 6. Current Functional State

Untuk test source `3-Selasa_oct.xls`, implementation telah
terverifikasi:

``` text
Workbook worksheets     : 9
Visible candidates      : 8
Hidden/excluded         : 1
Parser errors           : 0
Candidate records       : 232
```

Output:

``` text
Data_Mining
Control
Processing_Log
```

Source traceability dipertahankan melalui:

``` text
Source Worksheet
```

Hidden worksheet:

``` text
dbase
```

tidak diparse dan tetap tercatat dalam `Processing_Log`.

------------------------------------------------------------------------

# 7. Recovery Mechanism

## Lihat state

``` bash
git status
git log --oneline --decorate
```

## Kembali melihat baseline commit tanpa mengubah branch

``` bash
git show <commit>
```

## Membandingkan perubahan

``` bash
git diff <commit-a>..<commit-b>
```

## Recovery terhadap uncommitted modification

Gunakan hanya setelah memastikan perubahan memang ingin dibuang:

``` bash
git restore <file>
```

## Recovery point penting

``` text
5afe138  original verified single-worksheet baseline
0486509  worksheet discovery
1f7c992  visibility filter
5fc2936  diagnostic multi-worksheet
3e7150d  provenance
75261a9  service orchestration
69deea8  Excel export
375e9a3  Processing_Log
e874eef  integrated main execution path
```

------------------------------------------------------------------------

# 8. Acceptance Principle

Implementation multi-worksheet untuk source yang diuji dianggap
functionally verified berdasarkan evidence:

``` text
232 candidate records
9 worksheet results
8 visible/success
1 hidden/excluded
0 parser errors
```

Namun hasil parsing tetap harus dipahami sebagai:

``` text
candidate normalized QC records
```

bukan otomatis:

``` text
validated QC fact
approved QC result
release decision
formal system-of-record record
```

Multi-file regression terhadap source workbook lain merupakan
verification berikutnya yang terpisah dari capability yang telah
terverifikasi pada `3-Selasa_oct.xls`.

------------------------------------------------------------------------

# 9. Development Discipline untuk Perubahan Berikutnya

Untuk menghindari perubahan yang sulit dipulihkan:

``` text
CURRENT source
      ↓
review actual implementation
      ↓
minimum scoped change
      ↓
final file
      ↓
syntax verification
      ↓
functional verification
      ↓
Git commit / recovery checkpoint
```

Jangan melakukan substantive reconstruction berdasarkan asumsi ketika
CURRENT implementation dapat diperiksa langsung.

------------------------------------------------------------------------

**End of Reference**
