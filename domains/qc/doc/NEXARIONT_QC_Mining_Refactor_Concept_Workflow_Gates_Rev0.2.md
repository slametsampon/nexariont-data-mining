# NEXARIONT QC Mining Refactor — Concept, Workflow & Gate Check

**Revision:** Rev0.2  
**Status:** Proposed Refactor Baseline  
**Purpose:** Pegangan/rujukan refactor `main.py` dan `run_monthly_qc_mining.py`  
**Primary Objective:** Melepaskan coupling `run_monthly_qc_mining.py` terhadap `main.py` tanpa mengubah behavior QC mining yang sudah terbukti.

---

## 1. Governing Principles

Refactor harus:

- painless dan reversible;
- mengikuti **Separation of Concerns (SoC)**;
- mengikuti OOP principle dengan **high cohesion** dan **low coupling**;
- menggunakan **composition over inheritance**;
- bersifat **structural refactor, not behavioral redesign**;
- mempertahankan external CLI contract dan output contract;
- tidak mengubah parser/business logic tanpa blocker yang terverifikasi;
- tidak memasukkan cleanup/optimization yang bukan bagian objective.

Dependency direction:

```text
CLI / Entry Point
       ↓
Application Layer
       ↓
Existing QC Core
```

Tidak diperbolehkan:

```text
run_monthly_qc_mining.py
        ↓
      main.py
```

---

## 2. Target Architecture

```text
main.py
   ↓
QCWorkbookProcessor
   ↓
Existing QC Core


run_monthly_qc_mining.py
   ↓
MonthlyMiningOrchestrator
   ├── MonthlySourcePlanner
   ├── QCWorkbookProcessor
   └── MonthlyWorkbookConsolidator
```

### Responsibility

| Component | Primary Responsibility |
|---|---|
| `main.py` | Thin CLI single-workbook |
| `run_monthly_qc_mining.py` | Thin CLI monthly |
| `QCWorkbookProcessor` | Process satu workbook QC end-to-end |
| `MonthlySourcePlanner` | Monthly folder → deterministic work plan |
| `MonthlyMiningOrchestrator` | Coordinate monthly execution |
| `MonthlyWorkbookConsolidator` | Build consolidated monthly workbook |

Existing lower-level/core components tetap digunakan:

```text
MiningConfig
SamplingPointMaster
WorkbookReader
Parser
Models
QCDataMiningService
ExcelExporter
```

---

## 3. OOP / SoC Rules

### Single Responsibility

```text
Planner      → discovery / planning
Processor    → single-workbook processing
Orchestrator → coordination
Consolidator → monthly output
CLI          → user interface / command line
```

### Encapsulation

- filesystem discovery → `MonthlySourcePlanner`
- single-workbook processing → `QCWorkbookProcessor`
- monthly output construction → `MonthlyWorkbookConsolidator`
- monthly execution lifecycle → `MonthlyMiningOrchestrator`
- CLI parsing / console output → respective entry point

### Composition

`MonthlyMiningOrchestrator` menggunakan:

```text
MonthlySourcePlanner
QCWorkbookProcessor
MonthlyWorkbookConsolidator
```

Tidak membuat inheritance hierarchy atau abstract layer baru tanpa kebutuhan nyata.

---

# 4. Revised Workflow

```text
CURRENT BASELINE
      ↓
G0 — Baseline & Recovery
      ↓
G1 — Architecture / SoC Freeze
      ↓
G2 — Contract & Ownership Freeze
      ↓
G3 — Shared Application Structure
      ↓
Refactor main.py
      ↓
G4 — Single-Workbook Functional Parity
      ↓
Refactor Monthly Flow
      ├── MonthlySourcePlanner
      ├── MonthlyMiningOrchestrator
      └── MonthlyWorkbookConsolidator
      ↓
G5 — Monthly Structural Verification
      ↓
G6 — Monthly Functional Parity
      ↓
G7 — Decoupling Verification
      ↓
G8 — Stabilization
      ↓
G9 — Final Acceptance
```

---

# G0 — Baseline & Recovery

## Objective

Menetapkan recovery point sebelum source code berubah.

## Required Actions

1. Verifikasi:
   - `master` clean;
   - `master == origin/master`.
2. Catat baseline commit hash.
3. Buat branch refactor.
4. Commit dokumen refactor pada branch tersebut.
5. Push branch ke remote.
6. Jalankan CURRENT monthly baseline.
7. Simpan output baseline untuk regression comparison.

## Gate PASS

G0 PASS hanya jika seluruh kondisi berikut terpenuhi:

- [ ] baseline commit tercatat;
- [ ] branch refactor aktif;
- [ ] branch remote tersedia;
- [ ] working tree clean;
- [ ] source code belum berubah;
- [ ] CURRENT monthly execution berhasil;
- [ ] baseline output tersedia;
- [ ] baseline totals tercatat.

### Established G0 Regression Baseline

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

## FAIL

Stop. Jangan masuk G1.

---

# G1 — Architecture / SoC Freeze

## Objective

Membekukan architecture boundary sebelum coding.

## Required Architecture

```text
main.py
   └── thin CLI

run_monthly_qc_mining.py
   └── thin CLI

QCWorkbookProcessor
   └── process satu workbook

MonthlySourcePlanner
   └── source structure → work plan

MonthlyMiningOrchestrator
   └── coordinate monthly execution

MonthlyWorkbookConsolidator
   └── monthly output consolidation
```

## Gate PASS

- [ ] satu component mempunyai satu primary responsibility;
- [ ] `main.py` dan monthly runner tidak saling bergantung;
- [ ] CLI tidak melakukan filesystem discovery;
- [ ] CLI tidak melakukan QC processing;
- [ ] CLI tidak melakukan Excel consolidation;
- [ ] Orchestrator hanya mengkoordinasikan;
- [ ] Planner tidak melakukan processing;
- [ ] Consolidator tidak melakukan discovery;
- [ ] existing core tetap lower-level implementation;
- [ ] dependency hanya `Entry → Application → Core`;
- [ ] tidak ada circular dependency;
- [ ] composition digunakan;
- [ ] tidak ada inheritance/abstraction baru tanpa kebutuhan;
- [ ] tidak ada behavioral change pada gate ini.

## FAIL

Stop. Architecture belum boleh diimplementasikan.

---

# G2 — Contract & Ownership Freeze

## Objective

Menetapkan public contract dan ownership antar-component.

## 2.1 `QCWorkbookProcessor`

Conceptual contract:

```text
process(
    input_file,
    output_file,
    master_file=None,
    domain=None
)
    → MultiWorksheetProcessingResult
```

Ownership:

- `MiningConfig`;
- `SamplingPointMaster`;
- `QCDataMiningService`;
- `ExcelExporter`;
- existing per-workbook XLSX;
- existing processing result.

Tidak memiliki:

- CLI parsing;
- monthly discovery;
- monthly consolidation;
- console output;
- process exit code.

## 2.2 `MonthlySourcePlanner`

Input:

```text
input_root
domain routing/configuration CURRENT
```

Output:

```text
MonthlyPlan
 ├── month identity
 ├── detected weeks
 ├── work items
 └── discovery conditions
```

Work item minimal:

```text
month
week
domain
source workbook
source path
```

Discovery conditions mencakup CURRENT semantics seperti:

```text
MISSING_DOMAIN_FOLDER
NO_WORKBOOKS
```

## 2.3 `MonthlyMiningOrchestrator`

Ownership:

- monthly execution lifecycle;
- temporary workspace lifecycle;
- temporary output path;
- workbook execution isolation;
- continue-on-failure;
- `EXEC_FAILED`;
- coordination Planner → Processor → Consolidator.

## 2.4 `MonthlyWorkbookConsolidator`

Ownership:

```text
Data_Mining
Control
Processing_Log
Run_Summary
```

Termasuk:

- membaca temporary per-workbook XLSX;
- output/header compatibility check;
- provenance append;
- totals;
- formatting;
- save final monthly XLSX;
- `OUTPUT_READ_ERROR`.

## 2.5 Master Compatibility Adapter

- tetap dipertahankan;
- berada pada monthly application layer;
- dikoordinasikan oleh Orchestrator;
- tidak dipindahkan ke QC core;
- tidak dihapus pada refactor ini.

## 2.6 Temporary XLSX

Tetap dipertahankan:

```text
QCWorkbookProcessor
        ↓
temporary XLSX
        ↓
MonthlyWorkbookConsolidator
        ↓
final monthly XLSX
```

Ownership:

- lifecycle/path → Orchestrator;
- generation → Processor;
- consumption → Consolidator.

## 2.7 Monthly CLI Contract

Harus tetap:

```text
usage: run_monthly_qc_mining.py [-h]
                                --input-root INPUT_ROOT
                                --output-dir OUTPUT_DIR
                                [--master MASTER]
                                [--output-name OUTPUT_NAME]
                                [--overwrite]
```

Compatibility:

- `--input-root` mandatory;
- `--output-dir` mandatory;
- `--master` optional dengan CURRENT default;
- worksheet canonical tetap `Sampling-Point`;
- `--output-name` optional dengan CURRENT naming;
- `--overwrite` semantics sama;
- `-h/--help` sama;
- local/UNC input support sama;
- input/output boleh berbeda lokasi.

## Gate PASS

- [ ] Processor contract jelas;
- [ ] Planner contract jelas;
- [ ] Orchestrator contract jelas;
- [ ] Consolidator contract jelas;
- [ ] temporary output ownership jelas;
- [ ] master adapter ownership jelas;
- [ ] `EXEC_FAILED` ownership jelas;
- [ ] `OUTPUT_READ_ERROR` ownership jelas;
- [ ] `Run_Summary` ownership jelas;
- [ ] CLI contract freeze jelas;
- [ ] external behavior CURRENT harus sama;
- [ ] dependency direction jelas;
- [ ] non-goals jelas;
- [ ] tidak ada unnecessary abstraction.

## FAIL

Stop. Jangan membuat implementation sebelum ownership jelas.

---

# G3 — Shared Application Structure

## Objective

Membentuk reusable single-workbook application boundary.

## Required Actions

1. Tambahkan `QCWorkbookProcessor`.
2. Gunakan existing:
   - `MiningConfig`;
   - `SamplingPointMaster`;
   - `QCDataMiningService`;
   - `ExcelExporter`.
3. Jangan ubah parser/models/service/exporter kecuali blocker nyata.
4. Compile/import module.
5. Verifikasi Git delta hanya pada scope yang disetujui.
6. Commit checkpoint sebelum mengubah `main.py`.

## Gate PASS

- [ ] `QCWorkbookProcessor` tersedia;
- [ ] module compile;
- [ ] module import;
- [ ] existing QC core behavior tidak berubah;
- [ ] tidak ada duplicate single-workbook orchestration baru;
- [ ] tidak ada circular dependency;
- [ ] tidak ada unnecessary inheritance;
- [ ] perubahan tersimpan pada checkpoint commit;
- [ ] working tree clean setelah checkpoint.

## FAIL

Stop. Jangan refactor `main.py`.

---

# Refactor `main.py`

## Objective

Menjadikan `main.py` thin CLI yang menggunakan `QCWorkbookProcessor`.

## Required Actions

1. Pertahankan CLI parser CURRENT.
2. Pertahankan validasi pasangan `--master` / `--domain`.
3. Ganti direct orchestration dengan `QCWorkbookProcessor`.
4. Pertahankan console output.
5. Pertahankan exit behavior.
6. Jangan mengubah monthly runner.

---

# G4 — Single-Workbook Functional Parity

## Objective

Membuktikan refactor `main.py` tidak mengubah behavior.

## Gate PASS

CURRENT vs refactored dengan sample yang sama:

- [ ] CLI `--help` sama;
- [ ] candidate records sama;
- [ ] worksheet count sama;
- [ ] SUCCESS sama;
- [ ] EXCLUDED sama;
- [ ] ERROR sama;
- [ ] `Data_Mining` content parity;
- [ ] `Control` content parity;
- [ ] `Processing_Log` content parity;
- [ ] valid `--master + --domain` behavior sama;
- [ ] invalid argument combination behavior sama;
- [ ] source core lain tidak berubah tanpa kebutuhan.

## PASS

Commit `main.py` sebagai checkpoint.

## FAIL

Stop. Jangan masuk monthly refactor.

---

# Refactor Monthly Flow

## Objective

Membentuk monthly application flow sesuai architecture yang sudah dibekukan.

## Required Components

### A. `MonthlySourcePlanner`

Responsibility:

```text
monthly source folder
       ↓
deterministic MonthlyPlan
```

Wajib mempertahankan CURRENT:

- week discovery;
- week ordering;
- domain routing;
- domain folder matching;
- `.xls/.xlsx` discovery;
- temp Excel exclusion jika CURRENT;
- workbook ordering;
- `MISSING_DOMAIN_FOLDER`;
- `NO_WORKBOOKS`;
- provenance identity.

Tidak boleh:

- execute mining;
- create monthly XLSX;
- call CLI;
- classify execution failure.

### B. `MonthlyMiningOrchestrator`

Responsibility:

```text
MonthlyPlan
   ↓
execute work items
   ↓
coordinate Processor + Consolidator
```

Wajib menangani:

- temporary workspace lifecycle;
- master compatibility adapter coordination;
- per-workbook temporary output;
- `QCWorkbookProcessor` invocation;
- workbook failure isolation;
- `EXEC_FAILED`;
- continue next workbook;
- finalization coordination.

Tidak boleh:

- duplicate parser logic;
- duplicate Excel consolidation logic;
- duplicate filesystem discovery logic.

### C. `MonthlyWorkbookConsolidator`

Responsibility:

```text
per-workbook result/output
       ↓
consolidated monthly workbook
```

Wajib mempertahankan:

- `Data_Mining`;
- `Control`;
- `Processing_Log`;
- `Run_Summary`;
- provenance;
- monthly totals;
- header compatibility;
- `OUTPUT_READ_ERROR`;
- formatting;
- final save behavior.

### D. `run_monthly_qc_mining.py`

Setelah monthly flow refactor:

```text
parse CLI
   ↓
MonthlyMiningOrchestrator
   ↓
present result
   ↓
exit status
```

CLI contract harus tetap CURRENT.

---

# G5 — Monthly Structural Verification

## Objective

Memastikan monthly refactor sudah benar secara modular sebelum functional parity test.

## Gate PASS

### Component existence

- [ ] `MonthlySourcePlanner` tersedia;
- [ ] `MonthlyMiningOrchestrator` tersedia;
- [ ] `MonthlyWorkbookConsolidator` tersedia;
- [ ] `QCWorkbookProcessor` direuse.

### SoC

- [ ] Planner hanya discovery/planning;
- [ ] Processor hanya single-workbook processing;
- [ ] Orchestrator hanya coordination/execution lifecycle;
- [ ] Consolidator hanya monthly output;
- [ ] monthly CLI hanya parsing/presentation.

### Dependency

Target wajib:

```text
run_monthly_qc_mining.py
        ↓
MonthlyMiningOrchestrator
        ├── MonthlySourcePlanner
        ├── QCWorkbookProcessor
        └── MonthlyWorkbookConsolidator
```

- [ ] tidak ada circular dependency;
- [ ] existing core tidak bergantung pada monthly application layer;
- [ ] monthly runner tidak import `main.py`;
- [ ] monthly runner tidak berisi duplicate QC processing orchestration.

### Compatibility structure

- [ ] master adapter masih tersedia;
- [ ] temporary XLSX masih digunakan;
- [ ] failure isolation masih berada di Orchestrator;
- [ ] `Run_Summary` recording berada di Consolidator;
- [ ] CLI signature tetap;
- [ ] module compile/import PASS.

## FAIL

Stop. Jangan masuk G6.

---

# G6 — Monthly Functional Parity

## Objective

Membandingkan CURRENT G0 baseline dengan refactored monthly flow.

## Target Flow

```text
MonthlySourcePlanner
        ↓
MonthlyMiningOrchestrator
        ↓
QCWorkbookProcessor
        ↓
MonthlyWorkbookConsolidator
```

## Gate PASS

CURRENT vs refactored:

- [ ] discovered workbooks sama;
- [ ] Weeks detected sama;
- [ ] Week mapping sama;
- [ ] Domain mapping sama;
- [ ] provenance sama;
- [ ] workbooks attempted sama;
- [ ] candidate records sama;
- [ ] worksheets sama;
- [ ] SUCCESS sama;
- [ ] EXCLUDED sama;
- [ ] ERROR sama;
- [ ] execution failures sama;
- [ ] output read errors sama;
- [ ] missing domain folder sama;
- [ ] empty/no-workbook condition sama;
- [ ] workbook failure isolation sama;
- [ ] `Data_Mining` sama;
- [ ] `Control` sama;
- [ ] `Processing_Log` sama;
- [ ] `Run_Summary` sama secara semantics;
- [ ] output filename behavior sama;
- [ ] overwrite behavior sama;
- [ ] console summary semantics sama.

### G0 numerical reference

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

## PASS

Monthly behavior equivalent.

## FAIL

Stop. Jangan lanjut ke decoupling verification.

---

# G7 — Decoupling Verification

## Objective

Membuktikan objective utama refactor tercapai.

Tidak boleh ada runtime dependency:

```text
run_monthly_qc_mining.py
        X
        └── main.py
```

## Gate PASS

- [ ] monthly runner tidak menjalankan `main.py`;
- [ ] monthly runner tidak import `main.py`;
- [ ] tidak membentuk CLI argument untuk `main.py`;
- [ ] tidak bergantung pada exit code `main.py`;
- [ ] tidak bergantung pada stdout/stderr `main.py`;
- [ ] `main.py` dan monthly runner berbagi `QCWorkbookProcessor`;
- [ ] subprocess invocation ke `main.py` sudah tidak digunakan pada active monthly path;
- [ ] dependency direction sesuai architecture freeze.

## FAIL

Stop. Objective utama belum tercapai.

---

# G8 — Stabilization

## Objective

Memastikan refactor stabil sebelum final acceptance.

## Required Actions

1. Compile/import check relevant modules.
2. Jalankan monthly execution satu kali pada refactored branch.
3. Pastikan working tree hanya memiliki expected changes.
4. Review Git diff terhadap baseline.
5. Pastikan tidak ada accidental cleanup atau unrelated change.
6. Pastikan docs sesuai implementation aktual.
7. Commit/push seluruh accepted refactor changes.

## Gate PASS

- [ ] no syntax/import error;
- [ ] no regression dari G6;
- [ ] no unrelated change;
- [ ] no dead/duplicate active execution path yang membingungkan;
- [ ] working tree clean setelah commit;
- [ ] remote branch synchronized;
- [ ] documentation aligned.

## FAIL

Stop. Jangan final acceptance.

---

# G9 — Final Acceptance

## Architecture

- [ ] SoC PASS;
- [ ] high cohesion PASS;
- [ ] low coupling PASS;
- [ ] composition-based OOP PASS;
- [ ] no circular dependency PASS;
- [ ] thin CLI entry points PASS.

## Behavior

- [ ] single-workbook parity PASS;
- [ ] monthly parity PASS;
- [ ] failure isolation PASS;
- [ ] CLI contract PASS;
- [ ] output contract PASS.

## Maintainability

- [ ] tidak ada duplicate orchestration;
- [ ] tidak ada god object baru;
- [ ] responsibility setiap component jelas;
- [ ] existing QC core tetap terpisah;
- [ ] future change dapat diarahkan ke module pemiliknya;
- [ ] tidak ada abstraction yang tidak diperlukan.

## Recovery

- [ ] pre-refactor baseline tetap tersedia;
- [ ] refactor branch/history traceable;
- [ ] rollback path jelas.

## PASS

Refactor diterima.

---

# 5. Gate Failure Rule

Untuk seluruh gate:

```text
FAIL
 ↓
stop progression
 ↓
identify exact cause
 ↓
minimum correction
 ↓
repeat same gate
```

Tidak diperbolehkan:

```text
FAIL
 ↓
continue next stage
 ↓
accumulate unrelated changes
```

---

# 6. Explicit Non-Goals

Refactor ini bukan pekerjaan untuk:

- mengubah parser;
- mengubah QC recognition logic;
- mengubah business rules;
- mengubah workbook schema;
- menghapus temporary XLSX;
- menghapus master adapter;
- mengganti framework/technology;
- menambah design pattern yang tidak dibutuhkan;
- redesign keseluruhan QC mining application;
- melakukan cosmetic cleanup besar.

Perubahan di atas hanya boleh dilakukan sebagai change terpisah setelah G9 PASS dan ada kebutuhan yang jelas.

---

# 7. Final Governing Sequence

```text
BASELINE
  ↓
ARCHITECTURE
  ↓
CONTRACT
  ↓
SHARED APPLICATION
  ↓
SINGLE-WORKBOOK PARITY
  ↓
MONTHLY REFACTOR
  ↓
MONTHLY STRUCTURAL VERIFICATION
  ↓
MONTHLY FUNCTIONAL PARITY
  ↓
DECOUPLING VERIFICATION
  ↓
STABILIZATION
  ↓
FINAL ACCEPTANCE
```

Dokumen ini menjadi acuan urutan refactor. Tidak menambah, menghapus, atau memindahkan gate tanpa review dan persetujuan terlebih dahulu.
