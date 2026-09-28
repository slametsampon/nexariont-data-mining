# NEXARIONT QC Mining Refactor — Concept, Workflow & Gate Check

**Status:** Proposed Refactor Baseline  
**Purpose:** Pegangan/rujukan refactor `main.py` dan `run_monthly_qc_mining.py`  
**Primary Objective:** Melepaskan coupling `run_monthly_qc_mining.py` terhadap `main.py` tanpa mengubah business behavior QC mining.

---

## 1. Refactor Principle

Refactor harus:

- **painless dan reversible**;
- menggunakan **modularity / Separation of Concerns (SoC)**;
- mengikuti **OOP principle** dengan cohesion tinggi dan coupling rendah;
- mengutamakan **composition**, bukan inheritance tanpa kebutuhan nyata;
- bersifat **structural refactor, not behavioral redesign**;
- tidak mengubah parser, business rule, output contract, atau processing semantics tanpa kebutuhan yang terverifikasi;
- tidak melakukan cleanup/optimization yang berada di luar objective utama.

Prinsip dependency:

```text
CLI / Entry Point
       ↓
Application Layer
       ↓
Existing QC Core
```

Tidak diperbolehkan lagi:

```text
run_monthly_qc_mining.py
        ↓
      main.py
```

---

## 2. Target Architecture

```text
                         ENTRY POINTS
              ┌──────────────────────────────┐
              │                              │
              ▼                              ▼
           main.py              run_monthly_qc_mining.py
        Single-file CLI              Monthly CLI
              │                              │
              ▼                              ▼
     QCWorkbookProcessor       MonthlyMiningOrchestrator
              ▲                     │        │
              │                     │        ├── MonthlySourcePlanner
              │                     │        └── MonthlyWorkbookConsolidator
              └─────────────────────┘
                         │
                         ▼
                  EXISTING QC CORE
```

### 2.1 Responsibility

| Component | Primary Responsibility |
|---|---|
| `main.py` | Thin CLI untuk single-workbook: parse input, call application, present result |
| `run_monthly_qc_mining.py` | Thin CLI untuk monthly run: parse input, call application, present result |
| `QCWorkbookProcessor` | Memproses satu workbook QC end-to-end menggunakan existing core |
| `MonthlySourcePlanner` | Mengubah struktur folder monthly menjadi deterministic work plan |
| `MonthlyMiningOrchestrator` | Mengkoordinasikan monthly execution |
| `MonthlyWorkbookConsolidator` | Menggabungkan hasil menjadi monthly workbook |

Existing lower-level/core components tetap digunakan, misalnya:

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

## 3. SoC & OOP Rules

### 3.1 Single Responsibility

```text
Planner      → menemukan pekerjaan
Processor    → memproses satu workbook
Consolidator → menyusun output monthly
Orchestrator → mengkoordinasikan
CLI          → interaksi user / command line
```

Tidak boleh ada satu class/file yang sekaligus melakukan discovery, processing, consolidation, dan presentation.

### 3.2 Encapsulation

Detail berikut harus berada pada komponen pemiliknya:

- filesystem discovery → `MonthlySourcePlanner`;
- single-workbook processing → `QCWorkbookProcessor`;
- monthly workbook construction → `MonthlyWorkbookConsolidator`;
- execution coordination → `MonthlyMiningOrchestrator`;
- CLI argument / console output → entry point.

### 3.3 Composition Over Inheritance

`MonthlyMiningOrchestrator` menggunakan:

```text
MonthlySourcePlanner
QCWorkbookProcessor
MonthlyWorkbookConsolidator
```

Tidak membuat hierarchy seperti `BaseProcessor`, `AbstractRunner`, dan sejenisnya tanpa kebutuhan nyata.

### 3.4 Modularity Criterion

Modularity dinilai dari:

- responsibility;
- cohesion;
- coupling;
- dependency direction;
- testability;
- clarity of contract.

**Bukan** dari banyaknya file, class, layer, atau design pattern.

---

## 4. Conceptual Contracts

Nama final method/class dapat disesuaikan saat design review. Contract konseptual:

```text
MonthlySourcePlanner.plan(...)
    → MonthlyPlan

QCWorkbookProcessor.process(...)
    → ProcessingResult

MonthlyMiningOrchestrator.run(...)
    → MonthlyRunResult

MonthlyWorkbookConsolidator.add(...)
MonthlyWorkbookConsolidator.finalize(...)
```

### 4.1 Monthly Work Item

Work item membawa identity/provenance yang diperlukan:

```text
Month
Week
Domain
Source Workbook
Source Path
```

Provenance tidak perlu dibangun ulang oleh setiap layer.

### 4.2 Error Policy

Application module tidak memakai:

```text
sys.exit()
stdout/stderr contract
CLI return code sebagai application contract
```

Error policy berada pada caller:

```text
main.py
  exception → CLI error / exit status

MonthlyMiningOrchestrator
  workbook failure → record failure → continue next workbook
```

---

## 5. Current Compatibility to Preserve

Pada refactor ini, hal berikut **tetap dipertahankan** kecuali blocker nyata ditemukan:

- existing parser/business logic;
- existing service/model behavior;
- existing output workbook schema;
- temporary per-workbook XLSX;
- current master adapter mechanism;
- current monthly provenance semantics;
- current failure isolation;
- current `Data_Mining`;
- current `Control`;
- current `Processing_Log`;
- current `Run_Summary`;
- current domain routing.

Optimization seperti menghapus temporary XLSX atau master adapter merupakan pekerjaan terpisah setelah refactor diterima.

---

# 6. Refactor Workflow & Gate Check

```text
CURRENT BASELINE
      │
      ▼
[G0] Baseline & Recovery
      │ PASS
      ▼
[G1] Architecture / SoC Freeze
      │ PASS
      ▼
[G2] Contract Freeze
      │ PASS
      ▼
Create Shared Application Components
      │
      ▼
[G3] Structural Verification
      │ PASS
      ▼
Refactor main.py
      │
      ▼
[G4] Single-Workbook Parity
      │ PASS
      ▼
Refactor Monthly Flow
      │
      ▼
[G5] Monthly Functional Parity
      │ PASS
      ▼
Remove Runtime Coupling to main.py
      │
      ▼
[G6] Decoupling Verification
      │ PASS
      ▼
Stabilization
      │
      ▼
[G7] Final Acceptance
```

---

## G0 — Baseline & Recovery

### Objective
Menjamin seluruh refactor reversible sebelum source code diubah.

### Gate Check

- `master` clean.
- `master` aligned dengan `origin/master`.
- refactor branch dibuat sebelum coding;
- existing baseline execution dapat dijalankan;
- baseline output tersedia untuk comparison.

### PASS
Baseline aman dan reproducible.

### FAIL
Jangan mulai modifikasi.

---

## G1 — Architecture / SoC Freeze

### Required Architecture

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

### Gate Check

- satu class = satu primary responsibility;
- tidak ada `main.py → monthly.py` atau `monthly.py → main.py`;
- tidak ada filesystem discovery di CLI;
- tidak ada Excel consolidation di CLI;
- tidak ada parser/business logic di orchestrator;
- composition digunakan;
- inheritance baru hanya jika ada kebutuhan nyata;
- tidak ada circular dependency.

### PASS
Responsibility dan dependency direction jelas.

---

## G2 — Contract Freeze

### Objective
Menetapkan interface antar-module sebelum coding.

### Gate Check

Harus jelas:

- input/output `QCWorkbookProcessor`;
- bentuk `MonthlyPlan`;
- bentuk `MonthlyRunResult`;
- ownership temporary output;
- ownership master adapter;
- ownership error handling;
- ownership `Run_Summary`;
- provenance source;
- siapa yang memutuskan continue/stop saat workbook gagal.

Tambahan:

- caller tidak perlu tahu internal implementation;
- application object tidak menerima `argparse.Namespace`;
- tidak menggunakan stdout/stderr sebagai application contract;
- tidak membuat DTO/interface/class baru tanpa kebutuhan.

### PASS
Tidak ada ambiguity pada ownership dan interface.

---

## G3 — Structural Verification

### Objective
Menambahkan application components tanpa mengubah behavior existing QC core.

### Expected Treatment

```text
Parser                  NO CHANGE
Models                  NO CHANGE
QCDataMiningService     NO CHANGE kecuali blocker terverifikasi
ExcelExporter           NO CHANGE kecuali blocker terverifikasi
Business logic          NO CHANGE
Output schema           NO CHANGE
```

### Gate Check

- dependency:

```text
CLI → Application → Existing Core
```

- tidak ada circular dependency;
- tidak ada duplicate single-workbook orchestration;
- shared module dapat dipanggil oleh dua entry path;
- tidak ada behavioral change yang tidak diperlukan.

### PASS
Struktur baru tersedia dengan impact minimum.

---

## G4 — Single-Workbook Parity

### Objective
Memastikan `main.py` melalui `QCWorkbookProcessor` menghasilkan behavior yang sama.

```text
BEFORE
main.py → existing processing

AFTER
main.py → QCWorkbookProcessor → existing processing
```

### Gate Check

Dengan input/master/domain yang sama:

- candidate records sama;
- worksheet status sama;
- `Data_Mining` sama;
- `Control` sama;
- `Processing_Log` sama;
- failure behavior equivalent.

### PASS
Single-workbook behavior tetap.

### FAIL
Stop. Perbaiki root cause dan ulangi G4.

---

## G5 — Monthly Functional Parity

### Target Flow

```text
MonthlySourcePlanner
        ↓
MonthlyMiningOrchestrator
        ↓
QCWorkbookProcessor
        ↓
MonthlyWorkbookConsolidator
```

### Gate Check

CURRENT vs refactored:

- discovered workbooks sama;
- Week mapping sama;
- Domain mapping sama;
- provenance sama;
- candidate records sama;
- SUCCESS / EXCLUDED / ERROR sama;
- workbook failure isolation sama;
- `Data_Mining` sama;
- `Control` sama;
- `Processing_Log` sama;
- `Run_Summary` sama secara semantics.

### PASS
Monthly behavior equivalent.

### FAIL
Stop. Jangan lanjut ke decoupling final.

---

## G6 — Decoupling Verification

### Objective
Membuktikan objective utama refactor tercapai.

Tidak boleh ada runtime dependency:

```text
run_monthly_qc_mining.py
        X
        └── main.py
```

### Gate Check

Monthly flow:

- tidak menjalankan `main.py`;
- tidak import `main.py`;
- tidak membentuk CLI argument untuk `main.py`;
- tidak bergantung pada exit code `main.py`;
- tidak bergantung pada stdout/stderr `main.py`.

Valid dependency:

```text
main.py ───────────────┐
                      ▼
              QCWorkbookProcessor
                      ▲
                      │
MonthlyOrchestrator ──┘
```

### PASS
Coupling antar-entry-point telah hilang.

---

## G7 — Final Acceptance

### Architecture

- SoC: PASS
- High cohesion: PASS
- Low coupling: PASS
- Composition-based OOP: PASS
- No circular dependency: PASS
- Thin CLI entry points: PASS

### Behavior

- Single-workbook parity: PASS
- Monthly parity: PASS
- Failure isolation: PASS
- Output contract: PASS

### Maintainability

- responsibility masing-masing class jelas;
- tidak ada duplicate orchestration;
- tidak ada class/file menjadi "god object";
- tidak ada pattern/abstraction yang tidak dibutuhkan;
- perubahan berikutnya dapat dilakukan pada module terkait tanpa menyentuh layer lain secara tidak perlu.

---

# 7. Gate Failure Rule

Jika gate gagal:

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
continue coding next stage
 ↓
accumulate unrelated changes
```

---

# 8. Version-Control Strategy

Sebelum coding:

```text
master
  │
  └── refactor/<agreed-branch-name>
```

Prinsip:

- `master` tetap sebagai recoverable baseline;
- refactor dilakukan pada branch terpisah;
- commit dibuat per logical stage/gate;
- jangan mencampur structural refactor dengan cleanup/optimization;
- jika hasil tidak acceptable, baseline dapat dipulihkan tanpa reconstruct perubahan.

---

# 9. Explicit Non-Goals

Refactor ini **bukan** pekerjaan untuk:

- mengubah QC parser;
- mengubah business rule;
- mengubah recognition logic;
- mengubah data model tanpa kebutuhan;
- mengubah workbook schema;
- menghilangkan temporary XLSX;
- menghilangkan master adapter;
- mengganti framework/technology;
- menambah hierarchy OOP yang tidak diperlukan;
- melakukan cosmetic cleanup besar;
- redesign keseluruhan QC mining application.

Semua item tersebut hanya boleh menjadi change terpisah setelah G7 PASS dan apabila ada kebutuhan yang jelas.

---

# 10. Final Refactor Principle

```text
ENTRY
  ↓
PLAN
  ↓
EXECUTE
  ↓
CONSOLIDATE
  ↓
PRESENT
```

Target akhir:

> `main.py` dan `run_monthly_qc_mining.py` hanya menjadi thin entry points. Single-workbook processing dipusatkan pada `QCWorkbookProcessor`. Monthly source discovery menjadi tanggung jawab `MonthlySourcePlanner`. Monthly execution dikoordinasikan oleh `MonthlyMiningOrchestrator`. Monthly output dikelola oleh `MonthlyWorkbookConsolidator`. Existing QC core tetap menjadi implementation basis dan tidak diubah tanpa kebutuhan yang terverifikasi.

Refactor dinyatakan selesai hanya setelah **G0–G7 PASS**.
