# VBA → Python Mapping for xls_management

This document maps the VBA procedures/subs from the `ate_status_sequence.puml` flow to the corresponding Python classes, methods and files in this repository (`src/xls_management`). It is intended to guide translation of remaining VBA logic into Python.

## Overview
- Orchestrator (VBA): `ATE_Status` → Python: `xls_management.ate.tracking.ATEStatus` (`src/xls_management/ate/tracking.py`) with `perform_status()` driving the sequence.
- UI (VBA): `BoxAuswahlProjekt (UI)` → Python: `ProjectChoice` + `project_combo_box()` (`src/xls_management/tui/project_form.py`, `src/xls_management/ate/project.py`).
- Config: VBA initializer config → Python: `xls_management.config.ATEConfig` (`src/xls_management/config.py`).
- Workbook handling: VBA workbook open/close → Python: `xls_management.workbook.Workbook` (`src/xls_management/workbook.py`).
- File picking UI: VBA `GetOpenFilename` → Python: `path_from_file_picker()` (`src/xls_management/tui/file_picker.py`).
- Message boxes: VBA `MsgBox` → Python: `msgbox()` (`src/xls_management/tui/msgbox.py`).

## Direct mapping (PUML sequence → Python)

- BoxAuswahlProjekt (UI)
  - VBA: show selection dialog / unload
  - Python: `ProjectChoice` UI app
    - File: `src/xls_management/tui/project_form.py`
    - Helper: `project_combo_box()` in `src/xls_management/ate/project.py`

- ATE_Status (main orchestration)
  - VBA: `ATE_Status` Sub
  - Python: `class ATEStatus` in `src/xls_management/ate/tracking.py`
    - entry: `perform_status()` → implements sequence from PUML
    - helpers: `initialized()`, `status_deinitialize()`, and many read/write methods called from `perform_status()`

- ATE_Status_Initializer / ATE_Status_Deinitializer
  - VBA: `ATE_Status_Initializer`, `ATE_Status_Deinitializer`
  - Python: implemented as `ATEStatus.initialized()` and `ATEStatus.status_deinitialize()` in `src/xls_management/ate/tracking.py`. Use `ATEConfig` for config values.

- EinlesenLAHBlacklist
  - VBA: read LAH blacklist
  - Python: `ATEStatus.read_blacklist()` (method within `src/xls_management/ate/tracking.py`)

- EinlesenTDVKs (Testdesigns - Verifikationskriterien)
  - VBA: read TDVK workbook
  - Python: `ATEStatus.read_TDVKs()` in `src/xls_management/ate/tracking.py`
    - Uses: `DBInfo.einlesen_datei()` (`src/xls_management/ate/om/db_info.py`) for attribute discovery
    - Domain model: `verificationskriterium` in `src/xls_management/ate/om/verificationskriterium.py`

- EinlesenTDAAs (Testdesigns - Absicherungsaufträge)
  - VBA: read TDAA workbook
  - Python: `ATEStatus.read_TDAAs()` in `src/xls_management/ate/tracking.py`
    - Domain model: `absicherungsauftraege.py`

- EinlesenTFs (Testfälle)
  - VBA: read TF workbook
  - Python: `ATEStatus.read_TFs()` in `src/xls_management/ate/tracking.py`
    - Domain model: `testfaelle.py`

- EinlesenFRUTiming
  - VBA: read FRU timing data
  - Python: `ATEStatus.read_FRU_timing()` in `src/xls_management/ate/tracking.py`
    - Model: `FRUTiming` in `src/xls_management/ate/om/fru_timming.py`

- EinlesenAVWRohdaten / EinlesenAVWVorgaengerRohdaten / EinlesenAVWNachfolgerRohdaten
  - VBA: read AVW raw data and optionally master/predecessor+successor flows
  - Python:
    - AVW attributes: `AVW_ATTRIBUTE_DE` in `src/xls_management/ate/data.py`
    - Project-specific import: `ProjectDBInfo.einlesen_datei()` in `src/xls_management/ate/om/project_db_info.py`
    - Models: `BSMDaten` (`src/xls_management/ate/om/bsm_daten.py`), `AVWVorgaenger` (`src/xls_management/ate/om/avw_vorganenger.py`)
    - Orchestrator methods: `ATEStatus.read_raw_data_AVW()`, `ATEStatus.read_predecesor_raw_data_AVW()`, `ATEStatus.read_successor_raw_data_AVW()` in `src/xls_management/ate/tracking.py`

- AusgabeATEStatus / AusgabeTDStatus
  - VBA: write ATE and TD status to BsM workbook
  - Python: `ATEStatus.output_status()` and `ATEStatus.output_status_TD()` methods in `src/xls_management/ate/tracking.py`
    - Use `Workbook` (sheet read/write) helpers in `src/xls_management/workbook.py`.

- SchliessenWb
  - VBA: close open workbooks
  - Python: `ATEStatus.close_workbooks()` in `src/xls_management/ate/tracking.py` (using `Workbook` abstraction)

- AusgabeVerlauf
  - VBA: fill history sheets (ATE_Status_Verlauf, TD_Status_Verlauf)
  - Python: `ATEStatus.ausgabe_verlauf()` in `src/xls_management/ate/tracking.py`

## Data classes and helpers (where to add/extend)
- `src/xls_management/ate/data.py`
  - `AVW_ATTRIBUTE_DE` already present; add other shared constants or dataclass definitions if needed.
- `src/xls_management/ate/om/db_info.py`
  - `DBInfo` and `ProjectDBInfo` handle file picking + attribute detection; good place for reader utilities.
- `src/xls_management/ate/om/*.py`
  - Domain models are present (`bsm_daten.py`, `avw_vorganenger.py`, `fru_timming.py`) — extend these with parsers/mapping from DataFrame rows to instances.
- `src/xls_management/workbook.py`
  - Enhance with write helpers if the ATE output needs structured writing (e.g., `write_sheet_from_df()`), and context manager support for open/save/close.

## Control-flow mapping (high-level)
1. UI selection via `project_combo_box()` → returns `(project, use_predecesor_ids)`
2. `ATEStatus.perform_status()` calls `initialized()` → prepares `DBInfo/ProjectDBInfo` instances (attribute discovery), sets up `Workbook` for BsM
3. If initialization succeeds: call readers in order (`read_blacklist`, `read_TDVKs`, `read_TDAAs`, `read_TFs`, `read_FRU_timing`)
4. Branch on `use_predecesor_ids`:
   - False: `read_raw_data_AVW()`
   - True: `read_predecesor_raw_data_AVW()` then `read_successor_raw_data_AVW()`
5. Call outputs: `output_status()`, `output_status_TD()`, `close_workbooks()`
6. Call `ausgabe_verlauf()` to write history and aggregate `self.errors`
7. Display final `msgbox()` with success or errors and call `status_deinitialize()`

## Suggested next steps
- If you want, I can create `docs/STRUCTURE.md` or add code stubs for any missing `ATEStatus.*` methods in `src/xls_management/ate/tracking.py` or add unit tests that validate reading/writing flows.

---

*File generated from repository analysis on 2026-02-13.*
