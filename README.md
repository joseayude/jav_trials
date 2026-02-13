# jav_trials

## Repository Structure

High-level overview of the project layout with key files and directories.

```
.
├─ LICENSE                      # Project license
├─ pyproject.toml               # Project and dependency config
├─ README.md                    # Project overview and docs
├─ src/
│  └─ xls_management/           # Main package implementing XLS utilities
│     ├─ __init__.py            # Package initializer
│     ├─ config.py              # Configuration helpers
│     ├─ is_ole.py              # Detect OLE/legacy XLS files
│     ├─ workbook.py            # Workbook read/write helpers
│     ├─ ate/                   # ATE-related modules
│     │  ├─ __init__.py
│     │  ├─ data.py             # ATE data models/utilities
│     │  ├─ project.py          # Project-level ATE helpers
│     │  └─ tracking.py         # ATE tracking utilities
│     ├─ shell/                 # CLI / shell helpers
│     │  ├─ __init__.py
│     │  └─ ate.py
│     ├─ tui/                   # Terminal UI components
│     │  ├─ file_picker.py
│     │  ├─ main.py
│     │  ├─ msgbox.py
│     │  └─ project_form.py
│     └─ utils/                 # Support utilities used across package
│        ├─ __init__.py
│        ├─ aux.py              # Aux helpers used in tests and modules
│        └─ color.py            # Terminal color helpers
└─ test/                        # Unit tests and fixtures
	├─ test_config.py
	├─ test_ole.py
	├─ test_workbook.py
	└─ test_ate/
		└─ test_om/
			└─ test_db_info.py
```

- `LICENSE`, `pyproject.toml`, `README.md`: project metadata and dependency configuration.
- `src/xls_management/`: main package implementing XLS management utilities and tools.
- `src/xls_management/ate/`: ATE-related modules and domain-specific subpackage `om`.
- `src/xls_management/tui/`: terminal UI components and forms.
- `src/xls_management/utils/`: helper utilities used across the package.
- `test/`: unit tests and test fixtures mirroring the package structure


Notes:
- Use `src/xls_management/workbook.py` for primary XLS operations.
- Tests mirror the source layout under `test/`; run them with your chosen test runner.

