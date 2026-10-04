# Changelog

Dates are YYYY-MM-DD.

## 0.1.0 (2026-09-15)

First public release: marking_search, repository skeleton, CI, hygiene gates.

Added

- `marking_search/marking_search.py`: read-only scan of a document store for export-control,
  CUI, DFARS/CMMC, and flow-down vocabulary. One CSV of paths, categories, counts, and matched
  terms; a paste-safe report block of counts only. Standard library; `pypdf` optional for PDFs.
- `marking_search/VERIFICATION.md`: results of the five-suite stress test, local paths
  replaced with relative names.
- `tests/run_all.py`: builds a synthetic tree at runtime, runs the tool, checks exit codes,
  category counts (branching on whether `pypdf` imports), the content-leak sentinel, the
  refused output path, and that the scanned tree is untouched.
- `tests/hygiene_check.py`: fails on any forbidden string in a shipped file or a dirty
  working tree after the tests.
- CI on ubuntu and windows, Python 3.10, 3.12, 3.13, 3.14, with and without `pypdf`.
- `docs/SPEC.md`, `SECURITY.md`, MIT `LICENSE`.

Changed, relative to the verified internal build

- Docstring says to run the tool on the machine that holds the documents; no organisation
  name anywhere in the file.
- `pypdf` logger gets a `NullHandler` and level CRITICAL, and does not propagate (was ERROR).
- `--version` prints `marking_search 0.1.0`.
- The interpretation lines (GCC High, 252.204-7019 and SPRS) moved from stdout to the README.
  stdout keeps the counts. CLI, CSV schema, and the report block are otherwise unchanged.

Coming in 0.2.0

- `cui_scan`: a fail-closed gate for one file or folder before it goes near an AI assistant or
  a cloud service. Change list in `docs/SPEC.md`, section 4.
