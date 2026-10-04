# marking_search.py verification (2026-09-15)

Tool: `marking_search/marking_search.py` (this repository).
Suites under a local `stress/` folder (not shipped), run with `PYTHONIOENCODING=utf-8`, Python 3.12.10,
pypdf 6.16.2, cryptography NOT installed, LongPathsEnabled=0. Logs: `logs/before_*.log`, `after_*.log`.
py_compile clean; fixtures untouched.

## Per suite (before = harness readers fixed, tool as received; after = tool fixed)

| suite | before | after |
|---|---|---|
| regex, 65 probes | ok 64, false neg 0, false pos 1 | unchanged |
| formats, 82 fixtures | checker crashed; CSV: 72 OK, 8 accepted notes, 2 wrong (owner-only AES PDFs) | 72 OK, 8 accepted, 2 SKIPPED_PDF_AES (honest label, see below) |
| safety, 23 checks | PASS 21, FAIL 2 (SUM-1, SUM-2) | PASS 23, FAIL 0 |
| operator, A to K | 0 misses on 22 true docs, 4 decoys by design; C failed on a stale csv | all pass; K: 6254 files in 17.9 s |
| winops, T1 to T12 | T2/T3 crashed (exit 1, no CSV); legacy.doc missed; T7 peak 209 MiB | T2/T3 pass; legacy.doc found; T7 peak 150 MiB, 44 s |

## Tool defects fixed

1. Root given with the `\\?\` prefix (winops T2/T3): `os.path.relpath` raised ValueError mid-walk and the run died with no CSV or summary. Root is normalised with `display()`; the folder summary uses string math (`top_folder`).
2. A root over 260 chars was rejected as "not a folder": `isdir` now runs on `extended(root)`.
3. Any unexpected per-file exception ended the run: work moved into `scan_file()` under a guard that writes an ERROR row and continues.
4. Legacy OLE byte-scan: the 8-byte magic decodes to letters glued onto adjacent text, so `\bITAR\b` failed; magic and 0xFF/0xFE fill bytes are skipped now.
5. AES-128 PDFs: pypdf raises DependencyError inside `extract_text`; it was swallowed and shown as ERROR ValueError, now labelled SKIPPED_PDF_AES.
6. `normalize()` no longer copies pure-ASCII text (59 MB file: peak 209 to 150 MiB). Time is regex (21 s per 59 MB), acceptable; progress still prints every 500 files.

Note: the brief tied the ValueError to safety SUM-1/SUM-2. That was wrong: the 325-char deep file was always scanned. SUM-1/2 failed because the harness snapshot predates the deny ACL and expected a file inside an unlistable folder to appear by name. The tool already names that folder (ERROR_DIR row) and counts it under "folders not listable".

## Harness changes

- regex, formats check_results, operator (read_csv, H, K): read the CSV with `utf-8-sig`.
- formats check_results: `classify()` tolerates `extract()` raising, as main() does.
- formats run_suite.sh: asserts the default CSV lands next to the script and `--out` inside the tree exits 2.
- safety: S1b reads and removes the CSV next to the script; SUM-1 counts files inside unlistable folders and requires "folders not listable" >= 1; SUM-2 accepts the parent ERROR_DIR row for secret_itar.txt.
- operator: C removes a stale C_typo.csv and expects exit 2, no CSV; E expects exit 2 for `--out` inside and the CSV next to the script for cwd=root; F expects exit 2 before scanning.
- winops: T8 expects the CSV next to the script (then removes it); T9 expects exit 2 twice, no CSV.

## Still failing, and why

- regex `fp_export_controller_software`: "export controls" (UI) matches. Kept on purpose: "subject to U.S. export controls" is real marking language; the suite's own `proposed_patterns.py` fails the same probe.
- formats owner-only AES-128/256 PDFs: SKIPPED_PDF_AES because cryptography is missing for the tool's Python. With `PYTHONPATH=formats/_genlibs` both read as EXPORT_CONTROL. Install `cryptography` on the machine that runs the scan.
- Operator decoys (surname Cui, EXPT, HVAC "flow down", a resume): a string tool cannot tell these apart; run with `--exclude HR Personal`.
