# cui-tools: repository spec

Owner: builtbyisaiah (GitHub). Revision 2, 2026-09-15, after independent
review (approve with changes; all findings adopted). Scope rule: each release fits one sitting.

## 0. Release plan

- **v0.1.0**: marking_search, repo skeleton, CI, hygiene gates. One sitting. The tool is
  already verified against five adversarial suites.
- **v0.2.0**: cui_scan, after its source is copied in and every item in section 4 is done and
  tested. Its own sitting.

## 1. What this repo is

Two small, local, content-free triage tools for Controlled Unclassified Information (CUI):

- **marking_search** (v0.1.0): a read-only scoping scan across a whole document store (a synced
  SharePoint or OneDrive tree, a file server) that answers "do we hold export-controlled or
  CUI-marked material, and which contract clauses bind us" before a CMMC or NIST SP 800-171
  assessment. Output: counts, categories, a paste-safe report block; a CSV of paths and matched
  vocabulary, never content.
- **cui_scan** (v0.2.0): a fail-closed gate for one file or folder BEFORE it goes near an AI
  assistant or a cloud service. Output: verdicts and metadata only.

Which tool: "a whole store before an assessment" is marking_search; "one file before pasting"
is cui_scan.

Shared contract, stated in the README's first paragraph and enforced by tests:
1. Runs locally. No network calls. Python 3 standard library only; `pypdf` optional for PDFs.
2. Read-only on the scanned tree. Writes at most one output file, and refuses any output path
   inside the scanned tree (exit 2 before anything is written). Cloud placeholder files (OneDrive
   online-only) are detected by attribute and never opened, so the scan never triggers a download.
3. Never emits document content. Output is verdicts, counts, fixed vocabulary terms, paths.
   CI plants a sentinel string in fixtures and asserts it appears in no stdout, stderr, or CSV.
4. Detects markings and vocabulary. No unmarked-CUI detection, no OCR, no ML, not a compliance
   certification. "No markings found" means "no textual marking was found in the parts the
   scanner could decode," not "safe."

## 2. Why publish, and the honest framing

The useful idea is the output contract: a scanner that can sit inside an AI-assistant loop as a
boundary gate without itself exporting content. Regex over markings is not novel; the README
says so. The README says the tools were built with AI assistance and reviewed by the author,
and it states the known failure modes up front. The test corpus is entirely synthetic; nothing
from a real scan is in this repo; operational use is on your own machine, by design.

## 3. Layout (v0.1.0)

```
cui-tools/
  README.md                 contract, which tool, quick start, output meaning, known limits,
                            how tested, references, exit codes, one synthetic sample output
  LICENSE                   MIT
  SECURITY.md               report false negatives privately; never attach a real document
  CHANGELOG.md
  .gitignore                *.csv, __pycache__/, out/, fixtures/, *.bak, .venv/
  .gitattributes            * text=auto eol=lf
  .github/workflows/ci.yml  section 7
  docs/SPEC.md              this file, scrubbed of any client name
  marking_search/
    marking_search.py
    VERIFICATION.md         five-suite results, local paths scrubbed to relative names
  tests/
    run_all.py              builds a synthetic tree in a temp dir; runs the tool with --report and
                            --out to a sibling temp dir; asserts exit 0, expected category counts
                            (branching on pypdf presence), and the sentinel-absence check
    hygiene_check.py        fails if any tracked file matches the forbidden-string list or if
                            `git status --porcelain` is non-empty after tests
```
No per-tool READMEs in v0.1; `--help` plus the root README. Nothing under `_genlibs` or any
scratch helper ships.

## 4. cui_scan: v0.2.0 change list (from the 2026-09-01 evaluation plus review)

Each item has a fixture. All required before v0.2.0.

1. **Self-test cannot destroy anything.** `tempfile.mkdtemp()` it creates and removes itself.
2. **Self-test is stdlib-only.** Builds docx/xlsx/pptx with `zipfile`; PDF cases expect
   UNKNOWN when `pypdf` is absent, MARKED or NO_MARKINGS_FOUND when present.
3. **No content leak.** `pypdf` logger: `NullHandler` plus level CRITICAL. `sys.stdout` and
   `sys.stderr` reconfigured to UTF-8.
4. **Verdict `CLEAN` renamed `NO_MARKINGS_FOUND`.** Same ordering.
5. **Exit codes.** 0 = every file NO_MARKINGS_FOUND; 1 = any MARKED, UNKNOWN, or SUSPECT;
   2 = usage or refused output path.
6. **`--csv` under any scanned path exits 2** before writing (same `commonpath` check as
   marking_search). CSV cells guarded against formula injection; written `utf-8-sig`.
7. **Cloud placeholders** (OFFLINE, RECALL_ON_OPEN, RECALL_ON_DATA_ACCESS attributes) reported
   as UNKNOWN and never opened.
8. **Office extraction that does not glue words.** The run-joining extractor from
   marking_search: runs inside a paragraph or cell joined; paragraph and cell ends become
   newlines; tabs become spaces; headers, footers, footnotes, comments, text boxes, inline
   strings, sheet names, notes, masters read. Fixture: header `CUI<tab>Page 1` is MARKED.
9. **Encoding and type sniffing.** Magic bytes (`%PDF`, `PK`, OLE) checked before extension;
   BOM and UTF-16 sniffing; text with more than 1 percent U+FFFD after decoding is UNKNOWN.
   Fixture: UTF-16 text with a CUI banner is MARKED; a PDF renamed `.txt` is not cleared.
10. **SUSPECT vocabulary in `suspect_terms.json`** with profiles `generic` (default) and
    `physical-security`, chosen by `--profile`. Nothing hardcoded.
11. **Two false positives fixed.** The inverted `CLASSIFICATION` regex (fires on
    `ATTACHMENT(S) ARE`, misses real `(S)` marks). `Controlled by:` alone yields SUSPECT; with
    any other MARK hit on the file, MARKED. One fixture each.
12. **Skips its own reports** by name pattern.
13. **Header hygiene.** No client name anywhere. MIT header. `--version`.
14. **Known limits documented in the README, not hidden:** fonts without a ToUnicode map,
    ligatures and homoglyphs, form fields, stamp appearance streams, XFA, attachments, PDF and
    Office metadata including sensitivity labels, scanned images, container parts that fail to
    parse, Word all-caps formatting (`w:caps`, text stored lowercase), VML watermarks
    (`v:textpath string=`). Each a v0.3 candidate, none claimed fixed.

## 5. marking_search: changes for v0.1.0

1. Docstring: "run it on the machine that holds the documents"; no company names anywhere.
2. `pypdf` logger: `NullHandler` plus level CRITICAL (today it is set to ERROR only).
3. `--version`. CLI, CSV schema, and report-back block otherwise unchanged.
4. Move the GCC High and 7019 hint lines from stdout into the README; stdout keeps the counts.
5. VERIFICATION.md scrubbed to relative names.

## 6. Hygiene gates (CI and pre-push)

- Fresh `git init`; no history from any other machine.
- Commit identity set inside the repo only: display name plus the GitHub no-reply address
  (`<id>+builtbyisaiah@users.noreply.github.com`). The personal address never appears in
  history. Global git config stays empty.
- `tests/hygiene_check.py` runs a case-insensitive regex search for the forbidden list (the
  client's name and abbreviation, the machine username, the Windows profile path prefix, and
  the two local temp-folder names) over tracked files and fails on any hit; also fails if
  `git status --porcelain` is non-empty after the test run. Nothing is exempt, this file
  included, so the list is spelled out only in fragments inside the script.
- No CSV, no `out/`, no fixture outputs, no `.bak`, nothing generated from a real scan.
- Every fixture is synthetic and built at runtime; none is committed.
- `.gitattributes` pins LF. Subprocesses in tests pass `encoding="utf-8"`; CSV reads use
  `utf-8-sig`; tests always pass `--out` to a sibling temp dir.

## 7. CI

GitHub Actions on push and pull request. Matrix: os in {ubuntu-latest, windows-latest};
python in {3.10, 3.12, 3.13, 3.14}; deps in {full, minimal}. `full` installs `pypdf`; `minimal`
installs nothing (`if: matrix.deps == 'full'` on the install step). Steps: `py_compile` on
every tool; `python tests/run_all.py` (expected counts branch on whether `pypdf` imports);
`python tests/hygiene_check.py`.

## 8. README outline

1. The four-point contract, one paragraph.
2. Which tool (two lines).
3. Quick start, three commands max per tool.
4. What the output means: the report-back block, field by field; exit codes.
5. One synthetic sample output.
6. Known limits, honest.
7. How it was tested: the five-suite stress test for marking_search (counts verbatim); for
   cui_scan when it lands: "about 300 synthetic fixtures, 114 confirmed findings, 79 severity-1
   misses in v0; v0.2 fixes three classes, the rest are listed above."
8. References: ISOO CUI Marking Handbook v1.1, DoDI 5200.48, DFARS 252.204-7012.
9. Built with AI assistance, reviewed by the author. MIT.

## 9. Release

Tag `v0.1.0` after CI is green on every matrix cell. Release notes are the CHANGELOG entry.
Repo description: "Local, content-free triage tools for CUI: scope a document store before a
CMMC assessment; gate a document before an AI sees it." Topics: cui, cmmc, nist-800-171,
dfars, python.

## 10. Out of scope

OCR, ML classification, per-page verdicts, confidence scores, a GUI, PyPI packaging.
