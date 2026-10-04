# cui-tools

Small, local, content-free triage tools for Controlled Unclassified Information (CUI).

Every tool here keeps the same four-point contract, and the tests enforce it. (1) It runs
locally: no network calls, Python 3 standard library only, `pypdf` optional for PDFs. (2) It is
read-only on the scanned tree: it writes at most one output file, refuses any output path
inside the scanned tree (exit 2 before anything is written), and detects OneDrive online-only
placeholder files by attribute without opening them, so a scan never triggers a download.
(3) It never emits document content: output is verdicts, counts, fixed vocabulary terms, and
paths, and CI plants a sentinel string in a fixture and asserts it appears in no stdout, stderr,
or CSV. (4) It detects markings and vocabulary only: no unmarked-CUI detection, no OCR, no
machine learning, and it is not a compliance certification. "No markings found" means "no
textual marking was found in the parts the scanner could decode", not "safe".

## Which tool

- **marking_search** (this release): a whole document store before an assessment. Answers "do
  we hold export-controlled or CUI-marked material, and which contract clauses bind us".
- **cui_scan** (coming in v0.2.0): one file or folder before it goes near an AI assistant or a
  cloud service. A fail-closed gate; see `docs/SPEC.md`, section 4.

## Quick start: marking_search

Run it on the machine that holds the documents, pointed at the synced SharePoint or OneDrive
root. In PowerShell, do not end the quoted path with a backslash.

```
python -m pip install pypdf cryptography        # optional; PDFs are counted as skipped without it
python marking_search/marking_search.py "C:\Synced\Company" --out "C:\somewhere\else\results.csv"
python marking_search/marking_search.py "C:\Synced\Company" --report
```

`--exclude HR Personal` skips folders by name anywhere in the tree. `--version` prints the
version. `--help` prints the full usage text. The CSV defaults to a file next to the script,
never inside the scanned folder.

## What the output means

The CSV has four columns: `path`, `category`, `hit_count`, `terms_matched`. `category` is one
of the five vocabulary categories below, or a `SKIPPED_*` or `ERROR*` note saying why a file
was not read. `terms_matched` holds only the fixed vocabulary that matched (for example
`ITAR; SP-EXPT`), never surrounding text. Cells are guarded against spreadsheet formula
injection. Do not send the CSV anywhere; it contains file paths.

The REPORT-BACK block is the part that is safe to paste into a ticket or an email. Line by
line:

- `root folder`: the name of the scanned folder only, not its path. `run`: timestamp. `pypdf`:
  whether PDFs could be read.
- `files on disk / scanned / skipped/errors`: every file seen; those read to the end; those not
  read, with reasons in the `not read` line. Office lock files (`~$...`) are ignored.
- `EXPORT_CONTROL`: ITAR, EAR, USML, ECCN, NOFORN, 22 CFR 120-130, 15 CFR 730-774,
  Distribution Statements D through F, "U.S. persons only", and similar.
- `CUI_SPECIFIED`: a `CUI//SP-XXXX` or `CONTROLLED//SP-XXXX` banner, or the words "CUI
  Specified".
- `CUI_ANY`: any CUI, FOUO, "Controlled Unclassified", Controlled Technical Information, or
  Covered Defense Information wording.
- `DFARS_CMMC`: DFARS 252.204-7008, -7012, -7019, -7020, -7021, 252.225-7048, 252.239-7010,
  NIST SP 800-171 or 800-171A, CMMC, SPRS.
- `FLOWDOWN`: "flow down" and its variants, counted only when the file also mentions clauses,
  DFARS, FAR, purchase orders, subcontracts, or suppliers, so HVAC and plumbing text does not
  count.
- `252.204-7019 found in N files`: YES or NO, with counts for the other clauses.
- `export markings found in N files`: split into strong (ITAR, SP-EXPT, USML, 22 CFR,
  Distribution Statement, NOFORN) and weak-only (EXPT, ECCN, EAR99, the phrase "export
  control"). Weak terms are mentions, not markings.
- `CUI//SP- categories seen`: which Specified categories appeared and in how many files.
- `not read`: unsupported types, PDFs without a text layer, PDFs skipped for lack of `pypdf`,
  encrypted PDFs, files over the size cap, cloud placeholders, password-protected Office files,
  legacy Office files with no byte-level hits, empty files, errors, folders that could not be
  listed, junctions, excluded folders, partial reads.
- `top folders by files with hits`: first two folder levels below the root, by count.

How to read the counts: strong `EXPORT_CONTROL` terms mean ITAR or EAR data is present.
`CUI//SP-EXPT`, ITAR, EAR, or NOFORN hits point toward a GCC High style environment; other CUI
categories do not by themselves. `252.204-7019` present means a SPRS score was already
required. Send the REPORT-BACK block, not the CSV.

Exit codes:

| code | meaning |
|---|---|
| 0 | scan completed |
| 2 | bad arguments, root is not a folder, output folder missing, or output path inside the scanned tree (nothing written) |
| 3 | scan ran but no file was readable |
| 130 | interrupted with Ctrl+C; the partial CSV and summary are still written |

## Sample output

Produced by `python tests/run_all.py`'s synthetic tree (12 files in four folders: plain text,
a UTF-16 text file, a Word file with a `CUI` header, an Excel file with an inline string, a
text PDF, an image-only PDF, and decoys). The folder name is replaced with `<root>`.

```
REPORT-BACK (safe to paste: counts only, no file paths, no document text)
  root folder: <root>   run: 2026-09-15 14:52   pypdf: yes
  files on disk: 12   scanned: 11   skipped/errors: 1   Office lock files ignored: 0
  EXPORT_CONTROL  files with hits: 5
  CUI_SPECIFIED   files with hits: 1
  CUI_ANY         files with hits: 2
  DFARS_CMMC      files with hits: 2
  FLOWDOWN        files with hits: 1
  252.204-7019 found in 1 files -> YES   (other clauses  7012: 1)
  export markings found in 5 files -> YES   (strong ITAR/SP-EXPT/USML/22 CFR/Dist Stmt/NOFORN: 5; weak-only EXPT/ECCN/EAR99/'export control': 0)
  CUI//SP- categories seen: SP-EXPT 1
  not read: unsupported types 0, PDFs without text layer 1, PDFs no pypdf 0, PDF encrypted/AES 0, over size cap 0, cloud placeholders 0, password-protected Office 0, legacy Office no hits 0 (with hits 0), empty 0, errors 0, folders not listable 0, junctions 0, excluded folders 0, partial reads 0
  top folders by files with hits (relative to root):
       5  Engineering
       3  Contracts
       1  General
```

## Known limits

This is regex over extracted text. It is not novel and it is not complete. Things it will miss
or get wrong:

- **AES-encrypted PDFs** need the `cryptography` package as well as `pypdf`. Without it they
  are counted as `SKIPPED_PDF_AES`, not read. Password-protected PDFs and Office files are never
  opened.
- **OneDrive online-only files** are counted as `SKIPPED_CLOUD_PLACEHOLDER` and never
  downloaded. Right-click the library and choose "Always keep on this device" first if they
  must be scanned.
- **Image-only PDFs** (scans with no text layer) are counted as `SKIPPED_PDF_NO_TEXT`. There
  is no OCR.
- **PDF text extraction** depends on the file: fonts without a ToUnicode map, ligatures,
  homoglyphs, form-field appearance streams, stamps, XFA forms, and attachments can all hide a
  marking. Letter-spaced title blocks ("I T A R") get a second, de-spaced pass.
- **Word all-caps formatting** (`w:caps`, text stored lowercase) and **VML watermarks**
  (`v:textpath string=`) are not read as markings.
- **Metadata and sensitivity labels** in PDF and Office files are not read, apart from PDF
  title, subject, and keywords.
- **Legacy Office files** (`.doc`, `.xls`, `.ppt`, `.msg`) get a byte-level string scan, not a
  parse. A miss there is reported as `SKIPPED_LEGACY`; open the file in Office to confirm.
- **Decoys**: a person surnamed Cui, the word EXPT, a resume that lists ITAR experience, and
  "flow down" in a maintenance note look the same to a string tool. Use `--exclude` for HR and
  personal folders and read the CSV with that in mind. "Export controls" in ordinary prose
  matches on purpose: "subject to U.S. export controls" is real marking language.
- **Size caps**: 60 MB for text and Office files, 512 MB and 2000 pages for PDFs. Files over
  the cap are counted, not read.
- **Container parts that fail to parse** are skipped and reported as partial reads, so a corrupt
  embedded object does not hide its host document.

None of these are claimed fixed. Each is a candidate for a later release.

## How it was tested

Five adversarial suites, run on Windows with Python 3.12 and pypdf 6.16.2, cryptography not
installed. Counts are from `marking_search/VERIFICATION.md`, which has the detail.

| suite | result after fixes |
|---|---|
| regex, 65 probes | ok 64, false neg 0, false pos 1 |
| formats, 82 fixtures | 72 OK, 8 accepted notes, 2 SKIPPED_PDF_AES (honest label, see below) |
| safety, 23 checks | PASS 23, FAIL 0 |
| operator, A to K | all pass; K: 6254 files in 17.9 s |
| winops, T1 to T12 | T2/T3 pass; legacy.doc found; T7 peak 150 MiB, 44 s |

The one remaining regex false positive is the "export controls" case above. The two AES PDFs
read correctly once `cryptography` is present. The suites and their fixtures are local and are
not in this repository; `tests/run_all.py` is the shipped, synthetic, end-to-end test, and CI
runs it on ubuntu and windows across Python 3.10, 3.12, 3.13, and 3.14, with and without
`pypdf`. Nothing from a real scan is in this repository.

## References

- ISOO, CUI Marking Handbook, v1.1
- DoDI 5200.48, Controlled Unclassified Information
- DFARS 252.204-7012, Safeguarding Covered Defense Information and Cyber Incident Reporting

## About

Built with AI assistance and reviewed by the author. Report a missed marking form privately
(see `SECURITY.md`); never attach a real document. MIT License.
