#!/usr/bin/env python3
"""End-to-end test for marking_search. Standard library only.

Builds a synthetic document tree in a temp dir, runs the tool against it with --out in a
sibling temp dir (never inside the tree), and checks: exit codes, per-category file counts
(branching on whether pypdf imports), that a sentinel planted in a Word body never reaches
stdout, stderr, or the CSV, that an --out path inside the tree is refused before anything is
written, and that the scanned tree is byte-for-byte untouched afterwards.

Every fixture is built here at runtime. Nothing is committed and nothing from a real scan is
used. Prints PASS/FAIL per check and exits non-zero on any failure.
"""
import csv
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
TOOL = os.path.join(REPO, "marking_search", "marking_search.py")
SENTINEL = "ZQXJ7731"

try:
    import pypdf  # noqa: F401
    HAVE_PYPDF = True
except ImportError:
    HAVE_PYPDF = False

RESULTS = []


def check(name, ok, detail=""):
    RESULTS.append(bool(ok))
    line = ("PASS  " if ok else "FAIL  ") + name
    if detail and not ok:
        line += "  (" + str(detail) + ")"
    print(line, flush=True)


# --- fixture builders ---------------------------------------------------------------------------
def write_text(path, text, encoding="utf-8"):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding=encoding, newline="\n") as fh:
        fh.write(text)


def write_bytes(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(data)


XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
NS_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' \
       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
NS_X = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' \
       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def make_docx(path):
    """Word file: header holds `CUI<tab>Page 1` split across runs; body holds the sentinel."""
    content_types = XML_DECL + (
        f'<Types xmlns="{CT_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        "</Types>")
    rels = XML_DECL + (
        f'<Relationships xmlns="{NS_PKG_REL}">'
        f'<Relationship Id="rId1" Type="{NS_DOC_REL}/officeDocument" Target="word/document.xml"/>'
        "</Relationships>")
    doc_rels = XML_DECL + (
        f'<Relationships xmlns="{NS_PKG_REL}">'
        f'<Relationship Id="rId1" Type="{NS_DOC_REL}/header" Target="header1.xml"/>'
        "</Relationships>")
    document = XML_DECL + (
        f"<w:document {NS_W}><w:body>"
        f"<w:p><w:r><w:t>Meeting notes for the widget project. Reference {SENTINEL} is the tracking code.</w:t></w:r></w:p>"
        '<w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr>'
        "</w:body></w:document>")
    header = XML_DECL + (
        f"<w:hdr {NS_W}><w:p>"
        "<w:r><w:t>CUI</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Page 1</w:t></w:r>"
        "</w:p></w:hdr>")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/header1.xml", header)


def make_xlsx(path):
    """Excel file with one inline-string cell carrying an ITAR marking."""
    content_types = XML_DECL + (
        f'<Types xmlns="{CT_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>")
    rels = XML_DECL + (
        f'<Relationships xmlns="{NS_PKG_REL}">'
        f'<Relationship Id="rId1" Type="{NS_DOC_REL}/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>")
    wb_rels = XML_DECL + (
        f'<Relationships xmlns="{NS_PKG_REL}">'
        f'<Relationship Id="rId1" Type="{NS_DOC_REL}/worksheet" Target="worksheets/sheet1.xml"/>'
        "</Relationships>")
    workbook = XML_DECL + (
        f"<workbook {NS_X}><sheets>"
        '<sheet name="Parts" sheetId="1" r:id="rId1"/>'
        "</sheets></workbook>")
    sheet = XML_DECL + (
        f"<worksheet {NS_X}><sheetData>"
        '<row r="1"><c r="A1" t="inlineStr"><is><t>ITAR restricted drawing list</t></is></c></row>'
        '<row r="2"><c r="A2" t="inlineStr"><is><t>Bracket, left hand</t></is></c></row>'
        "</sheetData></worksheet>")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/worksheets/sheet1.xml", sheet)


def make_pdf(objs):
    """Assemble a PDF from object bodies (object 1 must be the catalog). Correct xref offsets."""
    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for i, body in enumerate(objs, 1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode("latin-1") + body + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n".encode("latin-1")
    out += b"0000000000 65535 f \n"
    for o in offsets:
        out += f"{o:010d} 00000 n \n".encode("latin-1")
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode("latin-1")
    return bytes(out)


def stream_obj(dict_body, data):
    return b"<< " + dict_body + b" /Length " + str(len(data)).encode() + b" >>\nstream\n" + data + b"\nendstream"


def make_text_pdf(path):
    """One page, Helvetica, a single ITAR line. Read as EXPORT_CONTROL when pypdf is present."""
    content = b"BT /F1 14 Tf 72 720 Td (ITAR controlled technical data) Tj ET"
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
        b"/Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        stream_obj(b"", content),
    ]
    write_bytes(path, make_pdf(objs))


def make_image_only_pdf(path):
    """One page that draws a 1x1 grey image and has no text layer at all."""
    content = b"q 200 0 0 200 72 500 cm /Im1 Do Q"
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
        b"/Resources << /XObject << /Im1 4 0 R >> >> /Contents 5 0 R >>",
        stream_obj(b"/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8",
                   b"\x80"),
        stream_obj(b"", content),
    ]
    write_bytes(path, make_pdf(objs))


def build_tree(root):
    j = lambda *p: os.path.join(root, *p)
    write_text(j("Engineering", "itar_notice.txt"),
               "Technical data in this folder is subject to ITAR. Do not share outside the program.\n")
    write_text(j("Engineering", "banner_sp_expt.txt"),
               "CUI//SP-EXPT\nDrawing package index for the widget assembly.\nCUI//SP-EXPT\n")
    write_text(j("Engineering", "utf16_itar.txt"),
               "ITAR NOTICE: this drawing set is restricted to the program team.\n", encoding="utf-16")
    make_xlsx(j("Engineering", "parts.xlsx"))
    make_text_pdf(j("Engineering", "marked.pdf"))
    make_image_only_pdf(j("Engineering", "scan_only.pdf"))
    write_text(j("Contracts", "clause_7019.txt"),
               "This order incorporates DFARS 252.204-7019 by reference.\n")
    write_text(j("Contracts", "clause_7012.txt"),
               "Safeguarding requirements under 252.204-7012 apply to this subcontract.\n")
    write_text(j("Contracts", "flowdown_notice.txt"),
               "The supplier shall flow down these clauses to lower-tier subcontractors.\n")
    write_text(j("Facilities", "hvac_note.txt"),
               "Cool air will flow down from the ceiling diffusers when the HVAC unit cycles on.\n")
    make_docx(j("General", "header_marked.docx"))
    write_text(j("General", "team_lunch.txt"),
               "Team lunch is on Friday at noon. Bring a dish to share.\n")


# --- expectations -------------------------------------------------------------------------------
# Files that are read in full regardless of pypdf: 10 text/Office files. marked.pdf is read only
# with pypdf; scan_only.pdf is always skipped (no text layer, or no pypdf).
EXPECT = {
    "on_disk": 12,
    "scanned": 11 if HAVE_PYPDF else 10,
    "skipped": 1 if HAVE_PYPDF else 2,
    "EXPORT_CONTROL": 5 if HAVE_PYPDF else 4,
    "CUI_SPECIFIED": 1,
    "CUI_ANY": 2,
    "DFARS_CMMC": 2,
    "FLOWDOWN": 1,
    "7019": 1,
    "pdf_no_text": 1 if HAVE_PYPDF else 0,
    "pdf_no_pypdf": 0 if HAVE_PYPDF else 2,
}
CATEGORIES = ("EXPORT_CONTROL", "CUI_SPECIFIED", "CUI_ANY", "DFARS_CMMC", "FLOWDOWN")


def snapshot(root):
    snap = {}
    for dp, dns, fns in os.walk(root):
        dns.sort()
        for fn in fns:
            p = os.path.join(dp, fn)
            st = os.stat(p)
            snap[os.path.relpath(p, root)] = (st.st_size, st.st_mtime_ns)
    return snap


def run_tool(args):
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    return subprocess.run([sys.executable, TOOL] + list(args), capture_output=True,
                          encoding="utf-8", errors="replace", env=env, cwd=REPO)


def report_int(text, pattern):
    m = re.search(pattern, text, re.M)
    return int(m.group(1)) if m else None


def main():
    print(f"python {sys.version.split()[0]}  pypdf {'present' if HAVE_PYPDF else 'absent'}")
    root = tempfile.mkdtemp(prefix="cui_tree_")
    outdir = tempfile.mkdtemp(prefix="cui_out_")
    try:
        build_tree(root)
        before = snapshot(root)
        check("fixture tree has the expected file count", len(before) == EXPECT["on_disk"], len(before))

        # --version
        r = run_tool(["--version"])
        check("--version exits 0", r.returncode == 0, r.returncode)
        check("--version prints marking_search 0.1.0", r.stdout.strip() == "marking_search 0.1.0", r.stdout.strip())

        # full run with --out in the sibling temp dir
        csv_path = os.path.join(outdir, "results.csv")
        r1 = run_tool([root, "--out", csv_path])
        check("full run exits 0", r1.returncode == 0, f"rc={r1.returncode} stderr={r1.stderr[-300:]}")
        check("full run wrote the CSV", os.path.isfile(csv_path))

        rows = []
        if os.path.isfile(csv_path):
            with open(csv_path, encoding="utf-8-sig", newline="") as fh:
                rows = list(csv.reader(fh))
        check("CSV has the fixed header", rows[:1] == [["path", "category", "hit_count", "terms_matched"]], rows[:1])
        body = rows[1:]
        by_cat = {}
        for row in body:
            if len(row) >= 2:
                by_cat[row[1]] = by_cat.get(row[1], 0) + 1
        for cat in CATEGORIES:
            check(f"CSV rows for {cat} == {EXPECT[cat]}", by_cat.get(cat, 0) == EXPECT[cat], by_cat.get(cat, 0))
        skip_notes = [row[1] for row in body if len(row) >= 2 and row[1].startswith("SKIPPED_PDF")]
        if HAVE_PYPDF:
            check("image-only PDF is SKIPPED_PDF_NO_TEXT",
                  len(skip_notes) == 1 and skip_notes[0].startswith("SKIPPED_PDF_NO_TEXT"), skip_notes)
        else:
            check("both PDFs are SKIPPED_PDF (no pypdf)",
                  len(skip_notes) == 2 and all(n.startswith("SKIPPED_PDF (") for n in skip_notes), skip_notes)
        check("CSV rows all start with the scanned root",
              all(os.path.normcase(row[0]).startswith(os.path.normcase(root)) for row in body if row), "")
        check("no ERROR rows in the CSV", not any(row[1].startswith("ERROR") for row in body if len(row) >= 2),
              [row[1] for row in body if len(row) >= 2 and row[1].startswith("ERROR")])

        # report-only run, second CSV, also in the sibling dir
        csv2 = os.path.join(outdir, "report_run.csv")
        r2 = run_tool([root, "--out", csv2, "--report"])
        check("--report run exits 0", r2.returncode == 0, f"rc={r2.returncode} stderr={r2.stderr[-300:]}")
        rep = r2.stdout
        check("--report stdout starts with the REPORT-BACK line", rep.startswith("REPORT-BACK"), rep[:60])
        check("--report stdout has no file paths", root not in rep and os.path.basename(outdir) not in rep, "")
        check("report: files on disk", report_int(rep, r"files on disk: (\d+)") == EXPECT["on_disk"],
              report_int(rep, r"files on disk: (\d+)"))
        check("report: scanned", report_int(rep, r"scanned: (\d+)") == EXPECT["scanned"],
              report_int(rep, r"scanned: (\d+)"))
        check("report: skipped/errors", report_int(rep, r"skipped/errors: (\d+)") == EXPECT["skipped"],
              report_int(rep, r"skipped/errors: (\d+)"))
        for cat in CATEGORIES:
            got = report_int(rep, rf"^\s+{cat}\s+files with hits: (\d+)$")
            check(f"report: {cat} files with hits == {EXPECT[cat]}", got == EXPECT[cat], got)
        got = report_int(rep, r"252\.204-7019 found in (\d+) files")
        check("report: 252.204-7019 count", got == EXPECT["7019"], got)
        check("report: 252.204-7019 flag is YES", re.search(r"252\.204-7019 found in \d+ files -> YES", rep) is not None, "")
        check("report: other clause 7012 counted", re.search(r"other clauses\s+7012: 1", rep) is not None, "")
        check("report: SP-EXPT category seen", "CUI//SP- categories seen: SP-EXPT 1" in rep, "")
        got = report_int(rep, r"PDFs without text layer (\d+)")
        check("report: PDFs without text layer", got == EXPECT["pdf_no_text"], got)
        got = report_int(rep, r"PDFs no pypdf (\d+)")
        check("report: PDFs no pypdf", got == EXPECT["pdf_no_pypdf"], got)
        check("report: no interpretation lines on stdout",
              "GCC High" not in r1.stdout and "SPRS score" not in r1.stdout and "GCC High" not in rep, "")

        # sentinel: the Word body text must never leave the tool
        csv_text = ""
        for p in (csv_path, csv2):
            if os.path.isfile(p):
                with open(p, encoding="utf-8-sig") as fh:
                    csv_text += fh.read()
        check("sentinel absent from both CSVs", SENTINEL not in csv_text)
        check("sentinel absent from stdout", SENTINEL not in r1.stdout and SENTINEL not in r2.stdout)
        check("sentinel absent from stderr", SENTINEL not in r1.stderr and SENTINEL not in r2.stderr)
        check("docx header marking was found via the CSV",
              any(row[0].endswith("header_marked.docx") and row[1] == "CUI_ANY" for row in body if len(row) >= 2), "")

        # refused output path: inside the tree, at the root and in a subfolder
        for rel in ("inside.csv", os.path.join("General", "inside.csv")):
            target = os.path.join(root, rel)
            r3 = run_tool([root, "--out", target])
            check(f"--out {rel} inside the tree exits 2", r3.returncode == 2, r3.returncode)
            check(f"--out {rel} inside the tree writes nothing", not os.path.exists(target))

        # the tree is untouched
        after = snapshot(root)
        check("scanned tree file list unchanged", sorted(before) == sorted(after),
              sorted(set(before) ^ set(after)))
        check("scanned tree sizes and mtimes unchanged", before == after,
              [k for k in before if before.get(k) != after.get(k)])
    finally:
        shutil.rmtree(root, ignore_errors=True)
        shutil.rmtree(outdir, ignore_errors=True)

    passed, failed = RESULTS.count(True), RESULTS.count(False)
    print(f"{passed} passed, {failed} failed")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
