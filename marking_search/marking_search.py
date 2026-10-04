"""
marking_search.py  --  read-only scan for export-control / CUI / DFARS markings.

Run it on the machine that holds the documents, pointed at the synced SharePoint or OneDrive
root. It walks every folder, pulls text out of Office / PDF / email / text-like files, and
writes ONE csv (default: next to this script, never inside the scanned folder): file path,
category, hit count, which fixed vocabulary terms matched. It never copies document text
anywhere and never writes into the scanned tree. Results stay on that machine.

Usage (PowerShell; do NOT end the quoted path with a backslash):
    python marking_search.py "C:\\Synced\\Company"
    python marking_search.py "C:\\Synced\\Company" --exclude HR Personal
    python marking_search.py "C:\\Synced\\Company" --report     # paste-safe block only
    python marking_search.py "C:\\Synced\\Company" --out "C:\\somewhere\\else\\results.csv"
    python marking_search.py --version

Needs only Python 3. PDFs are read if pypdf is installed:
    python -m pip install pypdf cryptography
If pip is blocked, run anyway: PDFs are listed as SKIPPED_PDF and counted, so the report
still says how many files were not read.

What to send back: the REPORT-BACK block (counts only). Do not send the CSV; it contains
file paths. OneDrive "online-only" files are NOT downloaded by this tool; they are counted as
SKIPPED_CLOUD_PLACEHOLDER. Right-click the library > "Always keep on this device" first if
those must be scanned.

Exit codes: 0 ok, 2 bad arguments / root / output path, 3 ran but nothing was readable,
130 interrupted with Ctrl+C (partial CSV and summary are still written).
"""
import argparse, collections, csv, email, email.policy, html, io, logging, os, re, sys, time, zipfile

__version__ = "0.1.0"

# pypdf logs fragments of malformed files at WARNING/ERROR. Nothing from a document may reach
# stdout or stderr, so its logger gets a NullHandler, level CRITICAL, and no propagation.
_pypdf_log = logging.getLogger("pypdf")
_pypdf_log.addHandler(logging.NullHandler())
_pypdf_log.setLevel(logging.CRITICAL)
_pypdf_log.propagate = False

PATTERNS = {
    "EXPORT_CONTROL": re.compile(
        r"(?-i:\bSP-EXPT\b|\bEXPT\b|\bNOFORN\b|\bREL\s+TO\s+USA\b)|"
        r"\bITAR\b|\bI\.T\.A\.R\b\.?|International\s+Traffic\s+in\s+Arms|"
        r"Export\s+Administration\s+Regulations|\bexport[\s-]*controll?(?:ed|s)?\b|\bEAR[\s-]+controlled\b|"
        r"22\s*C\.?F\.?R\.?\s*(?:Parts?\s*)?1[23]\d(?!\d)|15\s*C\.?F\.?R\.?\s*(?:Parts?\s*)?7[3-7]\d(?!\d)|"
        r"Arms\s+Export\s+Control\s+Act|\bAECA\b|\bUSML\b|\bECCN\b|\bEAR\s*99\b|"
        r"Distribution\s+Statement\s+[D-F]\b|\bU\.?S\.?\s+persons?\s+only\b|\bdeemed\s+exports?\b", re.I),
    "CUI_SPECIFIED": re.compile(
        r"(?-i:\b(?:CUI|CONTROLLED)\s*//\s*SP-(?:[A-Z]{2,7}\b)?)|\bCUI\s+Specified\b", re.I),
    "CUI_ANY": re.compile(
        r"(?-i:\bCUI\b|\bCONTROLLED\s*//|\bFOUO\b)|CONTROLLED\s+UNCLASSIFIED|FOR\s+OFFICIAL\s+USE\s+ONLY|"
        r"Controlled\s+Technical\s+Information|Covered\s+Defense\s+Information", re.I),
    "DFARS_CMMC": re.compile(
        r"252\.204[-\s]*70(?:08|12|19|20|21)\b|252\.225[-\s]*7048\b|252\.239[-\s]*7010\b|"
        r"(?<![\w.-])(?:NIST\s*)?(?:SP\s*)?800[-\s]?171A?\b(?![-\d])|\bCMMC\b|\bSPRS\b", re.I),
    "FLOWDOWN": re.compile(r"\bflow(?:s|ed|ing)?[\s-]*downs?\b", re.I),
}
# two-word "flow down" is ordinary English (HVAC, plumbing, finance); count it only when the
# document also talks about contract clauses. Checked against the text, never written anywhere.
FLOW_CTX = re.compile(r"\bclaus\w*|\bDFARS\b|(?-i:\bFAR\b|\bPOs?\b)|\bsubcontract\w*|\bsuppliers?\b|"
                      r"\bpurchase\s+orders?\b|\b252\.2\d\d\b|\bflow-?downs?\b", re.I)
WEAK_EXPORT = {"EXPT", "ECCN", "EAR99", "EAR 99", "EXPORT CONTROL", "EXPORT CONTROLS",
               "DEEMED EXPORT", "DEEMED EXPORTS"}   # mentions, not markings

TEXT_EXT = {".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".rtf", ".log",
            ".text", ".ini", ".yaml", ".yml"}
MAIL_EXT = {".eml"}
WORD_EXT = {".docx", ".docm", ".dotx", ".dotm"}
XL_EXT = {".xlsx", ".xlsm", ".xltx", ".xltm"}
PPT_EXT = {".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm"}
ODF_EXT = {".odt", ".ods", ".odp", ".ott", ".ots", ".otp"}
OOXML_EXT = WORD_EXT | XL_EXT | PPT_EXT
ZIPDOC_EXT = OOXML_EXT | ODF_EXT
LEGACY_EXT = {".doc", ".dot", ".xls", ".xlt", ".ppt", ".pot", ".msg"}     # OLE2 binary
KNOWN_EXT = TEXT_EXT | MAIL_EXT | ZIPDOC_EXT | LEGACY_EXT | {".pdf", ".zip"}
MAX_BYTES = 60 * 1024 * 1024           # text / binary files read whole
PDF_MAX_BYTES = 512 * 1024 * 1024      # pypdf reads lazily
PDF_PAGE_CAP = 2000
ZIP_BUDGET = 4 * MAX_BYTES             # total uncompressed bytes read out of one package
OLE_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")
ENC_MARK = "EncryptedPackage".encode("utf-16-le")
PLACEHOLDER_ATTRS = 0x1000 | 0x40000 | 0x400000    # OFFLINE | RECALL_ON_OPEN | RECALL_ON_DATA_ACCESS
REPARSE_LINK_TAGS = (0xA0000003, 0xA000000C)      # junction, symlink (cloud reparse tags are NOT pruned)
EXT_PREFIX = "\\\\?\\"

# --- in-memory normalisation before matching (nothing normalised is ever written) --------------
_NORM = {ord(c): "-" for c in "\u2010\u2011\u2012\u2013\u2014\u2212"}
_NORM.update({0x00AD: None, 0x200B: None, 0x200C: None, 0x200D: None, 0x2060: None, 0xFEFF: None, 0: None})

def normalize(text):
    if text.isascii() and "\0" not in text:    # nothing to fold: skip the second full-size copy (59 MB text files)
        return text
    return text.translate(_NORM)

def decode_bytes(raw):
    if raw[:3] == b"\xef\xbb\xbf":
        return raw[3:].decode("utf-8", "replace")
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16", "replace")
    sample = raw[:4000]
    if len(sample) >= 4:
        odd, even = sample[1::2].count(0), sample[0::2].count(0)
        if odd > len(sample) // 4 and even < len(sample) // 20:
            return raw.decode("utf-16-le", "replace")
        if even > len(sample) // 4 and odd < len(sample) // 20:
            return raw.decode("utf-16-be", "replace")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp1252", "replace")

# --- XML / HTML / RTF -> text ----------------------------------------------------------------
TAG = re.compile(r"<[^>]+>")
XML_BREAK = re.compile(r"</(?:w:p|a:p|si|c|f|row|w:tc|text:p|text:h|p:text|dc:[^>]+|cp:[^>]+|dcterms:[^>]+|vt:[^>]+)>|"
                       r"<w:(?:br|cr)\b[^>]*/>|<a:br\b[^>]*/>|<text:line-break\s*/>")
XML_SPACE = re.compile(r"<w:tab\b[^>]*/>|<text:tab\s*/>|<text:s\b[^>]*/>")
XML_NBH = re.compile(r"<w:noBreakHyphen\s*/>")
XML_SOFTH = re.compile(r"<w:softHyphen\s*/>")
XML_FALLBACK = re.compile(r"<mc:Fallback>.*?</mc:Fallback>", re.S)   # VML copy of a text box
XML_HF = re.compile(r"(<(?:odd|even|first)(?:Header|Footer)>)(.*?)(</(?:odd|even|first)(?:Header|Footer)>)", re.S)
HF_CODE = re.compile(r'&amp;(?:"[^"]*"|K[0-9A-Fa-f]{6}|\d+|[A-Za-z])')   # Excel &C &L &R &"font" codes
XML_NAME_ATTR = re.compile(r'<(?:sheet|p:cSld|definedName|table)\b[^>]*?\sname="([^"]*)"')
HTML_BLOCK = re.compile(r"<(?:br|hr|p|div|tr|li|/p|/div|/td|/th|/tr|/li|/h[1-6]|/title|/pre|/blockquote|/table|/dd|/dt|/option)\b[^>]*>", re.I)
HTML_DROP = re.compile(r"<(script|style)\b.*?</\1\s*>", re.I | re.S)

def xml_text(xml, joined=True):
    """joined=True: Office body parts. Runs inside one paragraph/cell are glued back together
    (Word splits words across <w:r> runs after edits); breaks only at paragraph/cell ends."""
    xml = XML_FALLBACK.sub("", xml)
    xml = XML_HF.sub(lambda m: m.group(1) + HF_CODE.sub(" ", m.group(2)) + m.group(3), xml)
    names = " ".join(XML_NAME_ATTR.findall(xml))
    xml = XML_NBH.sub("-", xml)
    xml = XML_SOFTH.sub("", xml)
    if joined:
        xml = XML_SPACE.sub(" ", xml)
        xml = XML_BREAK.sub("\n", xml)
        xml = TAG.sub("", xml)
    else:
        xml = TAG.sub(" ", xml)
    return html.unescape(xml + "\n" + names)

def html_text(s):
    s = HTML_DROP.sub(" ", s)
    s = HTML_BLOCK.sub("\n", s)
    return html.unescape(TAG.sub("", s))

def rtf_text(s):
    s = re.sub(r"\{\\\*(?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*\}", "", s)             # {\* ...} destinations
    s = re.sub(r"\\u(-?\d+)\s?(?:\\'[0-9a-fA-F]{2}|[^\\{}\r\n])?",
               lambda m: chr(int(m.group(1)) % 65536), s)                          # \uN<fallback>
    s = re.sub(r"\\'([0-9a-fA-F]{2})",
               lambda m: bytes([int(m.group(1), 16)]).decode("cp1252", "replace"), s)   # \'hh
    s = re.sub(r"\\(?:par|line|tab|cell|row|sect|page)\b", "\n", s)
    s = re.sub(r"\\[a-zA-Z]+-?\d* ?", "", s)                                        # other control words
    s = re.sub(r"\\([{}\\~_-])", lambda m: {"~": "\xa0", "_": "-", "-": ""}.get(m.group(1), m.group(1)), s)
    return s.replace("{", "").replace("}", "")

def ole_raw_text(data):
    """Word/Excel/PowerPoint 97-2003 and Outlook .msg: no parser in the stdlib, so read the
    strings straight out of the container in both encodings Office uses."""
    body = data[8:]      # drop the OLE magic: its bytes decode to letters that glue onto adjacent text and defeat \b
    views = [body.replace(b"\xff", b" ").replace(b"\xfe", b" ").decode("cp1252", "replace"),   # FAT fill bytes, same reason
             body.decode("utf-16-le", "replace"), body[1:].decode("utf-16-le", "replace")]
    return "\n".join(v.replace("\x00", " ") for v in views)

# --- extraction ------------------------------------------------------------------------------
def zip_text(z, depth, extra):
    out = []
    for zi in z.infolist():
        if zi.is_dir():
            continue
        low = zi.filename.lower()
        ext = os.path.splitext(low)[1]
        is_xml = low.endswith(".xml")
        nested = ext in ZIPDOC_EXT or ext == ".zip"
        other = ext in TEXT_EXT or ext in LEGACY_EXT or ext in MAIL_EXT or ext == ".pdf" or \
            (ext == ".bin" and "embeddings/" in low)
        if not (is_xml or nested or other):
            continue
        if zi.file_size > MAX_BYTES or extra.get("_budget", ZIP_BUDGET) - zi.file_size < 0:
            extra["partial"] = "PARTIAL_PACKAGE (a member over the size budget was not read)"
            continue
        extra["_budget"] = extra.get("_budget", ZIP_BUDGET) - zi.file_size
        try:
            data = z.read(zi)
        except Exception:
            extra["partial"] = "PARTIAL_PACKAGE (a member could not be read)"
            continue
        if is_xml:
            joined = low.startswith(("word/", "xl/", "ppt/")) or low in ("content.xml", "styles.xml", "meta.xml")
            out.append(xml_text(data.decode("utf-8", "replace"), joined))
        elif depth < 2:
            try:
                t, _note, sub = extract_bytes(data, ext, depth + 1)
                if t:
                    out.append(t)
                if sub.get("partial"):
                    extra["partial"] = sub["partial"]
            except Exception:
                pass    # a corrupt embedded object should not hide the host document
    return "\n".join(out)

def eml_text(raw, depth=0):
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    parts = [str(msg.get("subject", "") or "")]
    for part in msg.walk():
        if part.is_multipart():
            continue
        ctype = part.get_content_maintype()
        fname = part.get_filename() or ""
        ext = os.path.splitext(fname.lower())[1]
        if ctype == "text" and not fname:
            try:
                body = part.get_content()
            except Exception:
                body = decode_bytes(part.get_payload(decode=True) or b"")
            if not isinstance(body, str):
                body = decode_bytes(body if isinstance(body, bytes) else b"")
            parts.append(html_text(body) if part.get_content_subtype() == "html" else body)
        elif ext in KNOWN_EXT and depth < 2:
            try:
                t, _n, _e = extract_bytes(part.get_payload(decode=True) or b"", ext, depth + 1)
                if t:
                    parts.append(t)
            except Exception:
                pass
    return "\n".join(parts)

def pdf_text(fh):
    try:
        from pypdf import PdfReader
        from pypdf.errors import DependencyError, FileNotDecryptedError
    except ImportError:
        return None, "SKIPPED_PDF (pypdf not installed: python -m pip install pypdf)", {}
    extra = {}
    try:
        try:
            r = PdfReader(fh)
        except DependencyError:
            return None, "SKIPPED_PDF_AES (AES-encrypted; python -m pip install cryptography)", {}
        n = len(r.pages)
        texts, bad = [], 0
        for i, pg in enumerate(r.pages):
            if i >= PDF_PAGE_CAP:
                extra["partial"] = f"PARTIAL_PDF (pages {PDF_PAGE_CAP + 1}-{n} not read)"
                break
            try:
                texts.append(pg.extract_text() or "")
            except (FileNotDecryptedError, DependencyError):   # AES-128 only fails here, not in PdfReader()
                raise
            except Exception:
                bad += 1
        if bad and bad == n:
            raise ValueError("no page could be read")
        text = "\n".join(texts)
        try:                                   # fillable form values (DD254s, shipping declarations)
            fields = r.get_fields() or {}
            text += "\n" + "\n".join(str(v.get("/V", "")) for v in fields.values() if hasattr(v, "get"))
        except Exception:
            pass
        try:
            md = r.metadata
            if md:
                text += "\n" + " ".join(str(x) for x in (md.title, md.subject, md.keywords) if x)
        except Exception:
            pass
    except DependencyError:
        return None, "SKIPPED_PDF_AES (AES-encrypted; python -m pip install cryptography)", {}
    except FileNotDecryptedError:
        return None, "SKIPPED_PDF_ENCRYPTED (password required; not read)", {}
    if len(re.findall(r"\w", text)) < 10 * max(1, n):
        extra["thin"] = True                   # probably a scan with no text layer; main() decides
    # letter-spaced title blocks come out of pypdf as 'IT AR': try a de-spaced copy too
    extra["alt"] = re.sub(r"\b([A-Z]{1,3}) (?=[A-Z]{1,3}\b)", r"\1", text)
    return text, None, extra

def extract_bytes(data, ext, depth=0):
    head = data[:8]
    if head.startswith(b"PK\x03\x04") and (ext in ZIPDOC_EXT or ext in LEGACY_EXT or ext == ".zip"):
        extra = {}
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            return zip_text(z, depth, extra), None, extra
    if head == OLE_MAGIC:
        if ENC_MARK in data:
            return None, "SKIPPED_ENCRYPTED_OFFICE (password-protected Office file; not read)", {}
        if ext in LEGACY_EXT or ext in OOXML_EXT or ext == ".bin":
            return ole_raw_text(data), None, {"legacy": True}
        return None, None, {}
    if ext == ".pdf" or head.startswith(b"%PDF"):
        return pdf_text(io.BytesIO(data))
    if ext in ZIPDOC_EXT or ext == ".zip":
        with zipfile.ZipFile(io.BytesIO(data)) as z:   # wrong magic: let BadZipFile surface
            extra = {}
            return zip_text(z, depth, extra), None, extra
    if ext in MAIL_EXT:
        return eml_text(data, depth), None, {}
    if ext == ".rtf" or head.startswith(b"{\\rtf"):
        return rtf_text(decode_bytes(data)), None, {}
    if ext in (".html", ".htm"):
        return html_text(decode_bytes(data)), None, {}
    if ext == ".xml":
        return html.unescape(TAG.sub(" ", decode_bytes(data))), None, {}
    if ext in TEXT_EXT:
        return decode_bytes(data), None, {}
    if ext in LEGACY_EXT:                      # .doc that is really RTF/plain text
        return decode_bytes(data), None, {"legacy": True}
    return None, None, {}

def extract_full(path, ext=None, size=None):
    """-> (text, note, extra). text None + note = not read (reason); both None = unsupported."""
    if ext is None:
        ext = os.path.splitext(path)[1].lower()
    if size is None:
        size = os.path.getsize(path)
    if size == 0:
        return None, "SKIPPED_EMPTY (0 bytes)", {}
    with open(path, "rb") as fh:
        head = fh.read(8)
        if head.startswith(b"PK\x03\x04") and (ext in ZIPDOC_EXT or ext in LEGACY_EXT or ext == ".zip"):
            fh.seek(0)
            extra = {}
            with zipfile.ZipFile(fh) as z:
                return zip_text(z, 0, extra), None, extra
        if ext == ".pdf" or head.startswith(b"%PDF"):
            if size > PDF_MAX_BYTES:
                return None, f"SKIPPED_SIZE ({size >> 20} MB > {PDF_MAX_BYTES >> 20} MB cap)", {}
            fh.seek(0)
            return pdf_text(fh)
        if size > MAX_BYTES:
            return None, f"SKIPPED_SIZE ({size >> 20} MB > {MAX_BYTES >> 20} MB cap)", {}
        fh.seek(0)
        data = fh.read()
    return extract_bytes(data, ext, 0)

def extract(path, ext=None, size=None):
    """Compatibility wrapper: (text, note)."""
    text, note, _extra = extract_full(path, ext, size)
    return text, note

# --- Windows helpers -------------------------------------------------------------------------
def extended(p):
    if os.name != "nt" or p.startswith(EXT_PREFIX):
        return p
    if p.startswith("\\\\"):
        return EXT_PREFIX + "UNC\\" + p[2:]
    return EXT_PREFIX + p

def display(p):
    if p.startswith(EXT_PREFIX + "UNC\\"):
        return "\\\\" + p[8:]
    if p.startswith(EXT_PREFIX):
        return p[4:]
    return p

def is_link_dir(p):
    try:
        st = os.lstat(p)
    except OSError:
        return False
    if not getattr(st, "st_file_attributes", 0) & 0x400:      # FILE_ATTRIBUTE_REPARSE_POINT
        return False
    return getattr(st, "st_reparse_tag", 0) in REPARSE_LINK_TAGS

def guard(cell):
    s = str(cell)
    return "'" + s if s[:1] in ("=", "+", "-", "@", "\t", "\r") else s

def norm_terms(hits):
    seen = collections.Counter(re.sub(r"\s+", " ", h).strip().upper() for h in hits if isinstance(h, str))
    seen.pop("", None)
    return sorted(seen, key=lambda t: (t[0].isdigit(), -seen[t], t))[:12]

def die(msg):
    print("ERROR: " + msg, file=sys.stderr)
    sys.exit(2)

# --- main --------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("root", help='folder to scan, e.g. "C:\\Synced\\Company" (no trailing backslash inside the quotes)')
    ap.add_argument("--out", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "marking_search_results.csv"),
                    help="results CSV (default: next to this script; must be outside the scanned folder)")
    ap.add_argument("--exclude", nargs="*", default=[], metavar="NAME",
                    help="folder names to skip anywhere in the tree, case-insensitive, e.g. --exclude HR Personal")
    ap.add_argument("--report", action="store_true",
                    help="print only the paste-safe REPORT-BACK block (counts, no file paths)")
    ap.add_argument("--version", action="version", version=f"marking_search {__version__}")
    a = ap.parse_args()

    if '"' in a.root:
        die("the folder path contains a quote character. In Windows PowerShell a backslash before the closing\n"
            '       quote escapes it ("C:\\path\\" swallows the rest of the line). Remove the trailing backslash.')
    root = display(os.path.abspath(a.root))     # plain form for the CSV and summary; extended() re-adds \\?\ for every OS call
    if not os.path.isdir(extended(root)):
        die(f"not a folder: {a.root!r}")
    out = os.path.abspath(a.out)
    if not os.path.isdir(os.path.dirname(out)):
        die(f"output folder does not exist: {os.path.dirname(out)}")
    try:
        inside = os.path.commonpath([os.path.normcase(out), os.path.normcase(root)]) == os.path.normcase(root)
    except ValueError:
        inside = False
    if inside:
        die("the results file would be written INSIDE the scanned folder (it would sync to SharePoint and be\n"
            "       re-scanned next run). Use --out with a path outside it, or leave --out off.")
    for var in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        od = os.environ.get(var)
        if od and os.path.normcase(out).startswith(os.path.normcase(os.path.abspath(od)) + os.sep):
            print(f"WARNING: results file is inside a OneDrive-synced folder ({var}) and will sync upward.", file=sys.stderr)
            break
    if os.path.splitdrive(root)[1] in ("\\", "/") or os.path.normcase(root) == os.path.normcase(os.path.expanduser("~")):
        print("WARNING: that root is a whole drive or your whole profile. This will take a long time; "
              "point at the synced company folder if that is what you meant.", file=sys.stderr)
    excl = {e.lower() for e in a.exclude}
    try:
        f = open(out, "w", newline="", encoding="utf-8-sig")
    except OSError as e:
        die(f"cannot write results file {out}: {e}")
    try:
        import pypdf  # noqa: F401
        pypdf_ok = f"yes ({getattr(pypdf, '__version__', '?')})"
    except ImportError:
        pypdf_ok = "NO - PDFs will be listed as SKIPPED_PDF"
    if not a.report:
        print(f"scanning: {root}", file=sys.stderr)
        print(f"pypdf: {pypdf_ok}", file=sys.stderr)

    C = collections.Counter()               # every counter lives here
    totals = {k: 0 for k in PATTERNS}
    unsupported = collections.Counter()
    folder_hits = collections.Counter()
    dfars_files = collections.Counter()
    sp_files = collections.Counter()
    out_norm = os.path.normcase(out)
    w = csv.writer(f)
    w.writerow(["path", "category", "hit_count", "terms_matched"])
    t_start = time.time()
    last_progress = [t_start]

    def row(path, cat, n=0, terms=""):
        w.writerow([guard(path), guard(cat), n, guard(terms)])
        C["rows"] += 1
        if C["rows"] % 50 == 0:
            f.flush()

    def on_dir_error(e):
        row(display(getattr(e, "filename", "") or "?"), f"ERROR_DIR {type(e).__name__} (folder could not be listed)")
        C["skipped"] += 1
        C["dir_errors"] += 1

    def progress(force=False):
        now = time.time()
        if force or C["on_disk"] % 500 == 0 or now - last_progress[0] > 10:
            last_progress[0] = now
            print(f"  ... {C['on_disk']} files seen, {C['scanned']} scanned, {C['skipped']} skipped, "
                  f"{C['hit_files']} with hits ({int(now - t_start)} s)", file=sys.stderr, flush=True)

    root_nc = os.path.normcase(root)

    def top_folder(path):
        """first two folder levels below root, for the summary. String math only: os.path.relpath
        refuses to mix extended-length and plain forms and cannot resolve >260-char paths."""
        d, n = os.path.dirname(path), len(root)
        if os.path.normcase(d[:n]) == root_nc and (root[-1:] in "\\/" or d[n:n + 1] in ("", "\\", "/")):
            d = d[n:]
        rel = d.strip("\\/").replace("\\", "/")
        return "/".join(rel.split("/")[:2]) if rel else "(root)"

    def scan_file(p, fn, disp):
        try:
            st = os.stat(p, follow_symlinks=False)
        except OSError as e:
            row(disp, f"ERROR {type(e).__name__}")
            C["skipped"] += 1
            C["errors"] += 1
            return
        if getattr(st, "st_file_attributes", 0) & PLACEHOLDER_ATTRS:
            row(disp, "SKIPPED_CLOUD_PLACEHOLDER (online-only file; not downloaded)")
            C["skipped"] += 1
            C["placeholders"] += 1
            return
        ext = os.path.splitext(fn)[1].lower()
        if ext not in KNOWN_EXT:
            unsupported[ext or "(none)"] += 1
            C["skipped"] += 1
            C["unsupported"] += 1
            return
        try:
            text, note, extra = extract_full(p, ext, st.st_size)
        except KeyboardInterrupt:
            raise
        except Exception as e:      # locked, corrupt, truncated
            text, note, extra = None, f"ERROR {type(e).__name__}", {}
        if note:
            row(disp, note)
            C["skipped"] += 1
            key = note.split(" ")[0]
            C[key] += 1
            if key.startswith("ERROR"):
                C["errors"] += 1
            return
        if text is None:
            unsupported[ext] += 1
            C["skipped"] += 1
            C["unsupported"] += 1
            return
        text = normalize(text)
        alt = normalize(extra["alt"]) if extra.get("alt") else None
        found = {}
        for cat, rx in PATTERNS.items():
            hits = rx.findall(text)
            if not hits and alt:
                hits = rx.findall(alt)
            if cat == "FLOWDOWN" and hits and all(re.search(r"\s", h) for h in hits) and not FLOW_CTX.search(text):
                hits = []
            if hits:
                found[cat] = hits
        if not found and extra.get("thin"):
            row(disp, "SKIPPED_PDF_NO_TEXT (no text layer, probably a scan; OCR needed)")
            C["skipped"] += 1
            C["pdf_no_text"] += 1
            return
        if not found and extra.get("legacy"):
            row(disp, "SKIPPED_LEGACY (Office 97-2003/Outlook binary; byte-level scan found nothing, open in Office to confirm)")
            C["skipped"] += 1
            C["legacy_nohit"] += 1
            return
        if extra.get("legacy"):
            C["legacy_hits"] += 1
        if extra.get("partial"):
            row(disp, extra["partial"])
            C["partial"] += 1
        if found:
            C["hit_files"] += 1
            folder_hits[top_folder(disp)] += 1
        for cat, hits in found.items():
            totals[cat] += 1
            terms = norm_terms(hits)
            row(disp, cat, len(hits), "; ".join(terms))
            if cat == "DFARS_CMMC":
                for t in {re.sub(r"\s+", "", x) for x in terms}:
                    m = re.fullmatch(r"252\.(\d{3})-(\d{4})", t)
                    if m:
                        dfars_files[f"252.{m.group(1)}-{m.group(2)}"] += 1
            elif cat == "CUI_SPECIFIED":
                cats = {m.group(1) for m in (re.search(r"SP-([A-Z]{2,7})\b", x) for x in terms) if m}
                for c_ in cats or {"(category not captured)"}:
                    sp_files["SP-" + c_ if not c_.startswith("(") else c_] += 1
            elif cat == "EXPORT_CONTROL":
                strong = any(re.sub(r"[\s-]+", " ", x) not in WEAK_EXPORT for x in terms)
                C["export_strong" if strong else "export_weak"] += 1
        C["scanned"] += 1           # last, so a failure above cannot count a file as both scanned and skipped

    partial = False
    try:
        for dirpath, dirs, files in os.walk(extended(root), onerror=on_dir_error):
            keep = []
            for d in dirs:
                full = os.path.join(dirpath, d)
                if d.lower() in excl:
                    C["excluded_dirs"] += 1
                elif is_link_dir(full):
                    row(display(full), "SKIPPED_JUNCTION (folder link/junction not followed)")
                    C["skipped"] += 1
                    C["junctions"] += 1
                else:
                    keep.append(d)
            dirs[:] = keep
            for fn in files:
                p = os.path.join(dirpath, fn)
                disp = display(p)
                C["on_disk"] += 1
                progress()
                if fn.startswith("~$"):
                    C["lockfiles"] += 1
                    continue
                if os.path.normcase(disp) == out_norm:
                    continue
                try:
                    scan_file(p, fn, disp)
                except KeyboardInterrupt:
                    raise
                except Exception as e:      # one odd file must never end the whole run (no CSV, no summary)
                    row(disp, f"ERROR {type(e).__name__} (unexpected; file not fully processed)")
                    C["skipped"] += 1
                    C["errors"] += 1
    except KeyboardInterrupt:
        partial = True
    finally:
        try:
            f.flush()
            f.close()
        except OSError:
            pass
    progress(force=True)

    # ---- summary -------------------------------------------------------------------------------
    scanned, skipped = C["scanned"], C["skipped"]
    if not a.report:
        print(("PARTIAL RUN (interrupted) -- " if partial else "") +
              f"files scanned: {scanned}   skipped/errors: {skipped}   results: {out}")
    yn = lambda n: "YES" if n else "NO"
    lines = ["REPORT-BACK (safe to paste: counts only, no file paths, no document text)" + ("  ** PARTIAL RUN **" if partial else ""),
             f"  root folder: {os.path.basename(root) or root}   run: {time.strftime('%Y-%m-%d %H:%M')}   pypdf: {pypdf_ok.split(' ')[0]}",
             f"  files on disk: {C['on_disk']}   scanned: {scanned}   skipped/errors: {skipped}   Office lock files ignored: {C['lockfiles']}"]
    for cat, n in totals.items():
        lines.append(f"  {cat:15s} files with hits: {n}")
    others = "  ".join(f"{k.split('-')[1]}: {dfars_files[k]}" for k in sorted(dfars_files) if k != "252.204-7019")
    lines.append(f"  252.204-7019 found in {dfars_files['252.204-7019']} files -> {yn(dfars_files['252.204-7019'])}"
                 + (f"   (other clauses  {others})" if others else ""))
    lines.append(f"  export markings found in {totals['EXPORT_CONTROL']} files -> {yn(totals['EXPORT_CONTROL'])}"
                 f"   (strong ITAR/SP-EXPT/USML/22 CFR/Dist Stmt/NOFORN: {C['export_strong']}; "
                 f"weak-only EXPT/ECCN/EAR99/'export control': {C['export_weak']})")
    lines.append("  CUI//SP- categories seen: " + (", ".join(f"{k} {v}" for k, v in sp_files.most_common()) or "none"))
    unsup = ", ".join(f"{k} {v}" for k, v in unsupported.most_common(8))
    lines.append(f"  not read: unsupported types {C['unsupported']}" + (f" ({unsup})" if unsup else "") +
                 f", PDFs without text layer {C['pdf_no_text']}, PDFs no pypdf {C['SKIPPED_PDF']}, "
                 f"PDF encrypted/AES {C['SKIPPED_PDF_ENCRYPTED'] + C['SKIPPED_PDF_AES']}, over size cap {C['SKIPPED_SIZE']}, "
                 f"cloud placeholders {C['placeholders']}, password-protected Office {C['SKIPPED_ENCRYPTED_OFFICE']}, "
                 f"legacy Office no hits {C['legacy_nohit']} (with hits {C['legacy_hits']}), empty {C['SKIPPED_EMPTY']}, "
                 f"errors {C['errors']}, folders not listable {C['dir_errors']}, junctions {C['junctions']}, "
                 f"excluded folders {C['excluded_dirs']}, partial reads {C['partial']}")
    lines.append("  top folders by files with hits (relative to root):" if folder_hits else "  top folders by files with hits: none")
    for k, v in folder_hits.most_common(10):
        lines.append(f"    {v:4d}  {k}")
    print("\n".join(lines))
    if partial:
        sys.exit(130)
    if scanned == 0:
        print("WARNING: no readable files were found under that folder.", file=sys.stderr)
        sys.exit(3)

if __name__ == "__main__":
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    main()
