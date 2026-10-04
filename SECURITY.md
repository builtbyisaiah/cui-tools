# Security policy

## What to report

- A marking form the scanner misses that the README does not already list as a known limit
  (a false negative).
- Any way document content reaches stdout, stderr, or the CSV. The output contract is
  verdicts, counts, fixed vocabulary terms, and paths only.
- Any write inside the scanned tree, or any network activity.

Not in scope: unmarked CUI that carries no textual marking, OCR of scanned images, and
machine-learning classification. The tools do not claim any of those.

## How to report

Use GitHub's private vulnerability reporting on this repository (Security tab, "Report a
vulnerability"). Do not open a public issue for a false negative until it is fixed or listed.

Never attach a real document, a real file listing, or a real report block. Describe the
marking form in words (for example "CUI banner inside a Word text box, all-caps formatting")
or build a synthetic file that reproduces it and attach that. Everything in this repository's
test corpus is synthetic and yours should be too.

## Supported versions

Only the latest tagged release receives fixes.
