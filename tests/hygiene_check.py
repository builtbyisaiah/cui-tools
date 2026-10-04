#!/usr/bin/env python3
"""Hygiene gate. Standard library only.

Fails if any file that would ship contains a forbidden string (the client's identifiers, the
build machine's username, a local profile or temp path) or, inside a git repository, if the
working tree is dirty after the tests ran.

Nothing is exempt, this file included, so the patterns below are assembled from fragments and
the list is never spelled out in one piece anywhere in the repository.
"""
import os
import re
import subprocess
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

FORBIDDEN = [
    ("client name", "rock" + "et" + r"\s*" + "ci" + "ty"),
    ("client abbreviation", r"\b" + "r" + "c" + "s" + r"\b"),
    ("machine username", "ka" + "yn"),
    ("profile path prefix", "us" + "ers" + r"\\"),
    ("temp folder name", "scr" + "atch" + "pad"),
    ("temp folder name", "app" + "data"),
    ("author surname", "van" + r"\s*" + "cle" + "ave"),
]
PATTERNS = [(label, re.compile(rx, re.I)) for label, rx in FORBIDDEN]


def git(*args):
    try:
        r = subprocess.run(["git", "-C", REPO] + list(args), capture_output=True,
                           encoding="utf-8", errors="replace")
    except OSError:
        return 1, ""
    return r.returncode, r.stdout


def in_this_repo():
    rc, top = git("rev-parse", "--show-toplevel")
    if rc != 0 or not top.strip():
        return False
    return os.path.normcase(os.path.abspath(top.strip())) == os.path.normcase(REPO)


def files_to_check(in_repo):
    if in_repo:
        rc, out = git("ls-files", "-z", "--cached", "--others", "--exclude-standard")
        if rc == 0:
            return [os.path.join(REPO, n) for n in out.split("\0") if n]
    found = []
    for dp, dns, fns in os.walk(REPO):
        dns[:] = sorted(d for d in dns if d not in (".git", "__pycache__"))
        found.extend(os.path.join(dp, f) for f in sorted(fns))
    return found


def main():
    in_repo = in_this_repo()
    files = files_to_check(in_repo)
    hits = 0
    for path in files:
        if not os.path.isfile(path):
            continue
        with open(path, "rb") as fh:
            text = fh.read().decode("utf-8", "replace")
        for lineno, line in enumerate(text.splitlines(), 1):
            for label, rx in PATTERNS:
                if rx.search(line):
                    hits += 1
                    print(f"FAIL  {os.path.relpath(path, REPO)}:{lineno}: {label}")
    mode = "git-tracked and untracked files" if in_repo else "all files (not a git repository)"
    print(("PASS  " if hits == 0 else "FAIL  ") + f"forbidden strings: {hits} hit(s) in {len(files)} {mode}")

    dirty = 0
    if in_repo:
        rc, status = git("status", "--porcelain")
        if rc != 0:
            dirty = 1
            print("FAIL  git status could not run")
        elif status.strip():
            dirty = 1
            print("FAIL  working tree is dirty after the test run:")
            for line in status.splitlines():
                print("      " + line)
        else:
            print("PASS  working tree clean (git status --porcelain is empty)")
    else:
        print("SKIP  git status check (not a git repository)")

    return 0 if hits == 0 and dirty == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
