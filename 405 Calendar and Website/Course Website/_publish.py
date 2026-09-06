# -*- coding: utf-8 -*-
"""
"Now update the website" -- the one command that publishes BOTH sections and
refuses to ship anything stale.

    python _publish.py            # check, fix what it can, publish
    python _publish.py --check    # report only; nothing is published
    python _publish.py --section femba

WHY THIS EXISTS (2026-09-06, Nico asked how we avoid missing an updated
document). Committing to git does NOT update the website: git holds the
private source, `_deploy.py` publishes to the public repos that GitHub Pages
serves. And the gap sits upstream of both -- `_deploy.py` copies whatever
.pdf happens to be on disk, so when a .docx changes and the PDF is not
re-exported, the deploy ships the OLD one without a word. That is exactly
what happened: both sites served syllabus and calendar PDFs saying room A301
for two days while every HTML page correctly said G305.

So this script closes each hole in order:

  1. REBUILD both sites, so the HTML can never lag _calendar_content.py.
  2. RE-EXPORT any PDF whose .docx is newer (or that is missing at all).
  3. VERIFY each PDF opens, has plausible length, and belongs to its section.
  4. CHECK every local link in the built pages resolves to a file.
  5. REPORT what is still a placeholder -- the "(TBD)" handouts and slide
     decks and the "(link to follow)" videos, so nothing waits unnoticed.
  6. PUBLISH both sections.

Step 5 is the one to read after an upload: it is the list of things students
cannot get yet.
"""

import argparse
import glob
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, os.pardir))
PY = sys.executable

SECTIONS = ("emba", "femba")

# Exporting to PDF. ExportAsFixedFormat hung for days on this machine; what
# unblocked it was setting ActivePrinter to a local printer FIRST -- Word
# wants printer metrics to lay the PDF out, and it stalls when the default
# is unavailable.
#
# Do NOT go back to printing the document to the "Microsoft Print to PDF"
# driver as a workaround. That route returns quickly but FLATTENS the file:
# the calendar came out with 0 hyperlinks instead of 90, losing every video,
# podcast and slide-deck link, because they sit inside drawn text boxes.
# verify_pdfs() checks the link count for exactly that reason (2026-09-06).
EXPORT_HEAD = r"""
$ErrorActionPreference = 'Continue'
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$w.ActivePrinter = "Microsoft Print to PDF"
"""

EXPORT_JOB = r"""
try {
  if (Test-Path "%(out)s") { Remove-Item "%(out)s" -Force }
  $d = $w.Documents.Open("%(src)s", $false, $true)
  $d.ExportAsFixedFormat("%(out)s", 17)
  $d.Close(0)
} catch { Write-Output ("FAIL %(out)s : " + $_.Exception.Message) }
"""

EXPORT_TAIL = r"""
$w.Quit()
Start-Sleep -Seconds 2
"""



def sh(cmd, env=None, quiet=False):
    p = subprocess.run(cmd, cwd=HERE, capture_output=True, text=True,
                       env=env or os.environ.copy())
    if p.returncode and not quiet:
        print(p.stdout)
        print(p.stderr)
    return p


def section_env(section):
    e = os.environ.copy()
    e["MGMT405_SECTION"] = section
    return e


def section_facts(section):
    """(calendar basename, syllabus basename, repo, out dir) for a section --
    read from the content module rather than repeated here."""
    # the content module lives with the calendar, not here
    code = ("import sys; sys.path.insert(0, r'%s');"
            "import _calendar_content as C;"
            "print(C.CALENDAR_DOCX);print(C.SYLLABUS_DOCX);print(C.REPO)"
            % os.path.join(ROOT, "Course Calendar"))
    p = sh([PY, "-c", code], env=section_env(section))
    cal, syl, repo = p.stdout.strip().splitlines()[:3]
    out = HERE if section == "emba" else os.path.join(HERE, section)
    return cal, syl, repo, out


def doc_pairs(section):
    cal, syl, _repo, _out = section_facts(section)
    return [
        (os.path.join(ROOT, "Syllabus", syl + ".docx"),
         os.path.join(ROOT, "Syllabus", syl + ".pdf"), "syllabus"),
        (os.path.join(ROOT, "Course Calendar", cal + ".docx"),
         os.path.join(ROOT, "Course Calendar", cal + ".pdf"), "calendar"),
    ]


# ------------------------------ the steps ------------------------------

def rebuild(section):
    p = sh([PY, "_build_site.py", "--section", section])
    if p.returncode:
        sys.exit("  build FAILED for %s" % section)
    n = re.search(r"asset version (\S+)", p.stdout)
    print("    rebuilt (%s)" % (n.group(1) if n else "ok"))


def stale_pdfs(section):
    """(docx, pdf, why) for every PDF that is missing or older than its
    source. Mtime is the right test: both files are generated, and the PDF is
    only ever produced FROM the .docx."""
    out = []
    for docx, pdf, what in doc_pairs(section):
        if not os.path.exists(docx):
            print("    %-9s SOURCE MISSING: %s" % (what, docx))
            continue
        if not os.path.exists(pdf):
            out.append((docx, pdf, "%s PDF missing" % what))
        elif os.path.getmtime(pdf) < os.path.getmtime(docx):
            out.append((docx, pdf, "%s PDF older than its .docx" % what))
    return out


def export(jobs):
    """Export each (docx, pdf) in one Word session, then CONFIRM each
    target is now newer than its source. The confirmation matters: a silent
    no-op here would leave a stale PDF and report success."""
    script = EXPORT_HEAD + "".join(
        EXPORT_JOB % {"src": docx, "out": pdf} for docx, pdf, _w in jobs
    ) + EXPORT_TAIL
    p = subprocess.run(["powershell", "-NoProfile", "-Command", script],
                       capture_output=True, text=True)
    if p.stdout.strip():
        for line in p.stdout.strip().splitlines():
            print("    " + line.strip())
    ok = True
    for docx, pdf, _why in jobs:
        if not os.path.exists(pdf):
            print("    EXPORT FAILED -- %s was not written"
                  % os.path.basename(pdf))
            ok = False
        elif os.path.getmtime(pdf) < os.path.getmtime(docx):
            print("    EXPORT FAILED -- %s is still older than its .docx"
                  % os.path.basename(pdf))
            ok = False
        else:
            print("    wrote %s (%.0f KB)"
                  % (os.path.basename(pdf), os.path.getsize(pdf) / 1024.0))
    return ok


def verify_pdfs(section):
    """Each PDF must open, carry pages, and name its own section. Subset
    fonts make a raw byte scan useless here -- read the text with PyMuPDF."""
    try:
        import fitz
    except ImportError:
        print("    (PyMuPDF not installed -- content of the PDFs unchecked)")
        return True
    ok = True
    label = "FEMBA" if section == "femba" else "EMBA"
    for _docx, pdf, what in doc_pairs(section):
        if not os.path.exists(pdf):
            print("    %-9s PDF STILL MISSING" % what)
            ok = False
            continue
        try:
            d = fitz.open(pdf)
            txt = "".join(d[i].get_text() for i in range(d.page_count))
        except Exception as e:
            print("    %-9s PDF UNREADABLE: %s" % (what, e))
            ok = False
            continue
        wrong = [r for r in ("A301", "G-402") if r in txt]
        mine = ("%s Section" % label) in txt
        # A PDF with no hyperlinks means it was FLATTENED -- printing to a
        # PDF driver does that, and it silently strips every video, podcast
        # and slide-deck link out of the calendar (2026-09-06).
        links = sum(1 for i in range(d.page_count)
                    for l in d[i].get_links() if l.get("uri"))
        dead = links == 0
        good = mine and not wrong and not dead
        print("    %-9s %2d pages, %2d links, %s section%s%s"
              % (what, d.page_count, links, "correct" if mine else "WRONG",
                 (", stale room %s" % ", ".join(wrong)) if wrong else "",
                 "" if good else "   <-- CHECK"
                 + (" (no hyperlinks -- flattened?)" if dead else "")))
        ok = ok and good
    return ok


def check_links(out_dir):
    """Every local href/src in the built pages must resolve. This is what
    will catch a teaching note or a slide deck that is linked but never
    copied into the site folder."""
    bad = set()
    for f in glob.glob(os.path.join(out_dir, "*.html")):
        s = open(f, encoding="utf-8").read()
        for m in re.finditer(r'(?:href|src)="([^"]+)"', s):
            u = m.group(1)
            if u.startswith(("http", "#", "mailto", "data:")):
                continue
            target = os.path.join(out_dir, u.split("?")[0].split("#")[0])
            if not os.path.exists(target):
                bad.add((os.path.basename(f), u))
    for f, u in sorted(bad):
        print("    BROKEN LINK  %s -> %s" % (f, u))
    return not bad


def placeholders(out_dir):
    """What students still cannot get: "(TBD)" handouts and slide decks,
    "(link to follow)" videos and podcasts. Grouped by page."""
    rows = {}
    for f in sorted(glob.glob(os.path.join(out_dir, "week-*.html"))):
        s = open(f, encoding="utf-8").read()
        found = re.findall(r'class="txt">([^<]{0,70}?)\s*<span class="tba">'
                           r'\(([^)]*)\)', s)
        if found:
            rows[os.path.basename(f)] = found
    total = sum(len(v) for v in rows.values())
    print("    %d item(s) still to upload, on %d page(s)" % (total, len(rows)))
    for page, items in rows.items():
        kinds = {}
        for text, kind in items:
            kinds.setdefault(kind, []).append(text)
        bits = "; ".join("%d %s" % (len(v), k) for k, v in kinds.items())
        print("      %-14s %s" % (page, bits))
    return total


def publish(section, repo):
    p = sh([PY, "_deploy.py", "--section", section])
    for line in p.stdout.splitlines():
        if any(k in line for k in ("pushed", "no changes", "WARNING",
                                   "present", "live site")):
            print("    " + line.strip())
    if p.returncode:
        sys.exit("  deploy FAILED for %s" % section)


# ------------------------------ main ------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true",
                    help="report only; publish nothing")
    ap.add_argument("--section", choices=SECTIONS + ("both",), default="both")
    args = ap.parse_args()
    todo = SECTIONS if args.section == "both" else (args.section,)

    problems = []
    for section in todo:
        _cal, _syl, repo, out_dir = section_facts(section)
        print("\n=== %s  (%s) ===" % (section.upper(), repo))

        print("  1. rebuild")
        rebuild(section)

        print("  2. documents")
        stale = stale_pdfs(section)
        if not stale:
            print("    both PDFs are newer than their .docx")
        elif args.check:
            for _d, _p, why in stale:
                print("    STALE: %s" % why)
            problems.append("%s: %d stale PDF(s)" % (section, len(stale)))
        else:
            for _d, _p, why in stale:
                print("    re-exporting -- %s" % why)
            if not export(stale):
                problems.append("%s: PDF export failed" % section)

        print("  3. verify PDFs")
        if not verify_pdfs(section):
            problems.append("%s: PDF verification failed" % section)

        print("  4. links")
        if check_links(out_dir):
            print("    every local link resolves")
        else:
            problems.append("%s: broken local link(s)" % section)

        print("  5. still to upload")
        placeholders(out_dir)

        print("  6. publish")
        if args.check:
            print("    (--check: nothing published)")
        elif problems and any(p.startswith(section) for p in problems):
            print("    SKIPPED -- fix the problems above first")
        else:
            publish(section, repo)

    print("\n" + "=" * 60)
    if problems:
        print("PROBLEMS:")
        for p in problems:
            print("  - " + p)
        sys.exit(1)
    print("both sections are up to date" if not args.check
          else "check complete -- nothing was published")


if __name__ == "__main__":
    main()
