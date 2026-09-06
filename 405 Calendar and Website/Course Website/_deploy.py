# -*- coding: utf-8 -*-
"""
Publish the built website to its own PUBLIC GitHub repository and serve it
with GitHub Pages.

The source of truth stays in this (private) Teaching repository; only the
built HTML is copied out. Run after `python _build_site.py`:

    python _deploy.py              # build check, then commit + push + Pages
    python _deploy.py --dry-run    # show what would be pushed, change nothing

Requires the `gh` CLI, already authenticated (`gh auth status`).

The deploy target is cloned into a temporary folder each time, so no nested
git repository is ever created inside Teaching.
"""

import argparse
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
CAL = os.path.abspath(os.path.join(HERE, os.pardir, "Course Calendar"))
if CAL not in sys.path:
    sys.path.insert(0, CAL)

# --section has to land in the environment BEFORE _calendar_content is
# imported: that module reads MGMT405_SECTION at import time (2026-09-05).
for _i, _a in enumerate(sys.argv):
    if _a == "--section" and _i + 1 < len(sys.argv):
        os.environ["MGMT405_SECTION"] = sys.argv[_i + 1].lower()
    elif _a.startswith("--section="):
        os.environ["MGMT405_SECTION"] = _a.split("=", 1)[1].lower()

import _calendar_content as C  # noqa: E402

# EMBA publishes this folder; a second section publishes its subfolder --
# the same split _build_site.py writes to.
SITE = HERE if C.SECTION == "emba" else os.path.join(HERE, C.SECTION)

OWNER = "nvoigtla"
# GitHub repository names cannot contain spaces, so "MGMT 405 EMBA" becomes
# this; the published address is https://<owner>.github.io/<REPO>/
REPO = C.REPO
DESC = ("Course website for MGMT 405 Managerial Economics (%s Hybrid)"
        % C.SECTION_LABEL)

# Everything the site needs. Nothing else is copied -- the build script, the
# session notes and the calendar stay private.
#
# The page list is DISCOVERED, not hardcoded: a hardcoded list silently
# dropped all-videos.html and all-podcasts.html when they were added
# (2026-09-03), so both 404'd on the live site while the local build was
# fine. Every .html in this folder is a built page.
def page_files():
    return sorted(f for f in os.listdir(SITE)
                  if f.endswith(".html") and not f.startswith("_"))


ASSETS = ["site.css", "site.js", "search-index.js", "panopto-login.png"]

# Documents published NEXT TO the site, so the General Logistics page can
# link them (2026-09-04, Nico): (source path, published file name). The
# published names carry no spaces -- they are served straight off GitHub
# Pages -- and they have to match LINKS["syllabus_pdf"] / ["calendar_pdf"]
# in ../Course Calendar/_calendar_content.py, which is what the site, the
# syllabus and the calendar all link.
DOCS = [
    (os.path.join(HERE, os.pardir, "Syllabus", C.SYLLABUS_DOCX + ".pdf"),
     "MGMT-405-Syllabus-Fall-2026.pdf"),
    (os.path.join(HERE, os.pardir, "Course Calendar", C.CALENDAR_DOCX + ".pdf"),
     "MGMT-405-Calendar-Fall-2026.pdf"),
]

PUBLIC_README = """# MGMT 405 – Managerial Economics (%s Hybrid)""" % C.SECTION_LABEL + """

Course website: **https://%s.github.io/%s/**

These pages are generated from the course calendar and are published for the
students of MGMT 405 at UCLA Anderson. Do not edit them here -- they are
overwritten by the build script on every deploy.

Prof. Nico Voigtländer · UCLA Anderson
""" % (OWNER, REPO)


def run(cmd, cwd=None, check=True, quiet=False):
    if not quiet:
        print("   $ " + " ".join(cmd))
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if check and p.returncode != 0:
        sys.exit("FAILED (%d): %s\n%s%s"
                 % (p.returncode, " ".join(cmd), p.stdout, p.stderr))
    return p


def missing(skip_docs=False):
    out = []
    for f in page_files():
        if not os.path.exists(os.path.join(SITE, f)):
            out.append(f)
    for a in ASSETS:
        if not os.path.exists(os.path.join(SITE, "assets", a)):
            out.append("assets/" + a)
    if not skip_docs:
        for src, name in DOCS:
            if not os.path.exists(src):
                out.append(name + "  (export it from its .docx first: %s)" % src)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--section", default=os.environ.get("MGMT405_SECTION"),
                    help="emba (default) or femba; already consumed above")
    ap.add_argument("--dry-run", action="store_true")
    # Word's ExportAsFixedFormat has been hanging on this machine, so a
    # section can be published before its PDFs exist. The two links to them
    # 404 until the files are added, which is why this is never the default
    # (2026-09-05).
    ap.add_argument("--skip-docs", action="store_true",
                    help="publish without the syllabus / calendar PDFs")
    args = ap.parse_args()

    gone = missing(skip_docs=args.skip_docs)
    if gone:
        sys.exit("Run `python _build_site.py` first -- missing:\n  " +
                 "\n  ".join(gone))
    files = page_files()
    if not files:
        sys.exit("no built pages found -- run `python _build_site.py` first")
    have_docs = sum(1 for src, _ in DOCS if os.path.exists(src))
    print("all %d pages, %d assets and %d of %d PDFs present"
          % (len(files), len(ASSETS), have_docs, len(DOCS)))

    if args.dry_run:
        print("\nwould publish to https://github.com/%s/%s" % (OWNER, REPO))
        print("would serve at  https://%s.github.io/%s/" % (OWNER, REPO))
        for f in files:
            print("   " + f)
        for a in ASSETS:
            print("   assets/" + a)
        for _src, name in DOCS:
            print("   " + name)
        print("   README.md, .nojekyll")
        return

    slug = "%s/%s" % (OWNER, REPO)

    # ---- create the repository if it is not there yet ----
    if run(["gh", "repo", "view", slug], check=False, quiet=True).returncode:
        print("creating public repository %s" % slug)
        run(["gh", "repo", "create", slug, "--public", "-d", DESC])
    else:
        print("repository %s already exists" % slug)

    tmp = tempfile.mkdtemp(prefix="m405-deploy-")
    work = os.path.join(tmp, REPO)
    try:
        run(["git", "clone", "https://github.com/%s.git" % slug, work])

        # ---- sync the built site into the clone ----
        for name in os.listdir(work):
            if name == ".git":
                continue
            path = os.path.join(work, name)
            shutil.rmtree(path) if os.path.isdir(path) else os.remove(path)

        os.makedirs(os.path.join(work, "assets"), exist_ok=True)
        for f in files:
            shutil.copy2(os.path.join(SITE, f), os.path.join(work, f))
        for a in ASSETS:
            shutil.copy2(os.path.join(SITE, "assets", a),
                         os.path.join(work, "assets", a))
        for src, name in DOCS:
            if os.path.exists(src):
                shutil.copy2(src, os.path.join(work, name))
            else:
                print("   WARNING: %s not published (no %s)"
                      % (name, os.path.basename(src)))
        io.open(os.path.join(work, "README.md"), "w", encoding="utf-8",
                newline="\n").write(PUBLIC_README)
        # keep GitHub Pages from running the site through Jekyll
        io.open(os.path.join(work, ".nojekyll"), "w").write("")

        # a freshly created repository clones empty, with no upstream yet
        branch = (run(["git", "symbolic-ref", "--short", "HEAD"], cwd=work,
                      check=False, quiet=True).stdout.strip() or "main")

        run(["git", "add", "-A"], cwd=work)
        st = run(["git", "status", "--porcelain"], cwd=work, quiet=True)
        if not st.stdout.strip():
            print("\nno changes to publish -- the live site is already current")
        else:
            run(["git", "commit", "-m",
                 "Publish course website (built from the course calendar)"],
                cwd=work)
            run(["git", "push", "-u", "origin", branch], cwd=work)
            print("\npushed %d changed paths" % len(st.stdout.strip().splitlines()))

        # ---- turn Pages on (a no-op once it is already enabled) ----
        pg = run(["gh", "api", "repos/%s/pages" % slug], check=False, quiet=True)
        if pg.returncode:
            print("enabling GitHub Pages")
            body = os.path.join(tmp, "pages.json")
            io.open(body, "w", encoding="utf-8").write(
                json.dumps({"source": {"branch": branch, "path": "/"}}))
            r = run(["gh", "api", "-X", "POST", "repos/%s/pages" % slug,
                     "-H", "Accept: application/vnd.github+json",
                     "--input", body], check=False)
            if r.returncode:
                print("   could not enable Pages from here. Turn it on at")
                print("   https://github.com/%s/settings/pages" % slug)
                print("   (branch %s, folder /root), then re-run." % branch)
                print(r.stdout + r.stderr)
        else:
            print("GitHub Pages already enabled")

        print("\n" + "=" * 62)
        print("  repository:  https://github.com/%s" % slug)
        print("  live site:   https://%s.github.io/%s/" % (OWNER, REPO))
        print("=" * 62)
        print("Pages takes a minute or two to build on the first deploy.")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
