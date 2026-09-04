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

OWNER = "nvoigtla"
# GitHub repository names cannot contain spaces, so "MGMT 405 EMBA" becomes
# this; the published address is https://<owner>.github.io/<REPO>/
REPO = "MGMT-405-EMBA"
DESC = "Course website for MGMT 405 Managerial Economics (EMBA Hybrid)"

# Everything the site needs. Nothing else is copied -- the build script, the
# session notes and the calendar stay private.
#
# The page list is DISCOVERED, not hardcoded: a hardcoded list silently
# dropped all-videos.html and all-podcasts.html when they were added
# (2026-09-03), so both 404'd on the live site while the local build was
# fine. Every .html in this folder is a built page.
def page_files():
    return sorted(f for f in os.listdir(HERE)
                  if f.endswith(".html") and not f.startswith("_"))


ASSETS = ["site.css", "site.js", "search-index.js", "panopto-login.png"]

PUBLIC_README = """# MGMT 405 – Managerial Economics (EMBA Hybrid)

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


def missing():
    out = []
    for f in page_files():
        if not os.path.exists(os.path.join(HERE, f)):
            out.append(f)
    for a in ASSETS:
        if not os.path.exists(os.path.join(HERE, "assets", a)):
            out.append("assets/" + a)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    gone = missing()
    if gone:
        sys.exit("Run `python _build_site.py` first -- missing:\n  " +
                 "\n  ".join(gone))
    files = page_files()
    if not files:
        sys.exit("no built pages found -- run `python _build_site.py` first")
    print("all %d pages and %d assets present" % (len(files), len(ASSETS)))

    if args.dry_run:
        print("\nwould publish to https://github.com/%s/%s" % (OWNER, REPO))
        print("would serve at  https://%s.github.io/%s/" % (OWNER, REPO))
        for f in files:
            print("   " + f)
        for a in ASSETS:
            print("   assets/" + a)
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
            shutil.copy2(os.path.join(HERE, f), os.path.join(work, f))
        for a in ASSETS:
            shutil.copy2(os.path.join(HERE, "assets", a),
                         os.path.join(work, "assets", a))
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
