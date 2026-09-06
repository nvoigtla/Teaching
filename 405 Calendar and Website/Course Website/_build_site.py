# -*- coding: utf-8 -*-
"""
Build the MGMT 405 course website: one page per week, one per module, plus a
General Logistics page.

The course calendar's `_calendar_content.py` is the SINGLE SOURCE OF TRUTH for
dates, topics, videos, podcasts, readings, quizzes and exam windows. This
script imports it and writes static HTML -- no content is typed twice.

    python _build_site.py

Output (all overwritten on every run):
    index.html                 General Logistics / Before You Start
    week-01.html .. week-12.html
    module-1.html .. module-8.html
    assets/search-index.js     the search box's index

`assets/site.css` and `assets/site.js` are hand-authored and are NOT touched.
"""

import base64
import glob
import hashlib
import io
import json
import os
import re
import shutil
import sys

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

# EMBA keeps this folder, so its published site is untouched; a second
# section gets a subfolder of its own. SRC is where the hand-authored
# stylesheet and script live -- always this folder, for both sections.
SRC = HERE
OUT = HERE if C.SECTION == "emba" else os.path.join(HERE, C.SECTION)

SITE_NAME = "Managerial Economics Fall 2026 – %s" % C.SECTION_LABEL

# What the browser tab shows, on every page (2026-09-03, Nico).
TAB_TITLE = "Managerial Econ 405"

# Help and Questions carries two marks doing two jobs (2026-09-04, Nico):
# a question mark for "where do I get help", an envelope for "by email".
# The top-bar button takes the question mark; the box header takes both; the
# bullets take the envelope alone. All three are plain text characters, so
# they pick up the pale gold that the emoji glyphs cannot.
HELP_Q = "?"
HELP_MAIL = "&#9993;"
HELP_HD_GLYPH = HELP_Q + "&nbsp;" + HELP_MAIL

# The syllabus and the calendar are published as PDFs alongside the site
# (2026-09-04, Nico) -- _deploy.py copies both out of their folders. The
# addresses live in the calendar's LINKS registry, so the calendar, the
# syllabus and this page all read one string.
SYLLABUS_URL = C.LINKS["syllabus_pdf"]
CALENDAR_PDF_URL = C.LINKS["calendar_pdf"]

# The Panopto sign-in screenshot lives with the calendar; the build copies
# it into assets/ so the site can link it (2026-09-03, Nico).
PANOPTO_SHOT_SRC = os.path.join(CAL, "Images", "Panopto-Login-Picture.png")
PANOPTO_SHOT = "assets/panopto-login.png"

# The course's BruinLearn site. It moved INTO the calendar's LINKS registry
# on 2026-09-04, when the calendar's problem-set cards started linking it
# too, so the address is no longer typed twice.
BRUINLEARN_COURSE = C.LINKS["bruinlearn_course"]

# The instructor's faculty page, linked from the header (2026-09-03, Nico).
NICO_URL = "https://www.anderson.ucla.edu/faculty_pages/nico.v/"

# Prof. Nico's address (2026-09-04). It is never written into the HTML --
# see mail_link() below.
NICO_EMAIL = "nico.v@ucla.edu"

# General Logistics link that survives the "land on the current week"
# redirect below, so the page is always reachable from the sidebar.
GL_HREF = "index.html?stay=1"

# ===================== category cards (calendar format) =====================
# Title, header glyph and body tint per category, matching _build_calendar.py:
# videos yellow, podcasts light gray, reading and practice white. The header
# glyph for Videos is a FILM REEL (2026-09-03, Nico); the individual video
# bullets keep the play triangle.
CAT_ORDER = [
    ("video",    "Videos",                                  "\U0001F3AC"),
    ("podcast",  "Podcasts",                                "\U0001F3A7"),
    ("read",     "Suggested Reading",                       "\U0001F4D6"),
    ("practice", "Suggested Additional Practice Exercises", "✎"),
]
CAT_TITLE = {k: t for k, t, _ in CAT_ORDER}
CAT_GLYPH = {k: g for k, _, g in CAT_ORDER}

# Per-item bullet glyphs.
ITEM_GLYPH = {"video": "▶", "podcast": "♪", "topic": "◆"}
CAT_ITEM_GLYPH = {"read": "▤", "practice": "✎", "podcast": "♪"}

# ============================== modules ==============================
MODULES = [
    (1, "Basic Concepts and Economic Principles", "Basic Concepts", None),
    (2, "Demand Analysis", "Demand Analysis", None),
    (3, "Production & Costs", "Production & Costs", None),
    (4, "Competitive Markets, Interventions & Externalities",
        "Competitive Markets",
        ["Part I: Competitive Markets and Market Interventions",
         "Part II: Market Distortions / Externalities"]),
    (5, "Monopoly and Monopolistic Competition", "Monopoly", None),
    (6, "Complex Pricing and Advanced Pricing Strategies", "Complex Pricing", None),
    (7, "Oligopoly and Game Theory", "Oligopoly & Game Theory",
        ["Part I: Oligopoly with Homogenous Goods",
         "Part II: Oligopoly with Diff. Goods; Game Theory"]),
    (8, "Asymmetric Information; Auctions", "Asymmetric Information", None),
]

# Two index pages that sit in both menus, right after General Logistics
# (2026-09-03, Nico): href, current-key, title, sub-line, badge glyph.
# The glyphs are the ones the week headers use for those categories -- the
# clapperboard and the headphones (2026-09-03, Nico).
EXTRA_PAGES = [
    ("all-videos.html", "all-videos", "All Videos",
     "Every video, by module", "\U0001F3AC"),
    ("all-podcasts.html", "all-podcasts", "All Podcasts",
     "Intro and wrap-up, by module", "\U0001F3A7"),
]

# The TA's exercise site -- an outgoing link, not a page of ours, so it sits
# at the very bottom of the menu and opens in a new tab (2026-09-03, Nico).
PRACTICE_INDEX = "https://rafaelrubiao.github.io/mgmt405-practice/index.html"
EXTRA_LINKS = [
    (PRACTICE_INDEX, "all-practice", "All Practice Exercises",
     "Online exercises with solutions", "✎"),
]

_MOD_RE = re.compile(r"Module\s+(\d)")

# Groups whose module tags cannot be read off their own text. KEYED BY A
# DISTINCTIVE SNIPPET of the group's label -- or of its first item, when it
# has no label -- NOT by position. It used to be keyed by the group's index
# within the week, and deleting the "Recap of Module X" groups silently
# re-pointed every override that followed one (2026-09-05). check_overrides()
# fails the build if a key stops matching exactly one group.
GROUP_MODULE_OVERRIDES = {
    (1, "prep", "Optional Podcasts"): [],        # general-interest econ podcasts
    # "In preparation for class: Ch. 2.5" names no module; tagged with the two
    # modules that week 1's on-campus class covers, per that week's topics.
    (1, "prep", "In preparation for class:"): [1, 2],
    (2, "prep", "[Relevant textbook reading was already covered"): [2],
    (2, "prep", "Advanced reading (optional):"): [2],   # Ch. 5 (elasticity)
    (2, "prep", "Teaching notes (optional"): [2],       # MR, elasticity, regressions
    # The teaching-note groups name no module in their own text, so each
    # week's group needs its tag here (2026-09-06). The snippet matches the
    # group's LABEL -- group_text() returns the label whenever there is one,
    # and all three groups share it; the week number is what tells them apart.
    (3, "prep", "Teaching notes (optional"): [3],       # hiring, bang-for-buck
    (5, "prep", "Teaching notes (optional"): [5],       # MR = MC
    (3, "prep", "Advanced reading (optional):"): [3],   # Ch. 6.6 - 6.7
    (5, "prep", "Assigned articles for discussion"): [],
    (6, "prep", "Midterm Prep: TA Review Sessions"): [],
    (9, "prep", "Optional Podcasts"): [],        # general-interest econ podcast
    (9, "prep", "Assigned articles for discussion"): [],
    (11, "prep", "Exam prep time:"): [],
}

# Hooks for hiding a calendar line on the WEBSITE only, while it still
# prints in the calendar. Both are empty: everything Nico wanted gone was
# deleted from _calendar_content.py outright (2026-09-03). Add a distinctive
# substring of the line here if a website-only omission is ever needed.
DROP_TEXTBOOK_NOTES = []
DROP_MATH_ITEMS = []


def mods_in(text):
    out = []
    for m in _MOD_RE.finditer(text or ""):
        n = int(m.group(1))
        if 1 <= n <= 8 and n not in out:
            out.append(n)
    return out


# ============================== helpers ==============================

def esc(s):
    return (str("" if s is None else s).replace("&", "&amp;")
            .replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))


# Title Case for headers, the same rule the slide decks use: capitalize every
# significant word, leave articles / conjunctions / short prepositions lower
# unless they open or close the title, capitalize both halves of a hyphenated
# compound -- and NEVER lower-case a letter that is already capital, so
# acronyms survive (2026-09-03, Nico).
_SMALL = {"a", "an", "the", "and", "but", "or", "nor", "of", "in", "at", "to",
          "for", "with", "on", "by", "from", "as", "per", "vs"}


def title_case(s):
    words = str(s).split(" ")
    last = len(words) - 1

    def fix(word, is_first, is_last):
        parts = word.split("-")
        out = []
        for j, p in enumerate(parts):
            if not p:
                out.append(p)
                continue
            bare = re.sub(r"[^A-Za-z]", "", p).lower()
            keep_low = (bare in _SMALL) and not is_first and not is_last and j == 0
            out.append(p if keep_low else p[0].upper() + p[1:])
        return "-".join(out)

    return " ".join(fix(w, i == 0, i == last) for i, w in enumerate(words))


def link_for(key):
    return C.LINKS.get(key) if key else None


def week_of(n):
    for w in C.WEEKS:
        if w["num"] == n:
            return w
    return None


def week_span(w):
    if w.get("span_override"):
        (w0, d0), (w1, d1) = w["span_override"]
        return C.dt(w0, d0), C.dt(w1, d1)
    return C.dt(w["num"], "Mon"), C.dt(w["num"], "Sun")


def group_text(g):
    """The group's label, or its first item's text when it has no label --
    what GROUP_MODULE_OVERRIDES matches its snippets against."""
    if g.get("label"):
        return g["label"]
    for it in g["items"]:
        txt = it[1] if it[0] in ("t", "b", "note") else (
            it[2] if it[0] in ("v", "l", "p") else "")
        if txt:
            return txt
    return ""


def group_modules(g, wnum, where, idx):
    hay = group_text(g)
    for (w, wh, snippet), mods in GROUP_MODULE_OVERRIDES.items():
        if w == wnum and wh == where and snippet in hay:
            return list(mods)
    mods = mods_in(g.get("label"))
    if mods:
        return mods
    seen = []
    for it in g["items"]:
        txt = it[1] if it[0] in ("t", "b", "note") else (
            it[2] if it[0] in ("v", "l", "p") else "")
        for n in mods_in(txt):
            if n not in seen:
                seen.append(n)
    return seen


def week_modules(w):
    """Every module a week touches, from its topics and all its groups."""
    mods = []
    for t in w.get("topics") or []:
        for n in mods_in(t):
            if n not in mods:
                mods.append(n)
    for where, groups in (("prep", w.get("prep_groups") or []),
                          ("weekend", (w.get("weekend") or {}).get("groups") or [])):
        for i, g in enumerate(groups):
            for n in group_modules(g, w["num"], where, i):
                if n not in mods:
                    mods.append(n)
    return sorted(mods)


def item_modules(it, group_mods):
    txt = it[1] if it[0] in ("t", "b", "note") else (
        it[2] if it[0] in ("v", "l", "p") else "")
    own = mods_in(txt)
    return own or list(group_mods)


def has_video_groups(w):
    return any(g.get("cat") == "video" for g in (w.get("prep_groups") or []))


# ============================== item rendering ==============================

def render_segments(segs):
    out = []
    for s in segs:
        if s[0] == "t":
            out.append(esc(s[1]))
        elif s[0] == "l":
            url = link_for(s[1])
            out.append('<a href="%s" target="_blank" rel="noopener">%s</a>'
                       % (esc(url), esc(s[2])) if url else esc(s[2]))
    return "".join(out)


def podcast_label(text):
    """"Podcast: Intro to Module 1" -> "Podcast (<u>before</u> class): Intro
    to Module 1". The underline marks WHEN to listen (2026-09-03, Nico).

    The phrase after the underlined word comes from podcast_when() in the
    content module, which the calendar reads too -- a wrap-up says "after
    watching the Module 3 videos" where the on-campus session only did that
    module's applications (2026-09-05). Anything that is not one of the two
    module episodes is left alone."""
    w = C.podcast_when(text)
    if w is None:
        return esc(text)
    when, tail = w
    rest = text[len("Podcast: "):] if text.startswith("Podcast: ") else text
    return "Podcast (<u>%s</u> %s): %s" % (when, esc(tail), esc(rest))


def render_item(it, cat, gmods=None):
    """One calendar item tuple -> one <li>."""
    kind = it[0]

    if kind == "note":
        return ('<li class="note"><span class="g" aria-hidden="true"></span>'
                '<span class="txt">%s</span></li>' % esc(it[1]))

    glyph = ITEM_GLYPH.get(
        {"v": "video", "p": "podcast", "b": "topic"}.get(kind, ""),
        CAT_ITEM_GLYPH.get(cat, "•"))

    if kind in ("t", "b"):
        inner = esc(it[1])
    elif kind == "mix":
        inner = render_segments(it[1])
    elif kind in ("v", "p"):
        url = link_for(it[1]) if kind == "v" else it[1]
        text = podcast_label(it[2]) if kind == "p" else esc(it[2])
        inner = ('<a href="%s" target="_blank" rel="noopener">%s</a>'
                 % (esc(url), text)) if url else text
        # running time sits RIGHT AFTER the link, as in the calendar
        if not url:
            inner += ' <span class="tba">(link to follow)</span>'
        elif it[3]:
            inner += ' <span class="mins">(%d min)</span>' % it[3]
        # a video whose link works but whose running time is unmeasured
        # gets no marker at all: "(++)" was cryptic, and "(link to follow)"
        # would be untrue (2026-09-06, Nico)
        if kind == "v":
            # the slide deck behind this video, right after the running
            # time and separated by a dot -- the calendar's treatment,
            # which Nico prefers (2026-09-06)
            mods = item_modules(it, gmods or [])
            deck = C.slides_for(mods[0] if mods else None, it[2])
            if deck:
                inner += (' <span class="deck">\u00b7 '
                          '<a href="slides/%s">slides</a></span>'
                          % esc(C.slides_pub_name(deck)))
    elif kind == "l":
        url = link_for(it[1])
        inner = ('<a href="%s" target="_blank" rel="noopener">%s</a>'
                 % (esc(url), esc(it[2]))) if url else esc(it[2])
    else:
        raise ValueError("unknown item kind %r" % (kind,))

    cls = ' class="topic"' if kind == "b" else ""
    return ('<li%s><span class="g" aria-hidden="true">%s</span>'
            '<span class="txt">%s</span></li>' % (cls, glyph, inner))


def render_group(g, cat, items=None, gmods=None):
    """The italic navy lead-in, then the bullets. No per-week video total
    (2026-09-03, Nico)."""
    items = g["items"] if items is None else items
    h = ['<div class="grp">']
    if g.get("label"):
        h.append('<p class="grp-lab">%s</p>' % esc(g["label"]))
    h.append('<ul class="items">')
    h += [render_item(it, cat, gmods) for it in items]
    h.append("</ul></div>")
    return "".join(h)


def box_hd(title, glyph=None, small=False, when=None):
    """A navy header bar. `when` is the right-aligned date / room line."""
    return ('<div class="box-hd%s">%s%s%s</div>'
            % (" small" if small else "",
               ('<span class="cg" aria-hidden="true">%s</span>' % glyph)
               if glyph else "",
               esc(title_case(title)),
               ('<span class="when">%s</span>' % esc(when)) if when else ""))


def render_cat_card(cat, groups, filtered=None, gmods=None):
    body = []
    for i, g in enumerate(groups):
        items = None if filtered is None else filtered[i]
        if items is not None and not items:
            continue
        body.append(render_group(g, cat, items,
                                 gmods[i] if gmods else None))
    if not body:
        return ""
    return ('<section class="cat %s">%s<div class="cat-bd">%s</div></section>'
            % (cat, box_hd(CAT_TITLE[cat], CAT_GLYPH[cat]), "".join(body)))


def cat_cards(groups, wnum, where, only_module=None):
    """Group a week's groups by category and emit the cards in calendar
    order. `only_module` keeps just that module's items (module pages)."""
    buckets, others = {}, []
    for i, g in enumerate(groups):
        cat = g.get("cat", "other")
        gm = group_modules(g, wnum, where, i)
        if only_module is not None:
            keep = [it for it in g["items"]
                    if only_module in item_modules(it, gm)]
            if not keep:
                continue
        else:
            keep = g["items"]
        (others if cat == "other" else buckets.setdefault(cat, [])).append(
            (g, keep, gm))

    out = []
    for cat, _, _ in CAT_ORDER:
        if cat not in buckets:
            continue
        out.append(render_cat_card(cat, [g for g, _, _ in buckets[cat]],
                                   [k for _, k, _ in buckets[cat]],
                                   [m for _, _, m in buckets[cat]]))
    for g, keep, gm in others:
        out.append('<div class="cat other"><div class="cat-bd">%s</div></div>'
                   % render_group(g, "other", keep, gm))
    return "".join(out)


# ============================== navigation ==============================

def nav_week_label(w):
    kind = w["kind"]
    if kind == "midterm":
        return "Midterm Exam"
    if kind == "final":
        return "Final Exam"
    if kind == "examprep":
        return "Exam Preparation"
    if kind == "thanksgiving":
        return "Practice · Thanksgiving"
    mods = week_modules(w)
    if len(mods) == 1:
        return "Module %d" % mods[0]
    if mods:
        return "Modules %d – %d" % (min(mods), max(mods))
    return title_case(w["kind"])


def band_center(w):
    """What kind of week this is -- no dates (2026-09-03, Nico).
       Weeks with both read "Video Content & On-Campus Class"."""
    kind = w["kind"]
    if kind == "midterm":
        return "Midterm Exam"
    if kind == "final":
        return "Final Exam"
    video, campus = has_video_groups(w), kind == "oncampus"
    if video and campus:
        return "Video Content &amp; On-Campus Class"
    if campus:
        return "On-Campus Class"
    if video:
        return "Video Content"
    if kind == "examprep":
        return "Exam Preparation"
    return ""


def left_column(current):
    # The on-campus swatch mirrors the on-campus CARD -- a navy rule over a
    # transparent navy wash -- rather than the card's header alone
    # (2026-09-03, Nico).
    legend = "".join(
        '<div><span class="sw %s" style="%s"></span><span>%s</span></div>'
        % (cls, style, esc(title_case(label)))
        for cls, style, label in (
            ("campus", "", "On-campus class"),
            ("", "background:#F6E8C9", "Video content"),
            ("", "background:#E09F3E", "Exam")))

    def util(href, key, title, sub, glyph, external=False):
        """General Logistics, All Videos, All Podcasts, All Practice
        Exercises -- everything outside the week / module sequence."""
        # an empty sub-line renders nothing at all, rather than an empty row
        return ('<li><a href="%s" class="gl%s"%s%s>'
                '<span class="n gly" aria-hidden="true">%s</span>'
                '<span class="t">%s%s</span></a></li>'
                % (href, " out" if external else "",
                   ' target="_blank" rel="noopener"' if external else "",
                   ' aria-current="page"' if current == key else "",
                   glyph, esc(title),
                   ('<span class="d">%s</span>' % esc(sub)) if sub else ""))

    # General Logistics opens both lists; the two indexes close them
    # (2026-09-03, Nico).
    # no sub-line: "Before You Start" was dropped (2026-09-04, Nico)
    head = [util(GL_HREF, "index", "General Logistics", "", "◆")]
    tail = [util(href, key, title, sub, glyph)
            for href, key, title, sub, glyph in EXTRA_PAGES]
    tail += [util(href, key, title, sub, glyph, external=True)
             for href, key, title, sub, glyph in EXTRA_LINKS]

    # Each row leads with "Week 1" / "Module 1" and carries the dates and the
    # coverage underneath (2026-09-03, Nico).
    weeks = list(head)
    for w in C.WEEKS:
        d1, d2 = week_span(w)
        cur = current == "week-%d" % w["num"]
        weeks.append(
            '<li><a href="week-%02d.html" data-kind="%s"%s>'
            '<span class="n">%d</span><span class="t">Week %d'
            '<span class="d">%s &nbsp;·&nbsp; %s</span></span></a></li>'
            % (w["num"], w["kind"], ' aria-current="page"' if cur else "",
               w["num"], w["num"], esc(C.span(d1, d2)),
               esc(nav_week_label(w))))

    mods = list(head)
    for num, _title, short, _parts in MODULES:
        wks = [w["num"] for w in C.WEEKS if num in week_modules(w)]
        cur = current == "module-%d" % num
        kind = ("oncampus"
                if any(week_of(k)["kind"] == "oncampus" for k in wks)
                and not any(week_of(k)["kind"] == "deadline" for k in wks)
                else "deadline")
        mods.append(
            '<li><a href="module-%d.html" data-kind="%s"%s>'
            '<span class="n">%d</span><span class="t">Module %d'
            '<span class="d">%s &nbsp;·&nbsp; %s</span></span></a></li>'
            % (num, kind, ' aria-current="page"' if cur else "", num, num,
               esc(short),
               esc(("Week " if len(wks) == 1 else "Weeks ") +
                   ", ".join(str(k) for k in wks))))

    weeks += tail
    mods += tail

    return """<aside class="left">
  <div class="card">%s<div class="legend">%s</div></div>
  <div class="toggle" role="group" aria-label="Organize the course by">
    <button type="button" id="t-weeks" aria-pressed="true">By Week</button>
    <button type="button" id="t-mods" aria-pressed="false">By Module</button>
  </div>
  <nav class="list card" id="nav-weeks" aria-label="Weeks"><ul>%s</ul></nav>
  <nav class="list card" id="nav-mods" aria-label="Modules" hidden><ul>%s</ul></nav>
</aside>""" % (box_hd("Color coding", small=True), legend,
               "".join(weeks), "".join(mods))


# ======================= deadlines / video-watch rows =======================

def _ranges(nums):
    """[1,2,3,5] -> '1 – 3, 5'"""
    nums = sorted(set(nums))
    out, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        out.append(str(nums[i]) if j == i
                   else "%d – %d" % (nums[i], nums[j]))
        i = j + 1
    return ", ".join(out)


def video_watch_row(w):
    """The videos a week actually requires, as one deadline row -- e.g.
    "Watch Videos 1 - 4 for Module 1". Optional groups (the recaps, the
    "(optional)" practice sets) are left out; a group marked "(required)"
    is counted even though its items are practice videos (2026-09-03, Nico:
    "the videos beyond week 1 are missing").

    The deadline is the Friday class on an on-campus week, and the week's
    Sunday otherwise -- the same "suggested" date the calendar prints."""
    n = w["num"]
    numbered, practice = {}, {}

    for g in w.get("prep_groups") or []:
        if g.get("cat") != "video":
            continue
        label = g.get("label") or ""
        low = label.lower()
        if "optional" in low or low.startswith("recap"):
            continue
        label_mods = mods_in(label)
        for it in g["items"]:
            if it[0] != "v":
                continue
            text = it[2]
            found = mods_in(text) or label_mods
            if not found:
                continue
            mod = found[0]
            if re.search(r"Practice Video", text):
                practice[mod] = practice.get(mod, 0) + 1
                continue
            m = re.search(r"Video\s+(\d+)", text)
            if m:
                numbered.setdefault(mod, []).append(int(m.group(1)))

    if not numbered and not practice:
        return None

    parts = []
    for mod in sorted(set(list(numbered) + list(practice))):
        bits = []
        if mod in numbered:
            bits.append("Videos %s" % _ranges(numbered[mod]))
        if mod in practice:
            c = practice[mod]
            bits.append("%d practice video%s" % (c, "s" if c > 1 else ""))
        parts.append("%s for Module %d" % (" and ".join(bits), mod))

    due = C.dt(n, "Fri" if w["kind"] == "oncampus" else "Sun")
    return {"week": n, "label": "Watch " + "; ".join(parts),
            "note": None, "date": due.isoformat(),
            "when": "by %d/%d" % (due.month, due.day),
            "exam": False, "watch": True}


def mail_link(address, text, cls=""):
    """An email link that address harvesters cannot read.

    The address is split at the "@" and each half base64-encoded into a data
    attribute, so the rendered HTML contains no address and no "@" anywhere.
    `initMail()` in site.js reassembles it and sets the href on load. Without
    JavaScript the element is an anchor with no href -- inert text rather
    than a broken link (2026-09-04, Nico)."""
    local, _, domain = address.partition("@")

    def enc(part):
        return base64.b64encode(part.encode("utf-8")).decode("ascii")

    return ('<a class="mail%s" data-u="%s" data-d="%s">%s</a>'
            % ((" " + cls) if cls else "", enc(local), enc(domain), esc(text)))


def ta_link(text, cls=""):
    """The TA's name as a mail link -- or as plain text when the section has
    no mailbox yet, rather than routing students to the other section's TA
    (2026-09-05, FEMBA's address is still to come)."""
    if not C.TA_EMAIL:
        return esc(text)
    return mail_link(C.TA_EMAIL, text, cls)


def help_body():
    """Who to ask about what. Only "TA <name>" is the link, bold and
    underlined so it reads as one (2026-09-04, Nico); Prof. Nico's line
    deliberately carries no address."""
    return ('<div class="qlinks first">'
            '<div class="row"><span class="g" aria-hidden="true">%s</span>'
            # a no-break space keeps "to" attached to the name, so the
            # sentence can only wrap BEFORE "to" -- never between the two
            # (2026-09-04, Nico)
            '<span>All questions about Problem Sets, Practice Exams, '
            'Online Practice Problems etc. go to&nbsp;&nbsp;%s</span></div>'
            '<div class="row"><span class="g" aria-hidden="true">%s</span>'
            '<span>Conceptual questions about class material go to&nbsp;&nbsp;'
            '%s</span></div>'
            '</div>'
            % (HELP_MAIL,
               ta_link("TA " + C.TA_NAME, "who"),
               HELP_MAIL,
               mail_link(NICO_EMAIL, "Prof. Nico", "who")))


def jump_select(current):
    """The hamburger jump menu. It carries every page, so on a phone it can
    stand in for the whole left menu. A native <select> is deliberate:
    phones give it their own full-screen picker. `data-g` lets site.js show
    only the weeks or only the modules, following the View by button."""
    def opt(value, label, key, group):
        return ('<option value="%s" data-g="%s"%s>%s</option>'
                % (value, group, " selected" if current == key else "",
                   esc(label)))

    weeks = "".join(
        opt("week-%02d.html" % w["num"],
            "Week %d  ·  %s" % (w["num"], C.span(*week_span(w))),
            "week-%d" % w["num"], "weeks")
        for w in C.WEEKS)
    mods = "".join(
        opt("module-%d.html" % num, "Module %d  ·  %s" % (num, short),
            "module-%d" % num, "mods")
        for num, _t, short, _p in MODULES)

    return ('<span class="jumpwrap">'
            '<span class="jumpicon" aria-hidden="true"><i></i><i></i><i></i>'
            '</span>'
            '<select class="jumpsel" id="jump" '
            'aria-label="Jump to a week or module">%s%s%s'
            '</select></span>'
            % (opt(GL_HREF, "General Logistics", "index", "general"),
               weeks, mods
               + "".join(opt(href, title, key, "extra")
                         for href, key, title, _s, _g
                         in EXTRA_PAGES + EXTRA_LINKS)))


def assessments():
    """Problem sets, exams and pre-class video viewing, in date order."""
    out = []
    for w in C.WEEKS:
        row = video_watch_row(w)
        if row:
            out.append(row)
        for label, dw, dwd, note in (w.get("due") or []):
            d = C.dt(dw, dwd) if dw else None
            out.append({"week": w["num"], "label": label, "note": note,
                        "date": d.isoformat() if d else "9999",
                        "when": C.fmt(d, wd=True) if d else "t.b.a.",
                        "exam": False, "watch": False})
        if w.get("exam"):
            ex = w["exam"]
            (wd0, off0), (wd1, off1) = ex["window"]
            d0 = C.dt(w["num"] + off0, wd0)
            d1 = C.dt(w["num"] + off1, wd1)
            out.append({"week": w["num"], "label": ex["title"], "note": None,
                        "date": d0.isoformat(),
                        "when": "%s – %s" % (C.fmt(d0, wd=True),
                                             C.fmt(d1, wd=True)),
                        "exam": True, "watch": False})
    out.sort(key=lambda a: a["date"])
    return out


def week_has_deadlines(n):
    """Does this week have any row in the Deadlines & Exams column?"""
    return any(a["week"] == n for a in assessments())


def right_column(current_week):
    rows = []
    marked = False
    for a in assessments():
        cls = []
        if a["exam"]:
            cls.append("ex")
        if a.get("watch"):
            cls.append("watch")
        if current_week == a["week"]:
            cls.append("here")
        # the date rides along so the browser can flag what is due within
        # the next three days against TODAY (2026-09-03, Nico)
        # the first row of the current week is the phone link's target
        anchor_id = ""
        if current_week == a["week"] and not marked:
            anchor_id = ' id="dl-here"'
            marked = True
        rows.append(
            '<li%s%s data-week="%d" data-date="%s"><span class="w">Week %d</span>'
            '<span class="lb"><a href="week-%02d.html">%s</a>%s</span>'
            '<time>%s</time></li>'
            % ((' class="%s"' % " ".join(cls)) if cls else "", anchor_id,
               a["week"], "" if a["date"] == "9999" else a["date"],
               a["week"], a["week"], esc(a["label"]),
               (' <span style="font-size:12.5px;color:var(--ink-3)">(%s)</span>'
                % esc(a["note"])) if a["note"] else "",
               esc(a["when"])))
    return """<aside class="right" id="sidebar">
  <div class="searchwrap">
    <span class="mag" aria-hidden="true">⌕</span>
    <input type="search" id="q" placeholder="Search…"
           aria-label="Search the course" autocomplete="off">
    <div class="results" id="results" hidden></div>
  </div>
  <div class="card" id="deadlines">%s
    <div class="dl-legend"><span class="sw-due" aria-hidden="true"></span>
      <span>Upcoming within 3 days</span></div>
    <ul class="dl">%s</ul>
    <button class="dl-all" id="dl-all" type="button">Show all deadlines</button>
  </div>
</aside>""" % (box_hd("Deadlines & exams", small=True), "".join(rows))


# ============================== page shell ==============================

def back_link(href, label):
    if not href:
        return '<div class="back"></div>'
    return ('<div class="back"><a href="%s">'
            '<span class="ar" aria-hidden="true">←</span>%s</a></div>'
            % (href, esc(label)))


def next_link(href, label):
    if not href:
        return ""
    return ('<div class="turn"><a href="%s">%s'
            '<span class="ar" aria-hidden="true">→</span></a></div>'
            % (href, esc(label)))


def page(fname, page_title, nav_kind, current, main_html,
         current_week=None, head_extra=""):
    html = """<!DOCTYPE html>
<!-- light mode only (2026-09-03, Nico): the palette is
     stamped here so nothing can flip it -->
<html lang="en" data-theme="light">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<!-- the tab always reads the course, never the page (2026-09-03, Nico) -->
<title>%(tab)s</title>
%(head_extra)s<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&amp;family=Source+Sans+3:wght@400;600;700&amp;display=swap">
<link rel="stylesheet" href="assets/site.css?v=__ASSETV__">
<script src="assets/search-index.js?v=__ASSETV__"></script>
<script src="assets/site.js?v=__ASSETV__"></script>
</head>
<body data-navkind="%(navkind)s">
<div class="topbar"><div class="topbar-in">
  <div class="tb-left">%(jump)s<a class="tb-tag" href="%(gl)s">%(site)s</a></div>
  <span class="tb-sub"><a href="%(nico)s" target="_blank" rel="noopener">%(inst)s</a> &nbsp;·&nbsp; %(school)s</span>
  <div class="tb-right">
    <div class="helpwrap">
      <button class="helpbtn" id="helpbtn" type="button" aria-expanded="false"
              aria-controls="helppop" aria-label="Help and questions"><span class="q" aria-hidden="true">%(helpq)s</span><span class="m" aria-hidden="true">%(helpmail)s</span><span class="tip" aria-hidden="true">Contact the TA or Prof.</span></button>
      <div class="helppop" id="helppop" hidden>
        <div class="box-hd"><span class="cg" aria-hidden="true">%(helphd)s</span>Help and Questions</div>
        <div class="panel-bd">%(helpbody)s</div>
      </div>
    </div>
    <button class="viewbtn" id="viewmode" type="button"><span class="l1">View by</span><span class="l2" id="viewmode-label">Week</span><span class="tip" id="viewmode-tip">Switch to view by Module</span></button>
  </div>
</div></div>

<div class="shell">
%(left)s
<main>
%(main)s
</main>
%(right)s
<button class="sidebtn" id="sidebtn" type="button" aria-expanded="false"
        aria-controls="sidebar">Deadlines / Search</button>
</div>
</body>
</html>
""" % {
        "ptitle": esc(page_title),
        "tab": esc(TAB_TITLE),
        "helpq": HELP_Q,
        "helpmail": HELP_MAIL,
        "helphd": HELP_HD_GLYPH,
        "helpbody": help_body(),
        "site": esc(SITE_NAME),
        "navkind": nav_kind,
        "gl": GL_HREF,
        "inst": esc("Prof. Nico Voigtländer"),
        "school": "UCLA Anderson",
        "nico": esc(NICO_URL),
        "jump": jump_select(current),
        "left": left_column(current),
        "right": right_column(current_week),
        "main": main_html,
        "head_extra": head_extra,
    }
    with io.open(os.path.join(OUT, fname), "w", encoding="utf-8",
                 newline="\n") as f:
        f.write(html)


# ============================== week pages ==============================

def inclass_material(w):
    """A handout and a slide deck for every module the on-campus class
    covers, listed under the class's own module list. Both are "(TBD)" until
    Nico uploads them right before class (2026-09-03).

    The modules come from the week's TOPICS as well as the class card's own
    items. Week 9's card names Modules 7 and 8, but the class also covers
    Module 6 -- its Applications sit behind the Discussion item -- and the
    topics list catches that (2026-09-04, Nico)."""
    mods = []
    for t in w.get("topics") or []:
        for n in mods_in(t):
            if n not in mods:
                mods.append(n)
    for g in (w.get("weekend") or {}).get("groups") or []:
        for it in g["items"]:
            txt = it[1] if it[0] in ("t", "b", "note") else (
                it[2] if it[0] in ("v", "l", "p") else "")
            for n in mods_in(txt):
                if n not in mods:
                    mods.append(n)
    return sorted(mods)


def tbd_rows(mods, kinds):
    """Placeholder bullets -- "Module 3 – Handout (TBD)" -- one per module
    and kind. Nico uploads the real files right before each class."""
    rows = []
    for n in mods:
        for what in kinds:
            rows.append('<li><span class="g" aria-hidden="true">▤</span>'
                        '<span class="txt">Module %d – %s '
                        '<span class="tba">(TBD)</span></span></li>'
                        % (n, what))
    return "".join(rows)


def materials_box(w):
    """The week's downloadable material, in a box of its own at the foot of
    the page (2026-09-05, Nico). The In-Class Material list used to sit as a
    group inside the on-campus card.

    Same container format as Before Class / During the Week."""
    # On-campus weeks only (2026-09-06, Nico). The video decks used to have
    # a "Slides for Videos" rubric of their own here; they now sit on the
    # video's own bullet in the Videos card, so what is left is purely what
    # is handed out in class -- hence the name, and no inner sub-heading to
    # repeat it.
    if w["kind"] != "oncampus":
        return ""
    mods = inclass_material(w)
    if not mods:
        return ""
    body = ('<div class="grp"><ul class="items">%s</ul></div>'
            % tbd_rows(mods, ("Handout", "Slides")))
    return ('<div class="prep materials"><div class="head">%s</div>'
            '<section class="cat other"><div class="cat-bd">%s</div>'
            '</section></div>' % (esc("Slides from Class"), body))


def week_main(w):
    n = w["num"]
    d1, d2 = week_span(w)
    last = C.WEEKS[-1]["num"]
    h = []

    # Week 1 has no previous week, and both menus already offer General
    # Logistics, so it carries no back link at all (2026-09-03, Nico).
    back = ((None, None) if n == 1
            else ("week-%02d.html" % (n - 1), "Week %d" % (n - 1)))

    # On a phone the deadlines column sits below the content, so the header
    # offers a way down to this week's rows (2026-09-03, Nico). Hidden on
    # wider screens, where the column is right there beside the content.
    anchor = "#dl-here" if week_has_deadlines(n) else "#deadlines"
    h.append('<div class="band">'
             '<div class="who"><h1>Week %d</h1><span class="sp">%s</span></div>'
             '<div class="center">%s</div>%s'
             '<div class="dl-jump"><a href="%s" data-week="%d">'
             'Deadlines for this week '
             '<span aria-hidden="true">↓</span></a></div>'
             '</div>'
             % (n, esc(C.span(d1, d2)), band_center(w),
                back_link(back[0], back[1]), anchor, n))

    h.append('<div class="body">')

    for label, dw, dwd, note in (w.get("due") or []):
        when = ("Due: %s" % C.fmt(C.dt(dw, dwd), wd=True)) if dw else ""
        # Every problem set says where the solution goes (2026-09-03, Nico).
        # The practice final is not a problem set, so it is left alone.
        is_pset = label.lower().startswith("problem set")
        upload = ""
        if is_pset:
            upload = ('<p class="upload">Upload one solution per group on '
                      '<a href="%s" target="_blank" rel="noopener">BruinLearn'
                      '</a></p>' % BRUINLEARN_COURSE)
        # Only a problem set takes the dark-red treatment; the practice final
        # is not one, so it keeps the gold rule (2026-09-03, Nico).
        h.append('<div class="pcard due%s"><div class="pcard-bd">'
                 '<span class="g" aria-hidden="true">✎</span>'
                 '<div class="lead"><b>%s</b>%s%s</div>%s</div></div>'
                 % (" pset" if is_pset else "",
                    esc(label), (" — %s" % esc(note)) if note else "",
                    upload, ("<time>%s</time>" % esc(when)) if when else ""))

    if w.get("exam"):
        ex = w["exam"]
        (wd0, off0), (wd1, off1) = ex["window"]
        w0 = C.fmt(C.dt(n + off0, wd0), wd=True)
        w1 = C.fmt(C.dt(n + off1, wd1), wd=True)
        lines = "".join(
            '<li><span class="g" aria-hidden="true">•</span>'
            '<span class="txt">%s</span></li>'
            % esc(ln.format(w0=w0, w1=w1)) for ln in ex["lines"])
        h.append('<section class="pcard exam"><div class="pcard-hd">'
                 '<span>%s</span><span class="when">%s – %s</span></div>'
                 '<div class="pcard-bd"><ul class="items">%s</ul></div>'
                 '</section>'
                 % (esc(ex["title"]), esc(w0), esc(w1), lines))
    else:
        h.append('<section class="pcard topics">%s<div class="pcard-bd">'
                 '<ul>%s</ul></div></section>'
                 % (box_hd("Topics covered"),
                    "".join("<li>%s</li>" % esc(t) for t in w["topics"])))

    def class_card():
        we = w["weekend"]
        body = "".join(render_group(g, "other") for g in we["groups"])
        if w["kind"] == "oncampus":
            when = ("%s · Room %s"
                    % (C.class_when(n).replace("   ·   ", " · "),
                       C.CLASSROOM))
            # the classical building marks the in-person part of the week
            # (2026-09-03, Nico)
            return ('<section class="pcard class">%s<div class="pcard-bd">%s'
                    '</div></section>'
                    % (box_hd("On-campus class", "\U0001F3DB\uFE0F",
                              when=when), body))
        da, db = we["days"]
        when = ("suggested deadline: %s / %s"
                % (C.fmt(C.dt(n, da), wd=True), C.fmt(C.dt(n, db), wd=True)))
        # This branch is dormant -- only weeks 1, 5 and 9 carry a weekend
        # block and all three are on-campus -- but it takes the same
        # clapperboard as the Videos category, for completeness (2026-09-04).
        return ('<section class="pcard videos">%s<div class="pcard-bd">%s'
                '</div></section>'
                % (box_hd("Videos to watch", CAT_GLYPH["video"], when=when),
                   body))

    class_first = bool(w.get("weekend")) and w["kind"] == "oncampus"
    if class_first:
        h.append(class_card())

    if w.get("prep_groups"):
        head = ""
        if w.get("prep_days"):
            a, b = w["prep_days"]
            label = "Before class" if w["kind"] == "oncampus" else "During the week"
            head = ('<div class="head">%s<span class="when"> · %s – %s'
                    '</span></div>'
                    % (esc(title_case(label)),
                       esc(C.fmt(C.dt(n, a), wd=True)),
                       esc(C.fmt(C.dt(n, b), wd=True))))
        h.append('<div class="prep">%s%s</div>'
                 % (head, cat_cards(w["prep_groups"], n, "prep")))

    if w.get("weekend") and not class_first:
        h.append(class_card())

    h.append(materials_box(w))

    if w.get("holiday"):
        ho = w["holiday"]
        (wd0, off0), (wd1, off1) = ho["window"]
        h.append('<section class="pcard holiday">%s<div class="pcard-bd">%s'
                 ' (%s – %s)</div></section>'
                 % (box_hd("Holiday"), esc(ho["text"]),
                    esc(C.fmt(C.dt(n + off0, wd0), wd=True)),
                    esc(C.fmt(C.dt(n + off1, wd1), wd=True))))

    h.append("</div>")
    if n < last:
        h.append(next_link("week-%02d.html" % (n + 1), "Week %d" % (n + 1)))
    return "".join(h)


# ============================== module pages ==============================

def module_main(num, title, parts):
    wks = [w for w in C.WEEKS if num in week_modules(w)]
    h = []

    # Module 1 has no previous module, and both menus already offer General
    # Logistics, so it carries no back link -- as week 1 does not
    # (2026-09-04, Nico).
    back = ((None, None) if num == 1
            else ("module-%d.html" % (num - 1), "Module %d" % (num - 1)))

    h.append('<div class="band">'
             '<div class="who"><h1>Module %d</h1><span class="sp">%s</span></div>'
             '<div class="center">%s %s</div>%s%s</div>'
             % (num, esc(title), "Week" if len(wks) == 1 else "Weeks",
                ", ".join(str(w["num"]) for w in wks),
                back_link(back[0], back[1]),
                ('<p class="sub">%s</p>' % esc(" · ".join(parts)))
                if parts else ""))

    campus = [w["num"] for w in wks if w["kind"] == "oncampus"]
    video = [w["num"] for w in wks if w["kind"] == "deadline"]
    pills = []
    if video:
        pills.append('<span class="pill video">Video · week%s %s</span>'
                     % ("s" if len(video) > 1 else "",
                        ", ".join(str(k) for k in video)))
    if campus:
        pills.append('<span class="pill class">In class · week%s %s</span>'
                     % ("s" if len(campus) > 1 else "",
                        ", ".join(str(k) for k in campus)))

    h.append('<div class="body">')
    if pills:
        h.append('<div class="pills">%s</div>' % "".join(pills))

    for w in wks:
        d1, d2 = week_span(w)
        cards = cat_cards(w.get("prep_groups") or [], w["num"], "prep",
                          only_module=num)
        we = w.get("weekend")
        if we:
            cards += cat_cards(we["groups"], w["num"], "weekend",
                               only_module=num)
        topics = [t for t in w["topics"] if num in mods_in(t)]
        # a week's topic line often repeats what the on-campus card already
        # says ("Module 3: Applications"), so drop the duplicates
        shown = set(re.findall(r'class="txt">([^<]+)<', cards))
        topics = [t for t in topics if t not in shown]
        if not cards and not topics:
            continue
        h.append('<section class="mwk"><div class="mwk-hd">'
                 '<a href="week-%02d.html">Week %d</a>'
                 '<span class="d">%s · %s</span></div>%s%s</section>'
                 % (w["num"], w["num"], esc(C.span(d1, d2)),
                    band_center(w).replace("&amp;", "&"),
                    ('<p class="mwk-topics">%s</p>'
                     % esc(" · ".join(topics))) if topics else "",
                    cards))
    h.append("</div>")

    if num < MODULES[-1][0]:
        h.append(next_link("module-%d.html" % (num + 1),
                           "Module %d" % (num + 1)))
    return "".join(h)


def module_media(num, cat):
    """[(week, group, items)] for one module and one category, in week
    order -- the raw material for the All Videos / All Podcasts indexes."""
    out = []
    for w in C.WEEKS:
        for i, g in enumerate(w.get("prep_groups") or []):
            if g.get("cat") != cat:
                continue
            gm = group_modules(g, w["num"], "prep", i)
            items = [it for it in g["items"] if num in item_modules(it, gm)]
            if items:
                out.append((w, g, items))
    return out


def coursewide_media(cat):
    """The groups of one category that belong to no single module."""
    out = []
    for w in C.WEEKS:
        for i, g in enumerate(w.get("prep_groups") or []):
            if g.get("cat") != cat:
                continue
            if group_modules(g, w["num"], "prep", i):
                continue
            out.append((w, g, g["items"]))
    return out


def media_card(title, glyph, cat, blocks, mod=None):
    """One module's worth of videos or podcasts, in the category card style,
    with each source group's own lead-in kept and its week named."""
    body = []
    for w, g, items in blocks:
        lab = g.get("label") or ""
        body.append('<div class="grp"><p class="grp-lab">%s'
                    '<span class="wk">Week %d</span></p>'
                    '<ul class="items">%s</ul></div>'
                    % (esc(lab), w["num"],
                       "".join(render_item(it, cat, [mod] if mod else None)
                               for it in items)))
    if not body:
        return ""
    return ('<section class="cat %s">%s<div class="cat-bd">%s</div></section>'
            % (cat, box_hd(title, glyph), "".join(body)))


def media_index(kind_title, subtitle, cat, glyph):
    """The All Videos and All Podcasts pages: one card per module, in module
    order, then anything that belongs to no module (2026-09-03, Nico)."""
    h = ['<div class="band"><div class="who"><h1>%s</h1>'
         '<span class="sp">%s</span></div>'
         '<div class="center">By Module</div>%s</div>'
         % (esc(kind_title), esc(subtitle), back_link(None, None))]
    h.append('<div class="body">')
    for num, title, _short, _parts in MODULES:
        blocks = module_media(num, cat)
        if blocks:
            # pass the module, so a video whose title does not name it
            # ("Video 2: The Production Function") still finds its deck
            h.append(media_card("Module %d: %s" % (num, title), glyph, cat,
                                blocks, num))
        else:
            # Say so rather than silently skipping the module -- an absent
            # card reads like an omission (modules 5 and 8 have no videos).
            campus = [w["num"] for w in C.WEEKS
                      if num in week_modules(w) and w["kind"] == "oncampus"]
            why = ("This module is covered in class, in week %s."
                   % ", ".join(str(k) for k in campus)) if campus else                   "Nothing posted for this module yet."
            h.append('<section class="cat other">%s<div class="cat-bd">'
                     '<p class="none-yet">%s</p></div></section>'
                     % (box_hd("Module %d: %s" % (num, title)), esc(why)))
    wide = coursewide_media(cat)
    if wide:
        h.append(media_card("Not tied to one module", glyph, cat, wide))
    h.append("</div>")
    return "".join(h)


# ========================= general logistics page =========================

def panel(title, body, glyph=None):
    return ('<div class="panel">%s<div class="panel-bd">%s</div></div>'
            % (box_hd(title, glyph), body))


def logistics_main():
    keep_books = [t for t in C.TEXTBOOK_NOTES
                  if not any(d in t for d in DROP_TEXTBOOK_NOTES)]
    books = "".join("<li>%s</li>" % esc(t) for t in keep_books)

    keep_math = [it for it in C.MATH_REFRESHER_ITEMS
                 if not any(d in "".join(str(s) for s in it)
                            for d in DROP_MATH_ITEMS)]
    m_items = "".join("<li>%s</li>" % render_segments(it) for it in keep_math)

    h = []
    # "Before You Start" dropped from the band and the term set larger
    # (2026-09-04, Nico). The menu row still carries the sub-line.
    h.append('<div class="band">'
             '<div class="who"><h1>General Logistics</h1></div>'
             '<div class="center big">%s</div><div class="back"></div></div>'
             % esc(C.TERM))
    h.append('<div class="body">')

    # Two half-width boxes, side by side, above the three columns.
    h.append('<div class="gl-row">%s%s</div>' % (
        # Both PDFs sit in one box (2026-09-04, Nico): the syllabus first,
        # the calendar under it. The calendar has no card of its own any
        # more -- the website IS the calendar, and the PDF is the version
        # for those who want it on paper.
        panel("Class Syllabus and Calendar",
              '<div class="qlinks first">'
              '<a href="%s" target="_blank" rel="noopener">'
              '<span class="g" aria-hidden="true">&#9636;</span>'
              '<span class="u">Download the Class Syllabus here</span></a>'
              '<a href="%s" target="_blank" rel="noopener">'
              '<span class="g" aria-hidden="true">&#9636;</span>'
              '<span class="u">Download the Class Calendar as PDF here</span>'
              '</a></div>' % (SYLLABUS_URL, CALENDAR_PDF_URL),
              "▤"),
        panel("BruinLearn Class Site",
              '<div class="qlinks first">'
              '<a href="%s" target="_blank" rel="noopener">'
              '<span class="g" aria-hidden="true">&#9636;</span>'
              '<span class="u">Open the BruinLearn class site</span></a>'
              '</div>' % BRUINLEARN_COURSE,
              "&#127891;")))

    h.append('<div class="gl-grid">')

    # "How the Quarter Runs" was dropped and Class and Contact promoted to
    # the top left (2026-09-04, Nico).
    col1 = panel("Class and Contact",
                  "<dl>"
                   "<dt>Meets</dt><dd>%s<br>Room <b>%s</b></dd>"
                   "<dt>Instructor</dt><dd>%s</dd>"
                   "<dt>TA</dt><dd>%s</dd>"
                   "<dt>Term</dt><dd>%s</dd>"
                  "</dl>"
                  % (esc(C.CLASS_TIMES), esc(C.CLASSROOM),
                     mail_link(NICO_EMAIL, "Prof. Nico Voigtl\u00e4nder"),
                     ta_link(C.TA_NAME),
                     esc(C.TERM)))

    # Column 2: watching the videos, then the practice exercises directly
    # beneath it. Column 3: the math refresher above the textbook. Each
    # column is one grid cell, so neither pair can come apart.
    col2 = panel(
        "Watching the Videos",
        "<p>%s</p>"
        '<div class="qlinks">'
        '<a href="%s" target="_blank" rel="noopener">'
        '<span class="g" aria-hidden="true">&#128444;</span>'
        '<span class="u">How to sign in (screenshot)</span></a></div>'
        % (render_segments(C.SIGNIN_NOTE), PANOPTO_SHOT),
        "\U0001F3AC")
    col2 += panel(
        "Online Practice Exercises",
        "<p>Interactive practice exercises for each Module, with hints and "
        "step-by-step solutions.</p>"
        '<div class="qlinks">'
        '<a href="%s" target="_blank" rel="noopener">'
        '<span class="g" aria-hidden="true">&#9998;</span>'
        '<span class="u">Open the practice exercises</span></a></div>'
        % esc(PRACTICE_INDEX),
        "✎")

    # Column 3 opens with Class and Contact, so it sits in the top right
    # corner (2026-09-03, Nico), then the math refresher above the textbook.
    # The top right is now "who to ask about what" (2026-09-04, Nico). The
    # TA's line carries his address; Prof. Nico's deliberately does not.
    col3 = panel("Help and Questions", help_body(), HELP_HD_GLYPH)
    col3 += panel("Math Refresher",
                  "<p>%s</p><ul>%s</ul>"
                  % (esc(C.MATH_REFRESHER_INTRO), m_items),
                  "∂")
    col3 += panel("Textbook", "<ul>%s</ul>" % books, "\U0001F4D6")

    # Two balanced columns of three. The right-hand column opens with Class
    # and Contact, so it sits in the top right corner (2026-09-03, Nico);
    # within each column the pairs Nico asked for stay together -- practice
    # under videos, textbook under the math refresher.
    h.append('<div class="panel-stack">%s%s</div>'
             '<div class="panel-stack">%s</div>'
             % (col1, col2, col3))

    h.append("</div></div>")
    h.append(next_link("week-01.html", "Week 1"))
    return "".join(h)


def redirect_script():
    """On the site root, land on the CURRENT week once the quarter is
    running; before it starts and after it ends, stay here. The sidebar's
    General Logistics link carries ?stay so the page stays reachable."""
    rows = []
    for w in C.WEEKS:
        d1, d2 = week_span(w)
        rows.append('["%s","%s","week-%02d.html"]'
                    % (d1.isoformat(), d2.isoformat(), w["num"]))
    return ("<script>\n"
            "(function(){\n"
            "  if (location.search.indexOf('stay') !== -1) { return; }\n"
            "  var W=[%s];\n"
            "  var d=new Date(), p=function(n){return (n<10?'0':'')+n;};\n"
            "  var t=d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate());\n"
            "  for (var i=0;i<W.length;i++){\n"
            "    if (t>=W[i][0] && t<=W[i][1]) "
            "{ location.replace(W[i][2]); return; }\n"
            "  }\n"
            "}());\n"
            "</script>\n" % ",".join(rows))


# ============================== search index ==============================

def plain(html):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html)).strip()


def build_index(pages):
    rows = []
    for href, kind, title, sub, html in pages:
        hay = plain(html).lower().replace("–", "-").replace("—", "-")
        head = " ".join([kind.lower(), title.lower(), (sub or "").lower()])
        rows.append({"href": href, "kind": kind, "title": title, "sub": sub,
                     "head": head, "hay": head + " " + hay})
    body = json.dumps(rows, ensure_ascii=False, separators=(",", ":"))
    with io.open(os.path.join(OUT, "assets", "search-index.js"), "w",
                 encoding="utf-8", newline="\n") as f:
        f.write("/* GENERATED by _build_site.py -- do not edit by hand. */\n")
        f.write("window.SEARCH_INDEX = %s;\n" % body)


# ============================== main ==============================

def check_overrides():
    """Every GROUP_MODULE_OVERRIDES key must still match exactly one group.

    The table used to be keyed by a group's index within its week, so adding
    or deleting a group re-pointed the overrides after it without a word --
    which is how week 3's advanced reading and week 6's Module 4 practice
    video quietly fell off their module pages (2026-09-05)."""
    hits = {k: 0 for k in GROUP_MODULE_OVERRIDES}
    for w in C.WEEKS:
        for where, groups in (("prep", w.get("prep_groups") or []),
                              ("weekend",
                               (w.get("weekend") or {}).get("groups") or [])):
            for g in groups:
                hay = group_text(g)
                for key in GROUP_MODULE_OVERRIDES:
                    if key[0] == w["num"] and key[1] == where and key[2] in hay:
                        hits[key] += 1
    bad = {k: n for k, n in hits.items() if n != 1}
    if bad:
        for k, n in bad.items():
            print("  OVERRIDE %r matches %d groups (want 1)" % (k, n))
        sys.exit("group-module overrides are stale -- fix them and rebuild")


def stamp_assets():
    """Put a content hash on the asset URLs.

    Without it a browser can pair freshly published HTML with a site.css it
    already has cached -- which is exactly how the help tooltip shipped
    broken on 2026-09-05: the new markup arrived, the CSS that positions it
    did not, and the tip rendered as wrapped text inside the button. The
    stamp goes in as a post-pass because search-index.js is only written
    after the pages that reference it."""
    h = hashlib.sha1()
    for name in ("site.css", "site.js", "search-index.js"):
        with io.open(os.path.join(OUT, "assets", name), "rb") as fh:
            h.update(fh.read())
    v = h.hexdigest()[:10]
    n = 0
    for f in glob.glob(os.path.join(OUT, "*.html")):
        s = io.open(f, encoding="utf-8").read()
        if "__ASSETV__" not in s:
            continue
        io.open(f, "w", encoding="utf-8", newline="\n").write(
            s.replace("__ASSETV__", v))
        n += 1
    print("  asset version %s stamped into %d pages" % (v, n))


def main():
    os.makedirs(os.path.join(OUT, "assets"), exist_ok=True)
    check_overrides()
    if OUT != SRC:
        # site.css and site.js are hand-authored ONCE, in Course Website/.
        # Copy them in on every build so a second section can never drift
        # from the stylesheet Nico actually edits (2026-09-05).
        for _a in ("site.css", "site.js"):
            shutil.copy2(os.path.join(SRC, "assets", _a),
                         os.path.join(OUT, "assets", _a))
    if not os.path.exists(PANOPTO_SHOT_SRC):
        sys.exit("missing sign-in screenshot: %s" % PANOPTO_SHOT_SRC)
    shutil.copy2(PANOPTO_SHOT_SRC, os.path.join(OUT, PANOPTO_SHOT))

    # The video slide decks, under their published (space-free) names. These
    # are build OUTPUT here -- the originals live in 405 Slide Revisions
    # 2026/ -- so .gitignore keeps them out of the private repo, while
    # _deploy.py ships the folder to the public one (2026-09-06).
    sl = os.path.join(OUT, "slides")
    os.makedirs(sl, exist_ok=True)
    keep = set()
    for src in C.VIDEO_SLIDES.values():
        name = C.slides_pub_name(src)
        keep.add(name)
        dst = os.path.join(sl, name)
        if (not os.path.exists(dst)
                or os.path.getmtime(dst) < os.path.getmtime(src)):
            shutil.copy2(src, dst)
    for gone in set(os.listdir(sl)) - keep:
        os.remove(os.path.join(sl, gone))
    print("  slides/                 (%d decks)" % len(keep))
    pages = []

    gl = logistics_main()
    page("index.html", "General Logistics", "", "index", gl,
         head_extra=redirect_script())
    pages.append((GL_HREF, "General", "General Logistics",
                  "Before You Start", gl))

    for w in C.WEEKS:
        d1, d2 = week_span(w)
        html = week_main(w)
        fname = "week-%02d.html" % w["num"]
        page(fname, "Week %d" % w["num"], "weeks", "week-%d" % w["num"],
             html, current_week=w["num"])
        pages.append((fname, "Week %d" % w["num"],
                      w["topics"][0] if w["topics"] else nav_week_label(w),
                      C.span(d1, d2), html))

    for href, key, title, sub, glyph in EXTRA_PAGES:
        cat = "video" if key == "all-videos" else "podcast"
        html = media_index(title, sub, cat, CAT_GLYPH[cat])
        page(href, title, "", key, html)
        pages.append((href, title, title, sub, html))

    for num, title, _short, parts in MODULES:
        html = module_main(num, title, parts)
        fname = "module-%d.html" % num
        page(fname, "Module %d" % num, "mods", "module-%d" % num, html)
        wks = [w["num"] for w in C.WEEKS if num in week_modules(w)]
        pages.append((fname, "Module %d" % num, title,
                      "Weeks " + ", ".join(str(k) for k in wks), html))

    build_index(pages)
    stamp_assets()

    print("built %d pages" % len(pages))
    print("  index.html, week-01..week-%02d, module-1..module-%d"
          % (C.WEEKS[-1]["num"], MODULES[-1][0]))
    print("  assets/search-index.js  (%d entries)" % len(pages))
    watch = [a for a in assessments() if a.get("watch")]
    print("  pre-class video rows in the deadlines list: %d" % len(watch))
    for a in watch:
        print("     week %-2d  %s  %s" % (a["week"], a["label"], a["when"]))


if __name__ == "__main__":
    main()
