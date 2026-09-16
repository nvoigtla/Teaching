# -*- coding: utf-8 -*-
"""
Content source for the MGMT 405 EMBA Hybrid course calendar.

ALL dates derive from ONE anchor: ANCHOR_FRIDAY, the Friday of the first
on-campus class weekend. To roll the calendar to a new year, change that
single line (and any content edits), then re-run _build_calendar.py.

Week 1 = the calendar week (Mon-Sun) containing the first on-campus Friday.
Every date in the document is expressed as (week_number, weekday).
"""

import glob
import os
import re
from datetime import date, datetime, timedelta

# ============================== DATE ENGINE ==============================

ANCHOR_FRIDAY = date(2026, 9, 25)   # first on-campus Friday  <-- change yearly
TERM = "Fall 2026"

_WD = {"Mon": 0, "Tue": 1, "Wed": 2, "Thu": 3, "Fri": 4, "Sat": 5, "Sun": 6}

_WEEK1_MONDAY = ANCHOR_FRIDAY - timedelta(days=4)


def dt(week, weekday):
    """Date of `weekday` ('Mon'..'Sun') in course week `week` (1-based)."""
    return _WEEK1_MONDAY + timedelta(weeks=week - 1, days=_WD[weekday])


def fmt(d, wd=False):
    """'Oct 13' or 'Tue, Oct 13'."""
    s = f"{d.strftime('%b')} {d.day}"
    return f"{d.strftime('%a')}, {s}" if wd else s


def span(d1, d2):
    """'Sep 21 – 27' or 'Sep 28 – Oct 4' (month repeated only if it changes).

    A one-day span collapses to the single date: the final exam became a
    fixed Saturday slot on 2026-09-10, and "Dec 12 – 12" is not a span."""
    if d1 == d2:
        return fmt(d1)
    if d1.month == d2.month:
        return f"{d1.strftime('%b')} {d1.day} – {d2.day}"
    return f"{fmt(d1)} – {fmt(d2)}"


# ---- exam clock ----
# An exam is normally a WINDOW the student picks a start time inside, given
# as ("window": weekday, week-offset) pairs. The final exam is different
# from 2026-09-10: it has a fixed "slot" -- ((9, 0), (12, 0)) -- and these
# three helpers are what every renderer reads it through, so the time is
# written down once.

def clock(h, m):
    """(9, 0) -> '9:00 AM'."""
    return "%d:%02d %s" % (h % 12 or 12, m, "AM" if h < 12 else "PM")


def slot_label(ex):
    """'9:00 AM – 12:00 PM' for an exam with a fixed slot, '' otherwise."""
    s = ex.get("slot")
    return "%s – %s" % (clock(*s[0]), clock(*s[1])) if s else ""


def _pacific_offset(d):
    """Hours Pacific time runs behind UTC on `d`. US rule: DST from the 2nd
    Sunday in March to the 1st Sunday in November. Computed rather than
    hard-coded to 8 -- the current slot sits in December, but a slot moved
    into the spring would otherwise be published an hour late."""
    def nth_sunday(year, month, n):
        first = date(year, month, 1)
        s = first + timedelta(days=(6 - first.weekday()) % 7)
        return s + timedelta(weeks=n - 1)
    return 7 if nth_sunday(d.year, 3, 2) <= d < nth_sunday(d.year, 11, 1) else 8


def exam_utc(wk):
    """(DTSTART, DTEND) as UTC iCalendar stamps for a fixed-slot exam, or
    None. Written in UTC on purpose: it is unambiguous in every calendar
    client and needs no VTIMEZONE block kept in step with the tz database."""
    ex = wk.get("exam") or {}
    if not ex.get("slot"):
        return None
    (wd0, off0), _ = ex["window"]
    d = dt(wk["num"] + off0, wd0)
    off = _pacific_offset(d)
    return tuple(
        (datetime(d.year, d.month, d.day, h, m)
         + timedelta(hours=off)).strftime("%Y%m%dT%H%M%SZ")
        for h, m in ex["slot"])


# ============================== VIDEO SLIDES ==============================
# The slide deck behind each recorded video. DISCOVERED by scanning the
# module folders, not listed here, so a deck Nico drops into a "Videos Final"
# folder shows up on the next build with no code change (2026-09-06).
#
# The scan requires this filename convention:
#     Module <M> - Video <N> - <whatever>.pptx
# and keys on (module, video number) -- the same numbering the calendar uses
# for its video bullets, so a deck finds its video even where the Panopto
# link is still missing. `python _publish.py` reports any file it could not
# parse, and any video still without a deck.

SLIDES_ROOT = os.path.abspath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), os.pardir, os.pardir,
    "405 Slide Revisions 2026"))


def _scan_slides():
    found = {}
    if not os.path.isdir(SLIDES_ROOT):
        return found                      # another machine, or moved
    for folder in sorted(glob.glob(os.path.join(SLIDES_ROOT, "Module *",
                                                "Videos Final"))):
        for name in sorted(os.listdir(folder)):
            m = re.match(r"Module (\d+) - Video (\d+) - (.+)\.pptx$", name)
            if m:
                found[(int(m.group(1)), int(m.group(2)))] = \
                    os.path.join(folder, name)
    return found


VIDEO_SLIDES = _scan_slides()


def unparsed_slides():
    """.pptx files in a Videos Final folder that the convention above does
    not match -- so a mis-named deck is reported rather than ignored."""
    bad = []
    if not os.path.isdir(SLIDES_ROOT):
        return bad
    for folder in sorted(glob.glob(os.path.join(SLIDES_ROOT, "Module *",
                                                "Videos Final"))):
        for name in sorted(os.listdir(folder)):
            if name.lower().endswith(".pptx") and not re.match(
                    r"Module (\d+) - Video (\d+) - (.+)\.pptx$", name):
                bad.append(os.path.join(folder, name))
    return bad


def slides_pub_name(path):
    """Published file name: no spaces, since it is served straight off
    GitHub Pages. " - " collapses to a single hyphen, so
    "Module 1 - Video 1 - Introduction.pptx" publishes as
    "Module-1-Video-1-Introduction.pptx"."""
    base = os.path.basename(path)
    return re.sub(r"\s*-\s*", "-", base).replace(" ", "-")


def slides_for(module, title):
    """Absolute path of the deck behind one video bullet, or None.

    PRACTICE videos have no deck of their own, and their titles contain
    "Video 1" too, so they are excluded explicitly."""
    if module is None or "Practice Video" in title:
        return None
    m = re.search(r"Video (\d+)", title)
    if not m:
        return None
    return VIDEO_SLIDES.get((module, int(m.group(1))))


# ============================== SECTIONS ==============================
# MGMT 405 runs twice in Fall 2026 with the SAME material and the same due
# dates (2026-09-05, Nico). Only the meeting pattern, the room, the TA
# mailbox, the BruinLearn course and the published names differ, so both
# sections come out of this one file.
#
# Pick one with the MGMT405_SECTION environment variable; every builder
# also takes "--section femba" and sets it before importing this module.

SECTIONS = {
    "emba": {
        "label": "EMBA",
        "section_name": "EMBA Section 2",
        # EMBA is A301, FEMBA G305 (2026-09-08, Nico -- the two were
        # briefly the other way round earlier the same day)
        "classroom": "A301",
        # (weekday, time) per meeting of an on-campus weekend, in order
        "meetings": (("Fri", "4:00 – 5:30 pm"),
                     ("Sat", "9:00 am – 12:30 pm")),
        "class_times": "Fridays 4:00 – 5:30 pm  ·  Saturdays 9:00 am – 12:30 pm",
        "ta_email": "ta405.emba2@gmail.com",
        "bruinlearn_course": "https://bruinlearn.ucla.edu/courses/237825",
        # the class-recording tool, whose id differs per section
        "recordings_tool": "10996",
        "repo": "MGMT-405-EMBA",
        "calendar_docx": "Calendar EMBA Hybrid -- Fall 2026",
        "syllabus_docx": "Course Syllabus - 405 EMBA Fall 2026",
    },
    "femba": {
        "label": "FEMBA",
        "section_name": "FEMBA Section 2",
        "classroom": "G305",
        # FEMBA meets on the SAME three Saturdays as EMBA, but only on the
        # Saturday, and for a long afternoon (2026-09-05, Nico).
        "meetings": (("Sat", "2:00 – 8:00 pm"),),
        "class_times": "Saturdays 2:00 – 8:00 pm",
        "ta_email": "ta405.femba2@gmail.com",
        "bruinlearn_course": "https://bruinlearn.ucla.edu/courses/237860",
        # the class-recording tool, whose id differs per section
        "recordings_tool": "10995",
        "repo": "MGMT-405-FEMBA",
        "calendar_docx": "Calendar FEMBA Hybrid -- Fall 2026",
        "syllabus_docx": "Course Syllabus - 405 FEMBA Fall 2026",
    },
}

SECTION = os.environ.get("MGMT405_SECTION", "emba").lower()
if SECTION not in SECTIONS:
    raise SystemExit("unknown MGMT405_SECTION %r (want one of %s)"
                     % (SECTION, ", ".join(sorted(SECTIONS))))
SEC = SECTIONS[SECTION]

SECTION_LABEL = SEC["label"]
# "EMBA Section 2" -- the stored name, and the two strings built from it.
# Keeping the name rather than the finished subtitle means next year's edit
# is one string, not three (2026-09-13).
SECTION_NAME = SEC["section_name"]
SECTION_TITLE = "%s, %s" % (SECTION_NAME, TERM)   # "EMBA Section 2, Fall 2026"
SUBTITLE_TAIL = "%s (Hybrid)" % SECTION_NAME      # the documents' subtitle
REPO = SEC["repo"]
SITE_BASE = "https://nvoigtla.github.io/%s" % REPO
CALENDAR_DOCX = SEC["calendar_docx"]
SYLLABUS_DOCX = SEC["syllabus_docx"]
MEETINGS = SEC["meetings"]
TA_EMAIL = SEC["ta_email"]


def class_when(week):
    """The on-campus header line for a week: every meeting with its date and
    time, e.g. "Fri, Sep 25, 4:00 – 5:30 pm   ·   Sat, Sep 26, 9:00 am –
    12:30 pm", or just the Saturday for FEMBA."""
    return "   ·   ".join("%s, %s" % (fmt(dt(week, wd), wd=True), t)
                          for wd, t in MEETINGS)


def class_dates(week):
    """'Sep 25/26' -- an on-campus weekend's dates with NO weekday names.
    Split out of class_days_line so the website's "On campus" row and the
    calendar's sub-line cannot disagree (2026-09-13)."""
    days = [dt(week, wd) for wd, _ in MEETINGS]
    return fmt(days[0]) + "".join("/%d" % d.day for d in days[1:])


def class_days_line(week):
    """The agenda table's sub-line: "class: Fri/Sat Sep 25/26", or
    "class: Sat Sep 26" for a section that meets once."""
    days = [dt(week, wd) for wd, _ in MEETINGS]
    names = "/".join(d.strftime("%a") for d in days)
    return "class: %s %s" % (names, class_dates(week))


def oncampus_dates():
    """Every on-campus weekend, read off WEEKS rather than listed.

    A section that meets TWICE gets bare dates -- 'Sep 25/26' already says
    "a Friday and a Saturday" on its own. A section that meets ONCE keeps the
    weekday, because a lone 'Sep 26' does not say which day it is
    (2026-09-13, Nico, on the FEMBA box)."""
    one = len(MEETINGS) == 1
    out = []
    for w in WEEKS:
        if w["kind"] != "oncampus":
            continue
        d = class_dates(w["num"])
        if one:
            d = "%s %s" % (dt(w["num"], MEETINGS[0][0]).strftime("%a"), d)
        out.append(d)
    return out


# ============================== HYPERLINKS ==============================

LINKS = {
    # filled in from the section table below -- FEMBA has its own mailbox
    "ta_email":   "mailto:%s" % (TA_EMAIL or ""),
    "math_quiz":  "https://bruinlearn.ucla.edu/courses/195707?invitation=0GYbOXWd6mBK4bwem1dgAP7Jq2ad1PZo5Wp4kTYQ",
    "math_videos": "https://bruinlearn.ucla.edu/courses/195707/pages/econ-math-review-videos",
    # asm_panopto (dropbox screenshot) retired 2026-08-15: the screenshot is now
    # embedded on p.2 (Images/Panopto-Login-Picture.png) and the link points to
    # the Panopto site itself (Nico's hand-edit).
    "panopto_site": "https://ucla-anderson.hosted.panopto.com",
    # Module 1 -- videos re-recorded and re-uploaded as NEW Panopto
    # sessions 2026-08-28 (ids b4b3...; the old b1d9... ids were last
    # year's recordings). Running times still unmeasured -> "(++)".
    "m1v1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=b7207fac-c4f6-45fd-be9c-b4b800fc0e3f",  # re-uploaded 2026-09-01 (was 45ebea5c-...a768)
    "m1v2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=ea13ec98-9412-41fc-accd-b4b30104a761",
    "m1v3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=7ad4fc39-0070-4aa6-90a9-b4b30104a761",
    "m1v4": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=8bd7e577-22b2-4da6-888d-b4b30104a76a",
    "recap1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=2f3c674d-13da-4308-ada2-b08b0134d2ff",
    # Podcasts
    "pod_cb":  "https://podcasts.apple.com/us/podcast/core-principle-1-the-cost-benefit-principle-the-pros-cons/id1523898793?i=1000786324917",
    "pod_oc":  "https://podcasts.apple.com/us/podcast/core-principle-2-the-opportunity-cost-principle-or-what/id1523898793?i=1000786324632",
    "pod_freak": "https://freakonomics.com/podcast/should-we-really-behave-like-economists-say-we-do-a-new-freakonomics-radio-podcast/",
    "pod_tlae": "https://art19.com/shows/think-like-an-economist/episodes/d48893ba-1f44-43d6-a271-be306d55d0f9?fbclid=IwAR14eZxCKFr8FOC5jQ3ZvGia0WQZRs-zytfYxEUjfWsPlj8aKRzIF8zpmYQ",
    # Optional listening, added 2026-09-14. Each sits in the week BEFORE its
    # module is taught on campus. pod_lux is a VIDEO, not a podcast -- it is
    # listed with the podcasts because that is where a student looks for
    # optional extras, and its title says so.
    "pod_lux":     "https://www.youtube.com/watch?v=gDdq2rIqAlM",
    "pod_minwage": "https://freakonomics.com/podcast/the-true-story-of-the-minimum-wage-fight-ep-460/",
    "pod_rent":    "https://freakonomics.com/podcast/why-rent-control-doesnt-work/",
    "pod_groupon": "https://www.npr.org/sections/money/2011/04/08/135248177/the-friday-podcast-groupon-monty-python-price-discrimination",
    "pod_pd":      "https://www.npr.org/sections/money/2018/05/30/615622421/episode-844-nice-game",
    # Replaced 2026-09-15: this pointed at BruinLearn course 218078,
    # LAST year's site, which this year's students cannot open. The
    # Freakonomics episode needs no enrolment.
    "pod_penalty": "https://freakonomics.com/podcast/why-the-world-cup-is-an-economists-dream/",
    "pod_13000":   "https://www.npr.org/2020/01/09/794977811/episode-963-13-000-economists-1-question",
    # Module 2 -- videos 1-3 re-recorded and re-uploaded as NEW Panopto
    # sessions 2026-09-04 (ids b4bb...; the old b08b... ids were last
    # year's recordings). Running times read straight off Panopto the
    # same day: 25 / 23 / 24 min. The two PRACTICE videos were not
    # retaped and keep last year's sessions.
    "m2v1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=8fcb86ee-57e8-48ec-906d-b4bb0136917f",
    "m2v2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=406d26d1-60f9-4169-afd0-b4bb01369180",
    "m2v3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=885b4e89-7ffe-4305-8490-b4bb01369174",
    "m2p1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=d9a1abf9-e8e5-448b-bbbd-b08b012eeeb0",
    "m2p2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=cf0b4650-7a8b-41a0-9a6e-b08b012eeeb9",
    "recap2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=4979325b-0d4a-45d8-9401-b08b0134db2a",
    # Module 3 -- retaped as SEVEN videos and re-uploaded as new
    # Panopto sessions 2026-09-12. The six old b08b... sessions were
    # last year's, from before the retape. The keys now line up 1:1
    # with the bullet numbers: m3v1 IS Video 1, the introduction.
    "m3v1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=87a3e4bf-9f6b-47c5-be34-b4c10122bf71",
    "m3v2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=c6a66635-de5c-476d-a84e-b4c10122bf74",
    "m3v3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=20321bcb-3200-432d-af71-b4c10122bf79",
    "m3v4": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=90b82d2f-ae07-4af3-998c-b4c10122bf72",
    "m3v5": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=dc8e017e-0298-4984-ab06-b4c201290033",
    "m3v6": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=bd200584-fd9f-4c16-9d54-b4c2012941b0",
    "m3v7": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=db7b63c7-bc91-4099-a43a-b4bf01516e00",
    "recap3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=7803e078-6264-46e3-bdd6-b08b0134ff01",
    "m3pa": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=147753a0-3a02-40e0-a75c-b08b012f36b3",
    "m3pb": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=047b0460-fd8b-4a4e-8996-b08b012f5514",
    # Module 4 -- re-recorded as FIVE videos and re-uploaded as new
    # Panopto sessions 2026-09-09 (b4bf... ids). The old four b08b...
    # sessions were last year's, from before the module was re-split.
    "m4v1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=cbcea934-fb25-4602-ab96-b4bf014fb013",
    "m4v2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=3c840b96-8cb5-41d3-ad2c-b4bf014fb006",
    "m4v3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=96eeafea-a3a1-42cd-a705-b4bf014fb006",
    "m4v4": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=bb4df35b-7e1b-487d-9794-b4bf014fb00e",
    "m4v5": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=6be99363-bc5c-44d5-91b4-b4bf014ff951",
    "m4p1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=d9c62849-b585-46d1-9907-b1da01624e8d",
    # Module 6
    "m6v1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=aa860c73-b471-4d05-b4b1-b08b0132e83f",
    "m6v2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=72ee2af5-d9c8-4174-8a89-b08b0132e83f",
    "m6v3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=9c1bffb7-a64a-4c17-9fba-b08b0132e84d",
    "m6v4": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=d977a6ef-6102-4bd3-9066-b08b013312b6",
    "m6v5": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=520deb00-03a7-4c1d-b6ec-b08b013328d7",
    "m6v6": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=b3e080e7-fe4b-428c-b100-b08b013344f0",
    "m6v7": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=7185463d-c133-4f56-8e92-b08b01336154",
    "m6v8": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=3136caa5-3096-4826-8bd9-b08b01337eab",
    "m6p1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=c36ab9a6-b4ac-4c17-aa81-b08b0132e83e",
    "recap6": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=f594c324-d460-447f-992b-b08b0135435e",
    # Module 7
    "m7v1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=dd1b45a7-f553-49db-9150-b08b0133e5af",
    "m7v2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=39e72765-4f6f-4d7c-ad92-b08b0133e57f",
    "m7v3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=b3a9978a-5dff-4fa4-b26d-b08b0133ff7b",
    "m7v4": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=df2723c7-d63a-4908-a0ab-b08b01341319",
    "m7p1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=db03cc71-7075-47a5-a5c7-b08b01344881",
    "m7p2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=1c65c0ce-2e5a-4527-9078-b08b0133e583",
    "recap7": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=71beb745-5142-4a78-84f2-b08b01355a2c",
    "m7adv": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=055b647a-0998-4ab9-92cf-b08b0133e57c",
    # Teaching notes -- PDFs published next to the site by _deploy.py, so
    # the names here must match its DOCS list. They carry no year: the notes
    # are written to be reusable from one year to the next.
    "tn_mr":    "%s/MGMT-405-Teaching-Note-Marginal-Revenue.pdf" % SITE_BASE,
    "tn_elast": "%s/MGMT-405-Teaching-Note-Demand-Elasticity-and-Total-Revenue.pdf" % SITE_BASE,
    "tn_reg":   "%s/MGMT-405-Teaching-Note-Regressions.pdf" % SITE_BASE,
    "tn_bfb":   "%s/MGMT-405-Teaching-Note-Bang-for-the-Buck-Rule.pdf" % SITE_BASE,
    "tn_hire":  "%s/MGMT-405-Teaching-Note-Hiring-Decisions-Short-Run.pdf" % SITE_BASE,
    "tn_mrmc":  "%s/MGMT-405-Teaching-Note-MR-MC.pdf" % SITE_BASE,
    # Practice exercises (TA Rafael Rubiao's site), added 2026-08-31.
    # index: https://rafaelrubiao.github.io/mgmt405-practice/index.html
    "prac_m1": "https://rafaelrubiao.github.io/mgmt405-practice/module-1.html",
    "prac_m2": "https://rafaelrubiao.github.io/mgmt405-practice/module-2.html",
    "prac_m3": "https://rafaelrubiao.github.io/mgmt405-practice/module-3.html",
    "prac_m4a": "https://rafaelrubiao.github.io/mgmt405-practice/module-4-part-1.html",
    "prac_m4b": "https://rafaelrubiao.github.io/mgmt405-practice/module-4-part-2.html",
    "prac_m5": "https://rafaelrubiao.github.io/mgmt405-practice/module-5.html",
    "prac_m6": "https://rafaelrubiao.github.io/mgmt405-practice/module-6.html",
    "prac_m7a": "https://rafaelrubiao.github.io/mgmt405-practice/module-7-part-1.html",
    "prac_m7b": "https://rafaelrubiao.github.io/mgmt405-practice/module-7-part-2.html",
    "prac_m8": "https://rafaelrubiao.github.io/mgmt405-practice/module-8.html",
    # The course website and the BruinLearn class site (2026-09-04). Both
    # live here rather than in a build script, so the calendar, the website
    # and the syllabus all read one address.
    # Derived from SITE_BASE, not hardcoded: it used to be the literal EMBA
    # address, so every FEMBA build printed the FEMBA address (WEBSITE_TEXT
    # is section-derived) but linked the EMBA site (2026-09-06).
    "website": SITE_BASE + "/",
    "bruinlearn_course": SEC["bruinlearn_course"],
    # The Assignments page, which is where the problem sets are
    # downloaded AND uploaded from 2026-09-12 (Nico: they are no
    # longer on the class website).  Derived from the section's own
    # course URL for the same reason "website" is derived from
    # SITE_BASE -- a hardcoded pair would link EMBA from a FEMBA
    # build.
    "bruinlearn_assignments": SEC["bruinlearn_course"] + "/assignments",
    # Where the recording of each on-campus class is posted afterwards
    # (2026-09-14). Built from the course URL above, so the course id
    # stays in one place.
    "bruinlearn_recordings": ("%s/external_tools/%s"
                              % (SEC["bruinlearn_course"],
                                 SEC["recordings_tool"])),
    # The two PDFs the website hosts, so the calendar, the site and the
    # syllabus can all link them (2026-09-04). File names carry no spaces
    # because they are served straight off GitHub Pages.
    "calendar_pdf": SITE_BASE + "/MGMT-405-Calendar-Fall-2026.pdf",
    "syllabus_pdf": SITE_BASE + "/MGMT-405-Syllabus-Fall-2026.pdf",
}

# ============================== HEADER / INTRO ==============================

COURSE_TITLE = "MGMT 405 – Managerial Economics"
SUBTITLE = f"Course Calendar – {TERM} – {SUBTITLE_TAIL}"
CALENDAR_NOTE = ("Please check the course website "
                 "for the latest version of this calendar.")
# The course website, carried at the top of page 1 (2026-09-04, Nico). The
# link text is the bare address, so it stays usable in print; WEBSITE_TEXT
# and LINKS["website"] have to be kept in step.
WEBSITE_LEAD = "Course website:"
WEBSITE_TEXT = "nvoigtla.github.io/%s" % REPO
# The BruinLearn course site, shown the same way: the FULL address,
# scheme stripped, so it reads as a printable link and can never drift
# from LINKS["bruinlearn_course"] (2026-09-15, Nico -- "include the
# full BL site link").
BRUINLEARN_TEXT = LINKS["bruinlearn_course"].split("://", 1)[-1]
SYLLABUS_NOTE = ("Please check the course website "
                 "for the more detailed Class Syllabus.")
TA_NAME = "Rafael Rubiao"
CLASSROOM = SEC["classroom"]
CLASS_TIMES = SEC["class_times"]

TEXTBOOK_NOTES = [
    "“Microeconomics” by Goolsbee, Levitt, Syverson (4th edition). "
    "See the different purchase options in the class syllabus.",
    "The “Figure it out” boxes in each chapter should be considered "
    "as advanced (voluntary) readings.",
]

# How the course grade is made up. It lived only in _build_syllabus.py
# until 2026-09-13, when the website started showing it too -- two copies of
# 35/40/25 in two files is how a wrong number reaches students.
GRADE_WEIGHTS = (("Midterm Exam", "35%"),
                 ("Final Exam", "40%"),
                 ("Problem Sets", "25%"))

MATH_REFRESHER_INTRO = ("If you feel you could use a math refresher before "
                        "the class, I suggest:")
MATH_REFRESHER_ITEMS = [
    [("t", "Take the "), ("l", "math_quiz", "Math Quiz"), ("t", " (10 – 15 min)")],
    [("t", "Watch the "), ("l", "math_videos", "Math Review Videos")],
]

# The CALENDAR's line, inside the "Watching the course videos" card.
SIGNIN_NOTE = [("t", "This may require you to sign-in to Panopto. "
                     "Use option "),
               ("l", "panopto_site", "“ASM Panopto”")]

# The WEBSITE says it in two points instead (2026-09-14, Nico), and the
# whole point is the link -- so this is a plain string, not segments.
SIGNIN_WEB_TEXT = "You may have to sign-in to Panopto"

# ============================== PODCASTS ==============================
# Two NotebookLM Audio Overviews per module, hosted on Dropbox:
#   Module-<N>-Podcast-Intro.m4a     -- preview, listened to BEFORE
#   Module-<N>-Podcast-Wrap-Up.m4a   -- recap, listened to AFTER
#
# The FILE NAME follows that convention, but a Dropbox "scl/fi" share
# URL also carries a random per-file id and a random rlkey, so a new
# module's two links CANNOT be derived from the module number -- paste
# them in below as each module is uploaded.
#
# minutes: read straight off the audio file by _podcast_minutes.py
# (run it after pasting new links; it prints the numbers to copy here).
# (None, None) = not uploaded yet -> the bullet renders as plain text,
# with no link and no duration.
PODCASTS = {
    1: {"intro": ("https://www.dropbox.com/scl/fi/mghl9davv7sy6a7qeja20/Module-1-Podcast-Intro.m4a?rlkey=cof0c2gtljwud3kbtnoq87ryh&st=93b14l1z&dl=0", 6),
        "wrap":  ("https://www.dropbox.com/scl/fi/e010i3je3iaqdo0a6785l/Module-1-Podcast-Wrap-Up.m4a?rlkey=yqwuudwo4bqosg9yvu6ojawsy&st=4d59ssha&dl=0", 16)},
    2: {"intro": ("https://www.dropbox.com/scl/fi/w3ib5t12sxbekrtja6co7/Module-2-Podcast-Intro.m4a?rlkey=nujut9txx5zfyn8yrd76lfdju&st=rkmi76yc&dl=0", 4),
        "wrap":  ("https://www.dropbox.com/scl/fi/o2p9rnwnb276apb1jbfis/Module-2-Podcast-Wrap-Up.m4a?rlkey=m8tjqy0qnk5ts42etrljgd3ps&st=0dzo4j5u&dl=0", 20)},
    3: {"intro": ("https://www.dropbox.com/scl/fi/fto3ouj9gcum5gjs8128b/Module-3-Podcast-Intro.m4a?rlkey=vcpxzwex5r4ad6dm09nrr72qh&st=6cepogg8&dl=0", 7),
        "wrap":  ("https://www.dropbox.com/scl/fi/0geb6malz4suoruqbnp24/Module-3-Podcast-Wrap-Up.m4a?rlkey=j51911drl9t8173ep5zx4w3cy&st=gy4ocrqp&dl=0", 23)},
    4: {"intro": ("https://www.dropbox.com/scl/fi/k0vaqb6u7yej6sjagrauv/Module-4-Podcast-Intro.m4a?rlkey=foc9ajsc0e8wmzsq7ttiyobgd&st=3rivyeff&dl=0", 4),
        "wrap":  ("https://www.dropbox.com/scl/fi/n2ie2rqaud9n0u0477cs6/Module-4-Video-Wrap-Up.m4a?rlkey=zikuktu87plh008kil4yekt25&st=yieacety&dl=0", 22)},
    5: {"intro": (None, None), "wrap": (None, None)},
    # Module 6 uploaded 2026-09-16 (Nico). The wrap-up file is named
    # "Video-Wrap-Up" upstream, like Module 4's -- it is the audio
    # episode all the same.
    6: {"intro": ("https://www.dropbox.com/scl/fi/sa5u2pfaww9z93zbokkjy/Module-6-Podcast-Intro.m4a?rlkey=8212u4pxd41e7hcwn40xtd0xy&st=57qwpxsx&dl=0", 4),
        "wrap":  ("https://www.dropbox.com/scl/fi/5cmioelt39kxuqjy0c0ca/Module-6-Video-Wrap-Up.m4a?rlkey=e7wefrxfv9fuswqwt6oxpb1no&st=pcur8yel&dl=0", 20)},
    7: {"intro": (None, None), "wrap": (None, None)},
    8: {"intro": (None, None), "wrap": (None, None)},
}


def podcast_items(*specs):
    """Podcast bullets as ("p", url, text, minutes) items; url and minutes
    are None until the episode has been uploaded.

    Each spec is either a module number (both episodes, which is what the
    module pages of the website want) or a (module, kind) pair with kind
    "intro" or "wrap".

    The WEEKS below schedule the two episodes separately (2026-09-05, Nico):
    an intro has to be heard BEFORE the first section of its module is
    taught, a wrap-up AFTER the last one. Module 2 is the clearest case --
    its first section is taught in the week 1 class, so its intro sits in
    week 1 alongside Module 1's, while both wrap-ups sit in week 2."""
    TEXT = {"intro": "Podcast: Intro to Module %d",
            "wrap": "Podcast: Wrap-Up of Module %d"}
    out = []
    for spec in specs:
        mod, kinds = (spec, ("intro", "wrap")) if isinstance(spec, int) \
            else (spec[0], (spec[1],))
        ep = PODCASTS.get(mod, {})
        for kind in kinds:
            url, mins = ep.get(kind, (None, None))
            out.append(("p", url, TEXT[kind] % mod, mins))
    return out


# When each WRAP-UP episode should be listened to -- the phrase after the
# underlined "after" (2026-09-05, Nico). A module whose core material is all
# on video can be wrapped up in the same week as those videos; one taught in
# class waits until after the class.
#
#   "class"                 the module's last material is taught in class
#   "watching the ... "     all core material is on video; the on-campus
#                           session only does APPLICATIONS for that module
#   "class and watching ... " part I in class, part II on video (Module 2)
#
# An INTRO is always "before class".
WRAP_AFTER = {
    1: "class",                                     # remainder taught in wk 1 class
    2: "class and watching the Module 2 videos",    # part I in class, part II on video
    3: "watching the Module 3 videos",              # wk 5 class does Applications only
    # Week 4 holds ALL of Module 4's videos, so the wrap-up goes there
    # with them (2026-09-14, Nico). The week-5 class does Part II, but
    # Module 3 already set the precedent that later in-class work does
    # not hold the wrap-up back.
    4: "watching the Module 4 videos",
    5: "class",                                     # taught entirely in the wk 5 class
    6: "watching the Module 6 videos",              # wk 9 class does Applications only
    # Same case as Module 4: week 8 holds all of Module 7's content
    # videos, and the week-9 class does Part II (2026-09-14).
    7: "watching the Module 7 videos",
    8: "class",                                     # taught entirely in the wk 9 class
}


def podcast_when(text):
    """("before"|"after", rest of the phrase) for a module podcast bullet;
    None for anything that is not one of the two module episodes.

    Both builders call this, so the wording is written once. The first
    element is the word they underline."""
    rest = text[len("Podcast: "):] if text.startswith("Podcast: ") else text
    low = rest.lower()
    m = re.search(r"module\s+(\d+)", low)
    if "intro" in low:
        # A module taught entirely on video is previewed before those
        # videos, not before a class -- WRAP_AFTER already knows which
        # modules those are (2026-09-06, Nico).
        tail = WRAP_AFTER.get(int(m.group(1)), "class") if m else "class"
        return ("before", tail if tail.startswith("watching") else "class")
    if "wrap" in low and m:
        return ("after", WRAP_AFTER.get(int(m.group(1)), "class"))
    return None


# ============================== WEEKS ==============================
# item forms:
#   ("t", "plain text")                      plain bullet
#   ("b", "bold text")                       bold navy bullet (class topics)
#   ("v", linkkey|None, "Text", min|None)    video link + duration;
#                                            min None prints "(++)"
#   ("l", linkkey, "Link text")              plain link (podcasts, quiz, ...)
#   ("p", url|None, "Text", min|None)        module podcast (see PODCASTS)
#   ("note", "text")                         italic gray, no bullet
#   ("mix", [segments])                      segments as in MATH_REFRESHER_ITEMS
# group: {"label": "...", "items": [...]}    label is the italic navy lead-in
# due:   (label, week, weekday, note)        week/weekday None -> no date

WEEKS = [
    {
        "num": 1, "kind": "oncampus",
        "topics": ["Module 1: Basic Concepts and Economic Principles",
                   "Module 2: Demand Analysis"],
        "prep_days": ("Mon", "Fri"),
        "prep_groups": [
            {"cat": "video", "label": "Watch before class:",
             # names verbatim from the video title cards in
             # Module 1/_build_Module1.py (make_video_title); lengths
             # lengths read off Panopto 2026-08-28 by _video_minutes.py
             "items": [("v", "m1v1", "Module 1 – Video 1: Introduction", 9),  # 8:53, re-read 2026-09-01 after the re-upload
                       ("v", "m1v2", "Module 1 – Video 2: Markets", 10),
                       ("v", "m1v3", "Module 1 – Video 3: Demand and Supply", 8),
                       ("v", "m1v4", "Module 1 – Video 4: Market Equilibrium", 7)]},
            {"cat": "podcast", "label": "Podcasts About Class Material:",
             "items": podcast_items((1, "intro"), (2, "intro"))},
            {"cat": "podcast", "label": "Optional Podcasts (before class):",
             "items": [("l", "pod_freak",
                        "Should We Really Behave Like Economists Say We Do?"),
                       ("l", "pod_cb", "The Cost-Benefit Principle"),
                       ("l", "pod_oc", "The Opportunity-Cost Principle")]},
            {"cat": "read", "label": "In preparation for the Module 1 videos:",
             "items": [("t", "Ch. 1"),
                       ("t", "Ch. 2.1 – 2.4"),
                       ("t", "[Optional: Math Review Section 1 in the Appendix of the textbook]")]},
            {"cat": "read", "label": "In preparation for class:",
             "items": [("t", "Ch. 2.5")]},
        ],
        "weekend": {"days": ("Fri", "Sat"),
                    "groups": [
                        {"label": None,
                         "items": [("b", "Module 1 (remainder): Basic Concepts and Economic Principles"),
                                   ("b", "Module 2: Demand Analysis")]}]},
        "due": [],
    },
    {
        "num": 2, "kind": "deadline",
        "topics": ["Module 2: Demand Analysis (remaining videos)"],
        "prep_days": ("Mon", "Fri"),
        "prep_groups": [
            {"cat": "podcast", "label": "Podcasts About Class Material:",
             "items": podcast_items((1, "wrap"), (2, "wrap"))},
            {"cat": "video", "label": "Remaining videos for Module 2 [material not "
                      "covered in the on-campus class] – watch by the weekend:",
             # names verbatim from Module 2/Videos Final/ (and the
             # _video_title_slide calls in _build_Module2Video.py)
             # lengths read off Panopto 2026-09-04 (1527 / 1372 / 1412 s)
             "items": [("v", "m2v1", "Module 2 – Video 1: Elasticity and Revenue", 25),
                       ("v", "m2v2", "Module 2 – Video 2: Marginal Revenue", 23),
                       ("v", "m2v3", "Module 2 – Video 3: Demand Estimation", 24)]},
            {"cat": "video", "label": "Practice videos for Module 2:",
             "items": [("v", "m2p1", "Module 2 – Practice Video 1: Elasticity and Revenues", 7),
                       ("v", "m2p2", "Module 2 – Practice Video 2: Revenue Maximization", 6)]},
            {"cat": "read", "label": None,
             "items": [("note", "[Relevant textbook reading was already covered "
                                "in preparation for the previous class.]")]},
            {"cat": "read", "label": "Advanced reading (optional):",
             "items": [("t", "Chapters 5.1, 5.2, 5.4, 5.5")]},
            {"cat": "read", "label": "Teaching notes (optional):",
             "items": [("l", "tn_mr", "Teaching note: Marginal Revenue"),
                       ("l", "tn_elast",
                        "Teaching note: Demand Elasticity and Total Revenue"),
                       ("l", "tn_reg", "Teaching note: Regressions")]},
            # Practice exercises (TA site). Placement rule, confirmed
            # 2026-08-31: the week AFTER the module's own teaching week.
            # Modules 3 and 6 follow their VIDEO week, not the later
            # "Applications" week, so Module 3's set is available before
            # the Week 6 midterm (which covers through Module 3).
            {"cat": "practice", "label": None,
             "items": [("l", "prac_m1",
                        "Online quiz on Module 1: Basic Concepts and Economic Principles"),
                       ("l", "prac_m2",
                        "Online quiz on Module 2: Demand Analysis")]},
        ],
        "due": [],
    },
    {
        "num": 3, "kind": "deadline",
        "topics": ["Module 3: Production & Costs"],
        "prep_days": ("Mon", "Fri"),
        "prep_groups": [
            {"cat": "podcast", "label": "Podcasts About Class Material:",
             "items": podcast_items((3, "intro"), (3, "wrap"))},
            {"cat": "read", "label": "In preparation for the Module 3 videos:",
             "items": [("t", "Ch. 6.1 – 6.3, 6.5"),
                       ("t", "Ch. 7")]},
            {"cat": "video", "label": "Module 3: Production & Costs – videos to "
                      "watch by the weekend:",
             # 2026-08-28: Module 3 was retaped as SEVEN videos. Names
             # verbatim from M3_OUTLINE in Module 3/_m3_outline.py, the
             # single source of the deck's video title cards. All seven
             # were re-recorded and linked 2026-09-12, and the key number
             # now matches the video number on every line.
             "items": [("v", "m3v1", "Video 1: Introduction to Module 3", 3),
                       ("v", "m3v2", "Video 2: The Production Function", 8),
                       ("v", "m3v3", "Video 3: Short Run: Hiring Decisions", 31),
                       ("v", "m3v4", "Video 4: Wage Searchers", 6),
                       ("v", "m3v5", "Video 5: Long Run: The Optimal Input Mix", 19),
                       ("v", "m3v6", "Video 6: Cost Concepts", 33),
                       ("v", "m3v7", "Video 7: Economies of Scale and Scope", 13)]},
            {"cat": "read", "label": "Advanced reading (optional):",
             "items": [("t", "Ch. 6.6 and 6.7")]},
            {"cat": "read", "label": "Teaching notes (optional):",
             "items": [("l", "tn_hire",
                        "Teaching note: Hiring Decisions in the Short Run"),
                       ("l", "tn_bfb",
                        "Teaching note: The Bang-for-the-Buck Rule")]},
        ],
        "due": [("Problem Set 1", 4, "Tue", None)],
    },
    {
        "num": 4, "kind": "deadline",
        "topics": ["Module 4 (Part I): Competitive Markets and Market Interventions"],
        "prep_days": ("Mon", "Fri"),
        "prep_groups": [
            {"cat": "podcast", "label": "Podcasts About Class Material:",
             "items": podcast_items((4, "intro"), (4, "wrap"))},
            {"cat": "video", "label": "Practice on Module 3 (optional):",
             "items": [("v", "m3pa", "Practice Video: Costs: Make vs Buy Decision", 10),
                       ("v", "m3pb", "Practice Video: Short-Run and Long-Run Costs", 26)]},
            {"cat": "video", "label": "Module 4 (Part I): Competitive Markets and "
                      "Market Interventions – videos to watch by the weekend:",
             # Running times read off Panopto 2026-09-09.
             # 2026-08-30: Module 4 was re-split into FIVE videos when the
             # deck was converted for taping - Perfect Competition is now
             # a video of its own, and the module front matter sits inside
             # Video 1. All five were re-recorded and linked 2026-09-09.
             "items": [("v", "m4v1", "Video 1: Introduction to Market Structures", 6),
                       ("v", "m4v2", "Video 2: Perfect Competition", 8),
                       ("v", "m4v3", "Video 3: Profit Maximization of a Price Taker – Short Run", 45),
                       ("v", "m4v4", "Video 4: Firm-Level and Market Supply", 6),
                       ("v", "m4v5", "Video 5: Long-Run Competitive Equilibrium", 8)]},
            {"cat": "read", "label": "In preparation for the Module 4 (Part I) videos:",
             "items": [("t", "Ch. 8.1 – 8.3")]},
            {"cat": "practice", "label": None,
             "items": [("l", "prac_m3",
                        "Online quiz on Module 3: Production & Costs")]},
        ],
        "due": [],
    },
    {
        "num": 5, "kind": "oncampus",
        "topics": ["Module 3: Applications",
                   "Module 4 (Part II): Market Distortions / Externalities",
                   "Module 5: Monopoly and Monopolistic Competition"],
        "prep_days": ("Mon", "Fri"),
        "prep_groups": [
            {"cat": "podcast", "label": "Podcasts About Class Material:",
             "items": podcast_items((5, "intro"))},
            # The optional listening for every module taught in THIS
            # week's class (2026-09-14, Nico). The group spans Modules 4
            # and 5, so each title names its own -- item_modules() reads
            # the title ahead of the group. Luxottica is Module 5, the
            # eyewear near-monopoly named in Module 5's own decks.
            {"cat": "podcast", "label": "Optional Podcasts (before class):",
             "items": [("l", "pod_minwage",
                        "The True Story of the Minimum Wage Fight "
                        "(Module 4)"),
                       ("l", "pod_rent",
                        "Why Rent Control Doesn\u2019t Work (Module 4)"),
                       ("l", "pod_lux",
                        "Luxottica (Module 5) \u2013 note: this is a video, "
                        "not a podcast")]},
            {"cat": "read", "label": "In preparation for class:",
             "items": [("t", "For Module 4 (Part II): Ch. 3.1 – 3.4; Ch. 17 (pp. 513 – 524)"),
                       ("t", "For Module 5: Ch. 9.1 – 9.3; Ch. 9.5 – 9.7; Ch. 11.7")]},
            {"cat": "read", "label": "Teaching notes (optional):",
             "items": [("l", "tn_mrmc", "Teaching note: MR = MC")]},
            {"cat": "read", "label": None,
             "items": [("t", "Assigned articles for discussion (posted on BruinLearn)")]},
            {"cat": "practice", "label": None,
             "items": [("l", "prac_m4a",
                        "Online quiz on Module 4 (Part I): Competitive Markets and Market Interventions")]},
        ],
        "weekend": {"days": ("Fri", "Sat"),
                    "groups": [
                        {"label": None,
                         "items": [("b", "Module 3: Applications"),
                                   ("b", "Module 4 (Part II): Market Distortions / Externalities"),
                                   ("b", "Module 5: Monopoly and Monopolistic Competition")]}]},
        "due": [("Problem Set 2", 6, "Tue", None)],
    },
    {
        "num": 6, "kind": "midterm",
        "topics": ["Midterm Exam (covers through Module 3)"],
        "prep_days": ("Mon", "Wed"),
        "prep_groups": [
            {"cat": "podcast", "label": "Podcasts About Class Material:",
             # Modules 4 and 5 are taught in the week-5 class, so their
             # wrap-ups wait until after it. Module 3's went to week 3 with
             # its videos -- that class only does Module 3 APPLICATIONS
             # (2026-09-05, Nico).
             "items": podcast_items((5, "wrap"))},
            {"cat": "video", "label": "Practice on Module 4 (optional):",
             "items": [("v", "m4p1", "Practice Video: Optimization of a Price Taker", 19)]},
            {"cat": "other", "label": None,
             "items": [("t", "Midterm Prep: TA Review Sessions and Practice Sessions")]},
        ],
        "exam": {
            # 2026-08-31: header carries the name only; what used to sit in
            # the header ("covers through Module 3") is now bullet 1.
            "title": "Midterm Exam (online)",
            "window": (("Fri", 0), ("Sat", 0)),   # weekday, week offset from this week
            "lines": [
                "The midterm covers through Module 3.",
                "3-hour window at home – exact time window to be "
                "determined, will be announced in class.",
                "The midterm takes place online, with proctoring software.",
                "Open book, open notes. Calculator allowed.",
            ]},
        "due": [],
    },
    {
        "num": 7, "kind": "deadline",
        "topics": ["Module 6: Complex Pricing and Advanced Pricing Strategies"],
        "prep_days": ("Mon", "Fri"),
        "prep_groups": [
            {"cat": "podcast", "label": "Podcasts About Class Material:",
             "items": podcast_items((6, "intro"), (6, "wrap"))},
            {"cat": "video", "label": "Module 6: Complex Pricing and Advanced Pricing "
                      "Strategies – videos to watch by the weekend:",
             "items": [("v", "m6v1", "Video 1: Simple vs. Complex Pricing", 13),
                       ("v", "m6v2", "Video 2: First-Degree Price Discrimination", 11),
                       ("v", "m6v3", "Video 3: Segment Pricing", 14),
                       ("v", "m6v4", "Video 4: Versioning and Coupons", 13),
                       ("v", "m6v5", "Video 5: Flat Fee Pricing", 16),
                       ("v", "m6v6", "Video 6: Two-Part Tariffs", 8),
                       ("v", "m6v7", "Video 7: Block Pricing", 8),
                       ("v", "m6v8", "Video 8: Summary of Pricing Strategies", 6)]},
            {"cat": "read", "label": "In preparation for the Module 6 videos:",
             "items": [("t", "Ch. 10.1 – 10.4"),
                       ("t", "Ch. 10.6")]},
            {"cat": "practice", "label": None,
             "items": [("l", "prac_m4b",
                        "Online quiz on Module 4 (Part II): Market Distortions / Externalities"),
                       ("l", "prac_m5",
                        "Online quiz on Module 5: Monopoly and Monopolistic Competition")]},
        ],
        "due": [("Problem Set 3", 8, "Tue", None)],
    },
    {
        "num": 8, "kind": "deadline",
        "topics": ["Module 7 (Part I): Oligopoly with Homogenous Goods"],
        "prep_days": ("Mon", "Fri"),
        "prep_groups": [
            {"cat": "podcast", "label": "Podcasts About Class Material:",
             "items": podcast_items((7, "intro"), (7, "wrap"))},
            {"cat": "video", "label": "Practice on Module 6 (required):",
             "items": [("v", "m6p1", "Practice Video: Optimal Pricing in two Markets", 19)]},
            {"cat": "video", "label": "Module 7 (Part I): Oligopoly with Homogenous "
                      "Goods – videos to watch by the weekend:",
             "items": [("v", "m7v1", "Video 1: Oligopoly – Introduction", 10),
                       ("v", "m7v2", "Video 2: Collusion and Cartels", 9),
                       ("v", "m7v3", "Video 3: Cournot Competition", 22),
                       ("v", "m7v4", "Video 4: Bertrand Competition", 12)]},
            {"cat": "read", "label": "In preparation for the Module 7 videos:",
             "items": [("t", "Ch. 11.1 – 11.4")]},
            {"cat": "practice", "label": None,
             "items": [("l", "prac_m6",
                        "Online quiz on Module 6: Complex Pricing and Advanced Pricing Strategies")]},
        ],
        "due": [],
    },
    {
        "num": 9, "kind": "oncampus",
        "topics": ["Module 6: Applications",
                   "Module 7 (Part II): Oligopoly with Diff. Goods; Game Theory",
                   "Module 8: Asymmetric Information; Auctions"],
        "prep_days": ("Mon", "Fri"),
        "prep_groups": [
            {"cat": "podcast", "label": "Podcasts About Class Material:",
             "items": podcast_items((8, "intro"))},
            # The optional listening for every module taught in THIS
            # week's class (2026-09-14, Nico) -- Modules 6, 7 and 8. Each
            # title names its module, so the module pages sort them out.
            {"cat": "podcast", "label": "Optional Podcasts (before class):",
             "items": [("l", "pod_groupon",
                        "Groupon, Monty Python and Price Discrimination "
                        "(Module 6)"),
                       ("l", "pod_pd",
                        "The Prisoner\u2019s Dilemma and How to Solve It "
                        "(Module 7)"),
                       ("l", "pod_penalty",
                        "Applying Game Theory to Soccer Penalty Kicks "
                        "(Module 7)"),
                       ("l", "pod_tlae",
                        "Economics For All Your Decisions In Life "
                        "(Module 8)"),
                       ("l", "pod_13000",
                        "13,000 Economists. 1 Question (Module 8)")]},
            {"cat": "read", "label": "In preparation for class:",
             "items": [("t", "For Module 7: Ch. 11.6"),
                       ("t", "For Module 7: Ch. 12.1 – 12.2 (only until p. 373)"),
                       ("t", "For Module 8: Ch. 16.1 – 16.5")]},
            {"cat": "read", "label": None,
             "items": [("t", "Assigned articles for discussion (posted on BruinLearn)")]},
            {"cat": "practice", "label": None,
             "items": [("l", "prac_m7a",
                        "Online quiz on Module 7 (Part I): Oligopoly with Homogenous Goods")]},
        ],
        "weekend": {"days": ("Fri", "Sat"),
                    "groups": [
                        {"label": None,
                         # Was "Discussion: Application of Economic Concepts
                         # (articles will be assigned)" until 2026-09-14
                         # (Nico). The class card now mirrors the week's own
                         # Topics line, the way week 5's already does -- and
                         # naming the module is what puts this line on the
                         # Module 6 page.
                         "items": [("b", "Module 6: Applications"),
                                   ("b", "Module 7 (Part II): Oligopoly with Diff. Goods; "
                                         "Game Theory"),
                                   ("b", "Module 8: Asymmetric Information; Auctions")]}]},
        "due": [("Problem Set 4", 10, "Tue", None)],
    },
    {
        "num": 10, "kind": "thanksgiving",
        "topics": ["Practice videos on Module 7",
                   "Thanksgiving – no further videos"],
        "prep_days": ("Mon", "Wed"),
        "prep_groups": [
            {"cat": "podcast", "label": "Podcasts About Class Material:",
             # Modules 7 and 8 are taught in the week-9 class. Module 6's
             # wrap-up went to week 7 with its videos -- that class only
             # does Module 6 APPLICATIONS.
             "items": podcast_items((8, "wrap"))},
            {"cat": "video", "label": "Practice on Module 7 (required):",
             "items": [("v", "m7p1", "Practice Video: Cournot Competition – Math", 11),
                       ("v", "m7p2", "Practice Video: Oligopoly with different MC", 18)]},
            {"cat": "video", "label": "Practice on Module 7 (optional, advanced) [for "
                      "those interested in the math – not required for the exam]:",
             "items": [("v", "m7adv", "Practice Video: Oligopoly with Differentiated "
                                      "Goods – Math", 11)]},
            {"cat": "practice", "label": None,
             "items": [("l", "prac_m7b",
                        "Online quiz on Module 7 (Part II): Oligopoly with Diff. Goods; Game Theory"),
                       ("l", "prac_m8",
                        "Online quiz on Module 8: Asymmetric Information; Auctions")]},
        ],
        "holiday": {"window": (("Thu", 0), ("Sun", 0)),
                    "text": "Thanksgiving – no further videos"},
        "due": [("Problem Set 5", 11, "Thu", None)],
    },
    {
        "num": 11, "kind": "examprep",
        "topics": ["Exam preparation (1.5 weeks)"],
        "span_override": ((11, "Mon"), (12, "Thu")),
        "prep_days": None,
        "prep_groups": [
            {"cat": "other", "label": None,
             "items": [("t", "Exam prep time: if you follow the schedule, you will have "
                             "1.5 weeks to prepare. More videos will be posted with "
                             "review material by the TA."),
                       ("t", "Final Exam Prep: Solve the Practice Final"),
                       ("t", "TA Review Sessions and Practice Sessions")]},
        ],
        "due": [("Practice Final Exam", None, None, "solutions on BruinLearn")],
    },
    {
        "num": 12, "kind": "final",
        "topics": ["Final Exam (online, covers all material)"],
        "span_override": ((12, "Fri"), (12, "Sun")),
        "prep_days": None,
        "prep_groups": [],
        "exam": {
            # 2026-08-31: header carries the name only; coverage moved to
            # bullet 1, so the old "The exam will cover all material."
            "title": "Final Exam (online)",
            # 2026-09-10: the final is no longer a two-day window a student
            # starts somewhere inside. It is a FIXED 3-hour slot, so the
            # window is the one Saturday and `slot` carries the hours. Every
            # renderer -- the band, the Due column, the week page and the
            # .ics feed -- reads the time from here.
            "window": (("Sat", 0), ("Sat", 0)),
            "slot": ((9, 0), (12, 0)),
            "lines": [
                "The final exam covers all material, Modules 1 – 8.",
                "The final exam will take place online, and you will have 3 hours "
                "to solve the exam and upload your scanned solutions.",
                "The exam takes place on {w0}, {slot}.",
                "There will be about 20 multiple choice questions and "
                "3 – 4 problem-solving questions.",
                "Open book, open notes. Calculator allowed. See syllabus for further detail.",
                "We will use proctoring software.",
            ]},
        "due": [],
    },
]


# ============================== DERIVED VIEWS ==============================

_MOD_RE = re.compile(r"Module\s+(\d)")


def modules_named(text):
    """The module numbers a piece of calendar text names, in order."""
    out = []
    for m in _MOD_RE.finditer(text or ""):
        n = int(m.group(1))
        if 1 <= n <= 8 and n not in out:
            out.append(n)
    return out


def inclass_modules(wk):
    """Modules an on-campus class covers, for the "In-Class Material" list.

    Read from the week's TOPICS as well as the class card's own items: week
    9's card names Modules 7 and 8, but the class also covers Module 6 --
    its applications sit behind the Discussion item -- and the topics list
    catches that (2026-09-04, Nico). Both the calendar and the website
    build their list from this one function.
    """
    mods = []
    for t in wk.get("topics") or []:
        for n in modules_named(t):
            if n not in mods:
                mods.append(n)
    for g in (wk.get("weekend") or {}).get("groups") or []:
        for it in g["items"]:
            txt = it[1] if it[0] in ("t", "b", "note") else (
                it[2] if it[0] in ("v", "l", "p") else "")
            for n in modules_named(txt):
                if n not in mods:
                    mods.append(n)
    return sorted(mods)
