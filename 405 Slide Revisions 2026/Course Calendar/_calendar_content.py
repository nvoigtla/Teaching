# -*- coding: utf-8 -*-
"""
Content source for the MGMT 405 EMBA Hybrid course calendar.

ALL dates derive from ONE anchor: ANCHOR_FRIDAY, the Friday of the first
on-campus class weekend. To roll the calendar to a new year, change that
single line (and any content edits), then re-run _build_calendar.py.

Week 1 = the calendar week (Mon-Sun) containing the first on-campus Friday.
Every date in the document is expressed as (week_number, weekday).
"""

from datetime import date, timedelta

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
    """'Sep 21 – 27' or 'Sep 28 – Oct 4' (month repeated only if it changes)."""
    if d1.month == d2.month:
        return f"{d1.strftime('%b')} {d1.day} – {d2.day}"
    return f"{fmt(d1)} – {fmt(d2)}"


# ============================== HYPERLINKS ==============================

LINKS = {
    "ta_email":   "mailto:TA405.EMBA2@gmail.com",
    "math_quiz":  "https://bruinlearn.ucla.edu/courses/195707?invitation=0GYbOXWd6mBK4bwem1dgAP7Jq2ad1PZo5Wp4kTYQ",
    "math_videos": "https://bruinlearn.ucla.edu/courses/195707/pages/econ-math-review-videos",
    # asm_panopto (dropbox screenshot) retired 2026-08-15: the screenshot is now
    # embedded on p.2 (Images/Panopto-Login-Picture.png) and the link points to
    # the Panopto site itself (Nico's hand-edit).
    "panopto_site": "https://ucla-anderson.hosted.panopto.com",
    # Module 1 -- videos re-recorded and re-uploaded as NEW Panopto
    # sessions 2026-08-28 (ids b4b3...; the old b1d9... ids were last
    # year's recordings). Running times still unmeasured -> "(++)".
    "m1v1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=45ebea5c-ce47-4f69-9693-b4b30104a768",
    "m1v2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=ea13ec98-9412-41fc-accd-b4b30104a761",
    "m1v3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=7ad4fc39-0070-4aa6-90a9-b4b30104a761",
    "m1v4": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=8bd7e577-22b2-4da6-888d-b4b30104a76a",
    "recap1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=2f3c674d-13da-4308-ada2-b08b0134d2ff",
    # Podcasts
    "pod_cb":  "https://podcasts.apple.com/us/podcast/core-principle-1-the-cost-benefit-principle-the-pros-cons/id1523898793?i=1000488478204",
    "pod_oc":  "https://podcasts.apple.com/us/podcast/core-principle-2-the-opportunity-cost-principle-or-what/id1523898793?i=1000488478205",
    "pod_freak": "https://freakonomics.com/podcast/should-we-really-behave-like-economists-say-we-do-a-new-freakonomics-radio-podcast/",
    "pod_tlae": "https://art19.com/shows/think-like-an-economist/episodes/d48893ba-1f44-43d6-a271-be306d55d0f9?fbclid=IwAR14eZxCKFr8FOC5jQ3ZvGia0WQZRs-zytfYxEUjfWsPlj8aKRzIF8zpmYQ",
    # Module 2
    "m2v1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=acc0db41-0b2e-4b30-9c2a-b08b012eeeb1",
    "m2v2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=647e34af-e623-45af-882b-b08b012f00f2",
    "m2v3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=717874c9-f69f-4aa2-a66f-b08b012f0df4",
    "m2p1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=d9a1abf9-e8e5-448b-bbbd-b08b012eeeb0",
    "m2p2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=cf0b4650-7a8b-41a0-9a6e-b08b012eeeb9",
    "recap2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=4979325b-0d4a-45d8-9401-b08b0134db2a",
    # Module 3
    "m3v1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=59ea6ec3-dbb0-4938-9196-b08b012f829f",
    "m3v2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=28855216-e910-4fd1-9c52-b08b012ff530",
    "m3v3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=e70c9db1-92a1-424c-b6ca-b08b0130266f",
    "m3v4": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=fc0411b6-74ba-429c-8795-b08b01302eeb",
    "m3v5": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=2ce95a04-3ea0-4d09-860b-b08b01307162",
    "m3v6": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=05d552d7-d41f-45d7-a4b8-b08b0130fa7b",
    "recap3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=7803e078-6264-46e3-bdd6-b08b0134ff01",
    "m3pa": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=147753a0-3a02-40e0-a75c-b08b012f36b3",
    "m3pb": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=047b0460-fd8b-4a4e-8996-b08b012f5514",
    # Module 4
    "m4v1": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=5941b0ae-0671-465f-8547-b08b0131e697",
    "m4v2": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=ff155d66-d7f6-42cb-8243-b08b0131e6f7",
    "m4v3": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=3b55b68c-322f-4f07-963d-b08b0131e6bf",
    "m4v4": "https://ucla-anderson.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=5e595b0c-ac1d-4184-82de-b08b0131ed9b",
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
}

# ============================== HEADER / INTRO ==============================

COURSE_TITLE = "MGMT 405 – Managerial Economics"
SUBTITLE = f"Course Calendar – {TERM} – EMBA Section 2 (Hybrid)"
BRUINLEARN_NOTE = ("Please check Bruin Learn under “Syllabus” "
                   "for the latest version of this calendar.")
SYLLABUS_NOTE = ("Please check Bruin Learn under “Syllabus” "
                 "for the more detailed Class Syllabus.")
TA_NAME = "Rafael Rubiao"
CLASSROOM = "A301"
CLASS_TIMES = "Fridays 4:00 – 5:30 pm  ·  Saturdays 9:00 am – 12:30 pm"

TEXTBOOK_NOTES = [
    "“Microeconomics” by Goolsbee, Levitt, Syverson (4th edition). "
    "See the different purchase options in the class syllabus.",
    "The “Figure it out” boxes in each chapter should be considered "
    "as advanced (voluntary) readings.",
    "For additional and optional practice, you can find review questions at "
    "the end of each chapter (with answers at the end of the textbook).",
]

MATH_REFRESHER_INTRO = ("If you feel you could use a math refresher before "
                        "the class, I suggest:")
MATH_REFRESHER_ITEMS = [
    [("t", "Take the "), ("l", "math_quiz", "Math Quiz"), ("t", " (10 – 15 min)")],
    [("t", "Watch the "), ("l", "math_videos", "Math Review Videos")],
    [("t", "Textbook reading: Math review Appendix Section 1 + Section 2 "
           "(only first derivatives)")],
    [("t", "Achieve: the first section of the Achieve course provides math "
           "and graphing video tutorials")],
]

SIGNIN_NOTE = [("t", "Sign-in to watch videos: Use option "),
               ("l", "panopto_site", "“ASM Panopto”")]

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
    4: {"intro": (None, None), "wrap": (None, None)},
    5: {"intro": (None, None), "wrap": (None, None)},
    6: {"intro": (None, None), "wrap": (None, None)},
    7: {"intro": (None, None), "wrap": (None, None)},
    8: {"intro": (None, None), "wrap": (None, None)},
}


def podcast_items(mod):
    """The two podcast bullets for module `mod`, as
    ("p", url, text, minutes) items. url and minutes are None until the
    episode has been uploaded."""
    ep = PODCASTS.get(mod, {})
    out = []
    for key, text in (("intro", "Podcast: Intro to Module %d" % mod),
                      ("wrap", "Podcast: Wrap-Up of Module %d" % mod)):
        url, mins = ep.get(key, (None, None))
        out.append(("p", url, text, mins))
    return out


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
             "items": [("v", "m1v1", "Module 1 – Video 1: Introduction", 9),
                       ("v", "m1v2", "Module 1 – Video 2: Markets", 10),
                       ("v", "m1v3", "Module 1 – Video 3: Demand and Supply", 8),
                       ("v", "m1v4", "Module 1 – Video 4: Market Equilibrium", 7)]},
            {"cat": "podcast", "label": "About Class Material:",
             "items": podcast_items(1)},
            {"cat": "podcast", "label": "Other podcasts:",
             "items": [("l", "pod_cb", "The Cost-Benefit Principle"),
                       ("l", "pod_oc", "The Opportunity-Cost Principle"),
                       ("l", "pod_freak", "Should We Really Behave Like Economists?")]},
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
            {"cat": "podcast", "label": "About Class Material:",
             "items": podcast_items(2)},
            {"cat": "video", "label": "Recap (optional):",
             "items": [("v", "recap1", "Recap of Module 1", 7)]},
            {"cat": "video", "label": "Remaining videos for Module 2 [material not "
                      "covered in the on-campus class] – watch by the weekend:",
             # names verbatim from Module 2/Videos Final/ (and the
             # _video_title_slide calls in _build_Module2Video.py)
             "items": [("v", "m2v1", "Module 2 – Video 1: Elasticity and Revenue", None),
                       ("v", "m2v2", "Module 2 – Video 2: Marginal Revenue", None),
                       ("v", "m2v3", "Module 2 – Video 3: Demand Estimation", None)]},
            {"cat": "video", "label": "Practice videos for Module 2:",
             "items": [("v", "m2p1", "Module 2 – Practice Video 1: Elasticity and Revenues", 7),
                       ("v", "m2p2", "Module 2 – Practice Video 2: Revenue Maximization", 6)]},
            {"cat": "read", "label": None,
             "items": [("note", "[Relevant textbook reading was already covered "
                                "in preparation for the previous class.]")]},
            {"cat": "read", "label": "Advanced reading (optional):",
             "items": [("t", "Chapters 5.1, 5.2, 5.4, 5.5")]},
            {"cat": "read", "label": "Teaching notes (optional, posted on Bruin Learn):",
             "items": [("t", "Teaching note: Marginal Revenue"),
                       ("t", "Teaching note: Demand Elasticity and Total Revenue"),
                       ("t", "Teaching note: Regressions")]},
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
            {"cat": "podcast", "label": "About Class Material:",
             "items": podcast_items(3)},
            {"cat": "video", "label": "Recap (optional):",
             "items": [("v", "recap2", "Recap of Module 2", 8)]},
            {"cat": "read", "label": "In preparation for the Module 3 videos:",
             "items": [("t", "Ch. 6.1 – 6.3, 6.5"),
                       ("t", "Ch. 7")]},
            {"cat": "video", "label": "Module 3: Production & Costs – videos to "
                      "watch by the weekend:",
             # 2026-08-28: Module 3 was retaped as SEVEN videos. Names
             # verbatim from M3_OUTLINE in Module 3/_m3_outline.py, the
             # single source of the deck's video title cards. The new
             # "Introduction to Module 3" joins at the front and has no
             # Panopto link yet, so the six old link keys shift down one
             # topic (m3v1 = the production-function video, and so on).
             "items": [("v", None, "Video 1: Introduction to Module 3", None),
                       ("v", "m3v1", "Video 2: The Production Function", None),
                       ("v", "m3v2", "Video 3: Short Run: Hiring Decisions", None),
                       ("v", "m3v3", "Video 4: Wage Searchers", None),
                       ("v", "m3v4", "Video 5: Long Run: The Optimal Input Mix", None),
                       ("v", "m3v5", "Video 6: Cost Concepts", None),
                       ("v", "m3v6", "Video 7: Economies of Scale and Scope", None)]},
            {"cat": "read", "label": "Advanced reading (optional):",
             "items": [("t", "Ch. 6.6 and 6.7")]},
        ],
        "due": [("Problem Set 1", 4, "Tue", None)],
    },
    {
        "num": 4, "kind": "deadline",
        "topics": ["Module 4 (Part I): Competitive Markets and Market Interventions"],
        "prep_days": ("Mon", "Fri"),
        "prep_groups": [
            {"cat": "podcast", "label": "About Class Material:",
             "items": podcast_items(4)},
            {"cat": "video", "label": "Recap (optional):",
             "items": [("v", "recap3", "Recap of Module 3", 7)]},
            {"cat": "video", "label": "Practice on Module 3 (optional):",
             "items": [("v", "m3pa", "Practice Video: Costs: Make vs Buy Decision", 10),
                       ("v", "m3pb", "Practice Video: Short-Run and Long-Run Costs", 26)]},
            {"cat": "video", "label": "Module 4 (Part I): Competitive Markets and "
                      "Market Interventions – videos to watch by the weekend:",
             # 2026-08-30: Module 4 was re-split into FIVE videos when the
             # deck was converted for taping - Perfect Competition is now
             # a video of its own, and the module front matter sits inside
             # Video 1.  Every video therefore has to be re-recorded, so
             # the links and the running times are blank for now; a None
             # length prints "(++)" and suppresses the total line.  The
             # old Panopto URLs are still in VIDEO_LINKS (m4v1 - m4v4) and
             # can be re-pointed once the new cuts are up.
             "items": [("v", None, "Video 1: Introduction to Market Structures", None),
                       ("v", None, "Video 2: Perfect Competition", None),
                       ("v", None, "Video 3: Profit Maximization of a Price Taker – Short Run", None),
                       ("v", None, "Video 4: Firm-Level and Market Supply", None),
                       ("v", None, "Video 5: Long-Run Competitive Equilibrium", None)]},
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
            {"cat": "podcast", "label": "About Class Material:",
             "items": podcast_items(5)},
            {"cat": "read", "label": "In preparation for class:",
             "items": [("t", "For Module 4 (Part II): Ch. 3.1 – 3.4; Ch. 17 (pp. 513 – 524)"),
                       ("t", "For Module 5: Ch. 9.1 – 9.3; Ch. 9.5 – 9.7; Ch. 11.7")]},
            {"cat": "read", "label": None,
             "items": [("t", "Assigned articles for discussion (posted on Bruin Learn)")]},
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
            {"cat": "video", "label": "Practice on Module 4 (optional):",
             "items": [("v", "m4p1", "Practice Video: Optimization of a Price Taker", 18)]},
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
                "3.5-hour window at home – exact time window to be "
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
            {"cat": "podcast", "label": "About Class Material:",
             "items": podcast_items(6)},
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
            {"cat": "podcast", "label": "About Class Material:",
             "items": podcast_items(7)},
            {"cat": "video", "label": "Practice on Module 6 (required):",
             "items": [("v", "m6p1", "Practice Video: Optimal Pricing in two Markets", 19)]},
            {"cat": "video", "label": "Recap (optional):",
             "items": [("v", "recap6", "Recap of Module 6", 7)]},
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
            {"cat": "podcast", "label": "About Class Material:",
             "items": podcast_items(8)},
            {"cat": "podcast", "label": "Other podcasts:",
             "items": [("l", "pod_tlae", "Economics For All Your Decisions In Life")]},
            {"cat": "read", "label": "In preparation for class:",
             "items": [("t", "For Module 7: Ch. 11.6"),
                       ("t", "For Module 7: Ch. 12.1 – 12.2 (only until p. 373)"),
                       ("t", "For Module 8: Ch. 16.1 – 16.5")]},
            {"cat": "read", "label": None,
             "items": [("t", "Assigned articles for discussion (posted on Bruin Learn)")]},
            {"cat": "practice", "label": None,
             "items": [("l", "prac_m7a",
                        "Online quiz on Module 7 (Part I): Oligopoly with Homogenous Goods")]},
        ],
        "weekend": {"days": ("Fri", "Sat"),
                    "groups": [
                        {"label": None,
                         "items": [("b", "Discussion: Application of Economic Concepts "
                                         "(articles will be assigned)"),
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
            {"cat": "video", "label": "Practice on Module 7 (required):",
             "items": [("v", "m7p1", "Practice Video: Cournot Competition – Math", 11),
                       ("v", "m7p2", "Practice Video: Oligopoly with different MC", 18)]},
            {"cat": "video", "label": "Recap (optional):",
             "items": [("v", "recap7", "Recap of Module 7", 8)]},
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
        "due": [("Practice Final Exam", None, None, "solutions on Bruin Learn")],
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
            "window": (("Sat", 0), ("Sun", 0)),
            "lines": [
                "The final exam covers all material, Modules 1 – 8.",
                "The final exam will take place online, and you will have 3.5 hours "
                "to solve the exam and upload your scanned solutions.",
                "The exam window runs from {w0} to {w1} – exact time "
                "window to be determined, will be announced in class.",
                "There will be about 20 multiple choice questions and "
                "3 – 4 problem-solving questions.",
                "Open book, open notes. Calculator allowed. See syllabus for further detail.",
                "We will use proctoring software.",
            ]},
        "due": [],
    },
]
