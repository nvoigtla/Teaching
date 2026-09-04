# MGMT 405 Course Website

Static course website for MGMT 405 – Managerial Economics (EMBA Hybrid,
Fall 2026). One page per week, one page per module, plus a General Logistics
page. No server and no build step at view time: plain HTML, CSS and JavaScript.

- **Live site:** https://nvoigtla.github.io/MGMT-405-EMBA/
- **Public repo (built output only):** https://github.com/nvoigtla/MGMT-405-EMBA

## The course calendar is the single source of truth

Nothing on the website is typed twice. Every date, topic, video name, Panopto
link, podcast, reading, practice quiz, problem-set deadline and exam window
comes from `../Course Calendar/_calendar_content.py` — the same file that
builds the Word/PDF calendar.

```
../Course Calendar/_calendar_content.py     content (dates, links, topics)
            |
            |  python _build_site.py
            v
   index.html, week-01..week-12.html, module-1..module-8.html
   assets/search-index.js
            |
            |  python _deploy.py
            v
   github.com/nvoigtla/MGMT-405-EMBA  ->  GitHub Pages
```

**After any calendar edit:**

```
python _build_site.py      # regenerate all 21 pages + the search index
python _deploy.py          # push the built pages to the public repo
```

`_deploy.py --dry-run` lists what would be pushed without changing anything.

## Layout of a page

- **Navy top bar**, laid out in three tracks so the middle is centered on the
  page: a **hamburger jump menu** and the site name on the left, the
  instructor's name (linked to his faculty page, new tab) with the school in
  the center, and the **View by** button on the right. The jump menu is a
  native `<select>` at `opacity:0` laid over three CSS-drawn bars — it looks
  like a hamburger button but still opens the platform's own picker, which on
  a phone is a full-screen list.
- **Left column** — the color-coding key (the agenda legend from calendar
  page 1), a *By Week* / *By Module* toggle, and the page list. The list always
  starts with **General Logistics**, then reads "Week 1", "Week 2", … or
  "Module 1", "Module 2", … with each row's dates and coverage underneath, and
  closes with **All Videos**, **All Podcasts** and **All Practice Exercises**
  (the last an outgoing link to the TA's site, marked with a ↗ and opening in
  a new tab). The three carry the same glyphs the week headers use — 🎬, 🎧, ✎. The current page is
  highlighted; there is no breadcrumb anywhere. The on-campus badges (weeks 1,
  5 and 9) carry the same navy rule over a transparent navy wash as the color
  key, rather than a solid navy block; **weeks 10 and 11 are colored as video
  content**, since neither has an in-person component. The jump menu follows
  the same order.
- **Right column** — the search box at the top, exactly as wide as the
  deadlines card beneath it, so the week header sits level with it across the
  middle of the page. Search covers every week and module page: `/` focuses
  it, arrows and Enter navigate the results.
- **The band** (the navy header of every page) carries three things: the week
  or module identity on the **left**, what kind of week it is in the
  **center**, and the page-turn link back to the previous page on the
  **right**. Week 1 and Module 1 point back to General Logistics.
- **Next page** — bottom right of the content. The last week and the last
  module have none, and **week 1 has no back link** either: it has no previous
  week, and both menus already offer General Logistics.

The band's center names what the week is, with no dates:

| Week | Center |
|---|---|
| Videos to watch *and* an on-campus class | Video Content & On-Campus Class |
| Videos only | Video Content |
| On-campus class only (weeks 5 and 9 assign no new videos) | On-Campus Class |
| Midterm / final | Midterm Exam / Final Exam |
| Exam-prep week | Exam Preparation |

## Category cards match the calendar

A week's preparation material sits inside a thick-navy container, headed
"Before Class" on an on-campus week and "During the Week" otherwise. Inside
it, one card per category, in the calendar's order, each with a navy header
bar carrying the calendar's glyph and the group lead-in ("Watch before
class:") below it:

| Category | Header glyph | Card body |
|---|---|---|
| Videos | clapperboard 🎬 | pale gold `#F6E8C9` |
| Podcasts | headphones 🎧 | light gray `#EFF1F4` |
| Suggested Reading | open book 📖 | white |
| Suggested Additional Practice Exercises | pencil ✎ | white |

The **On-Campus Class** card carries a classical building 🏛️ — U+1F3DB
followed by **U+FE0F, the emoji presentation selector**. The selector matters:
a bare U+1F3DB renders as a flat monochrome text glyph on some platforms, and
only with FE0F does it come out as the color emoji. The same character is in
the calendar's on-campus header. It ships as text — no image asset — and in
the calendar it sits at the right tab stop like the other glyphs there; the
header text measures 5.14" against a 6.07" tab stop, so it has room and cannot
wrap the line.

Individual video bullets keep the play triangle ▶. A video's running time sits
immediately after its link, as in the calendar — `(9 min)`, or `(++)` where the
running time has not been measured yet, or `(link to follow)` where the video
has not been posted. **Per-week video totals are not shown.**

**On an on-campus week the class card comes first**, above the preparation
container. The same changes (clapperboard glyph, gray podcast card, class card on
top) were made in `_build_calendar.py` on 2026-09-03, so the two documents stay
in step.

**Every boxed heading uses the navy header bar and Title Case**, the same rule
the slide decks use for slide titles: significant words capitalized, articles
and short prepositions left lower unless they open or close the heading, both
halves of a hyphenated compound capitalized, and never a capital lowered.
`title_case()` in `_build_site.py` does this; it applies only to headings the
site itself authors, never to calendar content.

## Where a visitor lands

The site root lands on **the current week** once the quarter is running, and
on **General Logistics** before it starts and after it ends. The redirect is a
short inline script in `index.html` (generated, so the week dates come from the
calendar). The sidebar's General Logistics link carries `?stay=1`, which
suppresses the redirect, so the page is always reachable.

## Light mode only

The site is light-mode only (2026-09-03). `_build_site.py` stamps
`data-theme="light"` on `<html>` so nothing can flip it, `assets/site.css`
carries no dark palette, and the theme code — a sunrise/sunset calculation and
a time-zone table — is gone from `assets/site.js`. If a dark palette is ever
wanted again, it goes back as a `@media (prefers-color-scheme:dark)` block
redefining only the tokens.

## One view mode, three controls

"Weeks" or "modules" is a single piece of state in `assets/site.js`, and three
controls follow it:

- the **View by** button in the top bar (two lines, so it stays narrow on a
  phone: "View by" over "Week" / "Module");
- the **By Week / By Module** toggle in the left menu;
- **which entries the hamburger jump menu offers** — only weeks, or only
  modules, plus General Logistics either way.

The default is weeks. A module page opens in module view, because that is what
the reader is looking at; otherwise the last choice is remembered in that
browser. The jump menu's options are re-emitted from a snapshot on each
change, rather than hidden with the `hidden` attribute, which is not honored
on every platform.

**The View by button also returns to General Logistics**, so the new view is
applied somewhere it is visible. The left menu's toggle switches in place.

## All Videos and All Podcasts

Two index pages sit outside the week/module sequence, listed in both menus and
in the jump menu: **All Videos** (`all-videos.html`) and **All Podcasts**
(`all-podcasts.html`). Each holds one card per module, in module order, keeping
every source group's own lead-in from the calendar and naming the week it came
from. Practice videos and recaps are included — these are complete indexes, not
just the required viewing.

A module with nothing to list says so rather than being absent: modules 5 and 8
have no videos, so their cards read "This module is covered in class, in week
5 / 9". Podcasts that belong to no single module (the general-interest
episodes) get a final card of their own. Both pages are declared once in
`EXTRA_PAGES` in `_build_site.py`, which drives the pages, both menus and the
jump menu together.

## Titles, in-class material, podcast labels

- **Every browser tab reads "Econ 405"** (`TAB_TITLE`), whichever page is
  open. The page's own name lives in the band, not the tab.
- **On-campus cards carry an "In-Class Material" section**: a handout and a
  slide deck per module that class covers, both marked "(TBD)" until they are
  uploaded. The module list comes from the class card's OWN items, so the
  section says exactly what the card just listed.
- **Module podcasts read "Podcast (_before_ class):" and "Podcast (_after_
  class):"**, with the timing word underlined. Done in `podcast_label()` at
  render time rather than in the calendar content, because the calendar's item
  format carries no per-run formatting for podcasts — so the calendar still
  prints the plain "Podcast: Intro to Module 1".

## Problem sets

A problem set's card is set apart from every other card: a **dark-red rule
(`#C00000`) over a very transparent dark-red wash**, with the due date at
17 px semibold. It also carries **"Upload one solution per group on
BruinLearn"**, linking the course site (`BRUINLEARN_COURSE`).

All three are gated on the label starting with "Problem Set", which is why
week 11's Practice Final Exam keeps the gold rule and the smaller date — it is
not a problem set. The Deadlines & Exams column stays terse (label and date
only), because it is a narrow sidebar.

## Each column scrolls on its own

Above 861px the viewport is the frame: the top bar and footer are fixed bands,
`.shell` takes the space between them, and the three columns each get their own
`overflow-y:auto`. The page itself does not scroll.

Two traps are worth remembering here. `.shell` carries `align-items:start`,
which has to become `stretch` so the columns fill the height. And the side
columns are flex containers whose children **shrink by default** — which meant
the menu card was squeezed and its rows silently *clipped* by its own
`overflow:hidden` rather than scrolling, so weeks 8–12 disappeared on a short
window. `.left > *, .right > * { flex:none }` fixes it; verify by comparing
`scrollHeight` to `clientHeight` on each column, not by eye.

## Phones

The layout is responsive, with three breakpoints:

| Width | Layout |
|---|---|
| above 1240px | three columns: navigator, content, deadlines |
| 860 – 1240px | navigator and content side by side, deadlines below |
| below 860px | one column, and **the content comes first**. The whole left menu — the color key and the By Week / By Module choices — is hidden; the top-bar jump menu is the navigator on a phone. Search and deadlines follow the content, the **search box moves to the very top**, under the title bar and ahead of the week's content — done by dissolving the right column with `display:contents` so its two children can be ordered independently — and the week header grows a **"Deadlines for this week ↓"** link (hidden on wider screens). It jumps to that week's first row — `#dl-here` — **and narrows the list to that week alone**; a "Show all deadlines" button then restores the rest. The click is not intercepted, so the anchor jump still happens naturally. |
| below 700px | the top-bar byline is dropped (the footer repeats it) and the jump menu takes the corner |

Verified with no horizontal overflow down to 512 CSS pixels, which is as
narrow as headless Chrome will size a window on Windows; the sub-420px rules
(smaller body text, tighter padding) are therefore unverified by measurement.
Three traps worth remembering: `.shell` carries `align-items:start`, which
stops flex children from filling the width — the phone layout has to reset it
to `stretch`; `.box-hd` needs `flex-wrap:wrap` so a long header line (the
class times and room) can wrap instead of overflowing; and **anything sitting
on the navy bar needs its own link color** — the faculty-page link inherited
the navy body link color and rendered invisible until `.tb-sub a` was given
one.

## Deadlines & Exams

The right column lists every problem set and exam in date order, with the
current week highlighted, and the videos that have to be watched.

Every week that requires videos gets a row of its own, e.g. *"Watch Videos
1 – 7 for Module 3, by 10/11"*. A group marked "(optional)" and the recaps are
left out; a group marked "(required)" is counted even when its items are
practice videos, which is why weeks 8 and 10 appear. The deadline is the
Friday class on an on-campus week and the week's Sunday otherwise — the same
"suggested" date the calendar prints.

**Anything due within the next three days gets the problem-set treatment** —
dark-red rule over a transparent dark-red wash, with the date in red — and a
legend above the list explains the coding. This is the one thing on the site
that cannot be baked in at build time, because "within three days" moves every
day: each row carries a `data-date`, and `initDue()` in `assets/site.js`
compares it to today in the browser. Undated rows (the Practice Final Exam)
are skipped.

## How the module view is derived

The calendar is organized by week, so module membership is inferred. A group of
items is tagged with every module its label or its items name ("Module 3:
Production & Costs – videos to watch by the weekend" → module 3). Groups whose
text names no module are listed explicitly in `GROUP_MODULE_OVERRIDES` in
`_build_site.py`, each with a comment giving the reason; genuinely non-module
material (general-interest podcasts, discussion articles, exam prep) is tagged
with no module and appears on the week page only.

One judgment call is recorded there: week 1's "In preparation for class:
Ch. 2.5" names no module, so it is tagged with **both** modules that week 1's
on-campus class covers (1 and 2) rather than guessing one.

## The General Logistics page

A row of two half-width boxes, then two fixed columns of three. The order is
fully deterministic — the page is assembled from explicit column stacks, not an
auto-flowing grid, so nothing moves as the window changes size:

| | Left | Right |
|---|---|---|
| Row 1 | Class Syllabus ▤ | BruinLearn Class Site 🎓 |
| | **How the Quarter Runs** | **Class and Contact** |
| | Watching the Videos 🎬 | Math Refresher ∂ |
| | Online Practice Exercises ✎ | Textbook 📖 |

Class and Contact opens the right-hand column, so it sits in the top right
corner; within each column the pairs stay together — practice under videos,
textbook under the math refresher.

"Watching the Videos" links the **Panopto sign-in screenshot**, which
`_build_site.py` copies out of the calendar's `Images/` folder into
`assets/panopto-login.png` on every build; `_deploy.py` ships it with the site.

## One placeholder

- `SYLLABUS_URL` at the top of `_build_site.py` is `"#"`. Paste the Bruin Learn
  or PDF address there and rebuild; the General Logistics page then links it as
  "Download the Class Syllabus here".

`DROP_TEXTBOOK_NOTES` and `DROP_MATH_ITEMS` in `_build_site.py` can hide a
calendar line on the website while it still prints in the calendar. **Both are
empty** — every line Nico wanted gone was deleted from `_calendar_content.py`
outright, so the two documents now carry exactly the same content.

## Palette and type

The palette is the decks' palette, unchanged: navy `#0B2B4E`, gold `#E09F3E`,
pale gold `#F6E8C9`, cream `#FDF6E6`, gray `#555B66`, pale blue `#E7EDF4`, and
the podcast gray `#EFF1F4`. Week color coding matches the calendar — in
class = navy, video content = pale gold, exam = gold.

Body type is **Carlito**, the metric-compatible Calibri clone the decks already
use for charts, so the website reads in the same face as the slides; it falls
back to Calibri where Carlito is unavailable.

## Files

| File | |
|---|---|
| `_build_site.py` | the generator; run after every calendar change |
| `_deploy.py` | pushes the built pages to the public repo and enables Pages |
| `assets/site.css` | hand-authored stylesheet (not generated) |
| `assets/site.js` | hand-authored theme + search + nav toggle (not generated) |
| `assets/search-index.js` | GENERATED — do not edit |
| `assets/panopto-login.png` | GENERATED — copied from the calendar's `Images/` |
| `all-videos.html`, `all-podcasts.html` | GENERATED — do not edit |
| `index.html`, `week-*.html`, `module-*.html` | GENERATED — do not edit |

`_deploy.py` **discovers** the pages to publish (every `.html` in the folder)
rather than carrying a list — a hardcoded list silently dropped the two new
index pages on their first deploy, so they 404'd live while the local build was
fine. Only the built pages and the hand-authored assets are copied to the public
repository. The build script, this README and `Session-Notes.md` stay private.

## A note on what is public

The pages carry `noindex, nofollow`, so search engines will not list them. The
repository is public, though, which means the **Dropbox podcast links in the
page source are readable by anyone who finds the repository** — Dropbox share
links need no login. Panopto links require an Anderson sign-in, and the TA's
practice site is already public. If the podcasts should not be reachable that
way, the options are to move them behind Bruin Learn or to host the site
privately instead.
