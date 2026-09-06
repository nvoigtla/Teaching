# Session Notes – Course Calendar (MGMT 405 EMBA Hybrid, Fall 2026)

Deliverable: `Calendar EMBA Hybrid -- Fall 2026.docx` (+ `.pdf`), built by
`_build_calendar.py` from `_calendar_content.py`. Content file is the single
source of truth for dates, links, video/podcast lists; the build script owns
layout, palette and chrome.

**Where this stands (end of 2026-09-04):** the calendar is current and was
restyled on 2026-09-03 to match the new course website in
`../Course Website/` -- podcast cards light grey, a clapperboard for Videos,
the on-campus card above the prep container and carrying a classical-building
glyph, problem-set cards in dark red, weeks 10 and 11 coloured as video
content, and three lines deleted (end-of-chapter review questions, the
"Achieve" math item, the math-review appendix reading). Rebuild with
`python _build_calendar.py`, export the PDF through Word COM, then run
`_check_pagination.ps1` -- it PASSES at 14 pages.

**One divergence to remember:** the website prints "Podcast (before class):" /
"(after class):" with the timing word underlined, but the calendar still
prints the plain "Podcast: Intro to Module X". The calendar's
`("p", url, text, minutes)` item format carries no per-run formatting, so the
underline cannot travel with it.

---

## One placeholder: "(link to follow)", and "(++)" retired

2026-09-06, Nico: "some videos still have the (++). Please use '(link to
follow)' throughout."

"(++)" meant something different -- *the link is there, the running time is
unmeasured* -- so the rule is now:

| case | shows |
|---|---|
| no link | **(link to follow)** |
| link + known running time | (12 min) |
| link, running time unknown | nothing (the link works; the duration is not the student's problem) |

### Why this also removed six links

Every bullet still showing "(++)" was a **Module 3** video, and those six
keys (`m3v1` - `m3v6`) point at **LAST YEAR's Panopto sessions** -- the
`b08b...` ids. Module 2's were replaced with `b4bb...` on 2026-09-04;
Module 3's never were, and the open-items list has said so for days.

Relabelling alone would have printed "(link to follow)" next to a working,
clickable link to last year's recording. So the keys came off the items at
the same time: the seven Module 3 bullets now carry no video link, read
"(link to follow)", and keep their slide decks. The URLs stay in `LINKS`
for reference -- restore a key on its line as each new session id lands.

Verified live on both sites: 0 occurrences of "(++)", 20 of "(link to
follow)", and **0 reachable `b08b` links**. Both calendars still 14 pages;
their PDFs went 91 -> 85 hyperlinks, which is exactly the six removed.

## Video slide decks, and the PDF-export regression that came with them

**19 decks published** (Modules 1-4), linked from both websites and both
calendars. Module 3 Video 7 "Economies of Scale and Scope" has no deck yet
and shows "(TBD)".

The decks are **DISCOVERED, not listed**: `_scan_slides()` in
`_calendar_content.py` walks `405 Slide Revisions 2026/Module */Videos
Final/` and keys on `(module, video number)` parsed from the filename
convention `Module <M> - Video <N> - <whatever>.pptx`. Drop a new deck in
that folder, run `python _publish.py`, and it appears -- no code change.
`unparsed_slides()` reports anything that does not match the convention.

Keying on the video NUMBER rather than the Panopto link key matters: Module
4's five videos have no Panopto links at all yet, and their decks still find
their bullets.

- Website: the deck link sits on the VIDEO'S OWN BULLET -- "Video 2: The
  Production Function (++) · slides" -- the same treatment as the calendar,
  which Nico asked for after seeing both (2026-09-06). It reaches all three
  views: the week pages, the module pages and All Videos, 19 links each.
  A first version gave the decks a "Slides for Videos" rubric of their own
  in the materials box; that is gone.
- The materials box is now **"Slides from Class"** and appears on ON-CAMPUS
  weeks only (1, 5, 9). With the video decks moved onto their bullets it
  holds only what is handed out in class, so the inner "In-Class Material"
  sub-heading was dropped as well -- it just repeated the box title.
- Calendar: the link is appended INLINE to the video's own bullet
  ("... (9 min) · slides"). Week 3 lists seven videos and every week has to
  stay on one page, so a line each was not affordable.
- The decks are copied into the built site by `_build_site.py` and shipped
  by `_deploy.py`. `.gitignore` keeps them out of the private repo -- they
  are build OUTPUT here, 51 MB per section.

### The regression: printing to PDF FLATTENS the file

Do not go back to the "Microsoft Print to PDF" workaround. It returns
quickly, but the calendar came out with **0 hyperlinks instead of 90** --
every video, podcast and slide-deck link gone, because the calendar's links
sit inside drawn text boxes. The syllabus kept 14 (its links are in body
text), which is exactly why this was easy to miss: one document looked fine.
It was published that way before being caught.

**What actually fixes `ExportAsFixedFormat`:** set `$w.ActivePrinter` to a
local printer BEFORE exporting. Word wants printer metrics to lay a PDF out
and stalls when the default is unavailable. With that one line it returns in
seconds and produces 90 links.

`verify_pdfs()` in `_publish.py` now counts hyperlinks and fails a PDF that
has none, so a flattened file can never ship again. A page-count and
section-name check alone would not have caught this -- both were correct.

### Also worth knowing

The six **teaching-note PDFs** are already wired into `_deploy.py`'s DOCS and
linked from the module and week pages. They were in the working tree but not
in commit 0cfc6b8, so they are Nico's own edits; they work and were left
alone.


## "Now update the website" -> `python _publish.py`

One command, both sections. It exists because **committing to git does not
update the website** (git is the private source; `_deploy.py` publishes to
the public repos GitHub Pages serves) and because the real gap sits upstream
of both: `_deploy.py` copies whatever `.pdf` is on disk, so a changed `.docx`
whose PDF was not re-exported ships the OLD one silently. That is how both
sites served A301 PDFs for two days while every HTML page said G305.

```
python _publish.py            # check, fix what it can, publish both
python _publish.py --check    # report only, publish nothing
python _publish.py --section femba
```

Six steps, in order: rebuild both sites -> re-export any PDF older than its
`.docx` -> verify each PDF opens, has pages and names its own section ->
check every local link resolves -> **report what is still a placeholder** ->
publish. It refuses to publish a section whose checks failed.

**Step 5 is the one to read after an upload.** It lists the "(TBD)" handouts
and slide decks and the "(link to follow)" videos per week -- currently 28
items across 8 weeks. When a teaching note or a video deck is added, that
count should drop; if it does not, the file was not wired in.

Two bugs found while building it, both worth remembering:

- PowerShell flattens `@( @("a","b") )` into a 2-element array of STRINGS, so
  a single-document export unpacked to `$out = $null` and did nothing. The
  script now emits one explicit block per document instead of looping.
- That silent no-op then **reported success**. A false all-clear is worse
  than the problem the script exists to catch, so `export()` now re-checks
  each target's mtime afterwards and fails loudly.

## The PDF export: ExportAsFixedFormat, with ActivePrinter set first

**Superseded advice below.** An earlier version of this note recommended
printing the document to the "Microsoft Print to PDF" driver, because
`ExportAsFixedFormat` was hanging. That workaround FLATTENS the file: the
calendar came out with 0 hyperlinks instead of 90, losing every video,
podcast and slide-deck link, because those sit inside drawn text boxes. It
was published in that state before being caught.

**What actually works:** set `$w.ActivePrinter` to a local printer BEFORE
calling `ExportAsFixedFormat`. Word wants printer metrics to lay a PDF out
and stalls when the default is unavailable -- that was the hang all along.
With that one line it returns in seconds and keeps every link.
`python _publish.py` does this, and `verify_pdfs()` fails any PDF with zero
hyperlinks so a flattened file cannot ship again.

The original note is kept below for the mechanics of driving Word over COM.

### Original note (do NOT use PrintOut for the real export)

`ExportAsFixedFormat` (and `SaveAs2` to wdFormatPDF) HANGS on this machine --
Word opens, stays Responding, burns CPU and writes nothing, with no modal
dialog visible in an `EnumWindows` dump. It failed on ~8 attempts across
three sessions, read-only and read-write, visible and invisible, to the
canonical path and to a scratch path, before and after killing WINWORD and
clearing the Resiliency key.

**What works: printing to the "Microsoft Print to PDF" driver** (2026-09-06).
A different code path, and it returns in seconds:

```powershell
$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
$w.ActivePrinter = "Microsoft Print to PDF"
$d = $w.Documents.Open($src, $false, $true)
$m = [Type]::Missing
# PrintOut(Background, Append, Range, OutputFileName, From, To, Item,
#          Copies, Pages, PageType, PrintToFile, Collate)
$d.GetType().InvokeMember("PrintOut", "InvokeMethod", $null, $d,
  @($false, $false, 0, $out, $m, $m, 0, 1, $m, 0, $true, $true))
$d.Close(0); $w.Quit()
```

Delete the target first (the driver appends to an existing file), and give
the print spooler a few seconds after `Quit()` before checking for the file.

**Verify the result with PyMuPDF, not a regex.** These PDFs use subset fonts,
so scanning the raw bytes or the inflated streams for "G305" returns False
even when the page plainly shows it -- the same trap that made an earlier
syllabus check look like a failure. `fitz.open(path)` and `page.get_text()`
read it correctly. PyMuPDF and `pdftoppm` (MiKTeX) are both installed;
LibreOffice is not.

## 2026-09-06 - podcast labels, and no materials box where there is nothing to post

Three edits, all driven from `_calendar_content.py`, so both sections and
both documents followed from one change each.

**1. "Optional Podcasts (before class):"** replaces "Other podcasts:". Only
weeks 1 and 9 carry that group and both are in-person, so the rename covers
exactly the in-person weeks Nico asked for.

*This broke two GROUP_MODULE_OVERRIDES keys* -- they match on the label text
-- and `check_overrides()` is what would have caught it. Both keys now match
on the shorter, stabler snippet "Optional Podcasts".

**2. Practice videos no longer count as "videos"** for the materials box.
Module 4's slides were showing under week 6, which has neither a class nor
teaching videos -- only an optional practice video. `video_modules()` now
skips groups labelled "Practice…" and items named "Practice Video: …".

Result (same for both sections):

| week | box |
|---|---|
| 1 | In-Class (M1, M2) + Slides for Videos (M1) |
| 2, 3, 4, 7, 8 | Slides for Videos (M2 / M3 / M4 / M6 / M7) |
| 5, 9 | In-Class only (M3-5 / M6-8) |
| **6, 10, 11, 12** | **no box at all** |

Weeks 6 and 10 were the two Nico asked me to find: both carry practice
videos only. Weeks 4 and 8 carry both kinds and now list only the module
they actually teach (M4, M7 -- not M3's or M6's practice videos).

**3. An intro to a module taught entirely on video** says "Podcast
(<u>before</u> watching the Module 3 videos):" rather than "before class".
Derived from the same `WRAP_AFTER` table that drives the wrap-ups: a tail
beginning "watching" means the module never reaches a classroom, so the
intro is keyed to the videos too. Modules 3 and 6 today.

**Verification note.** Searching the .docx for the whole phrase reports
False -- Word splits it into runs around the underlined word ("Podcast (" /
"before" / " watching the Module 3 videos): "). Check the runs, not the
sentence.

Nothing to change in the CALENDARS for point 2: they never had a materials
box, only the single In-Class Material line inside the on-campus card, which
weeks 6 and 10 do not render.

**Confirmed by Nico (2026-09-06), so this is the rule, not an accident:** a
module split between video and class earns an entry in BOTH weeks. Module 4
appears under week 4 as "Slides for Videos" (its Part I videos) and under
week 5 as in-class material. Module 7 is the same case (videos week 8, class
week 9). Do not "tidy" one of them away.

---

## 2026-09-05 (FEMBA) - one content file, two sections

MGMT 405 runs twice in Fall 2026. Nico: "The material and due dates are the
same", only the meeting pattern and a few names differ. So FEMBA is NOT a
copy of anything -- both sections come out of the same content file and the
same builders, picked by **MGMT405_SECTION** (default `emba`).

```
python _build_calendar.py  --section femba
python _build_site.py      --section femba
python _build_syllabus.py  --section femba --md
python _deploy.py          --section femba
```

`SECTIONS` in `_calendar_content.py` holds everything that differs:

| | EMBA | FEMBA |
|---|---|---|
| subtitle | EMBA Section 2 (Hybrid) | FEMBA Section 2 (Hybrid) |
| room | ~~A301~~ **G305** | ~~G-402~~ **G305** (both sections, 2026-09-06) |
| meetings | Fri 4:00–5:30 pm + Sat 9:00 am–12:30 pm | **Sat 2:00–8:00 pm only** |
| Bruin Learn | courses/237825 | courses/237860 |
| TA mailbox | ta405.emba2@gmail.com | **still to come** |
| repo | MGMT-405-EMBA | MGMT-405-FEMBA |

Two helpers, `class_when(week)` and `class_days_line(week)`, build every
sentence that names a meeting day or time, so the hard-coded "Fri … 4:00 –
5:30 pm · Sat … 9:00 am – 12:30 pm" strings are gone from all three builders.

### The ordering rule that makes it work

`_calendar_content` reads `MGMT405_SECTION` **at import time**, and the
builders use `from _calendar_content import (...)`, which binds values at
import. So every builder parses `--section` into `os.environ` in a loop
placed ABOVE its imports. Move that loop below an import and the flag
silently does nothing -- there is a comment saying so at each of the four
sites.

### Output layout

EMBA keeps `Course Website/` exactly as it was, so the live EMBA site was
never at risk; FEMBA builds into `Course Website/femba/`. `site.css` and
`site.js` stay hand-authored in ONE place and are copied into the FEMBA
folder on every build, so the two cannot drift. The EMBA deploy only globs
top-level `*.html`, so `femba/` is invisible to it (checked: still 23 pages).

Calendars and syllabi share their folders and differ by filename
(`Calendar FEMBA Hybrid -- Fall 2026.docx`, `Course Syllabus - 405 FEMBA
Fall 2026.docx`).

### Open

1. ~~The FEMBA TA mailbox~~ -- done 2026-09-06: `ta405.femba2@gmail.com`
   (EMBA keeps `ta405.emba2@gmail.com`). One line in `SECTIONS`; both
   calendars, both websites rebuilt and redeployed. The websites obfuscate
   each address, so neither appears in plaintext on its page.

   **The SYLLABUS deliberately carries no address at all** and still does.
   Its own docstring records why (2026-09-04, Nico): the syllabus PDF is
   published on the public site, so printing an address there would undo the
   obfuscation. Adding it was started and then reverted -- ask before
   changing that decision.
2. **Both FEMBA PDFs are missing** -- Word's ExportAsFixedFormat hung again.
   The site is published with `--skip-docs`, so its syllabus and calendar PDF
   links 404 until the two files are exported and it is redeployed. That flag
   is never the default.

FEMBA calendar: 14 pages, one week per page. Live at
https://nvoigtla.github.io/MGMT-405-FEMBA/

---

## 2026-09-05 (week 1 "Other podcasts")

Reordered and relinked, in `_calendar_content.py` so both documents follow:

1. **Should We Really Behave Like Economists Say We Do?** — moved to the
   first bullet, and given the show's full title. The short form
   ("...Like Economists?") was ours; the episode's own URL slug carries the
   longer one.
2. The Cost-Benefit Principle — new Apple episode id `1000786324917`
3. The Opportunity-Cost Principle — new Apple episode id `1000786324632`

Both Apple links already existed and pointed at the 2020 postings
(`...488478204` / `...205`); only the `i=` episode ids changed. Verified in
the docx's `word/_rels/document.xml.rels` (Word stores hyperlink targets
there, NOT in `document.xml`) and on the live page. Calendar still 14 pages.

## 2026-09-05 (later) - recaps deleted, podcasts rescheduled, box taller

**Recaps gone.** All five "Recap of Module X" groups removed (weeks 2, 3, 4,
8, 10) -- old recordings, replaced by the podcasts. Their `recapN` entries
stay in `LINKS`, unused.

**Podcasts rescheduled.** An intro belongs to the week before a module's
first section is taught, a wrap-up to the week AFTER all its core material is
covered -- which is not necessarily after a class. The week-5 and week-9
sessions do Module 3 / Module 6 APPLICATIONS only, so those two modules are
fully covered by their videos and their wrap-ups sit with them, in weeks 3
and 7. `podcast_items()` takes `(module, kind)` specs.

The label says what to finish first. `WRAP_AFTER` in `_calendar_content.py`
holds one phrase per module and `podcast_when()` serves BOTH builders, so the
wording is written once and the calendar and the website cannot drift:

| module | wrap-up week | reads |
|---|---|---|
| 1 | 2 | after **class** |
| 2 | 2 | after **class and watching the Module 2 videos** |
| 3 | 3 | after **watching the Module 3 videos** |
| 4, 5 | 6 | after **class** |
| 6 | 7 | after **watching the Module 6 videos** |
| 7, 8 | 10 | after **class** |

**Label**: "Podcasts About Class Material:". **TA address**: lowercase
`ta405.emba2@gmail.com` (one string, `LINKS["ta_email"]`, plus the literal
the page-1 info block prints).

**Problem-set box taller.** 0.275" clipped the date's descenders;
`PSET_BOX_H` is 0.32" with 0.01" of vertical padding back. Paid for by
dropping the last point of the agenda's uniform row-height slack
(`median + 1` -> `median`). Still 14 pages, one week per page, legend on
page 1.

**The `.pdf` is STALE again.** `ExportAsFixedFormat` hung on the one attempt
Nico now allows per round (2026-09-05: "do one attempt to generate it. If
that's not working let me know and i'll generate it myself"). The docx is
correct and verified; Nico exports the PDF by hand. See the standing open
item below -- this is the third session it has failed in.

---

## 2026-09-05 - "Problem Set X" boxed in the agenda's Due / Exams column

Nico: the Due / Exams column on page 1 should spell the name out in full and
carry the same dark-red marker the website draws around a problem set inside
each week — "just narrow, so it fits inside the Due field".

- `agenda_due_text()` no longer abbreviates to "PS X". It now returns
  `(text, is_pset)` per entry, so the caller can tell a problem set from an
  exam window without re-parsing the label.
- Nico then asked for **rounded corners with the due date inside**. Word
  paragraph borders are always square and wrap one paragraph, so `pset_box`
  became a DRAWN roundRect -- `rounded_card` in a new `compact` mode -- 1.12"
  wide, holding "Problem Set X" (8.5 pt bold) over its date (8 pt), both
  NAVY. The two on-campus rows print white on navy, and white on the box's
  pale fill would be unreadable.
- Colors are the constants that were already there for the week cards:
  `DARKRED` `C00000` and `DUEWASH`, its 7 % wash over white, matching the
  website's `rgba(192,0,0,.07)`.

### Page 1 had ZERO slack, and the box costs real height

This is the part to remember. In the previous build the agenda table ended at
717.0 pt and the legend sat at 717.0 with the usable bottom at 741.6 -- page 1
was full to the point, and the whole block (heading, table, legend) has to
stay inside ONE drawn backing card. A drawn box costs ~4.5 pt per row more
than the two plain text lines it replaced, so five problem-set rows put the
legend on page 2 and the document at 15 pages.

What actually worked, after several dead ends:

- `compact` mode on `rounded_card`: no drop shadow (its `effectExtent` is
  what reserves layout space under an inline shape), `effectExtent` zeroed,
  no vertical text insets, no cushion, no paragraph gap. Defaults unchanged,
  so every other card in the calendar is untouched.
- **An explicit `height_in`.** `_measure_par` floors its font size at 9.5 pt,
  so a card set in 8.5/8 pt type measured LARGER than it renders -- which is
  why shrinking the fonts did nothing until the height was passed in
  (`PSET_BOX_H = 0.275`).
- Three small reclamations on page 1: agenda cell padding 50 -> 30 twips,
  the uniform row-height slack 2 -> 1 pt, and the legend separator 2/2 -> 0/1.

Dead ends worth not repeating: pinning the shape paragraph's line spacing to
the card height (Word does not clip an inline shape, so it changed nothing);
cell padding 30 -> 20 twips (rows were already content-driven); shrinking the
fonts alone (see the `_measure_par` floor above).

**Result: 14 pages, legend back on page 1 at 721.5 pt, every week on one
page.**

**Checked:** "Problem Set 1" measures 0.764" in Calibri Bold 9.5 against the
1.10" of usable width in the 1.32" column, so nothing wraps — the reason the
label had been abbreviated in the first place. `_check_pagination.ps1` still
PASSES at 14 pages, one week per page, and the docx carries 5 boxed
paragraphs and 5 `FBEDED` fills, one per problem set, with no "PS 1" left.

---

## 2026-09-04 (later) - Module 2's three videos retaped and measured

Nico supplied new Panopto links for Module 2 videos 1-3. `m2v1` / `m2v2` /
`m2v3` in `_calendar_content.py` now point at the NEW sessions (ids
`8fcb86ee...`, `406d26d1...`, `885b4e89...`, all `b4bb...`); the old `b08b...`
ids were last year's recordings and are gone from the deck.

**Running times read off Panopto, not typed:** 25:27, 22:51 and 23:31, so
**25 / 23 / 24 min** by the file's `round(seconds/60)` convention. Confirmed
twice - directly against `DeliveryInfo.aspx`, then with
`python _video_minutes.py m2`, which re-reads Panopto and flags disagreements
(none). The Panopto session names match the calendar's video titles exactly.

**Week 2 gained a line.** With all three lengths now measured the group
crosses `render_group`'s budget test (>= 2 videos, >= 20 min, all measured),
so week 2 now prints **"= 72 min of video in total"**. `_check_pagination.ps1`
re-run afterwards: **PASS**, 14 pages, week 2 still on page 4.

The two PRACTICE videos (`m2p1`, `m2p2`) were not retaped. They keep last
year's sessions and their stored 7 / 6 min; `_video_minutes.py` reports them
LOCKED because those sessions are not shared publicly.

---

**Follow-up the same day:** the glyph is now `U+1F3DB` **+ `U+FE0F`**, the
emoji presentation selector. A bare `U+1F3DB` is rendered as a flat monochrome
text glyph by some platforms -- which is what the first attempt produced -- and
only the selector forces the color emoji Nico attached. Worth remembering for
any future symbol: if an emoji comes out flat, it is missing FE0F.

## 2026-09-03 (on-campus glyph) - classical building in the class header

The on-campus card's header now carries **U+1F3DB, the classical building**,
matching the website. `render_weekend()` sets `hdr_glyph` per branch and
passes it to `card_header(..., glyph=hdr_glyph, inner_w_in=inner_w)`; the
"videos to watch" branch passes `None`.

**Measured before editing**, because this header is long and the calendar
right-aligns its glyphs at a tab stop -- a wrapped header would grow the card
and could break the one-page-per-week invariant. The text measures **5.14" in
Calibri Bold 11.5pt against the 6.07" tab stop**, so there is 0.73" of
clearance. `_check_pagination.ps1` **PASSES** afterwards (14 pages), and
`document.xml` carries 3 glyphs -- weeks 1, 5 and 9.

---

## 2026-09-03 (last change) - weeks 10 and 11 read as video content

`KIND_META["thanksgiving"]` and `KIND_META["examprep"]` now use `VIDEOYEL`
instead of `"FFFFFF"`, so the page-1 agenda rows for **weeks 10 and 11 are
pale gold rather than white**. Nico's reasoning: neither week has an in-person
component, so both count as video content. The band labels are unaffected
(both kinds have `legend=None`) and the three-item legend is hardcoded
separately, so it needed no change.

This was made alongside the same change on the website, where those weeks'
agenda badges now carry the video-content color. Rebuilt and re-exported
(backups rolled); `_check_pagination.ps1` **PASSES** at 14 pages.

---

## 2026-09-03 (later still) - problem-set cards in dark red

A problem set's due card now has a **dark-red rule (`DARKRED = "C00000"`) over
a very transparent dark-red wash (`DUEWASH = "FBEDED"`)**, matching the
website. Week 11's Practice Final Exam is not a problem set, so it keeps the
gold rule - the gate is `label.lower().startswith("problem set")`.

**Word has no alpha channel.** The website uses `rgba(192,0,0,.07)`; `DUEWASH`
is that colour flattened onto white by hand: `0.07 x 192 + 0.93 x 255 = 251`
(`FB`) for red and `0.07 x 0 + 0.93 x 255 = 237` (`ED`) for green and blue. If
the website's alpha changes, recompute rather than guessing a hex.

Rebuilt and re-exported (backups rolled). `_check_pagination.ps1` **PASSES** -
14 pages. Verified 5 `C00000` borders and 5 `FBEDED` fills in `document.xml`,
one per problem set.

---

## 2026-09-03 (last edit of the day) - math review appendix line deleted

Deleted `"Textbook reading: Math review Appendix Section 1 + Section 2 (only
first derivatives)"` from `MATH_REFRESHER_ITEMS`. That list is now down to two
items, the Math Quiz and the Math Review Videos.

This was the one line still printing in the calendar but hidden on the website.
With it gone, **the calendar and the website carry exactly the same content** -
the website's `DROP_*` hooks are both empty now.

Rebuilt the `.docx` and re-exported the `.pdf` (backups rolled first).
`_check_pagination.ps1` **PASSES** - 14 pages, one page per week.

---

## 2026-09-03 (later the same day) - clapperboard glyph, two lines deleted

Follow-up to the changes below, again for consistency with the course website.

1. **Videos card glyph is now a clapperboard** (U+1F3AC) - the Hollywood
   "action" symbol. It replaced the play triangle, then briefly the film reel
   (U+1F39E), which Nico did not like. Video bullets are unaffected.
2. **Deleted the textbook note** "For additional and optional practice, you
   can find review questions at the end of each chapter ..." from
   `TEXTBOOK_NOTES`.
3. **Deleted the "Achieve" item** from `MATH_REFRESHER_ITEMS`.

`TEXTBOOK_NOTES` is down to 2 entries and `MATH_REFRESHER_ITEMS` to 3.

**Still in the calendar, dropped on the WEBSITE only:** "Textbook reading:
Math review Appendix Section 1 + Section 2 (only first derivatives)". Nico
named only the two lines above for calendar deletion.

Rebuilt the `.docx` and re-exported the `.pdf` (backups rolled first).
`_check_pagination.ps1` **PASSES** - 14 pages, one page per week.

---

## 2026-09-03 - podcast cards grey, film-reel glyph, class card on top

Three changes requested while building the course website, so that the two
documents keep the same visual language.

### Files worked on

| File | Change |
|---|---|
| `_build_calendar.py` | `PODGRAY` constant; podcast card fill; Videos card glyph; `render_weekend()` hoisted |
| `Calendar EMBA Hybrid -- Fall 2026.docx` / `.pdf` | rebuilt and re-exported (`_t-1` backups rolled first) |

### What changed in the document

1. **Podcast cards are light grey** (`PODGRAY = "EFF1F4"`), not white, so the
   three categories read apart at a glance: videos pale gold, podcasts grey,
   suggested reading white.
2. **The Videos card header glyph is a film reel** (U+1F39E) instead of the
   play triangle. The calendar's video BULLETS never used a triangle, so
   nothing else changed. Same glyph on the website.
3. **On an on-campus week the class card now sits ABOVE the preparation
   container.** The class is the anchor of the week. The weekend block moved
   into a local `render_weekend()` that is called either before or after the
   prep section, keyed on a new `weekend_first` flag.

### Verified

- `_check_pagination.ps1` **PASSES** - every week still fits on one page, 14
  pages in total. The reorder does not change any card's height.
- In `word/document.xml`, the class-card header ("4:00 - 5:30 pm") now precedes
  the "Before class" prep heading in weeks 1, 5 and 9; week 2 (a video week)
  is unchanged and still reads "During the week"; there are 8 `EFF1F4` fills,
  one per week with a podcast group.
- PDF re-exported through Word COM (no LibreOffice on these machines).

### Context worth keeping

- **The website mirrors these three choices**, and its README records them.
  The site is generated from `_calendar_content.py` by
  `../Course Website/_build_site.py`, so a content edit here updates both
  documents: rebuild the calendar, then run the site build and deploy.
- Colour emoji ignore the CSS/Word text colour, so the glyphs render in their
  own colours rather than the pale gold the code asks for. That was already
  true of the headphones and book glyphs.

---

## 2026-08-28 – recolor, podcasts, video names, one-page weeks

Reworked the calendar to match the slide decks' color coding, gave every
module its two podcasts, adopted the module 1 – 3 video names from the module
folders, and made "every week on one page" an enforced invariant.

### Files worked on

| File | Change |
|---|---|
| `_build_calendar.py` | category colors, dark-gray agenda header + white gap row, `"Video content"` label, `"p"` podcast item kind, `(++)` video lengths, week-page spacing constants, band-label width guard |
| `_calendar_content.py` | `PODCASTS` registry + `podcast_items()`, podcast groups in 8 weeks, module 1 – 3 video names and lengths, new module 1 Panopto links |
| `_podcast_minutes.py` | NEW – reads podcast durations off Dropbox |
| `_video_minutes.py` | NEW – reads Panopto durations via `DeliveryInfo.aspx` |
| `_check_pagination.ps1` | NEW – asserts one page per week (Word COM) |

### What changed in the document

1. **Category colors, matching the decks.** In class = navy `0B2B4E`
   (white text), videos = light yellow `F6E8C9`, exams = darker yellow
   `E09F3E`. Applies to the page-1 agenda rows, the legend, the Videos card
   body, the on-campus card (navy header + pale blue `E7EDF4` body) and the
   exam card header.
2. **Agenda header row** is dark gray `555B66` at 12 pt with an 8 pt white
   spacer row beneath it, so it reads apart from Week 1's navy row.
3. **"Video deadline (suggested)" retired** in favour of **"Video content"**
   in the legend and `KIND_META`. The week band keeps the date and reads
   `Video content  ·  suggested: Sun, <date>`.
4. **Podcasts: two per module**, replacing "Posted on Bruin Learn" –
   `Podcast: Intro to Module X` and `Podcast: Wrap-Up of Module X`.
5. **Module 1 – 3 class video names** taken from the module folders. Module 3
   was retaped as **7** videos (was 6).
6. **Every week fits on one page**; the deck is 14 pages.

### Decisions

- **Week band stays navy on every page.** It is page chrome, like a slide's
  top bar, not a category marker. Only the content cards carry the coding.
- **The band label keeps its date.** "Video content" alone would drop the
  information, so the wording leads with the new term and keeps the date.
- **Practice videos and module recaps were NOT touched** by the video-name
  and `(++)` pass. Only the videos that teach the module itself changed.
  Recaps still read 7 / 8 / 7 min.
- **Rounding** for all durations is `round(seconds / 60)`.
- **Podcast module to week mapping** (the week that covers that module):

  | Week | 1 | 2 | 3 | 4 | 5 | 7 | 8 | 9 |
  |---|---|---|---|---|---|---|---|---|
  | Module | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 |

### Open items

- **Module 2 (3 links) and Module 3 (6 links) still point at last year's
  Panopto sessions.** Module 1's were re-recorded as new sessions
  (`b4b3…` ids), so modules 2 and 3 almost certainly need new URLs too.
- **Module 3's link keys are shifted one topic down** (`m3v1` now sits on
  "Video 2: The Production Function", and so on), because the new
  "Introduction to Module 3" joined at the front. That video has **no link
  at all** yet. Re-check the mapping when the new URLs arrive.
- **Module 2 and 3 video lengths are `(++)`**. `_video_minutes.py` fills
  them in, but only for sessions shared publicly.
- **Podcast links for modules 4 – 8** are placeholders (plain text, no link,
  no duration).
- **Week 1 has only 0.22 in of slack**; every other week has 1.15 in or more.
  Anything added to week 1 will push it to a second page. Run
  `_check_pagination.ps1` after any content change.
- **`.gitignore` ignores `*_t-1.pptx` / `*_t-2.pptx` but not the `.docx`
  equivalents**, so this folder's rolling backups show up as untracked.
- **Panopto session titles vs. deck names**: Nico renamed module 1's videos 2
  and 3 on Panopto so they now match the calendar. Worth checking the same
  for modules 2 and 3 when their links land.

### Commands

```
python _build_calendar.py                    # rebuild the .docx
python _podcast_minutes.py                   # podcast durations (Dropbox)
python _video_minutes.py            # all linked videos (Panopto)
python _video_minutes.py m2 m3      # just those modules
powershell -File _check_pagination.ps1       # one page per week?
```

PDF is exported through Word COM (no LibreOffice on these machines):

```powershell
$w = New-Object -ComObject Word.Application; $w.Visible = $false
$d = $w.Documents.Open("<abs path>.docx", $false, $true)
$d.ExportAsFixedFormat("<abs path>.pdf", 17); $d.Close(0); $w.Quit()
```

### Context worth keeping

- **`OUT` is now derived from the script's own folder** (overridable with
  the `CALENDAR_OUT` environment variable). It used to be a hardcoded `c:\`
  path, so the build could not run from the H: mapping at all.
- **`python-docx` had to be installed** on this machine; it was not present.
- **Durations are read, not typed.** Podcasts: a range request for the
  MP4/M4A `moov/mvhd` box, so nothing large is downloaded. Panopto: the
  viewer's own `DeliveryInfo.aspx` JSON, which answers anonymously only
  while a session is shared publicly and otherwise reports `LOCKED`. Both
  scripts re-check the stored numbers and flag disagreements, so they double
  as verification after editing.
- **Card heights are computed by `_measure_par` from the same spacing
  constants that lay out the text.** Changing a spacing constant moves the
  drawn box with the text. The one number to leave alone is the 0.18 in
  cushion inside `rounded_card`: it is what stops the approximation from
  clipping text out of a card.
- **This repo has concurrent activity.** Several commits appeared during the
  session that were not made from it, including the Course Calendar commit
  itself and a large Module 4 rebuild. Check `git log` before assuming the
  working tree is yours alone.

---

## 2026-08-30 – Module 4 re-split into five videos

**One-line summary:** Module 4's deck was converted for taping with a
different split than the calendar carried, so the Session-4 prep block was
updated to list five videos with placeholder lengths.

### What changed
`_calendar_content.py`, the Session-4 `"video"` group for Module 4 (Part I):

    Video 1: Introduction to Market Structures      (++)
    Video 2: Perfect Competition                    (++)
    Video 3: Profit Maximization of a Price Taker – Short Run  (++)
    Video 4: Firm-Level and Market Supply           (++)
    Video 5: Long-Run Competitive Equilibrium       (++)

Was four videos with real running times (13 / 47 / 6 / 11 min).

### Decisions
- **Lengths are `None`, which the builder already prints as `(++)`** — the
  same placeholder Modules 2 and 3 use. It also suppresses the
  "≈ N min of video in total" line, which would be meaningless.
- **Links are `None` too.** The four Panopto URLs (`m4v1` – `m4v4`) point at
  the OLD four-video cut, where Video 2 was the profit-max video and there
  was no separate Perfect Competition video — linking them would send
  students to the wrong recording. The URLs are still in `VIDEO_LINKS`,
  untouched, ready to be re-pointed.

### Open
1. Re-record the five videos, then put the real minutes in as the fourth
   element of each tuple (~line 313) and re-point `m4v1` – `m4v5`.
   `_video_minutes.py` can read the lengths off Panopto.
2. **The PDF was not regenerated** — it still shows the old four-video
   version. Export it from Word once the deck is final.


## 2026-08-31 - practice-exercise quizzes, podcast label, exam boxes rebuilt

Added the TA's practice site to every week that follows a module, relabelled
the podcast groups, and rebuilt both exam weeks as a single dark-yellow box.

### Files worked on

| File | Change |
|---|---|
| `_calendar_content.py` | 10 `prac_*` links; a `"practice"` group in 7 weeks; podcast label; both exam dicts rewritten (title, window, lines) |
| `_build_calendar.py` | `"practice"` category card; `exam_window()` helper; exam card moved into the topics-card slot |
| `Calendar EMBA Hybrid -- Fall 2026.docx` / `.pdf` | rebuilt and re-exported |

### What changed in the document

1. **Podcast group label** is now `About Class Material:` (was `On class
   material:`) in all 8 weeks. Nico's capitalisation, kept verbatim; note the
   sibling label `Other podcasts:` is still sentence case.
2. **New card `Suggested Additional Practice Exercises`**, a fourth category
   after Suggested Reading, glyph **U+270E** (the decks' problem-set mark),
   navy header / white body. Registered in the card tuple in
   `_build_calendar.py`, so it renders wherever a week has a `practice` group.
3. **Ten links**, all live (HTTP 200 checked), labelled
   `Online quiz on Module N: <topic>`. Source:
   `https://rafaelrubiao.github.io/mgmt405-practice/` - `module-N.html`, with
   `-part-1` / `-part-2` for Modules 4 and 7.
4. **Placement rule (Nico's call, this session):** the week AFTER the module's
   own TEACHING week. Modules 3 and 6 follow their VIDEO week, not the later
   "Applications" week, so Module 3's set is available before the Week 6
   midterm that covers through Module 3.

   | Week | 2 | 4 | 5 | 7 | 8 | 9 | 10 |
   |---|---|---|---|---|---|---|---|
   | Modules | 1, 2 | 3 | 4-I | 4-II, 5 | 6 | 7-I | 7-II, 8 |

5. **Exam weeks rebuilt.** Weeks 6 and 12 previously showed three separate
   boxes (a top banner added earlier the same day, a plain topics card, and a
   detail card at the bottom). They are now ONE box in the topics-card slot:
   dark-yellow `E09F3E` header carrying the name only (`Midterm Exam
   (online)` / `Final Exam (online)`), cream body, bullets below. The topics
   card and the bottom exam card are suppressed on exam weeks.
6. **Exam windows moved:** midterm Thu-Sun -> **Fri-Sat, Oct 30 - 31**; final
   Fri-Sun -> **Sat-Sun, Dec 12 - 13**. Page 1 reads `Midterm window` /
   `Final Exam window`.
7. **Exam wording:** exact time window to be determined and announced in
   class; both exams online with proctoring software; both `Open book, open
   notes. Calculator allowed.`

### Decisions

- **One source for the exam dates.** They had been hard-coded in three places
  (p.1's Due/Exams column, `band_right_label`, the exam card). All three now
  call the new `exam_window(wk)`, which reads the exam's own `window` tuple -
  change a date once in `_calendar_content.py` and every surface follows.
- **Practice links carry the CALENDAR's module names**, not the site's bare
  "Module 1" link text, so they match the wording students see elsewhere.
- **The redundant sentence was dropped** from the final's question-count
  bullet ("The exam will cover all material.") because the new first bullet
  now says it. Flagged to Nico; one word restores it.
- **Week 12's own span stays Dec 11 - 13** even though the exam window is now
  Dec 12 - 13: Week 11 ends Dec 10, so shortening Week 12 would leave Dec 11
  in no week at all.

### Open items

1. **Module 8's practice page is titled just "Auctions"** (11 questions),
   while the calendar's Module 8 is "Asymmetric Information; Auctions". The
   link uses the calendar's wording. A draft email to Rafael asking whether
   the asymmetric-information half is covered was written in chat this
   session but NOT saved to a file - rewrite it if still needed.
2. **The midterm box states no dates** (only the week band and p.1 do), while
   the final's box spells its window out. Deliberate for now - the midterm's
   time is TBD - but the asymmetry is worth a look.
3. **No practice card in Week 11 (exam prep).** A consolidated all-modules
   card was offered and not taken; Week 12 cannot host one (no prep section).
4. Everything from the 2026-08-30 entry still stands: Module 4's five video
   links and running times are still placeholders.

### Commands worth remembering

- Rebuild: `python _build_calendar.py` (writes the .docx in place).
- One-page-per-week invariant: `_check_pagination.ps1` (Word COM). Run after
  ANY content change; it still passes at 14 pages with all seven new cards.
- PDF: Word COM `ExportAsFixedFormat($path, 17)`; there is no pandoc step.
- Eyeball a page without opening Word:
  `pdftoppm -f N -l N -r 100 -png "<pdf>" out` (MiKTeX ships pdftoppm; there
  is no PyMuPDF / pdf2image on this machine).
- Measure a label before trusting it to fit: PIL `ImageFont.truetype`
  on `C:/Windows/Fonts/calibrib.ttf`; practice-card text width is 5.566 in.

---

## 2026-09-01 - Module 4 podcast links, new Module 1 Video 1 session

Pasted in Module 4's two podcast links, read their running times off Dropbox,
and swapped Module 1's Video 1 for the re-uploaded Panopto session.

### Files worked on

| File | Change |
|---|---|
| `_calendar_content.py` | `PODCASTS[4]` intro + wrap links and minutes; `LINKS["m1v1"]` new session id; Week 1 `m1v1` length |
| `Calendar EMBA Hybrid -- Fall 2026.docx` | rebuilt three times (intro only, then both podcasts, then the new video link) |

### What changed in the document

1. **Module 4 podcasts are live links**, replacing the plain-text
   placeholders: intro **3:35 -> 4 min**, wrap-up **21:47 -> 22 min**, both
   read by `_podcast_minutes.py` straight off the Dropbox files. Modules 5 - 8
   are still placeholders.
2. **Module 1, Video 1: Introduction** points at the re-uploaded session
   `id=b7207fac-c4f6-45fd-be9c-b4b800fc0e3f` (was `45ebea5c-...a768`), length
   **8:53 -> 9 min**. Week 1's four videos now read 9 / 10 / 8 / 7 min.

### Decisions

- **The wrap-up file is named `Module-4-Video-Wrap-Up.m4a`**, not
  `Module-4-Podcast-Wrap-Up.m4a` like modules 1 - 3. It is an `.m4a` and Nico
  supplied it as the wrap-up podcast, so it is wired to the
  *Podcast: Wrap-Up of Module 4* bullet as-is. Renaming it on Dropbox would
  change the share URL, so the naming inconsistency stays.
- **A re-recorded video loses its old running time.** When the new m1v1
  session was still private, the length was set to `None` (prints `(++)`)
  rather than carrying the old 9 min forward - the old number described a
  different file. It happened to come back as 9 min once the session was
  shared.
- **No PDF this round** (Nico: "you can create only the word file"), so the
  committed `.pdf` now lags the `.docx`.
- Backups were rolled before each content-changing rebuild, but NOT before
  the one rebuild that only regenerated identical content - rolling there
  would have pushed a good version off the end of the chain.

### Open items

1. **The `.pdf` is stale.** Word's COM `ExportAsFixedFormat` hung twice
   (Word started, stayed Responding, wrote nothing; no modal dialog was
   visible in an `EnumWindows` dump). Killing `WINWORD` and exporting to a
   scratch path did not help either. Retry, or Save-As from Word by hand.
2. **Dropbox share URLs cannot be derived from the file name** - the `fi` id
   and `rlkey` are random per file. Two of the links pasted for Module 4 were
   byte-identical, so the wrap-up had to be requested separately. Worth
   checking the two URLs differ before building.
3. **`(++)` remains on 15 videos** - Module 2 (3), Module 3 (7), Module 4 (5).
   Their Panopto sessions are not shared publicly.
   `python _video_minutes.py m2 m3 m4` fills them in as each is shared.

### Commands worth remembering

- Podcast lengths: `python _podcast_minutes.py` (range-GETs the Dropbox
  `.m4a` and parses `moov/mvhd`; prints the minutes to paste into `PODCASTS`).
- Video lengths: `python _video_minutes.py m1` - `LOCKED (not shared
  publicly)` means the session needs its share setting changed, not that the
  link is wrong. Panopto's `oembed` endpoint 404s here and `Embed.aspx`
  carries no duration, so `DeliveryInfo.aspx` is the only route.
- Confirm a link actually landed in the built file:
  read `word/_rels/document.xml.rels` out of the `.docx` with `zipfile` and
  grep the session id / file name.
