# Session Notes – Course Website (MGMT 405 EMBA Hybrid, Fall 2026)

Deliverable: a static multi-page course website, generated from
`../Course Calendar/_calendar_content.py`. See `README.md` for the pipeline,
the page layout and the palette.

- **Live:** https://nvoigtla.github.io/MGMT-405-EMBA/
- **Public repo (built output only):** https://github.com/nvoigtla/MGMT-405-EMBA

---

## Where this stands (end of 2026-09-04)

**The site is built, live and current.** 23 pages generated from the course
calendar, published to a separate PUBLIC repo and served by GitHub Pages. The
calendar was restyled in step with it over the same two days.

To pick this back up: `python _build_site.py` then `python _deploy.py`. The
calendar is `../Course Calendar/`, rebuilt with `python _build_calendar.py`
plus a Word-COM PDF export and `_check_pagination.ps1`.

**Open items, in the order they are likely to matter:**

1. `SYLLABUS_URL` in `_build_site.py` is `"#"` -- paste the Bruin Learn or PDF
   address and rebuild.
2. The In-Class Material rows are all "(TBD)". Nico uploads the handouts and
   slide decks right before each class; they need real links then.
3. The podcast "(before class)" / "(after class)" wording is website-only. The
   calendar still prints the plain "Podcast: Intro to Module 1", because its
   item format carries no per-run formatting for podcasts.
4. Content gaps the site surfaces honestly: Module 3 video 1 and all five
   Module 4 videos read "(link to follow)"; Modules 5 - 8 podcasts likewise;
   Module 2 and 3 video running times read "(++)".
5. The Dropbox podcast links are reachable from the public page. Reviewed and
   accepted -- repo visibility is irrelevant, since the page itself is public;
   the only real fix would be hosting the audio behind Bruin Learn.

---

## 2026-09-04 - contact box, help popover, and two layout bugs

The rounds after the tab-title change, recorded together.

### What changed

| | |
|---|---|
| General Logistics | "How the Quarter Runs" deleted; **Class and Contact** promoted to the top left; new **Help and Questions** box in the top right |
| Help and Questions | who-to-ask-about-what; only "TA <name>" / "Prof. Nico" are links, bold with a gold underline; "Practice Exercises," dropped from the list; an extra non-breaking space before each name |
| Top bar | a **help button** left of View by, opening Help and Questions as a popover (closes on outside click or Escape). "?" top-left and envelope bottom-right in one box, 48x47, matched to View by's height; glyphs 30px / 26px on the desktop, 19 / 17 on a phone |
| Email addresses | **obfuscated against harvesters** -- see below |
| Module 1 | lost its "Back to General Logistics" link, as week 1 had |
| In-Class Material | now reads the week's TOPICS too, so week 9 picks up Module 6 |
| Landing band | "Before You Start" removed, term set larger; also removed from the left menu |
| Tab | every page reads "Managerial Econ 405" (`TAB_TITLE`) |
| Dormant card | "Videos to watch" got the clapperboard for completeness |

### Email obfuscation

`mail_link()` splits an address at the "@" and base64-encodes each half into a
data attribute; `initMail()` in site.js reassembles it and sets the href at
load. **The served HTML contains no address, no `mailto:` and no "@"** -- the
only "@" characters left on any page are inside the Google Fonts URL. Applied
to Prof. Nico's new address AND the TA's, which had been sitting in the source
as plaintext on all 23 pages.

It defeats harvesters that regex raw HTML and do not run JS, which is nearly
all of them. It does NOT hide the address from a JS-executing crawler. With JS
off, the names are inert text rather than broken links.

### Two layout bugs, both found by MEASURING rather than looking

**1. The three columns shifted horizontally from page to page.** Nico noticed
the frames "moving around" between weeks. Cause: `.shell` carries
`margin:0 auto`, and on the desktop `body` is a flex column -- **auto margins
on the cross axis make a flex item shrink to fit its CONTENT instead of
stretching**. So the whole grid sized itself to whatever that week's content
needed: the left column began at x=143 on week 1, x=94 on week 2, x=215 on
week 3, x=22 on week 11. Fixed with an explicit `width:100%`; all six weeks
measured afterwards report an identical L 22|272, M 316|874, R 1212|244. The
same trap had been rendering the **footer** as a narrow centred strip.

**2. The name in Help and Questions sat on its own line.** It looked like a
stray `<br>`. It was not a wrap either: `.qlinks a` is styled as a flex ROW
(glyph + label) for the link lists, and the new inline mail link inherited
`display:flex`, which makes an element a BLOCK and forces a break before it.
My first attempt assumed wrapping and added a no-break space, which changed
nothing; an isolated test page showed the binding working fine, and only
diffing against the REAL css found the inherited rule. Fixed with
`.qlinks a.mail{display:inline}`.

**The lesson from both:** a screenshot at one window size hid the first bug
entirely and misled me about the second. Comparing rendered geometry across
pages, and testing against the real stylesheet rather than a minimal
reproduction, is what actually found them.


---

**Follow-up the same day:** the glyph is now `U+1F3DB` **+ `U+FE0F`**, the
emoji presentation selector. A bare `U+1F3DB` is rendered as a flat monochrome
text glyph by some platforms -- which is what the first attempt produced -- and
only the selector forces the color emoji Nico attached. Worth remembering for
any future symbol: if an emoji comes out flat, it is missing FE0F.

## 2026-09-03 (fourteenth pass) - tab title, in-class material, podcasts

### What changed

| Asked for | Done |
|---|---|
| The browser tab always reads "Econ 405" | `TAB_TITLE`; the page's own name stays in the band |
| Search box: desktop unchanged, phone at the very top | `.right{display:contents}` on the phone breakpoint, then `order` on the search box, main and `#deadlines` |
| On-campus classes: an italic "In-Class Material" section with Module X Handout / Slides, "(TBD)" | `inclass_material()`, appended to the class card |
| Module podcasts: "Podcast (before class): ..." / "(after class)", the timing word underlined | `podcast_label()` |

### The phone reorder needed `display:contents`

The search box has to stay inside the right column on the desktop, but sit
ABOVE the main content on a phone -- and CSS cannot move an element out of its
parent. `display:contents` on `.right` dissolves the wrapper at the phone
breakpoint, promoting the search box and the deadlines card to direct flex
children of `.shell`, where `order` can place them on either side of `main`.

Verified by measuring the rendered `top` of each region rather than by eye:
- 512px wide: search@84 < main@139 < deadlines@1341
- 1500px wide: search@91 = main@91 (side by side), deadlines@143 (under the
  search, in the right column) -- i.e. the desktop is untouched.

### Underlining lives on the website, not in the calendar content

The podcast label is rewritten at render time. Changing
`_calendar_content.py` instead would have put the new wording in the calendar
too, but the calendar's `("p", url, text, minutes)` item format carries no
per-run formatting, so the underline could not travel with it. The calendar
still prints "Podcast: Intro to Module 1". **Flagged to Nico.**

### One gap to confirm: week 9 and Module 6

`inclass_material()` reads the modules from the class card's own items. For
week 9 those are "Discussion: Application of Economic Concepts", "Module 7
(Part II)" and "Module 8", so the section lists Modules 7 and 8 only --
**Module 6 gets no handout or slides**, even though the week's `topics`
include "Module 6: Applications" (the Discussion item is what covers it).
Weeks 1 and 5 are complete (Modules 1-2 and Modules 3-5). Asked Nico whether
Module 6 should be added for week 9.

---

## 2026-09-03 (thirteenth pass) - on-campus glyph, bigger header

### What changed

| Asked for | Done |
|---|---|
| A symbol for the on-campus class header; Nico attached a classical building | 🏛 U+1F3DB, on the website AND in the calendar's on-campus header |
| Practice-exercises sub-title | "Online exercises with solutions" |
| Larger top-bar font on the desktop, phone unchanged | `.tb-tag` 19.5px, `.tb-sub` 16.5px inside `@media (min-width:861px)` |
| Delete the ↗ after the outgoing menu row | the `nav.list a.out .t::after` rule removed |

### No image asset was needed

Nico asked whether the attached image could be "processed". It reached the
session as inline pixels, not a file -- but it did not need to be: the picture
IS the standard rendering of U+1F3DB, so the symbol ships as a text character
in both documents. No hosting, no `assets/` entry, and it scales with the
font.

### Measuring before editing the calendar

The calendar puts header glyphs at a RIGHT tab stop, and the on-campus header
is long ("On-campus class · Fri, Sep 25, 4:00 - 5:30 pm · Sat, Sep 26,
9:00 am - 12:30 pm"). A wrapped header would grow the card and could break the
one-page-per-week invariant, so the text was measured in the real font first
(PIL, Calibri Bold 11.5pt): **5.14" against a 6.07" tab stop**, i.e. 0.73" of
clearance. Only then was the glyph added. Pagination re-checked afterwards:
**PASS**, 14 pages, and 3 building glyphs in the XML -- weeks 1, 5 and 9.

Also offered but NOT done, since Nico did not take it up: the same treatment
for the "Videos to watch" card header (which is currently dead code anyway --
only weeks 1, 5 and 9 carry a `weekend` block).

---

## 2026-09-03 (twelfth pass) - column scrolling, deadline filter, practice link

### What changed

| Asked for | Done |
|---|---|
| Each of the three columns should scroll separately on the desktop | body is a flex column with `overflow:hidden`; the top bar and footer are fixed bands; `.shell` takes the middle and each column gets `overflow-y:auto` |
| Phone: clicking "Deadlines for this week" shows ONLY that week | rows carry `data-week`; `initDueFilter()` hides the rest and reveals a "Show all deadlines" button |
| Menu glyphs should match the week headers | clapperboard and headphones, from the same constants |
| "All Practice Exercises" at the bottom of the menu | new `EXTRA_LINKS` -- an OUTGOING link (new tab, ↗ mark), in both menus and the jump menu |

### The scrolling bug was worse than "patchy"

Nico reported the left column scrolling only after the middle finished --
the symptom of the whole PAGE scrolling with two sticky side columns. Giving
each column its own scroller fixed that, but the first attempt hid a second
bug: the columns reported `scrollHeight == clientHeight`, i.e. no overflow at
all, even at a 640px window.

Cause: `.left` and `.right` are flex containers, and **a flex child shrinks by
default**. The menu card was being squeezed below its content height, and
`.card { overflow:hidden }` then CLIPPED the rows -- weeks 8 to 12 and the
three index rows were simply gone on a short window, with no scrollbar to
reveal them. `.left > *, .right > * { flex:none }` fixes it.

Worth keeping: this was only caught by MEASURING `clientHeight` against
`scrollHeight` per column. A screenshot at 900px tall looked perfectly fine,
because the clipping only bites when the window is short.

Now, at 1500x640: left 383/971, main 383/1118, right 383/1371 -- three real
scrollers.

### The deadline filter, verified by simulated click

A probe clicked the header link and counted visible rows: 15 before, **2
after** (week 3's video-watch row and Problem Set 1), the "Show all
deadlines" button appeared, and clicking it restored all 15. The click
handler deliberately does NOT preventDefault, so the browser still performs
the `#dl-here` anchor jump -- and that row is still visible after filtering.

### Auto-landing: re-confirmed, it was already in place

Nico asked whether landing on the current week was implemented. It was, from
the second pass; re-verified now with a stubbed clock across five dates:

| Simulated date | Lands on |
|---|---|
| Sep 3 (before the quarter) | General Logistics |
| Sep 23 (inside week 1) | Week 1 |
| Oct 28 (inside week 6) | Week 6 |
| Dec 12 (inside week 12) | Week 12 |
| Dec 28 (after the quarter) | General Logistics |

and `index.html?stay=1` -- the menu's own link -- always stays on General
Logistics.

---

## 2026-09-03 (eleventh pass) - menu order, weeks 10/11, phone deadline link

### What changed

| Asked for | Done |
|---|---|
| All Videos / All Podcasts move after the weeks | `head` / `tail` split in `left_column`; in the jump menu they became their own `extra` group, emitted last and shown in both modes |
| Weeks 10 and 11 count as "Video content" in the agenda | website: a badge rule for `data-kind="thanksgiving"` and `"examprep"`. **Calendar too**: `KIND_META` fills changed from white to `VIDEOYEL` |
| Phone only: a link in the week header down to that week's deadlines | `.dl-jump` in the band, `display:none` above 860px; targets `#dl-here` -- the first deadlines row belonging to that week -- with `scroll-margin-top` to clear the sticky bar |

### The calendar was changed too, deliberately

Nico said "in the agenda", and he has used that word for the calendar's page-1
table before (when asking for the color key). The website's week list carries
the same color scheme, so leaving the two disagreeing would have been worse
than changing both -- and he has asked repeatedly this session to keep the
documents in step. So `KIND_META["thanksgiving"]` and `["examprep"]` now use
`VIDEOYEL` as well, and the page-1 agenda rows for weeks 10 and 11 are pale
gold instead of white. Rebuilt, PDF re-exported, backups rolled;
`_check_pagination.ps1` **PASSES** at 14 pages. **Flagged to him** in case he
only meant the website.

### The phone deadline link

Every week happens to have at least one deadline row (a video-watch row, a
problem set, or an exam), so `#dl-here` always exists -- but the generator
still falls back to `#deadlines`, the card itself, if a week ever has none.
`week_has_deadlines()` decides. The id is written onto the FIRST matching row
only, guarded by a `marked` flag, so it can never be emitted twice.

### Verified

- Both menus and the jump menu end with All Videos then All Podcasts; the jump
  menu carries 15 entries in week view (General + 12 weeks + 2 indexes).
- `id="dl-here"` appears exactly once on each of weeks 1, 5, 11 and 12, and
  `dl-jump` appears on no index or module page.
- Weeks 10 and 11 render with the pale-gold badge; screenshot at 1500px.
- Live site re-checked by URL after the Pages build.

---

## 2026-09-03 (tenth pass) - All Videos / All Podcasts, menu labels

### What changed

| Asked for | Done |
|---|---|
| The View by button also returns to the main page | stores the new mode, then navigates to `index.html?stay=1` |
| The left menu says "Week 1" / "Module 1" | first line is the identity, second line the dates and coverage; left column widened 250 -> 272px |
| Weeks 1, 5, 9 take the transparent navy fill, like the color key | `nav.list a[data-kind="oncampus"] .n` -- navy rule over a 10% navy wash |
| Week 1 loses its "Back to General Logistics" | `back = (None, None)` for week 1 |
| New "All Videos" page, by module, practice videos included | `all-videos.html` |
| New "All Podcasts" page, intro and wrap-up by module | `all-podcasts.html` |

### One declaration drives four things

`EXTRA_PAGES` at the top of `_build_site.py` lists the two new pages once --
href, key, title, sub-line, badge glyph -- and that single list drives the
pages themselves, the week menu, the module menu and the jump menu. Adding a
third index page later is one tuple.

### Modules with nothing to list say so

Modules 5 and 8 have no videos at all, so `all-videos.html` would simply have
skipped them, which reads like an omission. They now get a card saying "This
module is covered in class, in week 5 / 9", derived from the week data rather
than typed. The podcast page gets a closing "Not tied to one module" card for
the general-interest episodes.

### A deploy bug worth recording

**`_deploy.py` carried a hardcoded page list**, so the two new pages were
built locally, pushed nowhere, and returned 404 on the live site while
everything looked right in the local build and in the git diff. Fixed by
discovering pages instead: `page_files()` globs every `.html` in the folder.
The lesson is narrow but sharp -- a publish step that enumerates its inputs by
hand will silently omit anything new. Verified afterwards by HTTP: index,
all-videos, all-podcasts, week-01 and module-3 all return 200.

### Notes

- **Module 1 still has its "Back to General Logistics" link.** Nico asked only
  about week 1; the same argument applies to module 1, so it is worth
  confirming rather than assuming.
- Both new pages include recaps and practice videos, since they are indexes
  rather than to-do lists. Unposted videos still show "(link to follow)".

---

## 2026-09-03 (ninth pass) - light only, View by button, all video deadlines

### What changed

| Asked for | Done |
|---|---|
| Always light mode | `data-theme="light"` stamped on `<html>`; the dark palette and all dark-specific rules stripped from `site.css` (26.7 -> 24.1 kB); the sunrise/sunset code and the 50-zone table deleted from `site.js` |
| Replace the button with "View by Week" / "View by Module", two lines, default Week | `.viewbtn` with `.l1` "View by" over `.l2` "Week"/"Module" |
| Capitalize By Week / By Module in the left menu | done |
| On-campus swatch: navy line + transparent navy fill, like the card | `.legend .sw.campus` |
| Hide the color key and week/module choices on a phone | the whole `.left` column is `display:none` below 860px |
| The hamburger shows only weeks or only modules, following View by | one shared mode drives three controls; the select is re-emitted from a snapshot |
| **Deadlines & Exams: the videos beyond week 1 were missing** | `video_watch_row()` rewritten -- now 7 rows, one per week that requires videos |

### The video-deadline rule, spelled out

The old version only looked at on-campus weeks, which is why just week 1
appeared. The rule now is:

- take every `cat == "video"` group whose label does NOT contain "optional"
  and does NOT start with "Recap";
- numbered items ("Video 3: ...") become a range, "Videos 1 - 7";
- unnumbered practice items ("Practice Video: ...") are counted, "2 practice
  videos" -- which is what makes weeks 8 and 10 appear at all, since their
  only required viewing IS practice videos;
- both are grouped per module, so week 2 reads "Videos 1 - 3 and 2 practice
  videos for Module 2" rather than repeating the module twice;
- the deadline is the Friday class on an on-campus week, the week's Sunday
  otherwise -- the same "suggested" date the calendar prints.

Result: weeks 1, 2, 3, 4, 7, 8, 10. Week 6's only video is "Practice on
Module 4 (optional)", so it correctly produces nothing.

### A mistake worth recording

**The round-7 splice deleted `jump_select()`.** I replaced everything between
`def video_watch_row(w):` and `def assessments():`, not noticing that
`jump_select()` had been inserted between those two markers in an earlier
round. The build failed immediately with `NameError`, so nothing shipped
broken -- but marker-based splicing is only safe when you have checked what
currently sits between the markers. Restored it with the round-7 signature.

### Notes

- **Default view is weeks**, but a module page opens in module view, since
  that is what the reader is looking at; otherwise the last choice is
  remembered per browser. Flagged to Nico in case he wants it to always
  reset to Week.
- The jump menu's options are **re-emitted from a snapshot** on each mode
  change rather than toggling `hidden` on `<option>` / `<optgroup>`, which is
  not honored on every platform.
- Verified after JS runs, not just in the source: a week page's menu offers 13
  entries (General + 12 weeks) and no modules; a module page offers 9 and no
  weeks, with the button reading "Module".

---

## 2026-09-03 (eighth pass) - dark red in the calendar, hamburger, 3-day flag

### What changed

| Asked for | Done |
|---|---|
| The same dark red for problem-set windows in the CALENDAR | `_build_calendar.py`: `DARKRED = "C00000"` border and `DUEWASH = "FBEDED"` fill on a problem set's due card. Only 5 of the 6 due cards get it - the practice final keeps gold, as on the website |
| The jump menu on the top LEFT with a hamburger symbol, navy ground | `.jumpwrap` / `.jumpicon`: three bars drawn in CSS with the real `<select>` over them at `opacity:0` |
| "Prof. Nico Voigtlaender - UCLA Anderson" centred in the top bar | `.topbar-in` is now a `1fr auto 1fr` grid with `.tb-left` / `.tb-sub` / `.tb-right` |
| Deadlines within 3 days in the same dark red, plus a legend | `data-date` on every row + `initDue()` in the browser; `.dl li.soon`; `.dl-legend` with a `.sw-due` swatch above the list |

### Word has no alpha, so the wash is composited by hand

The website uses `rgba(192,0,0,.07)`. Word fills are opaque, so `DUEWASH`
is that same colour flattened onto white: `0.07 x 192 + 0.93 x 255 = 251`
for red, `0.07 x 0 + 0.93 x 255 = 237` for green and blue - `FBEDED`. If the
website's alpha ever changes, recompute rather than eyeballing a new hex.

### The 3-day rule has to run in the browser

Every other thing on this site is baked in at build time, but "due within the
next three days" moves every day, and the pages are static. So each row
carries `data-date="YYYY-MM-DD"` and `initDue()` compares it to today's date
client-side, adding `.soon`. Rows with no date (the Practice Final Exam,
which is "t.b.a.") carry an empty attribute and are skipped.

**Tested with a stubbed clock**, not by reasoning: a probe copy of week-03
with `window.Date` replaced by a fixed 11 October 2026 (two days before
Problem Set 1 is due) flagged exactly one row - Problem Set 1 - and left the
past week-1 video row, the four later problem sets, both exams and the undated
practice final alone. The stub has to pass constructor arguments through to the
real `Date`, or `new Date(y, m, d)` inside `initDue()` returns the fixed date
too and every row matches.

### A note on the legend

It is rendered unconditionally, directly under the "Deadlines & Exams" header,
rather than only when something is actually within three days. It costs one
line and it explains the colour coding whether or not anything is red today.

### Calendar rebuilt

Backups rolled, `.docx` rebuilt, `.pdf` re-exported. `_check_pagination.ps1`
**PASSES** - 14 pages, one page per week. Colour changes do not move any card,
so this was expected. Verified 5 `C00000` borders and 5 `FBEDED` fills in the
document XML.

---

## 2026-09-03 (seventh pass) - problem-set styling, jump menu, phone nav

### What changed

| Asked for | Done |
|---|---|
| Problem-set boxes: dark-red rule, very transparent dark-red shading, larger due date | `.pcard.due.pset` -- `--due-edge #C00000` and `--due-bg rgba(192,0,0,.07)`; the date is 17px semibold. Dark theme uses a brighter edge (`#E0655C`) and a heavier wash (.17) so it reads on the dark ground |
| Faculty-page link on the instructor's name in the header, new tab | `NICO_URL`; `target="_blank" rel="noopener"` |
| General Logistics badge: dark gray, not dark yellow (gold is for exams) | `--gl-badge-bg #555B66` with white text |
| A jump dropdown in the top right corner, and the By week / By module box becomes a dropdown on a phone | native `<select id="jump">` with three optgroups (General / Weeks / Modules), 21 options, the current page marked `selected`. On a phone the toggle and both lists are hidden -- the jump menu IS the navigator |

### Only PROBLEM SETS get the red treatment

Nico said "for Problem Set boxes". Week 11's card is the Practice Final Exam,
which is not a problem set, so it keeps the gold rule and the smaller date.
The gate is a single `is_pset` check on the label, shared with the "upload one
solution" line, so the three changes can never drift apart.

### A bug worth remembering

**Anything on the navy bar needs its own link color.** The new faculty link
inherited `a{color:var(--ink)}` -- navy on navy -- and rendered as a blank
underlined gap. Caught in the screenshot, fixed with a `.tb-sub a` rule. This
is the same class of mistake as declaring a color only inside a dark-theme
block: the surface changed, the color did not follow.

### Why a native <select> for the jump menu

Phones render `<select>` with their own full-screen picker, which is far better
than any custom dropdown at that size, and it needs four lines of JavaScript.
It also means the phone navigator costs no vertical space at all - which is
what let the 13-item list be hidden outright below 860px.

### Verified

- Jump menu: 21 options in 3 optgroups, current page `selected`, on every page.
- `.pset` present on weeks 3, 5, 7, 9, 10 and absent on week 11.
- Faculty link with `target="_blank"`; byline now legible on the navy bar.
- Screenshots at 1500px and at 512px (the narrowest headless Chrome allows).
- Live site re-checked by URL after the Pages build.

---

## 2026-09-03 (sixth pass) - GL layout, syllabus/BruinLearn boxes, phones

### What changed

| Asked for | Done |
|---|---|
| "Class and Contact" to the top right corner - and is the order determinable? | Yes, fully: the page is built from explicit column stacks. Class and Contact now opens the RIGHT column |
| Class Syllabus as its own half-width box, with a matching box to its right linking the BruinLearn class site | new `.gl-row`: two `panel()`s side by side, above the columns. The cream strip is gone |
| "Upload one solution per group on BruinLearn" on all problem sets | added to the due card, gated on the label starting with "Problem Set", so the Practice Final Exam is untouched. Link: `BRUINLEARN_COURSE` |
| Is the site phone compatible? | It was mostly, with one real bug - see below |

### The General Logistics grid went from three columns to two

Putting Class and Contact at the top right left column 1 with a single box and
column 3 with three, which looked lopsided. Two columns of three satisfies
every constraint Nico has given AND balances:

| Left | Right |
|---|---|
| How the Quarter Runs | Class and Contact |
| Watching the Videos | Math Refresher |
| Online Practice Exercises | Textbook |

### Phone compatibility - one real bug, now fixed

- **`.shell` carries `align-items:start`.** Under the phone breakpoint the
  shell switches to `display:flex; flex-direction:column`, and `align-items:
  start` then stops the flex children from filling the width - the sidebar
  cards rendered at about 207px against full-width content. Fixed by resetting
  to `align-items:stretch` in the phone block. **Remember this whenever a grid
  container is switched to flex.**
- **`.box-hd` needed `flex-wrap:wrap`** so the long class-times-and-room header
  line can wrap rather than overflow.
- **Content now comes first on a phone** (`order` on main / left / right), so a
  student does not scroll past the 13-item navigator to reach the week the
  site just landed them on.
- The top-bar byline is hidden below 700px; the footer repeats it.

**Measurement limit, stated plainly:** headless Chrome on Windows will not
size a window below about 512 CSS pixels - `--window-size=360` and `390` both
reported `innerWidth=512`, in both old and new headless. So "no horizontal
overflow" is verified down to 512px only, which exercises the 860, 820, 700
and 640px rules but NOT the sub-420px block. That block only changes font size
and padding. Real-device confirmation is a 10-second check now that the site
is live.

### Notes

- **"BruinLearn" spelling.** Nico wrote "Bruinlearn"; the official UCLA name is
  "BruinLearn", which is what the new text uses. The footer still carries the
  calendar's own wording, "Bruin Learn" (two words), from
  `C.BRUINLEARN_NOTE`. Worth unifying if he cares.
- `BRUINLEARN_COURSE` lives in `_build_site.py`, not in the calendar's `LINKS`
  registry, because only the website links it.

---

## 2026-09-03 (fifth pass) - the last website-only omission removed

Nico asked for "Textbook reading: Math review Appendix Section 1 + Section 2
(only first derivatives)" to be deleted from the calendar as well.

- Removed from `MATH_REFRESHER_ITEMS` in `../Course Calendar/_calendar_content.py`
  (now 2 items: the Math Quiz and the Math Review Videos).
- **`DROP_TEXTBOOK_NOTES` and `DROP_MATH_ITEMS` are now both empty.** The hooks
  stay in `_build_site.py`, commented, in case a website-only omission is ever
  wanted again - but the calendar and the website now carry exactly the same
  content, with no divergence to remember.
- Calendar rebuilt, PDF re-exported, backups rolled. `_check_pagination.ps1`
  **PASSES** - 14 pages, one page per week.
- Website rebuilt and deployed. The line was already hidden on the site, so the
  only visible change is in the calendar; the website change was clearing the
  dead configuration.

### Still open

- **The Dropbox podcast links remain readable in the public repository.**
- The syllabus link is still `#`.

---

## 2026-09-03 (fourth pass) - American English, symbols, GL layout

Seven more edits from Nico. All live.

### What changed

| Asked for | Done |
|---|---|
| American English as the default, and added to the universal CLAUDE.md | new first bullet under "Style & Tone" in the OneDrive master `CLAUDE-Universal.md`; swept the website (colour->color, centre->center, organise->organize, grey->gray, judgement->judgment). The CSS class `.centre` and the function `band_centre()` were renamed too |
| New "How the Quarter Runs" wording | verbatim |
| Delete "Panopto - all class videos"; link the sign-in screenshot instead | `_build_site.py` copies `../Course Calendar/Images/Panopto-Login-Picture.png` to `assets/panopto-login.png` on every build; the box links it as "How to sign in (screenshot)". `_deploy.py` ships the file |
| Delete the "Textbook reading: Math review Appendix" item; move the Math box right, above Textbook; give it a math symbol | `DROP_MATH_ITEMS`; the General Logistics grid is now THREE explicit column stacks; the symbol is the partial-derivative sign |
| Videos symbol: the Hollywood "action" symbol | clapperboard U+1F3AC, on the website AND in the calendar. Video bullets still carry the play triangle |
| "Light Mode" / "Dark Mode" with a hover text box | button holds a label span and a `.tip` span; CSS shows the tip on hover/focus; `aria-label` carries the same text for screen readers |
| Delete the review-questions note and the Achieve item from the CALENDAR | removed from `_calendar_content.py` outright; the website's `DROP_TEXTBOOK_NOTES` is now empty |

### The General Logistics page is now three fixed columns

| Column 1 | Column 2 | Column 3 |
|---|---|---|
| How the Quarter Runs | Watching the Videos | Math Refresher |
| Class and Contact | Online Practice Exercises | Textbook |

Each column is ONE grid cell (`.panel-stack`), so no pair can be split apart
by the auto-fit grid at any window width. This is what "move the Math box to
the right, above Textbook" needed - the earlier auto-flow grid could not
guarantee it.

### Notes

- **The derivative symbol takes the pale-gold color, unlike the emoji.**
  U+2202 is ordinary text, so CSS `color` applies. The clapperboard,
  headphones, book and pencil are color emoji and keep their own colors -
  same as in the calendar.
- **The math-review-appendix line is website-only.** Nico named only two
  lines for deletion from the calendar, so that third one still prints there.
- The `WORDS` sweep deliberately used ``-anchored prefixes, so
  "organised" -> "organized" and "capitalised" -> "capitalized" were caught
  along with their base forms.

### Calendar rebuilt again

`_calendar_content.py` (two deletions) and `_build_calendar.py` (clapperboard).
Rolled `_t-1`/`_t-2`, rebuilt the `.docx`, re-exported the `.pdf` through Word
COM. **`_check_pagination.ps1` PASSES** - 14 pages, every week on one page.
Confirmed in the document XML that the clapperboard is present, the film reel
is gone, and neither deleted line survives.

### Verified

- All 21 pages rebuilt; no British spelling left in any generated page.
- Clapperboard on the Videos card, play triangle still on all 4 week-1 video
  bullets, derivative symbol on the Math box.
- Live site re-checked by URL after the Pages build finished (it took about a
  minute longer this round): new "How the Quarter Runs" text, "Light Mode",
  "Switch to Dark Mode", "How to sign in (screenshot)", "Color Coding", and
  `assets/panopto-login.png` returning HTTP 200.

### Still open

- **The Dropbox podcast links remain readable in the public repository.**
- The syllabus link is still `#`.

---

## 2026-09-03 (third pass) – Nico's revision round on the live site

Worked through a 13-point list of edits. All are live.

### What changed

| Asked for | Done |
|---|---|
| Theme button: no "Auto · light"; show "Dark", hover says "Switch to Light" | `syncBtn()` names the CURRENT theme, `title` offers the other; the click is now a plain light/dark toggle. Sunrise/sunset still decides the initial state |
| No "General Logistics" breadcrumb top-left; move the agenda up | `.subbar` and `.crumb` deleted; the left column starts right under the top bar |
| Search as narrow as the Deadlines column, header space in the middle | the search box moved INTO the right column, above the deadlines card, so the band sits level with it |
| "Back to ..." inside the header; the right-hand text to the center, no dates | `.band` is now a `1fr auto 1fr` grid: identity left, kind centerd, back link right |
| Center text wording | "Video Content & On-Campus Class" / "Video Content" / "Midterm Exam" / "Final Exam" |
| Syllabus: a link instead of the Bruin Learn sentence | cream strip, "Download the Class Syllabus here", `SYLLABUS_URL = "#"` |
| "How the Quarter Runs" + his replacement text | verbatim |
| Land on the current week; General Logistics before the quarter starts | inline redirect in `index.html`; the sidebar link carries `?stay=1` |
| Drop the textbook "review questions" note | `DROP_TEXTBOOK_NOTES` — website only |
| Videos box before Textbook; "Online Practice Exercises" box right under it | `.panel-stack` puts the two in ONE grid cell, so practice is always directly below videos |
| Drop the "Achieve" math item | `DROP_MATH_ITEMS` — website only |
| Title Case + navy background on all header boxes | `title_case()` + `box_hd()`; applied to panels, Topics Covered, the two sidebar cards, the class / videos / holiday cards |
| No per-week video totals | the "≈ N min of video in total" line is gone |
| Deadlines list: pre-class videos per on-campus week | "Watch Videos 1 – 4 for Module 1", "by 9/25" |

### Decisions and judgment calls

- **Title Case is capitalize-only.** `title_case()` never lowers a capital, so
  acronyms survive. It is applied ONLY to headings the site authors — never to
  calendar content (item text, group labels, exam titles, topics).
- **Weeks 5 and 9 read "On-Campus Class", not "Video Content & On-Campus
  Class"**, because those weeks assign no new videos of their own (their prep
  is podcasts, reading, articles and a quiz). Nico's rule assumed every
  on-campus week has videos; week 1 is the only one that does.
- **Week 11 reads "Exam Preparation"** — it has neither videos nor a class, so
  neither of the two given labels fits.
- **Only week 1 produces a "watch the videos" deadline row**, for the same
  reason. Video weeks have a suggested Sunday deadline shown on the week page;
  it was NOT added to the deadlines list, since Nico asked for on-campus weeks.
- **The two deletions are website-only.** The lines are still in
  `_calendar_content.py` and still print in the calendar; `DROP_*` lists in the
  generator hide them. Nico was asked whether he wants them gone from the
  calendar too.
- **`box_hd()` gained a `when` slot** rather than the string-replacement hack
  the first attempt used to inject the room/date line into a navy header.

### Verified

- All 21 pages rebuilt; band center text, back link and next link checked on
  every one of the 12 week pages and on modules 1 and 8.
- **The landing redirect was tested with a stubbed system clock**: real date
  (Sep 3, before the quarter) stays on General Logistics; a date inside week 3
  lands on `week-03.html`; `?stay=1` always stays.
- Light and dark both screenshotted after the CSS rewrite.
- Live site re-checked by URL: syllabus link, "How the Quarter Runs", "Online
  Practice Exercises" present; the Achieve item and the video totals absent;
  the week-1 watch row present.

### Still open

- **The Dropbox podcast links remain readable in the public repository.**
- The syllabus link is `#` until Nico supplies the address.
- Whether the two dropped lines should also go from the calendar.

---

## 2026-09-03 (second pass) – design A adopted, rebuilt as a multi-page site

Nico picked design A ("Ledger") and asked for a long list of changes. The
single-page prototype was replaced by a Python static-site generator, and the
site went live on GitHub Pages.

### What changed, against his list

| Asked for | Done |
|---|---|
| "Before you Start" as its own tab, first in both views, labelled "General Logistics / Before you Start" | `index.html`; first entry of the week list AND the module list |
| A separate page per week, reached from the left | `week-01.html` … `week-12.html` |
| "Week t−1" top right, "Week t+1" bottom right; week 1 top says "Back to General Logistics"; last week has no bottom link | `turn()` in the generator |
| Same for modules | `module-1.html` … `module-8.html`, module 1 points back to General Logistics |
| Category colors: videos yellow, podcasts light gray, reading white; navy header box per category with the group label below | `.cat` cards; podcast gray is the new `PODGRAY #EFF1F4` |
| Calendar's category symbols, with the Videos triangle → film reel | 🎞 / 🎧 / 📖 / ✎ in the card headers; video bullets keep ▶ |
| On-campus box on top in on-campus weeks | class card emitted before the prep container |
| Video length behind the link, not far right | `(9 min)` immediately after the link text |
| Auto dark after sunset / before sunrise, keep the button | SunCalc in `assets/site.js`; button cycles Auto → Light → Dark |
| Search box top right, under the navy bar | `.subbar`; `/` focuses it |
| Color-scheme legend where the search box used to be, kept small | left column, top, three swatches |
| A right sidebar listing deadlines per week | `.right`, all problem sets + exams in date order, current week highlighted |
| Site address and name | see below |

### The same three changes were made in the CALENDAR

Done in this session, in `../Course Calendar/_build_calendar.py` (Nico asked
whether it needed a separate session — it did not):

1. **Podcast cards now light gray** (`PODGRAY = "EFF1F4"`) instead of white.
2. **Videos card glyph is a film reel** (`\U0001F39E`) instead of `▶`. The
   calendar's video bullets never used a triangle, so nothing else changed.
3. **On an on-campus week the class card comes above the prep container.** The
   weekend block was hoisted into a local `render_weekend()` called either
   before or after the prep section, keyed on `weekend_first`.

Rebuilt the `.docx`, re-exported the `.pdf` through Word COM, rolled the
`_t-1` backups first. **`_check_pagination.ps1` still PASSES** — every week on
one page, 14 pages total. Verified in the document XML that the class-card
header ("4:00 – 5:30 pm") now precedes the prep heading in weeks 1, 5 and 9,
that week 2 is untouched, and that there are 8 `EFF1F4` fills.

### Address and naming

- **URL:** https://nvoigtla.github.io/MGMT-405-EMBA/ — GitHub repository names
  cannot contain spaces, so the requested "MGMT 405 EMBA" became
  `MGMT-405-EMBA`. A custom domain is the only way to get the spaces.
- **Site name / top-bar tag:** `Managerial Economics Fall 2026 – EMBA`, set
  once as `SITE_NAME` in `_build_site.py`. Page titles read
  "Week 3 · Managerial Economics Fall 2026 – EMBA".

### Decisions made

- **Static generation in Python, not client-side rendering.** Multi-page with
  prev/next navigation wants real HTML: it is fast, printable, works off disk,
  and needs no JavaScript for content. JS is only theme, search and the nav
  toggle.
- **The public repo holds the BUILT OUTPUT ONLY.** `_deploy.py` clones the
  public repo into a temp folder, syncs the 21 pages plus the 3 assets, commits
  and pushes, then enables Pages. No nested git repo inside Teaching, and the
  build script, README and these notes stay private.
- **`noindex, nofollow`** on every page, so the course pages are not indexed.
- **Sunrise/sunset needs a position**, and asking students for the geolocation
  permission would be intrusive, so ~50 IANA time zones are mapped to
  approximate coordinates, with a UTC-offset fallback.
- **Module-page topic lines are de-duplicated** against the cards; week 5 was
  printing "Module 3: Applications" twice.

### Open items / flagged to Nico

- **The Dropbox podcast links are now in a public repository.** Dropbox share
  links need no login, so anyone who finds the repo can fetch the audio.
  Panopto needs an Anderson sign-in and the TA site is already public. Remedies
  if he minds: move the audio behind Bruin Learn, or host the site privately.
- **"On top" was read as "above the preparation container", not "the very first
  card".** The order is band → due → Topics covered → class card → prep. If he
  meant above "Topics covered" too, it is a two-line change in `week_main()`.
- **`(++)` is the calendar's own convention** for an unmeasured running time
  and was kept for fidelity, with a tooltip. `(link to follow)` was ADDED for
  videos with no URL, because an unlinked title looks broken on a web page.
- **Emoji glyphs render in their own colors**, not the pale gold the CSS asks
  for, because color emoji ignore `color`. Same as the calendar. If the navy
  header looks muddy, the fix is a monochrome glyph set.
- Content gaps the site now shows honestly: Module 3 video 1 and all five
  Module 4 videos read "(link to follow)"; Modules 5 – 8 podcasts likewise;
  Module 2 and 3 video lengths read "(++)".

### Verified this session

- All 21 pages render in headless Chrome with no JS errors and no
  `undefined`/`NaN`; light and dark both checked by screenshot.
- Live site returns HTTP 200 for `index.html`, `week-01.html`, `module-3.html`
  and all three assets.
- Calendar pagination invariant still passes after the reorder.

### Commands

```
python _build_site.py          # regenerate all pages + search index
python _deploy.py --dry-run    # what would be published
python _deploy.py              # push to the public repo, enable Pages
```

Headless render check (Chrome is installed; no LibreOffice on these machines):

```
"/c/Program Files/Google/Chrome/Application/chrome.exe" --headless --disable-gpu \
  --allow-file-access-from-files --virtual-time-budget=4000 --hide-scrollbars \
  --window-size=1500,1600 --screenshot=out.png "file:///<abs path>/week-01.html"
```

To check dark mode, copy a page with `data-theme="dark"` on `<html>` and the
`site.js` tag removed, screenshot it, then delete the copy.

---

## 2026-09-03 (first pass) – three design alternatives

Built the data pipeline and three complete alternatives — A "Ledger" (fixed
sidebar + collapsible week rows), B "Term Rail" (editorial timeline), C
"Console" (two-pane app) — each reading the same generated content, in the
decks' palette, with a week/module toggle. Nico chose **A**.

Superseded and deleted in the second pass: `design-a/b/c.html`,
`_build_site_data.py`, `_bundle_preview.py`, `assets/course-lib.js`, `data/`,
`_preview/`. The `_preview/` entry in the repo-root `.gitignore` is now
unused but harmless.

Preview artifacts from that pass (they show the OLD single-page prototypes,
not the live site): A `fdcc7a40-5837-4008-a9e2-7637144b9c6a`,
B `a5e2c02e-f20b-41ff-ba5f-c3be0517f66f`,
C `abd1ef6c-fdb7-4260-a5b0-40f96062ed0e` under claude.ai/code/artifact/.
