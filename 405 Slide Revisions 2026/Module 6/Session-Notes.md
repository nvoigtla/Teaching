# Module 6 – Session Notes

## 2026-09-09 — Module 6 rebuild started: sources inventoried, outline drafted

**One-line summary.** Inventoried my Module 6 deck, the on-campus
applications deck and all three PG decks, then drafted
`Module 6 - Revised - outline.md` for approval. No deck built yet.

### Files created (all build inputs — keep)

| File | What it is |
|---|---|
| `Module 6 - Revised - outline.md` | **The deliverable awaiting review.** Slide-by-slide plan, PG adoption list, 9 open decisions |
| `_source_inventory_NV.md` | `Module 6.pptx` — 76 slides, 4:3, text + notes + media per slide |
| `_source_inventory_NV_apps.md` | `Module 6 -- Slides On-Campus Applications with Solutions.pptx` — 52 slides |
| `_source_inventory_PG1/2/3.md` | PG's three decks — 43 / 30 / 31 slides, already 13.33 × 7.5" |
| `_assets_manifest_NV.md`, `_assets_manifest_PG.md` | Image manifests |
| `_source_images/` | 52 NV + 30 NVapp + 65 PG images extracted |
| `_extract_pg_media.py` | Pulls PG's media — she builds screenshots as picture-**filled** autoshapes, which python-pptx does not report as pictures, so this walks the raw OOXML |
| `_pick_slides.py` | Prints selected slides from an inventory with the chrome lines filtered out |
| `_dump_deck.py`, `_extract_assets.py`, `_deck_guard.py`, `_diff_slides.py` | Copied from `Module 4/` |

### Decisions taken this session

1. **Built as a taped module in one pass** — 8 video title cards,
   `Module 6 · Video k · <topic>` tags, gold coverage pills. Video names come
   verbatim from the Fall 2026 course calendar, so pills / cards / calendar
   cannot drift.
2. **The applications deck feeds a full `In Class · Examples` appendix** at
   the end of the revised deck.
3. **BMW: the graphical comparison only in the main deck**; the algebra goes
   to the practice-video deck.
4. **A second deliverable**: `Module 6 - Practice Video - Optimal Pricing in
   Two Markets.pptx`, 10 slides, the complete PG BMW block, for retaping the
   19-minute practice video.
5. **Auctions are not Module 6 material.** PG's Part 3 (31 slides) is a full
   auctions unit, but the calendar teaches auctions in **Module 8**. Parked
   with its inventory and images intact — see section 7 of the outline.
6. **9-row outline** (four degrees / strategies at top level, the three
   advanced strategies as sub-items) so there is one row per video. Needs a
   0.615" row pitch, just inside the builder's 0.60" floor.

### Things I found in the sources

- **Two slides in my main deck are title-only**: 52 ("Two-part tariff
  Summary") and 56 ("Zoo Pricing – Solution"). The applications deck has the
  full zoo solution; PG2 18 can fill the two-part-tariff summary.
- **Three empty picture placeholders**: slide 59 (Costco), 71 (The Economist
  options), 76. PG2 7's Costco membership-tier image fills 59 exactly.
- **Slides 69–70 are exact duplicates of 66–67** (The Mad Optimist).
  Superseded 2026-09-09: The Mad Optimist is KEPT; only the verbatim second
  copy is not repeated.
- **The zoo numbers disagree between my two decks.** Main 57: F = $50,
  profit $50, MC = 0. Applications 41: F = $50, variable cost $20, profit
  $30, MC = $4. The applications version is consistent with its own two-part
  tariff solution (F = $32, P = $4), so I propose that one.
- **Slide 2's exam dates are the March 2022 finals** — need refreshing from
  the Fall 2026 calendar.
- **Backup slide 76** is the *Airplane Seating Chart* cartoon
  (TheCooperReview). It backs the versioning slide, which already says "More
  seating versions…".
- **PG's arithmetic in the BMW block checks out**: US Q\* = 15, P\* = 100,
  E_D = −1.667, Lerner 0.6; Germany Q\* = 10, P\* = 90, E_D = −1.8, Lerner
  0.556 — and Lerner = 1/|E_D| holds in both.
- **PG's decks are already in our chrome** (13.33 × 7.5", navy bar at
  0/0/13.33/0.42, title at 0.28/0.50, rule at 1.25, gold strip 2.20 × 0.05,
  footer rule at 7.15, numbered-circle outline). Adoption is a content
  question, not a formatting one.
- **The NV source decks were duplicated** at the folder root and in
  `Module 6 - NV Slides/`. Superseded 2026-09-09: the root set was deleted
  with Nico's OK; the subfolder copies are the build inputs.

### Open — waiting on Nico

Superseded by the round-2 section below (decisions A–H).

### Not yet looked at, by instruction

`Module 6 - Video 1.pptx` … `Video 8.pptx` — superseded later the same day:
Nico asked me to read them to verify the video boundaries. See the round-2
section below.

### Commands worth remembering

```
python _dump_deck.py "<deck>.pptx" _source_inventory_X.md
python _extract_assets.py "<deck>.pptx" _source_images TAG
python _extract_pg_media.py                      # PG picture-filled shapes
PYTHONIOENCODING=utf-8 python _pick_slides.py _source_inventory_PG1.md 31-40
```

`PYTHONIOENCODING=utf-8` is needed on this machine — the inventories contain
`⏎` and cp1252 stdout chokes on it.

## 2026-09-09 (later) — video decks read, outline rewritten as v2

**One-line summary.** Read the 8 original video decks, matched every video
slide to a main-deck slide, and rewrote the outline: v1 had put in-class
material inside the video blocks.

### The video map (evidence, not guesswork)

`_map_videos.py` fingerprints each slide by its whole text and scores it
against every main-deck slide. All 44 video content slides matched at
**>= 0.71**, most at 1.00. Result:

| Video | Main-deck slides, in tape order |
|---|---|
| V1 Simple vs. Complex | 4, 5 (BOTH outline slides), 6, 7, 8, 9, 10 |
| V2 First Degree | 11, 12, 14, 15, **13** (dystopian last) |
| V3 Segment Pricing | 17, 18, 19, 20, 21, 22, 23, **29**, 30 |
| V4 Versioning + Coupons | 31, 32, 33, 34, **59 (Costco)**, 37, 42 |
| V5 Flat Fee | 43, 44, 45, 46, 47, 49 |
| V6 Two-Part Tariffs | 50, 51, 52, 53 |
| V7 Block Pricing | 60, 61, 62 |
| V8 Summary | 64 only -- no agenda slide |

**In NO video** (= in-class): 2, 3, 16, 24-28, 35, 36, 38-41, 48, 54-58, 63,
65-73, 74, 75-76.

Three findings that explain oddities in the main deck:
- Slide 52 ("Two-part tariff Summary") is title-only **in the tape too** --
  a blank canvas Nico talks over, not an unfinished slide.
- Slide 56 (zoo solution) is title-only because it is worked live in class;
  the applications deck carries the finished version.
- The dystopian slide exists in two versions **by design**: the video
  *states* it ("Example: Uber"), the class version *asks* it.

### Decisions taken (round 2)

- **Video blocks are the retaping script** -- PG adoptions and enrichments go
  INSIDE the video blocks, so the deck is what Nico would tape next time.
  Section 5 of the outline lists exactly what now differs from the 2022 tapes.
- **Costco appears twice**: "Versioning: Costco" in Video 4 as taped, and a
  second slide in the in-class block with the 2024 title ("Combination of
  Versioning and Two-Part Pricing").
- **The in-class block keeps the applications deck's recap slides**, so it
  runs standalone. Block = applications deck slides 7-52 with the PG cases
  spliced in.
- **All 9 outline rows stay; no introduction section** (Video 1 is already a
  teaching video).
- **The Mad Optimist is kept** -- only the verbatim second copy at NV 69-70 is
  not repeated.
- **Zoo: the applications-deck numbers** (F = $32, P = $4/visit, MC = $4).
- **Root duplicate source decks deleted** (all ten sha1-identical to the
  `Module 6 - NV Slides/` copies, verified before deleting).

### Exam facts for slide 2, from the Fall 2026 calendar

Anchor Friday 2026-09-25. Final exam window **Fri Dec 11 - Sun Dec 13, 2026**;
**one 3.5-hour window** for everyone (NOT the old 5 hours / two TA windows);
online, open book, proctored; ~20 MCQs + 3-4 problems; Modules 1-8.
Module 6 applications class = **Sat Nov 21, 2026**. PS3 due Tue Nov 10,
PS4 Tue Nov 24, PS5 Thu Dec 3. **"Achieve" is gone from the course** -- the
practice material is the BruinLearn online quizzes.

### Open -- waiting on Nico

Decisions A-H in section 8 of the outline: two-part-tariff summary content,
the Video 8 agenda slide, the empty Barbie/dolls price labels, The Economist
options table, retiring the old practice-video example, which problem set the
pointer boxes name, whether in-class repeats get a "Recap:" prefix, and
whether to split the 122-slide deck into video + in-class files.

### New build inputs

`_map_videos.py` (video-to-main-deck matcher), `_source_inventory_V1..8.md`.


## 2026-09-09 (build) -- video half of the deck built and verified

**One-line summary.** `_build_Module6.py` now produces the front matter, all
eight video blocks and the summary closer -- **59 slides, opening in
PowerPoint, every slide checked on a PNG render.** The In-Class Examples
block and the practice-video deck are still to come.

### What is built

| Slides | Block |
|---|---|
| 1-3 | Title, Some Logistics (dates refreshed), Course Roadmap |
| 4-12 | Video 1 - Simple vs. Complex Pricing (card, overview + section agenda, 6 content) |
| 13-18 | Video 2 - First-Degree Price Discrimination |
| 19-30 | Video 3 - Segment Pricing (incl. Orbitz and the BMW three-panel figure) |
| 31-38 | Video 4 - Versioning and Coupons (incl. Costco) |
| 39-46 | Video 5 - Flat Fee Pricing (incl. ClassPass) |
| 47-51 | Video 6 - Two-Part Tariffs (summary slide filled) |
| 52-55 | Video 7 - Block Pricing |
| 56-58 | Video 8 - Summary of Pricing Strategies (agenda slide added) |
| 59 | Module 6: Summary (closer) |

### Still to build

- **In-Class Examples block** (~60 slides): applications-deck slides 7-52
  with the recaps kept, plus the Disneyland case, the two "In the News"
  openers and the second Costco slide.
- **Backup**: divider + the airline seating cartoon with its back button.
- `_splice_media.py` for the three poll trios, `_group_pass.py`,
  `_animate.py`.
- **`Module 6 - Practice Video - Optimal Pricing in Two Markets.pptx`**
  (10 slides, the full BMW block).

### Findings worth keeping

- **`build_freeform().convert_to_shape()` CORRUPTS the deck.** python-pptx
  saves it happily and `zipfile.testzip()` passes, but PowerPoint then
  refuses to open the file at all. Use the custGeom-into-a-RECTANGLE
  pattern instead -- now `_fig_poly` in `_m6_helpers.py`, lifted from
  Module 4.
- **`head_size` on `_add_arrow` must be `sm` / `med` / `lg`.** Passing
  `'small'` writes an invalid OOXML enum and corrupts the deck the same
  silent way. Both bugs were found by `_probe_build.py`, which builds one
  deck per slide function so PowerPoint can name the guilty slide.
- **PG's "Personalized Pricing in the Wild" is Nico's own material.** Her
  headline and photo are BYTE-IDENTICAL to slide 13's two images (sha1
  6471d638 and 4f2ec7c5), so adoption A5 was withdrawn -- there was nothing
  to adopt. The images are used on the dystopian slide, which is where they
  already were.
- **The 9-row outline does not survive measurement.** Nine rows floor the
  pitch at 0.60", leaving a 0.28" title row that a 25 pt title overflows.
  Eight rows fit at 0.709" with 24 pt titles, so the "Advanced Pricing
  Strategies" umbrella row was dropped and its three strategies became
  top-level items keeping "Advanced Pricing:" in their titles. Every
  section is still in the agenda and nothing is collapsed.
- **The Barbie prices were in the source images all along** -- Mattel
  screenshots reading $9.99, $150.00 and $13.99. Nothing had to be invented
  and the three empty text boxes are dropped.
- **BMW arithmetic is derived, not typed.** `_bmw_opt` solves each market,
  and the joint demand's kink at (5, 140) is computed from the horizontal
  sum. Segment profit 1,400 vs uniform 1,389, with an `assert` in the build
  script so the slide's claim cannot silently go false.

### Verification routine used

```
python _build_Module6.py                 # rebuild
python _probe_build.py [name ...]        # one probe deck per slide fn
# then, from PowerShell:
#   open the deck via COM -- python-pptx is too lenient to catch corruption
#   export slides to _probe\*.png and read them as a contact sheet
```


## 2026-09-09 (correction) -- animation choreography, and the AlternateContent blind spot

**Trigger.** Nico: "on the Barbie slide keep this as was, including the three
empty text boxes initially covering the prices and then being animated out"
and "make sure you keep exactly the order of animations from my original
slides."

### The Barbie boxes are MASKS, not empty labels

Three `bg1` (white) 1.25 x 0.39" boxes sit exactly over the three Mattel price
screenshots and are removed by EXIT fades, one per click. `_extract_anim.py`
gives the original order:

| click | reveals |
|---|---|
| 1 | $9.99 -- Barbie Farmer Doll (left) |
| 2 | $13.99 -- Barbie Robotics Engineer Doll (right) |
| 3 | $150.00 -- Barbie Yves Saint Laurent Doll (middle) |

The designer doll is the punchline, revealed last. My earlier plan to "drop
the three empty text boxes" would have destroyed the device.

### New: `_extract_anim.py` -- Nico's original choreography, click by click

Reads `<p:timing>` from any source deck, walks the main sequence, and
resolves every effect's `spid` to a shape SIGNATURE (kind, rendered
position in inches, text) because ids do not survive a rebuild. Output is
captured in `_anim_original_NV.md`, `_anim_original_APP.md`,
`_anim_original_V.md` -- **48 of the 76 main-deck slides carry
choreography.** The animation pass must reproduce these beat for beat, NOT
generate from the generic guidelines.

Gotcha found while writing it: `nodeType="mainSeq"` sits on the `<p:seq>`'s
CHILD `<p:cTn>`, not on the `<p:seq>` itself.

### The blind spot this exposed -- and two slides built on a false premise

Slide 52 showed FIVE animation clicks against a placeholder my inventory
said did not exist. Cause: **python-pptx does not enumerate a shape wrapped
in `mc:AlternateContent`**, which is what PowerPoint does to any textbox
holding OMML math. This is in my memory file and in Teaching CLAUDE.md, and
I still walked into it -- `_dump_deck.py` uses python-pptx.

New tool `_dump_text_raw.py` reads the raw OOXML instead. Hidden content
found: NV 49, 52, 53, 54, 57, 58 and APP 34, 35, 37, 39, 40, 41.

**Two slides had to be rebuilt:**
- **"What can go wrong with flat rates"** -- I had INVENTED two bullets.
  Nico's slide carries the AAirpass quote ("Each had paid American more than
  $350,000 for an unlimited AAirpass..."), his two questions, and his
  take-away. Restored verbatim; "only 65 customers" is the ANSWER to his
  question and belongs in the notes, where he had it.
- **"Two-part tariff: Summary"** -- I had filled it with PG2 18 (adoption
  A9) because it looked blank. His own six-paragraph summary was there all
  along and says the same things. **A9 WITHDRAWN.**

**Adoption A5 also withdrawn**: PG's "Personalized Pricing in the Wild"
headline and photo are byte-identical to my own slide 13's (sha1 6471d638
and 4f2ec7c5).

### New: `_check_coverage.py`

Compares every rebuilt slide, paragraph by paragraph, against the RAW source
XML, folding PowerPoint's math-italic codepoints back to ASCII. It found
seven further wording losses, all restored (NV 9, 14, 15, 46, 59, 62, 64).
It now reports 8 lines under 60%, every one deliberate: the five refreshed
exam lines, one retitled slide, and two OMML tokenization artifacts.

**Run it after every build.** Recovered numbers worth keeping: ZipCar demand
is `P = 1.5 - 0.01Q` (exactly the line I had derived), and the zoo is
`P = 20 - 4Q` with `MC = 0` in the main deck and `MC = $4` in the
applications deck -- the two versions differ only in MC, and Nico chose the
applications one.


## 2026-09-09 (build complete) -- both decks built, pipeline run end to end

### Deliverables

| File | State |
|---|---|
| `Module 6 - Revised.pptx` | **117 slides.** Front matter + 8 video blocks + summary closer + In-Class Examples block + backup. Opens in PowerPoint; full-screen slideshow probe PASSES with all three polls live. |
| `Module 6 - Practice Video - Optimal Pricing in Two Markets.pptx` | **10 slides.** The full PG BMW block; every number derived in code. |

### Pipeline (rerunnable, in this order)

```
python _build_Module6.py          -> 117 slides
python _splice_media.py           -> 3 polls, verbatim + notes + tags
python _group_pass.py             -> 13 groups
python _animate.py all apply      -> fade builds
python _build_M6_practice.py      -> the practice deck
python _check_coverage.py         -> must print only deliberate misses
```

Verification used throughout: open the deck via COM (python-pptx is too
lenient), export PNGs and read them as a contact sheet, and for the poll
slides run `_slideshow_probe.ps1`.

### The animation work

`_animate.py` gained EXIT-fade support, which it did not have. Nico's
"Versioning: dolls" reveals each price by fading OUT a white mask, and that
choreography could not have been reproduced otherwise. Emitted XML matches
what PowerPoint wrote on his slide 36 (presetClass="exit", animEffect
BEFORE the set, visibility "hidden" at delay 499 -- the reverse of the
entrance order). PowerPoint confirms slide 80: three effects, Exit=-1,
TriggerType=1, 0.5 s, MASK-1/2/3 in his order.

Skip sets are now DERIVED from the deck (no top bar -> title/card/divider;
titled "Outline of Module 6" -> agenda; a `tags` rel -> live poll), because
Module 4's hard-coded sets went stale as slides moved. 28 of 117 skipped.

**Module 4's PLANS were removed.** They are keyed by DISPLAY NUMBER, so
leaving them in would have applied Module 4's choreography to whatever
Module 6 slide shared the number. This was a live hazard, not a tidy-up.

### STILL OPEN -- the chart choreography

`_anim_original_NV.md` holds Nico's click-by-click choreography for **48 of
his 76 slides**. Only slide 80 has been translated into a plan. Every other
slide uses the DEFAULT plan: one top-level bullet per click, first bullet
visible. That MATCHES his text slides, but NOT his chart builds -- his
slide 62 has 15 clicks and slide 64 has 11, where the rebuild currently
puts the whole figure on one or two. The remaining job is to translate
those per-slide beat lists, chart by chart, and verify with
`_verify_anim.ps1`.

### Other things worth keeping

- `_make_poll_sidecar.ps1` carved `_handoff_polls_M6.pptx` (3 slides,
  2.5 MB) out of the 18 MB source deck with PowerPoint, so tags + notes +
  media travelled correctly. BUILD INPUT -- never delete, never round-trip
  through python-pptx.
- The AMC slide had been rebuilt with an INVENTED table of AMC Stubs tiers
  and prices, because the inventory could not read the .wmf clippings. His
  actual slide is a Forbes piece on sightline-based seat pricing -- nothing
  like it. His two .wmf images are now placed as they are. Lesson: an
  unreadable source is a reason to place the original, never to invent a
  replacement.
- The ice-cream block prices ($4.00 / $3.00 / $1.50) are NOT collinear,
  unlike the Walmart ones; an assert caught it. Demand is drawn piecewise
  through his three points so every block corner sits on the curve.


## 2026-09-09 (formatting pass) -- Nico's three comments, and a rules audit

### 1. The decision-tree slide, back to his original

His slide 64's colours come from shape STYLES pointing at theme accents,
which is why a plain `srgbClr` scan showed nothing. Resolved, his four
groups are:

| his accent | what it marks |
|---|---|
| accent2 `C0504D` | the four questions |
| accent6 `F79646` | "perfect competition" |
| accent3 `9BBB59` | "firm produces Q* where MR = MC" |
| accent1 `4F81BD` | advanced pricing, indirect, direct, and BOTH direct-PD terminals |

The blue family is **everything that IS complex pricing** -- that is the
point of the diagram, and my first version had flattened it to two colours.
Re-expressed as navy (questions) / gray / cream / gold, with his box
positions scaled from his 10" canvas (x times 1.333, y plus 0.43") and his
bottom-left direct-PD branch with the two terminals side by side under it.

### 2. Curve labels at the right end

Moved on slides 16, 17, 97 and 100. Slide 17 needed his own arrangement
from slide 15 -- callout high, demand label low near the axis -- and then a
further nudge left, because a label must not CROSS its own curve.

### 3. New: `_audit_format.py`

Checks a built deck against the rules that can be checked mechanically:
trailing periods, the 18 pt box-text floor, Title Case, off-canvas shapes,
live page-number fields, and the Back pill. **Run it after every build.**

The first run found **22 real problems behind ~300 false positives**; the
audit had to be tuned against documented geometry (the 13 pt coverage
pills, the footer band that sits 0.02" past the edge by design, the
narrow-object exception). What it caught:

- **Seven pictures / labels running off the canvas.** Wendy's was 7.15"
  tall on a 7.5" slide because I sized a PORTRAIT image by width. The PNG
  renders had not shown this -- the export crops to the canvas, so
  overflow is invisible in a render and only a geometry check finds it.
- **The Back pill was wrong**, exactly as Nico suspected: `_add_jump_pill`
  defaults to a WHITE pill with a navy label, and `back=True` does NOT
  change that. The colours have to be passed explicitly.
- Three trailing periods on bullets and labels.

Now down to 3 findings, all legitimate: two convention boxes carrying
running prose (which keeps normal punctuation) and one 15 pt corner pill
(the documented size).

### Gotcha worth remembering

The heredoc backslash trap bit FOUR times in this pass alone. Any patch or
note whose text contains a `u`-escape or a Python line continuation must be
written to a file with the Write tool. Doubling the escape in the patch is
also wrong: it yields a literal backslash-u that never matches the real
character. Build the backslash with `chr(92)` when a script must contain
one.

### State

Both decks rebuilt through the full pipeline and re-verified: 117 + 10
slides, deck opens, `_check_coverage.py` clean (8 deliberate misses),
`_audit_format.py` at 3 legitimate findings, and the full-screen slideshow
probe passes on the title, a chart, the tree, all three polls, the masked
dolls slide and the backup.

**Still open:** the chart choreography is translated for 11 slides
(8, 9, 16, 17, 22, 29, 34, 43, 49, 51, 80). The rest of the 48 slides that
carry original choreography still use the default one-bullet-per-click
plan. `_anim_original_NV.md` has the specs.

## 2026-09-09 (roadmap + CLAUDE.md)

**Slide 3 rebuilt to the house format.** It had diverged three ways from
Module 1's and Module 4's roadmap: no diamond connector arrows, a gold
"we are here" TEXT label to the right of the box instead of the gold block
ARROW plus italic label to its LEFT, and the bottom row 0.20" too high.
Also adopted the course-wide wording "4. Markets, Pricing, and Strategy"
(Module 6's own source deck omits the second comma; the recurring slide
wins, because students see it eight times).

**Six rules added to Teaching CLAUDE.md** -- none of them were there:

| where | rule |
|---|---|
| Figures/charts | a curve's label sits at its RIGHT END, or just above/below it; check the curve across the label's whole width, not just its anchor |
| Pictures | size by the CONSTRAINING dimension and compute the other from the real aspect ratio; a PNG export crops to the canvas, so overflow is invisible in a render |
| new `## Course-Roadmap Slide` section | the full diamond spec: box sizes and y positions, the four faded connector arrows, the gold block arrow + italic label to the LEFT, fixed wording |
| Back button | the `_add_jump_pill` trap -- it defaults to a WHITE pill and `back=True` does not change that |
| Reformatting a deck | a rebuilt diagram keeps the original's box positions, spacing and colour GROUPING; colours often live in a shape STYLE, so resolve `fillRef -> schemeClr` against the theme before concluding a diagram is unstyled |
| Build discipline | run a mechanical formatting audit on a new deck, tuned against documented geometry first; and a python-pptx inventory is BLIND to `mc:AlternateContent` |

`Teaching/CLAUDE.md` is git-tracked, so it needs committing when you next
ask for a commit.

## 2026-09-09 (label collisions)

**Slide 8.** The MR label was inside the revenue rectangle, which is
emitted AFTER it -- so the rectangle was painted over the label and took
the mouse, and the label was mid-curve besides. MR moved to just outside
the rectangle's bottom-right corner (which is MR's own right end), and
"Max. TR = $8" moved into the rectangle's LEFT half, the only part no line
crosses (MR enters the rectangle only from Q = 4 rightward). That is where
Nico's own slide 7 had it.

**New: `_check_labels.py`** reports two failure modes deck-wide --
COVERED (a later filled shape on top of a label: unclickable) and ONLINE
(a connector drawn through it). Neither shows in a PNG render.

Sequence: 46 raw -> 35 once leader lines anchored to their own label were
excluded (a callout's leader is the house pattern, not a defect) -> 24
after `_region_label` learned to clamp into its region and shrink to fit
-> 18 once judged on the label's inner 60% rather than the full box (slack
around the glyphs made steep curves look like strike-throughs) -> **0
COVERED and 2 ONLINE**, both marginal corner clips.

**The systematic fix** is `clear_of=((x0,y0),(x1,y1))` on
`_fig_curve_label`: give it the curve's logical endpoints and it pushes
the box along that line's normal by the exact clearance a rectangle needs,
`(|w*nx| + |h*ny|)/2 + margin`. Ten call sites use it. Nudging ten anchors
by hand would not have survived the next change to a curve.

Also reordered `s_netflix_lose` so the demand line and its label come
AFTER the three shaded regions -- previously the gray region was painted
over the "D" label.

**Teaching CLAUDE.md**: the curve-label rule was rewritten to cover all of
this -- comply with the right-end rule "to the extent possible", the two
failure modes, emit-order (regions, then curves, then labels), the
leader-line exception, and the two checkers.

## 2026-09-09 (in-class portion)

Nico asked whether the in-class block had been processed. Split answer:

**Formatting and labels: yes, it was.** The in-class copies call the SAME
slide functions as the video versions, so every fix that went into a shared
helper propagated. Both checkers report zero findings on any slide >= 60
(the single remaining `_audit_format` line is the documented 15 pt corner
pill on s84). The off-canvas images, the trailing periods and the region
labels that got clamped were in-class slides.

**Animations: no, and now yes.** Only slide 80 (the dolls) had a custom
plan; every other in-class chart was on the default one-bullet-per-click.
Twelve plans added, taking custom coverage from 11 slides to 23:

| slide | source |
|---|---|
| 55 | NV 62 -- the richest build in the deck, 15 clicks laying one block down at a time, each introduced by its own bullet. Comes out at 13 with axes static and curve+label grouped. |
| 58, 101 | the summary tree, video and in-class copies |
| 63, 70, 78, 89, 92 | in-class copies of video charts -- the SAME functions, so the same plans verbatim. Keeping them in step matters: a student sees the video build first and the class build second. |
| 95, 96, 97 | the three zoo panels, worked live in class |
| 100 | NV 63, the ice-cream blocks |

Verified in PowerPoint: on-click counts equal the beat counts in every
plan (s55 13, s63 7, s95 7, s100 9, s101 6).

**Trap worth remembering:** the tree plans initially used selectors from
the PRE-rebuild wording ("Every customer known"), which the engine caught
as a KeyError rather than animating the wrong shape. That failure mode is
the reason the engine raises on an unresolved selector, and it is why a
plan must be re-checked whenever the slide it targets is rewritten.

**Still open:** slides with original choreography that remain on the
default plan -- mostly bullet slides, where one-bullet-per-click already
matches what Nico's own decks do. `_anim_original_NV.md` has the specs if
any of them turn out to need a custom beat order.

## 2026-09-09 (completeness audit) — four gaps found and closed

Nico asked whether anything was missing. Measured against the shipped
Modules 3 and 4 rather than guessed:

| | Module 3 | Module 4 | Module 6 before | after |
|---|---|---|---|---|
| layouts | 1 | 1 | **11** | 1 |
| slides with notes | 91% | 73% | **32%** | 86% |

**1. Two approved slides had never been built.** NV 24 "Airline Pricing:
We Can Now Solve the 'Puzzle' from Week 1" and NV 25 "Lufthansa Pricing:
Solution" were in the approved outline as in-class 76-77 and the build
script had no trace of them. Now built: the two fare screenshots (LAX-IST
round trip $4,627 against the SHORTER LAX-FRA at $7,208), and his
two-panel answer rebuilt native. Both panels share one MC -- it is the
same LAX-FRA aircraft -- and the numbers are chosen so the result falls
out of MR = MC rather than being asserted: connecting P = 10 - Q gives
P* = 6, direct P = 14 - 2Q gives P* = 8, with an assert that the direct
flight comes out dearer.

**2. Eleven slide layouts.** The python-pptx template defaults were never
stripped; Module 4's `_strip_unused_layouts` is now wired into main().

**3. Speaker notes at 32%.** 65 written into `_m6_written_notes.py` and
applied by `_apply_written_notes`, which never overwrites a longer
existing note, so Nico's own notes always win (65 written, 0 of his
displaced).

**4. Podcast source docs missing** -- `Podcast Module 6 -- Intro.md` and
`-- Wrap-up.md`. Module 4 has both and the Fall 2026 calendar schedules
Module 6 podcasts. NOT built: they are a separate deliverable, flagged for
Nico.

### The renumbering trap, hit for real

Inserting the two slides shifted every display number after 74, and BOTH
the poll splice map and the animation plans are keyed by display number.
Left alone, the dolls plan (key 80) would have landed on the versioning
staircase, which is now slide 80. Fixed by reading the new numbers off the
built deck and remapping in DESCENDING order so no rename lands on a key
not yet moved: polls 72/85/96, plans 78->80, 80->82, 89->91, 92->94,
95->97, 96->98, 97->99, 100->102, 101->103.

**`WRITTEN_NOTES` therefore carries a TITLE GUARD** -- each entry is
`(title_prefix, note)` and the applier raises if the slide at that number
does not carry that title. A silent mis-attachment of 65 notes would have
been far worse than a crash.

## 2026-09-09 (arrow-origin pass) — every annotation label moved to its arrow's tail

Nico: "wherever we have arrows, put the corresponding text box to the
beginning of the arrow (the origin end). Make sure the text box does not
overlap with other lines (like on slides 16 and 17, which need to be
fixed)."

Two distinct defects were behind the slides he pointed at:

1. **The leader started at the label box's CENTRE**, so the line was drawn
   straight through the words. All three "…-MPV customer" callouts on
   slides 16 and 17 had a rule struck through "MPV". Fixed by a new
   `_arrow_from_box(slide, box_xy, box_wh, target, …)` in
   `_build_Module6.py`: it intersects the segment (box centre → target)
   with the box boundary, adds a 0.06" gap, and emits the arrow from
   there. `_fd_callout` now places the label FIRST and derives its leader
   from the box, instead of the other way round.
2. **Labels sitting at the arrow's head.** The decision tree's final "Yes"
   was down beside the outcome it pointed at rather than beside the
   question it left.

Positions that had to move to keep a label off another line (the rule
allows the arrow to grow):

| slide | label | to |
|---|---|---|
| 17 | "Do not sell to the low-MPV customer" | (9.80, 4.35), above MC |
| 17 | "Demand = MPV" | (3.15, 1.05) |
| 17 | "MC" | `FD_QMAX + 1.15, FD_MC + 0.15` |
| BMW panel | "D" | `pmax * 0.07` |
| Lufthansa panel | "D" | `pmax * 0.062` |

**New checker: ARROWEND** in `_check_labels.py` — for every arrow-headed
connector it finds the nearest label (within 1.2") and reports it when the
label's centre is nearer the head than the tail by more than 0.15".
**Axis arrows are exempt**, because an axis title is anchored to the arrow
TIP by the separate documented rule; they are recognised as axis-parallel
connectors longer than 2.5".

Final state of all three checks on `Module 6 - Revised.pptx`:
ARROWEND none, COVERED none, ONLINE none. `_audit_format.py` is at its 3
previously-triaged legitimate findings.

The rule is now in `Teaching/CLAUDE.md` under the curve-label section, in
both halves (label at the origin; arrow starts at the box EDGE), with the
axis exemption and pointers to `_arrow_from_box` and the ARROWEND check.

## 2026-09-10 — labels inside shapes, D not "demand", the tree opened up

Four items from Nico, all applied and both decks rebuilt.

### 1. Text inside a shape sits fully inside it, or goes outside with an arrow

His case was slide 9's "Consumer surplus" lying across the demand line.
The cause: `_region_label` centred the box on the region's centroid and
then *clamped it into the region's x-range*, which is not containment — a
triangle is only that wide at its BASE, and the label sits further up
where the hypotenuse (which IS the demand curve) has closed in.

`_region_label` now does real containment and tries three things in order:

1. one line, shrinking toward a floor, with a **search** for the nearest
   position inside the polygon that fits — this alone fixed slide 9, and
   "Consumer surplus" kept its full 17 pt;
2. the label re-wrapped onto **two lines**, shrinking again — this is what
   puts "Version 3" inside a 0.70"-wide staircase block on slide 34, where
   one line is wider than the block at any readable size;
3. failing both, the label goes **outside at full size with an arrow from
   its edge back into the region** (`_arrow_from_box`, so the leader starts
   at the box edge and never crosses the words).

New helpers: `_poly_clear` (signed distance to every edge, with the
winding sign taken from the shoelace area so no caller has to state it —
on a slide, where y grows DOWNWARD, guessing it is a coin toss),
`_box_in_poly` (four corners; every region here is convex) and
`_fit_in_poly` (anchor first, then a 32×32 scan for the nearest fit).

The floor stays at **11 pt**, not the 16 pt chart-label floor. A staircase
block or a thin profit rectangle is a "narrow object" in the Teaching
CLAUDE.md sense — its width is fixed by the data — and that exception says
fit the text to the object. Raising the floor to 14 sent all five
"Version k" labels outside, which would have wrecked slide 34.

`_region_label` prints a line whenever a label goes outside. On this build
it printed nothing: all 21 region labels fit inside.

### 2. Slide 29 — MR and D moved to their curves' right ends

They were at 0.42 and 0.80 of the way along, under the curve. Each now
sits just PAST its own line's intercept (demand at `a/b`, its MR at
`a/2b`) and a shade above the Q axis. Beyond the intercept there is no
line at all, so nothing can cross the words, and the two cannot collide
because the intercepts are `a/2b` apart.

The joint panel needed an override: its demand reaches the axis at Q = 68
of that panel's 72, leaving no room past the intercept, so its `D` goes
immediately ABOVE the right end, in the clear band between demand (P = 9
there) and the MC line at 40. That is the new `dend` argument on
`_bmw_panel`.

### 3. "demand" → "D" (slides 28 and 34)

Both spelled-out labels replaced, and their anchors moved 0.65-0.75 right
now that the box is narrower, so the label still ends at the curve's right
end. Rule added to Teaching CLAUDE.md, covering S / MR / MC / ATC too.

### 4. Slide 58 — the tree opened up

Both of his column ORIGINS are unchanged (`qx` 5.89, outcome right edge
13.05, which is the page margin). The room came out of the boxes, which
were far bigger than their text:

| | before | after |
|---|---|---|
| question box height | 0.92" | 0.78" (2 lines at 16 pt need 0.64") |
| gaps between question boxes | 0.17-0.23" | **0.34"**, uniform |
| gap to the outcome column | 0.53" | **0.71"** |
| Yes / No | 13 pt | **15 pt** |
| arrow weight | 1.5 pt | 1.75 pt |

**A trap worth recording: PowerPoint wraps wider than PIL measures.** The
first cut took `qw` to 2.85", which PIL said kept all four questions on 2
lines — PowerPoint put questions 3 and 4 onto 3 lines and burst their
0.78" boxes. Measuring the *slack* on the longest rendered line explains
it: at 2.85" question 3 had 0.01" left and question 4 had 0.06", while
question 2 survives at 3.15" with 0.10". So the real margin PowerPoint
needs is between 0.06" and 0.10", and the fix is to demand **≥0.20" of
wrap slack** rather than a bare fit. `qw` 3.05 and `ow` 3.40 both clear
that, which is where the 0.71" gap comes from.

**Why the gaps are 0.34" and not more.** Question 4 sits above the two
terminal cards, which reach across to x 6.71 and cannot move — their text
needs every bit of the 1.18" they have and they already stop 0.05" short
of the footer rule. The final "Yes" hangs below question 4 in that band,
so every 0.01" added to a row gap comes off that label's room. The
outcome boxes cap it from the other side: they stay 3 lines even at
ow 3.75, and 3.90 runs off the canvas.

Also: each "•" is now glued to its item with a no-break space, so a wrap
can only fall BEFORE a bullet. Narrowing the outcome column had left
"Two-part tariff  •" hanging at the end of a line.

### Verified

- ARROWEND / COVERED / ONLINE: **none** on both decks.
- `_audit_format.py`: the same 3 previously-triaged findings.
- Both decks open in PowerPoint (119 and 10 slides); slides 9, 29, 34, 58
  and the practice deck's slide 9 checked by render.
- The practice deck shares `_bmw_panel`, so it was rebuilt too (backups
  rolled first).

## 2026-09-10 (second pass) — film symbol, legend marks, and two grouping rules

### 1. The practice-video box now carries the FILM symbol

Nico: "follow the format that we used for other such boxes (i believe we
had a film symbol?)". He did. On **2026-08-23** the gold `▶` text glyph was
replaced deck-family-wide by the **action-button** markers, which say what a
link OPENS rather than which way it points — `ACTION_BUTTON_MOVIE` for a
video, `_SOUND` for a podcast, `_DOCUMENT` for an article. Module 6 was
still on the old glyph, and so was the Teaching CLAUDE.md spec for this box.
Both fixed. The button is seated the way `_add_jump_pill` seats its jump
button: a left inset, vertically centred, label indented past it.

**A second defect surfaced from the render:** the footer rule was drawn
straight through the box's label. The convention says "draw it AFTER the
footer", but `make_diagram_slide` runs the body callback BEFORE the footer,
so a box added inside `draw()` lands underneath it. Both call sites now add
the box to the returned slide instead.

### 2. Slide 49's legend carries a mark shaped like the thing it names

The legend-mark rule (a mark in the region's colour AND shape) was already
in the Teaching CLAUDE.md from Module 4, with `_welfare_rows` already
present in `_m6_helpers.py` but unused here. Slide 49 now uses it. Two
additions to the helper:

- a **`"line"` swatch kind** — a short solid bar for a legend line naming a
  CURVE rather than an area (the MC line), drawn without the area wash;
- **`"tri_h"`** — `MSO_SHAPE.RIGHT_TRIANGLE` mirrored horizontally. The
  unflipped preset came out as the mirror image of the flat-fee triangle,
  whose vertical leg is the y axis;
- rows now **size their text box to the lines it actually needs**, measured
  in the real font. "Flat fee F* = the entire consumer surplus" needs 4.99"
  in a 4.46" column, wrapped, and overflowed the fixed 0.34" box into the
  row below.

### 3. Arrows grouped with their labels; shapes grouped with their labels

Two new rules in `_group_pass.py` — 79 groups now, up from 12.

- **rule 6, arrow + label.** Includes the headless callout LEADERS (the six
  "…-MPV customer" callouts on slides 16 and 17), excludes axes and dashed
  guides.
- **rule 7, shape + label.** Region labels with their regions, and link
  buttons with the box they sit in.

**Two traps, both hit for real:**

- *A bbox test is not enough.* It grouped the CURVE labels `D` and `MR`
  with whatever shaded region they happened to sit inside. Fixed by pairing
  on a KEY: `_fig_poly` stamps `sdregion:<crc of the polygon>` and
  `_region_label` stamps `sdreglab:<same key>`. A label placed outside its
  region is not stamped, so it is never grouped with one.
- *A curve is a solid headless connector too*, and its endpoint lands on the
  axis exactly where a tick label sits — so the leader rule grouped demand
  lines with "$20". Fixed by naming the leaders at source
  (`_arrow_from_box` stamps `sdleader:`) and grouping only those.

### 4. The knock-on: grouping broke every plan that named the members

This is the part worth remembering. `_animate.py`'s plans address a region
and its label separately (`["osp:3", "t:Consumer surplus"]`). Once grouped,
the members are no longer top-level, so every `osp:` / `cxn:` ordinal after
a group SHIFTS — the engine raised `osp:4 resolves to CHROME`, which is
exactly the guard added in August after slide 52's `osp:7` hid a footer rule.

Rewriting fifteen plans by ordinal would have been the fragile answer. The
fix is in the engine instead: pair groups are **named** `sdpair:…`, and
`collect_shapes` indexes their MEMBERS as before while pointing each at the
group. Every existing plan keeps resolving, and the pair animates as one
object — which is what the house animation rule wants anyway. Three details
that each cost a debug cycle:

1. two members named in one beat resolve to the same target, so the beat
   deduplicates;
2. a pair group must NOT consume a `grp:` ordinal of its own — it did, and
   that pushed slide 55's box+text group from `grp:0` to `grp:3`;
3. the group pass is not idempotent over its own output. Re-running it on an
   already-grouped deck reports 0 groups, because the labels are no longer
   top-level. Always re-run from `_build_Module6.py`.

### Flagged, not changed

`SPLICED = {72, 83, 94}` in `_group_pass.py` is **stale** — the splice map
now targets displays 72, 85 and 96. So the pass is skipping slide 94 (an
in-class chart) and NOT skipping two real poll slides. Nothing was grouped
on 72/85/96 in this build, so no harm done, but it should be derived from
the splice map the way `OUTLINE_SLIDES` is derived, rather than hard-coded.

### Verified

- ARROWEND / COVERED / ONLINE: none, on both decks.
- `_audit_format.py`: the same 3 previously-triaged findings.
- Click counts unchanged on every grouped slide (beat count = plan length).
- Both decks open in PowerPoint (119 and 10 slides); slides 28 and 49
  checked by render.

## 2026-09-13 — slide 7: hand-edits adopted, new figure styled

Nico hand-edited slide 7 in PowerPoint and added a figure. Ported through
the side-path flow (`_build_Module6.py` frozen output to
`Module 6 - Revised_test.pptx`, diffed, ported, re-built, promoted).

**No backups and nothing in git.** The Module 6 folder is entirely
untracked and the `_t-1` / `_t-2` decks had been cleaned up, so his edits
existed in exactly one file. First action was to copy the deck to `_t-1`
before running anything.

### The three edits

| | |
|---|---|
| bullet column | 12.78" → **5.74"** wide, to clear the figure |
| "high price" / "low price" | set **bold** in the two sub-bullets |
| new figure | 1376 × 768 px, placed at 6.098 / 2.620", 6.846 × 3.821" |

`_add_media_image` is called with the WIDTH only, so python-pptx recovers
his 3.821" height from the native 1.7917 ratio exactly — width is the
constraining dimension here. Default `rounded=True, shadow=True` supplies
the "typical image edits"; he had pasted it flat, no corners, no shadow.
Saved as `_source_images/NV_s07_wtp_segments.png`.

### The bold was invisible to the geometry diff

`_diff_slides.py` matches shapes by (type, normalised TEXT), so a run-level
change inside an unchanged paragraph does not register — it reported only
the narrowed box and the new picture. The bold turned up in a run-by-run
read of `slide7.xml`, where the two sub-bullet paragraphs had split from
one run into three. **After any hand-edit port, read the runs of the edited
shapes, not just the geometry** — the existing CLAUDE.md note about
table-internal formatting applies to ordinary bullets too.

### Eight "differences" that were not differences

The diff also flagged slides 8, 28, 43, 51, 74, 91, 95 and 98, each with one
or two text-less shapes "only in the build". They are the documented
`mc:AlternateContent` artifact: his PowerPoint save wrapped every OMML math
box, python-pptx then stops enumerating them, and they read as phantom
deletions. Verified by counting the raw XML — his deck has
`AlternateContent=1, oMath=2` on slides 8 and 28 (and 2 / 4 on slide 98)
where the fresh build has `0 / 3` and `0 / 6`. The equations are present in
both. Nothing was lost, and nothing was ported from those slides.

### Flagged for Nico

The figure carries a **"Gemini Notebook" watermark** in its lower-right
corner, and the rounded corner clips its last letter. It cannot be cropped
away cleanly: it sits on the same band as the "Customer Segments" axis
title, so a bottom crop takes both, and a right crop takes the last bars.
Left as delivered pending his call.

### Verified

- `_diff_slides.py` on slide 7: clean.
- ARROWEND / COVERED / ONLINE: none. `_audit_format.py`: the same 3
  previously-triaged findings.
- Deck opens (119 slides); slide 7 renders with the bold, the narrowed
  column and the rounded/shadowed figure; 1 click, 4 effects (the picture
  joined the existing beat rather than adding a click).
- `_t-1` and `_t-2` both hold his hand-edited deck.

## 2026-09-13 — the Netflix example rescaled for realism

`P = 2 − Q/8` → **`P = 4 − Q/3`**, one subscriber's demand for movies in a
month. Nico: "$1 a movie and $8 a month were not Netflix numbers."

| | before | after |
|---|---|---|
| first movie worth | $2 | **$4** |
| stops watching at | 16 | **12 movies/month** |
| simple pricing (MR = 0) | 8 @ $1 | **6 @ $2** |
| Max. TR | $8 | **$12 / month** |
| consumer surplus, unexploited | $4, $4 | **$6, $6** |
| flat fee | $16 | **$24 / month** |

**Why 4 × 12.** Every derived figure stays a whole dollar only if both
maxima are even and their product divides by 8; 4 × 12 is the smallest such
pair in a believable range. The near alternative anchored on the Standard
tier (`P = 3 − Q/4`, flat fee $18) puts $1.50 and $4.50 on the slides. $24
also lands on Netflix's top tier, so the model's answer coincides with a
real price — the old $16 matched nothing.

**Nothing about the economics moved.** The flat fee is `PMAX·QMAX/2` and the
simple-pricing revenue `PMAX·QMAX/4`, so "exactly twice as much" holds for
any pair. Asserts in the constants block now state that, plus
`TR + 2·CS = flat` and that every figure is an integer.

### Everything is DERIVED now

`NFX_Q_SIMPLE` and `NFX_P_SIMPLE` used to be typed in (`8.0, 1.0`) beside
the maxima they are supposed to follow from. They are computed from
`NFX_PMAX / NFX_QMAX`, along with `NFX_SLOPE_DEN`, `NFX_TR_SIMPLE`,
`NFX_FLAT`, `NFX_CS_SIMPLE` and the plot ranges, so a figure and its
arithmetic cannot drift apart. Tick labels and the OMML equation are
formatted from the constants (`"$%g" % NFX_PMAX`), so they re-fit
themselves if the pair changes again.

### The five slides, plus two that ride along

8 and 9 (Video 1), 43 (flat fee), and the in-class copies 63 and 91. Slide
49 / 94 (two-part tariff) shares the demand: its `mc` is now
`0.375 × NFX_PMAX` — the same FRACTION of the choke price it always was —
so the picture keeps its proportions. That slide labels `F*` and `P*`
symbolically, so no number shows there.

### Two traps this pass

- **The animation plans select tick labels by TEXT.** `t:$2`, `t:$1`,
  `t:16`, `t:8`, `t:P = 2`, `t:Revenues = $16` — 10 selectors across slides
  8, 9, 43, 63 and 91. Changing a tick without its selector is a hard
  failure (the engine raises), which is the good outcome, but it is easy to
  forget that a *label* edit is also a *plan* edit.
- **A longer axis title collides with the last tick.** The x-axis title
  became "Movies per month" (from "Movies" / "Q"), because the dollar
  figures are monthly and the flat fee is a subscription price. The title
  is centred on the arrow TIP, so the long version reached back over the
  tick at `QMAX` and "12" printed under it. Fixed by widening the plot
  headroom from `QMAX × 1.125` to `× 1.25`, which buys 0.68" of clear air
  on the widest of these figures. Worth remembering: axis headroom is a
  function of how long the axis TITLE is, not just of the data.

### Flagged, not changed

`NOTES_NV[46]`, the verbatim source note on the flat-fee slide, describes a
different parameterisation entirely — "For a flat fee of $400, he would be
just willing to purchase the bundle of 20 units." It contradicted the slide
before this pass ($16 / 16 movies) and still does ($24 / 12 movies). Per the
source-notes-vs-slide rule this is Nico's call, so it was left alone. It
wants either a rewrite to the slide's numbers or a note in WRITTEN_NOTES.

### Verified

- ARROWEND / COVERED / ONLINE: none. `_audit_format.py`: the same 3
  previously-triaged findings. No region label had to go outside its shape
  at the new scale.
- Slides 8, 9, 43 and 49 checked by render; deck opens at 119 slides.
- Backups rolled before the rebuild.

## 2026-09-13 (third pass) — slides 8 to 12

Hand-edits adopted plus new work, through the side-path diff flow again.

| slide | done |
|---|---|
| 8 | demand equation moved to his spot (1.677, 2.340) and wrapped in a **red rounded box** — white fill, dark-red 1.75 pt border, soft shadow |
| 9 | his **28 pt** and his box (6.31, 2.15, 6.64 × 3.40); the bullet DOTS on rows 2 and 3 replaced by marks made of the figure's own regions |
| 10 | his cut of the "perfect competition → price taker" sub-bullet, box narrowed to 7.17", and his 3-D segments figure added with rounded corners + shadow |
| 11 | takeaway bar → **navy with white text**, "If buyers can resell, complex pricing is ineffective"; speaker note expanded |
| 12 | his per-row box heights (1.12 / 1.01 / 0.98, was a flat 1.62 for all three); tops unchanged, so only the gaps grew |

### Slide 9's region bullets

New `_region_mark` / `_region_bullets` in `_build_Module6.py`. Row 2 carries
the CS triangle and the unexploited triangle; row 3 carries those plus the
revenue rectangle, which together **are** the whole triangle under demand.
Both the colours/washes and the SHAPES are the regions' own, and the marks
keep the arrangement the regions have in the graph:

```
      +----+---+      cs  (red)   top-left
      | cs /   |      rev (gold)  below cs
      +----+---+      unx (grey)  right of rev
      | rev| u |
      +----+---+
```

This is the legend-mark rule applied to a bullet instead of a legend row.
Both triangles are `flipH` so their right angle is at the bottom left,
which is where the regions have theirs.

**It cost an animation rewrite.** The rows are three text boxes now (the
`_welfare_rows` pattern — each row carries its own drawn mark, and a mark
has to be positioned against the line it belongs to), so the plan's
`pr:Simplifying assumption:0:0 / 1:1 / 2:2` paragraph selectors no longer
addressed anything. They became `t:` selectors, and each bullet's beat now
also names its marks **by shape name** (`n:sdbullmark:1:cs`), which is
immune to the `osp:` ordinals shifting. Slides 9 and 63 both, 7 clicks and
18 effects each (was 12 — the marks joined their bullets rather than
sitting there from the start).

### Two things I got wrong and fixed

- **The first Dum-Dums note invented a channel structure** — "bulk bags to
  warehouse clubs, smaller packs to ordinary retailers". That is not in the
  Bloomberg clipping, which says only that rogue sellers order from Sam's
  Club and resell on Amazon at a markup, undercutting Spangler and
  violating Amazon policy. Rewritten to what the source supports, plus the
  general economics. The "violating Amazon policy" detail is worth keeping:
  an explicit contractual ban did not stop it.
- **The heredoc backslash trap, walked straight into.** The note patch was
  piped through a Bash heredoc; every `\n` in it arrived as a real newline
  and broke the string literals mid-file. The CLAUDE.md rule says to write
  such a patch to a `.py` file and run it by path — which is how it was
  repaired. Worth re-reading before the next multi-line note edit.

### Verified

ARROWEND / COVERED / ONLINE none; `_audit_format.py` at its 3 triaged
findings; deck opens at 119 slides; slides 8-12 checked by render; backups
rolled before the rebuild, `_t-1` holding his hand-edited copy.

## 2026-09-13 (fourth pass) — slides 8, 9, 15, 21, 33

| slide | done |
|---|---|
| 8 | red box moved to his spot (2.13, 2.26) |
| 9 | region triangles re-oriented to match the graph |
| 15 | his narrowed text box, his cut of the "hypothetical scenario" bullet, his figure, and a navy question box bottom-left |
| 21 | **his replacement figure — he did not ask for this one** |
| 33 | his wording ("airplane cabins", "cannot *directly* identify"), narrowed box, his figure |
| 49 | same triangle fix as slide 9 |

### The triangle flip was backwards, and had been since 2026-09-10

`rtTriangle` in DrawingML is `moveTo(0,h) -> lnTo(w,h) -> lnTo(0,0)` — the
right angle at the **bottom left**, hypotenuse falling from top-left to
bottom-right. That is already the orientation of every region in these
figures, so the `flipH` added on 2026-09-10 MIRRORED them.

It was a misread of a 0.24" mark in a 1400 px render. The fix this time
came from a 4000 px export cropped and doubled, which makes the
orientation unmistakable — and matches what the preset's path says it
should be. **Read a mark's orientation off the geometry or off a magnified
crop, never off a full-slide render.**

It also broke slide 9's row 3 in a way worth noting: two mirrored
hypotenuses zigzag, where two correctly-oriented ones are collinear. So
"overall a triangle" only works with the flip removed — the three marks now
share one straight hypotenuse and read as the whole area under demand.
Slide 49's legend mark carried the same defect and is fixed with it.

### Slide 8: the box is the anchor now

He dragged the red box to (2.13, 2.26) and the formula stayed behind at
(1.677, 2.340) — they are separate shapes, so selecting the box moved only
the box. The build now places the equation FROM the box constants
(`box + (0.20, 0.15)`), so the two cannot separate again. They are still
not grouped: an OMML shape wrapped in `mc:AlternateContent` loses its `a14`
namespace when extracted for a group, which is the standing exception.

### Slide 21 — a hand-edit he did not mention

The diff turned up a replaced figure on slide 21 (his 1024 x 1024 3-D
groups image at 7.59 / 1.703, in place of the old 5.30 x 3.68 one). It was
not on his list, but a rebuild would have discarded it, so it is ported.
**Flagged for him** in case the swap was exploratory.

### Verified

ARROWEND / COVERED / ONLINE none; `_audit_format.py` at its 3 triaged
findings; deck opens at 119 slides; slides 8, 9 (at 4000 px), 15, 21, 33
and 49 checked by render; backups rolled, `_t-1` holding his hand-edited
copy.

## 2026-09-13 — Tesla Model Y slide: parked, NOT built

Proposed as an international price-discrimination example after slide 28
(Segment Pricing), with a native two-bar Mexico-vs-US comparison. Nico
parked it. **No files were changed.**

**Why it stalled — do not re-attempt the fetch blindly.** `tesla.com`
returns **403 to everything**: WebFetch and curl alike, on `/modely`,
`/es_mx/modely`, `/modely/design`, the bare host, and the inventory JSON
API. Bot protection, not a bad URL. The archive route is no better — the
Wayback API rate-limited, and the 9 May 2026 snapshot of the Mexican
configurator is an 834 KB JavaScript shell: it carries option prices
(MXN 21,600 / 43,200 / 177,200) but no trim base prices, and its 133 "IVA"
matches are the string `CT_PRIVATE`, not the tax.

Press figures exist but conflict and are undated: Mexico MXN 799,000
(Standard RWD) to 1,129,000 (Performance AWD); the US search summary gave
Premium RWD as both $47,630 and $44,990 around a May 2026 increase.

**The trim names do not line up across markets** — US sells RWD / AWD /
Premium RWD / Premium AWD / Performance, Mexican listings name Standard
RWD / Long Range AWD / Launch Series / Performance AWD. Which pair counts
as the same car is Nico's call, not something to infer.

**Already in hand if it is revived:** FX USD 1 = MXN 16.9639
(exchangerate-api, 13 Sep 2026 00:02 UTC), and the PS 4 tie — Problem 4(e)
answers parallel imports with traders shipping San Diego avocados to
Atlanta and undercutting the scheme. The Tesla slide is the mirror image:
homologation, duties and a non-transferable warranty block the same
arbitrage, so the gap survives.

**Cost if revived:** the slide lands at display 29 and shifts everything
above it — animation plans, the poll splice map (72/85/96) and the
written-notes keys all need remapping in DESCENDING order. Build it once,
with verified numbers.

## 2026-09-13 (fifth pass) — slides 25/26 adopted, new slide 27, deck now 120

| slide | done |
|---|---|
| 25 | his EA Sports FC 26 cover + the game's own Amazon price history, replacing the old banner/chart stack |
| 26 | bullets narrowed to 7.27" and **"New Member Discount.png"** on the right |
| 27 | **NEW** — "Recall from Videos: Ways to Segment Consumers", his slide, adopted |

### The insertion, and the renumbering it forced

Slide 27 goes in after `s_ways_to_segment`, so every display number above it
moves up one. Remapped in DESCENDING order, per the note from 2026-09-09:

- `_animate.py` PLANS: 18 keys (29 … 103)
- `_m6_written_notes.py`: 52 keys, from 28 up
- `_splice_media.py`: 72 / 85 / 96 → **73 / 86 / 97**

Evidence it landed right: the build reports **65 notes written, 0 of mine
displaced** (every title guard matched), `_animate.py` raised no unresolved
selector, and all three poll slides still carry `tags` + `notesSlide` +
`image` at their new numbers — the notes-part trap avoided.

### `SPLICED` in the grouping pass is derived now

It said `{72, 83, 94}` and had been stale since 2026-09-09 — flagged then,
fixed now, because this insertion made it wrong twice over. A stale set is
worse than it sounds in both directions: it skips an ordinary slide, and it
does NOT skip a real poll slide, which is the one that must never be
touched because its notes part is load-bearing for the PollEv add-in. It is
read from `SPLICE_MAP` itself now, the way `OUTLINE_SLIDES` is derived.

### Two sizing calls worth recording

- Slide 25's two pictures and slide 27's are passed by WIDTH, which
  reproduces his heights exactly from the native ratios (0.932, 1.801,
  0.774) — nothing stretched.
- **Slide 26's is passed by HEIGHT.** "New Member Discount.png" is a 2:3
  portrait (1024 x 1536), so matching slide 27's *width* of 4.210" would
  have made it 6.32" tall and run it 0.68" past the footer rule. Matched to
  slide 27's HEIGHT (5.439") instead, right edge aligned with slide 27's so
  the pair reads as one column. This is the constraining-dimension rule
  doing its job.

### Flagged

Slide 27 sits INSIDE the Video 3 block, so it keeps `Module 6 · Video 3 ·
…` rather than the four-level `In Class · Examples` tag — the tag rule goes
by where a slide sits, not what it is for. If it is really destined for the
in-class section, it should move there and take the four-level tag with it.

### Still open — slide 29, Tesla image

The **Tesla US-vs-Mexico price image** is not inserted. There is no such
file in the deck, in `_source_images/`, or anywhere under the user profile.
Nico is sourcing one himself and will come back to it (2026-09-13).

When it arrives: drop it in `_source_images/`, add an `extras` callback to
`s_two_markets` (display 29), and place it with `_add_media_image` by the
CONSTRAINING dimension. Nothing else is pending on that slide, and no
renumbering is involved — it is an image on an existing slide, not a new
one. The earlier note on the abandoned *native two-bar chart* version, and
why tesla.com could not be scraped, is further up under
"Tesla Model Y slide: parked, NOT built".

### Verified

ARROWEND / COVERED / ONLINE none; `_audit_format.py` at its 3 triaged
findings; deck opens at **120 slides**; slides 25, 26, 27 checked by render;
backups rolled, `_t-1` holding his hand-edited copy.

---

## 2026-09-15 — hand-edits adopted, every speaker note rebuilt, podcasts written and corrected four times, folder cleaned

**One-line summary.** Adopted Nico's PowerPoint hand-edits round by round
(main-deck slides 96–100, then two full rounds on the practice-video deck),
rewrote or repaired the speaker notes of both decks against the slides,
produced and then four times corrected the two NotebookLM podcast source
docs, wrote two new podcast rules into `Teaching/CLAUDE.md`, and cleaned the
folder.

### State of the deliverables (verified by rebuild, not asserted)

| File | State |
|---|---|
| `Module 6 - Revised.pptx` | **100 slides**, 81 groups, **89 slides carry notes** (53 of them from `WRITTEN_NOTES`, the rest verbatim source captures) |
| `Module 6 - Practice Video - Optimal Pricing in Two Markets.pptx` | **12 slides** (was 10), **11 notes** |
| `Podcast Module 6 -- Intro.md` | 403-word body, 5 min ceiling, future tense |
| `Podcast Module 6 -- Wrap-up.md` | ~2,800-word body, 11 sections, 15–20 min |

The full pipeline reproduces both decks identically after the cleanup:
`_build_Module6.py` → `_splice_media.py` → `_group_pass.py` →
`_animate.py all apply`, and separately `_build_M6_practice.py`.

### Main deck — slides 96–100 (his hand-edits)

- Text edits adopted on 96, 97, 98, 99, 100; slide 100's last box deleted;
  slide 99's image replaced with the one he put in.
- `_dis_slide` gained `ix` / `iy` / `tx`, so one slide can place its picture
  where he put it and still pin the text column when the picture is narrower
  than the frame the column was sized against. Slide 99 is the reason.

### Practice-video deck — 10 slides to 12, over two rounds

**Round 1.** Slide 1 rebuilt to his new image and moved text boxes. Slide 9
reworked: the mark-up arrow moved to the left of the y axis, a `P*` added to
each panel, numbers on both axes, and the two panels matched to Module 5.
**Two new slides inserted after 9** — "What if they charged the same price in
both markets?": the calculation (including the aggregate demand) on the
first, the graph on the second. Slide 10 (now 12) retitled "Comparing Segment
Pricing to Uniform Pricing", its bottom box made navy, and the profit
comparison added (segment 1,700 vs uniform 1,600, +6.2 %).

**Round 2.** Slide 2's bottom box navy; **US** and **D** as the market
abbreviations everywhere; area letters dropped from the practice deck (the
main deck keeps B / C / A, so `_bmw_panel` takes `region_letter=None`);
"Your turn" boxes moved to the top of slides 3 and 6 under a **"What we
know"** header; one dark-green format for every Your-turn box; and
**"Solution"** as the first word of every solution slide's title.

New helpers in `_build_M6_practice.py`: `_yourturn`, `_section_header`,
`_panel_numbers`, `_markup_arrow`; constants `K_US` / `K_DE` and
`DARKGREEN`.

**Three traps worth remembering.**

1. `_add_convention_box(fill_alpha=…)` takes opacity as a **percent** and
   multiplies by 1000 itself — unlike `_fig_poly`'s `alpha`, which is already
   in 1/1000 %. `fill_alpha=14000` wrote `alpha=14000000` and PowerPoint
   refused to open the file, while `zipfile.testzip()`, lxml and python-pptx
   all accepted it happily. **Only the export probe caught it.**
2. The "Mark-up" label was struck through by the MR line in all three panels;
   it now sits below the MC line. `_check_labels.py` (ONLINE) is the check.
3. The kinked joint-demand panel labelled its y-intercept `$85` — where the
   extended lower segment would reach the axis — instead of Germany's choke
   price `$95`. `_panel_numbers(yint=…)` overrides it.

### Speaker notes — "update all notes to the slides"

- **`_check_notes.py` is new.** It audits notes against their own slide:
  NONOTE, EMPTY, and NUMBER (it folds spelled-out numbers to digits, so
  "four dollars" is caught against a slide reading `$8`). It found **14 stale
  notes and 2 missing**, which four rounds of reading would not have.
- **`NOTE_FIXES` in `_m6_written_notes.py`** is a new table beside
  `WRITTEN_NOTES`: 8 unconditional overrides (slides 26, 60, 61, 79, 89, 90,
  91, 93), each guarded by a substring that must appear on the slide, so a
  renumbering fails loudly instead of writing the note to the wrong slide.
  `_apply_note_fixes(prs)` runs before `_apply_written_notes`.
- **The zoo notes had drifted.** Slides 70 and 72–74 still quoted the
  pre-rescale demand `20 − 4Q` and `MC = 4`; they now read `40 − 8Q` and
  `MC = 8`. This is exactly the failure the NUMBER check exists for.
- `WRITTEN_NOTES` grew to **53** entries; 2, 83, 96–100 rewritten.
- The exam date is substituted at build time (`{final_date}` / `{final_slot}`)
  rather than typed into a note.
- The practice deck got its own `PV_NOTES` (11 entries) plus
  `_apply_pv_notes`, display-keyed with a title guard.

**`_m6_notes.py` still must not be reworded** — it holds verbatim captures
from the source decks, keyed by SOURCE slide number.

### Podcasts — produced, then corrected four times

Both docs follow the standing conventions: H1 exactly `Module 6 - Podcast
Intro` / `Module 6 - Podcast Wrap-Up`, a paste-ready Audio Overview prompt at
the top of the doc **and** printed in chat, one notebook per episode.

Four rounds of corrections, each from Nico, and each traceable to the same
root cause — **the wrap-up had been built from the speaker-notes dump rather
than from the slides**:

1. **Hospital merger.** Slide 11 is a WSJ clipping plus an open discussion
   question; the insurer-negotiation finding existed only in the notes. A
   sweep found three more of the same kind: Costco's $65 / $130, the three
   Barbie prices (notes-only *and* deliberately hidden for an in-class
   reveal), and the Kodak FunFilm story.
2. **Disneyland / "more elastic" Southern Californians.** Slide 28 is a
   picture with a title and nothing else. Re-sweeping *every* example then
   found **Windows was also notes-only** (slides 44–46 never name Microsoft;
   slide 49's on-slide list is internet speed tiers, FedEx/UPS, and the three
   airline cabins) — and that two of my own guards were too strict: the
   Lodine answer **is** on slide 37 and the United/Southwest answer **is** on
   slide 53.
3. **Netflix framing.** Slide 9 is titled "What if Netflix Used Simple
   Pricing per Movie? (Assume MC = 0)" and the episode had narrated it as
   company history. It is now the episode's opening thought experiment, and
   both the $12 and the $24 are introduced as figures from one stylised
   demand curve. The episode also always says "a simple price per movie",
   never a bare "single price" — the flat monthly fee is a single price too.
4. **Consumer welfare removed.** A host had said a simple per-movie price
   "would be a financial disaster for the company and for consumers", which
   has the sign backwards: under one price the customer **keeps** the
   surplus. The whole welfare question is now out of the episode rather than
   argued the other way, and that also removed an answer the episode was
   giving to slide 21's live discussion prompt. The measure throughout is the
   firm's profit; slide 10's own label, "wasted profits", is the register.

**Two rules added to `Teaching/CLAUDE.md`** (Podcasts section) so this cannot
recur:

- *Build the episode from the SLIDES, and check every claim against the
  on-slide text before it goes in* — notes are instructor notes and carry
  discussion answers and deliberately hidden numbers; discussion / poll /
  "In the News" slides pose questions answered live and never assert an
  answer; put the guard in the episode too.
- *A COUNTERFACTUAL on a slide is a thought experiment, never history* —
  read the title for its framing ("What if…", "Suppose…", "Assume…"), name
  the real product in the same breath, and treat a stylised example's figures
  as the example's, not the firm's accounts.

### Folder cleanup

Restricted to **definitionally regenerable artifacts** — render / probe
output, `__pycache__`, rolling `_t-1` / `_t-2` backups, side-path `_test`
copies. **Every script was kept**, because Module 6 was untracked in git and
a script deleted here would have had no other copy. Verified afterwards that
the whole pipeline still reproduces both decks identically and that
`_mkside.py` still regenerates the side-path hand-edit flow.

`_splice_media.py` reads only the sidecar `_handoff_polls_M6.pptx`, so the
107 MB of source decks are **no longer pipeline inputs**.

### Open / pending

1. **Two deletions await Nico's decision** (deleting source decks needs his
   confirmation): `Module 6 - NV Slides/` + `Module 6 - PG Slides/`
   (107 MB of original decks) and roughly 71 MB of unreferenced extracted
   assets inside `_source_images/`. Together they would take the folder from
   **239 MB to about 60 MB**. Doing it before the first commit would have
   kept them out of git history; from this commit on, history keeps them.
2. Module 6 had **never been committed** before this session's commit.
3. Pushing this repo over this network needs the Git Data API fallback for
   anything over about 10 MB (see the memory note on push 408s).

### Commands worth remembering

```
export PYTHONIOENCODING=utf-8          # required on EVERY command (cp1252 console)
python _build_Module6.py && python _splice_media.py && python _group_pass.py && python _animate.py all apply
python _build_M6_practice.py
python _check_notes.py                 # NONOTE / EMPTY / NUMBER
python _audit_format.py                # OFFCANVAS / BACKPILL / PAGENUM / OVERLAP / PERIOD / TITLECASE / SIZE<18
python _check_labels.py                # ARROWEND / COVERED / ONLINE
python _check_anim.py                  # CHROME / AXIS / NEWS / CURVE / PANEL / BAR / FIRST
powershell -File _export_probe.ps1     # the only check that catches a file PowerPoint refuses to open
```

Hand-edit flow: copy the deck to `_handedit_*.pptx`, run `_mkside.py`, build
to `*_test.pptx`, run `_diff_slides.py <handedit> <fresh build>`, port into
the build script, rebuild. Slides 2–8 of the practice deck always show
phantom MOVED chains in that diff — `mc:AlternateContent` hides OMML shapes
from python-pptx, and those are save artifacts, not edits.

---

## 2026-09-16 — podcasts: three examples dropped from the intro, both prompts hardened against two recurring host failures

**One-line summary.** Four more correction rounds on the podcast source
docs, all of them the same two failures — the hosts attaching an example to
the wrong case of the module's taxonomy, and treating a worked example's
invented figures as real company numbers. Both are now guarded in the
prompts, in the in-doc host instructions, and (for the figures) in the
wrap-up body. One new rule in `Teaching/CLAUDE.md`.

### Failure 1 — an example paired with the wrong case

Nico, on the intro episode: the hosts used **the arthritis drug** and **the
browser-based hotel prices** as illustrations of "customers sorting into
different segments", and then, a round later, explained the **Lufthansa
fare** as something the airline knew about the passenger. All three are
wrong in the same direction: the first two are cases where the seller works
out which GROUP a buyer is in, and the Lufthansa fare is about which market
the passenger is flying to.

**Cause.** The preview names examples without working them through — that is
what makes it a preview — so its prompt held a flat list of examples sitting
right beside the three abstract cases. The hosts pair the nearest example
with the case just described and invent the reasoning that justifies it. The
invented reasoning is always plausible, which is why it survives a
read-through.

**Fixes, and they are opposite in the two episodes:**

- **Intro:** all three examples deleted. The list is now **closed** ("use no
  example other than those") and the hosts are told to name each in passing
  only, never to say which case an example belongs to, and never to explain
  why the thing is priced as it is. Body down to 386 words.
- **Wrap-up:** silence is not available, because a recall episode has to
  place each example. Its prompt's flat run of twenty-odd examples is now
  **grouped by case** — one clause per case — so the pairing is given rather
  than inferred, followed by the three wrong pairings named explicitly. The
  grouping was taken from the body's own sections 2 and 5–9, which had
  already been checked against the slides on 09-15, not re-derived by reading
  the examples again. No example added or dropped.
- **Wendy's is deliberately given NO case.** The body files it under
  versioning, but varying prices by time of day is a timing segment and what
  the class draws from it is the backlash. A standing aside beats a category
  it half fits.
- **Two instructions that pull against each other are themselves a bug.**
  Once the intro carried both a blanket "do not attach an example to a case"
  and the older clause letting coupons illustrate self-selection, the hosts
  had licence to use whichever half suited the sentence. The coupon
  permission is now written as the single stated EXCEPTION, in the same
  breath as the rule.

### Failure 2 — a worked example's figures narrated as real

Nico: "They will often treat the examples from class such as the BMW example
as if those numbers were real." The guard for this already existed — it was
written on 09-15 as part of the Netflix counterfactual rule ("figures from a
stylised example are not the firm's accounts") — and I had applied it to
Netflix alone, because Netflix was the example he had raised. BMW, Zipcar,
the zoo and the ice cream kept stating their numbers flat, as though BMW had
reported 1.7 billion.

The fix is general this time. His own phrasing is now the model in the
prompt and the host instructions: **"as we saw in the BMW example, based on
the hypothetical demand curve for that case…"**, with the fallback that a
figure which cannot carry the hedge is left out and the comparison given in
words. What a worked example establishes is the DIRECTION of the gap between
two strategies and roughly how big it is, which survives whatever the real
numbers are.

The wrap-up body hedged the four cases that stated figures flat — BMW,
Zipcar, the zoo exercise, the ice-cream scoops. Body now 2,847 words.

**One exception, named in both places:** the **American Airlines lifetime
first-class pass** was a real product at a real price and the body tells it
as history, so hedging that one would introduce the opposite error.

### Also fixed

- The Lufthansa guard was first written as "which market the passenger is
  flying to **and how easily that passenger can substitute**", which
  contradicts the existing instruction not to attribute travel alternatives
  to any group. Narrowed; the body still carries the full reasoning.

### `Teaching/CLAUDE.md`

One rule added to the Podcasts section (50 lines, before "Lead with
real-world stories"): **an example belongs to exactly one case of whatever
taxonomy the module teaches — say which, or say nothing.** It carries the
cause, the opposite fixes for a preview and a recap, the instruction to name
the specific tempting mis-pairings, the no-case-for-a-straddler rule, the
ban on self-contradicting prompts, and the writing-time test (name each
example's case and say in one clause why; if the clause has to be invented,
the example does not belong in the episode).

**Still worth doing:** the existing "figures from a stylised example are not
the firm's accounts" sub-rule is written against Netflix specifically.
Generalizing it to every worked example, with Nico's phrasing as the model,
was offered and not yet approved.

### Open / pending (carried forward)

1. The two deletions still await his decision: `Module 6 - NV Slides/` +
   `Module 6 - PG Slides/` (107 MB) and the 116 unreferenced files in
   `_source_images/` (70 MB, including 30 MB of .mp4 extracted from the PG
   decks). Measured this session: only 81 files / 36 MB of `_source_images`
   are named by any script, and `_splice_media.py` reads the 3 MB sidecar
   alone.
2. He decided on 09-16 to **push Module 6 whole**, source decks included,
   rather than rewriting the unpushed commit to exclude them ("slow is
   fine"). Note for next time: the ~10 MB push ceiling on this network means
   this repo needs the Git Data API route regardless of what is excluded —
   the 27 MB deck alone is over it.
3. `405 Slide Revisions 2026/Module 5/` has **no Session-Notes.md and is
   entirely untracked** — the position Module 6 was in before 09-15.

## 2026-09-23 / 09-24 — introduction video added, then Module 6 FINALIZED

**One-line summary.** Added an introduction video (Video 1) to the Revised
deck and shifted the topic videos to 2–9. After Nico taped all the videos,
assembled `Module 6 - Final.pptx` from the taped decks with the new
"Finalize Module X" routine, now written into the project CLAUDE.md.

### 09-23: introduction video, videos renumbered
- `_build_Module6.py`: new `intro_video_card` ("Introduction to Module 6 /
  Module 6 · Video 1") as slide 1. Title, logistics, roadmap and the
  descriptive overview form Video 1's block and carry `Video 1` tags.
  `VIDEO_OFFSET = 2`, so the outline starts at Video 2. Umbrella pills now
  read "Videos 3+4" and "Videos 6–8".
- Every display-number key shifted by the one inserted slide:
  `WRITTEN_NOTES`, `NOTE_FIXES`, the `PLANS` in `_animate.py` (plus
  `SKIP_STATIC` 3 -> 4) and `SPLICE_MAP` (37/53/72). Click counts were
  checked against the old deck: 0 mismatches.
- Lesson: the first renumber script matched brackets through string
  literals and shifted `NOTE_FIXES` twice. The note title guard caught it.

### 09-24: Finalize Module 6
- Inputs: the 9 taped decks in `Recorded Video Slides/`. The practice-video
  deck is out of scope.
- Output: `Module 6 - Final.pptx`, 135 slides. Slides 1–93 are the main
  deck (intro card dropped). Slide 94 is the in-class divider (copied from
  Module 3 slide 93). Slides 95–128 are the 34 in-class copies, retagged
  `In Class · Examples`. Slides 129–135 are Disneyland, moved from after
  Backup. The seating chart stays in Backup (slide 93, linked from 53 and
  116).
- 55 taped slides replaced their Revised versions. All are verified
  structurally identical to their sources; links, polls and notes check out,
  and the slideshow probe passes on all five poll slides.
- Pipeline (read-only first, then write):
  `python _final_inventory.py` (pairing -> `_final_pairing.json`) ->
  `_final_plan.json` -> `powershell -File _finalize.ps1` ->
  `python _final_retag.py` -> `python _final_verify.py` and
  `python _final_verify_tree.py`.
- Retag gotcha: editing the tag through COM re-autofits the tag box
  (0.42" -> 0.27"). The retag is therefore a string edit inside the run.
- `_build_Module6.py` now carries a STALE banner. The **Final deck is the
  source of truth** and is edited in place.

### Decisions (Nico)
- The in-class examples appear twice: once in the main deck, once in the
  in-class section. Unlinked backup slides (Disneyland) move and are not
  copied.
- My calls, reported to Nico: logistics and the summary closer have no
  in-class copy. The "Only if we have time" divider is copied with its six
  slides. Disneyland stays one block.

### Open
1. Slide 80 (ice cream) sits in Video 8's block but its taped tag still
   reads `In Class · Examples · Block Pricing`.
2. Ten Video-tagged slides were not taped. They keep their `Video k` tags
   in the main deck.
3. The course calendar, website and podcasts may still number the videos
   1–8.
4. The Dum-Dums speaker note still says "an explicit contractual ban".
   Nico will fix that himself.
5. Carried forward: the deletion of the NV/PG source decks and the unused
   `_source_images` still awaits his decision.

## 2026-09-24 (later) — Folder cleaned up; support files moved to `_build_files/`

**One-line summary.** Module 6 is finished, so the folder was reduced to the
deliverables plus one support subfolder. `Module 6 - Final.pptx` remains the
source of truth and is edited in place; nothing is regenerated.

### What the module folder holds now

| Item | What it is |
|---|---|
| `Module 6 - Final.pptx` | **The deliverable.** 135 slides, edited in place |
| `Recorded Video Slides/` | The 9 taped decks + the taped practice-video deck. The only source if a taped slide has to be re-copied |
| `Podcast Module 6 -- Intro.md` / `-- Wrap-up.md` | NotebookLM sources for the two episodes |
| `Session-Notes.md` | This file |
| `_build_files/` | Everything else — scripts, dumps, `_source_images/`, the poll sidecar |

### Deleted (all recoverable from git history)

- `Module 6 - NV Slides/` (11 decks, 58 MB) and `Module 6 - PG Slides/`
  (3 decks, 49 MB) — the originals the rebuild started from.
- `Module 6 - Revised.pptx` (27 MB), superseded by Final.
- `Module 6 - Practice Video - Optimal Pricing in Two Markets.pptx` at the
  root, superseded by the taped copy in `Recorded Video Slides/`.
- `Module 6 - Revised - outline.md`.
- 12 one-off scripts: `_final_detail.py`, `_map_sources.py`, `_mkside.py`,
  `_pick_slides.py`, `_probe_build.py`, `_snap_titles.py`,
  `_extract_anim.py`, `_extract_assets.py`, `_extract_pg_media.py`,
  `_make_notes.py`, `_make_poll_sidecar.ps1`, `_deck_guard.py`.

This closes open item 5 of the previous entry.

### `_build_files/` — what is in it and why it was kept

- **The drawing layer**: `_m6_helpers.py`, `_build_Module6.py` (STALE),
  `_build_template_samples.py`, `_m6_notes.py`, `_m6_written_notes.py`,
  `_source_images/`. This is what a **new** slide in the deck's style needs:
  build it into a one-slide side deck, then `InsertFromFile` it into Final.
  The chain is `_animate.py` -> `_build_Module6.py` -> `_m6_helpers.py` ->
  `_build_template_samples.py`, so none of the five can go.
- **Builds and grouping**: `_animate.py`, `_group_pass.py`,
  `_splice_media.py` (with `_handoff_polls_M6.pptx`, the verbatim poll
  slides with their notes and `tags` parts — a poll slide that loses its
  notes part crashes the full-screen slideshow deck-wide).
- **Read-only auditors**: `_audit_format.py`, `_check_labels.py`,
  `_check_anim.py`, `_check_notes.py`, `_check_coverage.py`.
- **Dumps and diffs**: `_dump_deck.py`, `_dump_raw.py`, `_dump_text_raw.py`,
  `_diff_slides.py`. Use the raw dumps to read Final — python-pptx is blind
  to `mc:AlternateContent`, which wraps every OMML formula box.
- **Probes**: `_export_probe.ps1`, `_slideshow_probe.ps1`.
- **The finalize pipeline**: `_finalize.ps1`, `_final_inventory.py`,
  `_final_retag.py`, `_final_verify.py`, `_final_verify_tree.py`,
  `_map_videos.py`, `_final_pairing.json`, `_final_plan.json`. Kept because
  the project CLAUDE.md's "Finalize Module X" routine names these files, and
  Modules 7 and 8 will run it.
- **Research dumps**: `_source_inventory_*.md`, `_source_rawtext_*.md`,
  `_assets_manifest_*.md`, `_anim_original_*.md`. Kept precisely because the
  NV and PG decks are gone — these are their searchable text and animation
  record at 1/300 of the size.

### Path convention inside `_build_files/`

Every script was repointed when it moved, and the rule is uniform:

- `HERE` is `_build_files/` — `_source_images/`, `_handoff_polls_M6.pptx`
  and the `_final_*.json` plans all moved with the scripts, so those
  references did not change.
- `MODULE = HERE.parent` is the module folder, and every reference to a
  DECK or to `Recorded Video Slides/` goes through it. In PowerShell,
  `$folder = Split-Path $PSScriptRoot -Parent`.
- **Read-only auditors now default to `Module 6 - Final.pptx`**, so
  `python _check_labels.py` with no argument audits the deliverable.
- **Writer passes still default to the deleted `Module 6 - Revised.pptx`,
  deliberately** — an accidental `python _animate.py` then fails loudly
  instead of rewriting Final. Pass the deck explicitly to run one.
- `_build_Module6.py` reaches the course calendar at `parents[3]`, not
  `parents[2]`.

### Two scripts that need the deleted source decks

`_check_coverage.py` compares the rebuild against the NV decks, and
`_dump_text_raw.py`'s deck KEYS ("NV", "V1", …) point into
`Module 6 - NV Slides/`. Restore that folder from git history to use them.
`_dump_text_raw.py` now also accepts a `.pptx` PATH, so
`python _dump_text_raw.py "../Module 6 - Final.pptx" 9` works as it is.

### Verified

All 23 scripts compile and every local import resolves (checked statically
with `py_compile` + an AST walk — never by importing, since these passes do
their work at module level). The three `.ps1` files parse.
`_check_labels.py` and `_audit_format.py` were run against Final from the
new location: labels clean; the format audit reports 5 pre-existing
findings (one trailing period on slide 9, four 15 pt labels that are the
intended corner-pill and version-label sizes).

## 2026-09-24 (later still) — "Final" renamed to "Full"; a second step is coming

**One-line summary.** The deck assembled from the taped decks is the FULL
slide version, not the final one. `Module 6 - Final.pptx` is now
`Module 6 - Full.pptx`, the routine that builds it was renamed throughout,
and the name `Module X - Final.pptx` is reserved for a second step that
Nico will specify after he has edited the in-class example slides.

**Reading the earlier entries:** everything they call `Module 6 - Final.pptx`
is this same deck under its old name. They are left as written.

### Renamed

| Was | Is |
|---|---|
| `Module 6 - Final.pptx` | `Module 6 - Full.pptx` |
| `_finalize.ps1` | `_assemble_full.ps1` |
| `_final_inventory.py` | `_full_inventory.py` |
| `_final_retag.py` | `_full_retag.py` |
| `_final_verify.py` | `_full_verify.py` |
| `_final_verify_tree.py` | `_full_verify_tree.py` |
| `_final_pairing.json` | `_full_pairing.json` |
| `_final_plan.json` | `_full_plan.json` |

The scripts were renamed so that step 2 can have `_final_*` names of its
own without colliding. All cross-references were updated: the imports
between them, the two JSON filenames they read and write, the deck name in
the seven auditors and dumps, the rolling-backup names in
`_assemble_full.ps1`, and the STALE banner in `_build_Module6.py`. A repo
grep for `Module 6 - Final`, `_final_` and `_finalize` comes back empty.

### The routine, in the project CLAUDE.md

`405 Slide Revisions 2026/CLAUDE.md` now heads the section **"End of
Module, Step 1: Create the Full Slide Version"**, with the short name
Nico gave it: *"Create the full slide version (based on `Module X -
Revised` and on the recorded video slides)"*. The rules are otherwise
unchanged — only the output name and the wording moved from Final to Full.

Two things were added to that section:

1. A note that a **second step follows and is not specified yet**, that
   `Module X - Final.pptx` is RESERVED for its output, and that step 1
   never writes that name.
2. An instruction to **ask which step is meant** when a request says
   "final", until step 2 is written down.

The video decks are described as "**fixed** exactly as they are" rather
than "FINAL exactly as they are", so the word "final" keeps one meaning in
that section.

### Verified

All 23 scripts compile and every local import resolves, including the
renamed chain `_full_verify_tree` -> `_full_verify` + `_full_inventory`.
The three `.ps1` files parse. `_check_labels.py` and `_audit_format.py`
run against `Module 6 - Full.pptx` with no argument and report the same 5
pre-existing findings as before the rename.

### Open

- **Step 2 is undefined.** What it takes as input (the edited Full deck),
  what it drops or keeps, and what it verifies all still need to be
  written down before it can be run.

### Addendum — the cleanup is now a written-down step

`405 Slide Revisions 2026/CLAUDE.md` gained **"End of Module, Step 1b:
Clean Up the Module Folder"**, generalized from what was done here: the
five items the module folder keeps, what is deleted, the `_build_files/`
subfolder and its seven groups, the import chain to resolve before
deleting a script, the `HERE` / `MODULE` path convention, and the static
verification. Step 1's section now points forward to it. Module 6 is the
worked example of the routine.

**`In Class Material/`** (2026-09-24, Nico) is now a sixth top-level item
in a module folder, alongside the Full deck, `Recorded Video Slides/`, the
podcast sources, `Session-Notes.md` and `_build_files/`. It is Nico's own
folder for what he takes into the classroom, it is excluded from the
cleanup step even when empty, and it will exist in every module folder.
Nothing of ours goes in it. Note that git does not track an empty folder,
so it will not appear on another machine until it has a file in it.
