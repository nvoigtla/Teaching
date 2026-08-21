# Session Notes — Module 2 (In-Class + Video Part decks)

## 2026-08-16 — Video Part rebuild (Videos 1–3, 57 slides)

Built `Module 2 - Video Part Revised.pptx` from the 40-slide 4:3
original, adopting ALL 9 CT items (incl. Netflix — Nico may re-delegate
it to Practice Video 1): Ozempic sequence (facts verified: 2026 guidance
sales/profit −5–13%, shares −18%), McDonald's price-cut discussion
(verified, incl. ~$35M franchisee compensation for the notes), Mega
Millions revisited (corrected form + large-change caution), Inside Out 2
revenue maximization (MR=0 → P=$20), A/B-testing slides, expanded
airline application (Eᴅ=−0.92 → MR=−12.3 → raise price), "Why
randomization is key", module summary. NEW title slides for Videos 2+3.
Group Discussion badges (Nico's badge, relabeled + stretched via group
scaling) on Ozempic/Netflix/McDonald's/MegaMillions/InsideOut2.

**Pipeline (mirrors the In-Class one):**
```
python _build_Module2Video.py     # imports the In-Class helper layer
python _splice_video.py           # polls: new 9 <- old 9, new 28 <- old 20
python _group_pass.py "Module 2 - Video Part Revised.pptx" --spliced=9,28
python _animate_video.py all apply
```
- `_build_Module2Video.py` pulls the ENTIRE In-Class namespace via
  `globals().update(vars(_M))` (plain `import *` skips _-names!).
- Verbatim teleprompter notes auto-extracted to `_video_notes.py`
  (BUILD INPUT) and re-attached by OLD slide number.
- Agenda convention (2026-08-16, BOTH decks, CT format EXACT per
  Nico's correction): gold 0.58" circle at x=1.15 with 25 pt bold NAVY
  number; one text box at x=2.05 with the 25 pt bold navy title and a
  22 pt gray description line — description only for the current
  topic(s), all descriptions on the overview/summary. In-Class slides
  7/8/26/76/77 and Video slides 2/20/33/57 all use it. The 76/77
  post-work pointers moved to a bottom-right link box overlaying the
  footer (deck convention) because right-side boxes overlapped the
  wider description lines.
- Charts rebuilt native and exact: D + TR parabola (exact quadratic via
  cubic Bézier, peak aligned at Q=800), TR rectangles on P=10−Q, MR
  hits 0 exactly at 800 (and at half the intercept on the generic
  version), airline scatter as 16 shape dots + exact fitted line
  Q=479−1.64P.
- Verified: full-deck render review (2 fixes: A/B image height, MR
  solution line width), COM click check (all 47 animated slides match
  the plan), slideshow probe PASS (title, both live polls, A/B, summary).
- Original deck has NULL-target image rels — extraction/build guards
  skip them (python-pptx would choke; splice is zip+lxml as always).
- Minor conscious deviations: CT's full-bleed Netflix/lottery background
  stills not carried (clean bullet slides instead); slide 37's
  camelcamelcamel caption reveals as its own click.

# Session Notes — Module 2 In-Class deck (2026-08-15)

## 2026-08-15 — Full rebuild of "Module 2 - In Class with Solutions" into the new format

**One-line summary.** Built `Module 2 - In Class Revised.pptx` (**77
slides** after the 2026-08-15 cross-check additions, 16:9, new 405
format) from Nico's 69-slide 4:3 deck, adopting the CT innovations;
polls + pizza Excel embed spliced live; grouped; fade builds applied
and verified.

### CT cross-check round (2026-08-15, all approved)
Slide-by-slide diff against CT's deck found one missed adoption and six
refinements — all implemented (see the Addendum in the outline file):
bookend law-of-demand recap as NEW slide 14 (deck now 77 slides; all
pipeline configs renumbered ≥14 by +1), Netflix step chart + gold
"+$1 in 2014" callout, water-solution takeaway bar, two-direction
factors shift graphic + "(within the firm's control)" + "Anything
else?" prompt, Problem-Set pointer made generic ("Problem Set 2", no
exercise numbers — Nico wants numbering flexible), gas caption moved
above its picture and grouped (17 groups now). Full pipeline re-run;
click counts re-verified (53 slides match, incl. factors 6 clicks and
water 4); slideshow probe re-passed on slides 1/11/13/14/46/77.

### Pipeline (rerunnable, Module 7 pattern — 4 steps)
```
python _build_Module2InClass.py    # phase-1: all 59 scripted slides + stubs
python _splice_media.py            # 8 PollEv pairs + Excel slide, verbatim
python _group_pass.py              # box+text / shade+frame / pic+caption groups
python _animate.py all apply       # fade builds per per-slide plans
```
The build script is the source of truth; splice + animate re-run after any
rebuild. `_build_template_samples.py` (helpers) and `_animate.py` (engine)
copied from Module 7; `_splice_media.py` adapted with two M2 changes:
poll NOTES are copied from the source deck (PollEv reads the poll URL from
notes — missing notes crash the slideshow deck-wide), and the copy loop is
recursive so the Excel embed's vmlDrawing rels + EMF preview travel along.

### Decisions locked (2026-08-14, from chat + outline)
- Adopt CT items: descriptive outline (descriptions on slide 7 only;
  section agendas titles-only with cream-band highlight), three-types
  overview + re-anchor, law-of-demand D-curve graphic, native Netflix
  price chart, market-vs-firm slide (CT's gas image kept as-is),
  cheat sheet, Gjelina / CorePower naming, corrected Mega Millions
  example, 3 "In the News" slides (WSJ clippings from CT's file).
- CT's lottery example was FACTUALLY WRONG ("MA State Lottery 2024,
  142,170→100,297" — unverifiable). Replaced with the real event: Mega
  Millions, April 2025, $2→$5; NY sales ~1.9M→~560K per drawing
  (Hansen, Misra & Singh study), with a Method-1 caution line.
- "MOV of 4th slice" typo → MPV (slide 15). Stray WSJ links in old #45/#56
  notes deleted. Logistics dates left as [DATE] placeholders (slide 2).
- Target-vs-Walmart recession figure rebuilt as a NATIVE chart
  (Stevenson/Wolfers source line; series approximate, digitized from the
  printed figure).
- Slide numbering map (new↔old) is in `Module 2 - In Class Revised -
  outline.md`; spliced slides: 4,5,11,12,13,32,33,37,38,42,43,49,50,
  61,62,69,70 (SPLICE_MAP in `_splice_media.py`).

### Verification done
- Every scripted slide render-checked via PowerPoint COM PNG exports
  (several rounds of layout fixes applied).
- Deck opens clean in PowerPoint; 76 slides.
- Animation click structure verified via COM MainSequence TriggerType:
  ALL 53 animated slides match the plan beat-for-beat.
- Full-screen SLIDESHOW probe (screenClass PrintWindow captures on
  slides 1, 11, 13, 45, 76): PASS — live poll, Excel chart, and builds
  render; no "slide failed to open" banner.

### Content flags for Nico (reported in chat, awaiting his eyeball)
1. Slide 20 (aggregation): rebuilt with clean numbers (C1: 12→1 … 3→4;
   C2 shifted +1; aggregate = horizontal sum). The original dot values
   weren't recoverable from the XML geometry; economics (horizontal
   summation) preserved.
2. Slide 65: Target/Walmart stock series is an approximate digitization;
   the source line says "series approximate".
3. Old slide 66 ("Solution" popcorn) contained a stray word "Calculus" —
   not carried over.
4. Slide 23's network-effects screenshot (StudiVZ/Facebook, ~2007) is
   dated — possible refresh candidate, his call.
5. Slide 56's elasticity-estimates table is still the original screenshot
   (image53.png) — could be rebuilt as a native table later.

### Grouping pass (added 2026-08-15, approved by Nico)
`_group_pass.py` — geometric detection, zip+lxml surgery, spliced slides
skipped. 16 groups: 8 box+text callout pairs (slides 19, 40, 45, 46, 47,
52, 54, 76), 2 shade+graphicFrame pairs (27 Netflix chart, 73 cereal
table), 6 picture+caption groups (17 Inglehart, 64 ×3, 67 ×2 incl.
multi-picture captions). Slide 57's gas caption did not meet the
adjacency heuristic and stays ungrouped (co-reveals via its beat).
Animation plans updated to grp: selectors on the affected slides; full
pipeline re-run; click counts re-verified (all 53 match) and the
slideshow probe re-passed.

### Hand-edit port round (2026-08-15, second session)
Nico hand-edited slides 9 and 10 in the canonical deck (preserved in
`Module 2 - In Class Revised_backup_2026-08-15.pptx`) and gave format
instructions; all ported into the build script:
- **Slide 6:** course roadmap rebuilt in the Module-3 standard format
  (diamond layout, M3 wording "Basic Principles and Economic Way of
  Thinking", module 2 navy, gold up-arrow + "we are here" beneath it).
- **Slide 7:** descriptive outline redone in the CT format (gold number
  circles, bold title, gray description underneath). Slide 8 kept as-is.
- **Slide 9 (+ bookend 14):** his restructure ported — section headers
  ("Crucial assumption:", "The Law of Demand says:", "Reasons:") are
  unbulleted flush-left; NEW white rounded outlined boxes around the
  Assumption and Law-of-Demand sections (auto-grouped with their text by
  _group_pass; 21 groups deck-wide now). Animation replan: assumption
  box static, law box + D-curve one beat, reasons build (3 clicks).
- **Slide 10:** his geometry ported (bullets at y 1.495; pizza enlarged
  to 5.25" wide at (3.77, 3.05)).
- **Poll Break badge:** his hand-tuned badge (smaller gold parallelogram
  + navy 28 pt label, grouped, bottom-right overlapping the footer,
  IN FRONT) saved verbatim as `_handoff_pollbreak.xml` (BUILD INPUT —
  never delete) and injected via `_inject_handoff_group` AFTER the
  footer on all 7 poll-setup slides (10, 32, 37, 42, 49, 61, 69),
  replacing `_add_discussion_break`.
Verified: click counts match on all 53 animated slides (slide 9 now 3
clicks), renders checked, slideshow probe re-run.

### Round-3 feedback (2026-08-16, all implemented)
- **Slide numbers throughout:** every spliced slide (polls + Excel) now
  gets a live slide-number field injected by `_splice_media.py` (same
  look/position as the built footer); slide 7's missing footer fixed
  (the descriptions branch returned before drawing it).
- **Slide 7:** slide-8 font/format (30 pt bold navy numbered items,
  nothing shaded) + gray description line under each item.
- **Slide 13:** spliced title had NO xfrm (inherited the stub layout's
  default position) — pinned to the standard action-title position,
  30 pt bold navy, left-aligned; the old sldNum placeholder removed.
- **Slide 17:** Gates quote enlarged to 28 pt (attribution 24 pt).
- **Slide 19:** redesigned — each major block in its own white rounded
  shaded box (auto-grouped), gold MB = MC anchor star; build: recall
  box static → decision box → star (2 clicks). 23 groups deck-wide.
- **Slide 20:** MPV curve 4 pt; Q* computed as the TRUE Bézier/MC
  intersection (bisection in the build script); callout arrow ends ON
  the curve (Bézier-evaluated point). Slide 16's dashed guide now runs
  exactly through the bar-top values (1,12)→(5,0). Other chart slides
  audited: aggregation dots/lines, unit-elastic midpoint, special
  cases all exact.
- **Teaching CLAUDE.md:** new standing rule added (authorized by Nico):
  "Curves must be economically exact, not just suggestive" — marked
  intersections computed, dots on lines, arrows end on curves.
- Verified: renders of all changed slides; click spot-checks pass;
  slideshow probe PASS (slides 11/13/77 navigated in the real show).
  Note: a probe run right after repeated POWERPNT force-kills can fail
  with COM RPC errors from PowerPoint's dirty crash-recovery state —
  retry cleanly before suspecting the deck.

### Pending / next steps
- Speaker notes: substantive originals preserved verbatim (15, 16, 17,
  21, 26, 28, 52, 73 + poll notes via splice); NEW slides carry drafted
  2–4-sentence notes. No teleprompter pass was requested for this deck.
- "Module 2 - Video Part.pptx" is a separate, later task.
- Not committed to git yet (Nico confirms at session end).

### Gotchas learned this session
- **`<a:xfrm>` child order in a grpSp is off, ext, chOff, chExt.** Any
  other order is silently misparsed: group children collapse (text goes
  vertical one-letter-per-line, pictures/charts vanish). No error is
  raised — caught only by render comparison.
- CT's poll slides are static images and her deck has no animations —
  content reference only.
- OMML equation shapes DO expose their m:t text to the animation engine's
  shape collector — select them with `t:` prefixes ("t:E D"), not `osp:`.
- `t:` selectors containing `#` (e.g. "#Note") collide with the `#n`
  suffix syntax — write "t:#Note#1".
- `_add_convention_box` emits TWO shapes (rounded rect + text box) — put
  both in the same animation beat (and group them in the grouping pass).
- PowerShell COM: `$pp.Visible = -1` (MsoTriState), not `$true`; the
  slideshow window is found reliably via EnumWindows on the POWERPNT pid
  + class `screenClass` (FindWindow raced/failed; `SlideShowWindow.HWND`
  returns null through PS interop).
