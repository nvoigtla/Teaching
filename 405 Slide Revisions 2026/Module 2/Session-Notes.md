# Session Notes — Module 2 (In-Class + Video Part decks)

## 2026-08-23 (third pass) — Speaker notes on every slide

**One-line summary.** Every one of the 76 slides now carries speaker
notes; nothing that was already substantive was overwritten, and the
PollEverywhere notes were left untouched.

### What was written
- **41 slides had no notes at all** and got new ones: 1, 2, 3, 6, 7, 8,
  9, 10, 13, 14, 19, 20, 23, 24, 25, 30, 31, 34, 36, 39, 40, 46, 47, 48,
  51, 53, 54, 55, 56, 58, 59, 60, 63, 64, 65, 67, 68, 71, 74, 75, 76.
- **5 slides had a stub** (a citation, a URL, a one-liner) and were
  expanded, with the original line kept at the end of the note: 17
  (Inglehart citation), 27 (Netflix / Practice Video 1 pointer), 35
  (YouTube link), 41 (WSJ Amazon–Hachette URL), 44 ("Rounding to the
  closest integer").
- **14 slides kept their notes verbatim** — the substantive ones ported
  from Nico's original deck plus the ones drafted earlier: 15, 16, 18,
  21, 22, 26, 28, 29, 45, 52, 57, 66, 72, 73.
- **16 PollEverywhere slides untouched** (4, 5, 11, 12, 32, 33, 37, 38,
  42, 43, 49, 50, 61, 62, 69, 70). Their notes ARE the poll mechanism —
  the add-in reads the poll URL out of them at slideshow start, and a
  rewritten notes part crashes the renderer deck-wide.

Style follows the Teaching CLAUDE.md default: 2–4 sentences in spoken
voice, the concrete example named (Gjelina pizza, LADWP water at −0.4,
CorePower Yoga, Amazon vs. Hachette, Uber, Rivian R3, Target vs.
Walmart, movie tickets and popcorn), and the hand-off to the next slide.
Worked slides carry the full arithmetic so the notes stand alone as
student guidance when the deck is uploaded.

### Mechanics
- New BUILD INPUT `_notes_m2.py` holds `NOTES` (display number → text)
  and `SPLICED_NOTES`. `build()` applies `NOTES` at the very end, only to
  slides that do not already set notes of their own, so the per-slide
  `_set_notes` calls stay authoritative.
- **Slide 13 needs the splice route.** It is spliced in from the original
  deck, so its notes part is replaced wholesale; `_splice_media.py` now
  has `_with_notes_text()`, which writes `SPLICED_NOTES[disp]` into the
  copied notes part's body placeholder. Poll slides never go through it.
- Gotchas hit on the way: `xml.etree`'s `tostring()` has no `standalone`
  keyword (lxml does), and — the documented one — **a bash heredoc eats
  one backslash level**, so `

` inside the patch script arrived as
  real newlines and broke five string literals. Write patch scripts to a
  .py file with the Write tool, as the Teaching CLAUDE.md says.

### Verification
- Notes audit: all 76 slides carry notes; PowerPoint COM confirms every
  slide has ≥ 40 characters on its notes page.
- Member-level geometry diff vs. the previous deck: 46 slides differ and
  **every one of them differs only in NOTES** — no geometry moved.
- Click-by-click timing diff: **all 76 slides identical**.
- Deck opens clean, 76 slides; slideshow probe on 1/4/13/22/32/49/76:
  PASS, with all three sampled live poll slides still rendering.


## 2026-08-23 (later) — Second hand-edit round: slides 16, 19, 21, 22, 23

**One-line summary.** Adopted Nico's second pass of hand-edits (16, 19,
21, 23), added the snob-effect explainer box to 22, and wrote the new
upward-sloping-MC rule into the Teaching CLAUDE.md.

### Adopted hand-edits
- **Slide 16 (Gates).** The last leftover effect is gone — the slide is
  now fully static (moved from PLANS into `SKIP_STATIC`).
- **Slide 19 (optimal movies).** THE substantive change: the MC line is
  now **upward-sloping**, from fig (0.4173, 0.4884) to (8.3778, 4.3360),
  because the curve is labelled "incl. opportunity cost". The MC label
  moved to (8.676, 4.322). The "MPV is the demand curve" callout was
  resized to 3.800 x 1.028 at (5.500, 2.870) (text padding scaled to
  0.173 / 0.108) and **grouped with the gold arrow** that points at the
  curve, so the two reveal as one beat. New build order: MPV curve +
  labels → the callout group → MC + its label → Q*.
  **Q* recomputed** as the true Bézier/MC intersection against the sloped
  line (q* = 4.8234, x = 6.441"). Nico's hand-placed guide was at 6.468",
  so the drop line and the "Q*" label each shifted 0.027" left — below
  visual threshold, and required by the "curves must be economically
  exact" rule.
- **Slide 21 (factors affecting demand).** The demand-shift figure was
  redrawn and moved up beside the bullets: `SimpleFig(6.876, 4.003, 2.6,
  2.2, 10, 10)`; base D now (1.4769, 9.0) → (8.2, 1.6); the two shifted
  curves keep their dashed style but each got a short **diagonal** gold
  arrow off the base curve; labels re-worded to "Rising demand" /
  "Falling demand" and moved beside their own curve (11 pt italic gray,
  boxes 1.600 x 0.185 and 1.126 x 0.185). Cartoon moved to (10.400,
  3.831); "Anything else?" to (4.066, 6.473). Each shifted curve + arrow
  + label is ONE group. New 10-click build: falling panel → rising panel
  → then the bullets one at a time (p0 now animates too), cartoon riding
  on the Ryanair sub-bullet, "Anything else?" last.
- **Slide 23 (network effects).** Screenshot nudged to (3.200, 3.050).

### New work
- **Slide 22 — snob-effect explainer box.** Cream convention callout
  above the two WSJ clippings: "**Snob effect:** exclusivity is part of
  the value, so demand falls as more people own the good", 18 pt,
  centered, at (1.717, 1.400), 9.900 x 0.520. Both panels shifted down
  0.20" to make room. Revealed on its own first click, then the Ferrari
  panel, then the Birkin panel.
  Two sizing constraints drove the width: the line must not wrap (text
  measured with PIL on Calibri / Calibri Bold = 9.11"), and the box must
  stay **under 10"** or `_group_pass` rule 1 treats it as a layout band
  and refuses to group it with its text.
- **Teaching CLAUDE.md — new standing rule** (Nico's instruction): a
  marginal-cost curve that explicitly includes opportunity cost is drawn
  UPWARD-SLOPING (best alternatives are given up first); a flat MC is
  only correct for out-of-pocket cost alone; and when the slope changes,
  recompute every marked optimum as the true intersection.

### Tooling
- `_group_pass.py`: new `MANUAL_GROUPS_POST` pass that runs AFTER the
  geometric rules and may take an existing `grpSp` as a member — needed
  for slide 19, where the callout group nests inside a group with the
  arrow. `bbox()` now also reads `grpSpPr`.
- **lxml gotcha fixed:** the manual-group matcher tracked consumed
  shapes by `id(element)` while pulling elements fresh from `spTree` each
  pass. lxml frees and RECYCLES proxy ids, so a consumed id spuriously
  matched an untouched shape and slide 21's second group failed to
  resolve. It now snapshots the candidate list once and keeps the
  references.
- `_dump_cxn.py`: prints connector ENDPOINTS (flipH/flipV-aware) in
  rendered inches, descending into groups — the tool that made the
  redrawn slide-21 figure portable back into figure units.

### Verification
- Member-level geometry diff: slides 19, 21, 23 reproduce the hand-edits
  exactly (only PowerPoint spell-check run splits remain, plus the
  deliberate 0.027" Q* correction on 19); 22 differs by the new box and
  the 0.20" panel shift.
- Click-by-click timing diff: **73 of 76 slides identical**; the three
  flagged are 19 and 21 (0.001" rounding in the printed geometry — the
  beat sequences match shape-for-shape) and 22 (the new box beat).
- COM click check: 16 = 0 effects, 19 = 4 clicks, 21 = 10 clicks,
  22 = 3 clicks, 23 = 2 clicks.
- Deck opens clean, 76 slides, 51 animated; slide 13 still 4 shapes with
  the working OLE embed.
- Slideshow probe on 1/13/16/19/21/22/23/32/76: PASS.


## 2026-08-23 — Hand-edit adoption round on In-Class slides 1–20 (deck now 76 slides)

**One-line summary.** Ported every hand-edit Nico made to slides 1–20 of
`Module 2 - In Class Revised.pptx` back into the pipeline (geometry,
text, grouping, animation choreography), fixed the broken slide-13 Excel
embed, replaced the whole pizza on slide 10 with a single slice, added
the MPV/MB note to slide 18, and gave the slide-19 MPV curve a shade.

### Hand-edits found and adopted (canonical numbering)
Surfaced with a member-level geometry diff (`_diff_all.py`, copied from
Module 1) plus a new click-by-click timing diff (`_dump_timing.py` /
`_timing_all.py`) against a side-path rebuild.

- **Bookend slide deleted.** Nico removed the law-of-demand recap that
  had been inserted at display 14 in the CT cross-check round. Deck 77
  → **76 slides**; every pipeline config renumbered back to the
  pre-bookend numbering (build-script page numbers, `SPLICE_MAP`,
  `SPLICED`, `SKIP_*`, `PLANS`). The `_shifted_dict/_shifted_set` block
  in `_animate.py` is gone — the config numbering is live as written
  again, and the `slide_NN_*` function names line up with display
  numbers once more.
- **Slide 9.** "Price" axis label moved to (8.979, 2.285); the whole
  demand-curve mini figure (2 axis connectors + Price + Quantity +
  D-curve + "D") grouped into ONE object. Build re-cut to 3 clicks:
  law-of-demand box → the figure group → all three "Reasons"
  paragraphs on one click.
- **Slide 16 (Gates).** Build reduced to a single click that reveals
  only the oversized quote glyph; the quote text and the portrait are
  static.
- **Slide 17 (Inglehart).** Animation removed entirely — slide 17 added
  to `SKIP_STATIC`.
- **Slide 18.** The Module-1 recall line corrected to "Marginal benefit
  (MB) … = Marginal Cost (MC)"; the decision header now reads "Optimal
  consumption decision: We use “MPV”" with MPV in red. MB = MC star
  grouped with its label. Build order REVERSED vs. the old plan: star
  first, then the MPV rule.
- **Slide 19.** "MPV" curve label moved to (10.700, 6.059); "is" in the
  convention callout is now bold as well as underlined.
- **Slide 20.** Six hand-made groups adopted: each "=" sign grouped with
  the aggregate dot it produces (4×), all four "+" signs as one group,
  and the aggregate legend swatch+label. Build re-cut to 8 clicks —
  consumer 1, consumer 2, the plus signs, then one horizontal sum per
  row, aggregate curve last.

### New work this session
- **Slide 10 — one slice, not a whole pizza.** `_mk_slice.py` (BUILD
  INPUT generator) cuts a 56° wedge out of the same Gjelina photo
  (`_source_images/image14.jpeg`), upsamples 2.5× with an unsharp pass,
  and writes `_source_images/pizza_slice.png` (RGBA, transparent
  background). Placed at (4.890, 3.100), 3.55" wide, `rounded=False`
  (the wedge is not rectangular) with the standard soft shade.
- **Slide 13 — the Excel embed is openable again.** TWO bugs, both in
  `_splice_media.py`:
  1. **Rel-ID remap collision.** The old per-entry
     `slide_xml.replace('"rIdA"', '"rIdB"')` loop clobbered itself: old
     rId6 (the EMF preview) was rewritten to rId4, and the later
     rId4 → rId6 pass rewrote that same string, so the OLE fallback
     `<a:blip>` ended up pointing at the NOTES part. PowerPoint could
     not draw the embed's preview. Now a single-pass `re.sub` over
     `"rIdN"`.
  2. **VML shape left behind by the recentering shift.** `X_SHIFT_EMU`
     moved the OLE `graphicFrame` +1.667" but not the legacy VML shape
     the `p:oleObj spid` points at. With the two no longer coincident,
     PowerPoint rendered the VML as a SEPARATE picture lying on top of
     the OLE frame — that picture is what swallowed the double-click.
     `_shift_vml()` now moves `left:NNpt` by the same 120 pt.
  Also: the blanket `<a:off x>` regex was hitting the spTree's own
  `<p:grpSpPr>` transform. Harmless for the modern renderer (ext = 0)
  but the legacy VML path honours it and shifted the whole slide a
  second time — reset to 0 after the shift.
  Result: slide 13 has 4 shapes (was 5), the Pizza Demand chart renders
  as the OLE object itself, and `OLEFormat.Object` resolves with Edit /
  Open verbs.
- **Slide 18 notes.** Added the MPV-vs-MB note Nico asked for (same
  concept; MPV is MB specific to consumption).
- **Slide 19 MPV curve shade.** `_add_cubic_curve` result now gets
  `_add_drop_shadow(blur 3 pt, dist 2 pt, 40 % alpha)`.

### Tooling added (reusable)
- `_diff_slides.py` / `_diff_all.py` — member-level geometry+text+notes
  diff, canonical vs. `..._test.pptx` (copied from Module 1, retargeted).
- `_dump_timing.py` / `_timing_all.py` — click-by-click choreography
  diff, resolving `spid` → shape signature so it survives rebuilds.
- `_dump1.py`, `_dump_runs.py`, `_rawshape.py` — shape / run / raw-XML
  dumpers used to pin down run-level emphasis changes.
- `_slideshow_probe.ps1` — full-screen slideshow probe (from Module 1).
- `_group_pass.py` gained `MANUAL_GROUPS` (explicit member sets matched
  by rendered inches, reaching connectors too) and a `make_group(...,
  anchor="last")` mode — PowerPoint anchors a new group at the TOPMOST
  member's z-position, and matching that is what made the slide-20
  document order reproduce exactly.
- `_animate.py`'s `todo` range is derived from the slide count instead
  of the hardcoded `range(1, 78)` (it crashed silently after printing,
  leaving the deck un-animated, when the count dropped to 76).

### Verification
- Member-level geometry diff over all 76 slides: the only differences
  vs. Nico's deck are the three intended changes (slide 10 pizza, slide
  13 shape count, slide 18 notes) plus PowerPoint save artifacts
  (spell-check `err="1"` run splits, dropped empty `rPr`, autofit
  height recomputation).
- Click-by-click timing diff: **74 of 76 slides identical**; the two
  flagged are slide 10 (pizza geometry) and slide 16 (our engine emits
  the lone effect as `clickEffect` where PowerPoint left a
  `withEffect` — same on-click behaviour).
- COM click-structure check on 9/16/17/18/19/20 matches the plan.
- Deck opens clean in PowerPoint, 76 slides, 52 animated.
- Full-screen slideshow probe on 1/4/9/10/13/18/19/20/32/76: PASS — both
  live PollEv slides render, the Excel chart renders, no failure banner.

### Open questions for Nico
1. **Slide 16.** The deck carried a single leftover `withEffect` on the
   quote glyph (the text and portrait effects had been deleted). Adopted
   verbatim as a 1-click reveal of the glyph — say the word if the slide
   was meant to be fully static.
2. **Slide 18 build order** now shows the MB = MC star BEFORE the MPV
   rule, which inverts the Teaching CLAUDE.md rule that the star follows
   the concrete rule it abstracts. Adopted as hand-edited.
3. **"Shade" on the MPV curve** read as the deck's standard soft drop
   shadow. If shading the AREA under the curve was meant, say so.


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
