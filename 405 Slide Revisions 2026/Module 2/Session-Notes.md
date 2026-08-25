# Session Notes — Module 2 (In-Class + Video Part decks)

## 2026-08-24 (round 6) — Title case, shades on equation boxes, slide 36/33

**One-line summary.** Slide titles are now title-cased at the one place
titles are drawn, filled equation boxes carry the shade automatically,
video slide 36's table clears the footer, and slide 33 explains what
camelcamelcamel.com is for.

### Title case on every slide title
`_title_case()` wraps `_draw_action_title`, so BOTH decks are covered
from a single choke point and no title string had to be hand-edited.
The pass only ever RAISES a letter — it never lower-cases a word that is
already capitalised — so MR, TR, OLS, WTP, A/B, "Inside Out 2's" and
Nico's own choices survive. Articles, coordinating conjunctions and
short prepositions stay lower unless they open the title, close it, or
follow a colon; hyphenated compounds capitalise both parts.

25 of the 120 titles changed, almost all of them poll and poll-result
stubs: "Poll results: marginal revenue" → "Poll Results: Marginal
Revenue", "Poll: WTP for pizza" → "Poll: WTP for Pizza", "Class demand
curve (live Excel)" → "Class Demand Curve (Live Excel)". Nothing on a
content slide moved.

### Shades on the remaining cream boxes
The gap was `_add_math_equation`, which took `shadow=False` by default
while several call sites passed `fill=CREAM` without asking for a
shade — that is where video 21's three formula boxes and 25's came
from. `shadow` now defaults to `None`, meaning "on when the box is
filled"; `shadow=False` still suppresses it.

`_shade_audit.py` (new, keep) walks a built deck — group children
included — and lists every filled shape with no `outerShdw`. Both decks
now come back clean: what it still reports is the outline circles, the
scatter dots, the chart bars, the pale-gold revenue rectangles and the
deliberately flat DIMMED card on In-Class 58, all of which are meant to
be flat.

### Video slide 36 — the table ran into the footer
17 rows at 16 pt needed ~5.8", which pushed the last row past the
footer rule. Row height is line height plus the two vertical cell
margins, so `_add_styled_table` took a new `margin_v` argument and the
table is now 14 pt with 0.015" margins at (2.950, 1.560), 4.35 x 4.70".
It ends around y 6.45".

### Video slide 33 — what camelcamelcamel.com is doing there
CT's deck carries no note on it, so this one is written from scratch:
the site is a free Amazon price tracker that plots one listing's price
history and marks its all-time high and low. The chart shows a single
product swinging between roughly $35 and $84 over about five months.
The teaching point is that Amazon re-prices continuously rather than
running one clean A/B test, so the price variation is there all the
time — with the caveat, made a few slides later, that it is not
randomized. The note also says the product shot is a Coca-Cola branded
collectible plush keychain and that nothing turns on which product it
is.

### Verification
Both decks rebuilt through the full pipeline. Per-slide click counts
are IDENTICAL to the previous decks on all 57 and all 76 slides — the
title, shade and table changes touched no animation. Slides 21, 36 and
47 rendered and checked; slideshow probes on both decks (video
1/21/33/36/47/57, In-Class 1/13/20/58/62/76) show no failure banner,
and the live PollEverywhere slide still renders inside the show.

### Follow-up (same day): the product on slide 33, agenda title case
The plush on video slide 33 is Pop Mart's blind-box keyring from THE
MONSTERS x Coca-Cola series — the character is Labubu, and the packaging
in the shot reads "THE MONSTERS". Confirmed against retailer listings
(eBay, StockX, Showcase USA). The slide now carries a title-style
caption above the photo, "Pop Mart × Coca-Cola Labubu keychain",
measured in Calibri bold italic 13 pt at 2.93" so it stays on one line
inside a 3.20" box — which is also the widest the caption may be and
still be grouped with its picture (the group pass rejects an
above-caption wider than 1.5x the picture). The photo dropped to
y 1.780 and the price chart shrank to 6.60" wide at y 3.580 to make
room. The caption groups with the photo, so slide 33 still builds in
two clicks.

Agenda ITEM TITLES now take the same title case ("The Law of Demand",
"Demand and Revenue", "Marginal Revenue"); the one-line description
under an item is a sentence and keeps sentence case. Added as a
sub-clause of the title-case rule in the Teaching CLAUDE.md.

### Teaching CLAUDE.md — two rules added
* The reference glyphs are a fixed vocabulary: ✎ always means a
  problem set, ▤ always a teaching note, in every module and both
  decks; route every such box through one helper rather than styling
  call sites by hand.
* Slide titles are set in title case, including the throwaway poll and
  stub titles, and the pass capitalises without ever lower-casing.


## 2026-08-24 (round 5) — Cream-box shade, reference boxes, slide 7 note

**One-line summary.** Every cream callout now carries the soft shade the
Teaching CLAUDE.md always specified, post-work pointers became a single
`_add_reference_box()` with a glyph per kind, and slide 7's note is
adopted exactly as Nico set it.

### Slide 7
- The "[note: if P decreases, Q must increase]" line is its own
  paragraph: no bullet, `marL = 0`, `indent = 0`, **space-before 0**,
  indented with the eight leading spaces he typed.  Built with
  `bullet_style: 'arrow'` + `mar_l: 0` + `indent: 0`, which is the one
  combination in `_add_hierarchical_bullets` that emits exactly that.
- The two cream cards take his size: (0.726 / 6.950, 4.298), h 1.452
  (were at 4.150 with h 2.300).

### Cream callouts are shaded — deck-wide, both decks
`_add_convention_box` set `shadow.inherit = False` and stopped there, so
none of the cream boxes had a shade even though the course CLAUDE.md has
always said "soft drop shadow, so the box reads as a lifted card". It now
calls `_add_drop_shadow`. This touches every convention box in BOTH decks.

### Post-work reference boxes
New shared helper `_add_reference_box(..., kind="ps" | "tn")`:
gold-bordered rounded rect, soft shade, navy bold, with a leading glyph
saying what kind of reference it is —
  * **✎** a problem set (an exercise to work)
  * **▤** a teaching note (something to read)
Applied to Video 21 (both pointers), Video 29 and In-Class 40.

**Video 21** lost the exercise numbers: "Problem Set 2 · #4, #5" is now
just "Problem Set 2".

**Video 29** — the old two-line "✎ Problem Set 2 / On BL under Module 2
Post-Work" box is replaced by the standard reference box at
(9.45, 5.42), 2.60 x 0.50.

### Teaching CLAUDE.md — new standing rule
Added a "Post-work reference box" entry under the layout patterns: the
box style and glyph vocabulary, and two label rules —
  * **name the problem-set NUMBER only, never the exercise numbers**, so
    the slide survives next year's re-numbering;
  * keep the "on BruinLearn under …" line for the wrap-up / post-work
    slide, not on every pointer.

### Verification
- Video deck: 57 slides, 42 animated. In-Class: 76 slides, 51 animated.
  Both open clean.
- Renders of slides 7, 21, 29 (video) and 19, 40 (In-Class) checked by
  eye against Nico's hand-edited version.
- Slideshow probes on both decks (video 1/7/21/29/42/57; In-Class
  1/13/19/40/76): PASS, checked pixel-wise for the failure banner.
- Animation selectors that named the old pointer labels were updated in
  both `_animate.py` and `_animate_video.py`.

### Follow-up (same day): the Problem Set 2 box belongs on slide 42
He wanted the pointer ON the Application slide, not on the outline. So
the reference box is gone from slide 29 and sits at (1.000, 6.520) on
video slide 42, bottom-left, opposite the Poll Break badge. Slide 42
goes from 3 to 4 clicks; nothing else in the deck changed.

**Bug found while doing it.** `_group_pass.py` had ONE module-level
`SPLICED` set holding the In-Class display numbers, and it gates which
slides the pass touches. Running it on the VIDEO deck therefore skipped
video slides 4, 5, 11, 12, 13, 32, 33, 37, 38, 42, 43, 49, 50 - and
slide 4 is exactly where Nico's manual groups live, so `_animate_video`
died on `KeyError: 'grp:0'` unless the run remembered to pass
`--spliced=24`. `SPLICED` is now `SPLICED_BY_DECK`, resolved off `_STEM`
like the other deck-keyed tables. Verified: per-slide click counts are
identical to the previous deck everywhere except slide 42.


## 2026-08-24 (round 4) — Video slides 30-52 adopted from CT 34-60

**One-line summary.** Replaced our demand-estimation block wholesale with
CT's, including CT's poll positions as placeholders. Deck 53 -> **57
slides**.

### The mapping
Our 30-52 (23 slides) became CT 34-60 (27 slides), so the deck grows by
four. Two structural changes inside the block:
- CT folds our "Regression Results I" and "II" into one slide, so
  `v44_regression2` dropped out of `build_video()`.
- Our extra "Obtain Elasticity from Estimated Demand Curve" slide has no
  CT counterpart; CT carries that formula in the cream box at the foot of
  its OLS slide instead, so ours was dropped.

### Rebuilt from CT
30 = CT 34 · 33 = CT 37 (product shot above the price history) · 34 =
CT 38 · 35 = CT 39 (navy equation bar + the elasticity in a cream box) ·
36 = CT 40 (table left, aircraft beside it) · 37 = CT 41 (scatter with
two candidate lines) · 38 = CT 42 (fitted line, two residuals marked,
algorithm box) · 39 = CT 43 · 40 = CT 44 · 41 = CT 45 · 42 = CT 46 ·
45 = CT 49 · 48 = CT 52 · 51 = CT 55 · 52 = CT 56 · 53 = CT 57 (the wine
example replaces our coffee headline) · 54 = CT 58 · 55 = CT 59 (CT's two
Tyler Vigen charts) · 56 = CT 60.

`_scatter_fig()` was re-derived on CT's geometry: the x-axis starts at
P = 210 and Q runs 0-150, so P = 220 lands at x 4.415" and Q = 150 at
y 2.300", exactly as on CT's slides 41/42. Dots are gold with a navy
edge, as CT has them.

### Poll placeholders
CT runs three polls in this block (its 47/48, 50/51, 53/54). Those live
on CT's PollEverywhere account and are never spliced (project
CLAUDE.md), so they go in as OUR stubs at 43/44, 46/47 and 49/50 —
markers showing Nico exactly where to insert his own activities. All six
are in `SKIP_MEDIA`.

### Format adjustments (the ones our rules require)
Our chrome throughout (navy top bar with the Module 2 · Video 3 tag, our
action-title position, our footer). CT's equations that are built from
piles of little text boxes — the elasticity on its slide 39, the
multivariate equation on 44 — are native OMML here. Filled boxes are
rounded with a soft shade.

### New build inputs
`ct_airline_plane.png`, `ct_s57_image19/20.png`, `ct_s58_image19/20/21.png`,
`ct_s59_image22/23.png`, all extracted from CT's deck into
`_source_images_video/`.

### Verification
- Deck opens clean: **57 slides, 42 animated**.
- Every rebuilt slide rendered and compared side by side against its CT
  original, in five batches.
- Slideshow probe on 1/30/36/39/42/43/45/48/51/53/55/57: PASS, checked
  pixel-wise for the failure banner.

### Caught on the way
- `_add_styled_table`'s `col_widths` takes EMU, not inches — passing
  inches collapses the columns and stacks the text vertically.
- `image12.png` in `_source_images_video/` is from Nico's ORIGINAL deck,
  not CT's; CT's slide-40 image of the same name is a different picture.
  Extracted separately as `ct_airline_plane.png`.


## 2026-08-24 (round 3) — Video slides 3-7, 19, 21, 22

**One-line summary.** Adopted the hand-edits on slides 3 and 4 (including
grouping), made slide 4's three points separately animatable, added the
movement arrows on slide 5, and handled the smaller asks on 7, 19, 21
and 22. **The CT 30-52 block is NOT started** — see below.

### Hand-edits adopted
- **s3** retitled "Plotting the (Inverse) Demand Curve"; the cream box now
  spells out the rearrangement step with the inverse demand set bold.
- **s4** the two figure groups Nico made by hand are reproduced —
  {both axes + P + Q} as one object, and {demand line + D label + the
  $400 / 1600 ticks} as a nested pair — via a new Video entry in
  `MANUAL_GROUPS_BY_DECK`. "D" moved to (5.476, 3.471); the box moved to
  (7.923, 2.500) 5.030 x 3.040.
- **s7** the "[note: if P decreases, Q must increase]" line moved onto
  its own line under the question.
- **s21** step 3 reworded to "Compute Marginal Revenue from total
  revenue".

### New work
- **s4 — the three points animate one at a time.** The cream rectangle
  is now a plain background shape and the three points live in their own
  text box, because a box merged with its text can only animate as a
  single object. `_group_pass` has a new `NO_GROUP_BOXES` list that tells
  rule 1 to leave this pair alone. Five paragraphs (three bulleted, two
  un-bulleted continuations at level 1), revealed as three beats:
  `pr:Inverse demand:0:0`, `1:2`, `3:4`.
  **Gotcha:** a literal newline inside a run is read by PowerPoint as a
  PARAGRAPH break, not a line break — the continuation lines were each
  picking up their own bullet until they were made real paragraphs.
- **s5** gold down-arrow to the left of P0 / P1 and gold right-arrow
  under Q0 / Q1; the gold box now reads "Which area is bigger?".
- **s19** the arrow pointing at "Derivative" is now the same concept
  blue as the word.
- **s21** the single pointer became two boxes: "Problem Set 2 · #4, #5"
  and "Teaching Note: Marginal Revenue".
- **s22** the two braces are back over the demand curve, one across the
  elastic stretch and one across the inelastic stretch, from Nico's own
  Module 2 Video 2 slide 8. New `_brace_along()` helper lays a rightBrace
  along an arbitrary segment: rotate by (theta + 90) so it curls DOWN
  onto the line, then lift it clear along the upward normal.

### Verification
- Deck opens clean: **53 slides, 44 animated**.
- Slide 4 = 5 clicks (two figure groups, then one per point).
- Slideshow probe on 1/3/4/5/7/19/21/22/24/53: PASS, checked for the
  failure banner pixel-wise as well as by eye.
- Renders of every changed slide compared against Nico's hand-edited
  version and against his originals.

### NOT DONE — the CT 30-52 block
Nico also asked for our slides 30-52 to be replaced wholesale by CT video
slides 34-60 ("keep the location, font size, animations, grouping, all
exactly as in CT"), including CT's PollEverywhere slides as placeholders.
That is ~27 slides of pixel-faithful adoption and grows the deck by four;
it is a separate piece of work and has not been started. Everything else
from his 2026-08-24 round-3 list is done and verified.


## 2026-08-24 (round 2) — Video deck: hand-edits + slides 7, 18, 20-23, 27

**One-line summary.** Ported Nico's hand-edits on video slides 1-27,
reinstated his own slide 7, adopted CT 21 on slide 18, relaid out the
calculus slide, folded CT 23's sub-lines into slide 21 and deleted the
now-redundant slide 22 (deck 54 -> 53 slides).

### Hand-edits ported
- **s2** teaching-note link box lifted to (8.200, 6.272).
- **s3 / s4** cream boxes made taller and moved up (3.20 -> 1.95 with
  h 3.25; 2.40 -> 1.38 with h 4.20).
- **s6** TR label to (4.165, 4.775); both region arrows and both region
  note boxes repositioned and narrowed.
- **s7** "price decrease" in red, and P / Q italic inside the questions —
  folded into the reinstated slide (below).
- **s8** retitled "What can a Price Cut and Revenues tell us about the
  Elasticity?"; the bullet now says *revenues* rather than *sales*, and
  the speaker note follows. NOTE: his note edit read "sales reveunes";
  ported as "sales revenues" (unambiguous typo).
- **s19** "increases by 1 unit" -> "by one unit".
- **s21** each step label gained an un-bolded parenthetical.

### New work
- **Slide 7 — Nico's own slide reinstated.** Replaced the CT version
  adopted last round with his Module 2 Video 1 slide: TR = P · Q and
  %ΔTR = %ΔP + %ΔQ as native OMML, the "to assess a price decrease"
  line, and the two cream question cards. This reverses last round's
  CT adoption for slide 7 only; 3-6 stay CT.
- **Slide 18 = CT video slide 21.** The two rules moved out of the
  sub-bullets and onto their own navy bars. Kept close to CT except the
  bars are rounded with a soft shade, per the deck's filled-box rule.
- **Slide 20 — calculus refresher relaid out.** Was crammed into the top
  left; now a centred definition card, then "The general rule" and
  "A worked example" as two columns, each formula on its own line.
- **Slide 21 — CT 23's sub-lines folded in.** Under each bold step label
  sits CT's explanation ("Rearrange demand so P is a function of Q",
  etc.), with the formula boxes unchanged on the right, plus CT's
  Problem-Set pointer. Row pitch tightened to 1.40" so the pointer
  clears step 3.
- **Slide 22 deleted** (the "3-Step Method: Summary Notes" recap) now
  that its content lives on 21. Everything from 23 on shifted down one;
  splice map 25 -> 24, SKIP sets and all PLANS keys shifted.
- **Slide 23 (now 22)** — inverse demand line, its D label and the
  E_D = −1 label all navy, matching the inverse demand on slide 6; the
  elastic / inelastic portion labels take CT slide 26's 16 pt bold
  concept-blue treatment and sit in clear zones above the demand line.
- **Slide 27 (now 26)** — Poll Break badge instead of Group Discussion.

### Verification
- Video deck opens clean: **53 slides, 44 animated**.
- Geometry diff over slides 1-21: only the intended rebuilds (7, 18) and
  PowerPoint save artifacts (title run splits, autofit heights inside the
  cream boxes, `i` attribute normalisation).
- Slideshow probe on 1/7/18/20/21/22/24/26/53: PASS.
- In-Class deck rebuilt and diffed to confirm the shared-layer fixes
  (`_add_cubic_curve`, `is_chrome`, subscripts, links) leave it
  untouched: **0 slides differ, all 76 timings identical**.

### Still open
- Three slides continue to carry the "Group Discussion" badge (video 14,
  16, 17 in the new numbering). Nico has now named slide 8 and slide 27
  individually; the rest were left alone.
- The PollEverywhere slide (video 24) still renders "Activity not found"
  in the live show — an account-side issue, not a deck defect.


## 2026-08-24 — Agenda restructure (both decks) + CT adoption in the Video deck

**One-line summary.** Reworked the module outline on both decks (3a / 3b
sub-items, non-current items shaded), then rebuilt Video slides 3-7 from
CT's slides 4-8, dropped the Wrigley example, and ran two deck-wide
passes (real subscripts, CT source links).

### Agenda / outline — BOTH decks
- `M2_OUTLINE` now carries `(label, title, description, is_sub)`.
  "Elasticity and revenue" and "Marginal revenue" became sub-items **3a**
  and **3b** under "Demand and revenue", and demand estimation moved up
  to **4** — the structure of Nico's original deck, where those two sit
  at outline level 1. LIST INDICES are unchanged, so every
  `highlight_idx` / `highlight_set` call site kept working untouched.
- Sub-items render indented: circle Ø 0.46" at x 1.62 with a 17 pt
  label, title at 22 pt, and a correspondingly narrower cream band.
- **Dimming**: on a section agenda every item that is not the current one
  is now `#BFBFBF` (circle digit + title); the gold circle fill stays
  gold. Follows the Module-1 rule already written into the Teaching
  CLAUDE.md, so the descriptive overview and the summary closer — which
  have no current topic — keep every item navy.

### Video deck
- **Wrigley example deleted** (build displays 8-10: the chewing-gum
  setup, its poll and its solution). Deck 57 -> **54 slides**; page
  numbers, `SPLICE_MAP` (only 25 <- 20 left), `SKIP_*` and every PLANS
  key shifted down by three.
- **Slides 3-7 rebuilt from CT slides 4-8** (Nico's call: CT content in
  our chrome).
  - 3 = CT 4 "Plotting the Demand Curve": the firm note became an italic
    caption above the graph and the red bottom line went away. NOTE: this
    supersedes the earlier "make the bottom line a shaded box" request —
    Nico chose CT over his own edit when the conflict was surfaced.
  - 4 = CT 5 "From Demand to Total Revenue": the shaded P x Q rectangle
    with "Total revenue" inside it, and the full Demand -> Inverse ->
    TR = P·Q -> substitute box with the result in gold.
  - 5 = CT 6: base revenue box, the red slice lost to the lower price,
    the green slice gained on volume, a two-entry legend and the gold
    "Which is bigger?" box. All four corners sit ON the demand line.
  - 6 = CT 7 PLUS Nico's own video slide 6: the two rotated gold braces
    over the elastic / inelastic stretches, the unit-elastic midpoint,
    and — his addition — the rising / falling arrows with the region
    descriptions inside the bottom TR graph.
  - 7 = CT 8 "To Assess a Price Change, Ask About Elasticity": two
    gold-bordered cards.
  - New animation plans for all five.
- **Slide 8** now uses the deck-standard **Poll Break** badge instead of
  the "Group Discussion" relabel.

### Deck-wide passes (both decks)
- **`apply_subscripts()`** — every faked index becomes a real PowerPoint
  subscript (`baseline="-25000"`, what PowerPoint and CT both write).
  Fixes two defects: Unicode lookalikes typed inline (the small-capital D
  in E_D, subscript digits in P_0 / Q_1) and an index faked with a
  smaller font (what slide 9 was doing). 21 subscript runs in the Video
  deck, 19 in the In-Class deck; zero lookalikes left in either.
- **`apply_ct_source_links()`** — CT's own source URLs restored on the
  runs we adopted, keyed by run text so they survive renumbering:
  "Source: Pharmaceutical Technology" (slide 8), "Novo Nordisk shares
  tumbled ~18%" (10), "camelcamelcamel.com" (34),
  "tylervigen.com/spurious-correlations" (52). Checked CT's In-Class
  deck too — it carries no external links at all, so nothing is missing
  there.

### Two real bugs found and fixed
- **`_add_cubic_curve` emitted `<a:effectLst/>` BEFORE `<a:ln>`.**
  `shadow.inherit = False` appends the empty effectLst, and the ln built
  afterwards then violates schema order, so PowerPoint silently drops the
  whole line style and falls back to the theme line — thin and light
  blue. That is why the Video deck's TR parabolas had never rendered
  gold. The In-Class deck's MPV curve was accidentally immune because
  `_add_drop_shadow` re-appends the effectLst at the end.
- **`is_chrome()` was eating freeform curves.** A textless `sp` wider
  than 4" and taller than 2.5" is treated as a white chart backing;
  video slide 6's TR hill has a Bezier control-point bbox of 5.49 x 2.53,
  so it was being dropped from every animation plan. Now guarded on
  `custGeom`.
- Also scoped `_group_pass.py`'s MANUAL_GROUPS **by deck** — the pass is
  shared, and In-Class 9/18/19/20/21 are entirely different slides from
  Video 9/18/19/20/21, so the video build was picking up the In-Class
  groupings.

### Verification
- Both decks open clean: Video 54 slides / 45 animated, In-Class 76 / 51.
- In-Class geometry diff vs. the committed deck: 12 slides, all of them
  the new outline (7, 8, 25, 75, 76) or the subscript conversions.
- Video slideshow probe on 1/3/4/5/6/7/8/25/54: PASS.
- Renders of every rebuilt slide checked against the CT original.

### Open / flagged
- **Four more slides still carry the "Group Discussion" badge** (Video
  15, 17, 18, 30 in the new numbering). Nico named only slide 8, so the
  others were left alone — say the word to switch them too.
- **The PollEverywhere slide (Video 25) renders "Activity not found"** in
  the live slideshow. The splice is intact; the activity itself looks
  deleted or renamed in the PollEv account.
- The subscript pass sets the baseline but does NOT force the base letter
  italic, so the existing look of E_D is preserved. The Teaching
  CLAUDE.md rule for indexed symbols also asks for an italic base — worth
  deciding whether to apply that deck-wide.


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
