# Session Notes — Module 1 (combined In-Class + Videos deck)

## PENDING (updated 2026-08-23 — deck is 99 slides)

Nico's open requests on **Module 1 - Revised.pptx**:
0. **Podcast URL for display 12** — the Sound button is a marker until
   he supplies the link (one line in `_build_Module1.py`).
0b. **Display 31's speaker note mentions lithium, not tea** — the WSJ
   lithium link was kept underneath the new tea notes; it looks like it
   drifted onto the wrong slide. Awaiting his word.
0c. **Slide 82** — the AI-accelerator image has no attribution line;
   add one if the source needs crediting. Title ("AI and the Demand
   for Computer Chips") and the dropped In-Class-Discussion badge are
   mine, not his.
1. **Font-upsize rule, first batch** (then Nico reviews; only after
   convergence add the rule to Teaching CLAUDE.md — NOT yet):
   whenever bullet font < 28pt and space allows, enlarge up to 28pt
   main / 24pt sub. Proposed fitting rule: for each bullet box try
   (28/24) → (26/24) → (24/22), keep the largest pair whose PIL-
   measured wrapped height (Calibri, with 12/3pt spacing-before)
   fits the box height with ~5% headroom and no line >2 lines;
   sibling boxes on a slide get equal sizes; slide 6 must fit
   Nico's hand-set box (2.21"/4.13"). Nico's "slides 1–10" refer to
   the pre-insert numbering — that's now displays 1–12 (7–8 are
   spliced polls, nothing to resize there). NOTE: the click-overlay
   hotspots on slides 2 and 9 (`wire_backup_links`) are positioned
   for the CURRENT font sizes — recheck them after the font pass.
2. **Economic Models full-bleed rebuild — now display 10** (was 8):
   integrate Nico's background image (hiker+map+mountains, text-free
   variant) as background; overlay EDITABLE text exactly like his
   mock: 3 bullet lines top-left (bold navy key phrases), navy circle
   icons per line (head-bulb / quotes / map), "THE REAL WORLD" and
   "THE MODEL (A MAP)" labels right, navy takeaway bar bottom-right
   with compass icon: "Find a model that " white + "matches your
   needs." gold. Images saved as `Image_Hiker_Mountain_with_text` /
   `_no_text` (in `_source_images/`, .png). Update the slide's
   animation plan afterward (PLANS key 8 pre-shift → 10).

## 2026-08-23 — link symbols, hand-edit ports, subscripts, outlines, notes

**One-line summary.** A long iterative pass over `Module 1 - Revised.pptx`
(still 99 slides): replaced the gold ▶ link glyphs with PowerPoint action
buttons, ported ~15 rounds of Nico's hand-edits, darkened the deck's green
everywhere, converted all nine outline slides to the Module 2
numbered-circle format, added a deck-wide symbol-subscript pass, replaced
the COVID/flour slide with an AI-and-chips example, and gave every one of
the 99 slides speaker notes.

### Deck changes, in the order they were asked for
- **Link symbols.** The gold ▶ / ➜ glyphs are gone deck-wide. Backup jumps
  use `actionButtonEnd` (navy face, white glyph, 0.434 × 0.210" — 30%
  smaller than the first cut); external links use the same family keyed to
  what they open: `actionButtonSound` (podcast, display 12),
  `actionButtonDocument` (Economist article, 94), `actionButtonMovie`
  (econimate video, 94), all navy. The back button reverted to the original
  plain navy "← Back" pill after a detour — see the CLAUDE.md rule.
- **Invisible click overlays removed.** Transparent rects over the bullet
  boxes on displays 2 and 9 were why Nico could not select the text; the
  action button is now the click target itself.
- **Hand-edits ported** (all with dated comments in the build script):
  slide 1 comic raised to y 162547 EMU; display 11 third map enlarged;
  13 fox/hedgehog shrunk and stacked; 15 Einstein sub-bullet deleted and
  the block re-centred; 23 both figures raised clear of the footer; 26
  arrow + D′ label moved; 78 definition callout moved up; 79 cones panel
  grouped + his 3-click build; 81 D′ label / ii) arrow / ii) label moved,
  a new horizontal dashed segment, two groups, his 4-click build; 84 the
  P1↔P2 / Q1↔Q2 relabelling (the movement along S starts at the LOWER
  price — the build had it inverted), i)/ii) repositioned, a new dashed
  segment, four groups, his 2-click build; 89/90 arrow and label positions;
  92 reworded and hidden.
- **Slide 36 (avocados).** Rebuilt to Nico's original 10-click
  choreography from `Module 1 - In Class.pptx` slide 30, with his grouping
  of curves and labels. Fixed a real defect: the two shift arrows were on
  the WRONG beats (S→S1 fired with the demand shift, S→S2 with S1).
- **Green.** `#00B050` is gone from both decks; `GREEN_DK = #007A33` is the
  only green left. `GREEN_BR` and `GREEN_MB` were retired outright.
- **Symbols.** `apply_symbol_subscripts()` is a deck-wide build pass that
  splits any P/Q/D/S symbol followed by an index into an italic base run
  plus a true subscript run (66 paragraphs). Keyed to those four letters
  so ordinary text is untouched.
- **Slide 82** replaces the COVID/flour example: AI and the demand for
  chips, Nico's `AI_Accelerator_Chips` image, D → D′ outward shift, no
  supply curve (Video 3 has not reached equilibrium yet) and no Q2 guides
  (they would suggest the price stays constant).
- **Outline slides.** All nine now use the Module 2 numbered-circle format
  via a `make_m1_outline` copied from `make_m2_outline`, over a new
  6-item `M1_OUTLINE`. Item rows are pixel-identical across slides
  (rows from y 1.635", pitch 0.910"); bands all (0.900", 12.150 × 0.900").
- **Speaker notes.** `FILL_NOTES` (58 entries) + `apply_fill_notes()` fills
  every slide that had none. It never overwrites existing notes, which is
  what keeps the source-ported notes and — critically — the eight
  PollEverywhere payload notes intact. Verified byte-for-byte.

### Toolchain changes worth knowing
- **Shape names drive grouping and animation now.** `_sd_chart` and the
  hand-built charts name their shapes (`sdcurve:D`, `sdguide:h:Q3`,
  `sdarrow:ii`, `sdpic:chips`, …); `_group_pass.py` rule 5 pairs them by
  name from `CHART_GROUPS` and names the group `sdgroup:<key>`;
  `_animate.py` gained an `n:<name>` selector. This replaced a
  nearest-connector heuristic that silently grouped slide 36's "D" label
  with the *S2* curve — names removed that whole class of error.
- `_group_pass.py` also gained rule 4 (label + link button) and a
  width guard so outline bands (12.15") are not treated as callouts,
  while the 10.5" "Important" box on display 92 still groups.
- `_animate.py`: `t:`/`pr:` selectors now match a concatenated-run variant
  as well, because the subscript split turned "P0" into "P 0".
- Path-independence fixes: `_diff_slides.py`, `_verify_anim.ps1`,
  `_slideshow_probe.ps1` and `_export_probe.ps1` all take a deck argument
  and use `$PSScriptRoot`.
- New audit helpers: `_diff_all.py` (full-deck member-level hand-edit
  diff), `_check_jumps.py`, `_scan_glyphs.py`, `_audit_overlays.py`,
  `_audit_notes.py`, `_extract_timing.py` (pull a slide's click structure
  out of any deck — this is how his choreography was adopted),
  `_shape_idx.py`, `_dump_edits.py`, `_sheet_probe.py`, `_crop_probe.py`.

### Rules added to `Teaching/CLAUDE.md` (all at his request)
1. Overlay coverage — two figures should not overlap at all where the slide
   has room; where a build paints successive versions of the SAME figure,
   the later one must fully contain the earlier.
2. A backup link always sits in the lower-right corner; a podcast/article
   link goes wherever it fits best.
3. Indexed symbols get a real subscript, italic letter.

### Verification, every round
`_verify_anim.ps1` (COM click-count check, all 65 animated slides) +
`_check_jumps.py` + a full-screen `_slideshow_probe.ps1` run including a
live PollEv slide. Hand-edits were always captured with `_diff_all.py`
against a side-path build BEFORE rebuilding.

## 2026-08-22 — backup section + missing poll slides (deck 87 → 99)

Nico uploaded **"Module 1 - In Class with Solutions.pptx"** (68 slides,
his FEMBA variant) and asked to add everything the rebuild had missed.
Findings: all Solution slides were already adopted (AC #26, Swift #30,
flip-house #51) and the Kleven child-penalty chart too (#54). Missing
were (a) FIVE PollEv slides — a new "Econ & Coffee weekend slot" poll
pair after Office Hours, plus a results-view slide per existing poll
(each its own __PE_POLL_EMBED_ID) — and (b) the 7-slide BACKUP section
with its jump links. All added; pipeline rerun; verified.

New display map (old-87 → new-99): 1–6 same; 7–22 → +2; 23–25 → +3;
26–45 → +4; 46–87 → +5. New slides: 7/8 (Econ&Coffee poll pair),
25/29/50 (results views for AC / diamonds / flip polls), 93–99 =
BACKUP divider, National Leaders (Econometrica 2025), Money-Buy-
Happiness (Easterlin), Stevenson-Wolfers 2008, Anderson Faculty
(HIDDEN, as in source), Portland Street windows tax, Lufthansa fares.
Links: 2→94, 9→98, 12→95 (pointer pill), 17→99 (pointer pill); back
pills 94→2, 96→12, 97→12, 98→9, 99→17 (95 has none, flows to 96, as
in the source). Slide-6 hand-edit ported (bullets_top 2.21"/4.13" with
dated comment); full-deck geometry diff showed no other hand-edits;
`_test` deck deleted. Backups rolled to `_t-1`/`_t-2`.

Implementation notes:
- `_splice_media.py` SPLICE_MAP entries are now (source deck, display):
  "IC" = In Class, "WS" = In Class with Solutions (also 4:3, same
  +1.667" shift). `_group_pass.py` SPLICED and `_animate.py` skips /
  PLANS renumbered via `_m1_shift_key2`; `_verify_anim.ps1` table
  renumbered (still 65 animated slides — ALL COUNTS MATCH; slideshow
  probe incl. all-new polls + backups PASSED).
- **Hyperlinked text runs render UNDERLINED on this machine regardless
  of u="none"** (verified against a native PowerPoint save — even
  PowerPoint's own no-underline hyperlink run renders underlined).
  Slide-jump affordances therefore use SHAPES: invisible 100%-
  transparent-fill overlay rects on slides 2/9 (over the existing
  gold ▶ lines), `_add_ps_pointer` pills on 12/17, navy back pills
  (`_add_back_pill`) on the backup slides.
- Backup screenshots with baked-in red-circle annotations were cropped
  from 2400px slideshow exports of the WS deck
  (`_source_images/ws66_goodlife_crop.png`, `ws68_fares_*_crop.png`);
  plain images extracted verbatim (`ws63_*`, `ws64_*`, `ws65_*`,
  `ws67_*`).

Flags for Nico (not changed, awaiting word):
- WS slide 26 (heatwaves SOLUTION) also shows a flour-shortage
  clipping + "other examples of right-shifts" bullets that our native
  AC-solution slide (#26) doesn't have (flour appears in Video 3,
  display 82). Add to #26?
- WS backup "National Leaders" carried stale speaker notes (hedgehog
  text) — wrote fresh notes instead; Portland notes ported verbatim.
- WS deck uses the FEMBA TA email; deck keeps TA405.EMBA2@gmail.com.
- Spliced poll slides carry the source's static page-number text
  (e.g. "46" on display 50) — cosmetic, pre-existing behavior.

## 2026-08-20 (round 3) — "Module 1 - Example Candidates.pptx" (14 slides)

Nico asked for a thorough, careful web search for recent (2023–2026)
real-world, MBA-compatible examples for every Module 1 concept, delivered
as a separate review deck. Process: 5 parallel paper-writing-agent
research runs (market definition, S&D shocks, opportunity cost, sunk
costs, marginal analysis + fairness), each verifying facts against
primary/tier-1 sources and flagging confirmed vs. reported figures.
Deliverable: `Module 1 - Example Candidates.pptx` — cover + 11 candidate
slides + 2 "bench" slides (runners-up, one line each). Build script:
`_build_M1_candidates.py` (imports the `_build_Module1.py` helper
layer). Each candidate slide: concept tag, fact bullets, cream
teaching-angle card w/ proposed visual, gold discussion prompt, source
line; full URLs + verification flags in the speaker notes.

Candidates: Tapestry/Capri "accessible luxury" · Kroger/Albertsons
"Costco run" · Netflix+WBD (3 market definitions — updates existing
Netflix slides) · Eggs 2024–26 (BLS-verified) · DRAM/AI memory · AI
talent war ($100M implicit cost) · Return-to-office (AEA-published 72
min/day) · Apple Car + GM Cruise (SEC-confirmed exit math) · Meta
Reality Labs ($80B debate) · United/Delta marginal-flight cuts (2026
fuel shock) · LA-fires 10% rent cap (§396, local). Bench: FTC v. Meta,
Google/AI-chatbot market, FTC v. Amazon, beef, coffee, cocoa, GLP-1,
Berkshire cash, hyperscaler capex, NIL, CA HSR, Ørsted, Sony Concord,
Google demand response, Wendy's, egg rationing/DOJ.

Standing rule respected: press-reported figures are marked "reported"
on-slide; speaker notes name what must be re-verified before a
candidate graduates into the main deck. Awaiting Nico's picks.

**Round 3e (same day):** Deck-wide ≥18pt font pass on the candidates
deck (Nico: "font inside text boxes at least 18pt"): teaching-angle
cards, discussion lines, quote boxes, diagram boxes, timeline labels,
callouts all raised to 18+; chart labels to 16; photo/source captions
stay at caption size. Resolution lines on the Tapestry, Costco, and
Netflix-chart slides promoted to gold takeaway bars (19pt bold navy);
their on-slide source lines moved to speaker notes for space. Bench
split into three slides to fit 18pt. Deck now **21 slides**. Also:
Costco-run title reworded to "Everyday Shopping" (Nico), Kroger setup
wording varied ("Once again, everything would hinge on how you define
the market").

**Round 3d (same day):** Candidates deck now **20 slides**. (a) DRAM
case expanded to 3 slides: setup (datacenter photo, chronology-first
flag on wafer allocation), two-panel native S/D analysis (HBM: D shifts
right against steep supply; consumer DRAM: S shifts left as wafers
reallocate — gold arrow between panels), resolution (H100 + SK Hynix
DDR5 photos; magnitudes marked reported/TrendForce). (b) Tapestry and
Kroger setup slides restructured per Nico: deal first, then "market
definition would turn out to be crucial" with the two sides' market-
extent arguments as sub-points; photos now stacked vertically on the
right, text on the left two-thirds, fonts 24/22. (c) Share-figure
provenance verified from the Clifford Chance briefing (read directly):
58.7% = FTC's expert from largely third-party data; 77%/83% = Capri's/
Tapestry's internal ordinary-course documents produced in the merger
investigation (not leaked); "accessible luxury" was the firms' OWN term
from SEC filings and investor decks until the FTC sued (then
"expressive luxury") — now on the evidence slide + notes. New Commons
photos: web_dram, web_datacenter, web_h100.

**Round 3c (same day):** New standing preference from Nico, added to
`Teaching\CLAUDE.md` ("Case buildup: chronology first, resolution
second"): a two-slide mini-case builds the situation on slide 1 and
ends by flagging the crucial feature ("the definition of the market
would turn out to be crucial") WITHOUT revealing the outcome; slide 2
shows the resolution as the final beat/click. Applied to both the
Tapestry–Capri and Kroger–Albertsons pairs in the candidates deck
(court decisions moved to the second slide of each pair).

**Round 3b (same day):** Nico asked to expand candidates 1–3 + United
into 1–2 slides each with illustrative pictures. Deck now **18 slides**:
Tapestry–Capri (case + price-ladder/evidence slide with the 59/77/83%
share cards and the internal-message quote), Kroger–Albertsons (case +
"Costco run" in/out-of-market diagram with the Nelson quote),
Netflix–WBD (deal-saga timeline + native Nielsen TV-time bar chart vs.
the red SVOD callout), United (case with Kirby-quote callout + native
MB=MC chart with the crossing shifting left). Photos fetched from
Wikimedia Commons via `_fetch_web_images.py` (BUILD INPUT; images in
`_source_images/web_*.jpg`: Coach + Michael Kors stores, Kroger,
Albertsons Dallas, Costco, United 787, WB water tower) — all reviewed
before use; "Photos: Wikimedia Commons" caption lines on-slide. The
remaining 7 candidates + 2 bench slides unchanged.

## 2026-08-20 (round 2) — comic back + 2 more MW applications: 87 slides

Nico approved: (1) title-slide comic reintroduced; (2) NEW #23 "Shifts
of the Demand Curve for AC" (MW #51 solution after the AC poll, native
D→D′ chart); (3) NEW #37–38 copper mini-case (MW #65–66, two-stage
quantity/price figure + native both-shift-right chart, P1 = P0).
Old #23+ shifted +1, old #36+ shifted +3; polls now at 22/25/45. Page
numbers renumbered via `_renumber.py` (descending literal replacement);
`_animate.py` keeps pre-insert PLANS keys shifted by `_m1_shift_key`.
Full pipeline re-run; renders of 1/23/37/38 checked; click structure
re-verified (ALL 65 animated slides match).

## 2026-08-20 — Full rebuild into "Module 1 - Revised.pptx"

**One-line summary.** Built `Module 1 - Revised.pptx` (**84 → 87 slides**,
16:9, new 405 format) from Nico's 53-slide 4:3 In-Class deck plus the
four video decks (25 slides) appended at the end, adopting 5 approved
MW (Melanie Wasserman) items; 3 PollEv slides spliced live; grouped;
fade builds applied; click structure + slideshow probe verified.

### Structure
- Slides 1–58: In-Class part (front matter, models/philosophy, markets
  + S/D mini-cases, opportunity costs, sunk costs, CBA, summary).
- Slides 59–84: Videos 1–4, each with its own deck-format title slide
  (Nico will eventually split them back out into separate video decks).
- Outline slides keep Nico's order: videos listed first on the in-class
  outline (slide 16), in-class first on the video outlines (as in the
  video sources).

### Pipeline (rerunnable, Module 7/2 pattern — 4 steps)
```
python _build_Module1.py           # phase-1: all 81 scripted slides + 3 stubs
python _splice_media.py            # 3 PollEv slides verbatim (w/ notes+tags)
python _group_pass.py              # 7 groups (callouts, table shades)
python _animate.py all apply       # fade builds per per-slide plans
```
Helpers (`_build_template_samples.py`, `_animate.py`, `_group_pass.py`,
`_splice_media.py`, `_handoff_pollbreak.xml`) carried from Module 2.
`_animate.py` got one engine fix: shape text is whitespace-normalized so
`t:`/`pr:` prefixes can span run boundaries. `_group_pass.py` got a
rule-1 height cap (≤2.5") so the Homo-Economicus cream panel isn't
falsely paired with one of its text blocks.

### Decisions locked (2026-08-20, Nico)
- MW imports 1–5 adopted: flip-a-house Solution (new #43), shift-
  combination table (new #84), LA real-estate mini-case (new #34–35),
  Next Steps (new #58), Swiftonomics diamonds refresh (new #23 + #25);
  optional items 6–9 declined.
- Slide 4: "Fall 2025" → "Fall 2026" Achieve site.
- Slide 5: exam periods → [DATE] placeholders.
- Old slide 25's lithium-article note kept for now (new #26).
- Problem-Set pointers generic ("Problem Set 1", no exercise numbers).
- Nico's teaching order confirmed: videos watched FIRST, in-class
  applications stay where they are; MW imports slotted as extra
  applications only.

### PollEv caveat (IMPORTANT, Nico action)
New slide 24's spliced poll still asks "How does the DECLINE in
engagements affect the demand for diamonds?" — with the Swift example
the answer flips (demand shifts RIGHT). Reword the activity in the
PollEverywhere account (URL/embed stays valid); the static screenshot on
the slide will still show the old wording.

### Verification done
- All 84 slides render-checked via COM PNG exports (2 rounds; 6 layout
  fixes applied: s12 overlay, avocado TIFF alpha, s56/s72/s75 overlaps,
  s81 label, s83 P-label separation).
- Deck opens clean in PowerPoint (84 slides).
- Animation click structure verified via COM MainSequence TriggerType:
  ALL 62 animated slides match the plan.
- Full-screen slideshow probe (screenClass PrintWindow captures on
  slides 1, 22, 24, 42, 55, 84): PASS — all 3 live polls render their
  activities in the real slideshow; no "failed to open" banner.

### Content flags for Nico (reported in chat)
1. Title slide: the comic strip from the old title slide was dropped
   (new-format title slides are clean); the UCLA logo likewise.
2. Old #15's stray Lufthansa logo not carried into the new roadmap;
   roadmap wording standardized to the M3/M2 format ("1. Basic
   Principles and Economic Way of Thinking"); video-1 agenda's
   "2. Buyers, Value, and Demand" wording also standardized.
3. New #21 got an action title ("How Can Heatwaves Affect the Demand
   for ACs?") — the source slide had no title, only the question.
4. New #47 title rendered as "Similar Figures for the US, Estimated in
   2022" (source title had a line-break artifact).
5. Video title slides read "Module 1 – Video n" (source said "Week 1").
6. Fruit table (new #38) is now a native table — it reveals as one
   block, not cell-by-cell like the old shape-built version.
7. Next Steps (#58) keeps MW's two pre-class sub-bullets ("Read news
   article…", "Take survey…") — cut if they don't fit Nico's flow.
8. Exercise diagram (new #55) rebuilt natively with clean MB/MC values
   (blue net-benefit / red MC, indifferent at hour 4, STOP at hour 5).

### Suggested additional MW applications (awaiting Nico)
- **AC-heatwave Solution slide** (MW #50–51): after the AC poll (new
  #22) the deck jumps straight to Swiftonomics; MW closes the example
  with "demand shifts right". Could add a native D→D' solution slide.
- **Copper since 1880** (MW #65–66): both curves shift right → quantity
  ×100 at flat price; completes the shift taxonomy next to tea/avocado/
  LA and the #84 table.

### Pending / next steps
- Nico's eyeball pass of the deck + slideshow.
- PollEv rewording (see caveat above).
- Speaker notes: substantive originals preserved verbatim; MW-adopted
  and NEW slides carry drafted 2–4-sentence notes. No teleprompter pass
  requested yet for this deck.
- Not committed to git yet (Nico confirms at session end).

### Gotchas learned this session
- PowerShell COM: `New-Object -ComObject PowerPoint.Application`
  attaches to the RUNNING instance — never call `$pp.Quit()` when the
  user has decks open (it killed his PowerPoint once; only close
  presentations you opened, read-only).
- `_animate.py` joins runs with spaces when collecting shape text —
  without whitespace normalization, `pr:`/`t:` prefixes that cross run
  boundaries never match.
- Converting TIFF→PNG with `.convert('RGB')` flattens alpha to BLACK;
  composite on white first.
