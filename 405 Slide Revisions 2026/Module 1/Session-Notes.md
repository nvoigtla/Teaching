# Session Notes — Module 1 (combined In-Class + Videos deck)

## PENDING (2026-08-20 ~1:15 PM — in progress, resume here after /compact)

Nico's open requests on **Module 1 - Revised.pptx** (87 slides):
1. **Hand-edit port (diff DONE, port PENDING):** slide 1 = no diff
   (comic already in script). Slide 6 = bullets box moved by hand to
   top 2.21", height 4.13" (was 1.60"/5.35") — port into
   `slide_06_office_hours` in `_build_Module1.py` with a hand-tweak
   comment. Diff tool: `_diff_slides.py <displays>` (canonical vs
   `Module 1 - Revised_test.pptx` side build, already built). Backup
   made: `Module 1 - Revised_backup_2026-08-20c.pptx` (kept: b, c).
2. **Font-upsize rule, slides 1–10 first** (then Nico reviews; only
   after convergence add the rule to Teaching CLAUDE.md — NOT yet):
   whenever bullet font < 28pt and space allows, enlarge up to 28pt
   main / 24pt sub. Proposed fitting rule: for each bullet box try
   (28/24) → (26/24) → (24/22), keep the largest pair whose PIL-
   measured wrapped height (Calibri, with 12/3pt spacing-before)
   fits the box height with ~5% headroom and no line >2 lines;
   sibling boxes on a slide get equal sizes; slide 6 must fit
   Nico's hand-set box (2.21"/4.13").
3. **Slide 8 (Economic Models) full-bleed rebuild:** integrate Nico's
   attached background image (hiker+map+mountains+model-map, the
   text-free variant) as background; overlay EDITABLE text exactly
   like his attached mock: 3 bullet lines top-left (bold navy key
   phrases), navy circle icons per line (head-bulb / quotes / map),
   "THE REAL WORLD" and "THE MODEL (A MAP)" labels right, navy
   takeaway bar bottom-right with compass icon: "Find a model that "
   white + "matches your needs." gold. UNBLOCKED: Nico saved both
   images as `Image_Hiker_Mountain_with_text` and
   `Image_Hiker_Mountain_no_text` (locate them — likely the Module 1
   folder; extension unknown; use the no-text one as background, the
   with-text one as the layout reference).
   Also DONE meanwhile: all 8 session learnings adopted — 6 added to
   Teaching CLAUDE.md (case-resolution gold bar, provenance rule,
   18pt box floor, quote callout, example-candidates workflow,
   Commons photo workflow) and a NEW project-level CLAUDE.md created
   at `405 Slide Revisions 2026\CLAUDE.md` (adoption protocol, file
   naming, pipeline).
4. After porting + font pass: run full pipeline (`_build_Module1.py`
   → `_splice_media.py` → `_group_pass.py` → `_animate.py all apply`),
   verify via `_verify_anim.ps1` (87 slides, counts within), render-
   check slides 1–10, slideshow probe. Delete `_test` deck when done.

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
