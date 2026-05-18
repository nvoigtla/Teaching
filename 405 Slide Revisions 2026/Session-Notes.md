# 405 Slide Revisions 2026 – Session Notes

## 2026-05-17 – MRPL example reframed; bullet helper overhaul; cross-refs hyperlinked

**One-line summary.** Slides 19/20/21 reworked around L = 4,200 (so students
must locate the table interval) and switched to a net-revenue framing
(MR = $30k per car, not P = $80k), yielding MRPL ≈ $840 per worker per week.
Cross-references converted from natural-language hyperlinks to "(link)"
anchor pattern. `_add_hierarchical_bullets` substantially extended.
Global permission-rules file fixed (several rules were silently dead).

### Decisions and edits

1. **Cross-reference hyperlinks → "(link)" anchor pattern.**
   Earlier in the session we used natural-language anchors: e.g.,
   "From the production-function table" with "production-function table"
   itself hyperlinked. User reverted that: now the prose stays clean and
   a parenthesised "(link)" follows, with only the word "link" carrying
   the hyperlink (blue + underlined). Applies on slides 19, 21, 36.
   Helper `_add_slide_link_in_slide` skips already-hyperlinked runs, so
   calling it twice on one slide picks the next un-linked "link" in
   document order.

2. **L = 4,200 anchor on slides 19/20/21.**
   Previous example used L = 4,000 — too easy: students just read the
   table at that exact row. L = 4,200 forces them to first identify
   which production-function-table interval contains 4,200 (answer:
   4,000 → 4,500), per the MPL convention codified on slide 14
   (compute ΔQ/ΔL over a table interval).

3. **Net-revenue framing for MRPL on slide 19.**
   User hand-edited slide 19 to introduce the per-car contribution
   correctly: of the ~$80k price, ~$50k is material cost, so
   (Net) Revenue ≈ $30k per car. Assume approx. constant in quantity.
   This reconciles slides 19 and 21: MRPL = MPL × MR = 0.028 × $30,000
   ≈ $840 per worker per week (slide 21 now matches this — earlier draft
   was still using $80k × 0.028 = $2,240, which I had flagged as
   inconsistent; user accepted the reconciliation).

4. **Slide 21 restructured into four-step solution.**
   Main bullets (▪): "Check which interval contains the current workforce
   (L = 4,200)" → "From the production-function table (link)" →
   "Compute MPL = ΔQ / ΔL" → "MRPL = MPL × MR". Each followed by its
   derivation as sub-bullet(s). Punchline as a final paragraph in
   smaller (20 pt) font with Wingdings  arrow and underlined "$840 ".
   The "MPL convention" hyperlink that earlier pointed to slide 14 was
   removed — slide 21 no longer needs that pointer because the
   computation reads directly off the table.

5. **`_add_hierarchical_bullets` substantially extended.**
   - Bullet tuples now accept 2-tuple `(text, level)` OR 3-tuple
     `(text, level, opts)` with per-paragraph overrides.
   - `text` can be a string OR a list of `(run_text, run_opts)` for
     multi-run paragraphs (e.g., the punchline run with Wingdings
     symbol + underline emphasis).
   - New paragraph-level opts: `bullet_style` ('main' / 'sub' / 'arrow'
     / 'none'), `mar_l`, `indent`, `space_before_pts`, `size`, `color`,
     `bold`, `italic`.
   - New run-level opts: `font_name`, `size`, `color`, `bold`,
     `italic`, `underline`, `wingdings` (emits `<a:sym typeface="Wingdings"/>`
     so private-use-area chars render as Wingdings glyphs).
   - Empty paragraphs (`text=''`) produce visual spacers (no run).
   - Each paragraph with `level > 0` now receives `lvl="N"` attribute
     on its `<a:pPr>` — minimum-viable fix for PowerPoint's
     Tab / Shift-Tab outline navigation. **Open question:** whether
     this alone is enough, or whether `<a:lstStyle>` must also be
     populated with `lvl1pPr`/`lvl2pPr` definitions. **User to test.**

6. **24/22 sub-bullet sizing rule codified in Teaching CLAUDE.md** and
   propagated to all bullet slides from page 22 onward (12 slide
   builders updated). Slide 49 "Dictionary of Costs" deliberately left
   at 15/12 — dense reference card, full sizing pass deferred.

7. **R1T → R1 standardization.** "R1T" referred to the truck only;
   we want the R1 line in general. ~15 occurrences swept; "R1Ts/week"
   → "R1 vehicles / week".

8. **Permission-rules audit — global settings.json had dead rules.**
   Several entries were comma-joined into single strings with
   descriptive suffixes, e.g.
   `"Bash(ls *), Bash(cd *) — auto-approve navigation"` — which Claude
   Code reads as one literal rule (`ls *), Bash(cd *) — …`) that never
   matches. Fixed by splitting into proper individual rules. Added
   chain rules for `cd * && PYTHONIOENCODING=* python -c:*` etc., but
   the cleaner long-term pattern is unchained `python -c "..."` with
   absolute paths inside the snippet (covered by `Bash(python -c:*)`).

### Open items

- **Test Tab / Shift-Tab in PowerPoint** on the new deck. If it still
  doesn't promote/demote, populate `<a:lstStyle>` with `lvl1pPr` and
  `lvl2pPr` entries in `_add_hierarchical_bullets` (one-time helper
  change, touches every bullet slide).
- **Slide 49 "Dictionary of Costs"** still at 15/12 fonts — sizing
  pass deferred.
- **Material-cost number on slide 19 (~$50k).** Web research suggests
  Rivian's improved R1 BOM is in the mid-$50k–low-$60k range, of which
  the battery alone is ~$15k. The $50k material-cost figure on
  slide 19 is defensible but on the low end; ~$55k might be more
  accurate. Not changed pending user decision.

### Commands / workflows worth remembering

- **Side-path build pattern when hand-edits are reported:**
  `python _build_clean_deck.py "Module 3_clean_test.pptx"`
  (the build script takes the output path as `argv[1]`); diff slides
  via zipfile readback of `ppt/slides/slideN.xml`; then
  `mv -f Module 3_clean_test.pptx Module 3_clean.pptx`.
- **Hyperlink verification one-liner:** read each slide's `*.xml.rels`
  and match each `r:id` referenced by an `hlinkClick` to its
  `Target=` to confirm slide-jump destinations.
- **Helper bullet override examples** documented in
  `_add_hierarchical_bullets` docstring.

---

## 2026-05-16 – Slides 16–19 deep polish + discussion-break deck-wide cleanup

**One-line summary.** Polished slides 16–19 in detail (Black Death, Hiring
Decisions Context/Concepts, MRPL example), aligned slide titles with the
original source deck, added the MB=MC star anchor to slide 18, lifted the
discussion-break box font to 28 pt and dropped it cleanly above the gray
rule across all five instances, and codified the "reformatting an existing
deck vs. creating new content" workflow rule in the Teaching CLAUDE.md.

### Decisions and edits

1. **Title-alignment sweep, clean 16–25 against original 19–28.**
   Per a clarifying exchange, we agreed the offset is "−3" but with our
   customizations preserved (fewer dividers, our newer examples, MB=MC
   stars wherever the MR=MC concept appears). Titles updated to match
   the originals where ours diverged unnecessarily:
   - 16: "Black Death and the Return to Labor" → **"Famous Example for
     Diminishing Marginal Returns"**.
   - 17: "Rivian Hiring Scenario" → **"Hiring Decisions in the Short
     Run — Context & Scenario"**.
   - 18: "Hiring in the Short Run: An Important Concept" → **"Hiring
     Decisions in the Short Run — Concepts"**. Also added the MB=MC
     star anchor at the bottom (original 21 explicitly invokes
     "Recall: MB=MC").
   - 24: "Big Employers Bid Their Own Wages Up" → **"The Case of Wage
     Searchers"**.
   - 34: "Applying the Rule – Recipe for Exams" → **"Applying the
     'Bang for the Buck' Rule"**.

2. **Slide 16 (Black Death) full polish.**
   - Top bullets bumped to 26 pt main / 24 pt sub (per CLAUDE.md
     "sub-bullets ≥ 20 pt" convention).
   - Left text box: rectangle → **rounded rectangle** (~8 % corner),
     fonts to 24 / 26 pt, the question line now **navy** with a
     leading "→ " prefix.
   - Two label-arrows on the chart: **Population** label/arrow in
     dark grey, **Return to labor** label/arrow in black (so they
     visually match the curves they point at).
   - Captured manual moves: image at (6.653, 2.918), labels and
     arrows repositioned to land on the curves.

3. **Slide 17 (Hiring Context & Scenario).**
   - Bullets bumped to 26 / 22 pt.
   - Bottom takeaway: replaced rectangle takeaway bar with a
     **rounded narrower (8.5") box** with drop shadow; question
     prefixed with "→ ".

4. **Slide 18 (Concepts) layout flip + font bumps.**
   - Captured manual reflow: hero/decomp boxes tightened in size and
     pulled up; **MB=MC star moved to the right side**, decision-rule
     bar centred/left, arrow points leftward (star → bar).
   - Bumped Decomposition-box fonts: definitions 13 → 14 pt, three
     bullets 15 → 16 pt, "Even at constant price…" sub-bullet 13 → 14.

5. **Slide 19 (MRPL example) re-laid out.**
   - Compact PF table moved to (9.550, 2.013); helper called with
     `with_axes=False` so we manage axis labels directly.
   - **K (robots)** label: 16 pt navy italic, wider span, position
     (9.374, 1.689).
   - **L (workers)** label: 16 pt navy italic, **rotated 270°** so it
     reads bottom-up like a real row-axis label.

6. **Discussion-break box, deck-wide cleanup.**
   - Helper default `top` lifted **6.6 → 6.25** so the box sits
     cleanly above the gold/grey footer rule at y = 7.135.
   - Font bumped **20 → 28 pt** in the overlay text box.
   - All five call sites (slides 19, 44, 58, 72, 73) stripped of
     their per-call `top=…` so they inherit the new default;
     verified by audit that every box now sits at top 6.25, bottom
     6.97 (clear of the 7.135 rule by ~0.165").
   - Geometry unchanged — the helper was already generating the
     custom-geometry slanted parallelogram with rounded corners that
     user had hand-edited on slide 19.

7. **Teaching CLAUDE.md workflow rule codified.** New "Reformatting an
   existing deck vs. creating new content" subsection under "When
   working on PowerPoint slides…":
   - Reformatting = formatting only, NOT rewriting; preserve original
     titles, bullet wording, structural framing, pedagogical examples.
   - Three allowed deviations: refreshed examples (e.g., Tesla →
     Rivian); numerical updates to today's currency / wages; documented
     corrections.
   - Customizations already in the new deck stay (concept map, MB=MC
     anchors, merged hero-concept slides, …).
   - Speaker notes: preserve substantive originals; rewrite only when
     sparse, missing, or contradicted by the slide.
   - Section dividers: prefer our consolidated dividers over the
     original's recurring outline checkpoints.

### Files modified this session

| File | Status | Notes |
|---|---|---|
| [Module 3/_build_clean_deck.py](Module 3/_build_clean_deck.py) | Heavy edit | Slide 16, 17, 18, 19 builders rewritten; discussion-break helper default top lifted + font bumped; 4 call sites cleaned up. |
| [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) | Rebuilt | Audit clean. |
| [../CLAUDE.md](../CLAUDE.md) | Edit | New "Reformatting an existing deck vs. creating new content" subsection. |

### Pending — pick up tomorrow

- Continue slide-by-slide polish from slide 20 onward (per user's
  "we can polish one-by-one" plan).
- Spot-check slide 19's rotated L label in a few PowerPoint versions —
  rotation rendering can shift the text-box centre slightly.

### Carry-forward gotchas (no new ones today)

- PowerPoint file lock on every rebuild — already in the gotchas list.
- 28 pt "Discussion Break" fits comfortably in the 4.8"-wide inscribed
  rectangle of the rounded-parallelogram badge (no clipping).

---

## 2026-05-15 – Front-half rebuild: Announcements + slides 5–15 mirror source deck

**One-line summary.** Long iterative session reworking the front half of
the deck (slides 1 – 15) so it tracks the original deck more faithfully:
restored the Announcements slide (page 3, with placeholder dates),
brought back original titles and box wording on slides 5 / 6 / 10 / 14,
overhauled the slide-14 MPL-calculation table with per-column colors +
floating between-row annotations + a curved arc anchoring it to the
Convention box, and rebuilt slide 15 as a tangent-illustration that
visually links a gold-tangent left chart to a gold MPL right chart.
Two real bug fixes in the helper layer plus a substantial broadening
of the workflow rules in the course-layer `CLAUDE.md`. **Deck stays at
74 slides.**

### Final deck state

- [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) – **74 slides**, audit-clean (0 broken rels, 0 duplicate `<a:effectLst>`, all footer page-numbers match their position).
- New page 3 = **Announcements** (midterm logistics with `{{MIDTERM_WINDOW}}` / `{{TA_WINDOW}}` placeholders so the actual 2026 dates can be filled in later in PowerPoint). All subsequent page-num footers bumped +1.
- Slides 5 – 15 redesigned in sequence; details below.

### Slide-by-slide changes this session

| Slide | What landed |
|---|---|
| 3 (NEW) | `Announcements` — midterm logistics, 3.5-hour at-home window, TA availability window, material covered (Modules 1 + 2 + PS 1+2), review sessions. Placeholders for actual 2026 dates. |
| 4 | Title reverted to original `Recap of Module 2` (descriptive instead of action-titled). |
| 5 | Agenda flowchart: all 4 module boxes now `_add_rounded_filled_box` (rounded + drop shadow). Gold "we are here" up-arrow moved BELOW box 3 with horizontal shift ≈ 0.55". Connector weights bumped 1.5/2.0 → 3.0/3.5 pt. |
| 6 | Title: `Big Picture of Module 3` (was action title). All 4 boxes rounded + shadowed. Box wording restored to original slide 8 (`Production Functions`, `Costs`, `Demand`, `Output Decisions` — no sub-parens). Font 24/26 → 36. Arrow weights 2.0/1.5 → 3.5/3.0 pt. "Tonight: …" legend removed. |
| 7 | Concept-map arrow weights bumped modestly (2.0 → 2.5 navy, 2.5 → 3.0 gold). MB=MC star anchor now carries the deck-wide drop shadow (added inside `_add_anchor_burst`). |
| 9 | Title `The Production Function`. Big equation `Q = f (K, L, etc)`. Third sub-bullet rewritten as `"etc" can be raw materials, energy…`. Bottom navy takeaway bar **replaced** with a Concept-explanation Convention-style callout (cream-fill rounded rect, 20 pt navy). |
| 10 | Title `Short vs. Long Run: A Critical Distinction` (original slide 11). Picture captions moved ABOVE the pictures with original-deck wording (`Capital fixed in short run` / `Labor (ophthalmologists) fixed in short run`). Right-image caption confirmed to match the ophthalmologist photo. Speaker notes rewritten. |
| 11 | Bottom italic-gray caption replaced with Concept-explanation callout (17 pt navy bold, 2 lines). Table + top arrow + axis label shifted up 0.25" to make room. |
| 12 | In-plot legend bumped (font 12 → 13 pt; bbox `w=0.13 → 0.17`, `h=0.18 → 0.24`; `y=0.10 → 0.05`). `plotArea/manualLayout` added (inner 88 % × 82 % of frame) so the plot fills the chart tightly. Bottom banner converted to Convention-style box, tighter padding (`pad_h=0.15, pad_v=0.04`). Drop `;`, capitalise `The` on line 2. |
| 13 | Bullets now lead with `▪` / `–` glyphs. Sub-bullet font bumped 24 → 26 pt. |
| 14 | Title `Marginal Product of Labor (MPL): Calculation`. Caption + sub-caption replaced with a bullet + sub-bullet structure. Per-column data-cell colors: L/Q black, K red, ΔQ/ΔL green (darker `#1B5E20`), MPL accent blue + bold. Δ-columns + MPL **cells blanked**, values float as overlay textboxes at row boundaries (visual representation of the "compute relative to the initial point" convention). MPL floats have cream `MPL_FILL` rounded chrome; ΔL/ΔQ floats are transparent. 4 green down-arrows between Q rows (shifted ~5 mm right of column centre, then 2 mm left; weight 3.0 pt). 4 wavy green connector lines from each arrow to the corresponding ΔQ first digit (sine polyline, 36 segments, amplitude 0.02"). Convention box widened 5.20" → 4.20" → 5.80" through iterations; font 15 → 17 → 19 pt; ΔL / ΔQ rendered in green via OMML; added an `Interpretation:` second paragraph (`Between 0 and 500 workers, MPL is approximately 0.224`). Blue arrow from MPL column bottom to the Note; Note now lives in its own Convention-style cream box (8.20" × 0.75", centred); `declining` switched from red to ACCENT_BLUE. Smooth inverted-U green arc from the first Q-arrow to the Convention box, apex inside the L=0 empty band (`y=3.35`). |
| 15 | Title `Diminishing Marginal Product of Labor`, 4 bullets matching original slide 18 (3rd and 4th are sub-bullets). Both charts now use `smooth=1` at chart AND series level; `plotArea/manualLayout` pinned at `(0.15, 0.04, 0.80, 0.78)`. Right-chart MPL series switched to GOLD so it links visually with the gold tangents on the left chart. Two dashed gold tangents on the left chart with **hand-edited coordinates** ported into the script (steep: `(1.471, 5.666) → (2.536, 4.692)`; flat: `(4.464, 4.141) → (6.103, 3.750)` — after the 0.18" downward chart shift). Captions moved ABOVE the charts (`y=3.25, h=0.30`); chart frames shifted down to `y=3.58`; takeaway bar moved to `y=6.55`. "plot the slope" callout enlarged ~30 % (1.04" × 0.72" box + 0.72" × 0.39" block arrow) with drop shadow, shifted 0.32" (~8 mm) left of the gap midpoint. Bullets moved up to `bullets_top=Inches(1.53)`. Sub-bullet space-before set to 0 pt via new `sub_line_spacing_pts` parameter on `make_content_bulleted`. |

### Helpers added / extended

- `_add_convention_box(...)` — cream-fill rounded-rect callout, navy border, 12 % corner; supports `prefix + body` or `runs=[(text, opts), …]` form; optional `pad_h` / `pad_v` overrides for tighter chrome.
- `_add_rounded_filled_box(...)` — rounded variant of `_add_filled_box` with auto drop shadow (used on slides 5 and 6).
- `_add_wavy_line(...)` — sine-polyline via custGeom (used for the 4 ΔQ connector lines on slide 14).
- `_omml_run` / `_omml_text` now accept an optional `color=` kwarg (used for green ΔL / ΔQ in slide-14 Convention box).
- `_add_arrow` extended with `dash=` (used for the dashed tangent lines on slide 15) and `_add_arrow_shape` extended with `direction="up"` / `"down"` (used for slide-5 "we are here" up-arrow).
- `_add_anchor_burst` now adds a drop shadow to the 12-point star automatically.
- `make_content_bulleted` accepts `bullets_top` and `sub_line_spacing_pts` overrides.

### Bug fixes

1. **`_add_mixed_textbox` was clobbering OMML run colors.** Post-processing loop unconditionally stripped any `<a:solidFill>` from each OMML run's `a:rPr` and replaced it with `default_color`. Now only adds the default fill when the run doesn't already have one — lets `_omml_run('L', color=GREEN_NUM)` actually render in green.
2. **Chart-level `<c:smooth val="0">` overrides series-level `<c:smooth val="1">`.** python-pptx writes both; setting smooth=1 only on `<c:ser>` is silently ignored by PowerPoint. The chart-helper post-processor now updates BOTH the chart-level smooth (inside `<c:lineChart>`) and the per-series smooth. Without this, slide 15's "smooth" curve rendered piecewise-linear and the analytical tangents looked like secants.

### CLAUDE.md changes

- **Universal** ([Claude Code/CLAUDE.md](../../CLAUDE.md)): untouched this session.
- **Teaching** ([Teaching/CLAUDE.md](../CLAUDE.md)):
  - New `Concept-explanation textboxes (preferred format)` section: cream-fill rounded rect / thin navy border / slight rounding / bold-prefix + body. Documented as the canonical pattern for concept explanations and notational conventions across all teaching decks.
  - New `Sub-bullet sizing — err on the side of LARGER` bullet under Slide Design Principles: typical pairs 28 / 24–26 or 24 / 20–22; sub-bullets at 18 pt or below are almost always too small for EMBA viewing.
- **Course** ([Teaching/405 Slide Revisions 2026/CLAUDE.md](CLAUDE.md)):
  - `When manual tweaks beat the build` rewritten with three-mode workflow:
    1. **Default**: rebuild in place, no verification (no `_test.pptx`, no `mv -f`).
    2. **Opt-in side-path**: only when user says they've made hand-edits → build to side path → diff → port to script → `mv -f` over canonical.
    3. **Opt-in verification**: only when user reports a problem → readback / audit to diagnose.
  - Added `Exceptions require confirmation`: ban ad-hoc `_test` / `_temp` / `_v2` / `_new` files, forced moves, parallel scripts, hidden readbacks without explicit go-ahead.
  - Pictures section: caption position **flipped from below to ABOVE** (mirrors print-figure convention). Source attribution stays below in smaller italic gray.
  - Color discipline: new `Pair related visualisations by accent color` bullet, with slide-15 gold tangents ↔ gold MPL curve as the worked example.

### Decisions made this session

1. **Faithful-to-source rebuilding takes priority over action-title style** for chrome / structural slides. Slides 4, 6, 10, 14 reverted from action titles to original descriptive titles. The action-title rule still applies to content slides where the title carries the takeaway.
2. **Between-row floating annotations** are the right visualisation for transition / Δ values. Pattern: blank the table cell, float the value as an overlay textbox at the boundary between the two rows. MPL floats get the cream chrome (since they're the "headline" value); ΔL / ΔQ are transparent.
3. **Single-color anchoring across paired visualisations** — adopted as a canonical pattern (now in CLAUDE.md). Gold tangents on the Q-vs-L chart linked to a gold MPL curve on the right chart visually says "slope here = value there" without any extra annotation.
4. **Hand-edited coordinates are canonical** when the analytical computation doesn't match the rendered slide. Slide 15's tangent positions came from PowerPoint hand-tuning against the smooth-spline-rendered curve (which differs slightly from `Q = 5·√L`). Replaced the analytical `_draw_tangent` helper calls with hardcoded coordinates.
5. **Rebuild-in-place is the default**; the side-path + `mv -f` pattern is now opt-in only on user signal of hand-edits. Saves ~5 prompts per iteration.
6. **Verification (python-pptx readback / XML audit) is opt-in too** — only fires when the user reports a problem. Don't run by default.

### Gotchas (carry forward)

1. **PowerPoint chart smoothing is set at TWO levels.** Setting `<c:smooth val="1"/>` on `<c:ser>` is necessary but not sufficient — the chart-level `<c:smooth>` inside `<c:lineChart>` is honored over the series setting. Update both.
2. **OMML runs need color baked into `<a:rPr>/<a:solidFill>`.** The `_add_mixed_textbox` helper's default-color-application is now color-aware: it only adds the default fill when no fill is present.
3. **`_add_callout_box` returns the shape, but `_add_arrow_shape` had been returning None implicitly.** Both helpers now return the shape so the caller can apply a drop shadow / further styling.
4. **Inner-plot coordinates of a chart are unpredictable without `manualLayout`.** Default PowerPoint chart layout leaves variable margins for axis labels / titles. For ANY overlay (tangent lines, annotation arrows) that needs to land on a specific chart point, pin the inner-plot bounding box via `plotArea/manualLayout` so the overlay math works from known coordinates.
5. **Chart curves rendered with smooth=1 are spline-interpolated**, not exactly the underlying function. A tangent computed from `f'(x)` may visually "cut across" the rendered curve even though the math is correct. Hand-tune endpoints in PowerPoint and port back into the script.

### Pending – pick up next session

Carried over from 2026-05-14 (no progress this session):

- [ ] **A5 – picture captions + attributions** still missing on slides 29, 36, 41, 56, 57, 69. Now caption-on-top per the new course-CLAUDE.md rule; attribution below in italic gray.
- [ ] **B1 – OMML rendering for hero formulas** on slides 17, 18, 20, 22.
- [ ] **B3 – Layout-5 poll redesign** on slides 19, 27, 38, 52, 59, 73. Still source PollEv screenshots.
- [ ] **D1 – verify slide 28 ($8M) and slide 39 (illustrative MP values)** against the classroom material.
- [ ] **D2 – §2.1 stale Tesla content sweep** (slides 53–57).
- [ ] **D3 – animations.** Re-add Appear/Fade on click in PowerPoint after the deck is structurally settled.

### Useful commands

```powershell
# Rebuild the canonical deck in place (default — no verification).
# Close PowerPoint first if the deck is open.
$env:PYTHONIOENCODING = "utf-8"
python "h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3/_build_clean_deck.py"

# Opt-in: side-path build (only when user signals hand-edits).
python "h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3/_build_clean_deck.py" "Module 3_clean_test.pptx"
# … diff, port edits, then:
mv -f "Module 3_clean_test.pptx" "Module 3_clean.pptx"

# Opt-in: verification readback (only when user reports a problem).
python -c "from pptx import Presentation; p=Presentation(r'h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3/Module 3_clean.pptx'); print(len(p.slides))"
```

---

## 2026-05-14 – Full deck audit against course-layer CLAUDE.md + targeted re-implementations

**One-line summary.** Started the session by pulling the new course-layer
`CLAUDE.md` (committed overnight). Ran a thorough audit of all 74 slides
against the new design preferences, presented a categorised punch list
(A: chrome/styling, B: equations/formula rendering, C: layout / re-design,
D: content-verification), and implemented all of A–C except A29. Notable
re-implementations: full Ross Stores image (slide 50), Excel-driven cost
curves (slides 55–56–57), native ChatGPT illustration (slide 51),
Make-vs-Buy table data fidelity (slide 45), four cost-component boxes on
slide 43, shared production-function table on slides 18 & 35, and a
proper U-shaped LR-AC envelope on slide 65 with each SAC touching the LAC
at its minimum. Iterated twice on the Discussion-break badge after the
first integration broke the text-box / parallelogram coupling. Then
codified two universal preferences (bash bare-commands; "Important
Feedback Memories" workflow) in the top-level `CLAUDE.md`. **Deck stays at
73 slides**; previous count of 74 was pre-merge.

### Final deck state

- [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) – **73 slides** (one merge from yesterday; the audit reduced the count by one further redundancy, then a planned slide was reintegrated). Opens cleanly in PowerPoint; round-trip via `python-pptx` succeeds; non-integer-EMU and duplicate-effectLst audits both 0.
- Cost-side now data-driven from `Background Material/Module 3 - Make vs Buy.xlsx`. Single source of truth: `COST_TFC = 800_000`, `COST_VAR_COEF = 200`, `COST_Q_VALS = [10, 20, …, 110]`; helpers `_cost_tc/_cost_tvc/_cost_avc/_cost_atc/_cost_mc(Q)`. Charts on slides 55/56/57 are generated natively from these.
- Slide 18 (MRPL example) and slide 35 (Rivian optimal-plan check) now share `_add_compact_pf_table()` so the production-function values cannot drift between the two slides.
- Slide 45 (Waterworld) rewritten with the original-deck scenarios (June '94 / Sept '94 / Dec '94, decision "Make!" in all three) and exact column ordering.
- Slide 50 (Ross Stores) – previous copy was a partial fragment of the original PowerPoint object; replaced with a faithful reconstruction (row labels + two year columns + two red-overlay annotations + per-row context).
- Slide 51 (ChatGPT subscription tiers) – static `slide51_rId5` thumbnail replaced with a clean `_chatgpt_logo.png` (Wikimedia PD, 3840×2160).
- Slide 65 (LR-AC envelope) – three SAC curves are smooth cubic-Bezier U-shapes; LAC is a 2-segment envelope through each SAC's minimum (touches each U at its bottom, not floating above).
- All five Discussion-break callouts (slides 18, 44, 58, 72, 73) restored to the **rounded parallelogram + overlay-textbox** pattern, ensuring "Discussion Break" text never escapes the slanted edges.

### Decisions made this session

1. **No fundamental redesigns; respect the source.** Per the course-layer CLAUDE.md rule "stay as close to the original as possible", everything in category C was implemented as faithfulness-improvements (e.g., proper Ross Stores layout, correct Waterworld numbers) rather than reinvention. Slides explicitly flagged for redesign were limited to the four chart/table slides where native python-pptx beats the static image.

2. **Chart series colour discipline.** Three-series cost charts (TC components on slide 55; per-unit costs on slide 56) use NAVY + GOLD + warm-red `#C0504D` rather than the deck's structural NAVY + GOLD + neutral gray, because gray-on-white at line weight reads as "absent" in a chart legend. The warm red is reserved for chart-only use to keep it out of the structural palette.

3. **Compact production-function table is a shared helper.** Pulled the slide 18 table into `_add_compact_pf_table()` and reused on slide 35, so any future change to the K/L grid or per-cell formatting hits both slides at once. Captions and surrounding bullets remain slide-specific.

4. **Discussion-break: overlay textbox over the custGeom.** Setting `<a:rect>` on the parallelogram constrains text bounds inside python-pptx, but PowerPoint's runtime renderer still occasionally pushes text outside the slanted edges. Replaced with: render the parallelogram shape with no text, position a separate textbox at the inscribed rectangle (`ins_left = left + height`, `ins_w = width − 2·height`). Decouples the geometry from the text and is robust across PowerPoint versions.

5. **Slide 65 SAC curves: two cubic-Bezier segments per U.** Single-Bezier U-curves came out too flat at the minimum. Two-segment pattern: segment 1 from `(x_left, y_left)` to `(x_min, y_bottom)` with `CP1 = (10% horizontal, 70% vertical)` for steep descent and `CP2 = (30% horizontal, y_bottom)` for horizontal approach at the trough; segment 2 mirrors. Result: visibly U-shaped, with horizontal tangent exactly at the minimum.

6. **LAC envelope touches each SAC at its minimum, not above it.** Earlier draft had LAC as a single line offset above the SACs ("nice geometry, wrong economics"). Final LAC is a 2-segment polyline through the three SAC minima – the textbook envelope picture.

7. **Universal CLAUDE.md gets the bash bare-commands rule.** Mid-session feedback ("stop prepending `cd '<path>' && …`") was first captured to the per-workspace memory system, then promoted to the top-level CLAUDE.md as a new "Bash / shell commands" section because the rule is project-agnostic and worth durable enforcement.

8. **Universal CLAUDE.md gets an "Important Feedback Memories" section.** Adds a visible, git-tracked counterpart to the per-workspace memory store. Rule: I (Claude) must **ask** before adding an entry; format is `*YYYY-MM-DD* — one-line rule`; promotion threshold is "held up across at least two distinct sessions". Seeded with the bash-bare-commands rule.

### Files added / modified this session

| File | Status | Notes |
|---|---|---|
| [Module 3/_build_clean_deck.py](Module 3/_build_clean_deck.py) | Heavy edit | Added `_make_multi_line_chart()`, `_add_compact_pf_table()`, `_cost_tc/tvc/avc/atc/mc()` helpers and the `COST_*` data constants. Reworked `_add_discussion_break()` to overlay-textbox pattern. Rebuilt builders for slides 18 (compact PF table), 35 (compact PF table), 43 (four cost-component boxes), 45 (Waterworld with source data), 48 (Cost dictionary, native 3-card), 50 (Ross Stores full reconstruction), 51 (native ChatGPT illustration), 55 (TC components), 56 (per-unit costs), 65 (U-shaped SAC + envelope LAC). |
| [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) | Rebuilt | 73 slides, opens cleanly. |
| [Module 3/Background Material/Module 3 - Make vs Buy.xlsx](Module 3/Background Material/Module 3 - Make vs Buy.xlsx) | New | Source of truth for cost-curve data (quadratic TC fit). |
| [Module 3/_chatgpt_logo.png](Module 3/_chatgpt_logo.png) | New | Wikimedia PD ChatGPT logo, 3840×2160, used on slide 51. |
| [../../CLAUDE.md](../../CLAUDE.md) | Edit | New "Bash / shell commands" section. New "Important Feedback Memories" section at the end with workflow + first memory seeded. |

### Gotchas (carry forward)

1. **`Inches(<EMU value>)` double-wraps.** Slide 45's first build had `Inches(col_x0 − MARGIN − Inches(0.10))` where the inner expression was already in EMU; the outer `Inches()` multiplied by 914,400 again, producing a position ~220 000 inches off-canvas. PowerPoint refused to open the file with no clear error. Fix: only wrap raw float-inches inputs in `Inches()`; never wrap an expression that already contains an `Inches()` term.

2. **Single-Bezier U-shapes are unreliable.** Cubic Bezier with one segment produces an L or a J more often than a U, depending on where you place the control points. For U-shaped curves, use **two** segments meeting at the minimum with horizontal tangent there (CP at `(x_min ± Δ, y_min)`).

3. **LAC envelope drawn through SAC minima, not offset above.** The visually obvious "offset above" approach is economically wrong; the envelope **touches** each SAC at one point. Use the actual `(x_min, y_min)` triplet of the three SACs as the LAC control points.

4. **`<a:rect>` inside a custGeom shape doesn't always constrain text in PowerPoint.** Even with the inscribed-rectangle text-bounds rect set, some PowerPoint versions render text past the geometry. The safe pattern for non-rectangular shapes with text is **overlay**: draw shape with no text, add separate textbox positioned at the inscribed rectangle.

5. **Layered CLAUDE.md applies bottom-up.** Three layers in this project: universal (`h:/Claude Code/CLAUDE.md`), teaching (`h:/Claude Code/Teaching/CLAUDE.md`), course (`h:/Claude Code/Teaching/405 Slide Revisions 2026/CLAUDE.md`). Course-layer wins on conflicts. Course-layer is where slide-design taste lives; universal is for cross-project workflow rules (bash, drafting, memory).

### Pending – pick up next session

- [ ] **A1 (slide 29) – deferred by user during this session.** No content change required yet; revisit when next reviewing §1.1b.
- [ ] **A5 – picture captions + attributions** still missing on slides 29, 36, 41, 56, 57, 69. Each needs a source line (Wikimedia URL + CC licence, or photographer attribution). Best to handle in a single pass once user provides decisions on which images to keep vs. replace.
- [ ] **B1 – OMML rendering for hero formulas** on slides 17, 18, 20, 22. Currently using the Unicode-subscript middle ground from the course-layer CLAUDE.md (which permits this for inline bullets). If user wants the hero equations themselves promoted to OMML, route through the existing `_add_math_equation()` helper.
- [ ] **B3 – Layout-5 poll redesign** on slides 19, 27, 38, 52, 59, 73. Currently still source PollEv screenshots. Need the A/B/C/D option text from user before rebuilding as native Layout-5 poll cards.
- [ ] **D1 – verify slide 28 ($8M) and slide 39 (illustrative MP values)** against the classroom material. Numbers are plausible but unconfirmed.
- [ ] **D2 – §2.1 stale Tesla content sweep** (slides 53–57). I touched 55/56/57 for charts and verified the body text mentions Rivian, but a final pass against the source deck would close the loop.
- [ ] **D3 – animations.** Re-add Appear/Fade on click in PowerPoint after the deck is structurally settled. Highest-value slides: 13 (MPL "Note" reveal), 17 (Decomposition + Decision-rule reveals), 18 (PF-table staged reveal), 22 (callout).

### Useful commands

```powershell
# Rebuild (close PowerPoint first; absolute paths so cd is unnecessary).
$env:PYTHONIOENCODING = "utf-8"
python "h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3/_build_clean_deck.py"

# Quick deck health check.
python -c "from pptx import Presentation; p = Presentation(r'h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3/Module 3_clean.pptx'); print(len(p.slides))"
```

---

## 2026-05-13 – §1.1 polish: Tesla→Rivian, $80k price, native charts, design refresh

**One-line summary.** Heavy iteration session on §1.1 Short Run (slides 9-22):
swept Tesla → Rivian everywhere; dropped materials cost from the MRPL
framing (option A) to keep `MRPL = MR × MPL` clean for an EMBA audience;
rebuilt slide 12 to mirror the source deck's original slide 16, inserted
a new slide 13 mirroring original slide 17 (MPL-data table + "initial
point" Convention callout); merged old slides 17+18 into a single
"Hiring in the Short Run: An Important Concept" hero slide; replaced
several static images with native python-pptx charts; introduced a
deck-wide drop-shadow refresh for figures, tables and charts; redesigned
the Discussion-break badge as a custom-geometry rounded parallelogram with
45° slanted sides. **Deck is now 74 slides.**

### Final deck state

- [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) – **74 slides** (was 75 yesterday; slides 17 + 18 merged), 10.05 MB, audit clean (0 broken rels, 0 missing parts, 0 non-integer EMU, 0 duplicate effectLst).
- §1.1 Short-Run flow (slides 7-22):
  - 9 – Rivian Normal IL plant photo (CC BY-SA, Wikimedia) replaces the Tesla factory floor; pictures shifted up by 1 cm, rounded corners + drop shadow.
  - 10 – Production-function table at K = {100, 200, 300, 400} × L = {0, 1k, … 10k}; arrows + axis labels + Group annotation preserved; shadow rect behind.
  - 11 – Native 4-series line chart; legend in top-left with white fill + 70% line spacing; takeaway banner BELOW the chart; dashed light-grey gridlines on both axes.
  - 12 – **Rebuilt** to mirror original slide 16: "Short Run: Marginal Product of Labor" with K-bar OMML, L italic, "Important Concept" emphasis, big centred `MPL = ΔQ / ΔL` formula. Accent colours: blue (#0070C0) for concept name, dark-yellow (#B8860B) italic for "change".
  - 13 – **NEW**: `slide_mpl_data()` — "Example for MPL Calculation"; 6-column table (L | K | Q | ΔL | ΔQ | MPL) with K=100 fixed, L grid {0, 500, 1k, 2k, 3k}; cream-fill Convention callout on the right using the exact wording from original slide 17.
  - 14 – Native python-pptx line charts (Q vs L, MPL vs L) for K=100, blue; shadow rect behind each.
  - 15 – Diminishing MPL bullets + Black Death image kept.
  - 16 – Rivian Hiring Scenario: ~$80k R1T, fixed K. Single Rivian R1T photo on the right (rounded + shadow). Materials cost **removed** from title/bullets.
  - 17 – **Merged old 17 + 18** into "Hiring in the Short Run: An Important Concept". Hero NAVY box: "MRPL = Marginal Revenue Product of Labor" / "the extra revenue from one more worker". Cream Decomposition box with header `MRPL = MR × MPL`, MR & MPL one-line definitions, and three bullets (When MPL falls → MRPL falls; Decreasing MPL ⇒ value of marginal hire shrinks; For price takers, MR = P, so MRPL = P × MPL). Gold Decision-rule box at bottom: "If MRPL > w (wage), hire more workers".
  - 18 – Static image replaced with **native compact production-function table** (re-uses slide 10 data, smaller font/columns); bullet about price being approximately constant added; caption "Production-function table (slide 10)" moved below the table; redesigned Discussion-break badge.
  - 19 – Poll slide (PollEv screenshot) unchanged in structure; title "What Is Rivian's MRPL at 6,000 Employees?".
  - 20 – **Rewritten** to use ACTUAL table values per slide 13's "initial-point" convention: Q(6k)=387, Q(7k)=418, MPL = 31/1,000 = 0.031, MRPL = 0.031 × $80k ≈ $2,480 per worker per week.
  - 21, 22 – Existing bullets cleaned of materials-cost references; MRPL = MPL × $80k.
- §1.1b, §1.2, §2.1, §2.2 (slides 23-74): untouched except for the deck-wide page_num shift and the discussion-break redesign on the 5 slides that use it (18, 44, 58, 72, 73).

### Decisions made this session

1. **Materials cost: dropped from MRPL framing (Option A).** Original deck mislabelled the materials-cost term as `MC` ("Marginal Cost"), which is conceptually wrong — MC is the cost of producing one more unit, not the per-unit materials cost. To avoid confusing EMBAs, materials are now ignored on the production side; they'll be reintroduced cleanly on the cost side of the module. All references to $40k materials / $50k net-revenue removed from slides 16, 18, 20, 21, 22. Formula on slide 17 is the textbook `MRPL = MR × MPL ≈ P × MPL` (price-taker assumption).
2. **R1T price normalised to ~$80k** everywhere (was $90k). Reflects realistic R1T average transaction price for 2025-26 mid-trim.
3. **Slide 12 rebuilt** to be much closer to original slide 16 — more pedagogical, less "summary of MPL". Uses original-deck accent colours (blue + dark yellow) for visual fidelity; this is the one place in the deck where we deviate from the strict NAVY/GOLD palette.
4. **One slide added (page 13: MPL data) AND two merged (old 17 + 18 → new 17).** Net: 75 - 1 = 74 slides. Page-num bumps applied via regex (62 +1 first, then 57 −1 after the merge).
5. **Shadow treatment, deck-wide.**  Pictures: `_add_source_image()` adds an outerShdw by default (35 of 36 pictures shadowed; Zoom logo on slide 2 left alone; Karl Marx book on slide 8 set `shadow=False` per request).  Tables (slides 10, 13, 18) and charts (slides 11, 14): a white backing rectangle with an outerShdw effect (not blur-offset grey — that bled through transparent chart backgrounds and made the figures look grey).
6. **Discussion-break badge: custom-geometry rounded parallelogram.**  Top + bottom edges horizontal; left + right edges slanted at 45° in real space (skew = box height in path units); 4 lightly-rounded corners via cubic-Bezier transitions. Gold fill, navy bold text, drop shadow. The custGeom's `<a:rect>` defines the TEXT bounding rectangle as the parallelogram's inscribed rectangle (TL → BR vertices), so PowerPoint cannot render text past the slanted edges. Applied automatically on all 5 slides that use the helper.

### Files added / modified this session

| File | Status | Notes |
|---|---|---|
| [Module 3/_build_clean_deck.py](Module 3/_build_clean_deck.py) | Heavy edit | ~600 net lines added. New helpers: `_omml_acc_overline`, `_add_mixed_textbox`, `_make_simple_line_chart`, `_add_drop_shadow`, `_add_graphicframe_shadow`, `_apply_picture_style` extended use, `_inject_raw_xml`. `_add_discussion_break` rewritten with custGeom. `slide_12()` rewritten. `slide_mpl_data()` added. `slide_16()` rewritten (merged old 16+17). Many bullet edits in slides 13/15/18/20/21/22. |
| [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) | Rebuilt | 74 slides, 10.05 MB. |
| [Module 3/_rivian_plant.jpg](Module 3/_rivian_plant.jpg) | New | 3.87 MB — Rivian Normal IL plant floor (CC BY-SA, Wikimedia). Used on slide 9. |

### Gotchas (carry forward)

1. **`<a:effectLst>` must be UNIQUE inside an `<p:spPr>`.**  python-pptx's `add_shape` leaves an empty `<a:effectLst/>` in spPr by default.  If you append another effectLst (for shadow / blur / soft edge), PowerPoint will refuse to open the file ("file is corrupt, repair?"), but python-pptx is lenient and roundtrips fine — so the audit script can't catch this. ALWAYS `for old in spPr.findall(qn('a:effectLst')): spPr.remove(old)` before adding a new one.  Repro / debug: scan `<p:spPr>...</p:spPr>` blocks per slide and count `<a:effectLst` substrings — should be ≤ 1 per spPr.

2. **graphicFrames (tables, charts) cannot host `<a:effectLst>` directly.**  Workaround: paint a thin white rectangle with `outerShdw` BEHIND the table/chart. The chart's transparent plot area shows the white through; the shadow projects outside the rect. The helper `_add_graphicframe_shadow(slide, left, top, w, h)` does this — call it BEFORE adding the table/chart so z-order is correct.

3. **Chart backgrounds are transparent.**  An earlier attempt used a translucent-grey blur rectangle as the table/chart shadow; for tables (opaque cells) that looked fine, but for charts the grey bled through the transparent plot area and tinted the whole figure grey. Always use a WHITE backing rect with an outerShdw effect — not a blurred-offset grey rect.

4. **Custom-geometry rounded parallelogram.**  The asymmetric Discussion-break badge uses `<a:custGeom>`, not a preset.  The text bounding rect is set via `<a:rect l="{skew}" t="0" r="{100000-skew}" b="100000"/>` — that's the INSCRIBED rectangle, not the full bounding box. Without that explicit rect, PowerPoint happily renders text past the slanted edges of the parallelogram (margins on the text body are insufficient).  Skew is computed dynamically as `100000 × height_emu / width_emu` so the slant angle stays 45° regardless of the badge's actual size.

5. **Page-num renumbering is mechanical but error-prone.**  Today's session inserted a slide AND merged two — required two regex passes (`page_num >= 13` +1, then `>= 19` −1).  If you ever need to do this again, capture the bump as a one-off Python script and verify with `for i, s in enumerate(prs.slides): ... read footer page_num text ...` that index+1 == page_num for every slide.

6. **MRPL convention chosen for the lecture.**  `MRPL = MR × MPL` (cleanly).  For the example computation (slide 20), `MPL` uses the slide-13 "initial-point" convention: MPL "at L = n" is `(Q(n) − Q(n−1)) / (L(n) − L(n−1))`.  So slide 20's MPL at L=6,000 is actually computed across the L = 6,000 → 7,000 step using table values Q(6k)=387, Q(7k)=418, giving MPL = 31/1000 = 0.031 and MRPL ≈ $2,480/worker/week.  The user's verbal hint had a typo ("from 7k to 7k") which we interpreted as the 6k→7k interval.

### Pending — pick up next session

- [ ] **Cost side of Module 3 (§2.1, §2.2):** materials cost was dropped on the production side; introduce it cleanly when we hit cost concepts (slide 43+). The narrative should be: "now we add materials and other variable costs back in — and they affect the optimal scale of production differently from labour."
- [ ] **Slide 20 MPL number:** user's expected Q values were 391 and 423; the table actually produces 387 and 418 (Cobb-Douglas Q = 0.5·√(K·L)). The 4-5 unit discrepancy probably comes from a slightly different multiplier the user has in mind. If they want exact 391/423, change `PF_A` from 0.5 to ≈ 0.505 and rebuild — but that affects slides 10, 11, 13, 14 too.
- [ ] **Animations.**  None of the click-to-reveal animations from the source deck are preserved. After visual review, re-add Appear/Fade-on-click in PowerPoint for: slide 13 "Note: MPL is declining" (reveal after numbers), slide 17 Decomposition box (reveal after the hero box), slide 17 Decision-rule box (reveal last), slide 18 production-function table (reveal in two stages?), slide 22 callout.
- [ ] **Slides 19, 27, 38 (Poll slides):** still use the source PollEv screenshot — consider switching to the cleaner Layout-5 design (A/B/C/D auto-numbered + POLL pill) for visual consistency with the rest of the deck.
- [ ] **Stale Tesla images in §2.1 (slides 53-57):** session-2 notes flagged this; needs a sweep similar to today's §1.1 sweep.

### Useful commands

```powershell
# Rebuild the deck (close PowerPoint first!)
cd "h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3"
$env:PYTHONIOENCODING = "utf-8"
python _build_clean_deck.py                         # writes Module 3_clean.pptx
python _build_clean_deck.py "Module 3_clean_test.pptx"   # write to side-path

# Audit + duplicate-effectLst + non-integer-EMU
python _audit_package.py "Module 3_clean.pptx"
python -c "import zipfile,re; z=zipfile.ZipFile('Module 3_clean.pptx'); dups=0
for n in z.namelist():
    if not(n.startswith('ppt/slides/') and n.endswith('.xml')): continue
    x=z.read(n).decode('utf-8')
    for m in re.finditer(r'<p:spPr>(.*?)</p:spPr>', x, re.DOTALL):
        if m.group(1).count('<a:effectLst')>1: dups+=1
print('Duplicate effectLst:', dups)"

# Round-trip test (python-pptx will load the deck and report any errors)
python -c "from pptx import Presentation; p=Presentation('Module 3_clean.pptx'); print(len(p.slides))"
```

---

## 2026-05-12 (session 2) – Tail import, sweep, slide 10/11 rebuild

**One-line summary.** Finished the back half of Module 3 (12 new slides for §2.2 Long-Run Costs & Economies of Scale), restructured the deck (dropped textual outline at slide 6, promoted concept-map to slide 6, renumbered everything), applied slide-7 lessons across the deck (MB=MC anchor on slide 22, Unicode-subscript cleanups), and rebuilt slides 10 and 11 around a clean Cobb-Douglas production function with strict diminishing MPL **and** MPK. Deck is now **74 slides**. Slide 10 still to refine tomorrow from the office computer.

### Files modified

| File | Status | Notes |
|---|---|---|
| [Module 3/_build_clean_deck.py](Module 3/_build_clean_deck.py) | Heavy edit | Added 12 new slide builders (`slide_63`–`slide_74`), shared `_pf_value()` / `_pf_table()` helpers, new SECTION_TAG_P2_LR constant, slide_concept_map updated (Cost-types OMML header, MR=MC sub-label, downward arrow from Average Costs to scale annotation), slide_22 got MB=MC star anchor, slides 33/+ Unicode-subscript bullets. Old `slide_6` function removed. |
| [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) | Rebuilt | 74 slides. |
| [Module 3/_source_images/](Module 3/_source_images/) | New files | Renamed-for-use images: `slide63_rId1/2.jpg` (cost curves & MC chart), `slide69_rId1.png` + `slide69_rId2.jpg` (Embraer + Boeing 787), `slide71_rId1.jpg` (BA A380+A318). Also `_orig_chart_image18..21.png` extracted for chart-style reference. |
| [../CLAUDE.md](../CLAUDE.md) | Edit | Added "When working on PowerPoint slides..." section: Formulas (OMML, m:sty=p for acronyms), .pptx workflow (no python-pptx round-tripping, integer EMUs, build script is source of truth, single layout), Iteration is the norm, Source-vs-notes conflict policy. |

### Decisions made this session

1. **Slide 6 dropped, concept-map promoted to slide 6.** Textual outline replaced by the visual concept map. All later slide `page_num=` values decremented by 1 to keep footer numbering tight.

2. **§2.2 Long-Run Costs subsection added at the end (slides 63–74).** Mirrors what the original deck had as positions 68–79: More Complex Cost Functions (transition slide); §2.2 sub-divider; Short-Run vs. Long-Run Costs (two-column + OMML TC_SR ≥ TC_LR); LR-AC envelope schematic (three SAC bands with the LAC drawn as a gold lower envelope); Economies of Scale – three patterns; technological reasons for EoS; **Aviation case** with Embraer ERJ-145 vs. Boeing 787-9 (numbers updated to ~$25M / ~$290M list price, ~$28 vs. ~$31 per passenger-hour); Diseconomies of Scale (added Boeing-2024 quality-issues reference); Economies of Scope (Airbus A380+A318 photo); **Amazon case re-styled as text/discussion**, dropped the dated Bezos screenshot; Shark Tank mini-case (Vimeo link preserved, two PollEV questions); Shark Tank solution with the deal-comparison numbers.

3. **MB=MC star anchor extended to slide 22 (Optimal Hiring Rule).** Same 12-point star pattern as the concept map. Star sits at bottom-left with arrow pointing into the navy rule-statement bar – signals to students that MRPL = w is the "labor case" of MB = MC, consistent with the visual convention established on slide 6.

4. **Plain-text subscripts → Unicode subscripts in bulleted slides.** `p_K`, `MP_K`, `MP_L` rendered as `pₖ`, `MPₖ`, `MPₗ` in bullets on slides 33+ (long-run application slides). Cleaner than rebuilding bullets as OMML mixed-runs, which would have required a major bullet-rendering rewrite. Standalone formulas (the headline equations) remain in OMML.

5. **Cost-types header on slide 6 (concept map) re-rendered in OMML.** Six cost acronyms (TFC / TVC / AFC / AVC / ATC / MC) drawn in upright Cambria Math, matching the TeX-style convention.

6. **Slide 10 production-function table fully regenerated.**
   - **Root cause:** the original table values had MPK *constant* across the first three K-steps (then dropping). That violates strict diminishing returns to K – the same bug that earned a "CORRECTION" slide in the source deck (orig slide 13).
   - **Fix:** the table is now generated programmatically from **Q = 0.5 · √(K · L)** (Cobb-Douglas with α = β = 0.5, CRS overall but strictly concave in each input separately). K = {100, 200, 300, 400}, L = {0, 1000, …, 10 000}. Integer-rounded values verified to give strictly diminishing MPL down every column **and** strictly diminishing MPK along every row.
   - **Single source of truth:** lifted the function and grids into module-level `PF_A / PF_ALPHA / PF_BETA / PF_K_VALS / PF_L_VALS` constants and a `_pf_value(K, L)` helper. Slide 10 (table) and slide 11 (chart) both call `_pf_value()` so they can't drift apart.

7. **Slide 11 rebuilt as a native python-pptx chart.** Previously a static source image. Now a real `XL_CHART_TYPE.LINE` chart with four series (K=100 blue circles / K=200 red triangles / K=300 gray squares / K=400 gold diamonds), markers added per series via direct XML, legend pinned **inside** the plot area (top-left, via `c:manualLayout`), Y axis fixed 0–1000, X axis 0…10 000, no gridlines (removed via XML). Visually matches the original deck's chart styling.

### Pending — start here tomorrow

- **Slide 10 not yet finalised.** User flagged the slide as "not quite finished" before wrap-up. Likely candidates for cleanup: table cell padding / column widths, exact axis-label placement, possibly a derived-MPL/MPK callout next to the table, or surfacing the **Q = 0.5 · √(K · L)** formula explicitly somewhere on the slide. Pick this up first thing tomorrow.

### Gotchas (carry forward)

1. **PowerPoint file lock.** Every rebuild requires the .pptx to be closed in PowerPoint, otherwise `python _build_clean_deck.py` exits with `PermissionError: [Errno 13]`. If a rebuild fails for this reason, ask the user (or close it locally) before retrying – do NOT retry in a loop.

2. **Render-and-check pattern.** When changing chart styling or OMML rendering, always do one Python test-build to a side-path (e.g. `Module 3_clean_test.pptx`) before overwriting the canonical file – cheaper than waiting for PowerPoint to release the lock.

3. **Build script is source of truth.** Manual PowerPoint tweaks the user makes (resized boxes, repositioned labels, removed variables like `M` in `Q = f(K, L, M)`) need to be preserved in the build script so the next rebuild doesn't undo them. Note them when you see them.

4. **Renaming source images.** `_add_source_image(slide, N, rid)` resolves files matching `slide{N}_{rid}.*` in `_source_images/`. When importing new images from the original deck, rename them to match the NEW slide position (e.g. `orig_slide68_image55.jpg` → `slide63_rId1.jpg`).

### Useful commands

```powershell
# Rebuild the whole deck (close PowerPoint first!)
cd "d:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3"
python _build_clean_deck.py

# Audit: count slides + dump titles
python -c "import zipfile; from lxml import etree; ..."  # see prior scripts

# Verify strict diminishing returns on the production-function table
python -c "from _build_clean_deck import _pf_table, PF_K_VALS, PF_L_VALS; ..."
```


## 2026-05-12 – Clean deck rebuild on the "405 Slides Layout" (single layout)

**One-line summary.** Started a fresh `Module 3_clean.pptx` built from
scratch on the 6-type template system, with the PowerPoint Layout dropdown
stripped down to a single `405 Slides Layout`. Completed 42 of 78 slides
in three batches: front matter (1-6), §1.1 Short Run (7-22), §1.1b Wage
Searchers + §1.2 Long Run + Part 2 section divider (23-42). Added native
OMML (Cambria Math) equation rendering for `MP_K / p_K = MP_L / w` style
formulas, and a distinct "external-document reference" visual for
Teaching Note callouts (cream + dashed navy border + page icon + gold
`SEE TEACHING NOTE →` label).

### Final deck state

- [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) – **42 slides**, single layout (`405 Slides Layout`), audit clean (0 broken rels, 0 missing parts, 0 non-integer EMUs).
- Front matter (1-6): title cover, Zoom logistics, M2 recap, course-roadmap diagram, "every executive decision is a production-and-cost decision" diagram, hierarchical outline.
- §1.1 Short Run (7-22): section divider, production function with Karl Marx book, short-vs-long-run pair, rebuilt 9×9 production-function table, output curve, MPL with "plot the slope" callout, Black Death with half-page parchment textbox + chart, Tesla hiring setup, MRPL concept (with Teaching Note callout), MRPL detail, MRPL example, **poll slide using the source PollEv screenshot**, solution, MRPL > wage rule, optimum L\* (with "Revenue per car net of material cost" callout).
- §1.1b Wage Searchers (23-29): caution slide, big-employers-bid-wages-up, salary comparison chart, AI-researcher poaching example with Hassabis photo, poll, solution = **$8M**, UC wage search tool box.
- §1.2 Long Run (30-41): section divider (Part 1.2), Rivian Georgia plant context, optimal-input-mix concept with OMML formula, bang-for-the-buck rule (54pt OMML stacked-fraction headline), 5-step recipe with inline OMML inequalities, Rivian Georgia example, optimality check (production function), analysis with OMML inequality callout, poll, solution table, comparative statics (robot tax | union wages), grocery-shopping intuition reinforcer.
- Part 2 section divider (42): full Module 3 agenda with Part 2 now navy, Part 1 faded.

### Files added or modified this session

| File | Status | Notes |
|---|---|---|
| [Module 3/_build_clean_deck.py](Module 3/_build_clean_deck.py) | **New (main script)** | ~2000 lines. Builds the entire `Module 3_clean.pptx` from scratch. Imports primitives from `_build_template_samples.py`. Has builders for slides 1-42, the layout-stripping post-process, OMML helpers, and the new takeaway/Teaching Note/discussion-break/callout helpers. |
| [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) | **New (output)** | 42-slide draft. |
| Module 3/_source_images/ | New | Pictures extracted from `Module 3_NEW_draft.pptx` (slides 8-22 and 23-42), reused for image embedding in the clean deck. ~5 MB. |
| [Module 3/_zoom_logo.png](Module 3/_zoom_logo.png) | New | Extracted from source slide 2; used by `slide_2()`. |
| [Module 3/_batch2_dump.txt](Module 3/_batch2_dump.txt), [Module 3/_batch3_dump.txt](Module 3/_batch3_dump.txt) | New | Scratch reference dumps of source slides 7-22 and 23-42 with all shape positions, used as input for writing the per-slide builders. Kept for repeatability. |

### Decisions made this session

1. **Approach: per-slide builders, NOT wholesale XML surgery.** First considered lifting source slide XML wholesale and stripping/replacing chrome (the `_rebuild_via_zip.py` pattern), but landed on per-slide builders that explicitly add each source element (pictures, callouts, half-page textboxes, Teaching Notes). Reason: more readable, easier to iterate on per-slide visual issues, and we don't need to preserve animation timing XML (user can re-add in PowerPoint).

2. **Single-layout deck.** Stripped 10 of the 11 default python-pptx layouts via `strip_unused_layouts()`. Kept only the Blank layout, renamed to `405 Slides Layout`. The PowerPoint Layout dropdown now shows exactly one entry.

3. **Font sizes bumped for 2-up handout legibility.** Bullet text 32 pt / sub 28 pt where it fits; 24-26 pt in diagram boxes; section tag 16 pt; footer 12 pt. Slide 6 outline (13 lines) had to stay at 24/20 to fit; flagged as a candidate for splitting if you ever need bigger.

4. **Title case for all chrome.** Removed the `.upper()` from the top-bar helper. Section tags now read "Module 3 · Part 1 · Production" not "MODULE 3 · PART 1 · PRODUCTION". Action titles also title-cased.

5. **Lift-and-faithful-preserve over rewrite.** Every source element the user flagged got added back: Karl Marx book (slide 8), "plot the slope" callout (13), half-page parchment textbox (14), both Tesla images (15), Teaching Note + Optimal Hiring major-concept box (16), MRPL formula box (17), Discussion break parallelogram (18), full PollEv screenshot (19, 27, 38), Revenue-per-car callout (22), Teaching Note (32), Recipe-for-Exams Teaching Note (34), comparative-statics two-column (40), grocery pictures (41).

6. **Teaching Note redesigned as external-doc reference.** Previously rendered as a solid navy bar with gold arrow. Now: cream parchment fill, dashed navy border, navy folded-corner page icon on the left, gold `SEE TEACHING NOTE →` label, italic navy title. Visually distinct from in-slide takeaway/major-concept callouts.

7. **OMML formulas for TeX-feel rendering.** Used Office Math Markup Language (the same engine Insert > Equation uses) for `MP_K / p_K = MP_L / w` on slides 32, 33, 34, 37, 39 (inequality variant), and MRPL formula on slide 17. Native stacked fractions, italic Cambria Math variables, proper subscripts. NOT a hand-styled Unicode hack.

8. **Tesla → Rivian image correction on slide 35.** Source slide 35 still showed a Tesla plant despite the 2026 Tesla→Rivian currency revision. Replaced with `_rivian.jpg` (Rivian R1T, CC BY-SA Wikimedia) and corrected the caption.

### Gotchas — READ BEFORE CONTINUING ON HOME COMPUTER

1. **Non-integer EMU values silently break the file in PowerPoint.** Any
   `<a:off>` or `<a:ext>` with a decimal value (e.g. `3855567.5`) causes
   PowerPoint to refuse to open the deck. Cause: `int / 2` returns a
   float; `Inches(0.7) / 2` returns a float. Defense: always wrap shape
   coordinates with `int()` at the helper boundary, or use `//` for
   integer division. All five shape-primitive helpers
   (`_add_filled_box`, `_add_outlined_box`, `_add_arrow_shape`,
   `_add_arrow`, `_add_half_textbox`) now cast inputs defensively.
   **Test:** `python -c "..."` scanning for `(?:cx|cy|x|y)="-?\d+\.\d+"`
   in slide XML – expect 0 hits.

2. **OMML in a textbox needs `<a14:m>` wrapper to render.** Without it
   PowerPoint shows empty boxes instead of the equation. Required
   structure:
   ```xml
   <a:p>
     <a14:m>
       <m:oMathPara><m:oMath>…</m:oMath></m:oMathPara>
     </a14:m>
     <a:endParaRPr/>
   </a:p>
   ```
   `xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"`
   must be declared on (or above) the `<a:p>` element.

3. **For italic math variables, DO NOT emit `<m:rPr><m:sty m:val="p"/></m:rPr>`.**
   `m:sty="p"` forces "plain" (upright) style which silently overrides
   the italic property in `<a:rPr i="1">`. `_omml_run()` (italic var) is
   stripped of m:rPr; only `_omml_text()` (operators, numbers, acronyms)
   includes `m:sty="p"`.

4. **Animations (click-to-reveal) are NOT preserved on rebuild.**
   python-pptx has no clean API for `<p:timing>`. All the "appear on
   click" build animations from the original deck are flat in the clean
   deck. They have to be re-added in PowerPoint (select shape →
   Animations tab → Appear/Fade on click). Specific slides where this
   matters: 13 ("plot the slope" callout), 16 (Teaching Note), 32
   (Teaching Note), and any other animated callouts.

5. **The Teaching Note style is reserved for external-document references
   only.** Don't mix it with in-slide major-concept boxes. Teaching Note:
   cream + dashed navy border + page icon + gold `SEE TEACHING NOTE →`
   label. In-slide takeaway/major-concept: solid navy or gold fill,
   bold text. The visual separation is intentional.

6. **Source deck has stale Tesla images on Rivian-themed slides.** The
   2026 currency-revision pass swapped Tesla text → Rivian text but
   missed some images. Slide 35 was the one found and fixed; **watch
   for similar issues on slides 53-57** (cost-function chain that was
   also Tesla→Rivian-renamed) when working through batch 4 / §2.1.

7. **Numbers on slide 28 ($8M) and slide 39 (illustrative MP values) need
   user verification.** Slide 28: speaker notes say $8M ($5M + 2×$1.5M);
   stale source body had $3M – I went with $8M. Slide 39: MP_K ≈ 4 cars,
   MP_L ≈ 0.1 cars are illustrative numbers that produce the right
   directional conclusion – swap in actual classroom values when ready.

8. **Slide 32 and 34 content was built from speaker notes.** Source had
   only a title placeholder + Teaching Note callout on slide 32, and
   only a title on slide 34. I generated the major-concept content and
   the 5-step recipe from the bang-for-the-buck logic. Sanity check
   against your teaching notes.

### Pending – next session(s)

- [ ] **Batch 4: §2.1 Cost Concepts (slides 43-62, 20 slides).** Sunk costs, Waterworld decision tree, Apple Car opportunity-cost example, cost dictionary, Ross Stores annual report, ChatGPT subscription tier MC ≠ AC, Burn60-equivalent calculations, MC in finance, Rivian Georgia weekly cost function, iPhone teardown poll, naïve vs. complex cost functions. Several slides have **tables and complex group shapes** (Waterworld decision tree especially) that will need bespoke per-slide builders.
- [ ] **Batch 5: §2.2 Scale & Scope (slides 63-74, 12 slides).** Long-run AC envelope, economies of scale technological reasons, Embraer ERJ-145 vs Boeing 787 (also AI-training scale slides per session 2), diseconomies, scope (Airbus A380 / A318), Amazon scale-or-scope, Shark Tank case + two PollEv slides.
- [ ] **Batch 6: Closing synthesis (slide 75, 1 slide).** Use Layout 6 (the closing synthesis template) – top half Module 3 recap, bottom half Module 4 preview.
- [ ] **Animation pass** in PowerPoint after the deck is structurally final – add Appear/Fade on click for "plot the slope", Teaching Notes, takeaway bars where appropriate.
- [ ] **Verify slide 28 ($8M)**, **slide 39 numbers**, **slide 34 recipe**, **slide 32 content** with classroom material.
- [ ] **Slide 6 outline** (13 lines @ 24/20 with tight spacing) – decide if it should be split into two slides to allow bigger type for handouts.
- [ ] **Re-add manual styling** to slide 19 / 27 / 38 poll slides if you want the cleaner Layout 5 design (A/B/C/D auto-numbered options + POLL pill) instead of the source PollEv screenshot picture.
- [ ] **Stale Tesla images** – sweep slides 53-57 for the same Tesla→Rivian issue I caught on slide 35.

### Commands worth remembering

- **Rebuild the clean deck from scratch:**
  ```powershell
  cd "h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3"
  Remove-Item "Module 3_clean.pptx" -ErrorAction SilentlyContinue
  $env:PYTHONIOENCODING = "utf-8"
  python _build_clean_deck.py
  ```
- **Audit the deck (broken rels, missing parts, declared-but-missing):**
  ```powershell
  python _audit_package.py "Module 3_clean.pptx"
  ```
- **Scan for non-integer EMU values (should always be 0):**
  ```powershell
  python -c "import zipfile, re; z=zipfile.ZipFile('Module 3_clean.pptx'); bad=0
  for n in z.namelist():
      if n.startswith('ppt/slides/') and n.endswith('.xml'):
          x = z.read(n).decode('utf-8')
          for m in re.finditer(r'(?:cx|cy|x|y)=\"(\-?\d+\.\d+)\"', x):
              print(f'{n}: bad {m.group(0)}'); bad += 1
  print(f'Non-integer EMU values: {bad}')"
  ```
- **Extract source images for the next batch:**
  ```powershell
  # Edit slide range at top of the extraction script, then:
  python -c "import zipfile, re; from pathlib import Path
  OUT = Path('_source_images'); OUT.mkdir(exist_ok=True)
  src = zipfile.ZipFile('Module 3_NEW_draft.pptx')
  for sno in range(43, 63):  # next batch range
      try: rels = src.read(f'ppt/slides/_rels/slide{sno}.xml.rels').decode('utf-8')
      except KeyError: continue
      for m in re.finditer(r'<Relationship Id=\"(rId\d+)\" Type=\"[^\"]*image\" Target=\"([^\"]+)\"', rels):
          rid, target = m.group(1), m.group(2)
          media = 'ppt/' + target[3:] if target.startswith('../') else target
          try: data = src.read(media)
          except KeyError: continue
          (OUT / f'slide{sno}_{rid}.{media.rsplit(chr(46),1)[-1]}').write_bytes(data)"
  ```
- **Dump source slide content for a batch (input for writing builders):**
  ```powershell
  # Inside python:
  from pptx import Presentation
  from pptx.enum.shapes import MSO_SHAPE_TYPE
  # ...iterate slides, dump shapes with positions, see _batch2_dump.txt / _batch3_dump.txt for format
  ```

### Useful context for resuming

- Working directory is the network share `H:\Claude Code\Teaching\` on this machine. On the home computer it will be a different path – grep for hard-coded paths if any builds fail (there shouldn't be – the script uses `Path(__file__).parent`).
- `python-pptx 1.0.2` on Python 3.14 confirmed working. Use `lxml` for namespace-aware XML manipulation.
- The git remote is `https://github.com/nvoigtla/Teaching.git`, branch `main`.
- The 405 Slides Template definition is in [Module 3/_build_template_samples.py](Module 3/_build_template_samples.py) (single source of truth for visual constants `NAVY`, `GOLD`, `FADED`, `RULE`, `GRAY`, `WHITE`, `MARGIN`, `RULE_W`, `GOLD_W`, `SLIDE_W`, `SLIDE_H` and the chrome helpers `_add_text`, `_add_rect`, `_draw_action_title`, `_add_bulleted_list`, `_set_bullet_char`, `_blank_slide`, `MODULE_AGENDA`, `FOOTER_TEXT`).
- The clean-deck script `_build_clean_deck.py` adds: `_draw_top_bar_tc` (title-cased top bar, replacing the all-caps default), `_draw_footer` (with bigger 12pt type), `_add_filled_box` / `_add_outlined_box` / `_add_arrow_shape` / `_add_arrow` (diagram primitives), `_add_takeaway_bar` / `_add_teaching_note` / `_add_discussion_break` / `_add_callout_box` / `_add_half_textbox` (callout helpers), `_omml_run` / `_omml_text` / `_omml_sub` / `_omml_frac` / `_add_math_equation` / `_formula_bang_for_buck` / `_formula_mp_ratio` (OMML formula helpers), `strip_unused_layouts` (post-process layout pruning).
- **Critical visual constants** in `_build_template_samples.py`: NAVY=`#0B2B4E`, GOLD=`#E09F3E`, RULE=`#C8CDD3` (light grey horizontal rule), GRAY=`#555B66` (body text grey), FADED=`#B0B5BC` (inactive agenda items), Calibri throughout, 0.7 cm side margins (MARGIN ≈ 0.276" ≈ 251999 EMU).

---

## 2026-05-11 (session 2) – Module 3 rebuild + 2026 currency revision

**One-line summary.** Locked the open outline decisions, cycled the
template through three rounds of visual feedback, rebuilt
`Module 3.pptx` end-to-end into a 78-slide `Module 3_NEW_draft.pptx`
via direct zip+lxml surgery (python-pptx round-trip corrupts this
deck — see gotcha below), then ran a 2026-currency revision pass
(Tesla → Rivian, three new AI examples, Meta Reality Labs + AI training
scale inserts) and replaced the three placeholder logos with real
photographs.

### Final deck state

- [Module 3/Module 3_NEW_draft.pptx](Module 3/Module 3_NEW_draft.pptx) – **78 slides, 24.5 MB**, package audit clean.
- Title slide (#1) – cover-style white background, "Production and Costs / Module 3 / Management 405 EMBA / Prof. Nico Voigtlaender · UCLA Anderson".
- All 76 original notesSlides + 2 new ones populated with MBA-friendly speaker notes (3-6 sentences each, ~14k chars total).
- All character bullets normalized to navy squares (▪) with level-based sizing (lvl 0: 150%, lvl 1: 110%, lvl 2: 85%, lvl 3+: 70%).
- Chrome page numbers updated to match new positions.

### Files worked on (this session)

| File | Status | Notes |
|---|---|---|
| [Module 3/Module 3.pptx](Module 3/Module 3.pptx) | **Untouched** | Original 81-slide deck, never modified. |
| [Module 3/Module 3_NEW_draft.pptx](Module 3/Module 3_NEW_draft.pptx) | New (final output) | 78-slide rebuild with 2026 examples + real photos. |
| [Module 3/Module 3 - outline.md](Module 3/Module 3 - outline.md) | Modified | "Decisions locked" block added; all 5 open questions answered (section dividers in 405 template style; Tesla→Apple Car for opportunity cost; slide #5 kept; closing synthesis added; speaker-notes plan confirmed). |
| [Module 3/405 Slides Template.pptx](Module 3/405 Slides Template.pptx) | Modified | Now a 6-layout reference deck (title cover, section header / Agenda, content bulleted, two-column, poll, closing synthesis). Three rounds of feedback applied. |
| [Module 3/_build_template_samples.py](Module 3/_build_template_samples.py) | Modified | All 6 template layouts plus native PPT bullet helpers (`_set_bullet_char`, `_set_bullet_autonum`, `_add_bulleted_list`). |
| [Module 3/_rebuild_via_zip.py](Module 3/_rebuild_via_zip.py) | New | Main rebuild: 10 cuts, 5 NEW slides via sidecar, reorder, REVISE titles, Apple Car content swap, chrome overlay. **Direct zip+lxml surgery, never round-trips through python-pptx.** |
| [Module 3/_normalize_bullets.py](Module 3/_normalize_bullets.py) | New | Replaces every `<a:buChar>` with ▪ in Calibri (slides + masters). |
| [Module 3/_apply_bullet_hierarchy.py](Module 3/_apply_bullet_hierarchy.py) | New | Sets navy color + level-based `buSzPct` on every bullet pPr. |
| [Module 3/_polish_v2.py](Module 3/_polish_v2.py) | New | Forces white background on slide 1; first Apple Car (Apple Park image) build; adds/replaces notes on all 76 slides. |
| [Module 3/_replace_title_slide.py](Module 3/_replace_title_slide.py) | New | Clean title-slide swap (delete-and-reinject via sidecar) – the working fix after the background hack didn't take. |
| [Module 3/_revise_v1.py](Module 3/_revise_v1.py) | New | 2026 currency revision: Tesla→Rivian text swap, designer→AI researcher, Burn60→ChatGPT, iPhone 11→17, Airbus→Alphabet, + 2 NEW slides (Meta Reality Labs, AI training scale). Updates page numbers and all notes after structural changes. |
| [Module 3/_add_images.py](Module 3/_add_images.py) | New | Rebuilds slides 47 + 69 with bullets-left/picture-right layout; overlays images on 26, 31, 51, 72. First image pass (used logos for 3 slides). |
| [Module 3/_replace_logos_with_photos.py](Module 3/_replace_logos_with_photos.py) | New | Second image pass: removed the Anthropic / ChatGPT / Alphabet logos and replaced with real photographs (Demis Hassabis, smartphone running ChatGPT, Googleplex). |
| [Module 3/_speaker_notes.py](Module 3/_speaker_notes.py) | New | 78 entries, the single source of truth for speaker notes – edit here, re-run `_polish_v2.py`'s `add_notes_to_slide` to re-inject. |
| [Module 3/_audit_package.py](Module 3/_audit_package.py) | New | Sanity-checks declared parts vs zip files, plus broken `.rels` targets. Useful after every edit. |
| [Module 3/_extract_outline.py](Module 3/_extract_outline.py) | Carry-over | Dumps slide titles/body/notes from any .pptx; produced `_outline_dump.txt`. |
| Module 3/_*.jpg, _*.png (7 files) | New | Image assets: `_apple_car_concept.jpg`, `_chatgpt_phone.jpg`, `_googleplex.jpg`, `_hassabis.jpg`, `_meta_quest.jpg`, `_nvidia_h100.png`, `_rivian.jpg`. All CC BY-SA / PD from Wikimedia Commons. |

### Decisions made

1. **All 5 open questions resolved** (see top of `Module 3 - outline.md`): 405 template style for dividers; Tesla → Apple Car on opportunity-cost slide; old slide #5 kept (you'll draft agenda copy yourself); closing synthesis added; speaker-notes plan confirmed (new/merged slides get fresh notes; revised-title slides keep existing notes unless retitle changes framing).

2. **Implementation strategy: Hybrid** – restructure (cuts/inserts/reorders) + restyle (chrome overlay) on the original deck. Three options were offered (skeleton-only, restructure-only, hybrid); you picked hybrid.

3. **Template visual system** – navy `#0B2B4E`, gold `#E09F3E`, light-gray rule `#C8CDD3`, body gray `#555B66`, faded grey `#B0B5BC`; Calibri throughout; 0.7 cm side margins. Native PPT bullets ▪ in navy with level-based sizing.

4. **2026 currency revision scope (you chose "aggressive")** – Tesla → Rivian on long-run + cost-function half (9 slides); 3 AI examples added (AI researcher poaching, ChatGPT subscription tiers, AI training costs); Waterworld kept alongside Meta Reality Labs parallel; both Embraer/Boeing and AI scale slides kept (didn't replace one with the other).

5. **Slide-1 title fix** – first tried adding `<p:bg>` white override (didn't render in PowerPoint); the working fix was to delete slide 1 entirely and inject a fresh template-built one via sidecar.

6. **Image policy** – swapped the three placeholder logos (Anthropic, ChatGPT, Alphabet) for real photos after you flagged that they didn't add visual content. Final image set: Rivian R1T, Meta Quest 3, NVIDIA H100, Demis Hassabis (Nobel 2024), smartphone with ChatGPT, Googleplex Mountain View, plus the existing Vanarama Apple Car render.

### The big gotcha: python-pptx round-trip corrupts this deck

The first rebuild via `_rebuild_module_3.py` (now deleted) produced a file PowerPoint repaired to empty. Bisection showed even a no-op round-trip (open, save) of `Module 3.pptx` via python-pptx breaks the deck — diffing the rels files revealed python-pptx silently **strips `Target="NULL"` image rels** from slides 10, 11, 16, 17, 25 (and a couple others). The slide XML keeps the `r:id` references, so PowerPoint sees dangling pointers and bails.

**Workaround used everywhere now:** all edits to `Module 3_NEW_draft.pptx` are done via direct zip + lxml surgery (`_rebuild_via_zip.py`, `_polish_v2.py`, `_revise_v1.py`, `_add_images.py`, `_replace_logos_with_photos.py`). python-pptx is used only on freshly-built clean sidecar decks (no NULL rels to strip), whose slide XML is then extracted and injected.

### Open / pending for next session

- [ ] **Review the 78-slide deck visually** in PowerPoint and flag any layout issues. Specific things to scrutinize:
  - Slides 26, 31, 51, 72 – picture overlays on existing-content slides; might overlap with original elements at certain zoom levels.
  - Slide 4 ("Agenda for the class") – content is still a stub; you said you'd draft this yourself.
  - The original page-number placeholders from the master may show TWO numbers on some slides (the master's plus the chrome overlay's). My chrome page number is the correct one; the master's may be stale.

- [ ] **Speaker notes for affected slides** – they're plausible defaults, but a few use placeholder numbers (e.g., slide 60 iPhone 17 AVC ≈ $580 is my estimate, not a real teardown). Update with actual figures if you have them.

- [ ] **Rebuild Module 4** – when ready, the same pipeline applies. The scripts here are reusable (just change the cuts/inserts/replacements).

### Commands / workflows worth remembering

- **Rebuild from scratch** (re-runs entire pipeline on `Module 3.pptx`):
  ```powershell
  cd "d:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3"
  python _rebuild_via_zip.py
  python _normalize_bullets.py
  python _apply_bullet_hierarchy.py
  python _polish_v2.py
  python _replace_title_slide.py
  python _revise_v1.py
  python _add_images.py
  python _replace_logos_with_photos.py
  ```
- **Set UTF-8 stdout** when scripts print special characters (▪, →, ·):
  ```powershell
  $env:PYTHONIOENCODING = "utf-8"
  ```
- **Audit deck consistency** after any edit:
  ```powershell
  python _audit_package.py "Module 3_NEW_draft.pptx"
  ```
- **Dump deck content** for review:
  ```powershell
  python _extract_outline.py     # writes _outline_dump.txt from Module 3.pptx
  ```
- **Edit speaker notes** – modify `_speaker_notes.py` then re-run the notes step inside `_polish_v2.py` (or call its `add_notes_to_slide` directly).
- **If PowerPoint locks the file during `.tmp` rename**: the temp file is written successfully; just close PowerPoint and `mv "*.pptx.tmp" "*.pptx"`.
- **When a deck refuses to open after a python-pptx save**: don't try to repair via python-pptx; use direct zip+lxml and preserve the original rels (especially NULL-target image rels).

### Useful context for resuming

- Working directory is now `D:\Claude Code\Teaching\` (different machine from the May 4 session which was on the `H:\` network share). The OneDrive corruption issue from session 1 is gone here.
- `python-pptx 1.0.2` is installed on this machine (Python 3.14).
- The git remote is `https://github.com/nvoigtla/Teaching.git`, branch `main`.
- Folder is now clean: 4 deliverable files, 7 image assets, 12 reusable scripts, 1 outline doc, 1 reference dump. ~50 MB total.

---

## 2026-05-11 (session 1) – Module 3 structural review and template design

**One-line summary.** Reviewed the existing 81-slide Module 3 deck, drafted a revised 73-slide outline, designed a custom MBA-style PowerPoint template ("405 Slides Template"), and committed everything to GitHub.

### Files worked on

| File | Status | Notes |
|---|---|---|
| [Module 3/Module 3.pptx](Module 3/Module 3.pptx) | Untouched (restored from backup) | Original May 4 version, 23 MB, 81 slides. Was corrupted twice during the session by OneDrive / network-share sync (file size jumped to 42 MB and zip central directory broke). Restored from backup both times. |
| [Module 3/Module 3_backup_2026-05-11.pptx](Module 3/Module 3_backup_2026-05-11.pptx) | New | Timestamped backup of the May 4 deck, made before any planned edits. Keep until the revised deck is final. |
| [Module 3/Module 3 - outline.md](Module 3/Module 3 - outline.md) | New | Revised 73-slide outline with explicit Part 1/Part 2 structure, section-header dividers replacing the seven redundant "OUTLINE OF Module 3" slides, takeaway-style titles for ~20 slides, and a closing synthesis slide. |
| [Module 3/405 Slides Template.pptx](Module 3/405 Slides Template.pptx) | New | Custom MBA-style template – navy top bar with white section tag, action-title format, two gold left-edge segments (under title, above footer) creating a parallel rhythm, all main elements within 0.7 cm side margins. Built from scratch with python-pptx, no third-party assets, no attribution required. |
| [Module 3/_build_template_samples.py](Module 3/_build_template_samples.py) | New | Python-pptx script that regenerates the 405 Slides Template. Useful for future tweaks (colors, spacing, accent length). |
| [Module 3/_extract_outline.py](Module 3/_extract_outline.py) | New | Helper script that dumps the slide titles, bodies, and speaker notes from any .pptx to a UTF-8 text file. |
| [Module 3/_outline_dump.txt](Module 3/_outline_dump.txt) | New | Full text dump of the original 81-slide deck. Reference for the rebuild. |

### Decisions made

1. **Top-level structure for Module 3** – two-part, four-section: Part 1 Production (§1.1 Short Run, §1.2 Long Run) and Part 2 Costs (§2.1 Cost concepts, §2.2 Scale & scope). Cuts the seven redundant outline slides; replaces each with a true section-header divider.
2. **Front matter** – condensed from 9 slides to 5 (consolidated logistics+announcements, dropped the orphan "agenda for the class" stub and the duplicate outline, beefed up the "Big Picture" slide).
3. **Slide-title convention** – takeaway-first (per Teaching CLAUDE.md), e.g., "Cost types" → "Sunk costs should never drive decisions"; "Wage searchers" → "Big employers bid their own wages up".
4. **Template direction** – consulting/boardroom style as the base (closest to McKinsey/BCG aesthetic that EMBAs already read at work), with gold accents from an academic-modern style for visual rhythm.
5. **Template specifics** – navy `#0B2B4E`, gold `#E09F3E`, light-gray rule `#C8CDD3`, body gray `#555B66`. 0.7 cm side margins for all main horizontal elements. Two gold left-segments of 2.2" each, sitting on top of the title divider and the footer rule. Title sits close to the top bar (y=0.55") with the divider at y=1.25".
6. **Truly-free template sources** – verified none of the template-design work uses anything with hidden attribution or watermarks. Microsoft Create, SlidesCarnival, and certain GitHub repos identified as the only no-strings sources.

### Open / pending for next session

- [ ] **Build the full slide-master family** in the 405 Slides Template style: title slide, section-header layout, content layout (this), two-column variant, poll-question layout, closing synthesis layout.
- [ ] **Rebuild Module 3.pptx** from the 73-slide outline in `Module 3 - outline.md` using these layouts. Wait for explicit approval on the outline before doing this.
- [ ] **Answer the 5 open questions** at the bottom of `Module 3 - outline.md`:
  1. Section-divider visual style preference.
  2. Whether to swap a Tesla example (currently old slide #53 / new slide #46) for a non-Tesla company for variety.
  3. Confirm that the orphan slide #5 ("agenda for the class") is a leftover and can be cut.
  4. Whether to add a closing synthesis slide (currently in the plan as new slide #75).
  5. Speaker-notes treatment for retitled vs. merged vs. new slides.
- [ ] **Investigate the OneDrive / network-share corruption issue.** `Module 3.pptx` was corrupted twice today (file size jumped from 23 MB to 42 MB, ZIP end-of-central-directory bytes became UTF-8 replacement characters). Suspect OneDrive Files-on-Demand or network-share sync on `\\FSC1\Faculty_Personal\`. Workarounds used: restore from version history, work from a local copy, keep the timestamped backup.

### Commands / workflows worth remembering

- **Extract slide content for review** without opening PowerPoint:
  ```powershell
  cd "h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3"
  python _extract_outline.py
  # writes _outline_dump.txt (UTF-8 safe)
  ```
- **Rebuild the template after tweaks**:
  ```powershell
  cd "h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3"
  python _build_template_samples.py
  # writes "405 Slides Template.pptx"
  ```
- **If a .pptx becomes corrupted again** – check zip header/footer:
  ```bash
  head -c 8 "Module 3.pptx" | xxd   # should start with "PK\x03\x04"
  tail -c 32 "Module 3.pptx" | xxd  # should end cleanly with "PK\x05\x06" then small fields
  ```
  If the footer is followed by `EF BF BD` (UTF-8 replacement bytes), restore from backup or OneDrive version history.

### Useful context for resuming

- The Teaching folder lives on a network share (`\\FSC1\Faculty_Personal\nvoigtla\Claude Code\Teaching`), mapped to `H:\`. Git operations work but binary file sync is unreliable for large .pptx files. Consider working from a local copy and pushing to git rather than relying on the share to be the source of truth.
- python-pptx 1.0.2 is installed and works fine for reading and building decks.
- The git remote is `https://github.com/nvoigtla/Teaching.git` (branch `main`).
