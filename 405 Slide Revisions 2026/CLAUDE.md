# Slide-design preferences (course layer)

## Purpose of this file
This is the third layer of CLAUDE.md instructions for Claude Code,
sitting on top of the universal CLAUDE.md (`d:\Claude Code\CLAUDE.md`)
and the Teaching CLAUDE.md (`d:\Claude Code\Teaching\CLAUDE.md`).
Claude Code loads all three automatically when working in any
sub-folder of this course.

The higher-level files cover identity, audience, drafting workflow,
and general PowerPoint conventions. **This file is narrower**: it
captures my personal taste for how slides should look and feel –
typography, layout, color discipline, formula rendering, picture
treatment, chart styling, and the small visual conventions I've come
to prefer through iteration. The aim is for any new slide or deck I
build in this folder to feel consistent with the ones I've already
shaped, without me having to repeat the same preferences each session.

When two layers conflict, the more specific layer wins (this file
overrides the Teaching layer, which overrides the universal layer).
Where this file is silent, fall back to the higher layers.

## Equations and formulas
- **TeX style, always.** Anywhere a formula has subscripts, superscripts,
  fractions, Greek letters, or anything more than `a = b + c`, render it
  with OMML / Cambria Math – not plain text and not Calibri math.
- **Variables italic; multi-letter acronyms upright.** Single letters
  like Q, K, L, w stay italic by default (don't add `m:rPr`). Acronyms
  like MRPL, MPL, MC, TFC get `m:sty="p"` so they sit upright the way
  economics journals print them.
- **Stacked fractions for ratios** (e.g., bang-for-the-buck rules).
  Inline `/` is acceptable only inside running prose.
- **Hero formulas get their own box.** A navy filled rectangle with a
  small Calibri label on top and the OMML formula filling the rest is
  cleaner than dropping the equation into surrounding text. Use this for
  the headline equation on any "rule" or "concept" slide.
- **Bullets with embedded variables:** rendering full OMML inside a
  bullet is overkill. Unicode subscripts (`pₖ`, `MPₖ`, `MPₗ`) are an
  acceptable middle ground for in-line text. Standalone formulas stay in
  OMML.

## Pictures
- **Drop shadow + rounded corners by default.** Pictures should feel
  lifted off the slide. Soft shadow, modest blur, slight offset – not
  heavy.
- **Flat exceptions** are fine when a picture is functionally a logo
  (e.g., a brand mark) – there, drop the shadow and don't round.
- **Shadow opt-outs (explicit list):**
  - Logos and brand marks (e.g., Zoom, Anthropic) – flat, no shadow.
  - Book covers, posters, screenshots with their own visible border or
    background – no shadow.
  - Source images that already include a built-in shadow or frame –
    no shadow.
  - Set the image-helper's shadow flag to `False` in these cases.
- **Caption ABOVE the picture, small italic-bold navy, centered.**
  Source attribution / licence (when needed) goes below the picture
  in smaller italic gray. The "title-on-top, attribution-on-bottom"
  split mirrors the way figures are labelled in printed reports.
- **Prefer real photographs over logos** when the slide is illustrating
  a real-world example. A photo of the product or the place beats a
  brand mark every time.

## Tables and charts
- **Same shadow treatment as pictures.** Tables and charts also sit on
  a soft drop shadow – use a white backing rectangle with the shadow
  applied to it (OOXML doesn't let you put a shadow directly on a
  graphicFrame).
- **Native charts over static images.** Whenever the data exists in
  code, generate the chart natively. The chart and the table feeding it
  should be driven by the same helper so they can't drift apart.
- **Legend inside the plot area, top-left, with a white fill behind it.**
  Don't park the legend below or beside the chart unless space is
  desperate. White legend fill keeps it readable when it overlaps the
  rising part of a curve.
- **Markers per series, distinct shapes.** Circle / triangle / square /
  diamond – one per series. Color alone is not enough; markers
  reinforce the legend for handout printing.
- **No horizontal gridlines** unless they're pedagogically necessary.
  If you do show them, make them light gray and dashed.
- **Axis labels in bold italic, navy.** Tick labels regular Calibri,
  navy. Title only when the chart needs one beyond the slide title.
- **Round Y-axis maximum** to a clean number so the gridlines fall at
  intuitive intervals (e.g., 0–1000 in steps of 100, not 0–982).

## Color discipline
- **Three colors do the work:** one strong primary, one warm accent,
  one neutral. The primary carries headers / filled boxes / structural
  arrows; the accent is reserved for emphasis, takeaway bars, anchors,
  and "this is the point" callouts; the neutral handles captions and
  secondary text.
- **Gold (or whatever the accent is) loses its power if overused.** If
  more than ~20% of a slide is accent-colored, prune.
- **Backgrounds stay white.** Filled boxes are how I create visual
  weight, not background tints.
- **Pair related visualisations by accent color.** When two charts or
  shapes represent the same underlying concept (e.g., a tangent slope
  drawn on the production-function chart and the value of that slope
  plotted on a second MPL chart), use the SAME single accent color
  across both so the eye links them automatically. Worked example:
  Module 3 slide 15 uses gold tangents on the Q-vs-L chart and a gold
  MPL curve on the right-hand chart — students read "slope here ↔
  value there" without any extra annotation.
- **Reserved pedagogical accents** (off-limits for structural use; for
  concept-introduction slides only):
  - Blue `#0070C0` for the **concept name** being formally introduced
    (e.g., "Marginal Product of Labor" inside an "Important Concept"
    line).
  - Dark yellow `#B8860B` italic for the **emphasised word** in a
    definition (e.g., the word "change" in "change in output due to a
    change in labor input").
  - These two colors are not part of the structural primary / accent /
    neutral palette. Do not use them for chrome, structural boxes, or
    on slides that aren't formally introducing a new concept.

## Visual hierarchy: boxes, arrows, bridges
- **Filled boxes = primary content nodes** (e.g., a key concept, a rule,
  a definition). Filled in the primary color, white text.
- **Outlined boxes = annotations, bridges, "see also" notes.** White
  fill, accent-colored border, primary-color text. Lighter visual
  weight than the filled boxes.
- **Three-level box rhythm on concept slides.** Concept-introduction
  slides naturally settle into a vertical stack of three boxes:
  1. **Hero**: primary-color filled box (white bold) carrying the
     headline definition – "X = Y" or "X = the … from …".
  2. **Elaboration**: cream-fill / primary-color thin-border rounded
     rect holding the formal decomposition, glossary of terms, and
     2–3 bullets that expand the definition.
  3. **Action**: accent-color filled box (primary-color bold) holding
     the actionable rule or one-line takeaway.
  Primary → secondary → action. Slides 12 and 17 in Module 3 are
  worked examples of this pattern.
- **Arrows carry meaning.** Primary-color arrows for structural flow
  (parent → child, step 1 → step 2). Accent-color arrows for
  "this leads to that" cause-and-effect or "remember this here"
  pointers.
- **Bridge boxes between clusters.** A single outlined box that names
  a relationship and has one inflow + one outflow arrow beats a tangle
  of diagonal cross-cluster arrows.

## Recurring concepts get a distinctive shape
- **If prompted, pick one non-rectangular shape per recurring concept.**
  A 12-point star, a parallelogram, a starburst – whatever – and reuse
  it everywhere that concept appears. Repeated rectangles all blur
  together; one consistent oddity becomes a wayfinding cue across the
  deck.
- **Don't fight the shape with text.** If text doesn't fit cleanly
  inside a non-rectangular shape, layer a separate text box on top –
  don't shrink the text and don't deform the shape.

## Top bar and footer chrome
- **Three-level hierarchical section tag** in the top bar:
  `Module · Part · Section`. Title-cased. This is the wayfinding for a
  long deck.
- **Action title** as the slide title – the takeaway, not the topic.
  "Prices Coordinate Strangers Without Central Direction" beats "Supply
  and Demand."
- **Footer is minimal**: page number right-aligned, optional course
  footer text left-aligned, thin rule + accent strip above. No
  "Page X of N" – just the number.

## Layout patterns I reach for
- **Single slide-layout master.** One layout for the whole deck. Strip
  all the python-pptx defaults; one master keeps the deck coherent.
- **Two-column comparisons** for any "X vs. Y" content (short run vs.
  long run, option A vs. option B). Symmetric column widths, header
  cells on top, parallel bullet structure.
- **Three-card row** for "the three cases" content (e.g., falling /
  flat / rising). Equal widths, even spacing, parallel sentence
  structure inside each card.
- **Takeaway bar at the bottom of dense slides.** Accent-color filled,
  primary-color bold italic text, centered. The one-line punchline.
- **Discussion-break badge** for group-discussion cues. A distinctive
  slanted parallelogram shape, accent-colored, bottom-right corner.
- **Convention callout box.** A small cream-fill rounded rect with a
  thin primary-color border, slight rounding (~6%), soft drop shadow,
  holding a bold primary-color "Convention:" prefix and a single line
  of explanatory text or OMML formula (14–15 pt). Use this to record
  a notational or computational convention adopted in the deck – e.g.,
  "Compute ΔQ and ΔL relative to the initial point." Sits to the
  right of a table or below a hero definition.
- **Concept maps as section anchors.** A network-graph-style overview
  slide at the start of each major section, returned to at transitions.

## Text styling
- **Calibri throughout for slide text** (Cambria Math only inside OMML).
- **Short bullets, 5–10 words each.** No full sentences on the slide.
- **Bold for emphasis, italic for variables / captions / "soft" voice.**
  Don't combine bold + italic for body text – reserve it for takeaway
  bars and italicised callouts where you want the "this matters" cue.
- **Captions: small (11–13 pt), italic, gray, centered.** Same
  treatment for every image attribution.
- **Sub-bullets one indent level deep**, smaller font, secondary-color.
  Don't nest three levels deep – split the slide first.

## When manual tweaks beat the build
- I will sometimes hand-edit a slide in PowerPoint (resize a box,
  reposition a label, tighten line spacing, remove a variable). These
  edits reflect real visual preferences – preserve them in the build
  script so the next rebuild doesn't undo them. Treat manual edits as
  signal, not noise.
- **Default: rebuild the canonical deck in place, no verification.**
  The build script is the source of truth and the start-of-day Git
  snapshot is the safety net, so the normal flow is to write straight
  to the canonical filename (e.g., `Module 3_clean.pptx`) and stop.
  Do **not** run a python-pptx readback, footer-page-number check,
  or duplicate-`<a:effectLst>` audit by default — these add latency
  and noise without changing the outcome.
- **Opt-in verification when I report a problem.** Only run the
  readback / audit when I tell you something looks wrong (e.g., "the
  page numbers are off", "PowerPoint refuses to open the file",
  "shape X disappeared"). At that point, use the readback to diagnose,
  fix the script, and rebuild again.
- **Opt-in side-path pattern when I signal hand-edits.** If – and only
  if – I tell you in the prompt that I have made hand-edits in
  PowerPoint to the canonical deck (e.g., "I tweaked slide 12 by
  hand", "incorporate my manual edits"), switch to the safe
  round-trip:
  1. Build to a side path first (`<deck>_test.pptx`), not the
     canonical file.
  2. Diff against the canonical file (shape positions, page numbers,
     key text via a python-pptx readback) to surface the hand-edits
     and port them back into the build script.
  3. Re-run the build to the side path, verify, then `mv -f` the side
     path over the canonical deck.
  Do **not** invoke this flow on your own – wait for my explicit
  signal that hand-edits exist.
- **Exceptions require confirmation.** If a situation seems to call
  for a workflow outside the two opt-in cases above (verification on
  reported problems; side-path on signalled hand-edits), stop and ask
  me in plain text before acting. Do not invent ad-hoc safety
  mechanisms on your own – extra files (`_test`, `_temp`, `_v2`,
  `_new`), forced moves (`mv -f`, `Move-Item -Force`), parallel
  scripts, or hidden readbacks – without my explicit go-ahead.
- **Capture each hand-tweak with a one-line comment.** When porting
  a manual edit into the build script, annotate it with the prior
  value and date – e.g., `tbl_top = Inches(2.45)  # hand-tweaked from
  2.85 on 2026-05-12`. Saves the next reviewer from wondering why a
  value is "odd".

## Faithful-to-source rebuilding
- When rebuilding a slide that already exists in a prior version of
  the deck, **mirror the source's pedagogical structure** – text
  order, accent colors, callout shapes, definition wording – rather
  than reinventing. The source deck reflects classroom-tested choices.
- The default mode is "stay as close to the original as possible."
  Deviate only when I explicitly flag a slide for redesign.
- When the source's illustrative numbers (e.g., Q values, MPL values)
  differ slightly from what the build script computes, prefer the
  script's values for cross-slide consistency, and flag the
  discrepancy in either the speaker notes or the session notes.

## Things to avoid on every slide
- Walls of text. If a bullet runs past two lines, split or trim.
- Decorative imagery that doesn't carry information.
- Chart junk – legends that duplicate labels, three-letter gridlines,
  axis titles that repeat the slide title.
- Stock photos and clip art. Real product / real place / real person,
  or no image at all.
- Emojis anywhere on the slide.
- "Page X of N" footers, watermarks, "Confidential" stamps.
- Multiple slide-layout masters in one deck.

## Iteration norm
- Expect 2–3 rounds of "too cluttered → simplify" on any non-trivial
  diagram slide. The first build is rarely the final one. When in
  doubt, step back and ask what a busy executive sees in the first 30
  seconds – if the answer isn't the takeaway in the slide title, the
  slide isn't ready.
