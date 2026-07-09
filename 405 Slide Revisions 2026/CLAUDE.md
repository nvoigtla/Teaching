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
to prefer through iteration. It now also folds in the conventions
developed while harmonizing the *Demand Analysis* decks (Law of
Demand, Elasticities, Demand & Revenue, Demand Estimation) into one
visual language. The aim is for any new slide or deck I build in this
folder to feel consistent with the ones I've already shaped, without
me having to repeat the same preferences each session.

When two layers conflict, the more specific layer wins (this file
overrides the Teaching layer, which overrides the universal layer).
Where this file is silent, fall back to the higher layers or use good
editorial judgment and stay consistent with slides already shipped.

## The goal and two standing rules
Any new slide or deck should feel like it belongs in the same family
as the ones already shipped: same canvas, same chrome, same palette,
same typography, same way of drawing figures and equations. A busy
executive should be able to glance at any slide for 30 seconds and
walk away with the **takeaway in the title** – if not, the slide isn't
ready.

Two standing rules above all else:

1. **Stay as close to the original deck as possible.** The source
   decks reflect classroom-tested choices. Mirror their pedagogical
   structure – text order, wording, accent colors, callout shapes,
   definitions – and deviate only when a slide is explicitly flagged
   for redesign.
2. **Reproduce every table, chart, and equation as a native, editable
   PowerPoint object – never as a screenshot or flattened image.**
   Tables are real PowerPoint tables; charts are built from PowerPoint
   shapes (lines, freeforms, markers, text boxes); equations are
   native OMML / equation-engine objects. The only images that stay as
   images are genuine photographs, logos, news clippings, poll
   captures, and other screenshots that cannot be reconstructed (see
   Pictures). Details in the Equations, Figures, and Tables sections.

## Canvas and master
- **Widescreen 13.33 × 7.5"** (16:9). Convert legacy 4:3 (10 × 7.5")
  decks to widescreen.
- **One single slide master** for the whole deck. Strip the
  python-pptx defaults; one layout keeps the deck coherent. Never mix
  multiple masters.
- **Calibri** for all on-slide text (Cambria Math only inside
  equations). Charts use **Carlito**, the metric-compatible Calibri
  clone, so widths measured in code match what PowerPoint renders.

## Color palette and discipline
Three colors do the work: one strong primary, one warm accent, one
neutral. The primary carries headers / filled boxes / structural
arrows; the accent is reserved for takeaway bars, anchors, and "this
is the point" chrome; the neutral handles captions and secondary text.

The concrete palette used across the harmonized decks:

| Role | Color | Hex |
|---|---|---|
| Primary (headers, filled boxes, structural arrows, axes) | Navy | `0B2B4E` |
| Accent (emphasis bars, anchors, "this is the point") | Gold | `E09F3E` |
| Neutral (captions, secondary text) | Gray | `555B66` |
| Thin rule | Light gray | `C8CDD3` |
| Cream box fill | Cream | `FDF6E6` |
| Dimmed / "off" box | Gray box | `B0B5BC` |
| Pale gold (revenue rectangle fill) | Pale gold | `F6E8C9` |

Discipline:

- **Backgrounds stay white.** Create visual weight with filled boxes,
  not background tints.
- **Gold loses its power if overused** – if more than ~20% of a slide
  is accent-colored, prune.
- **Emphasis rule (important):** for ordinary emphasis in body text,
  use **bold navy**, not orange/gold. Gold is reserved for accent
  chrome, takeaway bars, and anchors.
- **Pair related visualisations by accent color.** When two charts or
  shapes represent the same underlying concept (e.g., a tangent slope
  drawn on a production-function chart and the value of that slope
  plotted on a second MPL chart), use the SAME single accent color
  across both so the eye links them automatically without extra
  annotation.

**Reserved pedagogical colors** (off-limits for structural / chrome
use; for concept-introduction and worked-solution slides only):

- **Concept blue `#0070C0`** (bold) for the **concept name** being
  formally introduced or used (e.g., "Marginal Product of Labor";
  *inelastic*, *elastic*, *unit-elastic*, *substitutes*,
  *complements*, *normal/inferior/luxury good*, *causal effect*). Keep
  it consistent across the deck.
- **Dark yellow `#B8860B`** (italic) for the single **emphasised word**
  inside a definition (e.g., the word "change" in "change in output
  due to a change in labor input").
- Deck-specific accents are fine as examples: e.g., green `#1B5E20`
  for "Market demand"; red `#C0504D` / green `#6E8B3D` for revenue
  lost / gained in a worked example. These are not part of the
  structural primary / accent / neutral palette – don't use them for
  chrome or structural boxes.

## Chrome: top bar, title slide, footer
Keep chrome **identical** across content slides – never enlarge one
slide's title relative to the others.

- **Navy top bar** with a **three-level hierarchical section tag**,
  white bold ~15–16 pt: `Module · Part · Section`, title-cased. This
  is the wayfinding for a long deck. (A deck family that isn't
  module-organized may use a two-level tag instead – e.g.,
  `Demand Analysis · Demand Estimation` – but the default keeps the
  module number.)
- **Action title** as the slide title – the takeaway, not the topic.
  "Prices Coordinate Strangers Without Central Direction" beats
  "Supply and Demand." Navy bold ~30 pt.
- Under the title and above the footer: a **thin gray rule** with a
  short **gold accent strip** on the left.
- **Footer is minimal:** left = optional course footer text (e.g.,
  `Management 405  ·  Demand Analysis`); right = page number only,
  auto-numbered. No "Page X of N", watermarks, or "Confidential"
  stamps.

**Title slide** – centered horizontally **and** vertically, no top
bar and no page number:

- Deck name (e.g., "Demand Analysis") – navy 60 pt bold.
- Section subtitle (e.g., "Demand Estimation") – gold 40 pt bold.
- Short gold strip.
- "Management 405" – gray **bold**.
- "Prof. Nico Voigtländer · UCLA Anderson" – gray regular.

## Typography and text
- **Calibri throughout for slide text** (Cambria Math only inside
  OMML).
- **Short bullets, ~5–10 words each.** No full sentences on the slide.
  If a bullet runs past two lines, split or trim.
- **Handout-legible sizes (go big).** Primary body bullets ~24–26 pt,
  sub-bullets ~21–23 pt. Keep one clear size step between level-0 and
  sub-bullets. Nothing in the body below ~21 pt. Cap ~26–27 pt to
  avoid overflow; render every slide and reduce anything that
  overflows.
- **Equal sizes within a group.** Sibling bullets at the same level
  get the **same** size; numbered steps "1./2./3." are all equal.
- **No awkward line breaks.** Keep key bullets/labels on **one line**.
  To stay big *and* on one line, in order: widen the text column →
  shrink/move the adjacent graphic → only then trim the font. Measure
  with the actual font (PIL ImageFont on Carlito) instead of guessing.
- **Vertically center body text.** Bullet blocks sit vertically
  centered in the content area (equal whitespace above the first line
  and below the last), centered around ~y4.2" between the title rule
  (~1.28") and footer rule (~7.15"). Use a MIDDLE-anchored text box.
  Applies even when text is paired with a side graphic – center the
  text column on its own.
- **Even spacing between main bullets.** When a slide has several
  top-level bullets, distribute the vertical space evenly between them
  – don't strand the last bullet at the bottom. Larger, equal gaps
  between main bullets; a smaller gap before their sub-bullets.
- **Sub-bullets are dark navy, not gray** – one indent level deep, one
  size step down. Don't nest three levels deep – split the slide
  first.
- **Bold = emphasis; italic = variables / captions / "soft" voice.**
  Don't combine bold + italic in body text – reserve that for takeaway
  bars and italic callouts where you want the "this matters" cue.
- **No trailing periods** on bullets/labels (a sentence-final period
  after a lowercase letter, digit, %, ), or ” is stripped deck-wide).
  Captions and running prose keep normal punctuation.
- **Captions: small (11–13 pt), italic, gray, centered.** Same
  treatment for every image attribution.

## Equations and formulas – native OMML, never images
Every equation in a delivered deck is a **native, editable PowerPoint
object** (OMML / equation engine), never a flattened PNG. Use
matplotlib/preview images only while drafting, never as the shipped
artifact.

- **TeX style, always.** Anywhere a formula has subscripts,
  superscripts, fractions, Greek letters, or anything more than
  `a = b + c`, render it with OMML / Cambria Math – not plain text and
  not Calibri math.
- **Variables italic; multi-letter acronyms upright.** Single letters
  like Q, K, L, w stay italic by default (don't add `m:rPr`). Acronyms
  like MRPL, MPL, MC, MR, TR, TFC get `m:sty="p"` so they sit upright
  the way economics journals print them.
- **Stacked fractions for ratios** (e.g., %ΔQ / %ΔP, bang-for-the-buck
  rules). Inline `/` is acceptable only inside running prose.
- **Spell out operators in definitions:** write "divided by", not
  "÷"; spell ratios in words in prose.
- **Elasticity symbol rule (deck-wide):** where a deck uses the
  elasticity symbol, render it with the **D as a subscript – Eᴅ** –
  everywhere it appears (titles, bullets, equations), consistently.
- **Hero formulas get their own box.** A navy filled rectangle with a
  small Calibri label on top and white bold OMML filling the rest is
  cleaner than dropping the equation into surrounding text. Use this
  for the headline equation on any "rule" or "concept" slide (e.g.,
  `Q = a + b·P`).
- **Bullets with embedded variables:** rendering full OMML inside a
  bullet is overkill. Unicode subscripts (`pₖ`, `MPₖ`, `MPₗ`) are an
  acceptable middle ground for in-line text. Standalone formulas stay
  in OMML.
- **Worked-solution slides** mirror the source's "1. / 2. / 3."
  numbered structure and wording. Color-code matched quantities: give
  each conceptual quantity its own color and reuse it everywhere it
  appears (equation + bullets) so the eye links them – but don't
  over-color; only the quantities carrying the pedagogical link.

## Figures and charts – native shapes, never screenshots
Whenever the data exists in code, **build the chart from native
PowerPoint shapes** (editable in PowerPoint), driven by the same
helper that builds any feeding table so the two can't drift apart.

- **Same shadow treatment as pictures.** Charts sit on a soft drop
  shadow – use a white backing rectangle with the shadow applied to it
  (OOXML doesn't let you put a shadow directly on a graphicFrame).
- **Axes:** straight connectors, navy, with a **triangle arrowhead**
  (XML `<a:tailEnd type="triangle"/>`). Y-axis title sits ABOVE the
  top arrow; X-axis title BELOW the right arrow. Axis titles italic
  bold navy; tick labels regular Calibri, navy.
- **Lines** (demand, MC, MR, fitted/regression line): connectors.
  **Curves** (parabolas, MPV, step functions): `build_freeform(...)`
  → `convert_to_shape()` → one editable freeform. **Dashed** guides:
  `<a:prstDash val="dash"/>`.
- **Make each curve independently movable.** Give every curve its **own
  tight bounding box** hugging just that curve – never one big box
  spanning the whole plot (overlapping full-plot boxes make the curves
  impossible to click and drag). Build the path from a **few Bézier
  anchor points** (`quadBezTo` / `cubicBezTo`), not a dense 40-point
  polyline, so "Edit Points" shows a handful of draggable handles. I care
  about being able to grab, reshape, and animate each curve by hand.
- **Bars:** gold-fill / navy-edge rectangles. **Markers / scatter
  points:** small OVAL / RECTANGLE / TRIANGLE shapes – a distinct
  marker shape per series (color alone isn't enough for handout
  printing).
- **Labels live INSIDE the plot area**, floating in clear zones; avoid
  annotation arrows where possible. Label the demand curve "D" at the
  end of the curve; keep Q* as a dashed drop line. Labels must not
  cross or overlap the curves.
- **All figure text large** for handouts: axis titles ~18–20 pt,
  in-chart labels ~16–20 pt; all label text the same size within a
  chart.
- **No horizontal gridlines** unless pedagogically necessary (then
  light gray, dashed). **Round the axis maximum** to a clean number so
  gridlines fall at intuitive intervals (e.g., 0–1000 in steps of 100,
  not 0–982).
- **Legend inside the plot area, top-left, on a white fill.** Don't
  park it below or beside the chart unless space is desperate. White
  fill keeps it readable when it overlaps the rising part of a curve.
- When a chart is the main content it **fills the slide** (~10.5 ×
  5.15") on the white backing rectangle.

## Tables
- **Native tables**, navy header row (white bold text), body rows
  alternating white / cream, navy text, thin borders.
- **Same shadow treatment as pictures** – a white backing rectangle
  with a soft shadow behind the table.
- Strip the python-pptx default table style (no banded theme colors);
  set fills explicitly.

## Pictures
- **Drop shadow + rounded corners by default.** Pictures should feel
  lifted off the slide. Soft shadow, modest blur, slight offset – not
  heavy.
- **Flat exceptions** (no shadow, no rounding); set the image-helper's
  shadow flag to `False`:
  - Logos and brand marks (e.g., Zoom, Anthropic) – flat, no shadow.
  - Book covers, posters, screenshots with their own visible border or
    background – no shadow.
  - Source images that already include a built-in shadow or frame.
- **Keep the original image assets** for screenshots, poll captures,
  news clippings, product photos, and photographs – pull the **actual
  asset** from the source deck, not a re-creation. Rebuild only
  charts / tables / equations natively. Preserve multi-part figures as
  separate shapes so they can be animated one at a time.
- **Prefer real photographs over logos** when illustrating a
  real-world example. A photo of the product or the place beats a
  brand mark every time. No stock photos, clip art, or emojis.
- **Crop tight to the subject** (square stock images have dead space),
  then size/place so the image never overlaps text. If a source footer
  is baked into an image, crop it out of the PNG (LibreOffice ignores
  PowerPoint's crop on preview).
- **Captions:** title-style caption ABOVE the picture (small
  italic-bold navy, centered); source/license BELOW (smaller italic
  gray) – only when it adds information the image or title doesn't
  already carry. **Remove redundant captions / labels / source lines**
  on or around pictures by default.

## Visual hierarchy: boxes, arrows, bridges
- **Filled boxes = primary content nodes** (e.g., a key concept, a
  rule, a definition). Filled in the primary color, white text.
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
  Primary → secondary → action.
- **Arrows carry meaning.** Primary-color arrows for structural flow
  (parent → child, step 1 → step 2). Accent-color arrows for
  "this leads to that" cause-and-effect or "remember this here"
  pointers.
- **Bridge boxes between clusters.** A single outlined box that names
  a relationship and has one inflow + one outflow arrow beats a tangle
  of diagonal cross-cluster arrows.
- **Recurring concept → distinctive shape.** If prompted, pick one
  non-rectangular shape per recurring concept (a 12-point star, a
  parallelogram, a starburst) and reuse it everywhere that concept
  appears – one consistent oddity becomes a wayfinding cue. Don't
  fight the shape with text: if text doesn't fit cleanly, layer a
  separate text box on top – don't shrink the text and don't deform
  the shape.

## Layout patterns I reach for
- **Single slide-layout master.** One layout for the whole deck. Strip
  all the python-pptx defaults; one master keeps the deck coherent.
- **Two-column comparisons** for any "X vs. Y" content (short run vs.
  long run, option A vs. option B). Symmetric column widths, header
  cells on top, parallel bullet structure.
- **Three-card row** for "the three cases" content (e.g., falling /
  flat / rising). Equal widths, even spacing, parallel sentence
  structure inside each card.
- **Definition slides share ONE layout:** a cream rounded formula box
  centered near the top with the native formula centered inside, then
  ~26–28 pt bullets below. All definition slides in a deck match each
  other in box size/position and bullet size.
- **Takeaway bar at the bottom of dense slides.** Accent-color filled,
  primary-color bold italic text, centered. The one-line punchline.
  **Use sparingly** – only when it states a genuine takeaway; do not
  add one by default.
- **Discussion / poll badge** for group-discussion or poll cues. A
  distinctive slanted gold parallelogram, bottom-right corner.
- **Convention callout box.** A small cream-fill rounded rect with a
  thin primary-color border, slight rounding (~6%), soft drop shadow,
  holding a bold primary-color "Convention:" prefix and a single line
  of explanatory text or OMML formula (14–15 pt). Use it to record a
  notational or computational convention adopted in the deck – e.g.,
  "Compute ΔQ and ΔL relative to the initial point." Sits to the right
  of a table or below a hero definition.
- **Concept maps / outline anchors as section anchors.** A
  network-graph-style overview slide at the start of each major
  section, returned to at transitions; the section divider highlights
  the current section (cream band, navy/gold badge, others dimmed).

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
- "Keep the original text/picture" means pull the **exact** wording
  and the **actual** image asset(s) from the source deck.

## Workflow and iteration
1. **Inventory the original first** – dump every slide's text, tables,
   chart data, and image map; render the original deck to images so
   you can *see* every slide.
2. **Decide reuse vs. rebuild:** keep real image assets (polls,
   screenshots, photos); rebuild charts, tables, and equations
   natively.
3. **Build the whole deck** in the clean style in one pass.
4. **Render every slide and review**, then iterate slide-by-slide on
   my feedback. Expect 2–3 rounds of "too cluttered → simplify" on any
   non-trivial diagram slide. When in doubt, step back and ask what a
   busy executive sees in the first 30 seconds – if the answer isn't
   the takeaway in the slide title, the slide isn't ready.
5. **Preserve hand-edits as signal** (see next section).

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
  or duplicate-`<a:effectLst>` audit by default – these add latency
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

## Build mechanics (for the shared toolkit)
The decks are generated with **python-pptx** from a small set of
reusable modules:

- A **base/helpers module** – palette constants, the single master,
  `slide()`, `rect()`, `textbox()`, `bullets()`, `picture()`, boxes
  (`filled_box`, `outlined_box`, `cream_box`), `badge()`,
  `takeaway()`, and `chrome()`. `chrome()` auto-computes the page
  number from the current slide count, so inserting/reordering slides
  renumbers automatically.
- An **equation + figure module** – the `eqn()` / `eqn_centered()`
  engine (stacked fractions built from numerator/denominator text
  boxes + a connector bar), `curve()` freeforms, and a `Fig` class
  providing a logical→inches transform with `axes / line / vdash /
  parabola / point / lbl / brace` primitives.
- A **slides module** per deck – one function per slide plus a
  `main()` that runs them in order and applies deck-wide
  post-processors (trailing-period stripping; Eᴅ subscripting).

**Sandbox / mount caveats:**

- The sandbox mount caps each script file at its **first-write byte
  size** and null-pads the rest. Give new files generous trailing
  comment padding for edit headroom, keep them under the cap, and
  strip nulls (`tr -d '\000'`) when assembling.
- `/tmp` and `/dev/shm` reset between bash calls. Do the whole
  build → convert → render in a single bash call (copy scripts to
  `/dev/shm`, build there).
- **Deploy with an integrity check:** copy to the destination, verify
  the zip (`zipfile.testzip()`), retry on transient lock errors.
- **Render / verify — this machine has NO LibreOffice.** Drive
  **PowerPoint via COM** from PowerShell: export a slide to PNG to *see*
  it, and **open the file in PowerPoint as the real integrity check**.
  python-pptx and `zipfile.testzip()` are far too lenient — this session
  they both accepted a file PowerPoint rejected as corrupt (0x80070570).
  Open in PowerPoint after **any structural edit** (insert / delete /
  reorder slides, add notes or media); kill any stale `POWERPNT` process
  first. (The `/tmp` · `/dev/shm` · `tr -d` bullets above apply only when
  building inside a Linux sandbox, not on this Windows machine.)
- **PowerPoint renumbers part filenames on save.** After I save the deck
  in PowerPoint, `slideN.xml` is **NOT** display-slide N (and
  `imageN` / `notesSlideN` shift too). Resolve display→part by parsing
  `ppt/presentation.xml` `<p:sldIdLst>` order → `r:id` →
  `presentation.xml.rels` (use ElementTree, **not** a regex over the
  rels). Assuming `slideN.xml = display N` corrupted the deck this
  session — a duplicate notesSlide relationship landed on the wrong
  slide.
- **custGeom curve gotchas** (each makes a shape render invisible or get
  dropped): the path segment element is `<a:lnTo>`, **not** `<a:lineTo>`;
  inside `<a:ln>` the child order is fill → dash → join (put `<a:round/>`
  AFTER `<a:solidFill>`); and `<a:prstGeom>` must sit after `<a:xfrm>`
  and before the fill.
- **After I hand-edit the canonical .pptx, edit it in place** with
  python-pptx (swap an image blob, remove a paragraph, etc.) rather
  than rebuilding from scripts – and do it within one bash call, then
  copy back. If reading the file intermittently fails (corrupt central
  directory), ask me to re-save/revert and retry.

## Rebuilding a module from its old deck: phase order (do build.py FIRST)
When reformatting a new module from its old PPT, the phases must run in
this order – and the order is **not reversible**:

1. **build.py scaffold.** Generate ALL script-buildable slides (text,
   bullets, native charts / tables / equations, chrome) in the clean
   style. Finish *every* such slide before moving on.
2. **Freeze build.py.** The moment phase 3 begins, the `.pptx` becomes
   the source of truth and build.py must not be re-run – re-running
   regenerates from scratch and **overwrites everything from phase 3**
   (see the STALE-banner note on the Module 3 build script).
3. **OOXML surgery + hand-edits.** Splice in video / poll / interactive
   slides from the old deck (these **cannot** go through python-pptx –
   it strips NULL video rels and poll `tags`), then fine-tune and port
   my PowerPoint hand-edits.

- **WARN ME if I try to start phase 3 too early.** If I ask you to
  import a video, a poll, or any live-content slide – or to hand-edit /
  do OOXML surgery – while the build.py pass is **not yet complete for
  all slides**, STOP and warn me first: doing so freezes build.py
  prematurely, so any later build.py run would wipe the imported video
  and the edits. Confirm with me that the build.py step is finished for
  every slide before proceeding.

## Things to avoid on every slide
- Walls of text. If a bullet runs past two lines, split or trim.
- Orange/gold used for ordinary emphasis (use bold navy).
- Gray sub-bullets (use dark navy).
- Decorative imagery that doesn't carry information.
- Chart junk – legends that duplicate labels, three-letter gridlines,
  axis titles that repeat the slide title.
- Redundant captions / labels / source lines on or around pictures.
- Stock photos and clip art. Real product / real place / real person,
  or no image at all.
- Emojis anywhere on the slide.
- Trailing periods on bullets.
- "Page X of N" footers, watermarks, "Confidential" stamps.
- Multiple slide-layout masters in one deck.
