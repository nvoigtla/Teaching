# Teaching CLAUDE.md

This file adds teaching-specific instructions on top of the universal
preferences in the parent CLAUDE.md (`h:\Claude Code\CLAUDE.md`).

## Context for This Folder
For work in this folder, you are assisting with **teaching materials**
– primarily PowerPoint lecture slides and related course content.

## Audience
- Students are **Executive MBA students**: experienced senior managers
  and executives, typically mid-career, returning to graduate education
  part-time.
- They are **smart, busy, and skeptical of unnecessary complexity**.
  They want clear takeaways relevant to real business decisions, not
  textbook recitations.
- They have **strong business intuition but variable formal economics
  background**. Assume comfort with general concepts (margins,
  incentives, supply and demand) but introduce formal models, equations,
  and econometric terminology slowly and with motivation.
- They respond well to **case-based reasoning, real-world examples,
  and stories** that connect concepts to executive decision-making.

## Language and Framing
- **Big-picture first.** Every new topic in a slide deck should lead with
  the "why this matters" before any technical content.
- Frame concepts in terms of **strategic implications and decision
  consequences**, not abstract theory. 
- Keep each bullet point short, maximum 2 lines. 
- Keep slides "light", with max. 4 bullet points each, and where possible use graphics and animations
- When technical content is necessary, **explain it very carefully**:
  - Motivate why we need it before introducing it.
  - Walk through it step by step, not in a single dense slide.
  - Tie it back to a concrete example immediately after.
  - Use plain-language definitions for any jargon on first use.
- Avoid graduate-style econometrics terminology unless the lecture is
  explicitly about methods. Prefer "the effect we estimate" over
  "the coefficient on the treatment variable in our specification".

## Slide Design Principles
- **Visuals are the priority.** Executive MBA slides should be heavily
  visual. Each major point should be supported by a chart, diagram,
  image, or schematic – not a wall of text.
- **Every new slide should carry at least one illustrative figure.**
  When I ask you to create a new slide, include at least one relevant
  image or figure (photo, map, chart, diagram, or schematic), with a
  **soft shade**, a **source / attribution line**, and **rounded edges
  where applicable** (photos and maps get rounded corners; charts and
  tables sit on a shadowed white backing card). Only skip the figure
  when the slide is inherently text-only — a section divider, a
  formula / definition slide, or a quiz prompt.
- **Filled boxes and callouts get rounded edges + a soft shade.** Any
  filled box — a concept / "Convention" callout, a takeaway bar, a
  link / "button" box, a back button, and the **header cells of
  two-column / multi-column comparisons** ("Cournot" / "Bertrand" bars
  and the like) — uses slightly rounded corners and a soft drop shadow
  so it reads as a lifted card. (Kept flat, no rounding/shadow: the navy
  top bar, the thin rules and gold accent strips, and table-cell number
  highlights.) **Build-script trap:** the reused Module-3 helper layer
  contains a legacy flat `_add_filled_box` next to
  `_add_rounded_filled_box`; calling the flat one for content boxes is
  how flat headers keep slipping in. In build scripts, treat
  `_add_rounded_filled_box` as the ONLY helper for content-filled boxes
  and audit any remaining `_add_filled_box` call sites (they are almost
  always wrong outside chrome internals).
- **A figure's shade is part of the figure — group them.** Pictures
  carry their own drop shadow directly, so they need nothing extra. But
  PowerPoint can't put a shadow on a table or chart frame
  (`graphicFrame`), so the shade lives on a **separate white backing
  rectangle** behind it. In that case, **group the figure and its
  backing/shade into a single PowerPoint group** (`<p:grpSp>`, backing
  behind + figure in front; group `off/ext` = the two shapes' bounding
  box, with `chOff/chExt` equal to `off/ext` so the children keep their
  absolute positions) so the shade always moves, resizes, and animates
  **with** the figure rather than lagging behind as a stray shape. Do
  this everywhere the split occurs (e.g. every table slide). **Grouping
  invalidates any existing animations on that slide, so re-run the
  animation build afterward.**
- **A colored box and the text inside it are ONE object — group them.**
  Same procedure as the figure-shade rule, applied to callouts: any box
  built as a filled / rounded rectangle plus a **separate** text box on
  top (takeaway bars, cards, question boxes, "Convention" callouts)
  must be merged into a single PowerPoint group (`<p:grpSp>`, box behind
  + text in front; group `off/ext` = the two shapes' bounding box, with
  `chOff/chExt` equal to `off/ext`). Otherwise the box fades in on its
  own and the text lags a beat behind. Do this everywhere the pattern
  occurs (detect the box by geometry — a filled `roundRect` with no text
  of its own — not by shape name, since python-pptx often names rounded
  rects "Rectangle N"). Same two exceptions as the grouping rule in the
  Animations section (table-cell number-highlights; OMML-math callouts).
  **Grouping invalidates the slide's animations, so re-run the animation
  build afterward.**
- **Minimize text on slides.** Aim for short bullets (5 – 10 words),
  not full sentences. Longer explanation belongs in **speaker notes**,
  not on the slide itself.
- **One idea per slide.** If a slide has multiple distinct ideas,
  suggest splitting it.
- **One list = one text box.** When a column or region is really a
  bulleted list, put it in a **single text box with real hanging
  bullets**, not a stack of separate one-line text boxes. One box is
  cleaner to enlarge, edit, and animate, and wrapped lines align under
  the text instead of under the bullet. (I repeatedly ask for these
  stacked one-liners to be consolidated – do it by default.)
- **Title each slide with the takeaway**, not the topic. Bad title:
  "Supply and Demand". Better: "Prices Coordinate Strangers Without
  Central Direction".
- **Use real-world examples** wherever possible – named companies,
  named industries, identifiable events. Generic "Firm A and Firm B"
  examples should be a last resort.
- **Case buildup: chronology first, resolution second.** When a
  mini-case spans two slides, the FIRST slide builds up the situation
  as it unfolded – the facts, the stakes, the two sides – and ends by
  flagging the feature that will matter ("the definition of the market
  would turn out to be crucial for the case"), WITHOUT anticipating
  the outcome. The SECOND slide then shows how the case was resolved
  or decided, with the resolution as the final beat (and the final
  animation click). Never reveal the ending on the setup slide.
- **Case resolutions get the gold takeaway bar.** The outcome line of
  a mini-case ("the court blocks the deal…") is set as a gold rounded
  takeaway bar (~19 pt bold navy text), revealed as the slide's final
  animation click – the visual counterpart of the chronology-first
  rule above.
- **State the provenance of litigation/filing numbers.** When a figure
  comes from a court case or regulatory filing, say how it surfaced
  (internal ordinary-course documents produced in the investigation
  vs. company disclosure vs. press reports) – on-slide when it carries
  pedagogical weight, otherwise in the speaker notes. Never put a
  press-reported magnitude on a slide without marking it "reported".
- **Data visualization should be clean and uncluttered.** Strip out
  chart junk (gridlines, legends that duplicate labels, unnecessary
  axes). Highlight the one element the audience should notice.
- **Consistency matters.** Match colors, fonts, and layout patterns
  across slides in a deck. Do not introduce new visual styles unless
  asked.
- **Sub-bullet sizing — err on the side of LARGER.** Claude has a
  recurring tendency to make sub-bullet text too small for EMBA
  readability. Default to sub-bullets only **~2 pt smaller** than the
  main bullet, NOT 4 – 10 pt smaller.
- **Default bullet sizing for EMBA decks: 24 pt main / 22 pt sub.**
  This is the canonical pair for content-bulleted slides; only deviate
  when a slide has unusually little or unusually much text. Other
  acceptable pairs: 28 / 26, 26 / 24. Sub-bullets at 20 pt or below
  are almost always too small for this audience.
- **Preferred sizing when space allows: 28 pt main / 24 pt sub.** Go
  bigger by default — use **28 / 24** whenever the bullets fit
  comfortably; **fall back to 24 / 22 on more crowded slides**. Either
  way keep an overall balanced look and make sure the larger font never
  spills past the slide borders (measure with the actual font — PIL
  ImageFont on Calibri/Carlito — rather than guessing).
- **Bullet spacing-before: 12 pt before main bullets, 3 pt before
  sub-bullets** (first bullet in a box gets none), as long as the result
  looks balanced.

## Concept-explanation textboxes (preferred format)
- When a slide needs a short, visually-distinct callout to **explain a
  concept**, define a notational **convention**, or record any other
  brief conceptual aside that sits alongside the main slide content,
  use the **cream-fill rounded-rect "Convention" callout** pattern.
- Style:
  - Rounded rectangle, ~12 % corner radius (slight rounding).
  - Cream / soft-yellow fill (e.g., `#FDF6E6`).
  - Thin primary-color border (~1 pt, navy).
  - Soft drop shadow, so the box reads as a lifted card.
  - Primary-color text – bold for any prefix/header (e.g.,
    `Convention:`), regular for the rest.
  - One or two short lines of body text, 14 – 16 pt, left- or
    center-aligned depending on slide context.
- Usage examples:
  - "Convention:  Compute ΔQ and ΔL relative to the initial point."
  - "A production function transforms inputs into outputs. The more
    efficient this process, the higher is productivity."
- Reuse the same visual style across all such callouts in a deck so
  students recognise the box on sight as "this is a concept
  clarification". Course-layer `CLAUDE.md` files can override
  colors / sizes if a course uses a different palette.

## Canvas, palette, and chrome
Any new slide or deck should feel like it belongs to the same family as
the ones already shipped: same canvas, chrome, palette, and typography.

**Canvas and master**
- **Widescreen 13.33 × 7.5"** (16:9). Convert legacy 4:3 decks to widescreen.
- **One single slide master** for the whole deck; strip the python-pptx
  defaults. Never mix multiple masters.
- **Calibri** for all on-slide text (Cambria Math only inside equations).
  Charts use **Carlito**, the metric-compatible Calibri clone, so
  code-measured widths match what PowerPoint renders.

**Color palette** – three colors do the work: one strong primary, one warm
accent, one neutral.

| Role | Color | Hex |
|---|---|---|
| Primary (headers, filled boxes, structural arrows, axes) | Navy | `0B2B4E` |
| Accent (emphasis bars, anchors, "this is the point") | Gold | `E09F3E` |
| Neutral (captions, secondary text) | Gray | `555B66` |
| Thin rule | Light gray | `C8CDD3` |
| Cream box fill | Cream | `FDF6E6` |
| Dimmed / "off" box | Gray box | `B0B5BC` |
| Pale gold (revenue rectangle fill) | Pale gold | `F6E8C9` |

- **Backgrounds stay white** – create weight with filled boxes, not tints.
- **Gold loses its power if overused** – if more than ~20% of a slide is
  accent-colored, prune.
- **Ordinary emphasis = bold navy, not gold.** Gold is reserved for accent
  chrome, takeaway bars, and anchors.
- **Pair related visualizations by accent color** – two shapes/charts for
  the same concept share one accent color so the eye links them.
- **Reserved pedagogical colors** (off-limits for structural / chrome use;
  concept-introduction and worked-solution slides only): **concept blue
  `#0070C0`** (bold) for the concept name being introduced (e.g. "Marginal
  Product of Labor", *elastic* / *inelastic*, *causal effect*); **dark
  yellow `#B8860B`** (italic) for the single emphasised word inside a
  definition. Deck-specific accents (green `#1B5E20` for "Market demand";
  red / green for revenue lost / gained) are fine as examples but are not
  part of the structural palette – don't use them for chrome.

**Chrome** – keep identical across content slides; never enlarge one slide's
title relative to the others.
- **Navy top bar** with a three-level hierarchical section tag, white bold
  ~15–16 pt: `Module · Part · Section`, title-cased. (A non-module deck
  family may use a two-level tag.)
- **Action title** = the takeaway, not the topic; navy bold ~30 pt.
- Under the title, above the footer: a **thin gray rule** with a short
  **gold accent strip** on the left.
- **Minimal footer:** left = optional course footer text; right = page
  number only (live slide-number field, auto-numbered). No "Page X of N",
  watermarks, or "Confidential" stamps.
- **Title slide** – centered horizontally and vertically, no top bar and no
  page number: deck name navy 60 pt bold; section subtitle gold 40 pt bold;
  short gold strip; course line ("Management 405") gray **bold**;
  "Prof. Nico Voigtländer · UCLA Anderson" gray regular.

## Typography and text
(Bullet sizing and spacing-before are under Slide Design Principles; these
are the additional layout rules.)
- **Vertically center body text.** Bullet blocks sit vertically centered in
  the content area (equal whitespace above the first line and below the
  last), around ~y 4.2" between the title rule (~1.28") and footer rule
  (~7.15"). Use a MIDDLE-anchored text box – even when text is paired with a
  side graphic (center the text column on its own).
- **Even spacing between main bullets** – distribute the vertical space
  evenly; don't strand the last bullet at the bottom.
- **No awkward line breaks.** Keep key bullets / labels on one line. To stay
  big AND on one line, in order: widen the text column → shrink / move the
  adjacent graphic → only then trim the font. Measure with the actual font
  (PIL ImageFont on Calibri / Carlito), don't guess.
- **Equal sizes within a group** – sibling bullets at the same level get the
  same size; numbered steps "1./2./3." are all equal.
- **Bold = emphasis; italic = variables / captions / "soft" voice.** Don't
  combine bold + italic in body text – reserve that for takeaway bars and
  italic callouts.
- **No trailing periods** on bullets / labels (a sentence-final period after
  a lowercase letter, digit, %, ), or ” is stripped deck-wide). Captions and
  running prose keep normal punctuation.
- **Captions: small (11–13 pt), italic, gray, centered** – the same
  treatment for every image attribution.
- **Box/callout text floor: 18 pt.** Text inside boxes and callouts
  (teaching cards, quote boxes, diagram boxes, timeline labels) is at
  least 18 pt; chart-internal labels at least 16 pt. Only photo
  attribution and source captions stay at caption size (11–13 pt
  italic gray). If 18 pt does not fit, shorten the text or grow the
  box – don't shrink the font.

## Figures, charts, and tables – native, not screenshots
**Standing rule: reproduce every table, chart, and equation as a native,
editable PowerPoint object – never a screenshot or flattened image.** Tables
are real PowerPoint tables; charts are built from PowerPoint shapes (lines,
freeforms, markers, text boxes); equations are native OMML. The only images
that stay images are genuine photographs, logos, news clippings, poll
captures, and other screenshots that can't be reconstructed (see Pictures).
Use matplotlib / preview images only while drafting, never as the shipped
artifact.

**Charts (native shapes):**
- The shade lives on a white backing rectangle behind the chart (OOXML can't
  shadow a `graphicFrame`); **group the backing with the chart** (see the
  shade-grouping rule under Slide Design Principles).
- **Axes:** straight navy connectors with a triangle arrowhead
  (`<a:tailEnd type="triangle"/>`). Y-axis title above the top arrow, X-axis
  title below the right arrow; axis titles italic-bold navy, tick labels
  regular navy.
- **Lines** (demand, MC, MR, fitted line) = connectors. **Curves**
  (parabolas, step functions) = one editable freeform built from a **few
  Bézier anchor points** (`quadBezTo` / `cubicBezTo`), not a dense polyline,
  so "Edit Points" shows a handful of handles. **Dashed** guides use
  `<a:prstDash val="dash"/>`.
- **Each curve gets its own tight bounding box** hugging just that curve –
  never one big box spanning the plot (overlapping full-plot boxes are
  impossible to grab and drag).
- **Curves must be economically exact, not just suggestive.** Every marked
  point, intersection, or guide must lie mathematically on the curves it
  annotates: a marked Q* sits at the TRUE intersection of the two curves
  (compute it – e.g. solve the Bézier numerically – never eyeball it, and
  make sure the curves actually cut through each other there); dots sit
  exactly on their lines; a dashed guide through declining values passes
  exactly through those values; an annotation arrow ends ON the curve it
  points to (evaluate the curve, don't approximate). When a curve and a
  marker disagree, fix the geometry, not the label. Audit every chart
  slide for this before handing a deck over.
- **A marginal-cost curve that explicitly includes opportunity cost must be
  drawn UPWARD-SLOPING.** Whenever the chart labels the curve "MC (incl.
  opportunity cost)" – or the slide otherwise says the cost of one more
  unit counts the forgone next-best use – the curve rises with quantity,
  never flat and never falling. The best alternative uses are given up
  first, so each additional unit sacrifices a less valuable but still
  positive alternative. A flat MC is correct only when the slide means
  out-of-pocket cost alone. When the slope changes, recompute every marked
  optimum (e.g. the Q* where MPV = MC) as the true intersection with the
  sloped line, per the exactness rule above.
- **Bars** = gold-fill / navy-edge rectangles. **Markers** = small oval /
  rectangle / triangle shapes, a distinct shape per series (color alone
  isn't enough for handout printing).
- **Labels live inside the plot area**, in clear zones; avoid annotation
  arrows. Label the demand curve "D" at its end; keep Q* as a dashed drop
  line. Labels must not cross the curves.
- **All figure text large:** axis titles ~18–20 pt, in-chart labels ~16–20
  pt, uniform within a chart.
- **No horizontal gridlines** unless pedagogically necessary (then
  light-gray dashed). Round the axis maximum to a clean number.
- **Legend inside the plot area** (see the Chart legends section for the
  manual-layout mechanics).
- When a chart is the main content it fills the slide (~10.5 × 5.15") on its
  white backing rectangle.

**Tables:**
- Native tables: navy header row (white bold), body rows alternating
  white / cream, navy text, thin borders. Strip the python-pptx default
  table style; set fills explicitly.
- Same shadow-on-backing-rectangle treatment, grouped with the table.

## Pictures
- **Drop shadow + rounded corners by default** – pictures should feel lifted
  (soft shadow, modest blur, slight offset, not heavy).
- **Flat exceptions** (no shadow / rounding): logos and brand marks; book
  covers, posters, screenshots with their own border / background; source
  images that already include a shadow or frame.
- **Keep the original image assets** for screenshots, poll captures, news
  clippings, product photos, and photographs – pull the actual asset from
  the source deck, not a re-creation. Rebuild only charts / tables /
  equations natively. Preserve multi-part figures as separate shapes so they
  can animate one at a time.
- **Prefer real photographs over logos** – a photo of the product / place
  beats a brand mark. No stock photos, clip art, or emojis.
- **Crop tight to the subject**; size / place so the image never overlaps
  text.
- **Captions:** title-style caption ABOVE the picture (small italic-bold
  navy, centered); source / license BELOW (smaller italic gray) – only when
  it adds information the image / title doesn't already carry. Remove
  redundant captions / labels / source lines by default.
- **Captions sit RIGHT UNDERNEATH their picture** (~0.05–0.1" gap, box
  centered on the picture) – never floating at the bottom of the slide or
  drifting away from the figure they describe. If a picture moves or
  resizes, its caption moves with it.
- **A picture and its caption/source line are ONE object — group them.**
  Any caption, source, or attribution text box that belongs to a picture
  (the footer line beneath it, or a title line above it) is merged with
  the picture into a single PowerPoint group (`<p:grpSp>`, group
  `off/ext` = the pair's bounding box, `chOff/chExt` equal to `off/ext`),
  so the pair always moves, resizes, and animates together. Build scripts
  should do this in a deck-wide post-processing pass (match a picture to
  any small all-italic ≤13 pt text box sitting within ~0.3" of its edge
  and horizontally centered on it). Grouping invalidates existing
  animations on the slide — re-run the animation build afterward.

## Visual hierarchy: boxes, arrows, bridges
- **Filled boxes = primary content nodes** (a key concept, rule, definition)
  – primary color, white text.
- **Outlined boxes = annotations / bridges / "see also"** – white fill,
  accent border, primary-color text; lighter weight.
- **Three-level box rhythm on concept slides:** (1) **Hero** –
  primary-filled box (white bold) with the headline definition; (2)
  **Elaboration** – cream-fill / thin-navy-border rounded rect with the
  decomposition, glossary, and 2–3 bullets; (3) **Action** – accent-filled
  box (primary bold) with the actionable rule / takeaway. Primary →
  secondary → action.
- **Arrows carry meaning:** primary-color arrows for structural flow
  (parent → child, step 1 → 2); accent-color arrows for "this leads to that"
  or "remember this here".
- **Bridge boxes between clusters** – one outlined box naming a relationship,
  one inflow + one outflow arrow, beats a tangle of diagonal arrows.
- **Recurring concept → distinctive shape.** If asked, pick one
  non-rectangular shape per recurring concept (12-point star, parallelogram,
  starburst) and reuse it everywhere that concept appears. Don't fight the
  shape with text – layer a separate text box on top rather than shrinking
  text or deforming the shape (then group the two, per the grouping rule).

## Layout patterns
- **Two-column comparison** for any "X vs. Y" content – symmetric widths,
  header cells on top, parallel bullet structure.
- **Three-card row** for "the three cases" – equal widths, even spacing,
  parallel sentence structure.
- **Definition slides share ONE layout** – a cream rounded formula box near
  the top with the native formula centered inside, then ~26–28 pt bullets
  below; all definition slides match in box size / position and bullet size.
- **Takeaway bar** at the bottom of dense slides – accent-filled,
  primary-color bold-italic centered text, the one-line punchline. Use
  sparingly (only for a genuine takeaway); it gets its own final animation
  click.
- **Discussion / poll badge** – a slanted gold parallelogram, bottom-right
  corner (never coral). A **poll-break** slide's badge reads **"Poll Break"**
  (navy bold Calibri 28 pt on the gold parallelogram), not "Discussion
  Break". Fit the parallelogram to the text: the navy text box is auto-fit
  (`wrap="none"` + `spAutoFit`) at ~70% of the parallelogram width, right
  edge anchored in the corner; measure the label in Calibri Bold, don't
  guess.
- **Convention callout** – see "Concept-explanation textboxes" above.
- **Quote callout** – verbatim quotes (CEO memos, court opinions,
  named executives) go in the cream convention box: italic quote
  (≥18 pt) followed by a bold "— attribution" line one step smaller.
- **Practice-Video link box** – the deck-standard chrome for links to
  practice videos: rounded rect (~28% corner) with a **vertical gray
  gradient fill** (schemeClr bg1 lumMod 65% → 95% → 65%, linear 90°,
  scaled), **gold 1.75 pt border**, soft drop shadow, and a gold "▶" play
  glyph followed by the navy bold label (single shape — text lives inside
  the rect). **Default position: bottom-right corner, overlaying the
  footer** (left 6.92", top 6.83", 5.85 × 0.58" — the slide-24 reference
  position); draw it AFTER the footer so it sits on top of rules and
  page number, and in front of any chart elements. On dedicated
  video-index slides (e.g., "Cournot: Computation") larger centered boxes
  mid-slide are fine; everywhere else use the default corner position.
- **Concept maps / outline anchors** – a network-graph overview slide at the
  start of each major section, returned to at transitions; the section
  divider highlights the current section (cream band, navy / gold badge,
  others dimmed).

## Module-Outline / Agenda Slides (numbered-circle format)
The standard for a module's outline slides — the descriptive overview near
the front, the section agendas at each transition, the wrap-up / post-work
slides, and the summary closer. Finalized on Module 2 (2026-08-16, adopted
from CT's format + Nico's band and spacing rules); apply IDENTICALLY in
other modules. Reference implementation: `make_m2_outline` in
`Module 2/_build_Module2InClass.py` — copy it and swap the item list.

- **Data**: one module-level list of ~6 `(title, description)` pairs; the
  description is one plain-language line.
- **Chrome**: normal content chrome. Top-bar tag `Module N · Outline`
  (summary closer: `Module N · Summary`); action title "Outline of
  Module N" (closer: "Module N: Summary").
- **Item row — identical on every agenda slide**:
  - Gold (`E09F3E`) circle Ø 0.58" at x = 1.15", y = row_y + 0.02";
    the item number centered in it, Calibri **25 pt bold NAVY**
    (`0B2B4E`); no circle outline, no shadow.
  - ONE text box at x = 2.05", y = row_y, w = 11.0": paragraph 1 = the
    title, Calibri **25 pt bold navy** (first letter capitalized);
    paragraph 2 (only when shown) = the description, Calibri **22 pt
    gray** (`555B66`). Space-before 0 on both paragraphs; no bullets.
- **Uniform pitch — reserve the description row on EVERY item**, shown or
  not: title row 0.42" + description row 0.38" + gap 0.11" → pitch
  0.91" per item. Center the block vertically between y 1.60" and
  7.02". Result: item positions are pixel-identical across all agenda
  slides of the deck, so nothing jumps between consecutive slides.
- **Description visibility**: the descriptive overview and the summary
  closer show ALL descriptions; every other agenda slide shows the
  description ONLY for the current topic(s).
- **Current-topic band** (section agendas only — never on the overview
  or the summary): cream (`FDF6E6`) rounded rectangle BEHIND
  circle + title + description: x 0.90", y row_y − 0.06", w 12.15",
  h 0.92", corner adjustment 0.35, gold border 1.0 pt, soft drop
  shadow. One band per highlighted item; no fading of the other items.
- **Pointer / link boxes** (post-work videos, problem set, teaching
  notes): gold-bordered rounded outlined box (white fill, navy bold
  14–15 pt, corner 0.20, soft shadow), two lines — the links line,
  then `On BL under "Module N Post-Work"` — at bottom-right overlaying
  the footer (x ≈ 8.15", y 6.68", w ≈ 4.9", h 0.72"), drawn AFTER the
  footer so it sits in front. **Exception:** when the LAST item is
  highlighted (its band reaches the bottom-right corner), raise the box
  to y 5.30", beside the empty reserved row of the previous item.
- **Grouping pass**: exclude the bands from the box+text grouping rule
  (a filled roundRect wider than ~10" is a layout band, not a callout).
- **Animation**: agenda slides are static (in the SKIP set); wrap-up
  slides may reveal the link box (and any note box) on their own
  clicks.

## Working with .pptx Files
- PowerPoint files are **binary**, so VS Code visual diffs do not
  work for them.
- Before making any edits to a .pptx file, **roll the backups** as
  described in the universal CLAUDE.md: overwrite `Deck_t-2.pptx` with
  the old `Deck_t-1.pptx`, overwrite `Deck_t-1.pptx` with the outgoing
  `Deck.pptx`, then write the new `Deck.pptx`. The two previous
  versions stay in the same folder; the oldest drops off.
- For substantive edits (rewording, restructuring, adding slides),
  **summarize the proposed changes in chat first** in a clear list
  before touching the file. Wait for my confirmation before applying.
- For minor edits (typo fixes, single-word changes), proceed but
  describe what changed afterward.
- **Slide titles and factual content: check with me FIRST.** Never
  change a slide's title, and never "correct" names, places, dates, or
  other facts in slide content, captions, or speaker notes — even fixes
  that seem obviously right, and even in passing while doing other
  work — without asking me first. If something looks factually wrong
  (a mislabeled place, artwork, or person; a note contradicting the
  slide), flag it in chat with the evidence and wait for my
  confirmation; report such findings as proposals, never as done.
  **Exception — obvious spelling mistakes:** unambiguous misspellings
  (e.g. "Ceasar" → "Caesar", "recieve" → "receive") may be fixed
  directly, including in titles; report the fix afterward. This covers
  only typos where the intended word is beyond doubt — anything that
  changes a name, date, or substantive claim still requires asking
  first.
- When editing slides, **preserve the existing visual style** (fonts,
  colors, master slide layout, header/footer) unless I explicitly
  ask to change it.
- When adding new slides, **match the layout of the surrounding
  slides** so the deck feels coherent.
- **A backup link always sits in the lower-right corner.** When a slide
  carries more than one link, the backup / jump-to-slide link takes the
  lower-right corner; a podcast, article, or other external link goes
  wherever it fits best visually — centred in the space that is left, or
  attached to the bullet it belongs to. (Module 1 slide 12 is the
  reference: backup pill in the corner, podcast link centred to its left.)
- **"Back" navigation buttons go in the lower-right corner.** Any
  jump-back button (e.g., on a backup/detail slide that a content slide
  links to) is a navy rounded-rect pill with white bold "← Back",
  placed in the **lower-right corner just above the footer** — the same
  fixed position for every back button in the deck (≈ x 11.72", y 6.6",
  size ≈ 1.55 × 0.46"). It links back to the source slide. Keep this
  position consistent even when it overlays a full-bleed image.
- **Footer page numbers are LIVE slide-number fields, by default.** The
  page number in the footer must be a PowerPoint slide-number field
  (`<a:fld type="slidenum">` in OOXML, or Insert → Slide Number in the UI),
  never a hand-typed static number. Live fields auto-renumber whenever I
  insert, delete, or reorder slides — including slides I add by hand (e.g.
  PollEverywhere activity slides) — so the numbering never drifts. When
  building a new deck, emit the footer number as a `slidenum` field; when I
  hand it a deck with static footer numbers, convert them to fields (touch
  only the footer number run, leave slides that intentionally have no number
  — title, poll, backup — alone). Each field gets its own GUID `id`; keep
  the cached `<a:t>` set to the current number so it looks right before
  PowerPoint recomputes.

## Speaker Notes
- Speaker notes should be **substantive, not bullet repeats**. They will be used for students as guidance when I upload the slides. So they can state again what the slides actually talks about. 
- Default speaker-note style: 2 – 4 sentences per slide, written in
  natural spoken voice (not academic prose), including the key example
  or anecdote to use, and the transition to the next slide.
- If a slide has heavy technical content with little on-slide text,
  the speaker notes should contain the full explanation I'll deliver
  verbally.

## Teleprompter Notes
I tape video lectures and read the script off a teleprompter. When I ask
for a "teleprompter script," it is the verbatim, read-aloud text for each
slide. My preferences:

- **Home = the slides' speaker notes.** Put the read-aloud script in each
  slide's speaker notes (that box below the slide), NOT a separate Word
  doc unless I ask. Notes are per-slide, so the script always matches the
  slide number and shows in Presenter View / teleprompter tools. This
  script replaces the guidance-style notes (it doubles as student guidance
  when I upload the deck).
- **Write for a tired end-of-day viewer.** Assume the person has worked
  all day and is watching a recording. Be clear, well-signposted, and easy
  to follow; guide their attention to what's on the slide.
- **On complex slides, walk through an example.** When a slide is
  technical or abstract, talk the viewer through the concrete example on
  it (the numbers, the named firm, the case) instead of just restating the
  concept. A tired viewer follows a worked example far better than a
  definition.
- **Feel free to add extra examples on complex slides.** Beyond what's on
  the slide, you may bring in one more illustration to help the idea land.
  Search the web for a fitting, current case, or use one from your own
  knowledge. Keep them real and accurate — no invented facts or figures;
  prefer named, verifiable examples, and flag the source (or your
  uncertainty) for anything specific, per the no-hallucination rule.
- **Digits, not spelled-out numbers.** Write "5", not "five"; "28 times",
  not "twenty-eight times"; "50,000", not "fifty thousand"; "16th century",
  not "sixteenth century". Exceptions: "one" stays a word (usually a
  pronoun), discourse ordinals stay words ("the first emperor", "Third
  Italy"), and proper names are never touched ("Two Sicilies"). Keep mixed
  forms consistent ("1 to 5 to 28", never "one to 5 to 28").
- **Verbatim speech ONLY — no stage directions.** Every word in the notes
  must be speakable as-is. Never write instructions to myself like "let the
  relief map do the work", "point to the map", "linger on the images", or
  "name the buildings as you point". If the audience's attention needs
  directing, write the words I would actually say ("Look at the map with
  me…", "Take a moment with these images…").
- **No lecture-management phrases.** Never announce what the narration is
  about to do or frame it as a favor to the audience — no "Let me walk you
  through/down/around…", "let me give you the story behind them", "let me
  pull this together", "What I want you to notice is…", "The theme I want
  you to hold onto is…". Go straight to the substance instead ("The
  structure on the slide has 3 layers.", "Here is the story behind
  them.", "At the top left is…"). Plain signposting ("First…", "Here is
  the point…") and sanctioned attention-directing imperatives ("Look at
  the map with me…") remain fine.
- **Substance over gloss.** When a note compresses a major development into
  one clause (e.g. "by 500 BC Rome had absorbed them"), give it 2–3
  sentences of real, well-established detail on the process instead. Only
  uncontroversial facts — no invented numbers or quotes; when unsure, leave
  the gloss and flag it.
- **Natural, flowing spoken English.** Conversational, first person, as if
  reading to camera. Not choppy or staccato — longer sentences are fine
  when they read naturally.
- **Prefer a full stop over a comma-splice.** When two complete thoughts
  run together with "and" or a comma, break them into two sentences. E.g.
  "…then flattens out, and that flattening is diminishing MPL" becomes
  "…then flattens out. That flattening is diminishing MPL." Likewise start
  a fresh sentence for a contrast: "…held constant. In contrast, labor is
  the input that can still be adjusted." Shorter sentences read better on
  the teleprompter — just don't tip into staccato.
- **Lead with the number the slide highlights.** When a slide visually
  highlights or circles a specific figure (e.g. the boxed 467 on slide 11,
  the first MPL value on slide 15), that highlighted number is the FIRST
  worked example in the script — point at it and read it before making any
  general point. Don't bury it or replace it with a different example.
- **No " – " dash-asides.** Don't set off sub-clauses with dashes; use
  plain sentences instead.
- **Never say "tonight"** (or other live-moment words). The videos are
  watched anytime, so keep it time-neutral ("in this module", "now",
  "here").
- **No unverified emphasis claims.** Do NOT assert things like "this is one
  of the most important ideas in the course", "the key point", or "crucial"
  on your own — those judgments are mine. Don't inherit them from the old
  notes either. Where emphasis might help, flag the spot and ask me first.
- **Skip the poll slides** (PollEverywhere / "respond at PollEv…" and
  poll-break slides) — no teleprompter narration there.
- Work in **verified batches** and confirm the voice on the first batch
  before scripting the whole deck.

## Podcasts (module audio overviews)
When I ask for a "podcast" for a module, I mean an audio conversation I
generate with **Google NotebookLM's Audio Overview** from a source document.
**You don't produce the audio** – you write the self-contained Markdown
**source doc(s)**; I upload each to NotebookLM and hit Generate. Steering
lives in the doc, so it works even if NotebookLM's "Customize" box is hidden.

- **Two episodes per module, by default:**
  - **Intro / preview – about 5 minutes.** *Prepares* students before class.
    Don't be vague – a listener should come away with an **intuitive grasp
    of each core concept**, so briefly explain what each one means in plain
    language (e.g. when you say "economies of scale and scope," say in a
    sentence what each is). What you DO hold back is the **worked examples,
    the specific numbers, and how each case resolves** – those are the
    payoff for class; name the example *types* as illustrations without
    working them through. **Future tense** ("in Module X we'll look at…").
    Open with a line like "Here's a preview of what to expect from
    Module X." Two hosts: **one who did the reading and is prepared**,
    briefing **one who did not and just wants a quick sense of what's
    coming.**
  - **Wrap-up / recap – about 15 minutes.** Lets students *recall* what they
    saw. Covers all the main ideas. **Past tense** ("as we saw in
    Module X…"). Two hosts: **two students who just took the class, talking
    it through** to lock it in.
- **File convention:** one **self-contained** `.md` per episode, named
  `Podcast Module X -- Intro` and `Podcast Module X -- Wrap-up`. Each must
  stand alone (its own instructions + its own content) because each is
  uploaded separately. Draft in Markdown, keep alongside the deck.
- **Document title (H1) must be exactly `Module X - Podcast Intro` /
  `Module X - Podcast Wrap-Up`.** NotebookLM keys the episode/notebook title
  off the doc's first heading, so keep the H1 in this exact, consistent form
  (the titles were drifting when the H1 was something else). A descriptive
  subtitle line underneath (e.g. "PRE-CLASS PREVIEW (about 5 minutes)") is
  fine.
- **Always supply a ready-to-paste Audio Overview prompt for each episode** –
  a short single-paragraph version of that episode's instructions. Give it
  in chat **and** as a labeled block at the very top of the source doc
  ("Audio Overview prompt — paste this into NotebookLM's Audio Overview /
  Customize box:") so I can drop it straight into NotebookLM's Customize box
  as a redundancy on top of the in-doc instructions.
- **Every source doc starts with an "Instructions for the audio hosts (read
  this first)" block** carrying these standing rules:
  - **Audience = "executives pursuing an MBA at UCLA Anderson"** (this
    covers both my Fully Employed and Executive MBA sections). Never say
    "Executive MBA students."
  - **Call it "Module X," never a "masterclass."**
  - **Refer to the material as "the class talks about…" / "Module X
    covers…"** – never "the source material," "this document," or "the
    notes."
  - **Concrete numbers only when the example truly needs them** (e.g. keep
    exact salaries in a talent-poaching case, dollar figures in a
    package-pricing or loan example, cost-per-seat-mile in an aviation
    comparison). Don't recite fine-grained figures where the idea stands on
    its own (e.g. don't quote that marginal product "falls from 0.66 to
    0.04").
  - **Explain decisions by intuition, not arithmetic** (hiring = "does the
    next worker bring in at least as much extra revenue as we pay them?").
  - **Always be clear about context, including which regime a rule lives
    in.** State explicitly that the hire-until-MRPL-=-w rule is a
    **short-run** rule (capital fixed) while the bang-for-the-buck input-mix
    rule is a **long-run** rule (both inputs flexible), and remind the
    listener periodically which world they're in.
  - **If you're unsure how an example should be framed, ask me before
    writing it — don't guess.** Get the *point* of each example exactly
    right. E.g., Waterworld's disaster was the runaway, over-budget
    **production cost** (now sunk), not weak box-office revenue; saying only
    "the movie was a disaster" is ambiguous and misleads.
  - **Convey each concept correctly — don't overstate the result.** The
    hosts ad-lib and tend to exaggerate. Guard the classic slip:
    **diminishing** marginal product means output keeps **rising**, just by
    smaller and smaller amounts as you add workers — it does NOT mean output
    falls or that "adding workers slows production." Output only actually
    declines under **negative** marginal product, which this course does not
    reach unless we deliberately exaggerate. Keep those two distinct, and in
    general state each result no more strongly than the economics supports.
  - **Keep returning to the module's one unifying idea** (the throughline).
  - **Lead with real-world stories** and let them carry the ideas.
  - Warm, curious, conversational tone – smart colleagues (or two students)
    connecting the dots, not a lecture; define terms in plain language, go
    light on formulas.
- **Usage / mechanics:** put **each episode in its own NotebookLM notebook**
  (NotebookLM blends all sources in a notebook into one audio), so the two
  files must never share a notebook. There is **no editor for the finished
  audio** – to change an episode, edit the source doc (especially the
  instruction block) and regenerate. Length isn't exact; the "about N
  minutes" instruction plus a correspondingly short/long source doc pushes
  it the right way. No invented facts or numbers, and spot-check the
  generated audio since the hosts ad-lib.

## Wrap-Up Video (NotebookLM Video Overview)
A module can also get a **video** wrap-up, produced with **NotebookLM's
Video Overview** from a source doc I write (I don't produce the video). It's
the video sibling of the Wrap-up podcast: a ~15-minute, **past-tense** recap
for students to recall the module.

- **Reuse the Wrap-up podcast body verbatim** (same corrected content), with
  a **video-specific** instruction block and a ready-to-paste **Video
  Overview prompt** at the top of the doc (same paste-in convention as the
  podcasts — give it in chat *and* in the doc).
- **All the standing podcast rules apply** (audience = executives pursuing an
  MBA at UCLA Anderson; "Module X," not "masterclass"; "the class covered…,"
  never "the source material"; numbers sparingly; intuition over arithmetic;
  be explicit about which regime a rule lives in — MRPL = w is short-run,
  bang-for-the-buck is long-run; precise example framing, e.g. Waterworld =
  runaway *production* cost, not weak box office; diminishing ≠ negative
  marginal product; keep returning to MB = MC).
- **Video-specific:** tell it to let clean visuals carry each point — one
  simple diagram / label / image per idea (a rising-then-flattening output
  curve, a U-shaped average-cost curve, a short-run vs. long-run split, the
  named companies for each example) — and NOT to crowd the frame with text.
  A single narrator or two voices are both fine, as long as it stays past
  tense.
- **File + title convention:** name the file **`Video Module X -- Wrap-up`**;
  set the H1 title to **`Module X - Video Wrap-Up`**.
- **Set expectations — this is a supplement, not the real lecture video.**
  NotebookLM generates its OWN visuals from the text; it does **not** use my
  slides or animations. Spot-check both the narration and the AI-made visuals
  for distortions. Video renders are slower than audio and may have a smaller
  daily quota. NotebookLM auto-titles the output, so rename the downloaded
  file to `Module X - Video Wrap-Up`. Put it in its own notebook.

## Animations
I like slides to build up step by step so the audience (often watching a
taped video) follows one idea at a time. Calibrated defaults:

- **Tell the story of the slide.** Put yourself in my head as the
  instructor and reveal the bullets and figures in the order the story
  unfolds as I talk through it — not just mechanically top-to-bottom.
  Group whatever belongs to one beat of the story so it appears together.
- **Effect: Fade, about 0.5 s, on click.** Fade is the default entrance for
  every revealed element. Not instant "appear," not flashy motion. Keep one
  effect style across the whole deck.
- **Chrome stays put.** The top bar, section tag, title, thin rule, footer,
  page number, and a chart's axes and axis labels are visible from the
  start and are never animated.
- **Two figures on one slide should not overlap at all — and if they must,
  the later one has to cover the earlier one completely.** First try to
  size and place them so they sit apart: shrinking both and stacking them
  is the better answer whenever the slide has room (the fox and the
  hedgehog on Module 1 slide 13 were fixed this way). The coverage
  requirement applies only where the figures genuinely overlap, i.e. where
  a build deliberately paints successive versions of the SAME figure on
  the same spot (Module 1's LA-to-SF map sequence, the two-stage copper
  figure). There, size and place each revealed figure so its bounding box
  contains every figure revealed before it: a printed handout shows the
  final state, so an uncovered strip of an earlier figure reads as a messy
  fringe. Check the whole chain, not just the last pair — figure 2 must
  cover figure 1, and figure 3 must cover both. When they differ in aspect
  ratio, grow the later figure (keeping its own proportions) rather than
  cropping or stretching it, and re-check that it still sits inside the
  content area.
- **Build the content, one step per click:**
  - Text slides: reveal one bullet (top-level point) per click.
  - Charts / diagrams: reveal the pieces one at a time, and reveal each
    curve or series together with its own label on the same click. Follow
    "guides before regions" (a dashed guide line before the shaded area it
    marks), and reveal the building blocks before the synthesis (e.g. the
    individual short-run curves before the long-run envelope).
- **First bullet shows WITH the slide — don't animate it.** On a
  text/bullet slide, the FIRST top-level bullet is visible the moment the
  slide appears (no click); the build starts from the second bullet, so
  the slide never opens completely empty when its content is bullets.
  **Exception:** slides whose build is of pictures, boxes, formulas, or a
  diagram (e.g. the Big-Picture map, the concept map, definition/formula
  slides) may still open empty and fade everything in — there, the
  empty-then-build is the point.
- **Group things that belong to one beat on a single click:**
  - a graphic and its own label;
  - a picture and its source / attribution line beneath it — **always**
    reveal the source at the same time as the picture, on **every** slide
    (e.g. slide 47) — never let the picture land a click before its source;
  - **a figure and its label / legend reveal on the SAME click — never
    staggered.** This includes legend cards, callout labels, and source
    lines, and applies equally on figure-only slides with no bullets (a
    map plus its legend plus its source line = ONE click, not three);
  - **one caption per picture, revealed WITH that picture.** Never use a
    single combined caption for several pictures ("A · B" at the bottom of
    the slide); split it so each description sits with its own picture and
    fades in on that picture's click;
  - an image and the text box that gives its title, header, or context —
    reveal them together, whether that text sits **above** the picture (a
    header/label) or below it (a caption) (e.g. slides 10, 44);
  - **a whole labelled panel reveals as one beat.** When a picture has a
    header above it and/or a caption/source below it, the header +
    picture + caption/source are ONE click, not three. In a **multi-column
    / multi-panel** layout (e.g. slide 10's short-run vs. long-run
    columns), reveal each column as its own single beat — header + picture
    + source together — and go **column by column**. Never reveal all the
    headers first and then all the pictures.
  - a context picture and the bullet points it supports — reveal them
    together so they land as one thought (e.g. slide 47);
  - a formula and the description of its terms (its glossary / legend) —
    reveal them together (e.g. slide 33).
- **Don't trust a generic shape-by-shape auto-rollout on picture slides.**
  A blind "one body shape per click" pass has no idea that a text box is
  the *title* of the picture next to it, so it splits them onto separate
  clicks (this is exactly how slide 10 ended up with headers on clicks 1–2
  and pictures on clicks 3–4). Any slide with pictures gets a **custom
  per-panel grouping** — bind each picture to its adjacent header and
  caption/source before assigning clicks. Verify picture slides by eye, not
  just by click count.
- **A rounded/filled shape and the text box layered on it are ONE object.**
  When a callout is built as a background rounded-rect plus a separate text
  box on top (the "layer a text box on top" pattern), merge the two into a
  single PowerPoint **group** (`<p:grpSp>`) so they move, resize, and
  animate as one — never as two shapes that can fade in on separate clicks.
  Do this **everywhere** the pattern occurs, not just where it's currently
  wrong. Group offset/extent = the children's bounding box, with
  `chOff/chExt` equal to `off/ext` so the children keep their absolute
  positions.
  - **Two exceptions to the grouping (do NOT group these):**
    - **Table-cell number-highlights** — the little rounded box behind a
      single computed value in a table column (e.g. `$29,700`, `0.660`).
      That's a different device from a callout: it stays locked to its cell,
      isn't dragged around, and already co-reveals with its column. Leave
      it as two shapes.
    - **Shapes whose text is native OMML math wrapped in
      `mc:AlternateContent`** (the `a14` namespace is declared on the inner
      `mc:Choice`, outside the `<p:sp>`). Extracting just the `<p:sp>` to
      group it orphans that namespace and the equation fails to render.
      Leave such math callouts ungrouped; just make sure they co-reveal.
      (When extracting a `<p:sp>` for grouping, match the **balanced**
      closing tag by depth — an `mc:Fallback` can nest another `<p:sp>`.)
- **Worked-computation / derivation tables build in the order you'd teach
  them, not row-by-row top-to-bottom** (e.g. slide 15's MPL table). The
  pattern I want:
  - **Setup stays on screen from the start** — title, subtitle, and the
    empty/partly-filled table are visible before the first click; the
    build fills in the computed values.
  - **Teach the mechanic once, in full, on the FIRST instance.** Reveal the
    first cell's whole apparatus together — the arrow / connector, the line
    that points to the result, and the result itself (e.g. the down-arrow
    0→165, its rounded line, and the ΔQ value 165).
  - **Attach the explanation to that first instance.** Immediately AFTER the
    first computed value, reveal the convention / legend / definition box
    (with the connector that points to it) — it explains the step just
    shown. Not up front, not saved for the end.
  - **Do the second instance in full too, then BATCH the rest.** Second cell
    gets its own click (arrow + line + value); then one click reveals all
    remaining arrows + lines + values in that column at once.
  - **Go column by column, and introduce each new column's first value
    together with the box/formula that defines it** (e.g. the first MPL
    value revealed with the `MPL = ΔQ/ΔL` box + its connector), then batch
    the rest of that column.
  - **End with the concluding observation** — the summarising arrow and its
    note box (e.g. the down-arrow under the MPL column + "MPL is declining
    as we add workers") as the final click.
- **The MB=MC star reveals AFTER the rule it abstracts, never before.**
  The recurring 12-point "MB = MC" (or "MB > MC") star is the concept the
  concrete decision rule is an instance of. Reveal the concrete rule first
  (e.g. slide 19: "If MRPL > w, hire more"), then reveal the star as the
  "…and this is really the MB = MC idea" payoff on the next click. Same on
  slide 23 (optimal-workers rule → star) and slide 24 (optimal-hiring
  interval → star). Show the star together with its little label and its
  connector line to the rule.
- **The takeaway / conclusion bar gets its own final click** so the
  punchline lands last.
- **Skip (no animation):** the title slide, the agenda / Part-X roadmap
  slides, poll slides (PollEverywhere), the embedded video slide (it has
  its own click-to-play trigger), and BACKUP / backup slides.
- **Build mechanics (this machine):** inject a `<p:timing>` block via OOXML
  surgery (no LibreOffice here). I can't watch playback, so verify the
  effect count and targets via PowerPoint COM
  (`Slide.TimeLine.MainSequence`), confirm the file opens in PowerPoint,
  and have me eyeball the slideshow. Work in verified batches.
- **Reuse the animation engine, don't rewrite it.** `_animate.py` (first
  built for Italy IBR, adapted for Module 7) is the proven generator: it
  emits timing XML byte-pattern identical to what PowerPoint itself writes
  (verified by diffing against a native `MainSequence.AddEffect` save).
  Per-slide story plans use a small selector language (`t:PREFIX` /
  `t:PREFIX#n` / `pic:N` / `cxn:N` / `osp:N` / `pr:BOX:st:end`); anything
  not selected stays visible from the start, which is exactly how chart
  axes/ticks stay static. Run it as a rerunnable post-build pass.
- **Animation-engine gotchas (learned on Module 7):**
  - A single effect targeting a MULTI-paragraph range (`pRg st≠end`) gets
    re-expanded by PowerPoint into one effect per paragraph on separate
    CLICKS. To reveal a text section as one beat, emit one effect per
    paragraph inside the same click group (first `clickEffect`, rest
    `withEffect`).
  - Verify CLICK STRUCTURE, not just effect counts: iterate
    `MainSequence.Item(i).Timing.TriggerType` (1 = on-click, 2 =
    with-previous) and compare on-click counts to the plan.
  - Detect chrome by GEOMETRY (top-bar/rule/footer bands, page-number
    position), never by shape name — python-pptx names content shapes
    "Rectangle N" too, so name-based chrome rules eat payoff cells and
    callouts.
  - `t:` text selectors match first-unused in document order — order
    "50 units" before "50" in a beat list, and use `t:PREFIX#n` when the
    same tick label ("200") appears several times on one chart.
- **The slideshow renderer is a SEPARATE surface — verify it too.** The
  editing canvas, PNG export, and COM slideshow *stepping* can all pass
  while the real full-screen Slide Show crashes ("The slide failed to open
  properly" — and the failing slide may be ANY slide in the deck, since
  add-ins scan the whole deck at show start; the banner just surfaces on
  the current slide). I CAN see what the user sees: open the deck via COM
  with a visible window, run `SlideShowSettings.Run()`, and screenshot the
  `screenClass` window via `PrintWindow` (flag 2 for DirectX content);
  classify pass/fail by pixel and bisect with `sldIdLst`-subset copies.
  After any media/OOXML surgery, run at least one full-deck slideshow
  probe before handing the deck over.

## Drafting Workflow for Slide Content
- For new slide content or substantial restructuring, **first draft
  the deck outline in Markdown** (one section per slide, with bullets
  and speaker notes) so I can review and iterate quickly using visual
  diffs in VS Code.
- Once the Markdown outline is approved, then build or update the
  .pptx file to match.
- Keep the Markdown outline in the folder alongside the .pptx so
  future revisions can edit the outline first.
- File-naming convention for drafts:
- If I give you an initial slide deck, use that deck's name. Otherise, ask me for a "Slides Name."
  - Outline: `[Slides Name] - outline.md`
  - Deck: `[Slides Name].pptx`
  - Previous versions: `[Slides Name]_t-1.pptx`, `[Slides Name]_t-2.pptx`

## Example-Candidates Workflow (new real-world examples)
When I ask for fresh, current examples for a module's concepts:
- Research with parallel agents against primary / tier-1 sources
  (filings, court documents, official data, tier-1 press); separate
  confirmed from press-reported figures explicitly.
- Deliver a separate review deck named `Module X - Example
  Candidates.pptx` – one slide per candidate: concept tag in the top
  bar, fact bullets, a cream teaching-angle card (with the proposed
  visual), a gold "Discussion:" prompt, and a source line; full URLs
  and confirmed-vs-reported flags go in the speaker notes.
- Mark press-reported figures "reported" on the slide itself; list in
  the notes exactly what must be re-verified before adoption.
- I pick the winners; only then do candidates graduate into the main
  deck with the full visual treatment (images, native charts,
  animations, renumbering via the pipeline).

## Folder Structure
- Each distinct course gets its own subfolder under `Teaching\`
  (e.g., `Teaching\405-Fall-2026\`, `Teaching\Macro-EMBA-Spring-2027\`).
- Within a course folder, organize by lecture or topic as I direct.
- **A `Session-Notes.md` lives in each distinct deck / module
  subfolder** – one per deck or course-unit (e.g. `Module 3/`,
  `Italy IBR/`), not one shared file at the parent level. A folder that
  merely groups several distinct decks does **not** carry a shared
  Session-Notes; each deck keeps its own.
- **When a session works on a given deck, read and update the
  `Session-Notes.md` in THAT deck's subfolder.** If the subfolder
  doesn't have one yet, create it. Never merge notes for different
  decks / courses into a single file. (A small course folder holding a
  single deck can keep one Session-Notes at its top level.)
- Shared materials (general visual templates, recurring case examples,
  reusable diagrams) can live in `Teaching\Shared\` and be referenced
  from any course.

## Default Behavior for Slides
- **Never produce a full slide deck unprompted.** When given a topic
  or source material, ask a clarifying question first about scope,
  level, lecture length, and where it fits in the course.
- For a new lecture, propose an **outline first** (slide-by-slide
  titles and one-line summaries) before drafting any content.
- Estimate slide count from lecture length: roughly **one slide per
  2 – 3 minutes** of lecture time for EMBA pacing, since slides are
  visual-heavy and discussion-anchored.

## When working on PowerPoint slides...

### Formulas
- **Use OMML / Cambria Math, not plain text**, anywhere a formula
  contains subscripts, superscripts, fractions, or Greek letters.
  Convert things like `p_K`, `MP_L`, `MRPL = w` to proper math runs.
- **Variables italic, acronyms upright.** In OMML, set `m:sty=p` for
  multi-letter acronyms (MRPL, MPL, MC, TFC). Single-letter variables
  (Q, K, L, w) stay italic by default. This matches journal-style
  notation and is the cue economists expect.
- **Indexed symbols get a real subscript: italic letter, subscript
  index.** Write P₀, Q₁, D₂, S′ the way TeX would set `$P_0$` — the
  letter italic, the index a true subscript. Outside formulas use
  PowerPoint's own subscript (an OOXML `baseline` on the index run),
  not a raised-looking digit typed inline and not a Unicode subscript
  character. This applies everywhere the symbol appears: chart axis
  labels, curve labels, table cells, and mentions inside running bullet
  text ("the market clears at a higher price P₁"). Any letter standing
  for a price, quantity, or curve is italic even without an index.
  Module 1's build does this in a deck-wide pass over every text run
  (`apply_symbol_subscripts`), keyed to the letters P, Q, D and S so
  ordinary text is never touched — copy that approach rather than
  editing label call sites one by one.
- **TeX style, always.** Anything beyond `a = b + c` is OMML / Cambria
  Math, never plain text or Calibri math.
- **Stacked fractions for ratios** (e.g. %ΔQ / %ΔP); inline `/` only inside
  running prose. Spell operators out in definitions ("divided by", not "÷").
- **Hero formulas get their own box** – a navy filled rectangle with a small
  Calibri label on top and white bold OMML filling the rest, for the
  headline equation on a "rule" / "concept" slide (e.g. `Q = a + b·P`).
- **Bullets with embedded variables:** Unicode subscripts (`pₖ`, `MPₗ`) are
  an acceptable middle ground inline; standalone formulas stay OMML.
- **Elasticity symbol (econ decks):** render with the **D as a subscript –
  Eᴅ** – everywhere it appears (titles, bullets, equations), consistently.
- **Worked-solution slides** mirror the source's "1. / 2. / 3." numbered
  structure and wording; color-code matched quantities (each quantity its
  own color, reused in equation + bullets), but only the quantities carrying
  the pedagogical link.

### Chart legends
- **Legends stack vertically, one entry per line.** When a chart has
  more than one series, the legend entries should appear vertically
  one under the other – never side-by-side. Use a narrow + tall
  manual-layout box (e.g., `w ≈ 0.15`, `h ≈ 0.07 × N_entries`) so
  PowerPoint is forced to render single-column.
- **Entries packed close together.** Don't pad the legend box with
  extra vertical whitespace; the entries should sit just a small
  gap apart. The user can read three TC / TFC / TVC labels in
  ~0.22 of chart height comfortably.
- **Legend lives inside the chart**, not below or beside it. Set
  `chart.legend.include_in_layout = False` so the legend overlays
  the plot area in an empty corner (typically upper-left when the
  curves rise from left to right). Place it where it does NOT
  occlude any series.
- **White fill + thin primary-color border.** A 0.75 pt navy border
  with a solid white fill makes the legend read as a self-contained
  badge even when it overlaps gridlines or low data points.
- **18 pt navy Calibri text by default** – matches the deck's
  oversized-for-EMBA-readability axis-label scale.

### Workflow with existing .pptx decks
- **Never round-trip an existing deck through python-pptx.** It
  silently strips NULL hyperlink rels and other elements PowerPoint
  expects, corrupting the file. For modifications, use direct zip +
  lxml surgery on the OOXML parts.
- **All EMU values must be integers.** Decimal EMUs break PowerPoint
  silently (the file opens but shapes vanish or misposition). Always
  wrap computed positions in `int(...)`.
- **The build script is the source of truth.** When the user makes a
  manual tweak in PowerPoint (resized box, deleted variable,
  repositioned label), preserve it in the build script so the next
  rebuild doesn't undo their work. Note these as visual-preference
  signals, not edge cases.
- **Single-layout master.** For new decks, keep one slide layout for
  the whole deck. Multiple layouts invite drift and make consistent
  rebuilds harder.

### Reformatting an existing deck vs. creating new content
- **When I supply an existing deck to "rebuild in the new format",
  the job is formatting, NOT rewriting.** Preserve the original's
  slide titles, bullet wording, structural framing, and pedagogical
  examples verbatim where possible. Apply only the new visual layer:
  palette, layout primitives, OMML for math, drop shadows, section-tag
  hierarchy, etc.
- **Three allowed kinds of deviation from the original:**
  - **Refreshed examples** I have explicitly retired (e.g., Tesla →
    Rivian, iPhone 11 → 17, Burn60 → ChatGPT). Don't sweep examples on
    your own initiative; confirm, or follow established prior choices
    for that deck.
  - **Numerical updates** — currency, wages, prices brought to today.
    Same rule: established prior choices propagate; don't invent new
    ones.
  - **Documented corrections** I have flagged in the source (e.g., a
    CORRECTION slide noting a math error).
- **Customizations already in the new deck stay.** Concept-map slides,
  MB=MC anchor patterns, merged hero-concept slides, or any element
  that doesn't exist in the source but has been added in our work —
  preserve through subsequent rebuild passes.
- **When in doubt about wording, lean toward the original.** The new
  deck is a reformat of my own pedagogical material; rewriting prose
  is not the goal and risks introducing subtle changes I may not want.
- **Copied text keeps its run formatting — italics, bold, underline.**
  When porting text from the original slides, carry over the
  character-level emphasis on each run (bold, italic, underline —
  also strikethrough and colored words), not just the words. This
  emphasis is deliberate pedagogy (the stressed word in a definition,
  the underlined contrast), and flattening it to plain text loses
  meaning. When rebuilding via a script, extract the original's run
  properties from the slide XML and reproduce them run by run; don't
  retype from a text-only dump.
- **Speaker notes:** preserve substantive notes from the original
  verbatim; only rewrite when the original notes are sparse, missing,
  or contradicted by the slide content (per the existing
  source-vs-notes conflict rule).
- **Section dividers:** the new deck uses fewer dividers than the
  original deck's recurring "outline" checkpoints. Prefer the new
  deck's consolidated dividers over reinstating every original
  outline-of-module slide.

### Rebuilding a deck that has animations, hidden slides, or live content
When the source deck I hand you includes builds/animations, hidden slides,
polls, or videos, faithfulness extends beyond text and figures:

- **Match the animation choreography slide-by-slide.** Reproduce the
  original's build order and click count – the exact interleaving of graph
  pieces, labels, and text. Per-bullet builds are fine on text slides; match
  click counts where they carry pedagogical meaning.
- **Guides before regions.** Reveal a dashed guide line to a value first,
  then the shaded area / region it annotates – never the fill before its
  guide.
- **Keep figures economically correct, not just visually matched.** E.g., a
  kinked joint-demand's top segment equals the higher-cost firm's demand and
  the flat part is the horizontal sum; demand and MR share the same vertical
  intercept. Correctness beats pixel-matching.
- **Preserve hidden slides** – rebuild them and keep them hidden; un-hide
  only temporarily for review, then re-hide.
- **Preserve live / interactive content** – PollEverywhere slides and
  embedded or online (YouTube) videos must stay functional. Never drop poll
  URLs, poll `tags` relationships, or video links; size online videos as in
  the original.
- **Keep slide count and order identical** for a faithful rebuild of a
  specific deck, so speaker notes and any spliced-in original slides line up
  – **unless I explicitly ask to add, delete, or renumber slides** (that
  overrides this; and see the "Section dividers" note above for the divider
  exception).
- **Toolchain-independent build gotchas:** large decks can read back
  corrupted over a working mount – save the final file straight to its
  destination and verify it re-opens; and large in-place edits to a generated
  build script can truncate – prefer writing a fresh versioned file. (The
  specific way animation timing gets injected depends on the machine's tools
  – a LibreOffice "normalize-then-inject" round-trip where LibreOffice is
  installed, or direct OOXML `<p:timing>` surgery where it isn't; follow the
  course layer for the current machine.)

### Rebuilding game-theory / payoff-matrix decks
These conventions apply **only** when rebuilding a game-theory deck (payoff
matrices, best responses, Nash equilibria); ignore them for other decks.

- **Payoff matrix (2×2) layout.** Column player's name centered on top
  (accent color); row player's name rotated at the left (a reserved concept
  color); strategy labels around a 2×2 grid of white cells with navy borders.
  Each cell shows *"a , b"* with player-1's payoff and player-2's payoff in
  the two players' respective colors. Caption below: "Payoffs to (Player 1,
  Player 2)."
- **Best-response method.** For each column, draw an arrow (player 1's color)
  to player 1's best row and circle that number; for each row, draw an arrow
  (player 2's color) to player 2's best column and circle that number. A
  **Nash equilibrium** is a cell where both numbers are circled – draw a
  larger oval plus a "Nash equilibrium" callout. No cell with both circled ⇒
  **no pure-strategy equilibrium** (say so on the slide). Multiple equilibria
  ⇒ a "*N* Nash Equilibria" gold pill.
- **Animate the matrix step-by-step.** Reveal one step at a time – **arrow
  first, then its circle**, a separate click each, for all four best
  responses. The **equilibrium is the finale:** on the last click, reveal
  together the Nash oval, a "Nash equilibrium" conclusion box on the right,
  and a connector line from that box to the equilibrium cell, with a slow
  (~1.4 s) fade-in so it lands as the punchline.
- **Worked-solution answers.** Step through the algebra; put the **final
  numeric answer in deep red**.
- **Discussion / poll badges** use the gold parallelogram style (never coral).

### Iteration is the norm
- **Expect 2 – 3 rounds of "too cluttered → simplify"** on any
  diagram slide. Don't try to land it in one shot. Propose a layout,
  build, look at it through EMBA eyes, cut.
- **When in doubt, step back into the student's shoes.** What does
  an executive see in the first 30 seconds? If the answer isn't
  immediately the takeaway in the slide title, the slide isn't ready.

### Conflicts between source slides and notes
- **When a previous deck's speaker notes contradict its slide
  content, prefer the slide content.** The slide is what was actually
  shown; the notes may be outdated drafts. Flag the discrepancy in
  chat so I can decide whether the notes had a good reason.

### Build discipline: verification and manual tweaks
- **Default: rebuild the canonical deck in place, no verification.** The
  build script is the source of truth and the start-of-day Git snapshot is
  the safety net – write straight to the canonical filename and stop. Don't
  run a python-pptx readback, footer-page-number check, or
  duplicate-`<a:effectLst>` audit by default.
- **Opt-in verification when I report a problem** ("the page numbers are
  off", "PowerPoint won't open it", "shape X disappeared") – then use the
  readback to diagnose, fix the script, and rebuild.
- **Opt-in side-path only when I signal hand-edits** (e.g. "I tweaked slide
  12 by hand"): build to a side path (`<deck>_test.pptx`), diff against the
  canonical file to surface the hand-edits and port them into the script,
  re-run to the side path, verify, then move it over the canonical deck.
  Don't invoke this flow on your own.
- **Preserve hand-edits as signal** – port each manual tweak into the build
  script with a one-line comment giving the prior value and date (e.g.
  `tbl_top = Inches(2.45)  # hand-tweaked from 2.85 on 2026-05-12`).
- **Exceptions require confirmation.** Anything outside those two opt-in
  cases – extra files (`_test`, `_temp`, `_v2`), forced moves (`mv -f`),
  parallel scripts, hidden readbacks – stop and ask first.

### Build mechanics and machine gotchas
(These extend "Workflow with existing .pptx decks" – never round-trip through
python-pptx, integer EMUs only, one master.)
- **Reusable modules:** a base / helpers module (palette, single master,
  `slide()`, `rect()`, `textbox()`, `bullets()`, `picture()`, `filled_box` /
  `outlined_box` / `cream_box`, `badge()`, `takeaway()`, `chrome()` –
  `chrome()` auto-computes the page number from the slide count); an
  equation + figure module (the `eqn()` stacked-fraction engine, `curve()`
  freeforms, and a `Fig` class with a logical→inches transform and
  `axes / line / vdash / parabola / point / lbl / brace` primitives); and a
  per-deck slides module (one function per slide + a `main()` applying
  deck-wide post-processors such as trailing-period stripping).
- **Render / verify — this machine has NO LibreOffice.** Drive PowerPoint via
  COM from PowerShell: export a slide to PNG to *see* it, and **open the file
  in PowerPoint as the real integrity check** after any structural edit
  (insert / delete / reorder slides, add notes or media). python-pptx and
  `zipfile.testzip()` are too lenient – both have accepted files PowerPoint
  rejected as corrupt (0x80070570). Kill any stale `POWERPNT` process first.
- **PowerPoint renumbers part filenames on save.** After a save, `slideN.xml`
  is **NOT** display-slide N (and `imageN` / `notesSlideN` shift too).
  Resolve display→part by parsing `ppt/presentation.xml` `<p:sldIdLst>`
  order → `r:id` → `presentation.xml.rels` (ElementTree, not a regex).
  Assuming `slideN.xml = display N` has corrupted a deck.
- **custGeom curve gotchas** (each makes a shape invisible or dropped): the
  path segment element is `<a:lnTo>`, not `<a:lineTo>`; inside `<a:ln>` the
  child order is fill → dash → join (put `<a:round/>` after `<a:solidFill>`);
  `<a:prstGeom>` sits after `<a:xfrm>` and before the fill.
- **Inside `<a:rPr>`, `<a:solidFill>` must come BEFORE `<a:latin>`** (schema
  order: ln → fill → effects → … → latin/ea/cs). A fill APPENDED after the
  font elements is silently ignored by PowerPoint — text renders in the
  default color with no error. When post-processing runs (e.g. coloring
  OMML runs), `insert(0, fill)`, never `SubElement`-append.
- **A PowerPoint save hides OMML shapes from python-pptx.** When I save the
  deck in PowerPoint, every textbox containing OMML math gets wrapped in
  `mc:AlternateContent` — python-pptx then no longer enumerates those shapes,
  so a hand-edit diff reports them as phantom deletions. Treat missing OMML
  boxes as save artifacts; confirm real hand-edits via COM slide renders or
  the raw slide XML, never via python-pptx enumeration alone.
- **Check for `~$<deck>.pptx` lock files before COM automation.** A lock file
  means I have that deck open in PowerPoint — do NOT kill `POWERPNT`
  processes then (the "kill stale POWERPNT" advice above applies only when no
  deck is open); open presentations read-only alongside instead.
- **A PollEverywhere slide's NOTES are part of the poll mechanism — splice
  them with the slide.** The PollEv add-in scans the whole deck at SLIDESHOW
  start, finds its `__PE_POLL_EMBED_ID` tag, and reads the poll data from
  that slide's notes ("Poll Title: Do not modify the notes…" + the poll
  URL). A poll slide whose tag is present but whose notes part is missing
  crashes the slideshow renderer DECK-WIDE — the error surfaces as "The
  slide failed to open properly" on slide 1, while the editing canvas,
  exports, and COM slideshow stepping all look fine. When splicing poll
  slides, always carry slide XML + tags part + image + the NOTES part
  together. Debug technique that found this: drive the real slideshow via
  COM, screenshot the `screenClass` window via PrintWindow, and bisect with
  `sldIdLst`-subset decks.
- **My hand-edits often arrive as scaled groups.** PowerPoint hand-edits come
  as `grpSp` with `off/ext` differing from `chOff/chExt`; decode each child
  to its RENDERED position (`off + (child − chOff) × ext/chExt`) and port
  into the build script as ungrouped shapes at those coordinates — grouping
  is redone wholesale in the phase-3 grouping/animation pass. Also diff
  speaker notes separately (a shape diff won't catch notes edits), and
  remember table-internal formatting (cell fills, per-run colors/italics) is
  invisible to a shape-level diff — do a render comparison for table slides.
- **My hand-edits can include ANIMATION re-choreography — diff `<p:timing>`
  too.** Once a deck has builds, "adopt all my changes" covers the click
  order, beat composition, and which shapes I demoted to static. Extract my
  timing per slide (parse mainSeq → per-click effect lists, resolving
  `spid`s to shape signatures of type + position + text, since ids differ
  across rebuilds), translate the beats into the `_animate.py` per-slide
  plan, and VERIFY by extracting the rebuilt deck's timing and diffing the
  signature sequences against mine — accept only a beat-for-beat match.
  Never regenerate animations from the generic guidelines when the source
  deck carries my hand-tuned choreography.
- **Close every hand-edit port with an AUTOMATED member-level geometry
  diff — never sign off from a visual read of the dump.** Transcribing
  coordinates from a dump by hand is how a resized group member slips
  through (a region rect inside a group kept its old height while the
  dump plainly showed my new one): the eye pattern-matches to the value
  already in the build script. The closing check is a script that decodes
  EVERY shape in both decks to rendered inches — including group children
  via the chOff/chExt transform — matches shapes by type + text (after
  normalizing PowerPoint's math-italic codepoints 𝑃→P to avoid phantom
  mismatches on OMML), and compares position AND size to ~0.01".
  "Adopted" means that diff prints clean, the same way the timing check
  must be beat-for-beat and the slideshow probe must PASS.
- **Company logos:** fetch from Wikimedia Commons via
  `Special:FilePath/<File>.svg?width=N` (the server rasterizes SVG → PNG);
  Wikipedia's REST page-summary API helps locate a company's logo file. Store
  under the module's `_source_images/_logo_*.png` and keep them flat (logo
  exception: no rounding, no shadow).
- **Photos from Wikimedia Commons:** search via the Commons API
  (`action=query&list=search&srnamespace=6`), download the top ~3 hits per
  query at low width, and VISUALLY REVIEW before choosing — searches
  misfire ("Coach store" returns trains, "Albertsons" returns Publix
  conversions). Re-download the chosen file at ≥1400 px into the module's
  `_source_images/web_*.jpg` (BUILD INPUTS — never delete). Caption
  "Photo: Wikimedia Commons" on-slide; record the exact `File:` titles in
  the speaker notes.
- **After I hand-edit the canonical .pptx, edit it in place** with
  python-pptx (swap an image blob, remove a paragraph) within one call, then
  copy back – don't rebuild from scripts. If reads intermittently fail
  (corrupt central directory), ask me to re-save and retry.
- **Bash-heredoc Python scripts lose one backslash level on this machine.**
  A `\\n` in Python source inside a Bash heredoc arrives as `\n` (and a
  lone `\` becomes a line-continuation error), so exact-string matches
  against the build script silently fail or the script won't parse. Any
  edit script whose match strings contain backslashes (`\\n` line breaks
  in box labels, regex escapes) must be WRITTEN TO A .py FILE with the
  Write tool and executed by path — never piped through a heredoc.
  Alternative for tiny cases: build the backslash with `chr(92)`.

### Phase order: rebuilding a module from its old deck (build.py FIRST)
The phases run in this order and it is **not reversible**:
1. **build.py scaffold** – generate ALL script-buildable slides (text,
   bullets, native charts / tables / equations, chrome) in the clean style;
   finish every such slide first.
2. **Freeze build.py** – once phase 3 begins the `.pptx` is the source of
   truth; re-running build.py regenerates from scratch and overwrites the
   phase-3 work.
3. **OOXML surgery + hand-edits** – splice in video / poll / interactive
   slides from the old deck (these can't go through python-pptx – it strips
   NULL video rels and poll `tags`), then fine-tune and port my hand-edits.
- **Warn me if I try to start phase 3 too early.** Importing a video / poll /
  live-content slide, or doing OOXML surgery, while build.py isn't complete
  for all slides freezes build.py prematurely – confirm the build.py pass is
  finished for every slide first.
- **PREFERRED (Module 7 pattern): make phase 3 a RERUNNABLE pipeline instead
  of freezing build.py.** Keep every phase-3 step as its own script run
  after each rebuild — `_build_ModuleX.py` → `_splice_media.py` (polls /
  videos from sidecar files) → `_animate.py all apply` — so build.py stays
  the source of truth forever and hand-edit porting keeps working through
  phase 3. Requirements that make this safe: splices start from fresh build
  output (nothing accumulates); hand-inserted live content is preserved as
  verbatim sidecar files (`_handoff_*.xml` / `.txt` / images — BUILD INPUTS,
  never delete); each pass is idempotent. Heavy embedded media (e.g. an
  11 MB video) goes behind an opt-in flag (`--with-video`) used only for
  the final teaching copy, so the working deck and its git commits stay
  small. Only fall back to the freeze-then-surgery flow when content truly
  can't be scripted or sidecar'd.

### File naming
- **Revised deck:** `Module X - Revised.pptx` (canonical, in-place-edited).
- **Per-module build script:** `_build_ModuleX.py` (the one-time scaffold;
  add a STALE banner once it's frozen).
- **Source images** live in that module's `Images/` subfolder.
- **Previous versions:** `Module X - Revised_t-1.pptx` and
  `Module X - Revised_t-2.pptx` (rolled before each edit, per the .pptx
  rules above).
- For ad-hoc / non-module decks, the drafting-workflow naming applies.

### Things to avoid on every slide
Walls of text (split / trim any bullet past two lines); orange / gold for
ordinary emphasis (use bold navy); gray sub-bullets (use dark navy);
decorative imagery that carries no information; chart junk (duplicate
legends, needless gridlines, axis titles that repeat the slide title);
redundant captions / labels / source lines; stock photos and clip art;
emojis anywhere on the slide; trailing periods on bullets; "Page X of N"
footers / watermarks / "Confidential" stamps; multiple slide-layout masters
in one deck.

## Standing Authority Within Module Subfolders
Claude has full authority to create, modify, move, and delete files
inside each module's subfolder (e.g. `405 Slide Revisions 2026\Module 2\`,
`Module 3\`, `Italy IBR\`) without asking — build scripts, decks, images,
notes, temp files. Outside a module subfolder, and for these regardless
of location, confirmation is still required:
- deleting or overwriting **source/input decks** (the original .pptx
  files a rebuild starts from);
- git commits and pushes (per the session-end ritual);
- any change to CLAUDE.md files or settings.
