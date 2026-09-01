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

## Economics Terminology Conventions
Fixed wording for concepts the courses return to. Once a convention is
listed here it holds in EVERY slide, speaker note, podcast script and
teleprompter script of the deck – if a change lands on one slide, sweep
the whole deck for the old wording rather than fixing the one instance.

- **Short run = "stop production"; long run = "shut down"** (2026-08-30,
  Nico). The two horizons are different decisions and must not share a
  verb:
  - In the **SHORT run** the firm can only **stop production** (equivalently
    "not produce" / "continue to produce"). Its fixed costs are already
    committed and are owed whether or not it produces, so stopping is a
    temporary halt with the option to restart.
  - In the **LONG run** the firm **shuts down** – it exits the industry
    altogether, and thereby also escapes the fixed costs.
  - So the short-run rule reads "**continue to produce** if P ≥ AVC,
    **stop production** if P < AVC", never "operate / shut down". Reserve
    "shut down" and "exit" for the long-run rule.
  - Do NOT use "operate" for either decision – it blurs exactly the
    distinction the two slides are drawing.

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
- **Rounding is for CARDS, not for grid cells** (2026-08-27, Nico). The
  rule above applies to a box that stands on its own — a scenario card, a
  comparison column header, a callout, a takeaway bar. It does NOT apply
  to the small cells of a hand-built table or number grid: those stay
  flat and square, because a rounded 2.5 × 0.4" cell in a row of like
  cells reads as a mistake rather than as a lifted card. The practical
  test is whether the box sits in a row / column of similarly sized
  siblings (grid cell → flat) or on its own (card → rounded + shade).
  Module 3 slides 49 and 54 are the reference: the 6.00 × 0.70" navy
  headers, the 6.00 × 2.65" cream panels and the 5.50 × 1.40" comparison
  cards are rounded and shaded, while the 2.50 × 0.40" cost chips inside
  slide 54's container and the whole of slides 48 and 56 stay flat.
- **Set the corner radius as a RENDERED length, not one shared `adj`.**
  A `roundRect`'s `adj` is a fraction of the shape's SHORT side, so the
  same `adj` gives a thin header bar and a tall panel visibly different
  corners. Pick the radius in inches (~0.08" is the deck's "slightly
  rounded") and compute `adj = radius / min(w, h) × 100000` per shape.
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
- **Slide titles are set in TITLE CASE**, the way a paper title is:
  every significant word starts with a capital letter, while articles
  ("a", "an", "the"), coordinating conjunctions ("and", "but", "or"),
  and short prepositions ("of", "in", "at", "to", "for", "with") stay
  lower case – unless they open the title, close it, or follow a colon.
  "Poll results: marginal revenue" is wrong; "Poll Results: Marginal
  Revenue" is right. This covers EVERY title, including the throwaway
  ones – poll slides, poll-result slides, stubs and backup slides.
  Hyphenated compounds capitalise both parts ("Cross-Price", "E-Book").
  - **Capitalise, never lower-case.** A pass that fixes titles only
    ever raises a letter; it must not lower-case a word that is
    already capitalised, so acronyms (MR, TR, OLS, WTP, A/B), product
    names, and deliberate capitalisation survive untouched.
  - **The same rule applies to the agenda / outline slides – to the
    ITEM TITLES only.** "3. Demand and Revenue", "3b. Marginal
    Revenue". The one-line description under an item is a sentence,
    not a heading, so it keeps ordinary sentence case.
  - Do this in the build script at the one place titles are drawn
    (Module 2: `_title_case()` wrapping `_draw_action_title`, and the
    item-title row of `make_m2_outline`), not by editing title strings
    one by one.
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
- **Sparse slides get MORE space between items, not more empty space at
  the bottom.** When a slide carries only a few items – roughly three or
  four main bullets, or three sentences on a poll-setup or question
  slide – raise the space-before to about **18 pt** so the lines are
  distributed over the content area instead of clustering at the top.
  Combine this with the larger end of the sizing range (28 pt main).
  The reference case is Module 2's In-Class slide 44 ("Obtain Price
  Elasticity from the Demand Function"): three lines at 28 pt with 18 pt
  before each, which balances the slide. Do this by default on sparse
  slides – it is the same principle as "even spacing between main
  bullets", applied when there is spare vertical room.

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
- **A DEMAND curve is dark red `C00000` – curve and label alike**
  (2026-08-30, Nico). This is the deck-wide default for every demand
  curve: the market demand in a supply-and-demand diagram, a firm's
  `d = MR` line, a labour-demand curve, the `D (MB of …)` line on an
  externality panel. Supply stays navy, so the two curves read apart at a
  glance and the price line and demand share one colour. Sweep the whole
  deck when it changes rather than fixing one panel – a lone dark-red
  demand curve among navy ones reads as an error. A slide may still make a
  deliberate exception (Module 4 slide 10 draws market demand in dark
  yellow because that slide is *about* telling the two curves apart), but
  it has to be a choice, not a leftover.
- **Consumer / producer surplus areas have fixed colours** (2026-08-30,
  Nico, adopted from MW's deck). Consumer surplus is a **red wash
  `C0201B` at ~26 %** with a `C0201B` outline; producer surplus is a
  **blue wash `4E79B5` at ~34 %** with a `4E79B5` outline. The explanatory
  card beside each panel is WHITE-filled with a border in the matching
  tone (`C0201B` for CS, `2E5AA8` for PS), and its text names the term in
  full with the abbreviation in parentheses – "Consumer surplus (CS) is
  the area …", "Producer surplus (PS) is the area …". The `CS` / `PS`
  label sits INSIDE its own triangle (place it at the triangle's
  centroid, computed – not eyeballed).
- **A surplus triangle's hypotenuse IS the curve it lies on.** Build the
  polygon from the curve's own intercept and the equilibrium point, so
  the sloped edge coincides with the demand (CS) or supply (PS) line
  exactly. And derive the equilibrium from the two lines rather than
  typing it in: Module 4 carried `(3.75, 5.5)` as "the exact crossing"
  for months when it was a point on the supply line only, which put the
  dot, `P*`, `Q*` and both triangles off the demand curve.
- **Profit / loss region shading on cost-curve panels** (2026-08-30, Nico).
  Wherever a diagram shades the rectangle between price and average cost,
  the fill says which case it is, at ~20 – 25 % opacity so the curves stay
  readable through it:
  | Case | Fill |
  |---|---|
  | Positive profit | dark red `C00000` |
  | Loss, but the firm keeps producing | gray `555B66` |
  | Loss so large the firm shuts down / stops producing | DARK gray |
  The dark-gray step is the visual cue that this is the shut-down case, so
  keep it clearly darker than the ordinary-loss gray.
  - **Write the word into the box.** Whenever the shaded rectangle is big
    enough to hold it, put **Profit** or **Loss** (with `π > 0` / `π < 0`
    where it fits) INSIDE the shaded region rather than captioning it
    underneath. Where the region is too thin or a curve cuts through it,
    the short form `π > 0` / `π < 0` alone is enough – the panel heading
    already names the case.

**Chrome** – keep identical across content slides; never enlarge one slide's
title relative to the others.
- **Navy top bar** with a section tag, white bold ~15–16 pt, title-cased.
- **The top-bar tag names the CURRENT AGENDA ITEM** (2026-08-26, Nico).
  A content slide's tag is `Module N · <agenda item title>`, copied
  verbatim from the module outline — slide 15 of Module 3 sits under
  outline item 2, so its tag reads `Module 3 · Short Run: Hiring
  Decisions`. A student can therefore always tell from the top bar where
  in the agenda s/he is. Rules:
  - **In a TAPED module the video number sits in the middle:
    `Module N · Video k · <topic>`** (2026-08-27, Nico). Once a module
    has been converted for taping, every content slide inside video k's
    block carries that video's number, so a student can see which video
    a slide came from as well as which topic it teaches – slide 6 of
    Module 3 reads `Module 3 · Video 1 · Course Roadmap`, slide 12
    reads `Module 3 · Video 2 · The Production Function`. **The block
    is delimited by the video title cards**: a slide belongs to video k
    if it sits between card k and card k+1. The `<topic>` level is
    whatever the slide already said, so the video number is SPLICED in
    rather than the topic retyped.
  - **Agenda / outline slides read `Module N · Video k · Agenda`** in a
    taped module (`Module N · Agenda` in an untaped one). The word
    stays "Agenda"; only the video level is added.
  - **Rename the item, rename the tag.** The outline list is the single
    source of both, so a build script derives the tag from it rather
    than repeating the wording per slide.
  - The module front matter (logistics, announcements, recap of the
    previous module, course roadmap, big picture) is INSIDE the
    introduction video's block once the module is taped, so it takes
    that video's number – `Module 3 · Video 1 · Recap`, `Module 3 ·
    Video 1 · Big Picture`. In an untaped module it keeps its own
    two-level tag.
  - **Two slides are deliberately exempt and keep a two-level tag:** the
    **concept map** (`Module N · Concept Map`) and the **summary
    closer** (`Module N · Summary`). They are reference slides that
    belong to no single video, and the concept map is exempt in every
    copy of it – the one in the video block and the one in the appendix
    alike.
  - Slides with no top bar at all stay as they are, and none of these
    ever gets a tag: the deck title slide, the video title cards, the
    section / appendix dividers, and the PollEverywhere slides.
  - **A backup slide DOES get a tag: `Module N · Backup`** (2026-08-27,
    Nico – this corrects an earlier version of this rule that put backup
    slides on the no-tag list). Everything after the BACKUP divider is
    ordinary chrome: navy top bar, `Module N · Backup`, action title,
    footer. Module 1's backup section (displays 91–95 of
    `Module 1 - Revised.pptx`) is the model, and the tag is the same on
    every backup slide – it does not name the topic the slide backs up.
    - **The one exception is a backup slide built as a FULL-BLEED
      figure.** Where the picture has been enlarged to fill the canvas
      and the text pulled on top of it – Module 3's two MPL photos,
      slides 91–92 of `Module 3 - Revised.pptx` – there is no top bar
      to put a tag in, and none is added. That is a deliberate choice
      about a particular slide (the figure needed the room), not the
      default. **Do not strip the top bar off an ordinary backup slide
      to reach it**, and do not add a bar to a full-bleed one.
    - The BACKUP divider itself carries no top bar, like every other
      divider.
  - This SUPERSEDES the older three-level `Module · Part · Section` tag
    (`Module 3 · Production · Short Run`), whose middle level drifted
    away from the agenda wording. (A non-module deck family may still
    use a two-level tag.)
  - **In-class examples of video material get a four-level tag**
    (2026-08-27, Nico). When a module is taught videos-first, some
    slides APPLY the taped material and are kept back to be shown on
    campus — the mini-cases, worked applications and polls I work
    through with the class after they have watched the videos. Those
    slides read `Module N · In Class · Examples · <topic>`: the module
    number, then `In Class`, then `Examples`, then the topic the
    examples cover (`Module 1 · In Class · Examples · Markets`,
    `Module 1 · In Class · Examples · Supply and Demand`). The student
    can then see at a glance that the slide is an application of
    something s/he has already watched, not a new agenda item — which
    is exactly why these slides are not one of the outline's items and
    sit under a divider of their own instead.
    - **The test is what the slide is FOR, not where it sits.** A slide
      gets this tag when it applies taped material and is planned for
      class. Where such slides are parked is a per-deck layout choice
      and changes nothing about the tag. Two arrangements exist so far:
      an **"Applications" divider mid-deck**, with the block running to
      the next agenda slide (Module 1, displays 51–65), and an
      **examples APPENDIX at the end** of a taped deck (Module 3). A
      deck uses whichever fits; **a module with no appendix simply has
      no appendix**, and the mid-deck arrangement carries the whole
      rule on its own. Module 1 is that case — it has no "slides not
      used in the videos" section, and nothing in the appendix
      sub-rules below applies to it.
    - **`Examples` never appears inside a video block.** A slide sitting
      between two video title cards is part of that video and carries
      `Module N · Video k · <topic>`, even when its content is a worked
      example. The four-level tag is for the copies kept back for class.
    - The `<topic>` is the wording of the VIDEO topic being applied,
      not a new heading invented for the slide.
    - The tag covers the whole applications block, from the
      applications divider to the next agenda slide. Never let it stop
      in the middle of a two-slide mini-case (set-up tagged one way,
      resolution another).
    - Slides in the in-class part that teach a topic of their OWN (the
      agenda items taught only in class) keep the ordinary agenda-item
      tag; `Examples` is only for applications of taped material.
    - Route it through the deck's tag constants, one per topic, never
      per call site. Module 1's `TAG_MARKETS` / `TAG_SD` in
      `_build_Module1.py` are the reference.
    - A slide that was copied out of a video block into the in-class
      part keeps its CONTENT but loses the video's tag: the in-class
      copy of Module 1's course roadmap is tagged `Module 1 ·
      Introduction`, while the copy inside Video 1 keeps `Module 1 ·
      Course Roadmap`. A `Video k` tag never appears outside that
      video's own block.
    - **The same tag applies to an examples APPENDIX at the end of a
      taped deck** (2026-08-27, Nico) — *when the deck has one; this
      sub-rule and the two after it are about that arrangement only.* A
      video deck may close with an
      appendix of slides that were cut from the videos – Module 3's
      "SLIDES NOT USED IN THE VIDEOS" – and the example slides in it are
      shown on campus. They get the same four-level tag, so the block
      reads `Module 3 · In Class · Examples · Wage Searchers`, `… · Cost
      Concepts`, and so on. The block runs from the first example slide
      to the END of the deck.
    - **The appendix's own front matter keeps its tag.** The appendix
      divider has no top bar, and a reference slide carried into the
      appendix (Module 3's concept-map copy) keeps its own two-level
      tag. Only the example slides are retagged – in Module 3 that is
      slides 95–110, not 93–94.
    - **Splice the level in, never retype the topic.** An appendix slide
      is a copy of a video slide, so it already carries the video
      topic's wording. Insert `In Class · Examples ·` after `Module N ·`
      and leave the rest of the string untouched. That satisfies the
      "topic = the video topic's wording" rule automatically and
      survives a later renaming of the outline item.
    - **Measure the result.** Four levels is a long tag, so check the
      label in the real font (PIL ImageFont, Calibri Bold at the tag's
      own size) against the tag box width and fail loudly rather than
      shipping a wrapped or clipped tag. Module 3's longest, `Module 3 ·
      In Class · Examples · Long Run: The Optimal Input Mix`, measures
      4.4" in a 12" box. Reference pass: `_retag_inclass.py` in
      `Module 3/`.
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
  - **Exception – a label written INSIDE a narrow object** (2026-08-27,
    Nico). Where a label has to sit inside a narrow figure element, a
    narrow table cell, or a similarly constrained object whose width is
    fixed by the data rather than by layout – the caption inside a thin
    bar of a bar chart, a value written into a slim column, a label on
    a narrow timeline segment – the floor does not apply. Fit the text
    to the object at whatever size it takes, and keep it as large as
    the object allows. The alternative (a shorter, vaguer label, or a
    wider bar that misrepresents the data) is worse than small type.
    Module 1's exercise chart is the reference: "Net Benefit of Hour 1"
    is set at 9 pt inside a 0.95"-wide bar and stays that way.
  - The exception is narrow, in both senses. It covers labels *inside*
    a constrained object only. Body bullets, callouts, teaching cards,
    quote boxes and free-standing chart labels keep their floors – a
    crowded slide is not a narrow object. Don't reach for the exception
    before trying to shorten the text or grow the box.

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
  (`<a:tailEnd type="triangle"/>`). Axis titles italic-bold navy, tick
  labels regular navy.
- **Axis titles are anchored to the ARROW TIP, in a box sized tight to the
  label** (2026-08-30, Nico – geometry measured off the placement he set
  by hand on Module 4 slide 36, after two earlier versions of this rule
  guessed wrong). Not floating above the plot, not trailing off the end of
  it, and never in a generous fixed-width box with an alignment: a wide box
  puts the glyphs somewhere the arithmetic cannot see, so the title reads
  as detached from its arrow even when the anchor is right.
  - **Measure the label's rendered width in the real font** (PIL
    ImageFont, Calibri Bold Italic at the title's own size) and make the
    text box exactly that wide. At 17 pt, "P" measures 0.135" and "Q"
    0.167" – the box is that narrow.
  - **Y-axis title:** its vertical MIDDLE sits exactly at the arrow tip,
    and its right border **0.08" clear of the axis line** – sitting ON
    the axis reads as too tight. So `left = axis_x − width − 0.08` and
    `top = tip_y − height/2`. Route the clearance through ONE constant
    (`Y_TITLE_GAP`), never per call site.
  - **X-axis title:** its horizontal MIDPOINT sits exactly at the
    x-arrow's tip, and its TOP **0.05"** below the axis. So
    `left = tip_x − width/2` and `top = axis_y + 0.05`.
  - The tip is **0.18" beyond the plot bound** (the arrowhead's length),
    so it is computed from the figure, never eyeballed.
  - **A price axis is labelled `P`**, not `$/Q`.
  - This is the DECK-WIDE default, not an opt-in per chart.
- **EVERY label in a figure sits in a box sized to the label** (2026-08-30,
  Nico). Not just the axis titles – curve labels, tick labels, region
  letters, annotation text. Measure the string in the real font (PIL
  ImageFont on Calibri at the label's own size, plus ~0.08" of slack so a
  wide glyph cannot wrap) and make the box that wide, keeping whatever
  edge the label is anchored by: a right-aligned y tick keeps its right
  edge on the axis, a centred x tick keeps its midpoint on the tick, a
  left-aligned curve label keeps its left edge. The box then re-fits
  itself when the wording changes, and a stray wide box never sits
  invisibly across the plot catching the cursor. Route this through the
  label helpers, never per call site.
- **A legend mark for an AREA copies that area's shape – and a
  combination of areas copies their arrangement** (2026-08-30, Nico).
  Wherever a diagram shades lettered regions (A, B, C …) and a list
  beside it refers to them, each line carries a small mark in the
  region's colour AND shape: a square for a rectangular region, a right
  triangle for a triangular one, flipped when the region's right angle is
  on top. At ~0.20" and the same wash as the region itself.
  - **When a line names several regions, lay their marks out the way the
    regions sit in the graph.** Side by side (~0.02" apart) when the
    areas sit side by side; **stacked** (~0.01" apart) when one sits above
    the other. "Firms lose A + B" gets the two marks in a row because A
    and B are neighbours; "Deadweight loss: B + C" gets them one above the
    other because B sits on top of C. The reader can then match the legend
    to the figure without reading the letters.
  - The line's text is vertically **centred on the mark block**, so a
    stacked pair pushes its text down, and the next row clears the taller
    block rather than keeping a fixed pitch.
  - Reference implementation: `_welfare_rows` in `Module 4/_m4_helpers.py`,
    with rows given as `(layout, marks, text)` where layout is `"h"` or
    `"v"`. Module 4 slides 64, 68 and 72 are the worked examples.
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
- **No institutional branding mark on ANY slide** (2026-08-29, Nico). The
  UCLA Anderson wordmark – or any other school, university, or department
  logo – does not belong on the title slide, in the chrome, or anywhere
  else in a deck. The affiliation is already carried by the title slide's
  "Prof. Nico Voigtländer · UCLA Anderson" byline and by the footer, so the
  mark adds nothing; and it is set in UCLA blue, which fights the deck's
  navy / gold palette. When a rebuild inherits one from an original deck,
  DROP it rather than reproducing it in the build script – this is how
  Module 4 kept one at 0.60 / 6.42" on slide 1 through the whole rebuild.
  - **This does NOT cover company logos used as CONTENT.** A brand mark
    that illustrates the point of the slide stays – Boeing and Airbus
    standing for an oligopoly, LADWP for a regulated monopoly (Module 4
    slide 7), a firm's logo in a mini-case. Those keep the flat treatment
    above. The test is whether the logo is the SUBJECT of the slide or the
    letterhead on it; only letterhead is banned.
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
- **The same rule covers charts and tables, not just photographs.** A
  native table or chart, its white backing card, and its source line are
  ONE object. Group all three. The one deliberate exception is a figure
  built in pieces so the pieces can be revealed on separate clicks (a
  table split into sections, for example): there the backing card and
  the source line are grouped with the FIRST piece and the remaining
  pieces stay separate — and the shared card must keep the z-order of
  that first piece, or it will be drawn on top of the pieces below it.

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
  - **A poll runs over three slides, and the badge alternates** (2026-08-27,
    Nico). Slide 1 is the set-up / question slide and carries the **Poll
    Break parallelogram**. Slide 2 is the PollEverywhere slide itself – the
    live poll or its screenshot – and carries the **round POLL pill**
    instead: a fully-rounded rectangle (`roundRect`, adj 50000) with gold
    `E09F3E` fill, no border, soft drop shadow, and navy `0B2B4E` bold
    Calibri **"POLL"** centered, about 1.49 × 0.51". Slide 3 is the
    solution slide and carries the **Poll Break parallelogram** again.
    Reference: slides 9–11 of `Module 3 - Video 3 - Short Run Hiring.pptx`
    (set-up → PollEv screenshot → "Solution: MRPL of Rivian").
  - **Both marks sit in the bottom-RIGHT corner, overlaying the footer**, and
    both are grouped box + text. Neither is ever animated (see Animations).
  - **The poll marks OWN their corner; other corner boxes give way**
    (2026-08-30, Nico). The Poll Break parallelogram and the round POLL
    pill always take their designated bottom-right position – they never
    move to make room. Any OTHER corner-dwelling box that would otherwise
    occupy that spot (the Yi-family example tab, a problem-set pointer, a
    practice-video link) is the one that moves: **up, to sit directly
    above the badge, or left, to sit beside it – whichever direction the
    slide has more free space in.**
    - Moving UP, keep the box's own right-hand alignment so the two read
      as one column. On the 13.33 × 7.5" canvas that puts the displaced
      box at **top 6.10"**, directly above the badge at 6.77" (Module 4
      slides 30 – 34 are the reference).
    - Moving LEFT, keep it on the badge's own baseline so the two read as
      one row.
  - **Both are drawn LAST, so they sit in the FOREGROUND** (2026-08-30,
    Nico). The mark straddles the footer rule by design, so anything
    emitted after it – the rule, the page number, a chart element –
    cuts a line straight across it. Append the badge or pill at the end
    of the `spTree`, after the footer, on EVERY slide that carries one.
  - **The Poll Break parallelogram has ONE fixed geometry, and this rule
    is the whole spec** (2026-08-27, Nico — rewritten the same day so a
    build script can generate the badge from these numbers alone, with no
    reference deck or hand-tuned sidecar file to copy from). On the
    13.33 × 7.5" canvas:

    | | |
    |---|---|
    | Badge position (`off`) | left **10.2377"**, top **6.7695"** (EMU **9361444 / 6190030**) |
    | Badge size (`ext`) | **2.9500 × 0.5326"** (EMU **2697480 × 487009**) |
    | Slant offset `S` | **0.72"** horizontally, on each slanted side |
    | Label box | left = badge left **+ S**, same top, **1.51 × 0.5326"** (EMU x **10019812**, cx **1380744**) — it fills the parallel middle, `W − 2S` |
    | Label | "Poll Break", **Calibri Bold 28 pt**, navy `0B2B4E`, centred, `wrap="none"` + `spAutoFit` |
    | Fill / border | gold `E09F3E`, **no** border |
    | Shadow | `outerShdw` blurRad **50800**, dist **38100**, dir **2700000**, black at **50 %** alpha |
    | Grouping | one `grpSp` (parallelogram behind, label in front), `off`/`ext` as above with `chOff`/`chExt` **equal** to them |

    - **It must be in the FOREGROUND, covering the footer rule.** The badge
      spans y 6.7695" – 7.302" and the footer's thin horizontal rule sits at
      y **7.15"**, so the badge STRADDLES it by design — tucked into the very
      corner rather than floating above the rule. **Draw the badge LAST**
      (append it at the end of the `spTree`, after the footer, the rule and
      the page number) so it renders in front of them. Drawn earlier, the
      footer rule cuts a line straight across the badge.
    - **The parallelogram is a `custGeom`, not a preset**, authored in the
      normalised 100000 × 100000 path box. With `s` = the slant as a
      fraction of the width (0.72 / 2.95 → **24406**), `r` = **5000** for the
      rounded corners and `d` = `r·s/100000` = **1220**, the path is:
      `moveTo(s+r, 0)` → `lnTo(100000−r, 0)` →
      `cubicBezTo (100000,0)(100000,0)(100000−d, r)` →
      `lnTo(100000−s+d, 100000−r)` →
      `cubicBezTo (100000−s,100000)(100000−s,100000)(100000−s−r, 100000)` →
      `lnTo(r, 100000)` → `cubicBezTo (0,100000)(0,100000)(d, 100000−r)` →
      `lnTo(s−d, r)` → `cubicBezTo (s,0)(s,0)(s+r, 0)` → `close`; with
      `<a:rect l="24406" t="0" r="75594" b="100000"/>`. Because the path is
      normalised, the height can be set straight to 0.5326" — the shape
      squashes and the 0.72" slant survives. The label stays 28 pt.
    - Route the position and size through ONE constant in the build script
      (`POLLBREAK_XY` / `POLLBREAK_WH`), never per call site. Reference
      implementation: **`_add_pollbreak_badge` in Module 1's
      `_build_Module1.py`**, which builds the badge from exactly these
      numbers.
    - **Decks retrofitted to this geometry:** all 22 badges in
      `Module 3 - Revised.pptx` and all three in `Module 1 - Revised.pptx`
      (2026-08-27). Only Module 3's seven per-video decks still carry a mix
      of the older positions (y 6.25 at full 0.72" height, and a
      2.585 × 0.512" variant); they are finished and are not retrofitted
      unless asked.
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
- **Post-work reference box** – the deck-standard pointer to a problem set
  or a teaching note: gold-bordered rounded rect (~25% corner), white fill,
  soft drop shadow, navy bold ~15 pt, with a **leading glyph** that says
  what kind of reference it is – **✎ for a problem set**, **▤ for a
  teaching note**. **Default position: the bottom-RIGHT corner,
  overlapping the footer** (left 10.17", top 6.53", 3.00 × 0.50" on the
  13.33 × 7.5" canvas), the same corner the practice-video box uses.
  Put every problem-set pointer there unless that corner is already
  taken on the slide (by a Poll Break badge or a video box); only then
  move it left along the bottom edge. Route the position through one
  constant in the build script (`PS_BOX_XY`), never per call site, so
  the corner cannot drift. Two rules for the label:
  - **Name the problem set NUMBER only – never the exercise numbers.**
    "Problem Set 2", not "Problem Set 2 · #4, #5". The exercises get
    renumbered from year to year and the slide should survive that.
  - Keep the label to the reference itself. The "on BruinLearn under
    …" line belongs on the wrap-up / post-work slide, not on every
    pointer.
  - **The glyphs are a fixed vocabulary – reuse them everywhere.** ✎
    always means "a problem set", ▤ always means "a teaching note", in
    every module and both the in-class and the video deck. Students
    should learn to recognise the mark, so never invent a second symbol
    for the same kind of reference and never leave one pointer bare
    while its neighbours carry a glyph. Build scripts route every such
    box through ONE helper (Module 2: `_add_reference_box(…,
    kind="ps" | "tn")`) rather than hand-styling call sites; a new kind
    of reference means adding a glyph to that helper.
- **Concept maps / outline anchors** – a network-graph overview slide at the
  start of each major section, returned to at transitions; the section
  divider highlights the current section (cream band, navy / gold badge,
  others dimmed).

## Module-Outline / Agenda Slides (numbered-circle format)
The standard for a module's outline slides — the descriptive overview near
the front, the section agendas at each transition, the wrap-up / post-work
slides, and the summary closer. Finalized on Module 2 (2026-08-16, adopted
from CT's format + Nico's band and spacing rules); apply IDENTICALLY in
other modules. Reference implementation: `make_m1_outline` in
`Module 1/_build_Module1.py` — copy it and swap the item list. (It
supersedes `make_m2_outline` in `Module 2/_build_Module2InClass.py`,
which predates the dimming rule below; Module 2's outline slides still
need that pass.)

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
  shadow. One band per highlighted item.
- **Dimming** (2026-08-24, Nico): on a section agenda the items that are
  NOT currently covered are shaded — the circle digit and the item title
  both `#BFBFBF` (white, 25 % darker; PowerPoint writes it as `schemeClr
  bg1` with `lumMod 75%`). The gold circle fill stays gold on every item.
  The descriptive overview and the summary closer, which have no current
  topic, keep every item navy.
- **Coverage pills – every outline item says where it is taught**
  (2026-08-25, Nico). A small rounded pill sits at the right of each
  row: **navy fill / white text** for a topic done **in class**, **gold
  fill / navy text** for one on **video**. This is now standard on
  EVERY module's outline and agenda slides – including a module taught
  entirely in class and one that is entirely on video, so the format
  stays uniform and a student can tell at a glance whether there is
  anything to watch beforehand.
  - **Geometry:** height 0.36", top = row_y + 0.02", 13 pt bold, corner
    0.30, soft drop shadow. A main item's pill is 1.55" wide; a
    **sub-item's is narrower (1.14") and right-aligned to the SAME right
    edge** (12.85"), so it reads as part of its parent's rather than a
    peer. Take the width from the item's own sub-item flag, never from
    a hand-kept list.
  - **Name the video a student actually has to watch, not a sequence
    number.** If a topic spans two videos the pill says so ("Videos
    1+2") and its sub-topics each name their single video. Numbering
    topics 1, 2, 3, 4 when there are only three videos points students
    at the wrong one, and the video deck's own agenda then contradicts
    itself on screen.
  - **Dimming:** on a section agenda the pill dims with its row –
    `#BFBFBF` fill, white text, and no shadow – so only the current
    topic keeps its colour. The descriptive overview and the summary
    closer, which light every item, keep every pill in colour.
  - **The pill sits vertically CENTRED in its row – 0.16" below the top
    of the reserved two-row box** (2026-08-26, Nico). A section agenda
    shades most of its rows down to a single line, which is itself nudged
    down to centre it in that box; left at the top, the pill floats away
    from the item it belongs to. Use the same drop on EVERY agenda slide
    of the deck – section agendas, the descriptive overview and the
    summary closer alike – so the pill column is identical everywhere and
    nothing jumps between consecutive slides. Route it through one
    constant in the outline builder (`PILL_DROP`), never per row.
  - **A description must clear the pill by at least a five-letter word**
    (2026-08-26, Nico). The description row runs the full width of the
    text box, so a long one reaches under the pill and reads as running
    into it. Measure each description in the real font (PIL ImageFont on
    Calibri at the description's own size) and require
    `width ≤ pill_left − text_left − width("costs")`; on the standard
    geometry (text box at x 2.05", pill left edge 11.30", 22 pt) that is
    8.63". If a description is over, **shorten the wording** – never
    shrink the type or narrow the text box. Module 3's "Cost Concepts"
    line lost its closing "for decisions" this way. Put the check in the
    build script so it fails loudly rather than shipping a collision.
  - **Implementation:** one table mapping item → pill text, read by the
    outline builder (Module 2: `COVERAGE_LABEL` in the shared layer).
  - **Watch the right edge.** The pills occupy it, so a pointer / link
    box on an agenda slide has to sit at the convention position
    (y 6.68, overlaying the footer) – higher up it collides with the
    last item's pill.
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

## Converting a Module for Video Taping
When I say "convert / adapt Module X for my video purposes" (or for
taping), apply these two steps to **every section that will be taped**.
A module can be taped in full (Module 3: all six sections are videos) or
only in part (Modules 1 and 2: half video, half in class) — so first
establish which sections are taped, then treat only those.

1. **A video title card at the start of each taped section**, placed
   immediately BEFORE that section's agenda slide, matching the order the
   video decks use (title card → outline → content).
   - **The INTRODUCTION video's card is always the VERY FIRST slide of
     the deck** (2026-08-26, Nico) — ahead of the deck title slide, so a
     video-mode deck opens by naming the video the viewer is about to
     watch. The introduction is item 1 of the outline (Module 3:
     "Introduction to Module 3 — a brief overview of what we cover in
     Module 3"), and it has no section agenda of its own: the module
     overview slide is its agenda.
   The card is the
   title-slide layout with no top bar, no footer text and no page number:
   the section name navy bold 60 pt at y 2.10"; `Module N  ·  Video k`
   gold bold 40 pt at y 3.25"; a 4" gold strip centered at y 4.28";
   "Management 405" gray bold 26 pt at 4.62"; "Prof. Nico Voigtländer ·
   UCLA Anderson" gray 22 pt at 5.32"; the footer rule and its gold strip.
   Reference implementation: `_video_title_slide` in
   `Module 2/_build_Module2Video.py`. The card carries **no speaker
   notes**. The card's name is the outline item's title, so the agenda
   and the video announce the same thing.
2. **Coverage pills on every agenda slide** (the "video banners"), per the
   Coverage-pills rule in the Module-Outline section: gold fill / navy
   text for a taped topic, navy fill / white text for one done in class,
   dimming with its row on a section agenda. Numbering follows the videos
   a student actually has to watch, not the item numbers.

Order matters: insert the title cards FIRST, then rebuild the agenda
slides, so the pills and the cached page numbers land on the final slide
order. Module 3's pipeline is the reference — `_video_prep.py` then
`_retrofit_agenda.py`.

Everything else in the deck stays as it is: the top-bar tags keep naming
the agenda item (they are not rewritten to "Video k"), and the module
front matter (logistics, announcements, recap, roadmap, big picture,
concept map) sits ahead of the first title card, outside any video.

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
    - **The preview gives INTUITION, never the technical apparatus**
      (2026-08-31, Nico – the Module 4 preview came out "very detailed and
      technical"). "Intuitive grasp of each core concept" means the plain
      idea a listener could repeat to a colleague over coffee – NOT the
      rule, the test, or the formula that the class will spend an hour
      building. So name the QUESTION the module answers, not the machinery
      that answers it:
      - "how a firm decides how much to make when it has no say over the
        price it gets" – **not** "produce where marginal cost equals price";
      - "when a business that is losing money should keep going anyway, and
        when it should stop" – **not** "compare price to average variable
        cost";
      - "why unusually good profits in a competitive business never last" –
        **not** "entry and exit drive economic profit to zero";
      - "how to see who really gains and who loses when a government steps
        in, and how much value simply vanishes" – **not** "consumer surplus,
        producer surplus and deadweight loss".
      A term may be NAMED as a label when the class will use it constantly
      (*price taker*), but it is then dissolved into everyday words in the
      same breath, and no second technical term is stacked on top of it.
      **Test before sending:** if a sentence in the body contains two or
      more pieces of course jargon, or any "X equals Y" rule, rewrite it.
      No equations, no curve names, no acronyms.
      - This constraint is on the PREVIEW only. The **wrap-up** is a recall
        episode and SHOULD use the technical vocabulary, because the
        listener has already sat through the class.
    Open with a line like "Here's a preview of what to expect from
    Module X." Two hosts: **one who did the reading and is prepared**,
    briefing **one who did not and just wants a quick sense of what's
    coming.**
  - **Wrap-up / recap – 15 to 20 minutes.** Lets students *recall* what they
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
  in chat **and** as a labeled block at the very top of the source doc, so I
  can drop it straight into NotebookLM as a redundancy on top of the in-doc
  instructions.
  - **Print BOTH prompts in the chat window every time, unprompted**
    (2026-08-28, Nico). Whenever I ask for a module's podcast scripts – a
    first draft, a revision, or a fix to one episode – end the turn by
    printing the intro prompt and the wrap-up prompt in the chat as two
    separate copyable blocks, each labelled with its episode. I paste them
    into NotebookLM's Customize box from the chat, not out of the file, so
    "it is in the doc" is not enough. Print both even when only one episode
    changed, so I never have to hunt for the other one, and print the CURRENT
    text – re-read it from the docs rather than reproducing an earlier
    version from the conversation. A module with a Wrap-Up **Video** gets its
    Video Overview prompt the same way, as a third block.
  - **Write it as an ANSWER to the question the Customize box actually
    asks** (2026-08-27, Nico): *"What should the AI hosts focus on in this
    episode?"* So it opens "Focus on …" and stays in that register
    throughout – instructions TO the hosts, not a description of the
    episode. Label the block in the doc with that question, verbatim, so
    the destination is unambiguous.
  - **No Markdown inside the prompt.** Asterisks and underscores are read
    as literal characters in that box; write the emphasis out in words.
  - The workflow the prompt belongs to: **Add sources → upload the
    episode's `.md`**, then **Studio → Audio Overview → Customize → paste
    the prompt → Generate**. One episode per notebook.
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
  - **Diminishing marginal product: smaller and smaller, and possibly
    negative** (2026-08-28, Nico — an earlier version of this rule produced
    a factual error in the Module 3 intro episode). As we add more and more
    workers to the same fixed capital, the **marginal** return to each extra
    worker gets smaller and smaller, and eventually it **can even turn
    negative** — a factory floor so crowded that people get in each other's
    way, or too many cooks in one kitchen. Say that in a sentence; it does
    not need detail. Then be clear that **for the purposes of this class we
    never end up in that range**: we work where the marginal product is
    still **positive**, so each extra worker does add value, and the
    question is how many extra workers it makes sense to hire at the going
    wage.
    - **Do NOT have a host deny the negative range.** The previous rule
      said diminishing returns "does NOT mean output falls", and the hosts
      hardened it into "So it never means that production slows down or
      falls?" / "Absolutely not." That is wrong: marginal returns to labor
      really can go negative, and Module 3's backup slides make exactly
      that point with the cramped factory floor. What is true is that
      economists usually work with a total-output curve that flattens
      toward zero marginal product, and that the class stays in the
      positive range — not that the negative range does not exist.
  - **In general, state each result no more strongly than the economics
    supports.** The hosts ad-lib and tend to exaggerate in whichever
    direction the instructions lean, so a rule phrased as a flat denial
    will come back as a flat denial.
  - **Keep the language measured – go easy on "massive"** (2026-08-27,
    Nico). The hosts reach for extreme words – "massive", "huge",
    "enormous", "incredible", "game-changing" – every other sentence, and
    it wears out fast. The words are not banned; ask for them
    **occasionally**, a handful of times across an episode, where the size
    of something is the actual point. Otherwise say it plainly and let the
    fact carry the emphasis ("the largest supermarket merger ever
    proposed" needs no "massive" in front of it). Put this in BOTH the
    in-doc host instructions and the paste-ready Audio Overview prompt.
  - **Keep returning to the module's one unifying idea** (the throughline).
  - **Follow the MAIN part of the deck, and ignore any appendix**
    (2026-08-28, Nico). A finished module deck can end with a block of
    slides held back for in-class use – Module 3's "SLIDES NOT USED IN THE
    VIDEOS – FOR In-Class APPLICATIONS", slides 95 onward. Take the
    episode's structure from the main part only; those extra examples are
    for the classroom, not the podcast.
  - **Lead with real-world stories** and let them carry the ideas.
  - Warm, curious, conversational tone – smart colleagues (or two students)
    connecting the dots, not a lecture; define terms in plain language, go
    light on formulas.
- **Usage / mechanics:** put **each episode in its own NotebookLM notebook**
  (NotebookLM blends all sources in a notebook into one audio), so the two
  files must never share a notebook. There is **no editor for the finished
  audio** – to change an episode, edit the source doc (especially the
  instruction block) and regenerate. No invented facts or numbers, and
  spot-check the generated audio since the hosts ad-lib.
- **Length: the source doc controls it, not the prompt** (2026-08-27, Nico).
  NotebookLM expands whatever it is given, so an "about 5 minutes" line on a
  full-length doc is ignored -- Module 1's intro was written that way and
  came out at **18 minutes**. Three levers, in order of strength:
  1. **The panel's length control.** If the Audio Overview panel offers
     Shorter / Default / Longer, set it. It beats any wording.
  2. **The length of the source doc.** For a 5-minute episode the body
     wants to be ~400 words -- one sentence per idea, no example worked
     through, no connective narrative. For a 15-20 minute episode it can
     run to ~2,400. Write the doc to the target; do not write a long doc
     and ask for a short reading of it.
  3. **A hard, quantified instruction, stated more than once.** "No longer
     than 5 minutes -- roughly 700 spoken words in total", plus a structural
     cap ("about a dozen short exchanges", "name each idea, give it one
     plain sentence, and move on") and a closing "Above all, stay under 5
     minutes." Put it at the START of the prompt as well as the end.
  Current targets: **intro 5 minutes MAXIMUM**, **wrap-up 15-20 minutes**
  (20 is the ceiling). Say "maximum"/"ceiling" rather than "about" wherever
  the length actually matters -- "about" reads as advisory.

## Wrap-Up Video (NotebookLM Video Overview)
A module can also get a **video** wrap-up, produced with **NotebookLM's
Video Overview** from a source doc I write (I don't produce the video). It's
the video sibling of the Wrap-up podcast: a 15- to 20-minute, **past-tense** recap
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
  runaway *production* cost, not weak box office; diminishing marginal
  product means smaller and smaller additions that can eventually turn
  negative, with the class staying in the positive range; keep returning
  to MB = MC).
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
- **Poll chrome is never animated.** The gold "Poll Break" parallelogram and
  the gold round POLL pill are visible from the moment the slide appears – no
  click, no fade. Students should see that a poll is coming while I talk
  through the slide's build, rather than be surprised by it at the end. Treat
  both marks as chrome: leave them out of the per-slide animation plan
  entirely (in `_animate.py`, drop any `t:Poll Break` / `t:POLL` selector –
  anything not selected stays visible from the start). The same applies to a
  "Discussion Break" / "Group Discussion" badge. Because the badge sits on
  top of the footer rule, it is also drawn LAST – see its fixed position
  under Layout patterns.
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
- **Profit is a LOWER-CASE pi** (2026-08-30, Nico). In every equation write profit as π (and a change in profit as Δπ), never the capital Π. Set it as an OMML variable run so it comes through in Cambria Math italic like any other symbol. Where a line needs to name the quantity as well, label it "Profit:" in upright text and then give the formula in π.
- **Elasticity symbol (econ decks):** render with the **D as a subscript –
  Eᴅ** – everywhere it appears (titles, bullets, equations), consistently.
- **Worked-solution slides** mirror the source's "1. / 2. / 3." numbered
  structure and wording; color-code matched quantities (each quantity its
  own color, reused in equation + bullets), but only the quantities carrying
  the pedagogical link.
- **The final solution is set in DARK RED** (2026-08-26, Nico). On every
  solution slide the line that delivers the answer - the last equation of
  the derivation - is dark red (`C00000`), while the setup and the
  intermediate steps stay navy. This is the baseline for all decks, not
  just game-theory ones (the payoff-matrix section below repeats it for
  numeric answers). One red line per solution slide: if every equation is
  red, none of them reads as the answer.

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
- **NEVER verify a pass by importing it.** (2026-08-27 – this destroyed
  work.) These passes do their job at MODULE level; many have no `main()`,
  so `import _retrofit_agenda` runs the entire pipeline and rewrites the
  canonical deck on the spot. Checking "does this script still work?" that
  way silently regenerated seven agenda slides and overwrote every top-bar
  tag, including a hand edit made minutes earlier. Verify **statically**
  instead: `ast.parse` for structure, `py_compile.compile(f, doraise=True)`
  for syntax, and an AST walk to confirm the symbols a pass calls exist in
  the helper module it imports. Only ever RUN a pass deliberately, by path,
  with `--dry-run` first.
- **Two guards belong on every pass that rewrites a deck**, and it is worth
  adding them the moment a script grows past a one-off:
  - an **import guard** – `if __name__ != "__main__": raise ImportError(...)`
    for a pass whose work sits at module level (or put the work in a
    `main()` behind `if __name__ == "__main__":`);
  - a **pre-flight git check** – refuse to run when the deck has
    uncommitted changes, overridable with `--force`. The rolling `_t-1` /
    `_t-2` backups are in `.gitignore`, so they never protected against a
    bad script, and a finished module has had them deleted: **git is the
    only way back.** Reference: `_deck_guard.require_committed(DECK)` in
    `Module 3/`, called by all six of that folder's passes (skipped on
    `--dry-run`, which writes nothing).
- **Do NOT ask to commit before a routine rebuild** (2026-08-29, Nico –
  this replaces an earlier "commit the deck BEFORE running anything
  destructive over it – a pass, a splice, a rebuild", which turned every
  rebuild into a commit prompt). Running the build script is
  **reproducible**: the script is the source of truth, so a rebuild
  destroys nothing and needs no commit. The safety net for ordinary work is
  the start-of-session git snapshot plus the rolling `_t-1` / `_t-2`
  backups. Commits and pushes happen when I ask for them – the universal
  CLAUDE.md governs.
  - **The one exception is UNPORTED HAND-EDITS.** When the deck carries
    changes I made in PowerPoint that are not yet in the build script, a
    rebuild would destroy work that no script can regenerate. Then port the
    hand-edits first, roll the `_t-1` / `_t-2` backups, and **offer** a
    commit in one line – never block on it, and never stop work waiting for
    an answer.
  - The `--force`-able pre-flight git check above stays as it is. It guards
    against a script running by accident; it is not a reason to raise a
    commit prompt in conversation.
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
- **Splice live content from a SIDECAR deck, not from the full source deck**
  (2026-08-27, Nico). `_splice_media.py` needs the original only for a
  handful of slides, but pointing it at the source decks keeps tens of
  megabytes in the folder forever (Module 1: 65 MB of source decks serving
  6 poll slides). Carve the needed slides into a small sidecar instead:
  - **Build the sidecar with PowerPoint via COM**, not with python-pptx and
    not by hand: copy the source deck, delete every slide except the ones
    the splice map names, and Save. PowerPoint does all the rel / tags /
    media bookkeeping correctly, and drops the orphaned media on save
    (Module 1: 68 slides / 34 MB -> 5 slides / 2.4 MB). One sidecar per
    source deck, named `_handoff_polls_<KEY>.pptx` to match the splice
    map's source keys.
  - **Re-key the splice map to the sidecar's own slide numbers** (the kept
    slides keep their relative order, so 7, 8, 25, 29, 46 become 1-5), and
    **delete map entries whose target slide no longer exists** rather than
    leaving them pointing at slides the sidecar does not contain.
  - **The sidecars are BUILD INPUTS — never delete them**, and never
    round-trip them through python-pptx.
  - **Verify before deleting the originals:** rebuild, diff the result
    against the shipped deck (it must be identical), confirm every spliced
    slide still has its `tags` + `notesSlide` + `image` rels, and run the
    full-screen slideshow probe on EVERY poll slide — the notes-part trap
    above is exactly what a sidecar built the wrong way would spring.
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
- **A repeated instruction means the MEASUREMENT is wrong, not the edit**
  (2026-08-30, Nico – after the same slide-28 box size had to be pointed
  out four times). If the same hand-edit has to be pointed out twice,
  STOP. Do not re-apply the same change harder, and do not re-read the
  same dump. The second request is evidence that the number being read is
  not the number on screen – almost always a `grpSp` whose `ext` differs
  from its `chExt`, so the child shape still reports its ORIGINAL size
  while the card renders at another. Re-measure with a different,
  transform-aware tool before touching the build script, and say which
  number changed.
  - **One geometry reader per folder, and it decodes groups.** A probe
    that prints raw child coordinates is a trap dressed as a
    convenience. Where a second script exists because it shows run-level
    formatting (fonts, bold, colour), use it for FORMATTING ONLY – every
    coordinate comes from the transform-aware reader.
  - **Run the closing diff and REPORT it.** `_diff_slides.py <deck>
    <fresh build>` belongs at the END of a hand-edit port, not only at
    the start to find the edits. "Adopted" is a claim about that output:
    if the diff was not run, say the port is unverified rather than
    done.
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
