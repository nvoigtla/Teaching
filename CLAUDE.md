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

## Concept-explanation textboxes (preferred format)
- When a slide needs a short, visually-distinct callout to **explain a
  concept**, define a notational **convention**, or record any other
  brief conceptual aside that sits alongside the main slide content,
  use the **cream-fill rounded-rect "Convention" callout** pattern.
- Style:
  - Rounded rectangle, ~12 % corner radius (slight rounding).
  - Cream / soft-yellow fill (e.g., `#FDF6E6`).
  - Thin primary-color border (~1 pt, navy).
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

## Working with .pptx Files
- PowerPoint files are **binary**, so VS Code visual diffs do not
  work for them.
- Before making any edits to a .pptx file, **create a timestamped
  backup**: e.g., `slides_backup_2026-05-11.pptx`. Keep it in the
  same folder until I confirm I'm satisfied with the new version.
- For substantive edits (rewording, restructuring, adding slides),
  **summarize the proposed changes in chat first** in a clear list
  before touching the file. Wait for my confirmation before applying.
- For minor edits (typo fixes, single-word changes), proceed but
  describe what changed afterward.
- When editing slides, **preserve the existing visual style** (fonts,
  colors, master slide layout, header/footer) unless I explicitly
  ask to change it.
- When adding new slides, **match the layout of the surrounding
  slides** so the deck feels coherent.

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
- **Natural, flowing spoken English.** Conversational, first person, as if
  reading to camera. Not choppy or staccato — longer sentences are fine
  when they read naturally.
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
  - Backup: `[Slides Name]_backup_YYYY-MM-DD.pptx`

## Folder Structure
- Each distinct course gets its own subfolder under `Teaching\`
  (e.g., `Teaching\405-Fall-2026\`, `Teaching\Macro-EMBA-Spring-2027\`).
- Within a course folder, organize by lecture or topic as I direct.
- A `Session-Notes.md` lives in each course subfolder for that
  course's continuity (per the universal CLAUDE.md rules).
- **Session-Notes location is at the course-folder level only.** If a
  session is started from any nested subfolder of a course (e.g., a
  lecture, module, or topic folder), walk **up** the directory tree
  until you find the course folder's `Session-Notes.md`, and read that
  one. There is exactly one `Session-Notes.md` per course; do not
  create per-subfolder session-notes files.
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
