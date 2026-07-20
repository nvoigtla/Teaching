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
- **Only the two most recent backups are kept.** When you create a new
  backup, automatically delete any older ones for that deck, so at most
  the two newest backup files remain in the folder at any time.
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
