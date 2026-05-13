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
