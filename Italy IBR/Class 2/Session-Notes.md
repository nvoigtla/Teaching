# Italy IBR (MGMTEX 421, Class 2) – Session Notes

Continuity log for Class 2 of the International Business Residential (Italy).
Course: **MGMTEX 421 – International Business Residential in Italy** (trip to
Milan & Turin, Sep 6 – 11, 2026 (flights home the 12th)). Audience wording everywhere: "executives
pursuing an MBA at UCLA Anderson."

Canonical deck: **`Class 2 - Revised.pptx`** (16 slides, widescreen 13.33 × 7.5",
Class 1 visual language). Original source (untouched): **`Class 2.pptx`** (2023,
4:3, 9 slides). The design/build rulebook is `Teaching/CLAUDE.md`; Class 1's
build log lives in `../Class 1/Session-Notes.md`.

---

## Resume here (start of next session)

- **Status:** deck built (16 slides); presentations are identified by **group
  number only** and run **in ascending group order**. Opens cleanly in
  PowerPoint, animations verified via COM, full-screen slideshow probed (no
  renderer failure). **Not yet committed** – awaiting Nico's go-ahead.
- **No open questions.** Poll: Nico updates it himself. Garbled tech/AI
  statement: fixed at the source. PRO/CONTRA sides: settled (see below).
  Committed and pushed 2026-08-22.

## Rerunnable pipeline

Phase-3-as-pipeline (Module 7 pattern) – `_build_Class2.py` stays the source of
truth, every later pass is idempotent and re-run after each rebuild:

```
python _build_Class2.py      # 15 script-buildable slides
python _splice_poll.py       # inserts the PollEverywhere slide at display #4
python _animate.py all apply # injects <p:timing> builds
```

- `_slideshow_probe.ps1` – runs the real full-screen Slide Show and screenshots
  the `screenClass` window via `PrintWindow` (flag 2). Run after any OOXML
  surgery; the editing canvas and PNG export can pass while the show fails.
- Render to PNG for eyeballing: PowerPoint COM `Slide.Export` (no LibreOffice
  on this machine). Outputs land in `_probe/`.
- `_make_itinerary.py` regenerates `_source_images/itinerary_2026.png` from the
  PDF, applying the two edits described in the 2026-08-22 entry below. Run it
  before `_build_Class2.py` if the itinerary PDF is ever replaced.
- Build inputs in `_source_images/` – **never delete**: `itinerary_2026.png`
  (generated, 2028 × 1546) and the five company logos pulled from Wikimedia
  Commons.

## Deck map (display numbers)

| # | Slide | Animation |
|---|---|---|
| 1 | Title – "Presentations and Debates" / Class 2 | static |
| 2 | Roadmap – 3 numbered-circle bands | static (agenda) |
| 3 | Logistics – The Schedule from Here ("Schedule" → 16) | 2 clicks |
| 4 | **PollEverywhere** – "I will arrive at the hotel in Milan" | static (spliced) |
| 5 | Divider – Presentations and Debates | static |
| 6 | Today's Line-Up – 9 Presentations, in group order | static (agenda) |
| 7 | Debate – Italy and the EU / the Euro (Grp 1 CONTRA / Grp 3 PRO) | static |
| 8 | Kering – Group 2 | static |
| 9 | Debate – Italy's Future in Tech and AI (Grp 4 PRO / Grp 7 CONTRA) | static |
| 10 | Prada Group – Group 5 | static |
| 11 | Inter Milan FC – Group 6 | static |
| 12 | Pirelli – Group 8 | static |
| 13 | EssilorLuxottica – Group 9 | static |
| 14 | Thank You! | static |
| 15 | Backup Slides divider | static |
| 16 | Trip Schedule (Sep 6 – 11) – itinerary + ← Back to 3 | static |

**Slide 3 is the only animated slide in the deck** – 2 clicks on its bullets.
Everything from slide 7 on lands complete (Nico, 2026-08-22).

Page numbers are live `slidenum` fields on 13 of 16 slides; title, poll and
Thank-You are deliberately unnumbered (Class 1 convention).

## Sessions

### 2026-08-22 – Class 2 rebuilt in the Italy IBR format

**What was done.** The 2023 4:3 deck was rebuilt in the Class 1 visual language
and repopulated with the 2026 presentation assignments and the current
itinerary.

- **Chrome lifted verbatim from `Class 1 - Revised.pptx`** (EMU for EMU): navy
  top bar 0 → 384048 with the `Italy IBR · <Section>` tag at 16 pt white bold;
  action title 30 pt navy bold at y 566928; gray rule + gold strip at
  1188720 / 1175004; footer rule + gold strip + "International Business
  Residential – Italy" + live slide-number field. Bullets use Class 1's exact
  `▪` / `–` buChars at marL 342900 / 731520, 28 pt main / 24 pt sub, spcBef
  1200 / 300.
- **Presentation slides driven by `Assigned Groups and Grades.xlsx`, in that
  file's order** (Topic 4 → Topic 2 → Kering → Prada → Inter → Pirelli →
  EssilorLuxottica). Position statements come from
  `../Presentation Topics -- Italy.xlsx`, mapping **PRO = Team A, CONTRA =
  Team B**.
- **Debate slides** are two-column: navy header bars ("PRO · Group 3"), cream
  rounded cards with the position in 26 pt italic quotes, gold takeaway bar
  ("2 – 3 min executive summary from each side, then open debate") on the final
  click.
- **Company slides** pair a Wikimedia logo (flat – logo exception) with the
  sector and a navy group pill on the left, and two cream cards on the right:
  "Your brief" (introduce it, then argue why (not) to invest) and "We visit them
  on the trip" (slot + the itinerary's own description of the session).
- **Dates updated to 2026** (Nico approved): trip Sep 3–9 → **Sep 6 – 12**;
  "Class on Sep 24" → **Oct 4** (matches Class 1 slide 4). Today's class is
  Aug 23.
- **Itinerary on the last slide** as a picture (Nico's choice over a native
  table): PDF page rendered at 3.2× with PyMuPDF, cropped to the grid
  (126,168)–(2560,1714), placed 8.55 × 5.42" centred at y 1.36 with the source
  line beneath and a ← Back button to slide 3.
- **Poll slide spliced, not rebuilt.** `_splice_poll.py` carries slide XML +
  `tags` part (`__PE_POLL_EMBED_ID`) + image + **notes part** together, and
  recentres the 4:3 poll graphic on the 16:9 canvas. The notes part is
  load-bearing: the add-in reads the poll URL from it, and a tagged poll slide
  with no notes crashes the slideshow renderer deck-wide.
- **Animations** (`_animate.py`, Class 1 timing primitives, explicit per-slide
  plan rather than heuristics): 8 slides, 18 click-beats. Verified through
  `MainSequence.Item(i).Timing.TriggerType` – slide 3 = 1,2,2,1 (first bullet
  static, subs ride with their main bullet); 7/8 = 1,2,1,2,1; 9–13 = 1,1.

**Decisions made.**
- Group **numbers only** on the slides, no student names.
- Itinerary as an **image of the PDF page** rather than a native table (the
  7-day grid is too dense to stay readable natively; Class 1 slide 112 set the
  precedent).
- **One slide per company** rather than compact multi-company cards.
- Logos over photographs on the company slides (the slide's job is to name the
  firm the group is presenting).

**Differences between this itinerary and the one in Class 1 slide 112** (both
Legacy Ventures, so the deck now shows the newer one): **Stellantis** (Fri 11
Sep, 09.00 – 11.00, plus a circular-economy facility tour) replaces
**Italdesign**, and the Wednesday 11.00 – 12.30 slot that was "Company Visit
(TBD)" is now **Dolce&Gabbana or Eataly**. Class 1 slide 112 still shows the old
version if that matters for consistency.

### 2026-08-22 (later) – numbering settled: GROUP NUMBERS ONLY

Two steps, ending where it matters: **presentations are identified by group
number, and by nothing else.**

1. First pass made the leftover numbering from the 6-topic / 11-company menu
   sequential in the grade sheet's own order (Topic 4 → Topic 1; Kering →
   Company 1; …), and put matching `Topic N` / `Company N` pills on the slides.
2. Nico then dropped that scheme entirely – the grade sheet and the menu file
   could not agree on company numbers without rewriting the menu students chose
   from, so the numbers went instead of the mismatch.

**Final state.**
- **`Assigned Groups and Grades.xlsx`** – the `Assigned Topic` column now holds
  the plain name: "Italy and the EU/Euro", "Italy's future in tech/AI",
  "Kering", "Prada Group", "Inter Milan FC", "Pirelli", "EssilorLuxottica".
  The `Type` column ("Broad Topic" / "Company") already carries the kind, so
  nothing was lost. Rolled backups: `_t-1` = sequentially numbered version,
  `_t-2` = the original as received.
- **`Group` column untouched throughout** – 3, 1, 4, 7, 2, 5, 6, 8, 9, exactly
  as the students signed up. **Every group number in the deck is read from that
  column; none is inferred or renumbered.**
- **Slides** – no `Topic N` / `Company N` pills. The only numbers are the navy
  group pill on each company slide, "PRO · Group 3" / "CONTRA · Group 1" on the
  debate headers, and the line-up slide's circles (running order) with
  "Debate · Groups 3 & 1" / "Company · Group 2" labels.
- `../Presentation Topics -- Italy.xlsx` was never modified – it stays the
  record of the menu students chose from.
- Re-verified after the rebuild: 16 slides, no leftover pill shapes (COM shape
  scan), click structure unchanged (3 = 1,2,2,1 · 7/8 = 1,2,1,2,1 · 9–13 = 1,1),
  slideshow probe clean.

### 2026-08-22 (later still) – slides ordered by group number

Nico located the group numbers and asked the deck to run Group 1, 2, 3, …

- **Running order** is now driven by one table, `SEQUENCE` in
  `_build_Class2.py`. A debate sits at its **lower** group number so both sides
  of a question stay on one slide: **Groups 1 & 3** (EU/Euro) → **2** Kering →
  **4 & 7** (tech/AI) → **5** Prada → **6** Inter → **8** Pirelli → **9**
  EssilorLuxottica. Debates and company slides therefore interleave, and the
  top-bar tag alternates accordingly.
- **Within each debate the columns are ordered by group number too**, so the
  lower group takes the left column and reveals first. On slide 7 that puts
  CONTRA (Group 1) on the left and PRO (Group 3) on the right; on slide 9 PRO
  (Group 4) is on the left. **If you would rather PRO always sat on the left,
  reorder the `sides` list for that debate in `DEBATES`.**
- `DEBATES` entries now hold `sides=[(side, group, claim), …]` in ascending
  group order instead of fixed `pro_`/`con_` fields; the debate shapes are named
  `Hdr1 / Card1 / Hdr2 / Card2` by column, and `_animate.py` beats follow.
- **Roadmap trimmed to 3 bands** ("Logistics", "Arrival in Milan – a quick
  poll", "Seven group presentations – in order of group number"), since the old
  "debates then companies" split no longer describes the running order. Band
  block re-centres itself on the item count.
- Verified after the rebuild: display order and titles dumped via COM (7 =
  EU/Euro, 8 = Kering, 9 = tech/AI, 10–13 = Prada / Inter / Pirelli /
  EssilorLuxottica); triggers 3 = 1,2,2,1 · 7 = 1,2,1,2,1 · 8 = 1,1 ·
  9 = 1,2,1,2,1 · 10–13 = 1,1; slideshow probe clean.

### 2026-08-22 (end of day) – garbled statement fixed at the source

Nico confirmed the reading and asked for the wording to be fixed.

- **`../Presentation Topics -- Italy.xlsx`, cell D10** (Topic 2, Team A):
  *"Italy must is lagging far behind on tech/AI"* → **"Italy is lagging far
  behind on tech/AI"**. The stray "must" was the only thing wrong; the sentence
  now reads as the opposite of Team B's "Italy is already on a good path", which
  is what the debate needs. Previous version kept as
  `Presentation Topics -- Italy_t-1.xlsx`. This is the first edit ever made to
  that file – it is otherwise the untouched record of the menu students chose
  from.
- **The deck needed no change**: slide 9 already carried the corrected wording,
  rendered as **"Italy is lagging far behind on tech and AI"** (the slash
  expanded to "and" to match the slide title "Italy's Future in Tech and AI" and
  the rest of the on-slide prose). Verified there is no "must is" anywhere in
  the package.
- Nico is updating the PollEverywhere activity himself, so the 2023 arrival
  options are no longer an open item here.

### 2026-08-22 (end of day) – tech-AI debate sides corrected

The PRO/CONTRA-to-statement mapping was the last thing I had guessed at. Nico
confirmed it was backwards, so slide 9 now reads:

- **PRO · Group 4** — "Italy is already on a good path"
- **CONTRA · Group 7** — "Italy is lagging far behind on tech and AI"

**Consequence for the mapping rule:** PRO is *not* uniformly Team A of
`Presentation Topics -- Italy.xlsx`. For the EU/Euro debate PRO = Team A; for
tech/AI PRO = **Team B**. What is consistent is the sense of the label — PRO
takes the affirmative / optimistic position in both debates. `DEBATES` in
`_build_Class2.py` carries a comment saying so, since the source files do not
record it. The `Side` column of the grade sheet was not touched: it already says
PRO for Group 4 and CONTRA for Group 7; only which claim goes with which side
changed.

Verified: slide 9 shape text read back via COM (`Hdr1` PRO · Group 4 →
`Card1` "already on a good path"; `Hdr2` CONTRA · Group 7 → `Card2` "lagging far
behind"), triggers still 1,2,1,2,1, column order still ascending by group.

### 2026-08-22 (review pass) – Nico's comments applied

- **"9 Presentations", not 7.** Slide 6's title is now "Today's Line-Up –
  9 Presentations" (9 groups present across the 7 slots, since each debate pairs
  two groups). For consistency the same count was corrected on slide 2's roadmap
  band ("9 group presentations – in order of group number") and in the speaker
  notes on slides 1, 2, 5 and 6 – leaving "Seven" on slide 2 would have
  contradicted slide 6. Slide 6's note now spells out the 9-groups / 7-slots
  relationship.
- **Trip dated Sep 6 – 11 throughout** (the 12th is the flight home): slide 3's
  bullet and slide 16's title. Slide 3's note adds "flights home on Saturday the
  12th" so the 12th is not simply lost.
- **Itinerary regenerated** by a new script, `_make_itinerary.py`:
  - **Wednesday 11.00 – 12.30 is now Eataly** (uncertainty resolved). Done
    without re-typing any text: the time line and the "Eataly / Food & Beverage"
    block are lifted out of the render as image patches, the cell interior is
    cleared with the cell's own background colour, and the two patches are pasted
    back as a vertically centred stack. The table's own font (Metropolis), the
    cyan company colour and the bold sector line are therefore preserved exactly,
    and "Dolce&Gabbana / Luxury Fashion / or" is gone.
  - **Saturday column dropped** – the crop now stops on the Friday|Saturday
    rule, which becomes the new right border.
  - Nothing is hard-coded: vertical/horizontal rules are found by scanning for
    dark rows and columns, the target cell is identified by the rules that
    bracket the "Dolce&Gabbana" text span, and all text positions come from the
    PDF's own spans via PyMuPDF. Asserts fail loudly if the PDF changes shape.
  - New image is 2028 × 1546 (aspect 1.312 vs 1.574 before), placed
    7.14 × 5.44" centred at x 3.10, y 1.36.
- Re-verified: 16 slides, click structure unchanged (3 = 1,2,2,1 · 7 = 1,2,1,2,1
  · 8 = 1,1 · 9 = 1,2,1,2,1 · 10–13 = 1,1), slideshow probe clean on 4 / 6 / 9 / 16.
- **Gotcha re-learned:** a Bash heredoc ate one backslash level in a patch
  script (the `\` line-continuation inside a match string), so the match
  silently failed. Per `Teaching/CLAUDE.md`, edit scripts containing backslashes
  must be written to a `.py` file with the Write tool and run by path.

### 2026-08-22 (final) – animations dropped from slide 7 on

Nico's last note: no animations from slide 7 forward. Once the class hands over
to the students, a presentation slide should be complete the moment it appears –
there is no instructor narration to pace a build against.

- `PLAN` in `_animate.py` is now just `{3: dict(bullets="TextBox 6")}`. The
  debate beats (`Hdr1/Card1` → `Hdr2/Card2` → `Takeaway`) and the company beats
  (`BriefCard` → `VisitCard`) are gone.
- Slides 7–13 are deliberately kept **out of** `SKIP` rather than added to it, so
  a rerun still strips any `<p:timing>` an earlier pass may have written. They
  report "0 click-beats", which is the intended outcome, not a failure.
- Deck-wide effect count is now 4 effects on one slide (was 18 click-beats over
  8 slides). Verified by walking `MainSequence` on all 16 slides: only slide 3
  returns a non-empty sequence (triggers 1,2,2,1); the other 15 report static.
- Slideshow-probed 7 / 9 / 11: each shows its full content at entry (headers,
  cards, logo, brief card), where slide 7 previously opened with chrome only.

## Open items

1. ~~The poll is stale~~ – **Nico updates the PollEverywhere activity himself**
   (2026-08-22). Slide 4's embedded graphic is only the cached 2023 result
   image; the live options come from polleverywhere.com at show time.
2. ~~Garbled tech/AI statement~~ – **fixed 2026-08-22**, see the session entry
   below.
3. ~~PRO / CONTRA ↔ Team A / B for the tech-AI debate~~ – **settled
   2026-08-22**: Group 4 argues the optimistic side. See the session entry below.
4. ~~Group 8's roster looks short~~ – **confirmed fine by Nico (2026-08-22)**;
   no name is missing from the sheet.
5. ~~Push is pending~~ – **pushed 2026-08-22**: `43817e2` (rebuild) and
   `b2f243b` (review-pass fixes) are on `origin/main`.
6. **Optional:** the itinerary picture is now 7.14 × 5.44" – with Saturday gone
   the grid is squarer, so at the maximum height the chrome allows it no longer
   spans the slide width (dead margins left and right). The table text is the
   same physical size as before; rebuilding the schedule as a NATIVE PowerPoint
   table is the only way to fill the width and enlarge the text.
