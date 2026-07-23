# Italy IBR (MGMTEX 421, Class 1 – "The Italian Economy") – Session Notes

Continuity log for the International Business Residential (Italy) deck. The
course is **MGMTEX 421 – International Business Residential in Italy**; this
deck is **Class 1**, the online big-picture class (first class held online;
the trip to Milan & Turin is in September). Audience wording everywhere:
"executives pursuing an MBA at UCLA Anderson."

Canonical deck: **`Class 1 - Revised.pptx`** (widescreen, 405 visual language).
Original source (untouched): **`Class 1.pptx`** (4:3, 123 slides).

---

## Resume here (start of next session)

- **Status:** deck is **132 slides**, opens cleanly, and everything is
  **committed and pushed to `main`** (working tree clean). This folder now
  lives at `Teaching/Italy IBR/` (moved up out of `405 Slide Revisions 2026/`).
  The single design/build rulebook is `Teaching/CLAUDE.md`.
- **The deck is the source of truth** (edited in place via OOXML; never
  round-trip through python-pptx — it strips poll `tags` + null links). Only
  the **reusable engine scripts** remain in the folder (the one-off build
  scripts were deleted — recoverable from git history):
  - `_animate.py` — rebuild animations after any structural change:
    `python _animate.py all apply` (or specific slides). Holds `FIG_GROUP`
    (figure→bullet) and `PIC_BULLET` (per-picture→bullet, e.g. slide 6).
  - `_group_shades.py` / `_group_boxes.py` — group figure+shade / colored
    box+text; **re-run `_animate.py` afterward** (grouping wipes timing).
  - `_resize_bullets.py apply` — re-fit bullet sizes.
  - `_add_notes.py apply` + `_notes_data.py` — edit speaker-note text in
    `_notes_data.py`, then re-apply (edits notesSlides in place; poll notes
    are protected).
- **Open items:**
  - Delete the empty stale folder `405 Slide Revisions 2026/Italy IBR/` — it
    was locked while the prior session was running; remove it now that this
    session is closed.
  - **Slide 132** (poll "Italian-Speakers in 1861?") is an orphan backup that
    nothing links to → it got **no** Back button. Add one → slide 61 if wanted.
  - Podcast **Parts I–IV** emails are drafted (link placeholders) — paste the
    NotebookLM links and double-check the per-part summaries.
  - Optional, previously flagged: slide **20** has no figure (too text-dense —
    could split); slide **47**'s plague image is Naples 1656 (captioned so);
    slides **49 / 100 / 101** reveal their image clusters on a single click.
- **Continuity note:** the chat transcript does *not* carry into a session
  opened from this folder — this file is the handoff.

---

## 2026-07-23 (session 2) – map, box grouping, back buttons, folder move, cleanup

- **Slide 15:** replaced relief map rounded + shadowed; normalized its messy
  embedded media name (`image8.jpg&w=828&q=100` → `relief15.jpg`).
- **Colored box + text grouping:** grouped each takeaway bar / card with its
  text into one `<p:grpSp>` on slides **15, 25, 29, 43, 64, 90, 91, 102** (fixes
  "box appears before its text"); updated `_animate.py` to detect the groups
  (`TakeawayGroup` = final click; `CardGroup` = column beats) and re-ran
  animations. Added the rule to `Teaching/CLAUDE.md`.
- **Slide 6:** each hotel photo now fades in with its bullet (top → "Hotel in
  Milan", second → "Hotel in Turin") via the new `PIC_BULLET` map.
- **Back buttons:** added to backups **115–117 → 31, 119 → 53, 120 → 55,
  121–131 → 57** (cloned from slide 114). **Deleted empty slides 133–135**
  (deck 135 → 132).
- **Structure:** consolidated the course-layer CLAUDE.md into `Teaching/CLAUDE.md`
  and moved `Italy IBR/` up to `Teaching/`; Session-Notes are now per-subfolder.

---

## 2026-07-23 – Consolidated log for the Class 1 build (covers 07-18 → 07-23)

**One-line summary.** Rebuilt the old "Class 1" deck into the 405 visual
language (135 slides), enriched the economic-history content with new slides
and figures, refreshed the presentation-topics slide from Excel, wrote
teleprompter-style speaker notes for the whole deck, added Fade/on-click
animations, and grouped table shades into their figures. Deck committed and
pushed once (`e0fa35d`); later edits (shade grouping, CLAUDE.md consolidation,
this notes restructure) are still local.

### Current deck state
- **`Class 1 - Revised.pptx` — 132 slides** (was 135; 3 empty trailing slides
  deleted 2026-07-23), opens cleanly in PowerPoint.
- Preserves the **7 live PollEverywhere embeds** (slides 11, 13, 21, 35, 48,
  69, 132) and all **internal slide-jump links** (e.g. slide-4 "Schedule" →
  trip itinerary; company/topic → backup detail slides; "← Back" buttons).
- Footer page numbers are **live `slidenum` fields**.

### What was done (grouped)
1. **Reformat.** Rebuilt every script-buildable slide in the 405 style
   (navy/gold chrome, Calibri, three-level section tags, live footer fields,
   native tables). Kept original slide titles and bullet wording; converted
   4:3 → widescreen 13.33×7.5.
2. **Economic-history enrichment** (from the podcast research): new slides —
   geography/rugged-map (15), Rome "First Integrated Market" (25), Empire at
   Height map (26), the Fall/fragmentation (28–29), **Renaissance Economy
   (43)**, **Late Industrialization & IRI (64)**, **Economic Miracle (90)**,
   **Third Italy (91)**, **The Paradox (102)**; plus enrichment bullets on
   31, 33, 46, 48, 59, 60, 97 (Arab Sicily, Legnano/Peace of Constance,
   maritime republics, Columbus/Atlantic shift, refeudalization, 1630 plague,
   Napoleon reforms, 1887 tariff, emigration, fourth capitalism).
3. **Slide 107 presentation topics** rebuilt from `Presentation Topics --
   Italy.xlsx`: 6 debate topics (Topic / Team A / Team B) + a single
   full-width **"Company Option" block with no A/B split** (company pills;
   #6 a dashed "TBD (Wed 11am)" placeholder). Back button → slide 8.
4. **Figures.** Added 11 CC-licensed illustrative figures to text-only High +
   Medium slides (16, 22, 24, 27, 33, 34, 40, 42, 47, 89 — slide 31 excluded
   by request, slide 20 dropped as too text-dense). Each has a sourced italic
   caption, rounded corners, soft shadow. Sourced via a research subagent from
   Wikimedia/PD/CC. **Slide 47's plague image is Micco Spadaro's Naples 1656
   painting** (no clean Milan-1630 image existed) — captioned honestly.
5. **Speaker notes** on every content slide: teleprompter-style narration for
   a general reader (first person, read-aloud, extra background), time-neutral,
   no dash-asides, each pointing at the posted picture. ~1 anecdote in 3, set
   off as `ANECDOTE — …` with a verified `Source:` link. Logistics (3–9),
   agenda/dividers, and the 7 poll slides get **no** notes (poll notes are
   protected); backups get one brief line. Slide 17's original Celts link kept.
6. **Animations** (Fade, ~0.5 s, on click): 85 slides built. Content slides
   build one top-level bullet per click (first bullet shows with the slide;
   sub-bullets ride with their parent). **Side figures fade in on the click of
   the bullet they illustrate** (default = first built bullet; overrides:
   16 → "Advanced art", 90 → "Icons/Fiat 500", 47 → "1630 plague"). Galleries
   reveal one picture per click; two-column 102 = left card → right card →
   takeaway; featured-research = paper then authors; takeaway bars last.
   Logistics 3–9 also animated (per-main-bullet). Skipped: title (1–2),
   agenda/dividers, poll slides, backups (105–135).
7. **Shade grouping.** Slides 7 and 107 had the table's shadow on a separate
   backing rectangle; grouped each backing + table into one `<p:grpSp>` so the
   shade travels/animates with the figure, then re-ran animations.
8. **Emails** drafted for students (not in the deck): IBR Custom Podcast
   Parts I–IV announcements (link placeholders; Nico pastes NotebookLM links).

### Key decisions / conventions
- **Canonical deck edited IN PLACE via OOXML surgery** (zip + lxml); **never
  round-trip through python-pptx** — it strips poll `tags` rels and NULL-target
  internal links. New slides are built in a temp deck then spliced.
- **Display→part mapping** via `ppt/presentation.xml` `<p:sldIdLst>` → rels
  (never assume `slideN.xml` = display N).
- **Render/verify via PowerPoint COM** (no LibreOffice); open in PowerPoint as
  the real integrity check after any structural edit.
- `_manifest.json` keys the 7 poll GUIDs + internal links to build functions so
  splices are insertion-robust.
- Figures: pictures carry their own shadow; tables/charts put the shade on a
  backing rect and **group it with the figure** (now a CLAUDE.md rule).
- Backups: keep only the two newest + the original; **gitignored** (see root
  `.gitignore`, `*_backup_*.pptx`).

### Build scripts (source of truth, in `Italy IBR/`)
`_build_Italy_Class1.py` (STALE scaffold; imports Module 3 template helpers),
`_splice_polls.py`, `_wire_links.py`, `_manifest.json`, `_resize_bullets.py`
(28/24 or 24/22 fit, measured with PIL), `_audit_fix.py`, `_build_new_slides.py`
/ `_add_slides.py` / `_slide_polish.py` / `_enrich_more.py`, `_build_more_slides.py`
/ `_add_more.py` (the 5 new slides), `_topics107.py`, `_add_figures.py`
(+ `_notes_data.py` / `_add_notes.py` for speaker notes), `_animate.py`
(animation engine; `FIG_GROUP` maps figure→bullet), `_group_shades.py`.
Rerun animations with `python _animate.py all apply`.

### Pending / open items
- **Commit + push the local changes:** shade grouping (deck), `_group_shades.py`,
  the consolidated `Teaching/CLAUDE.md` (+ deleted `405 .../CLAUDE.md`), and this
  Session-Notes restructure. (User to confirm the commit.)
- **Podcast links** — Nico pastes the four NotebookLM links into the drafted
  emails; verify the Part II/III/IV content summaries match his actual splits.
- **Slide 20 (Roman Republic)** has no figure (too text-dense) — offer to split
  into two slides if a figure is wanted there.
- **Slide 47** plague image is Naples 1656, not Milan 1630 — swap if preferred.
- **Slides 49 / 100 / 101** reveal their image clusters on one click (they have
  ≥2 bullets); switch to one-image-per-click if desired.

### Useful commands
```powershell
# Render a slide to PNG + integrity-check open (kill stale POWERPNT first).
Get-Process POWERPNT -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
# ... COM: $pp.Presentations.Open(deck,$true,$false,$false); $slide.Export(png,"PNG",w,h)

# Re-run all animations after any structural change (grouping, splices):
python _animate.py all apply        # or: python _animate.py 16 90 apply  (specific slides)

# Re-apply speaker notes (edits notesSlides in place; poll notes untouched):
python _add_notes.py apply
```
