# 405 Slide Revisions 2026 — project CLAUDE.md

Project-specific conventions for the module-by-module rebuild of my
Management 405 decks. The universal and Teaching CLAUDE.md layers
apply in full; this file only adds what is specific to this project.

## Comparison-Deck Adoption Protocol
Each module rebuild starts from MY old deck and consults a colleague's
parallel deck for innovations (CT's deck for Module 2, MW's for
Module 1). Rules:
- Propose imports item by item in the outline file; I approve or
  reject each. Never sweep in colleague content on your own.
- Tag adoptions in the outline as `[NEW – MW]` / `[NEW – CT]` with the
  source slide number.
- My existing application / mini-case slides keep their place and
  wording; colleague material only ADDS (or replaces after my
  explicit approval).
- Colleague polls run on THEIR PollEverywhere accounts — never splice
  those. When an adopted example changes what my own poll should ask,
  flag that I need to reword the activity in my PollEv account (the
  embed URL stays valid; the static screenshot keeps the old wording
  until I re-capture it).

## File Naming (this project)
- Canonical deck: `Module X - Revised.pptx` (backups per the Teaching
  rules — two newest only).
- Example-candidates review deck: `Module X - Example Candidates.pptx`
  with build script `_build_MX_candidates.py`, which imports the
  module build script's helper layer.
- Research dumps (build inputs, keep): `_source_inventory.md`
  (per-deck text / notes / media), `_runfmt_dump.md` (run-level
  emphasis), `_assets_manifest.md` (extracted images + positions).
- Web-fetched photos: `_source_images/web_*.jpg`, fetched and reviewed
  via `_fetch_web_images.py`.

## Pipeline (per module — the Module 7 rerunnable pattern)
`_build_ModuleX.py` → `_splice_media.py` (polls / live content,
verbatim with notes + tags) → `_group_pass.py` (box+text, shade+frame,
pic+caption groups) → `_animate.py all apply` (fade builds per
per-slide plans). Verify with the COM click-count check
(`_verify_anim.ps1`) and at least one full-screen slideshow probe
before handing a deck over. Hand-edits are surfaced with a
member-level geometry diff (`_diff_slides.py`) against a side-path
build and ported into the build script with a dated comment.

## End of Module: Assembling the Final Deck from the Taped Slides
(2026-09-24, Nico.) Run this once a module's videos are all taped.
**Trigger:** something like "I have now taped the videos and made many
more edits along the way. All final slides are in the 'Recorded Video
Slides' folder. Go ahead and create the final Module X slide deck."
Short name: **"Finalize Module X"**. Recognize the routine from context.
If it is unclear whether I mean it, ask whether I mean the finalize
routine before starting.

**Inputs and output**
- `Module X/Recorded Video Slides/` holds one deck per video
  (`Module X - Video k - <topic>.pptx`). These decks are **FINAL
  exactly as they are**.
- `Module X - Revised.pptx` is the reference for slide order, the
  in-class slides, and the backup section. It is read, never
  written.
- The output is `Module X - Final.pptx` in the module folder. It is a
  main deliverable, so roll the `_t-1` / `_t-2` backups before each
  write.

**Rules**
1. **Do not regenerate or rebuild anything.** No build script, splice,
   grouping or animation pass runs on the taped slides. Their `.pptx`
   is the source of truth from here on. The build script stops being
   the source of truth for this module; mark it STALE. The Final deck
   is edited in place from then on.
2. **Compare every slide of `Module X - Revised` with its taped
   counterpart**, and check for EVERY kind of change: text (wording,
   color, font, size, bold / italic / underline, bullets, spacing),
   shapes (position, size, fill, line, rounding, shadow), pictures
   (the image itself, crop, size, formatting), grouping, animations
   (`<p:timing>` beat for beat), speaker notes, hyperlinks and hidden
   status. Use the member-level geometry diff (`_diff_slides.py`), a
   run-level format diff, a notes diff and a timing diff. Pair slides
   by content fingerprint (`_map_videos.py`), never by slide number.
3. **A changed slide is replaced by the taped version.** Copy it over
   from the video deck with PowerPoint COM
   (`Slides.InsertFromFile`, keeping the source formatting), never
   through python-pptx.
4. **Slide order.** Within each video's block, follow the TAPED order.
   Everything outside the video blocks keeps its Revised order.
   - **A slide added during taping** (in a video deck, not in Revised)
     goes into Final at its taped position.
   - **A slide used in more than one video** appears once, at its
     Revised position.
   - **A Revised slide that looks REPLACED by a taped one** (e.g. one
     slide split into two during taping): do not decide. List the
     case and ask me.
5. **The opening of the deck.**
   - Drop the "Introduction to Module X" title card that opens
     `Module X - Video 1`. The Final deck opens with the deck title
     slide. The other video title cards stay.
   - If the main deck's title slide carries a picture, keep it.
6. **Example slides that are in NO video deck appear TWICE**
   (2026-09-24, Nico). The Final deck must stay complete even if
   everything after the in-class divider is deleted, so that I can
   still teach the whole module in person.
   - **Copy 1 stays in the main deck**, at its Revised position, with
     its tag unchanged.
   - **Copy 2 goes to the in-class section at the very end**, behind a
     divider titled "SLIDES NOT USED IN THE VIDEOS – FOR In-Class
     APPLICATIONS". Model the divider on slide 93 of
     `Module 3 - Revised.pptx`. The copies keep their Revised order.
   - **Copy 2 is retagged `Module X · In Class · Examples · <topic>`**,
     where `<topic>` is the title of the outline item those examples
     belonged to. Splice the level in (see the in-class tag rule in the
     Teaching CLAUDE.md); do not retype the topic.
   - If a video deck carries an edited version of such a slide, both
     copies take that version.
7. **Backup section, and the order of the closing sections.** The deck
   ends: main deck → Backup → in-class section.
   - **A backup slide that a main-deck slide links to stays in Backup**
     (Module 6: the airplane seating chart). It keeps its Revised
     position and order.
   - **A backup slide that nothing links to MOVES to the in-class
     section** (Module 6: the Disneyland case). It is not copied: the
     deck holds one copy only, in the in-class section. (Backup was
     never part of the main deck, so the rule 6 guarantee does not need
     it.) Where its place in that section is unclear, ask me. In
     Module 6 the top-bar tags already name the right section.
   - **Every jump link and "← Back" button must land on the right
     slide** after the reordering, including the links on copy 2 of an
     in-class slide. Links point at slide ids, so re-point every link
     that was copied from another deck.
8. **Live content stays functional.** Poll slides move with their
   notes part and `tags` part (the PollEv add-in reads both; see the
   splice rule in the Teaching CLAUDE.md). Both copies of an in-class
   poll need them. Embedded and online videos keep their links.
9. **Practice-video decks are not part of this routine.** They stay
   separate decks and are not copied into Final.
10. **Flag an unusually small video deck before building.** A video
    deck with only a few slides sends many Revised slides to the
    in-class section. Confirm with me that it is complete first.

**Verification before hand-over**
- The file opens in PowerPoint. Its slide count equals Revised, plus
  the taped-only slides, plus the in-class copies, plus the new
  divider, minus the dropped introduction card.
- Deleting everything from the in-class divider onward must leave a
  complete deck, with every example in its original position. The
  unlinked backup slides that moved to the in-class section are the
  one exception.
- Re-run the comparison of rule 2 between Final and each video deck.
  Every taped slide must diff clean against its source.
- Run the COM click-count check against the video decks, and one
  full-screen slideshow probe of the whole deck.
- Report, slide by slide: slides replaced (with what changed), slides
  moved to the in-class section, slides left untouched, and every
  pairing the fingerprint could not settle.
