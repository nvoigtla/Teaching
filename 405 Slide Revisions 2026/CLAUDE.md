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

## End of Module, Step 1: Create the Full Slide Version
(2026-09-24, Nico.) Run this once a module's videos are all taped. The
step creates the **full slide version** of a module from `Module X -
Revised.pptx` and the recorded video slides.
**Trigger:** something like "I have now taped the videos and made many
more edits along the way. All taped slides are in the 'Recorded Video
Slides' folder. Go ahead and create the full Module X slide deck." I say
"all taped slides" or "all recorded slides" for this, never "all final
slides" – "final" belongs to step 2.
Short name: **"Create the full slide version (based on `Module X -
Revised` and on the recorded video slides)"**. Recognize the routine
from context. If it is unclear whether I mean it, ask before starting.

**What follows step 1.** The folder cleanup (step 1b, below) runs once
I have looked at the Full deck. Then, once I have taught the module and
hand-polished the in-class slides, **step 2 produces the very final
teaching version** – see "End of Module, Step 2" below. Step 1 never
writes `Module X - Final.pptx`: that name belongs to step 2 and to no
other step. In a request, the word **"final"** means step 2, while
**"taped"** or **"recorded"** means step 1.

**Inputs and output**
- `Module X/Recorded Video Slides/` holds one deck per video
  (`Module X - Video k - <topic>.pptx`). These decks are **fixed
  exactly as they are**.
- `Module X - Revised.pptx` is the reference for slide order, the
  in-class slides, and the backup section. It is read, never
  written.
- The output is `Module X - Full.pptx` in the module folder. It is a
  main deliverable, so roll the `_t-1` / `_t-2` backups before each
  write.

**Rules**
1. **Do not regenerate or rebuild anything.** No build script, splice,
   grouping or animation pass runs on the taped slides. Their `.pptx`
   is the source of truth from here on. The build script stops being
   the source of truth for this module; mark it STALE. The Full deck
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
     goes into Full at its taped position.
   - **A slide used in more than one video** appears once, at its
     Revised position.
   - **A Revised slide that looks REPLACED by a taped one** (e.g. one
     slide split into two during taping): do not decide. List the
     case and ask me.
5. **The opening of the deck.**
   - Drop the "Introduction to Module X" title card that opens
     `Module X - Video 1`. The Full deck opens with the deck title
     slide. The other video title cards stay.
   - If the main deck's title slide carries a picture, keep it.
6. **Example slides that are in NO video deck appear TWICE**
   (2026-09-24, Nico). The Full deck must stay complete even if
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
   separate decks and are not copied into Full.
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
- Re-run the comparison of rule 2 between Full and each video deck.
  Every taped slide must diff clean against its source.
- Run the COM click-count check against the video decks, and one
  full-screen slideshow probe of the whole deck.
- Report, slide by slide: slides replaced (with what changed), slides
  moved to the in-class section, slides left untouched, and every
  pairing the fingerprint could not settle.

## End of Module, Step 1b: Clean Up the Module Folder
(2026-09-24, Nico. Written up from the Module 6 cleanup.) Run this after
`Module X - Full.pptx` exists and I have looked at it, and AGAIN after
step 2 has produced `Module X - Final.pptx`. It is its own step, triggered
by something like "clean up the Module X folder" – never done on your own
initiative. The universal CLAUDE.md's "Cleaning Up a Finished Project"
governs; what follows is what is specific to a taped module.

**Commit first.** This step deletes source decks and the Revised deck, and
git history is then the only way back. Refuse to delete anything the repo
does not already hold, and say so.
- **`Module X - Full.pptx` is the one exception** (2026-09-24, Nico). When
  `Module X - Final.pptx` exists alongside it, delete Full whether or not
  git holds it. Final IS the backstop: it contains every slide Full has, so
  nothing is lost. Do not hold Full back waiting for a commit.

**The cleanup depends on HOW FAR the module has got** (2026-09-24, Nico).
The same trigger runs at two different points, and what survives differs.
**Establish which point you are at before deleting anything**, from the
newest deliverable in the folder:
- **Phase A – Full is the newest deck.** The module has been taped but not
  yet taught, so the in-class slides are not polished and step 2 has not
  run. `Module X - Full.pptx` is the deliverable and is KEPT, and so is
  `Module X - In Class.pptx` **if that file exists at all** – it is the
  input step 2 will read.
- **Phase B – Final exists.** Step 2 has run and the in-class slides are
  finalized. `Module X - Final.pptx` is the deliverable; **whenever the two
  decks sit side by side, `Module X - Full.pptx` is superseded and goes** –
  the way Revised did when Full arrived, and regardless of whether git holds
  it yet. `Module X - In Class.pptx` is KEPT, because it is the source of
  truth for every in-class slide and the only way to re-copy one.
- Where it is not obvious which phase you are in, say so and ask. Deleting
  Full in phase A destroys the deliverable.

**The test for every file:** would I need it to edit the newest deck
tomorrow, or to run the next step on it? Where the answer is unclear, keep
the file and name it in the report instead of guessing.

**The module folder ends with six items** (five where the module has no
in-class deck)
- **The deliverable** – `Module X - Full.pptx` in phase A, `Module X -
  Final.pptx` in phase B.
- `Module X - In Class.pptx` – kept in both phases (when it exists). In
  phase A it is step 2's input; in phase B it is the source of truth for
  every in-class slide.
- `Recorded Video Slides/` – the taped decks. The only source if a taped
  slide has to be re-copied. Kept in both phases.
- The podcast / video-overview sources (`Podcast Module X -- Intro.md` and
  the rest).
- `Session-Notes.md`.
- `_build_files/` – everything else, see below.

**Delete**
- The colleague's source decks and mine (`Module X - NV Slides/`,
  `Module X - PG Slides/` and the like). These are source decks, so **ask
  before deleting them**, per the standing-authority rule in the Teaching
  CLAUDE.md.
- `Module X - Revised.pptx`, superseded by Full.
- `Module X - Full.pptx` **whenever `Module X - Final.pptx` exists**,
  superseded by it. This is the one deletion that does not wait for a
  commit: Final holds every slide Full does, so it is the backstop. In
  phase A, where there is no Final, Full is the deliverable – never delete
  it there.
- A pre-taping copy of a practice-video deck at the folder root, where the
  taped copy is already in `Recorded Video Slides/`.
- `Module X - Revised - outline.md`.
- The rolling `_t-1` / `_t-2` backups.
- One-off scripts written for a single fix, a single probe, or a single
  extraction.
- Render / probe / export folders, `__pycache__`, `~$*.pptx` lock files.

**Move everything else into `_build_files/`**, including `_source_images/`.
Keeping the images beside the scripts is deliberate: `_mX_helpers.py`
resolves `SRC_IMG_DIR` relative to its own file, so moving both together
leaves that reference untouched.

What belongs there, by role:
1. **The drawing layer** – `_mX_helpers.py`, `_build_ModuleX.py` (STALE),
   `_build_template_samples.py`, the notes modules, `_source_images/`.
   This is what a NEW slide in the deck's style needs: build it into a
   one-slide side deck, then `InsertFromFile` it into the deliverable.
2. **Builds and grouping** – `_animate.py`, `_group_pass.py`,
   `_splice_media.py`, and the poll sidecar `_handoff_polls_MX.pptx`.
3. **Read-only auditors** – `_audit_format.py`, `_check_labels.py`,
   `_check_anim.py`, `_check_notes.py`.
4. **Dumps and diffs** – `_dump_deck.py`, `_dump_raw.py`,
   `_dump_text_raw.py`, `_diff_slides.py`. The raw dumps are how the
   deliverable gets read: python-pptx is blind to `mc:AlternateContent`,
   which wraps every OMML formula box.
5. **Probes** – `_export_probe.ps1`, `_slideshow_probe.ps1`.
6. **The step-1 pipeline** – `_assemble_full.ps1`, `_full_inventory.py`,
   `_full_retag.py`, `_full_verify.py`, `_full_verify_tree.py`,
   `_map_videos.py` and their JSON plans. Kept because the next module
   runs the same routine.
7. **The research dumps** – `_source_inventory_*.md`, `_source_rawtext_*.md`,
   `_assets_manifest_*.md`, `_anim_original_*.md`. Kept **because** the
   source decks are gone: they are that material's searchable text and
   animation record at a fraction of the size.

**Resolve the import chain before deleting any script.** The scripts import
each other, and the chain is not obvious: `_animate.py` imports
`_build_ModuleX.py`, which imports `_mX_helpers.py` plus the notes modules,
and `_mX_helpers.py` imports `_build_template_samples.py`. Deleting a
"one-off" out of that chain silently breaks the animation engine. Walk the
`import _` lines of every script you keep, and check each target exists.

**Repoint every path when the scripts move down a level.** One convention,
applied to all of them:
- `HERE` stays the script's own folder – `_source_images/`, the poll
  sidecar and the JSON plans moved with the scripts, so those references do
  not change.
- `MODULE = HERE.parent` is the module folder, and every reference to a
  DECK or to `Recorded Video Slides/` goes through it. In PowerShell,
  `$folder = Split-Path $PSScriptRoot -Parent`.
- A `parents[N]` that reaches outside the module folder (the course
  calendar, a sibling module) shifts by one.
- **Read-only auditors default to the phase's deliverable** – `Module X -
  Full.pptx` in phase A, `Module X - Final.pptx` in phase B – so running
  one with no argument audits the right deck. Repoint them when step 2
  runs.
- **Writer passes keep pointing at the deleted `Module X - Revised.pptx`,
  deliberately.** An accidental `python _animate.py` then fails loudly
  instead of rewriting the deliverable. To run one, name the deck
  explicitly.

**Say which tools no longer work.** A check that reads the deleted source
decks (a coverage check, a dump keyed to the source decks) stays only if it
carries a header comment saying what to restore from git to use it again,
and the session note repeats that.

**Verification before hand-over**
- Every kept script compiles and every local import resolves. Check this
  **statically** – `py_compile` plus an AST walk over the import
  statements. Never verify by importing: these passes do their work at
  module level, so an import RUNS them.
- The `.ps1` files parse (`[System.Management.Automation.Language.Parser]::ParseFile`).
- Run one read-only auditor against the deliverable **from its new
  location**, and compare the findings to a run from before the move. They
  must match.
- The module folder holds exactly the items listed above for the phase it
  is in, and the phase was established before anything was deleted.

**Then write the session note**, listing what was deleted, what
`_build_files/` holds and why, the `HERE` / `MODULE` convention, and any
tool that now needs a restore from git. The note is the point of the step:
the next session has to be able to find the toolkit without reading the
folder twice.

## End of Module, Step 2: Create the Final Slide Version
(2026-09-24, Nico.) Run this once I have taught the module and
hand-polished the in-class deck. The step creates the **final teaching
version** of a module from `Module X - Full.pptx` and `Module X - In
Class.pptx`. It is step 1 run a second time, one layer up: where step 1
checked `Module X - Revised` against the taped video decks, step 2 checks
`Module X - Full` against the hand-edited in-class deck.
**Trigger:** something like "I have now updated the in-class slides by
hand. Go ahead and create the final Module X deck." The word **"final"**
belongs to this step and to no other; "taped" / "recorded" belongs to
step 1.
Short name: **"Create the final slide version (based on `Module X - Full`
and on the hand-edited `Module X - In Class`)"**. Recognize the routine
from context. If it is unclear whether I mean it, ask before starting.

**Inputs and output**
- `Module X - In Class.pptx` is **the source of truth for every slide it
  holds**, exactly as the taped decks are in step 1. I have edited it by
  hand, so it is **fixed exactly as it is** and is read, never written.
- `Module X - Full.pptx` is the reference for slide order, for the video
  blocks and for the backup section. It is read, never written.
- The output is `Module X - Final.pptx` in the module folder. It is a main
  deliverable, so roll the `_t-1` / `_t-2` backups before each write.

**Rules**
1. **Do not regenerate or rebuild anything.** No build script, splice,
   grouping or animation pass runs. The build script is already STALE from
   step 1 and stays that way. `Module X - Final.pptx` is edited in place
   from then on, and the freeze guard covers it as well.
2. **Compare every slide of `Module X - Full` with its in-class
   counterpart**, and check for EVERY kind of change: text (wording,
   color, font, size, bold / italic / underline, bullets, spacing), shapes
   (position, size, fill, line, rounding, shadow), pictures (the image
   itself, crop, size, formatting), grouping, animations (`<p:timing>`
   beat for beat), speaker notes, hyperlinks and hidden status. Same
   tooling as step 1 – the member-level geometry diff with groups decoded,
   a run-level format diff, a notes diff and a timing diff. **Pair slides
   by content fingerprint, never by slide number.** Exclude the live
   `slidenum` field's cached text: it moves with a slide's position and is
   not a difference between the decks.
3. **A changed slide is replaced by the in-class version.** Copy it over
   with PowerPoint COM (`Slides.InsertFromFile`, keeping the source
   formatting), never through python-pptx.
4. **THE VIDEO BLOCKS ARE FROZEN.** A slide inside one of Full's video
   blocks is never replaced by an in-class version – step 1 made the taped
   decks the source of truth for those slides, and step 2 must not undo
   it. Where the in-class deck holds an edited copy of a video slide, that
   edit applies to the slide's IN-CLASS copy only. List every such case in
   the report.
5. **Slide order.** Within the in-class part, follow the IN-CLASS order.
   Everything outside it keeps Full's order.
   - **A slide added in the in-class deck** (not in Full) goes into Final
     at its in-class position.
   - **A slide used in both a video block and the in-class part** appears
     in both places, as it does in Full.
   - **A Full slide that looks REPLACED by an in-class one** (one slide
     split into two while teaching, say): do not decide. List the case and
     ask me.
6. **A slide I DROPPED from the in-class deck is kept, not deleted.**
   Slides usually come out of the in-class deck for time, not because they
   are wrong, and Final is the complete teaching record. Such a slide
   keeps its Full position and its tag. **List every one of them in the
   report** so I can say which should go; never delete a slide on your own
   (the standing rule in `Session-Notes.md`).
7. **Closing sections keep step 1's order:** main deck → Backup →
   in-class section, where the module has one. A backup slide edited in
   the in-class deck keeps its Backup position.
8. **Every jump link and "← Back" button must land on the right slide**
   after the reordering. Links point at slide ids, so re-point every link
   that came across from another deck, and check the whole map before and
   after.
9. **Live content stays functional.** Poll slides move with their notes
   part and `tags` part (the PollEv add-in reads both). Where the same
   poll exists in both decks, the in-class copy wins – it is the one with
   the current activity. Embedded and online videos keep their links.
10. **An EMPTY placeholder slide is reported, never silently shipped.**
    The in-class deck can carry blank slots I added for a PollEv sheet or
    a screenshot I had not captured yet. List each one with its neighbors
    and ask whether it is still wanted.
11. **Flag an unusually small in-class deck before building**, the way
    step 1 flags a short video deck: many dropped slides under rule 6
    means the deck may be a partial copy. Confirm with me first.
12. **Practice-video decks are not part of this routine.**

**Verification before hand-over**
- The file opens in PowerPoint. Its slide count equals Full, plus the
  in-class-only slides, minus anything I explicitly said to drop.
- Re-run the comparison of rule 2 between Final and `Module X - In
  Class`. Every in-class slide must diff clean against its source.
- **Re-run step 1's verification too:** every taped slide in Final must
  still diff clean against its video deck, and every untouched slide
  against Full. This is the guard that step 2 did not disturb step 1.
- Run the COM click-count check, and one full-screen slideshow probe of
  the whole deck. Confirm every live poll renders in the show.
- Report, slide by slide: slides replaced (with what changed), slides
  added from the in-class deck, slides dropped from the in-class deck and
  therefore kept (rule 6), in-class edits to video slides that were NOT
  applied (rule 4), slides left untouched, and every pairing the
  fingerprint could not settle.

**What follows step 2.** The folder cleanup (step 1b) runs again, this
time in phase B: `Module X - Final.pptx` is the deliverable, `Module X -
Full.pptx` is deleted as superseded, and `Module X - In Class.pptx` and
`Recorded Video Slides/` are kept.
