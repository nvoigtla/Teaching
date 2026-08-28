# Session Notes — Module 1 (combined In-Class + Videos deck)

## PENDING (updated 2026-08-27 — deck is 95 slides)

Open items on **Module 1 - Revised.pptx**:
0. **Podcast URL for display 12** — the Sound button is a marker until
   Nico supplies the link (one line in `_build_Module1.py`).
0c. **Display 24** — the AI-accelerator image has no attribution line.
1. **PollEv rewording (Nico's account).** The diamonds activity on
   display 59 still asks "How does the decline in engagements affect the
   demand for diamonds?" — with the Swift-engagement example the answer
   flips (demand shifts RIGHT). The embed URL stays valid; the static
   screenshot keeps the old wording until it is re-captured.
2. *(resolved 2026-08-27)* The backup slides keep `Module 1 · Backup` —
   the rule was corrected, not the deck. See "Round 4" below.
3. Two smaller flags remain open from round 1: display 53 shares one
   "Photos: Wikimedia Commons" caption across two photos, and the four
   video title cards (1, 10, 17, 27) carry one-line notes the
   video-conversion rule says they should not have.

## 2026-08-27 — adopting Nico's hand pass; format audit; podcasts

**One-line summary.** Nico's hand-edited 95-slide deck was diffed against
a fresh pipeline build of the 104-slide deck; every one of his edits was
ported back into the build scripts, a formatting audit against the latest
Teaching CLAUDE.md was run and its findings fixed, and the two NotebookLM
podcast source docs were written.

### His hand-edits, all adopted into the pipeline
Found with a whole-deck text-signature match plus a member-level
geometry / text / format / notes / timing diff (`_diff2.py`, new —
canonical vs. a side-path build). It matches slides across a REORDER,
which the older `_diff_slides.py` cannot.

- **12 slides deleted:** the "Slides Not Used in the Videos" divider; the
  ADM mini-case; both COVID/tea slides; all three avocado slides; both
  copper slides; the "Related Work by Anderson Faculty" backup; and the
  QUESTION slide of the AC and the diamonds poll pairs (each of those
  polls now runs set-up → PollEv results → solution; the Econ & Coffee
  and flip-house polls keep both views).
- **3 slides added:** the Kroger–Albertsons pair from `Module 1 - Example
  Candidates.pptx` (after Netflix), and a second copy of the
  "Introduction / about the instructor" slide opening the in-class part.
- **2 moves:** the shift-combination table now CLOSES Video 4 (display
  34, still hidden); Homo Economicus moved after the course roadmap.
- **1 geometry tweak:** display 76's child-cost figure raised 2.55 → 2.21.
- The whole front matter now simply belongs to the in-class part.

`_m1_order.py` was rewritten to the 95-slide ORDER; every downstream pass
re-keys through it, so `_splice_media.py`, `_group_pass.py` and
`_animate.py` needed no manual renumbering. `_build_Module1.py` gained
`slide_kroger_case`, `slide_kroger_costco` and
`slide_02_introduction_again`; `apply_fill_notes` now skips keys whose
slide was deleted.

### Formatting fixes applied
- **Tags + footers on 53/54:** the candidates deck's "Module 1 ·
  Candidates · Market Definition" tag and its "Example Candidates (for
  review)" footer replaced by the main deck's.
- **Grouping on 53/54** (group pass): the quote box with its text, and
  each picture with its caption.
- **Animation on 53/54:** 3 clicks on the case slide (chronology first —
  the deal and the photos are static, the market-definition argument
  builds), 4 on the resolution slide with the gold decision bar last.
- **Title case (7 titles):** "…: The Case of Netflix" (×2), "Shortages
  When Disasters Loom", "…Costs When Buying a Present", "What Is the Full
  Economic Cost…", "…(You Should Ignore!)", "…That You'll See Throughout
  the Class", "Can These Prices Be Optimal?".
- **Problem-set pointers:** glyph ➜ → ✎ (the fixed-vocabulary rule),
  routed through new `PS_GLYPH` / `PS_BOX_XY` constants. The convention
  bottom-right corner was TRIED on displays 33 and 49 and reverted with a
  dated comment: on 33 the box covers the chart's "Quantity" axis title,
  on 49 a 3.00"-wide box collides with the last coverage pill.
- **Display 86:** "MC" and "indifferent" raised 14/13 → 16 pt (the
  chart-label floor).
- **Display 54:** the navy market header box widened 6.90 → 7.10" so the
  line stops wrapping with "…" alone on line 2 (measured at 6.81" in
  Calibri Bold 18 pt).
- **Speaker notes:** build-provenance and to-do text removed from the
  notes on displays 34, 53, 54, 57, 58, 74, 89 and rewritten as
  student-facing guidance; the Pike Place demand-curve note, which sat
  IDENTICALLY on displays 20, 21 and 22, was split into three
  (`DEMAND_DEF_NOTE` / `CETERIS_PARIBUS_NOTE` / `DEMAND_CURVE_NOTE`).
  The PollEv rewording reminder was kept, marked "(Nico: …)".

### Flagged, NOT changed (awaiting Nico)
- **Display 36** duplicates display 2 and sits BETWEEN the "In-Class
  Part" divider and the module title slide. Intentional, or should it
  follow the title slide?
- **Display 47** ("Agenda for the Class") carries the tag "Module 1 ·
  Video 1 · Introduction" while sitting in the in-class part; display 7
  is the same slide tagged "Module 1 · Course Roadmap" inside Video 1.
  The two tags look swapped.
- **Display 86:** "Net Benefit of Hour 1 / 2" are 9 pt inside 0.95"-wide
  bars. Reaching the 16 pt floor needs shorter wording ("Net benefit") or
  the label moved outside the bar — a content call.
- **Poll Break badges** (55, 58, 71) sit at 9.991 / 6.770, 2.525 × 0.590
  rather than the CLAUDE.md constant 10.238 / 6.769, 2.95 × 0.533. They
  come verbatim from Nico's hand-tuned `_handoff_pollbreak.xml`, so they
  were left alone.
- **Display 53** uses one shared "Photos: Wikimedia Commons" caption for
  two photos; the animation rule prefers one caption per picture.
- Video title cards (1, 10, 17, 27) carry one-line notes; the
  video-conversion rule says a title card carries none.

### Verification
- `_diff2.py` canonical vs. rebuild: every remaining difference is one of
  the fixes above — nothing of Nico's was lost.
- `_verify_anim.ps1` (expected map regenerated for the 95-slide deck):
  ALL CLICK COUNTS MATCH, 63 animated slides; the deck opens in
  PowerPoint.
- `_slideshow_probe.ps1` on 1, 42, 53, 54, 56, 59, 72, 86, 95: PASS — all
  four live PollEv activities render in the real slideshow.
- All 95 slides eyeballed as contact sheets.

### Podcasts (NotebookLM source docs, new)
`Podcast Module 1 -- Intro.md` (~1,580 words, about 5 minutes, future
tense, prepared host briefing an unprepared one) and `Podcast Module 1 --
Wrap-up.md` (~3,250 words, about 15 minutes, past tense, two students
talking it through). Both follow the Module 3 pattern: H1 exactly
`Module 1 - Podcast Intro` / `Module 1 - Podcast Wrap-Up`, a paste-ready
Audio Overview prompt at the top, then the standing host instructions.
Module-1-specific guard rails in those instructions: movement ALONG a
curve vs. SHIFT of the curve (never "the price went up so demand went
up"); a sunk cost is ignored, which is not "always quit"; opportunity
cost is the NEXT-BEST alternative, not the sum of alternatives; Concorde
failed on operating costs, not merely on expense; the house flip is
accounting-profit-positive and economic-profit-negative; Los Angeles was
BOTH curves shifting left; and the supermarket ruling was "some
substitution, not enough substitution". Upload each to its OWN notebook.

### Round 2 the same day — Nico's answers on the four flagged items

1. **Display 36** (the duplicated introduction slide before the module
   title) — keep as is.
2. **Tags.** Display 47 is now `Module 1 · Introduction` (it is the
   in-class copy of the course roadmap; display 7, inside Video 1, keeps
   `Module 1 · Course Roadmap`). The whole applications block — displays
   **51–65** — moved to the new four-level tag: `Module 1 · In Class ·
   Examples · Markets` (51–54) and `Module 1 · In Class · Examples ·
   Supply and Demand` (55–65). Nico asked for "47–64"; the block was
   taken through 65 so the two-slide Los Angeles mini-case does not end
   up with two different tags. Done in the two tag constants
   (`TAG_MARKETS`, `TAG_SD`), never per call site. **New rule added to
   `Teaching/CLAUDE.md`** ("In-class examples of video material get a
   four-level tag") so this is the convention wherever slides that apply
   taped material are moved into an in-class deck.
3. **Display 86** — the 9 pt bar labels stay. **New exception added to
   `Teaching/CLAUDE.md`** under the box/callout text floor: a label
   written INSIDE a narrow object (a thin bar, a slim column, a narrow
   table cell) is exempt from the 18/16 pt floors; fit the text to the
   object. The exception covers labels inside constrained objects only —
   not crowded slides.
4. **Poll Break badge** — the CLAUDE.md rule was rewritten as a complete,
   self-contained spec (position, size, slant, label box, font, fill,
   shadow, grouping, the parametric `custGeom` path, and the
   draw-it-LAST / covers-the-footer-rule requirement), so a build script
   can generate the badge from the rule alone. Module 1's three badges
   (55, 58, 71) were then retrofitted to it: `_add_pollbreak_badge` in
   `_build_Module1.py` now BUILDS the badge from `POLLBREAK_XY` /
   `POLLBREAK_WH` instead of injecting `_handoff_pollbreak.xml`. The
   badge moved from 9.45 / 6.77, 3.607 × 0.590 to **10.238 / 6.769,
   2.950 × 0.533** — the same geometry as Module 3's 22 badges.
   `_handoff_pollbreak.xml` is now unused by the pipeline (kept as the
   provenance of the design).

Re-verified after both rounds: `_diff2.py` shows only the intended
changes; `_verify_anim.ps1` ALL CLICK COUNTS MATCH (63 animated slides);
slideshow probe on 1, 55, 56, 58, 59, 71 PASS.

**Backups:** deliberately NOT rolled a third time within the session. The
two kept predecessors are the more useful pair — `_t-2` is Nico's own
hand-edited 95-slide deck of 17:08 and `_t-1` is the first rebuild after
the port. Rolling again would have dropped his hand-edited original for a
build that differs only in tags.

### Podcasts — one addition
Both source docs (and the standing rules in `Teaching/CLAUDE.md`) now ask
the hosts to **keep the language measured**: "massive", "huge",
"incredible" and the like only occasionally, not every other sentence.
Added to the in-doc host instructions AND the paste-ready Audio Overview
prompt.

**How to produce the audio (it is Google NotebookLM, not a Microsoft
tool):** notebooklm.google.com → **Create new** notebook → **Add source**
→ upload `Podcast Module 1 -- Intro.md` → **Studio** panel → **Audio
Overview** → **Customize**, paste the prompt block from the top of that
file → **Generate**. Wait a few minutes, play it back, then download.
Repeat in a **separate** notebook for the Wrap-up file — NotebookLM
blends every source in a notebook into one audio, so the two episodes
must never share a notebook. There is no editor for the finished audio:
to change an episode, edit the source doc and regenerate.


### Round 3 the same day — retag under the revised top-bar rule

The `Teaching/CLAUDE.md` tag rule changed after round 2 (the video number
now sits in the MIDDLE level of a taped module's tag, agenda slides read
"… · Agenda", the summary closer is exempt, and BACKUP slides get no
tag). Module 1 was re-tagged against it, using `Module 3 - Revised.pptx`
as the reference implementation.

**New machinery.** `_build_Module1.py` gained `_m1_top_bar_tags()` — ONE
table, display-keyed — and `apply_top_bar_tags(prs)`, a deck-wide pass run
at the end of `build()`. The slide builders still draw whatever tag they
were written with; the pass is now the single source of truth and rewrites
them. It is loud on drift: a top bar the table does not cover, or a table
entry with no bar to write to, is reported (the six spliced PollEverywhere
displays are excluded, since their stubs are thrown away).

**34 tags rewritten:**

| Displays | was | now |
|---|---|---|
| 2, 3, 6 | `Module 1 · Introduction` | `Module 1 · Video 1 · Introduction` |
| 4, 5 | `Module 1 · Economic Models` | `Module 1 · Video 1 · Economic Models` |
| 7 | `Module 1 · Course Roadmap` | `Module 1 · Video 1 · Course Roadmap` |
| 9 | `Module 1 · Video 1 · Introduction` | `Module 1 · Video 1 · Agenda` |
| 11 | `Module 1 · Video 2 · Markets` | `Module 1 · Video 2 · Agenda` |
| 18 | `Module 1 · Video 3 · Demand and Supply` | `Module 1 · Video 3 · Agenda` |
| 28 | `Module 1 · Video 4 · Market Equilibrium` | `Module 1 · Video 4 · Agenda` |
| 49, 66, 79, 84 | `Module 1 · Outline` / `… · In Class · <section>` | `Module 1 · Agenda` |
| 67–78 | `Module 1 · In Class · Opportunity Costs` | `Module 1 · Economic Costs Include Opportunity Costs` |
| 80–83 | `Module 1 · In Class · Sunk Costs` | `Module 1 · Ignore Sunk Costs` |
| 85–87 | `Module 1 · In Class · Cost-Benefit and Marginal Analysis` | `Module 1 · Use Cost-Benefit and Marginal Analysis` |
| 88 | `Module 1 · Wrap-Up` | `Module 1 · Summary` |

The three in-class agenda items are now the outline item titles, taken
from `M1_OUTLINE` through `_title_case()` rather than retyped, so renaming
an outline item renames the tags with it. The old three-level
`Module · Part · Section` form is gone from the deck.

**Unchanged, deliberately:**
- 12–16, 19–26, 29–34 already had the video number in the middle.
- 51–65 keep the four-level `Module 1 · In Class · Examples · <topic>`.
- 36, 38–41, 44–48 are front matter sitting OUTSIDE every video block (it
  lives in the in-class part in this deck), so it keeps its own two-level
  tag — the rule's "front matter takes the introduction video's number"
  applies only when that matter is inside the video block, as in Module 3.
- 89 ("Next Steps") keeps `Module 1 · Wrap-Up`: it is post-work
  logistics, not the summary closer.
- 1, 8, 10, 17, 27, 35, 37, 50, 90 and the six poll slides have no top bar
  and are untouched.

**Open question — the five backup slides (91–95).** The rule lists BACKUP
slides among those that "never get a tag", but its justification ("a
backup slide is a full-bleed figure with a caption") describes Module 3's
backups, which have no top bar at all and never did. Module 1's five are
ordinary content slides with a navy bar, an action title and bullets, so
obeying the rule literally means STRIPPING the top bar from all five. That
was left alone pending Nico's word; they still read `Module 1 · Backup`.

**Verification:** `_diff2.py` vs. the previous canonical shows the top-bar
text shape as the only difference on any slide; `_verify_anim.ps1` ALL
CLICK COUNTS MATCH (63 animated slides); slideshow probe on 1, 9, 42, 67,
88, 95 PASS; top bars on 9 / 67 / 88 rendered and checked (the longest tag
fits the bar comfortably).

**Backups:** again not rolled — `_t-2` remains Nico's hand-edited 95-slide
deck of 17:08 and `_t-1` the first rebuild after the port. Every
intermediate since is script-reproducible; his hand-edited original is not.


### Round 4 the same day — the backup-tag rule was wrong, not the deck

Nico's call on the open question from round 3: **a genuine backup slide
DOES carry `Module N · Backup`**, and Module 1's five (91–95) are the
model. Module 3's two tagless backups are the exception, not the pattern —
there he enlarged the picture to fill the canvas and pulled the text on top
of it, so there is no top bar to put a tag in. **No change to Module 1's
deck**; the fix was to `Teaching/CLAUDE.md`:

- BACKUP slides were REMOVED from the "never gets a tag" list, and a new
  sub-rule says every backup slide carries `Module N · Backup` with
  ordinary chrome, the same tag on all of them (it does not name the topic
  being backed up). Module 1 displays 91–95 are cited as the model. The
  full-bleed backup is written up as the single exception, with an explicit
  "do not strip the top bar off an ordinary backup slide to reach it".
- The `In Class · Examples` clause was sharpened so the TEST is what a
  slide is FOR, not where it sits: it applies to slides that apply taped
  material and are kept back to be shown in class. Where they are parked is
  a per-deck layout choice — an "Applications" divider mid-deck (Module 1,
  displays 51–65) or an examples appendix at the end (Module 3) — and a
  module with no appendix simply has none, so the appendix sub-rules do not
  apply to it. Added alongside: `Examples` never appears inside a video
  block; a slide between two title cards carries `Module N · Video k ·
  <topic>` even when its content is a worked example.
- `_build_Module1.py`'s tag-pass comment gained a line on the backup case,
  so the script and the rule say the same thing. Comment only — the deck
  was not rebuilt.

Audited the shipped deck against the finalized rule: 15 slides with no top
bar (title slides, video title cards, three dividers, six polls), the four
video blocks tagged `Module 1 · Video k · <topic>` with `… · Agenda` on
each video's outline slide, `Module 1 · Agenda` on the four in-class agenda
slides, the two four-level Examples tags across displays 51–65, the three
in-class agenda items under their outline titles, `Module 1 · Summary`,
`Module 1 · Wrap-Up`, and `Module 1 · Backup` five times. Conforms.

## 2026-08-24 — porting the polished "Videos Final" decks back in

**One-line summary.** Nico deleted the four old video decks, split polished
per-video decks into `Videos Final/`, and asked for every edit in them to be
carried back into `Module 1 - Revised.pptx` (still 101 slides). All 35 video
slides were mapped to their main-deck twins, diffed, and the differences
ported through the pipeline — build script, group pass, animation plans —
never by hand.

### How the diff was done (reusable)
Shape ids SURVIVE an extract-and-polish round trip, so the video slides
could be paired to the main deck exactly, by id. Three passes:
- `_vdiff.py` / `_vdiff_all.py` — id-keyed geometry + text + run-format +
  group-path diff, plus a click-by-click comparison of `<p:timing>`.
  `_diff_cross.py` is the earlier positional version (kept; weaker).
- `_rawdiff.py` — raw `spPr` / rels diff. This is what caught the two
  arrowheads on displays 83 / 86; a shape-level diff cannot see them.
- `_map_vids.py`, `_pair_ids.py`, `_vtext.py`, `_outline_probe.py`,
  `_shape_xml.py`, `_find_text.py`, `_vid_inv.py` — probes used along the way.
The port itself is `_port_video_edits.py`, `_port_video_groups.py`,
`_port_video_anim.py` — each replacement asserts its match count.

### Video → main-deck mapping
| Video deck | main displays |
|---|---|
| Video 1 – Introduction (11) | 67, 2, 9, 10, 11, 13, **17**, 1, 69, 95, 100 |
| Video 2 – Markets (7) | 70–76 |
| Video 3 – Demand and Supply (10) | 77–86 |
| Video 4 – Equilibrium (7) | 87–93 |
Video 1's agenda slide is main **17** (In-Class), NOT 68 — it matches 17's
tag and notes exactly, minus the backup pill. Main 68 was untouched.

### Adopted
- **Agenda shading (deck-wide).** On a section agenda the items NOT
  currently covered are dimmed — circle digit and item title both
  `#BFBFBF` (Nico's decks write `schemeClr bg1 lumMod 75%`); gold circle
  fill stays gold. The descriptive overview (18) and the Video-1 outline
  (69) keep every item navy. `DIM` added to the palette; `make_m1_outline`
  gained `lit = descriptions or i in hi`.
- Outline data: "the value of the **next-best** alternative you gave up"
  (one shared list → displays 18, 19, 43, 69).
- **2** new sub-bullet "Visiting positions at Harvard, University College
  London"; backup action button 4.555" → 4.765".
- **9** "as consumers, students" → "partners".
- **72** "Geographic boundaries"; "vs. gasoline retail vs. App purchases";
  "Simple test" bold; 18 pt above "Extent of market"; build 5 → 6 clicks
  (new PLANS key 64).
- **74** provenance line loses " — not leaked".
- **75** Covid bullet deleted; discussion badge grouped with its text
  (new `sdbadge:` names + `CHART_GROUPS[75]`); slide is now STATIC.
- **77 + 78–86** "Demand and Supply" — title slide AND `TAG_V3`.
- **83 / 86** the red movement arrows got their missing arrowheads
  (`_set_line_ends`; 83 = `headEnd`, 86 = `tailEnd` because it is flipV).
- **84** opening bullet + chips photo now static; 3 → 2 clicks.
- **91 / 92 / 93** two groups each (shifted curve + label + shift arrow;
  new-equilibrium guides + P1/Q1) via named shapes in `_v4_shift_chart`
  and `CHART_GROUPS`; builds recomposed to his beats. On 93 the *S* in the
  Note line is italic, "Price ($)" moved to (7.300, 2.020) to clear the
  two-word header, and "➜ Problem Set 1" is the final click.
- **100** caption 6.00" → 5.06"; gained a 2-click build — the ONLY backup
  slide that animates (`SKIP_STATIC = (SKIP_STATIC | {75}) - {100}`).

### NOT adopted, and why
- Page numbers (live fields), and PowerPoint's autofit re-fits of bullet
  boxes: those shrink the box around an UNCHANGED centre (1.600 + 5.350/2
  = 2.554 + 3.442/2 = 4.275), so they are save artifacts, not hand-edits.
  Worth remembering — they look like real moves in a naive diff.
- **73 / 74 (Tapestry)**: his video copies came out of `Module 1 - Example
  Candidates.pptx`, not the main deck — "Candidates" tag, "(for review)"
  footer, candidate research notes, no photo+caption group, and shape ids
  off by one. The main deck is ahead; only the "not leaked" deletion was
  taken.
- **91**: he nested the D’ label around an inner group; ported as ONE flat
  group of the same four shapes, matching how he grouped 92 and 93.
- **94** ("Effect of Shifts … in Isolation") is absent from his Video 4
  deck; read as a video-only cut and left in the main deck.
- Outline-slide run colors losing their explicit `srgbClr` on 71/78/88 —
  a copy-paste theme artifact, NOT the dimming (the dimming shows up as
  `schemeClr bg1 lumMod`, which an `srgbClr`-only probe misses entirely;
  this is how the shading was nearly missed on the first pass).

### Verification
- id-keyed diff vs. all 35 video slides: clean (leftovers are page
  numbers, autofit re-fits, and group NAMES — every group's bounding box
  matches his exactly).
- click structure matches his beat-for-beat on every changed slide.
- `_verify_anim.ps1` (expected counts updated: 72 → 6, 84 → 2, 91/92 → 2,
  75 removed, 100 → 2): ALL 67 animated slides match; deck opens in
  PowerPoint at 101 slides.
- `_slideshow_probe.ps1` on 1, 24, 72, 75, 91, 93, 100: PASS, live poll
  renders.
- Renders of all 17 changed slides eyeballed (`_chk_sheet_1/2.png`).

### Rule changes (Nico approved)
`Teaching/CLAUDE.md`, Module-Outline section: "no fading of the other
items" is GONE, replaced by the dimming rule above; the reference
implementation now points at `make_m1_outline` in `Module 1`, not
`make_m2_outline`.

### Toolchain notes
- **`_anim_config_m1.txt` IS STALE** — it was the one-off payload for
  `_splice_anim_config.py`. `_animate.py` has been the live config ever
  since. A STALE banner was added to the top of the .txt.
- `_animate.py` PLANS keys are in the 84-deck numbering; display =
  `_m1_shift_key3(_m1_shift_key2(_m1_shift_key(k)))`. Post-shift overrides
  (`PLANS[73]`, `PLANS[74]`, and now the 75/100 block) are keyed by
  DISPLAY number instead.
- Display 100's caption, now close to its picture, is picked up by group
  rule 3 — the main deck has a pic+caption group there that his slide does
  not. Consistent with the standing rule; flagged to him.

### Open / next
- **Module 2's outline slides still lack the dimming** — its build script
  predates the rule. Same one-line guard in `make_m2_outline`, but it needs
  its own backup roll, rebuild and animation re-verify. Awaiting his word.
- Everything in the PENDING block above is still open.


## 2026-08-23 — link symbols, hand-edit ports, subscripts, outlines, notes

**One-line summary.** A long iterative pass over `Module 1 - Revised.pptx`
(still 99 slides): replaced the gold ▶ link glyphs with PowerPoint action
buttons, ported ~15 rounds of Nico's hand-edits, darkened the deck's green
everywhere, converted all nine outline slides to the Module 2
numbered-circle format, added a deck-wide symbol-subscript pass, replaced
the COVID/flour slide with an AI-and-chips example, and gave every one of
the 99 slides speaker notes.

### Deck changes, in the order they were asked for
- **Link symbols.** The gold ▶ / ➜ glyphs are gone deck-wide. Backup jumps
  use `actionButtonEnd` (navy face, white glyph, 0.434 × 0.210" — 30%
  smaller than the first cut); external links use the same family keyed to
  what they open: `actionButtonSound` (podcast, display 12),
  `actionButtonDocument` (Economist article, 94), `actionButtonMovie`
  (econimate video, 94), all navy. The back button reverted to the original
  plain navy "← Back" pill after a detour — see the CLAUDE.md rule.
- **Invisible click overlays removed.** Transparent rects over the bullet
  boxes on displays 2 and 9 were why Nico could not select the text; the
  action button is now the click target itself.
- **Hand-edits ported** (all with dated comments in the build script):
  slide 1 comic raised to y 162547 EMU; display 11 third map enlarged;
  13 fox/hedgehog shrunk and stacked; 15 Einstein sub-bullet deleted and
  the block re-centred; 23 both figures raised clear of the footer; 26
  arrow + D′ label moved; 78 definition callout moved up; 79 cones panel
  grouped + his 3-click build; 81 D′ label / ii) arrow / ii) label moved,
  a new horizontal dashed segment, two groups, his 4-click build; 84 the
  P1↔P2 / Q1↔Q2 relabelling (the movement along S starts at the LOWER
  price — the build had it inverted), i)/ii) repositioned, a new dashed
  segment, four groups, his 2-click build; 89/90 arrow and label positions;
  92 reworded and hidden.
- **Slide 36 (avocados).** Rebuilt to Nico's original 10-click
  choreography from `Module 1 - In Class.pptx` slide 30, with his grouping
  of curves and labels. Fixed a real defect: the two shift arrows were on
  the WRONG beats (S→S1 fired with the demand shift, S→S2 with S1).
- **Green.** `#00B050` is gone from both decks; `GREEN_DK = #007A33` is the
  only green left. `GREEN_BR` and `GREEN_MB` were retired outright.
- **Symbols.** `apply_symbol_subscripts()` is a deck-wide build pass that
  splits any P/Q/D/S symbol followed by an index into an italic base run
  plus a true subscript run (66 paragraphs). Keyed to those four letters
  so ordinary text is untouched.
- **Slide 82** replaces the COVID/flour example: AI and the demand for
  chips, Nico's `AI_Accelerator_Chips` image, D → D′ outward shift, no
  supply curve (Video 3 has not reached equilibrium yet) and no Q2 guides
  (they would suggest the price stays constant).
- **Outline slides.** All nine now use the Module 2 numbered-circle format
  via a `make_m1_outline` copied from `make_m2_outline`, over a new
  6-item `M1_OUTLINE`. Item rows are pixel-identical across slides
  (rows from y 1.635", pitch 0.910"); bands all (0.900", 12.150 × 0.900").
- **Speaker notes.** `FILL_NOTES` (58 entries) + `apply_fill_notes()` fills
  every slide that had none. It never overwrites existing notes, which is
  what keeps the source-ported notes and — critically — the eight
  PollEverywhere payload notes intact. Verified byte-for-byte.

### Toolchain changes worth knowing
- **Shape names drive grouping and animation now.** `_sd_chart` and the
  hand-built charts name their shapes (`sdcurve:D`, `sdguide:h:Q3`,
  `sdarrow:ii`, `sdpic:chips`, …); `_group_pass.py` rule 5 pairs them by
  name from `CHART_GROUPS` and names the group `sdgroup:<key>`;
  `_animate.py` gained an `n:<name>` selector. This replaced a
  nearest-connector heuristic that silently grouped slide 36's "D" label
  with the *S2* curve — names removed that whole class of error.
- `_group_pass.py` also gained rule 4 (label + link button) and a
  width guard so outline bands (12.15") are not treated as callouts,
  while the 10.5" "Important" box on display 92 still groups.
- `_animate.py`: `t:`/`pr:` selectors now match a concatenated-run variant
  as well, because the subscript split turned "P0" into "P 0".
- Path-independence fixes: `_diff_slides.py`, `_verify_anim.ps1`,
  `_slideshow_probe.ps1` and `_export_probe.ps1` all take a deck argument
  and use `$PSScriptRoot`.
- New audit helpers: `_diff_all.py` (full-deck member-level hand-edit
  diff), `_check_jumps.py`, `_scan_glyphs.py`, `_audit_overlays.py`,
  `_audit_notes.py`, `_extract_timing.py` (pull a slide's click structure
  out of any deck — this is how his choreography was adopted),
  `_shape_idx.py`, `_dump_edits.py`, `_sheet_probe.py`, `_crop_probe.py`.

### Rules added to `Teaching/CLAUDE.md` (all at his request)
1. Overlay coverage — two figures should not overlap at all where the slide
   has room; where a build paints successive versions of the SAME figure,
   the later one must fully contain the earlier.
2. A backup link always sits in the lower-right corner; a podcast/article
   link goes wherever it fits best.
3. Indexed symbols get a real subscript, italic letter.

### Late evening — Tapestry insert (deck 99 -> 101) + slide 12
- Nico copied the two Tapestry-Capri slides out of `Module 1 - Example
  Candidates.pptx` (slides 2-3) into the deck at **displays 73-74** and
  animated them. Ported into `_build_Module1.py` as
  `slide_tapestry_case` / `slide_tapestry_evidence` (from
  `exp_tapestry_case` / `exp_tapestry_evidence`) rather than spliced, so
  build.py stays the source of truth and future style passes reach them.
  His edits adopted: the gold "the FTC's market ..." line deleted, the
  quote box grouped with its text, and his choreography (73: three
  clicks on the setup bullets; 74: seven clicks - ladder, share cards
  with the provenance line riding the 77% card, internal quote, court
  decision). MY changes on top: the candidates chrome swapped for
  `TAG_V2` + the normal footer, and the two photos grouped with their
  shared caption.
- **Everything from display 73 on shifted +2**, via `_m1_shift_key3` in
  `_animate.py` and `_m1_disp_shift` in the build script (FILL_NOTES
  keys). Backup jump targets are now 96 / 97 / 100 / 101. Deck is 101
  slides with 67 animated. `_group_pass.py`'s outline-band guard is now
  scoped to `OUTLINE_SLIDES` by number instead of a width threshold -
  the old 11.5" rule would have blocked the 12.08" Tapestry quote box
  from grouping.
- **Slide 12**: Nico dropped `pic:0` from the first animation beat, so
  the Homo-Economicus icon cluster is visible from the start. Adopted in
  `PLANS[10]`.
- **Slide 12 title bug, resolved.** After his edit the only effect left
  on beat 1 targeted the slide TITLE, not the panel heading: the
  `t:Homo Economicus` selector matches the action title first in
  document order, so the slide opened untitled. Nico: "don't animate the
  title on slide 12" - the beat was removed outright, which also matches
  his own edit (he had already made the row-1 icon static, and the row-1
  text always was). Slide 12 is now 2 clicks: Real Human row, then the
  podcast label; the title and the whole Homo-Economicus row are static.
  Worth remembering: a `t:` selector in a custom plan BYPASSES
  `is_chrome`, so a prefix that also matches the action title will
  silently animate chrome.

### Verification, every round
`_verify_anim.ps1` (COM click-count check, all 65 animated slides) +
`_check_jumps.py` + a full-screen `_slideshow_probe.ps1` run including a
live PollEv slide. Hand-edits were always captured with `_diff_all.py`
against a side-path build BEFORE rebuilding.

## 2026-08-22 — backup section + missing poll slides (deck 87 → 99)

Nico uploaded **"Module 1 - In Class with Solutions.pptx"** (68 slides,
his FEMBA variant) and asked to add everything the rebuild had missed.
Findings: all Solution slides were already adopted (AC #26, Swift #30,
flip-house #51) and the Kleven child-penalty chart too (#54). Missing
were (a) FIVE PollEv slides — a new "Econ & Coffee weekend slot" poll
pair after Office Hours, plus a results-view slide per existing poll
(each its own __PE_POLL_EMBED_ID) — and (b) the 7-slide BACKUP section
with its jump links. All added; pipeline rerun; verified.

New display map (old-87 → new-99): 1–6 same; 7–22 → +2; 23–25 → +3;
26–45 → +4; 46–87 → +5. New slides: 7/8 (Econ&Coffee poll pair),
25/29/50 (results views for AC / diamonds / flip polls), 93–99 =
BACKUP divider, National Leaders (Econometrica 2025), Money-Buy-
Happiness (Easterlin), Stevenson-Wolfers 2008, Anderson Faculty
(HIDDEN, as in source), Portland Street windows tax, Lufthansa fares.
Links: 2→94, 9→98, 12→95 (pointer pill), 17→99 (pointer pill); back
pills 94→2, 96→12, 97→12, 98→9, 99→17 (95 has none, flows to 96, as
in the source). Slide-6 hand-edit ported (bullets_top 2.21"/4.13" with
dated comment); full-deck geometry diff showed no other hand-edits;
`_test` deck deleted. Backups rolled to `_t-1`/`_t-2`.

Implementation notes:
- `_splice_media.py` SPLICE_MAP entries are now (source deck, display):
  "IC" = In Class, "WS" = In Class with Solutions (also 4:3, same
  +1.667" shift). `_group_pass.py` SPLICED and `_animate.py` skips /
  PLANS renumbered via `_m1_shift_key2`; `_verify_anim.ps1` table
  renumbered (still 65 animated slides — ALL COUNTS MATCH; slideshow
  probe incl. all-new polls + backups PASSED).
- **Hyperlinked text runs render UNDERLINED on this machine regardless
  of u="none"** (verified against a native PowerPoint save — even
  PowerPoint's own no-underline hyperlink run renders underlined).
  Slide-jump affordances therefore use SHAPES: invisible 100%-
  transparent-fill overlay rects on slides 2/9 (over the existing
  gold ▶ lines), `_add_ps_pointer` pills on 12/17, navy back pills
  (`_add_back_pill`) on the backup slides.
- Backup screenshots with baked-in red-circle annotations were cropped
  from 2400px slideshow exports of the WS deck
  (`_source_images/ws66_goodlife_crop.png`, `ws68_fares_*_crop.png`);
  plain images extracted verbatim (`ws63_*`, `ws64_*`, `ws65_*`,
  `ws67_*`).

Flags for Nico (not changed, awaiting word):
- WS slide 26 (heatwaves SOLUTION) also shows a flour-shortage
  clipping + "other examples of right-shifts" bullets that our native
  AC-solution slide (#26) doesn't have (flour appears in Video 3,
  display 82). Add to #26?
- WS backup "National Leaders" carried stale speaker notes (hedgehog
  text) — wrote fresh notes instead; Portland notes ported verbatim.
- WS deck uses the FEMBA TA email; deck keeps TA405.EMBA2@gmail.com.
- Spliced poll slides carry the source's static page-number text
  (e.g. "46" on display 50) — cosmetic, pre-existing behavior.

## 2026-08-20 (round 3) — "Module 1 - Example Candidates.pptx" (14 slides)

Nico asked for a thorough, careful web search for recent (2023–2026)
real-world, MBA-compatible examples for every Module 1 concept, delivered
as a separate review deck. Process: 5 parallel paper-writing-agent
research runs (market definition, S&D shocks, opportunity cost, sunk
costs, marginal analysis + fairness), each verifying facts against
primary/tier-1 sources and flagging confirmed vs. reported figures.
Deliverable: `Module 1 - Example Candidates.pptx` — cover + 11 candidate
slides + 2 "bench" slides (runners-up, one line each). Build script:
`_build_M1_candidates.py` (imports the `_build_Module1.py` helper
layer). Each candidate slide: concept tag, fact bullets, cream
teaching-angle card w/ proposed visual, gold discussion prompt, source
line; full URLs + verification flags in the speaker notes.

Candidates: Tapestry/Capri "accessible luxury" · Kroger/Albertsons
"Costco run" · Netflix+WBD (3 market definitions — updates existing
Netflix slides) · Eggs 2024–26 (BLS-verified) · DRAM/AI memory · AI
talent war ($100M implicit cost) · Return-to-office (AEA-published 72
min/day) · Apple Car + GM Cruise (SEC-confirmed exit math) · Meta
Reality Labs ($80B debate) · United/Delta marginal-flight cuts (2026
fuel shock) · LA-fires 10% rent cap (§396, local). Bench: FTC v. Meta,
Google/AI-chatbot market, FTC v. Amazon, beef, coffee, cocoa, GLP-1,
Berkshire cash, hyperscaler capex, NIL, CA HSR, Ørsted, Sony Concord,
Google demand response, Wendy's, egg rationing/DOJ.

Standing rule respected: press-reported figures are marked "reported"
on-slide; speaker notes name what must be re-verified before a
candidate graduates into the main deck. Awaiting Nico's picks.

**Round 3e (same day):** Deck-wide ≥18pt font pass on the candidates
deck (Nico: "font inside text boxes at least 18pt"): teaching-angle
cards, discussion lines, quote boxes, diagram boxes, timeline labels,
callouts all raised to 18+; chart labels to 16; photo/source captions
stay at caption size. Resolution lines on the Tapestry, Costco, and
Netflix-chart slides promoted to gold takeaway bars (19pt bold navy);
their on-slide source lines moved to speaker notes for space. Bench
split into three slides to fit 18pt. Deck now **21 slides**. Also:
Costco-run title reworded to "Everyday Shopping" (Nico), Kroger setup
wording varied ("Once again, everything would hinge on how you define
the market").

**Round 3d (same day):** Candidates deck now **20 slides**. (a) DRAM
case expanded to 3 slides: setup (datacenter photo, chronology-first
flag on wafer allocation), two-panel native S/D analysis (HBM: D shifts
right against steep supply; consumer DRAM: S shifts left as wafers
reallocate — gold arrow between panels), resolution (H100 + SK Hynix
DDR5 photos; magnitudes marked reported/TrendForce). (b) Tapestry and
Kroger setup slides restructured per Nico: deal first, then "market
definition would turn out to be crucial" with the two sides' market-
extent arguments as sub-points; photos now stacked vertically on the
right, text on the left two-thirds, fonts 24/22. (c) Share-figure
provenance verified from the Clifford Chance briefing (read directly):
58.7% = FTC's expert from largely third-party data; 77%/83% = Capri's/
Tapestry's internal ordinary-course documents produced in the merger
investigation (not leaked); "accessible luxury" was the firms' OWN term
from SEC filings and investor decks until the FTC sued (then
"expressive luxury") — now on the evidence slide + notes. New Commons
photos: web_dram, web_datacenter, web_h100.

**Round 3c (same day):** New standing preference from Nico, added to
`Teaching\CLAUDE.md` ("Case buildup: chronology first, resolution
second"): a two-slide mini-case builds the situation on slide 1 and
ends by flagging the crucial feature ("the definition of the market
would turn out to be crucial") WITHOUT revealing the outcome; slide 2
shows the resolution as the final beat/click. Applied to both the
Tapestry–Capri and Kroger–Albertsons pairs in the candidates deck
(court decisions moved to the second slide of each pair).

**Round 3b (same day):** Nico asked to expand candidates 1–3 + United
into 1–2 slides each with illustrative pictures. Deck now **18 slides**:
Tapestry–Capri (case + price-ladder/evidence slide with the 59/77/83%
share cards and the internal-message quote), Kroger–Albertsons (case +
"Costco run" in/out-of-market diagram with the Nelson quote),
Netflix–WBD (deal-saga timeline + native Nielsen TV-time bar chart vs.
the red SVOD callout), United (case with Kirby-quote callout + native
MB=MC chart with the crossing shifting left). Photos fetched from
Wikimedia Commons via `_fetch_web_images.py` (BUILD INPUT; images in
`_source_images/web_*.jpg`: Coach + Michael Kors stores, Kroger,
Albertsons Dallas, Costco, United 787, WB water tower) — all reviewed
before use; "Photos: Wikimedia Commons" caption lines on-slide. The
remaining 7 candidates + 2 bench slides unchanged.

## 2026-08-20 (round 2) — comic back + 2 more MW applications: 87 slides

Nico approved: (1) title-slide comic reintroduced; (2) NEW #23 "Shifts
of the Demand Curve for AC" (MW #51 solution after the AC poll, native
D→D′ chart); (3) NEW #37–38 copper mini-case (MW #65–66, two-stage
quantity/price figure + native both-shift-right chart, P1 = P0).
Old #23+ shifted +1, old #36+ shifted +3; polls now at 22/25/45. Page
numbers renumbered via `_renumber.py` (descending literal replacement);
`_animate.py` keeps pre-insert PLANS keys shifted by `_m1_shift_key`.
Full pipeline re-run; renders of 1/23/37/38 checked; click structure
re-verified (ALL 65 animated slides match).

## 2026-08-20 — Full rebuild into "Module 1 - Revised.pptx"

**One-line summary.** Built `Module 1 - Revised.pptx` (**84 → 87 slides**,
16:9, new 405 format) from Nico's 53-slide 4:3 In-Class deck plus the
four video decks (25 slides) appended at the end, adopting 5 approved
MW (Melanie Wasserman) items; 3 PollEv slides spliced live; grouped;
fade builds applied; click structure + slideshow probe verified.

### Structure
- Slides 1–58: In-Class part (front matter, models/philosophy, markets
  + S/D mini-cases, opportunity costs, sunk costs, CBA, summary).
- Slides 59–84: Videos 1–4, each with its own deck-format title slide
  (Nico will eventually split them back out into separate video decks).
- Outline slides keep Nico's order: videos listed first on the in-class
  outline (slide 16), in-class first on the video outlines (as in the
  video sources).

### Pipeline (rerunnable, Module 7/2 pattern — 4 steps)
```
python _build_Module1.py           # phase-1: all 81 scripted slides + 3 stubs
python _splice_media.py            # 3 PollEv slides verbatim (w/ notes+tags)
python _group_pass.py              # 7 groups (callouts, table shades)
python _animate.py all apply       # fade builds per per-slide plans
```
Helpers (`_build_template_samples.py`, `_animate.py`, `_group_pass.py`,
`_splice_media.py`, `_handoff_pollbreak.xml`) carried from Module 2.
`_animate.py` got one engine fix: shape text is whitespace-normalized so
`t:`/`pr:` prefixes can span run boundaries. `_group_pass.py` got a
rule-1 height cap (≤2.5") so the Homo-Economicus cream panel isn't
falsely paired with one of its text blocks.

### Decisions locked (2026-08-20, Nico)
- MW imports 1–5 adopted: flip-a-house Solution (new #43), shift-
  combination table (new #84), LA real-estate mini-case (new #34–35),
  Next Steps (new #58), Swiftonomics diamonds refresh (new #23 + #25);
  optional items 6–9 declined.
- Slide 4: "Fall 2025" → "Fall 2026" Achieve site.
- Slide 5: exam periods → [DATE] placeholders.
- Old slide 25's lithium-article note kept for now (new #26).
- Problem-Set pointers generic ("Problem Set 1", no exercise numbers).
- Nico's teaching order confirmed: videos watched FIRST, in-class
  applications stay where they are; MW imports slotted as extra
  applications only.

### PollEv caveat (IMPORTANT, Nico action)
New slide 24's spliced poll still asks "How does the DECLINE in
engagements affect the demand for diamonds?" — with the Swift example
the answer flips (demand shifts RIGHT). Reword the activity in the
PollEverywhere account (URL/embed stays valid); the static screenshot on
the slide will still show the old wording.

### Verification done
- All 84 slides render-checked via COM PNG exports (2 rounds; 6 layout
  fixes applied: s12 overlay, avocado TIFF alpha, s56/s72/s75 overlaps,
  s81 label, s83 P-label separation).
- Deck opens clean in PowerPoint (84 slides).
- Animation click structure verified via COM MainSequence TriggerType:
  ALL 62 animated slides match the plan.
- Full-screen slideshow probe (screenClass PrintWindow captures on
  slides 1, 22, 24, 42, 55, 84): PASS — all 3 live polls render their
  activities in the real slideshow; no "failed to open" banner.

### Content flags for Nico (reported in chat)
1. Title slide: the comic strip from the old title slide was dropped
   (new-format title slides are clean); the UCLA logo likewise.
2. Old #15's stray Lufthansa logo not carried into the new roadmap;
   roadmap wording standardized to the M3/M2 format ("1. Basic
   Principles and Economic Way of Thinking"); video-1 agenda's
   "2. Buyers, Value, and Demand" wording also standardized.
3. New #21 got an action title ("How Can Heatwaves Affect the Demand
   for ACs?") — the source slide had no title, only the question.
4. New #47 title rendered as "Similar Figures for the US, Estimated in
   2022" (source title had a line-break artifact).
5. Video title slides read "Module 1 – Video n" (source said "Week 1").
6. Fruit table (new #38) is now a native table — it reveals as one
   block, not cell-by-cell like the old shape-built version.
7. Next Steps (#58) keeps MW's two pre-class sub-bullets ("Read news
   article…", "Take survey…") — cut if they don't fit Nico's flow.
8. Exercise diagram (new #55) rebuilt natively with clean MB/MC values
   (blue net-benefit / red MC, indifferent at hour 4, STOP at hour 5).

### Suggested additional MW applications (awaiting Nico)
- **AC-heatwave Solution slide** (MW #50–51): after the AC poll (new
  #22) the deck jumps straight to Swiftonomics; MW closes the example
  with "demand shifts right". Could add a native D→D' solution slide.
- **Copper since 1880** (MW #65–66): both curves shift right → quantity
  ×100 at flat price; completes the shift taxonomy next to tea/avocado/
  LA and the #84 table.

### Pending / next steps
- Nico's eyeball pass of the deck + slideshow.
- PollEv rewording (see caveat above).
- Speaker notes: substantive originals preserved verbatim; MW-adopted
  and NEW slides carry drafted 2–4-sentence notes. No teleprompter pass
  requested yet for this deck.
- Not committed to git yet (Nico confirms at session end).

### Gotchas learned this session
- PowerShell COM: `New-Object -ComObject PowerPoint.Application`
  attaches to the RUNNING instance — never call `$pp.Quit()` when the
  user has decks open (it killed his PowerPoint once; only close
  presentations you opened, read-only).
- `_animate.py` joins runs with spaces when collecting shape text —
  without whitespace normalization, `pr:`/`t:` prefixes that cross run
  boundaries never match.
- Converting TIFF→PNG with `.convert('RGB')` flattens alpha to BLACK;
  composite on white first.
