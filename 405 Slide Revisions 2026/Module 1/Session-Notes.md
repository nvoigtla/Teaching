# Session Notes — Module 1 (combined In-Class + Videos deck)

## PENDING (updated 2026-09-23)

**STANDING RULES for Module 1 (2026-09-23, Nico)**
- **`Module 1 - Revised.pptx` is the SOURCE OF TRUTH.** The build pipeline
  is retired and refuses to write it (`_m1_frozen.py`). Edit the .pptx.
- **Never delete a slide from it without asking.** Only update the slides
  that changed in the in-class deck.
- `Module 1 - In Class.pptx` is FROZEN — adopted exactly as Nico left it.

C1. *(resolved 2026-09-23)* `Module 1 - Revised.pptx` IS the source of
    truth. The pipeline is retired and guarded — see round 5 below. Edit
    the .pptx directly from here on.
C2. The backup slide's Back button returns to RV 45 (the in-class
    roadmap), while RV 7 (the Video-1 roadmap) also links to it. Fine, or
    re-point?
C3. *(resolved 2026-09-23)* Nico removed the stranded effect himself, in
    REVISED: slide 42 is now 0 clicks / 0 effects, honestly static, with
    all four illustrations showing as the slide opens. **In-Class slide 8
    still carries the orphan** — the two decks now differ there on purpose,
    because Revised is the source of truth and In-Class is a frozen
    archive. `_xml_diff.py` reports that one pair, as expected; it is the
    only content divergence between the decks (slide 92's page number
    aside).

**Checking this deck:** `_xml_diff.py` is the conclusive In-Class ↔ Revised
check (normalized raw XML). `_style_diff.py` covers geometry/fills/effects.
`_ic_vs_rev.py` covers text and layout. Use the first one before claiming
the two decks agree — the other two each have blind spots.

*(All six merge questions below were answered on 2026-09-23 and are done —
kept for the record.)*
B1. The three live poll slides – port In-Class's full-bleed / POLL-pill
    captures over Revised's inset ones? They also open the activity live
    rather than on the closed chart.
B2. Revised keeps RV 72, the flip-house poll's question view; In-Class
    collapsed that poll to one slide. Drop it from Revised?
B3. Revised keeps RV 42/43, the Econ & Coffee poll pair, dropped from
    In-Class. Drop from Revised too?
B4. Revised keeps RV 94 "People Respond to Incentives", deleted from
    In-Class – but **RV 3 jump-links to it**. Keep, or delete and rewire?
B5. Add In-Class slide 7, "Recall from Video 1: Economists as Hedgehogs",
    to Revised's in-class part?
B6. The roadmap / "Can These Prices Be Optimal?" pair (IC 11 → RV 47,
    IC 58 → RV 95): where should the jump pill live in Revised, and where
    should the Back button point?

## PENDING (older, updated 2026-09-22)

Open items on **Module 1 - In Class.pptx** (now 59 slides):
A2. **Slide 50's five olive "Total Benefit of Hour N" boxes** sit BEHIND
   the bars and show only as a hairline; clicks 1–5 reveal them invisibly.
   Two are mislabelled "Hour 1", and "Net Loss of Hour 3" sits on the
   Hour-5 bar. See the 2026-09-23 (round 2) entry.
A. *(resolved 2026-09-23)* The empty poll slots: Nico deleted the
   flip-house placeholder and collapsed that poll to one live view. Three
   live activities remain (21, 24, 37). Original question kept below.
A-orig. **Do the three new empty poll slots stay?** Each of the three
   poll set-ups is now followed by an empty POLL placeholder AND by
   the LIVE PollEverywhere activity that was already there – see the
   table in the 2026-09-22 entry. Slots 21 and 25 look right (the
   question-view slides for those two polls were deleted on
   2026-08-27, leaving only the results view); slot 39 is doubtful,
   because flip-a-house already carries both views. Nothing was
   deleted; awaiting Nico's call.

Open items on **Module 1 - Revised.pptx** (95 slides):
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

## 2026-09-23 (round 3) – In-Class merged INTO "Module 1 - Revised"

**One-line summary.** Nico froze `Module 1 - In Class.pptx` ("no more
changes") and asked for Revised to be brought up to it wherever anything
differs, the in-class version being dominant. 57 slide pairs were compared;
35 were already identical, **17 were ported**, and **5 are held back pending
his answers** (below).

### The mapping – it cannot be found by similarity alone
`_ic_vs_rev.py` (new, read-only) uses an EXPLICIT In-Class -> Revised map,
because a text-similarity matcher pairs in-class slides with their
near-identical twins in Revised's VIDEO block: Netflix, the hedgehogs slide
and all the outline slides appear twice in Revised. Every pairing is
re-scored and anything under 0.50 is flagged, so a wrong pairing cannot pass
silently. The In-Class deck (58 slides) maps onto Revised's in-class part
(35-95); the two decks order their front matter differently (IC puts the
"In-Class Part" divider at 14, Revised at 35) but the content pairs 1:1.

### Ported – the slide part is copied WHOLESALE
`_port_ic_to_rev.py` (kept) copies the slide XML, its notes, its images and
its layout rel, rather than re-deriving the differences. That is the only
way to carry the hand-built choreography across: slide 50's 23-click
exercise chart and slide 33's 15-shape table would not survive being rebuilt
from the generic rules. Teaching CLAUDE.md is explicit – "never regenerate
animations from the generic guidelines when the source deck carries my
hand-tuned choreography".

| IC | RV | what came across |
|---|---|---|
| 9 | 45 | Making the Most 2 – body rebalanced, 2 clicks → 1 |
| 10 | 46 | Teaching Philosophy – body tightened |
| 13 | 49 | outline – pills nudged down, "✎ Problem Set 1" box removed |
| 14 | 35 | In-Class Part divider – taller title box, strip moved |
| 17 | 52 | Netflix – 4 clicks → 3 |
| 26, 29, 35 | 61, 64, 70 | Class-Discussion badge dropped from the build (it is chrome) |
| 33 | 68 | **the fruit table** – native table → 15 cell shapes, the two fruit photos, 3 clicks → 10 |
| 34 | 69 | **$30 → $50** per present |
| 39, 45, 46 | 75, 81, 82 | bodies moved / resized; Concorde 2 clicks → 3 |
| 49 | 85 | Cost-Benefit – 4 clicks → 6 |
| 50 | 86 | **the exercise chart** – rebuilt, 7 clicks → 23, with the shades and 0.05" rounding |
| 51 | 87 | "MC (Marginal Cost)" raised 0.20" |
| 53 | 89 | Next Steps – real dates ("Due on Tuesday, Oct 13") |

**Gotcha, cost one round:** the first version renumbered each copied slide's
relationship Ids sequentially and remapped the slide XML in place. In-Class
slide 33's rels are stored out of order (rId3, rId2, rId1, rId4), so the
remap renamed the same `r:embed` twice and the apple and the STOP sign came
through as "The picture can't be displayed". **Keep the source's own rel Ids
verbatim** – the slide XML is copied unchanged, so its references already
point at them, and a brand-new rels file cannot collide with anything.
Revised was restored from `_t-1` and the port re-run.

### Verified
Revised opens at 95 slides. Click counts on every ported slide match the
in-class original (RV 68 = 10 clicks / 12 effects, RV 86 = 23 / 33).
`_ic_vs_rev.py` now reports **52 of 57 pairs identical**, the five
exceptions being exactly the ones held back. Full-screen slideshow probe on
1, 35, 42, 49, 56, 59, 68, 72, 73, 86, 89, 95: no failure banner, and all
five live PollEverywhere activities render. Slides 68 and 86 rendered and
eyeballed after the image fix.

### HELD BACK – awaiting Nico (see PENDING)
1. **Three live poll slides** IC 21/24/37 -> RV 56/59/73. Same activity ids,
   but the in-class captures are full-bleed with the round POLL pill and
   open the activity in the LIVE state (`state=opened`), while Revised's are
   inset captures opening the CLOSED chart (`display_state=chart`).
2. **RV 72** – the flip-house poll's question view. In-Class collapsed that
   poll to a single slide; Revised still carries both views.
3. **RV 42 / 43** – the Econ & Coffee poll pair, dropped from the in-class
   deck but still in Revised.
4. **RV 94** – "People Respond to Incentives". Deleted from In-Class, but
   **RV 3 jump-links to it**, so removing it from Revised breaks a link in
   the video part.
5. **IC 7** – "Recall from Video 1: Economists as Hedgehogs", a new in-class
   slide with no twin in Revised (Revised has the un-retitled original at 6,
   inside the Video-1 block).
6. **IC 11 -> RV 47** and **IC 58 -> RV 95** – the roadmap / "Can These
   Prices Be Optimal?" pair. In-Class puts the jump pill on its only roadmap
   (IC 11) and links it to IC 58; Revised wires the pill to slide 7, the
   Video-1 roadmap, and RV 47 has none. Porting either one needs the link
   target decided first. IC 58 is also a substantial rebuild (13 clicks).

### Round 4 the same day — the merge finished on Nico's answers

His answers to the six questions: (1) import the poll slides; (2) keep just
one flip-house view; (3) drop the Econ & Coffee pair; (4) do NOT delete
"People Respond to Incentives"; (5) the hedgehogs slide stays only in the
intro, not duplicated; (6) adopt the roadmap / "Can These Prices" pair from
the in-class deck.

**Standing rule he gave with them: NEVER delete a slide from
`Module 1 - Revised.pptx` without asking. Only update the slides that
changed in the in-class deck.** The three deletions below are the ones he
authorised in that message and nothing else was removed.

### Ported (`_merge_ic_round2.py`, kept)
| IC | RV (old #) | carried |
|---|---|---|
| 11 | 47 | the roadmap, now with the "Can these prices be optimal?" jump pill |
| 21, 24, 37 | 56, 59, 73 | the three LIVE PollEverywhere slides – full-bleed captures with the round POLL pill, opening the activity live rather than on the closed chart |
| 58 | 95 | "Can These Prices Be Optimal?", his 13-click rebuild |

Both jump links were re-pointed at their Revised targets, the in-class way:
the in-class roadmap now links to the backup slide and that slide's Back
button returns to it.

### Deleted (authorised)
`RV 42` + `RV 43` (Econ & Coffee poll pair) and `RV 72` (the flip-house
poll's second view) – slide part, rels, notes and tags parts, their
`[Content_Types].xml` overrides, the `sldIdLst` entry and the presentation
rel, in descending order so earlier positions stay valid. Deck **95 → 92
slides**; 43 cached `slidenum` fields refreshed afterwards.

### Verified
- `_ic_vs_rev.py` (re-mapped for the new numbering): **all 57 pairs
  identical**. The only residue it prints is slide 92's page number reading
  92 instead of the in-class 58 – a live field, so that is correct.
- Every slide→slide jump link resolves: 2↔88, 3↔91, 7→92, 36→88, 45↔92,
  46→89, 90→46. **RV 3 → RV 91 still works**, which is why "People Respond
  to Incentives" was kept.
- Click counts survive the copy: RV 66 = 10 (fruit table), RV 83 = 23
  (exercise chart), RV 92 = 13 (backup rebuild), RV 45 = 1.
- Deck opens at 92 slides; slideshow probe on 1, 45, 54, 57, 66, 70, 83, 91,
  92 – no failure banner, and the three remaining live polls render.

### Two consequences worth knowing
1. **`Module 1 - Revised.pptx` is now OFF-PIPELINE.** `_build_Module1.py`
   still produces a 95-slide deck from the old content and knows nothing
   about any of this. Rebuilding would throw the whole merge away. Treat the
   .pptx as the source of truth for Module 1 from here, or budget a session
   to port the in-class work back into the build script.
2. **One asymmetry in the backup link.** Revised has two copies of the
   course roadmap: RV 7 (inside the Video-1 block) and RV 45 (in-class).
   Both carry a pill to the backup slide, but that slide's single Back
   button returns to RV 45, because the in-class wiring won. Jumping from
   RV 7 therefore lands back at RV 45. Say the word to point it at RV 7
   instead, or to drop one of the two pills.

### Round 5 the same day — the .pptx is the source of truth; pipeline retired

Nico's answer to C1: **use `Module 1 - Revised.pptx` as the source of
truth.** So the build pipeline is retired for this module, and — because
this is exactly the failure that destroyed a morning's work in Module 3 on
2026-08-27 — it is retired with a GUARD, not just a banner.

### `_m1_frozen.py` (new)
One guard module, two entry points:
- `refuse_if_canonical(target)` — exits unless the target is something
  other than `Module 1 - Revised.pptx`. **The guard is on the TARGET, not
  on the script**, so building to a side path still works and experimenting
  is unaffected; it refuses only when a pass is about to clobber the
  canonical deck. `--force` overrides.
- `refuse_already_applied(name)` — for the two one-off merge scripts, whose
  slide numbers refer to the deck as it was BEFORE the merge. Re-running
  `_merge_ic_round2.py` would delete slides 42, 43 and 72 of a deck in
  which those numbers now mean something else entirely.

### Wired into six scripts, each with a RETIRED banner at the top
| script | where the guard sits |
|---|---|
| `_build_Module1.py` | inside `build()`, **not** at module level — `_build_M1_candidates.py` imports this module for its helper layer and must keep working |
| `_animate.py` | top of `main()` |
| `_group_pass.py` | top of `main()` |
| `_splice_media.py` | before `splice()` |
| `_port_ic_to_rev.py`, `_merge_ic_round2.py` | top of `main()`, "already applied" |

### Verified
All six refuse when aimed at the canonical deck. A side path still runs:
`python _animate.py _sidetest.pptx` got past the guard and started
animating — and then died with `KeyError: 'pic:0'`, which is itself the
proof that the pipeline no longer describes this deck (its PLANS are keyed
to the old 95-slide content). `import _build_Module1` still works, so the
candidates deck can still be built. The deck was not touched: 92 slides,
unchanged since the merge.

### What this means going forward
- **Edit Module 1 by editing the .pptx**, with targeted zip + lxml surgery —
  `_slide33_edit.py` and `_slide50_shade.py` are the worked patterns, and
  `_m1_frozen.py`'s docstring points at them. Never round-trip this deck
  through python-pptx: it carries three live PollEverywhere activities.
- `_build_Module1.py` and its passes stay in the folder as the record of
  how the deck was built and as a source of helper functions; they are no
  longer the source of truth.
- The project `CLAUDE.md` for `405 Slide Revisions 2026` still documents the
  per-module pipeline as the way Module 1 is produced. That is now wrong for
  this module. **Not changed — a CLAUDE.md edit needs Nico's say-so.**

### Round 6 the same day — the merge check had a blind spot; closed

Nico: "on slide 39 you did not import the picture with rounded edges and
shade. Then check all slides again to make sure you didn't forget anything
else." He was right, and the miss was systematic rather than a slip.

### Why it was missed
`_diff2.py` compares **position, size, text and run formats only**. It does
not look at geometry presets, fills, outlines or effects. So Revised 39's
textbook cover — which gained `roundRect` adj 8000 + the deck shadow in the
in-class deck on 2026-09-22 — compared as *identical* and was never ported.
Its click comparison was weak too: it counts click GROUPS without checking
`TriggerType`, so a `withEffect` reading as a `clickEffect` passed.

**Lesson worth keeping: "the diff prints clean" is only as strong as what
the diff samples.** Before trusting one, state what it does NOT compare.

### Two new checks, and the tool chain now used for this deck
| tool | sees |
|---|---|
| `_diff2.py` / `_ic_vs_rev.py` | position, size, text, run formats, notes, click counts |
| `_style_diff.py` (new) | prstGeom + adjusts, fills, outlines, effects, table-cell fills |
| `_xml_diff.py` (new) | **everything** — normalized raw slide XML |

`_xml_diff.py` is the conclusive one. It normalizes away what differs
between the two decks by design — shape ids and names, relationship ids, the
cached text of the live `slidenum` field, and PowerPoint's per-save `extLst`
blobs — then c14n-compares the whole part. Nothing can hide from it.

### What the re-check found
- **RV 39** — the textbook cover, flat instead of rounded + shaded. Fixed by
  `_slide39_picstyle.py`, applied to the picture rather than by re-copying
  the slide part, since the pair was otherwise identical.
- **RV 42** ("Making the Most of the Course", the first one) — found only by
  the XML diff. In-Class has **0 clicks / 1 effect**: a single stranded
  `withEffect` on one of the four illustrations, so all four show as the
  slide opens. Revised still had the original **1 click / 4 effects** build
  (all four fading in together, `_animate.py` PLANS key 12). Ported with the
  new `_ic_port.py`.
- Nothing else. All three checks now agree.

### `_ic_port.py` (new, reusable)
`python _ic_port.py IC:RV [IC:RV ...]` — the wholesale slide-part copy,
generalised out of the two one-off merge scripts, carrying both hard-won
rules: keep the SOURCE's relationship Ids (renumbering them corrupted two
pictures on 2026-09-23) and reset the cached `slidenum` text to the Revised
number. It refuses a slide that carries a slide-to-slide jump, because the
target has to be chosen.

### Verified
`_xml_diff.py`: **"IDENTICAL at the XML level - nothing is missing"**, all
57 pairs. `_style_diff.py`: 0 differences. `_ic_vs_rev.py`: 56 identical,
the 57th differing only in the live page number (58 vs 92), which is
correct. Deck opens at 92 slides; RV 39 rendered and the cover now carries
the rounded corners and the shade; RV 42 = 0 clicks / 1 effect, matching
In-Class; RV 66 = 10, RV 83 = 23, RV 92 = 13 unchanged. Slideshow probe on
1, 39, 42, 54, 57, 70, 92 clean, live polls rendering.

### Flagged, not changed — RV 42 / In-Class 8
The stranded `withEffect` has nothing before it to run with, so all four
illustrations simply appear with the slide. That reads like an editing
artifact — three of the four effects deleted in PowerPoint and the fourth
left behind — rather than a choice. It is now faithfully in both decks.
Say the word and it can be cleaned either way: drop the orphan so the slide
is honestly static, or restore the one-click four-picture build.

## 2026-09-23 (round 2) – Nico's hand pass adopted; slide 50 shaded

**One-line summary.** Nico hand-edited the in-class deck further (61 → **59
slides**) and asked for shades and very slightly rounded edges on slide 50,
keeping his animation exactly. Nothing of his needed porting – this deck has
no build script and is edited in place – so the work was to inventory his
changes, verify none was lost, and make slide 50 a properties-only edit.

### Why "adopt" is a no-op here – read this before the next round
`Module 1 - In Class.pptx` is **not produced by any script**.
`_build_Module1.py` builds the 95-slide `Module 1 - Revised.pptx` and knows
nothing about this file. So a hand-edit to the in-class deck has nowhere to
be ported TO: editing in place preserves it by construction. The duty is to
*inventory* his changes and prove nothing was clobbered, not to port them.
`_whatchanged.py` (new, read-only) does the inventory: it wraps `_diff2`'s
`Deck` + `match` so it survives deletions and reorders, and reports
member-level geometry / text / run-format / notes / click differences
between any two versions of the deck.

### His changes since the morning version (61 → 59)
| slide | what he did |
|---|---|
| 9 Making the Most 2 | body rebalanced (y 2.29→2.77, h 3.97→3.01); 2 clicks → 1 |
| 10 Teaching Philosophy | body tightened (y 1.60→2.65, h 5.35→3.26) |
| 13 outline | the three "In class" pills nudged down ~0.14"; **the "✎ Problem Set 1" box deleted** |
| 26, 29, 35 | the Class-Discussion badge dropped from the build – it is chrome now, per the "poll / discussion chrome is never animated" rule |
| 34 Buying a Present | **$30 → $50** per present; body resized; notes updated |
| 37–39 flip-house poll | the EMPTY poll placeholder `_inclass_edits.py` added on 2026-09-22 is gone, and the live pair collapsed to ONE view – the question view, which keeps its `tags` part and the round POLL pill. This closes PENDING item A. |
| 39 Another Opp. Cost, 45 Sunk-cost examples, 46 Concorde | bodies moved / resized; Concorde 2 clicks → 3 |
| 49 Cost-Benefit | rebuilt as a 6-click build (was 4) |
| **50 exercise chart** | **rebuilt** – see below |
| 51 | "MC (Marginal Cost)" label raised 0.20" |
| 53 Next Steps | real dates ("Due on Tuesday, Oct 13", "Submit one copy per group"); 2 clicks → 1 |

**Live PollEverywhere activities: 4 → 3** (slides 21, 24, 37), each with its
`tags` part intact. Verified by a full-screen slideshow probe – all three
render live, no failure banner.

### Slide 50 – what he built
He rebuilt the exercise chart into named groups and a **23-click** build:
each bar segment is now a group of rectangle + label (`Group 42-48`, `81`,
`84`), the "Hour N" headers animate with five new overlay groups, and he
added "Net Benefit of Hour 3", "Net Loss of Hour 3", a `Left Arrow 36`, and
three extra rectangles on the left "Total Net Benefit" stack.

### Slide 50 – my edit: properties only
`_slide50_shade.py` (kept) changes **shape properties and nothing else** – no
shape added, removed, renamed, regrouped or reordered, and `<p:timing>` is
never touched. The animation targets shapes by id, so the build survives
verbatim.

- **13 bar rectangles** (the left stack's four, plus the nine hour-column
  segments): the deck's standard `outerShdw` (blurRad 50800 / dist 38100 /
  dir 2700000 / black at alpha 50000) **plus** `roundRect` at a **rendered
  0.05" radius** – "very slightly", per his wording.
- **adj is computed per shape**, not shared: `adj = 0.05 / min(w, h) x
  100000`, so it runs 5263 on a 1.90"-tall block and 8772 on a 0.57" one and
  the corner is the same length everywhere. This is the Teaching CLAUDE.md
  "radius as a rendered length" rule; one shared adj would have given the
  short blocks visibly rounder corners.
- **4 verdict arrows** take the shade only – an arrow preset's adjust
  handles size the head, not the corners.
- The shapes already carried an EMPTY `<a:effectLst/>`; the schema allows
  only one, so the pass fills that element rather than appending a second.
- **Left alone:** the STOP sign (a picture, and a sign – the flat-exception
  case), the two connectors beside the left stack (a shadow on a thin line
  reads as blur), all text, all chrome.

**Verified:** deck opens at 59 slides; slide 50 reports **23 on-click steps /
33 effects** via COM and the target sequence matches his build beat for beat
(TextBox 11 + Group 64 … Picture 37 + TextBox 10). `_whatchanged.py` between
his version and mine prints **(none)** – no geometry, text, notes or click
difference anywhere in the deck, which is the point: shadows and rounding are
not in that tool's signature, so a clean report means the edit touched
nothing else.

### Flagged, NOT changed – slide 50
1. **Five olive boxes hide behind the hour columns.** `Group 64 / 67 / 70 /
   73 / 76`, labelled "Total Benefit of Hour N", are filled **accent3
   `9BBB59`** with `<a:grpFill/>` children, and each spans its column's FULL
   height. They sit at the very back, so the bars cover them – except for a
   0.01" misalignment that leaks an olive hairline down the left edge of
   Hour 3 and along the top of Hour 5. They are animated on clicks 1-5,
   paired with the "Hour N" headers, so right now those clicks reveal
   something invisible. This looks like work in progress: as drawn, a
   "total benefit" box behind each column is exactly the right idea, it just
   needs to be in FRONT (or the bars made translucent). Untouched, because
   removing the hairline would hide the only visible sign that they exist.
2. **Two of those overlays are mislabelled** – the Hour-1 and Hour-2 columns
   both read "Total Benefit of Hour 1".
3. **"Net Loss of Hour 3" sits on the Hour-5 bar.** Reads as a copy-paste
   slip from the Hour-3 group; not corrected, per the never-change-content
   rule.
4. The 0.05" corner cut makes the olive hairline very slightly more visible
   at Hour 3's corners. It goes away when item 1 is resolved.

## 2026-09-23 – slide 33: fruit pictures styled, table rebuilt for per-cell animation

**One-line summary.** Nico dropped in a newer `Module 1 - In Class.pptx`
(now **61 slides**) and asked for one change on slide 33, "Opportunity
Cost: A Simple Example": adopt the two fruit pictures he had inserted, and
make each table entry animatable on its own.

### Edit 1 – the two new figures
Slide 33 carries two pictures that the Revised deck's version of this
slide does not have: a banana at (0.630, 2.470) 2.077 x 1.731" and an
apple at (10.455, 2.867) 2.352 x 2.083". Both arrived flat – `prstGeom`
`rect`, no `effectLst`. Both now carry the deck's house picture style,
`roundRect` adj 8000 + `outerShdw` blurRad 50800 / dist 38100 / dir
2700000 / black at alpha 50000, the same treatment `_inclass_edits.py`
gave slide 4's textbook cover. Both PNGs are fully opaque with white
corners, so the rounded card reads the way it does on the deck's other
pictures rather than cropping a transparent cut-out.

### Edit 2 – the table is no longer a native table
**A native PowerPoint table is ONE `graphicFrame`, so PowerPoint animates
it as a single object; there is no way to reveal a cell at a time.** To
get per-entry animation the table had to be rebuilt as 15 individual cell
shapes, one `sp` per cell carrying its own text. The group that held the
white backing card plus the table was dissolved as well, for the same
reason a group animates as one.

- Geometry, fills, borders, fonts and alignment were READ OFF the native
  table, so the final state is unchanged: columns 2.450 / 1.600 / 1.800"
  and five 0.510" rows from (3.750, 2.400); header navy `0B2B4E` with
  white bold 20 pt; column 0 alternating `FFFFFF` / `FDF6E6`, navy bold,
  left-aligned; "My Value" green `92D050`; "My Opp. Cost" red `FF5050`;
  cell borders `C8CDD3` at 9525 EMU; insets 0.1" / 0.04", anchored middle.
- The cells stay **flat and square** – the rounding rule is for cards, not
  for cells in a grid.
- Each cell is NAMED so the animation targets it by name and survives a
  re-run: `oc:hdr:0-2`, `oc:alt:1-4`, `oc:val:1-4`, `oc:opp:1-4`.
- The backing card keeps its place in the z-order and its shade.

### The build – 10 clicks
| | |
|---|---|
| static | the card, the header row, the four fruit names, both pictures |
| 1 – 4 | **My Value** 9, 5, 3, 1 – one per click |
| 5 – 8 | **My Opp. Cost** 5, 9, 9, 9 – one per click |
| 9 – 10 | the two bullets below |

The header row is static so the table reads as a table from the start and
the columns fill in under their own headings. The alternative – letting
each column heading arrive with its first value – was not taken; say the
word if that is the beat wanted.

Because each cell draws its own border, the grid is only as complete as
the cells revealed so far: the initial state shows the header row full
width and the fruit column, then the table fills to the right. Checked on
screen and it reads as intended.

### Mechanics
`_slide33_edit.py` (kept) does both edits as pure zip + lxml surgery – no
python-pptx round-trip, because this deck carries four live
PollEverywhere activities. The timing generators (`effect_par`,
`click_group`, `timing_xml`) are copied verbatim from `_animate.py` rather
than imported, per the "never verify a pass by importing it" rule.
`_animate.py` itself is keyed to the 95-slide Revised deck and does not
apply here. The script refuses to run twice (it looks for `oc:val:1`).

### Verified
Backup rolled to `Module 1 - In Class_t-1.pptx` before writing. Deck opens
in PowerPoint at 61 slides; slide 33 reports **10 on-click steps / 10
effects** via COM, and the animated shapes in order are `oc:val:1-4`,
`oc:opp:1-4`, then TextBox 9 twice – no picture and no chrome animates.
Both pictures confirmed present and static (COM: Picture 15 at L=753 pt,
Picture 17 at L=45 pt). Export render of the final state is
pixel-identical to the original table; the full-screen slideshow probe on
1, 22, 26, 33, 40, 61 shows no failure banner and the live polls still
render, which is the check that matters after OOXML surgery on this deck.
(The probe's PrintWindow capture is scaled and crops the right edge and
the footer – the apple and the footer are missing from it for that reason,
not from the slide.)

### Flagged, not changed
The Teaching CLAUDE.md standing rule is "reproduce every table as a
native, editable PowerPoint table". Slide 33 is now a deliberate exception
– per-cell animation is impossible any other way. Two consequences for
editing: a number is changed by clicking its own box rather than tabbing
through a table, and adding a row means adding three boxes. Worth writing
the exception into the Teaching CLAUDE.md if more tables need this.

## 2026-09-22 – new "Module 1 - In Class.pptx": picture style + poll slots

**One-line summary.** Nico dropped a NEW deck into the folder –
`Module 1 - In Class.pptx`, the in-class-only cut – and asked for two
edits: the picture he had inserted on slide 4 styled like the rest of the
deck, and an empty poll slide after each poll set-up. Both applied and
verified; the deck went 60 → 63 slides.

### What this deck IS – read this before touching it
- 63 slides, 16:9, hand-assembled from the Revised deck (its chrome, tags
  and Poll Break badges are the 2026-08-27 ones). It is **NOT built by the
  pipeline**: `_build_Module1.py` still produces the 95-slide
  `Module 1 - Revised.pptx`, which knows nothing about this file. So this
  deck is edited IN PLACE, and `_verify_anim.ps1` (keyed to the 95-slide
  deck) does not apply to it.
- It carries **four LIVE PollEverywhere activities** – slides 22, 26, 40
  and 41, each with its `tags` part and the "Poll Title: Do not modify the
  notes…" payload, plus the four `ppt/tags/tag*.xml` parts. Never round-trip
  this deck through python-pptx and never touch those notes.

### Edit 1 – slide 4's picture
The textbook cover Nico inserted got the deck's house picture style:
`roundRect` adj 8000 + `outerShdw` blurRad 50800 / dist 38100 / dir
2700000 / black at alpha 50000. The style was READ OFF this deck (11 of
its pictures already carry it; display 17 is the reference), not assumed
from the build script.

**Flagged, not changed:** Teaching CLAUDE.md lists book covers among the
flat exceptions that take neither rounding nor shadow. Nico asked for both
explicitly, so both were applied.

### Edit 2 – an empty poll slot after each poll set-up
Three slides carry the Poll Break badge (displays 20, 23, 36 in the
60-slide deck). Each now has an empty poll slide after it, copying display
34 of `Module 4 - Revised.pptx`: navy top bar + section tag, footer
chrome, and the round gold POLL pill at (11.567", 6.460") 1.490 × 0.510",
`roundRect` adj 50000, fill `E09F3E`, `outerShdw` alpha 45000, label
"POLL" 20 pt bold navy Calibri. No title, no body, no animation, no notes.

- **Each placeholder is built from the set-up slide it follows**, then
  stripped of everything in the 0.50"–7.10" band – title, its rule and gold
  strip, the body, and the Poll Break badge. So it inherits that slide's
  own tag and footer, which is exactly how Module 4's slide 34 relates to
  its slide 33. Slots 21 and 25 read `Module 1 · In Class · Examples ·
  Supply and Demand`, slot 39 `Module 1 · Economic Costs Include
  Opportunity Costs`.
- The pill's box and label are GROUPED, which Module 4 leaves ungrouped –
  the deck-wide "a filled box and the text on it are one object" rule, and
  this deck's own Poll Break badge is grouped too.

### The resulting sequences – and the open question
| Poll | set-up | NEW empty | existing LIVE poll | solution |
|---|---|---|---|---|
| Heatwaves / AC | 20 | **21** | 22 | 23 |
| Swiftonomics | 24 | **25** | 26 | 27 |
| Flip a house | 38 | **39** | 40 + 41 | 42 |

Nico said he would "later include the PollEV sheet" on the new slides,
which suggests he had not registered that live activities are already
there. See item A of PENDING.

### Mechanics
`_inclass_edits.py` (kept) does both edits as pure zip + lxml surgery – no
python-pptx round-trip. It builds each new slide part by hand
(`slide61/62/63.xml` + a rels file carrying only the slideLayout), registers
it in `[Content_Types].xml`, `presentation.xml.rels` and `sldIdLst`, and
inserts the three in DESCENDING source order so the earlier positions stay
valid. It refuses to run twice (it looks for the `PollPill` group), and
`--renumber` refreshes the cached `<a:fld type="slidenum">` text, which the
insertion left one to three behind on 36 slides.

### Verified
Opens in PowerPoint at 63 slides; full-screen slideshow probe
(`_slideshow_probe.ps1 -Deck "Module 1 - In Class.pptx"`) on slides 1, 4,
20, 21, 22, 25, 26, 39, 40, 63 – no failure banner and all three live polls
render live, which is the check that matters after OOXML surgery on a deck
with polls. Backup rolled to `Module 1 - In Class_t-1.pptx` first; that copy
mattered more than usual, since the deck was still untracked in git.

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

### Round 5 the same day — poll sidecars, folder cleanup, podcast lengths

**Cleanup.** The folder went from ~230 MB / 100 items to **103 MB / 24
items**. Deleted: both rolling backups, `Videos Final/`, the render and
probe folders, 36 one-off diagnostic / probe / video-port scripts, 4 extra
export scripts, the stale `_anim_config_m1.txt`, all loose contact sheets,
`Module 1 - MW.pptx` (the colleague comparison deck — its five approved
imports live in the build script), and `_handoff_pollbreak.xml` with its
now-dead injector.

**Poll sidecars.** `_splice_media.py` was reading two 65 MB source decks
for six slides. Both were carved down with PowerPoint via COM —
`_handoff_polls_WS.pptx` (68 slides → 5, 2.4 MB) and
`_handoff_polls_IC.pptx` (53 → 1, 2.2 MB) — and the splice map re-keyed to
the sidecars' own numbering (WS 7/8/25/29/46 → 1–5, IC 39 → 1). The two
entries for the deleted AC and diamonds question polls were removed rather
than left dangling. Verified before deleting the originals: the rebuild is
identical to the shipped deck, every spliced slide still carries
tags + notesSlide + image, click counts match, and the full-screen
slideshow probe on all six poll slides renders all four live activities.
The pattern is now a rule in `Teaching/CLAUDE.md`.

**Kept, decided 2026-08-27:** `Module 1 - Example Candidates.pptx` and its
build script stay in the folder. They are not needed to rebuild the module
deck, but they hold the researched runners-up Nico has not picked from yet
(DRAM / AI memory, United's marginal flights, Meta Reality Labs,
return-to-office, eggs, the Apple Car exit math, plus two bench slides).
Do not propose deleting them again. `_source_images/` also stays whole: it
stopped being redundant when the two In-Class source decks were deleted, so
the extracted `slide{N}_{rId}` images now have no other copy.

**Podcast lengths.** Nico generated the episodes: the wrap-up came out at
15 minutes (fine, ceiling raised to 20), but the **intro ran 18 minutes
against a 5-minute target**. Diagnosis: NotebookLM expands whatever it is
given, so "about 5 minutes" on a full-length doc is ignored — the intro
body was 788 words of module narrative. Fixes: the intro body cut to **396
words** (one sentence per idea, no worked examples, no connective
narrative); the prompt rewritten with a hard quantified cap stated at both
ends ("no longer than 5 minutes — roughly 700 spoken words", "about a dozen
short exchanges", "Above all, stay under 5 minutes"); and a note added at
the top of the doc to set the panel's length control to **Shorter** if it
offers one. The wrap-up was relaxed to 15–20 minutes in its subtitle,
prompt and host instructions; its body was left alone. Recorded in
`Teaching/CLAUDE.md` as "Length: the source doc controls it, not the
prompt", with the three levers and the per-episode targets.

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

---

## 2026-09-06 — Adopted the dark-red demand rule (deck-wide sweep)

Nico asked for the colour rule to be adopted here. Teaching CLAUDE.md
(2026-08-30): *"A DEMAND curve is dark red `C00000` — curve and label
alike. Supply stays navy."* Module 1 was built on 2026-08-20 and predated
the rule, so it still drew demand in GOLD and supply in STEEL (the source
deck's light blue `95B3D7`). Module 4's `_sd_curves` is the reference.

**What changed** — 33 lines in `_build_Module1.py`, applied by
`_sweep_demand_color.py` (kept in the folder; it verifies every anchor
line before writing and aborts if any has drifted):

| | before | after |
|---|---|---|
| demand curve + label | `GOLD` (label sometimes `NAVY`) | `RED` |
| supply curve | `STEEL` | `NAVY` |
| supply label | `NAVY` | unchanged |

**`RED` is the right constant, not `DARKRED`.** In this script `RED` is
`C00000` (the rule's colour); `DARKRED` is `A2162A`, the source MC-bar
red. Easy to get wrong.

**Deliberately NOT changed:**
- **Shifted curves (D′, S′) keep `GREEN_DK` / `BLUE_PED`.** Green marks
  *"this is the shifted curve"*, which is a different job from naming the
  curve type, and Module 4 keeps it for exactly that (a shifted supply in
  `GREEN_DK` on a slide whose base demand is `RED`).
- **The gold "Excess demand" band on display 29** is not a demand curve.
- All other gold (chrome, badges, table headers, roadmap dots).
- In the tuple-driven charts the *dashed* shifted curve follows its base
  curve's new colour, since there the dash already marks the shift —
  display 65 (LA real estate) now reads solid/dashed dark red for D₀/D₁
  and solid/dashed navy for S₀/S₁, which is clearer than before.

The tea, avocado and copper chart edits are **dead code**: those slides
are mapped to `None` in `_m1_order.py`. Swept anyway so the script stays
consistent if they ever come back.

**Verified:** `95B3D7` is gone from all 95 slides; displays 22, 29, 31
and 65 exported and eyeballed; full pipeline re-run
(`_build_Module1.py` → `_splice_media.py` → `_group_pass.py` →
`_animate.py all apply`); full-screen slideshow probe PASSED on 1, 22,
29, 31, 57, 65, 72, 73, 95 — including the two PollEv slides, which
render live. Backups rolled before the rebuild. Deck was committed and
clean beforehand, so there were no unported hand-edits at risk.

**Note for other modules:** Module 2's deck has not been checked against
this rule. Module 3 likewise.
