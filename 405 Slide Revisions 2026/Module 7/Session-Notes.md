# Module 7 (Managerial Economics 405) – Session Notes

## 2026-08-20 – Agenda slides retrofitted to numbered-circle format

**One-line summary.** `make_module_outline` in `_build_Module7.py` was
replaced with the Module-2 numbered-circle implementation (Teaching
CLAUDE.md "Module-Outline / Agenda Slides"); new module-level
`M7_OUTLINE` = Nico's approved flat 6-item list with one-line
descriptions (Collusion / Cournot / Bertrand / Differentiated Goods /
Strategic Thinking / Classic Games). The legacy `part_idx`/`sub_idx`
call-site signature is kept and mapped onto the flat index (part 0
subs 0–3 → items 1–4; part 1 subs 0–1 → items 5–6); the page-4
overview call now passes `descriptions=True` (all descriptions shown);
section agendas at 11/17/29/38/48/55 band + describe only the current
item. Full pipeline rerun (build → splice → animate; click counts
match the documented plans), agenda renders verified, slideshow probe
PASS. Backup: `Module 7 - Revised_backup_2026-08-20.pptx` (kept with
`..._2026-08-14`).

---

## 2026-08-14 – Nico's slides 40–47 hand edits adopted (3 gates green)

**One-line summary.** Side-path pass #18: all hand edits in slides
40–47 ported into the scripts and gate-verified. **S42/43** (Airbus
reaction slides): P-axis label nudged per slide ((1.89, 1.52) on 42,
(1.93, 1.52) on 43); on 42 the MR line + "MR A" label + its "300"
foot tick are now a hand GROUP that clicks as beat 4 AFTER the
circle-note beat; on 43 the 200→400 circle + "400" tag are a GROUP
riding WITH the new-demand beat. **S44** (equilibrium prices) rebuilt
in his 7-group topology — A-reaction (line + label + 200 tick),
B-reaction, starting point (150 tick + guide + dotted card), reaction
points I and II (guides + dot + ticks + label + pointer each), the
267-267 guides + ticks, and the equilibrium (dot + gold callout +
arrow) — with his 11-click story: TITLE first, then the two reaction
POINTS before the lines through them, then start, the 3 staircase
steps one by one, the 267 guides, equilibrium finale. **S45**
(BruinLearn) now STATIC (SKIP_STATIC) and the nav box moved up to
(3.67, 5.14). **S47**: both discussion paragraphs on ONE click, gold
badge static. Engine gotcha logged: a group's concatenated text can
start with a tick's digits, so bare `t:300`/`t:400` selectors matched
the new GROUPS instead of the y-ticks — fixed with doc-order-indexed
`t:300#2` / `t:400#2` / `t:400#3`. Gates: geometry identical 1–47;
timing text-insensitively identical 1–47 (one cosmetic concat-space
artifact on the S44 click-4 group signature); slideshow probe PASS.
Swapped over the canonical deck. Backups: `..._2026-08-13b` +
`..._2026-08-14`.

---

## 2026-08-13 (evening) – Simultaneity: S39 bullet dropped, S40 notes

**One-line summary.** Pedagogical decision (Nico, after my
assessment): "Set prices simultaneously" DROPPED from slide 39 (the
motivation slide) — it duplicated slide 40's third assumption; kept
on slide 40, the only place the assumption is listed on-slide.
Ported his deletion into the builder (bullet removed; box height
adopted at his auto-fit 3.59). Slide 40 gained speaker notes with
the why-simultaneity explanation: neither firm observes the rival's
price nor can commit to moving first → equilibrium = mutual best
responses, not leader-follower; answers "why can't the second mover
just wait and undercut" (no last mover); and Part 2's commitment
material deliberately BREAKS this assumption. Gates: geometry
identical 39–40, timing identical, slideshow PASS. Swapped. Backups
`..._2026-08-13` + `..._2026-08-13b`.

---

## 2026-08-13 (later) – S27 title + Nico's 35–39 edits adopted

**One-line summary.** Side-path pass #17. **S27 title renamed
"Cournot: Computation" → "Cournot: Math"** (his change, ported to
the builder). **S35:** "100" x-tick rebuilt tight at (6.79, 6.31,
0.54×0.27); the Bertrand cluster (vertical guide + gold point +
"90" tick + 3 labels + pointer) is now a hand GROUP — his
choreography: the horizontal price guide clicks ALONE (beat 6), the
grouped cluster is the finale (beat 7). **S36:** given-formula moved
to y 1.58 and BOLDED — gotcha: OMML ignores `a:rPr b="1"`; math
weight comes from `<m:sty>` (bi for variables, b for upright runs) —
the builder now post-sets m:sty on the equation; separator-3 moved
up to 5.53; top separator now STATIC; his 4-click plan = Monopoly
paras 0–4 / sep2 + Cournot para 7 / sep3 + paras 10–12 / the
practice-video box as the finale (engine chrome rule doesn't block
explicit t:▶ selectors). Second gotcha: `_add_mixed_textbox`
"break" segments produce N−1 empty paragraphs (a break STARTS a new
paragraph, it doesn't append one) — the earlier "extra air" note
mis-stated the para indices; his indices (7, 10–12) match the
ORIGINAL 3-break structure, which was restored after a render diff
showed separators striking text (geometry gate can't see paragraph
layout — pixel-diff caught it). Gates: geometry identical 1–39
(remaining: two auto-fit box-height cache artifacts on S27
title/S36 formula, pixel-verified invisible), timing identical
1–39, slideshow PASS. Swapped. Backups `..._2026-08-12b` +
`..._2026-08-13`.

---

## 2026-08-13 – Nico's slides 14–35 hand edits adopted (3 gates green)

**One-line summary.** Side-path pass #16 with the full gate battery.
**Geometry ports:** S24 "Firm A's Demand" label shrunk to
2.46×**0.50** (his explicit callout; `firm_lbl` gained an optional
height); S34 tick boxes "60"/"61" rebuilt tight at his coords
(4.88/5.15, 0.36/0.41 wide). **New hand groups →** builders: S23/24
market-quantity guide + navy dot + tick = "MktGuide" group
(`_fig_x/ytick` now return their textbox so ticks can be grouped);
S26 gained SEVEN groups (ReactionA, ReactionB incl. pointer,
Fifteen = tick + first staircase segment, Thirty = eq guides + both
"30" ticks, Equilibrium = gold point + arrow + gold box + OMML,
Callout, Starting). **Choreography (per his timing):** 16 + 18 now
STATIC (SKIP_STATIC); S19 = 1 click (Bertrand column only); S20
headers static; S21 = his bullet partition (paras 5 / 7–8 / 10 / 12,
logos static); S23/24 = mkt-guide group gets its own click AFTER the
conclusion; S26 = 9 clicks (A → B → starting card → 15+step1 →
steps 2/3/4 → 30-30+equilibrium → callout); S28 = 1 click: steel
bullet + whole pie panel INCLUDING the chart graphicFrame
(`_animate.py` gained graphicFrame targeting, selector `gf:N`);
S34 "30" tick reveals with the strips. Gotcha hit: grouping shifted
the textless-shape indices — S23/24 callout became osp:6 (caught by
the timing gate). Cached footer numbers 22/23 → 23/24 fixed in the
Cournot wrappers. **Gates: GEOMETRY IDENTICAL 1–35, TIMING IDENTICAL
1–35 (type+position+prg), slideshow probe PASS.** Swapped. Backups
`..._2026-08-12` + `..._2026-08-12b`.

---

## 2026-08-12 (evening) – S13 green box size + geometry-diff gate

**One-line summary.** Nico caught a missed hand-edit: the green
"Extra Profit" region on slide 13 was resized by him to 1.63×**0.78**
(the rewrite had transcribed the old 0.94 from the previous builder
— classic eye-pattern-matching error when porting from a dump).
Fixed in `slide_13_cartels`. Built **`geom_diff.py`** (scratchpad):
decodes EVERY shape in two decks to rendered inches (group children
via chOff/chExt), matches by type+text (normalizing math-italic
codepoints 𝑃→P), compares position AND size at 0.01" tolerance —
result after the fix: **GEOMETRY IDENTICAL for slides 1–13** vs his
backup. Slideshow probe PASS, swapped. New Teaching CLAUDE.md rule:
every hand-edit port closes with this automated member-level
geometry diff (plus the beat-for-beat timing diff and slideshow
probe) — never sign off from a visual read of the dump.

---

## 2026-08-12 (later) – Nico's ANIMATION choreography adopted exactly

**One-line summary.** Correction to pass #15: Nico's hand edits on
slides 1–13 included RE-CHOREOGRAPHED ANIMATIONS, which the first
port regenerated from the generic guidelines instead of preserving.
Built `extract_timing.py` (scratchpad): parses a deck's `<p:timing>`
per slide into ordered click → effect lists with spids resolved to
shape SIGNATURES (type, x, y, text — ids differ across rebuilds).
Extracted his beats from `backup_2026-08-12` and re-targeted
`_animate.py`: **S7 = ONE click** (spectrum fully static, only the
WE-ARE-HERE group reveals); **S9 = 5 clicks** (paras 1–2 together;
para 3; para 4 + chart group; merger card ALONE; paras 5–7);
**S12 = 2 clicks** (definition card STATIC; aims card; quote card);
**S13 = 10 clicks, competition-first story** (competitive outcome R
→ P_Comp+d → q_Comp/zero-profit → MR → cartel outcome R → P_Cartel
+ text line 1 → q_Cartel + line 2 → profit region → q_Dev + line 3
→ extra-profit region; panels/curves/headers all static). S2/5/6/10
already matched. VERIFIED: extract_timing on the rebuilt deck diffs
IDENTICAL beat-for-beat against his backup for all of slides 1–13;
slideshow probe PASS; swapped. New Teaching CLAUDE.md rule:
hand-edits include timing — diff `<p:timing>` and accept only a
beat-for-beat match; never regenerate over hand-tuned choreography.

---

## 2026-08-12 – Nico's slides 1–13 hand edits adopted (groupings)

**One-line summary.** Side-path pass #15. Diffed all 13 slides
(geometry + run formatting + OMML colors + notes); only 7/9/12/13
changed — mostly Nico applying the GROUPING conventions by hand.
Added **`_group_shapes(slide, shapes, name)`** (build-time grpSp
creator, off/ext = bbox, chOff/chExt = off/ext) +
`_shape_ids`/`_new_shapes_since` snapshot helpers; builders now group
at creation. Ported: **S7** WE-ARE-HERE trio (label + red arrow +
red Oligopoly box) = one group, animated as the finale beat; **S9**
bullet now "Limitations **of these measures**", chart panel moved UP
0.6" (chart_t 2.25→1.65, everything keyed off it) and grouped
("WirelessChart"), merger card rebuilt at his bottom-right position
(6.666, 6.093, 6.391×0.838, logos at his exact coords) and grouped
("MergerCard"); **S12** three cards (definition frame+label+hero /
cream aims / Smith quote) each = one group, animated as 3 beats;
**S13** builder REWRITTEN in his grouped topology — 10 groups
(P_Cartel line+tick; P_Comp+d+tick; q_Cartel, q_Dev guides+ticks;
profit and extra-profit regions each with label+arrow;
q_Comp+zero-profit; MR+label; cartel outcome guides+dot+ticks;
competitive outcome) — and his EXTENDED bottom takeaway ported
verbatim with subscript runs ("…, and this reduced output leads to
a higher market price P_Cartel"; "(that is, each firm must reduce
its output)"); NOTE his line 1 has lowercase "Q_cartel" vs
"P_Cartel" — ported verbatim, flagged to Nico. Ignored as save
artifacts: OMML math-italic codepoints (𝑃 vs P) and run splits with
identical formatting. `_animate.py` PLANS 7/13 re-targeted to group
beats (7: 6 clicks; 9: 2 clicks/9 effects — bullets + 2 groups; 12:
3 card beats; 13: 9 clicks/21 effects). Render-matched vs his deck,
slideshow probe PASS, swapped. Backups `..._2026-08-11b` +
`..._2026-08-12`.

---

## 2026-08-11 (evening) – FIXED: slideshow crash ("slide failed to open")

**One-line summary.** Nico's slideshow error on slide 1 root-caused
and fixed: **the spliced PollEv slide 33 had its live
`__PE_POLL_EMBED_ID` tag but NO NOTES PART** — the PollEverywhere
add-in scans the deck at slideshow start and reads the poll data
from the slide notes; missing notes → the add-in crashes the
slideshow renderer DECK-WIDE, surfacing as "The slide failed to open
properly" on slide 1. Editing canvas, exports, and COM slideshow
stepping all render fine (different code paths), which is why the
first diagnosis missed it. **Debug method (reusable):** drive the
real slideshow via COM, screenshot the `screenClass` window with
PrintWindow (PW_RENDERFULLCONTENT), auto-classify pass/fail by
pixel, and bisect with `sldIdLst`-subset decks
(scratchpad `bisect_show.ps1` + `make_subset.py`). Bisection:
animations exonerated (TEST-A no-anim failed too) → prefix 1-30 PASS
/ 1-38 FAIL → pair [1,33] FAIL → tag removed PASS → notes added PASS.
**Fix in `_splice_media.py`:** the slide-33 branch now always
generates `ppt/notesSlides/pe33_notes.xml` (from a template
notesSlide, body swapped with the poll notes) + its rels + content
type + rId3 from the slide; new BUILD-INPUT sidecar
`_handoff_s32_notes.txt` (poll title + PollEv URL — his poll's exact
notes text). `_animate.py` now accepts a deck-path argument. Teaching
CLAUDE.md gained the poll-notes gotcha + debug technique. Full
pipeline rebuilt (build → splice → animate), full-deck slideshow
probe PASS, poll tag + notes verified via COM, swapped; TEST-A/B
decks deleted. NOTE: slide 33's speaker notes are now the PollEv
notes (that's what makes it work — never replace them).

---

## 2026-08-11 (later) – Animations for slides 1–47

**One-line summary.** New **`_animate.py`** (adapted from the Italy
IBR engine): injects Fade 0.5 s on-click `<p:timing>` blocks.
**Pipeline is now: build → splice → `python _animate.py all apply`.**
34 slides animated; skipped per guidelines: 1 (title), 3 (roadmap),
4/11/17/29/38/48 (outline/agenda), 8 (table slide — static), 22/31
(section cards), 15 (video), 33 (poll), 27 (practice-video index).
Defaults: text slides build per top-level bullet (first one static),
figures + captions ride on the FIG_GROUP bullet (10 → bullet 1),
takeaway bars last; chrome/backings/video-link boxes/nav pills never
animate (geometry-based detection — M7 uses "Rectangle N" for
content, so the Italy name-based rule was replaced). Custom PLANS
(selector language t:/pic:/cxn:/osp:/grp:/pr:) for 5 (concept map,
10 clicks ending on the gold bridge), 6 (red circle only), 7
(spectrum column-by-column), 13 (cartel 2-panel, left→right, 9
clicks), 19/20 (column beats), 23/24 (residual demand: market
demand → MC → callout → residual+brace → MR → optimum → conclusion),
25 (points before the line), 26 (reactions → staircase → gold
finale), 34 (price war: demand→MC→callout→$40→$39→strips→eq→
takeaway), 35 (three equilibria in welfare order), 36 (3 section
beats — NOTE: multi-para pRg ranges get re-expanded into separate
clicks by PowerPoint; emit one effect per paragraph within the click
group instead), 41–44 (reaction-function stories, gold equilibrium
last). Gotchas: t: selector is first-UNUSED-match — order "50 units"
before "50"; `t:PREFIX#n` = nth doc-order match for duplicate tick
labels. COM-verified click counts on all 34 slides (TriggerType 1 =
on-click); deck reopens (88 slides). Nico to eyeball the slideshow.
Backups `..._2026-08-11` + `..._2026-08-11b` (pre-animation).

---

## 2026-08-11 – Slide 5 hand edits adopted

**One-line summary.** Ported Nico's hand edits on the new concept-map
slide 5: (1) Bertrand box wording now "identical **goods**:\nP = MC\n
differentiated **goods**:\nP > MC" (explicit line breaks instead of
his padding-space hack); (2) Nash box: "mutual best responses" →
"**individually** best responses"; (3) NEW second gold 3 pt inflow
arrow into the bridge box from the Cournot side, (3.753, 4.952) →
(6.55, 6.125). Rebuilt → spliced → render-matched against his hand
version → swapped; reopens fine (88 slides). Backups
`..._2026-08-10b` + `..._2026-08-11`.

---

## 2026-08-10 (later) – NEW slide 5: concept map; deck is now 88 slides

**One-line summary.** Inserted **slide 5 "How the Pieces of Module 7
Connect"** (tag "Module 7 · Concept Map"), modeled on Module 3's
concept map but shorter: two clusters — OLIGOPOLY (Part 1: root +
Collusion/Cartel, Cournot, Bertrand children with their price
outcomes) and GAME THEORY (Part 2: root + Dominant strategy, Nash
equilibrium, Commitment) — all boxes rounded + shaded per the
current standard (NOT M3's flat boxes); MB=MC anchor star under
Cournot ("given the rival's action"); gold bridge box "Cournot &
Bertrand are games: their equilibria are Nash equilibria" with
inflow from Bertrand and outflow to the Nash child. **ALL
SUBSEQUENT SLIDES RENUMBERED +1 (deck 87 → 88):** slide functions
`slide_NN_*`, `_draw_footer` page numbers, `make_stub` /
`make_module_outline` / `page_num=` / `_airbus_reaction_slide`
literals, and build() comments all bumped via descending regex;
`_add_ticket_jump` now links display 68 → 88's sibling 87 (indices
[67]/[86]); `_splice_media.py` poll display 32 → **33**, video
14 → **15**. IMPORTANT: all slide numbers in EARLIER session-notes
entries refer to the pre-insert numbering (subtract nothing before
this entry; add 1 to compare with the current deck for slides ≥5).
Poll tag re-verified on display 33 after splice; deck reopens (88
slides). Backups `..._2026-08-10` + `..._2026-08-10b` (pre-insert).
PENDING: speaker-notes batch for old-numbering slides 1–46 drafted
in chat, awaiting Nico's voice confirmation (slide refs there use
OLD numbering).

---

## 2026-08-10 – Adopted Nico's updated PollEv (32) + slide-14 note

**One-line summary.** Nico hand-inserted the UPDATED PollEv poll on
slide 32 ($120-anchor bands, new embed ID
`a7ca8c79-93bb-46c5-82d0-9bda44e13d46`, picture already 16:9-sized
by the add-in). Preserved as sidecar files —
`_handoff_s32_poll.xml` (his slide XML verbatim),
`_handoff_s32_tag.xml` (his tagLst), and
`_source_images/_s32_pollev.png` (his screenshot) — and
`_splice_media.py` now splices slide 32 FROM THE SIDECAR (no
x-shift; rels built with the exact rId1=tags / rId4=image ids his
XML references) instead of from the original deck. These three
sidecar files are BUILD INPUTS — never delete. Slide 14: his
changed stub note ported into `build()` ("copy video over by hand
(not done yet to keep file size manageable)"; make_stub adds its
own brackets, so the text is passed bare). Rebuilt → spliced →
swapped; COM-verified the new embed ID and that the deck reopens
(87 slides). Backups `..._2026-08-09` + `..._2026-08-10`.

---

## 2026-08-09 (evening) – Phase 3 begins: rerunnable media splice

**One-line summary.** New **`_splice_media.py`** — phase-3
poll/video splicing WITHOUT freezing build.py. **Pipeline is now:
`python _build_Module7.py` → `python _splice_media.py`** (rerun both
after every edit round; the splice always starts from fresh build
output, so nothing accumulates and the file size stays constant).
Pure zip + ElementTree surgery (python-pptx would strip poll `tags`
and NULL video rels): copies the original slide part, remaps rels
(stub's layout kept, media/tags copied under collision-safe
`peNN_*` names, External rels kept verbatim, content-type entries
added), and shifts all `<a:off>` x by +1.667" to center 4:3 content
on our 16:9 canvas (string-level regex — an ET round-trip would
break the `mc:AlternateContent Requires="p14"` prefix binding; NOTE:
also shifts group-child offsets, fine for group-less poll slides,
recheck for slide 14). **Spliced now: slide 32** (concrete-pricing
PollEv; +76 KB; `__PE_POLL_EMBED_ID
488c41e0-4fa6-40ee-9325-75d3d3b03070` verified live via COM Tags;
screenshot still shows old $50 bands — add-in pulls Nico's updated
poll by ID). **Slide 14 video (embedded 11.5 MB mp4) deferred
behind `--with-video`** — run only for the final/teaching copy so
the working deck and its git commits stay ~12 MB (Nico's size
concern; answer: no per-round growth, only the one-time video
weight). Remaining media (47+): polls 52/69/72/77, videos 74/76/78
(externally linked, cheap) — extend the `displays` list later.
Deck reopens fine at 11.9 MB.

---

## 2026-08-09 (later) – Slides 63/66/67/75/79: original techniques

**One-line summary.** Side-path pass #14. (1) **Slide 63:** original
commons image restored — the original deck shows `image46.png`
CROPPED via `srcRect` (t 16.1 %, b 38.1 %, r 1 %) to just the
3-panel sheep strip; extracted + pre-cropped to
`_source_images/_s63_commons.png`, placed rounded + shaded, centered
(h 2.95). (2) **Slide 66 (Group Work):** the matrix now carries the
FULL answer under a cover — payoffs 0,0 / −100,−100,
"(Cell impossible due to M.A.D.)" notes, best-response arrows, and a
small gold oval around the "0, 0" pair (full-cell nash oval poked
out from under the cover) — all hidden by a light-blue **gradient
cover box** over the cell area (new helper `_add_cover_box`:
F2F7FC→C9DCEE vertical, no line; instructor deletes it live).
(3) **Slide 67:** red X now sits in the LOWER half of each
impossible cell, BELOW the caption ("under the text" = positioned
below, per the original layout — caption top-anchored at 12 pt so it
wraps 2 lines); added the missing bottom-right link: original's
action button jumps to slide 86 (ticket-commitment example), ported
as a navy "Tickets →" pill in the standard corner position with a
`ppaction://hlinksldjump` link (post-pass `_add_ticket_jump` since
slide 86 must exist first). (4) **Slide 75:** the 4 gradient cover
squares over the payoffs restored, plus the original's 4
best-response arrows; matrix ovals drawn manually AFTER the covers
so ovals/arrows sit on top. (5) **Slide 79:** crosses below text
(as 67) + the original's small red X through Steven's "Split" row
label. `_payoff_cell` override captions: TOP-anchored 12 pt
deck-wide. Backups `..._2026-08-08b` + `..._2026-08-09`.
Render-verified vs originals, swapped, reopens fine.

---

## 2026-08-09 – Payoff-table covers (56/57) + arrow geometry (all)

**One-line summary.** Side-path pass #13. (1) **Cover technique from
the original deck** on the strategy-analysis slides 56/57: the
non-relevant player's payoff is now hidden behind a small GRAY box
(0.46×0.34 at the payoff anchor, GRAY fill, no line) instead of
being rendered faded; the hidden run is drawn WHITE so wide numbers
("-10") can't peek out from under the box. Implemented generically
in `_add_payoff_matrix` (dim='row'/'col'). (2) **`_br_arrow`
rewritten to the ORIGINAL arrow geometry, applied to ALL
payoff-table slides** (56, 57, 58, 61, 73, 82, 83, 84): a VERTICAL
best-response arrow runs number-to-number (start 0.26" below/above
the source payoff, stop 0.40" clear of the target's circle); a
HORIZONTAL arrow sits 0.30" ABOVE the payoff line (e.g. slide 57:
from -1 to 0 and from -10 to -8, above those numbers) so it never
strikes through them — the old version shrank the anchor-to-anchor
line 30 % per side, which left horizontal arrows crossing the
numbers. Render-verified 56/57/58 side-by-side vs the original and
spot-checked 61/62/71/73/82/83. Backups `..._2026-08-08` +
`..._2026-08-08b`. Swapped, reopens fine.

---

## 2026-08-08 (later) – Picture+caption grouping rule; slides 48/50

**One-line summary.** Side-path pass #12. (1) **New Teaching
CLAUDE.md rule** (Pictures section): a picture and its
caption/source line are ONE object — group them into a `<p:grpSp>`
(off/ext = bbox, chOff/chExt = off/ext); build scripts do it in a
deck-wide post-pass. (2) **Implemented `_group_pic_captions(prs)`**
in the build script, called in `build()` before save: matches each
picture to a small all-italic ≤13 pt text box within ~0.3" of its
top/bottom edge AND centered on it (|caption center − picture
center| ≤ 0.35" — the looser "within span" test falsely grouped
slide 27's pie-chart TITLE with the oil photo). Groups created on
slides 13, 15, 48, 50. Deferred to phase 3: slide 8/27 chart
source-lines (they belong to the chart+backing grouping pass, not a
picture). (3) **Slide 48 hand edits:** Nash portrait enlarged
(8.813, 1.401, h 2.949) + caption (9.038, 4.43); Nobel medal to
top-right (10.933, 1.95) drawn FIRST so it sits behind the
portrait; Beautiful-Mind poster to (6.466, 4.62, h 2.775).
(4) **Slide 50 hand edits:** bullets box 6.854×3.256; Keynes photo
enlarged to (7.4, 3.51, 5.58×2.988 — his slightly squashed aspect,
height set explicitly so the caption grouping matches) + caption
at (7.4, 6.577). Backups `..._2026-08-07b` + `..._2026-08-08`.
Render-verified vs hand deck, swapped, reopens fine.

---

## 2026-08-08 – Outline bands, gold equilibrium boxes, 40–43 hand edits

**One-line summary.** Side-path pass #11. (1) **Outline highlight
band fix (all 6 outline slides):** band top now `y + 0.10 +
sub_idx × 0.477"` — the actual row pitch measured from Nico's
hand-placed band on slide 37; the old 0.52"/row assumption drifted
the band lower with each row (screenshot complaint). Verified on
slides 10/16/28/37/47/54 for every sub_idx. (2) **Slides 33 + 43:
equilibrium callouts in the slide-25 gold style** (gold rounded box,
corner 0.12, navy bold, shade): 33 "Bertrand Equilibrium" at
(6.05, 4.92); 43 gold box at (7.05, 3.62) with the OMML
"P_A = P_B = 267" overlay (replaces his white/gray-border box).
(3) **Slide 40 hand edits:** pictures enlarged (A350 8.67/1.465
w 4.08; 787 7.708/3.35 w 5.042), gold oval (4.783, 4.56,
1.047×0.62) at 2.25 pt, plain 3 pt gold line (no arrowhead),
"substitutes" at (6.475, 4.47), Note box moved to (6.92, 5.2,
4.27×0.75) at 18 pt **+ drop shade** (his request). (4) **Slides
41/42 hand edits:** MR_A inside the plot; demand labels at
(2.814, 2.862)/(2.995, 2.182); small ovals circling "100"/"P_B"
with repositioned tags + 2 pt connectors; notes moved + NEW 2 pt
connector from each note toward its circle; MC label bold-ITALIC
inside plot (41 below line at 7.95/4.469; 42 above at 8.056/4.033);
callout text 18 pt; 42 prev-demand label single-line at
(6.981, 5.914) + 2 pt pointer + green shift arrow (6.287, 6.059)→
(6.772, 5.486). (5) **Slide 43 hand edits:** y-title (0.538, 3.486);
reaction labels (8.96, 2.058)/(6.27, 1.41); 150-guide now has an
arrowhead; starting box (8.718, 5.0); RP II label BELOW its point
(8.718, 3.323) w/ 2 pt connector; RP I connector 2 pt; NEW 3-step
gray dotted convergence staircase with arrowheads
((5.958,4.464)→(5.958,3.0); (5.96,3.042)→(6.31,3.042);
(6.3,3.047)→(6.3,2.891)). Backups `..._2026-08-07` +
`..._2026-08-07b`. Render-verified vs hand deck, swapped, reopens.

---

## 2026-08-07 (later) – Run-formatting sweep, slides 45+

**One-line summary.** New CLAUDE.md rule ("Copied text keeps its run
formatting — italics, bold, underline", under Reformatting-vs-new-
content), then a programmatic sweep comparing every emphasized run in
original slides 45–87 against the revised deck. Fixed: slide 45
underlined "also raise" + "optional" (Note line de-italicized to
match original); slide 53 italic "individual"; slide 58 box text now
plain with italic "dominant strategy" (was all-bold); "Interpretation
of payoffs/arrows:" lead-ins on 55, 56, 57, 61, 62, 71, 73, 82 now
UNDERLINED not bold (original style). Helper upgrades: `_add_text`
(in `_build_template_samples.py`) and `_add_convention_box` run
styler now support `underline=`. Skipped as deck-standard chrome:
bold-italic on original "Poll/Discussion Break" badges (our badge
spec is bold non-italic). Sweep re-run: 0 remaining issues. Swapped,
reopens fine. Sweep script pattern: extract (text, b, i, u) per run
from original slide XML, flag partial-emphasis runs missing in the
revised deck — reusable for future modules.

---

## 2026-08-07 – Slides 37–43: hand-edit ports + 41–43 original layout

**One-line summary.** Side-path pass #10. Slides 37–39 hand edits
ported: divider sub-list nudge (+0.04/+0.10, applied in
`make_module_outline` so all dividers stay consistent), slide 38
underline emphasis ("identical products", "Differentiated",
"differentiated") with box h 4.04, slide 39 red C00000
"Differentiated" run **in the title**, red "differentiated" +
underlined "simultaneously" in the assumptions card. Slide 40:
firm-color equations (Q_A/P_A dark red C00000; Q_B/P_B dark green
ACC3_50 4D5D2C; MC_A red, MC_B green) and the ORIGINAL substitutes
illustration — gold oval around the "+ P_A" term of Q_B's demand, one
gold arrow, italic "substitutes". **Root-cause fix:**
`_add_math_equation` was clobbering per-run OMML colors (it stripped
and rewrote every run's solidFill with the equation default); now it
only fills runs WITHOUT an existing solidFill, inserted at position 0
(fill-before-latin rule). Slides 41–43 rebuilt to the ORIGINAL deck's
arrangement per Nico: **NO shade** (overlaps text; deliberate
exceptions, commented), paired red-bordered callouts top-right
(assumption + optimal response), olive oval circling the cross-price
term with "200"/"400" tag + note ("A higher price P_B increases
demand…"), MR_A at the MR foot, ticks navy/gray like the original;
42 adds the red previous-demand line + pointer label + green shift
arrow; 43 has plain "Equilibrium P_A = P_B = 267" label + navy arrow
into the intersection, gray DOTTED starting-point box, single dotted
150-guide (staircase removed), reaction-point labels + connectors,
red/green rotated axis titles. Charts stay TRUE-LINEAR (originals
had internally inconsistent stylized axes; "economically correct
beats pixel-matching"). Backups: `..._2026-08-06b` +
`..._2026-08-07`. Verified renders vs originals, swapped, reopens.

---

## 2026-08-06 (evening) – 33/34 hand-edit ports, 35 separators, box standard

**One-line summary.** Side-path pass #9. Slide 33 rebuilt at Nico's
exact geometry (fig 1.7/6.25, 6.4×4.25 over 112; equation-only demand
label at 2.187/2.588; strips 40–35 / 35–10; $39 + 61 ticks GOLD, 20→
x-tick offsets 57/64 for the 60/61 pair; "Bertrand Equilibrium" at
6.134/4.977 with the down-arrow (7.01,5.3)→(6.88,5.796) into the
(90,10) point; callout border 1.0 pt). Slide 34 rebuilt likewise
(fig 6.0 wide; annotations inside the plot: Monopoly at 4.657/2.32
w/ arrow (5.015,3.03)→(4.211,4.108); Cournot at 5.157/3.961 w/ arrow
(5.714,4.293)→(4.953,4.695); Bertrand block at 8.02/4.43 w/ long
arrow (7.915,4.65)→(6.546,5.811)) — **backing card has NO shade**
(deliberate exception: shade would overlap the annotation text; navy
"MR" label). Slide 35: title now "…: Math (for your reference)";
block moved to (0.63,2.27); extra blank rows before Cournot / Perfect
sections; **3 gold separator rules with drop shade** above Monopoly /
Cournot / Perfect competition (y 2.2 / 4.52 / 5.62). Slide 36 headers
(and slide 85's, same violation) switched to
`_add_rounded_filled_box` (rounded + shade). **Teaching CLAUDE.md
amended** with the root cause: the legacy flat `_add_filled_box` in
the reused M3 helper layer is how flat headers kept slipping in —
content boxes must use `_add_rounded_filled_box`; audit confirmed no
slide-level flat calls remain. Backups: `..._2026-08-06` +
`..._2026-08-06b`. Verified renders, swapped, reopens fine.

---

## 2026-08-06 (later) – Slide 31: MC updated to $120 per cubic yard

Per Nico's decision after the fact-check: slide-31 bullet now reads
"Same marginal cost: MC = $120 per cubic yard" (no market-price
context line, per his preference). Speaker notes updated with the
sourcing (NRMCA: materials $89–96/yd³ + delivery ≈ $15–25 → MC ≈
$120; 2025 avg price ≈ $180/yd³). **Nico is updating the
PollEverywhere answer bands himself** (slide 32 splice unchanged);
suggested bands were <$120, $120, $121–130, $131–145, $146–160,
$161–180, >$180. Rebuilt in place, verified.

---

## 2026-08-06 – ROOT CAUSE of lost formula colors + 21/30 background design

**One-line summary.** Side-path pass #8. **Root-cause fix:** the OMML
color post-processing in `_add_hierarchical_bullets` /
`_add_mixed_textbox` / `_add_math_equation` APPENDED `<a:solidFill>`
after `<a:latin>` inside `a:rPr` — schema-invalid, so PowerPoint
silently IGNORED the color (this is why Nico's red/green formula
colors kept reverting to black and he had to re-fix by hand). Fix:
`insert(0, fill)`; gotcha codified in Teaching CLAUDE.md. Ports:
slide 22 firm-demand label all dark red (works now); slide 23 label
all red at his newest geometry (2.342/2.793, w 2.686) with the pointer
arrow now DARK RED (was theme blue); slide 25 y-title moved to
(0.728, 3.795), fully red incl. Q_A; x-title fully olive incl. Q_B
(colors now actually render). **Slides 21/30 redesigned by Nico:**
full-bleed AI-illustration backgrounds (extracted to
`_source_images/_s21_image27.png` / `_s30_image43.png`; picture added
FIRST so chrome/footer stay on top) + his assumptions cards as SCALED
GROUPS — injected verbatim from `_handoff_s21_group.xml` /
`_handoff_s30_group.xml` via new `_inject_handoff_group` helper
(group scaling scales text, which a rebuild can't replicate; shape
ids re-based to avoid collisions). KEEP the two _handoff_*.xml files —
the build depends on them. Backups: `..._2026-08-05b` +
`..._2026-08-06`. Verified renders, swapped, reopens fine.

---

## 2026-08-05 (evening) – Firm-color formulas, 26–33 edits, concrete check

**One-line summary.** Side-path pass #7. (1) Slides 22–25: assumption
callouts back to 18 pt (his edit; box 9.44/2.477, 3.15×0.71); firm
color-coding applied to ALL text-box formulas per Nico's rule —
**A / Q_A references dark red `C00000`, Q_B references dark green
`4D5D2C`** (accent3@50): callout "(Q_B = 50/20)" now green; reaction
equations mixed-color (Q_A red / Q_B green on 24, both directions on
25, incl. inside the gold equilibrium box, which he also resized to
3.12×0.73); slide-23 firm label at his (2.342, 2.793, w 2.553).
(2) Slide 26: "…different MC:" + his geometry. (3) Slide 27: OPEC
image → (6.09, 1.41, 3.98 w), pie card/chart/source → centered-left
(3.39/3.83 …) per his layout. (4) Slide 29: bullets box narrowed to
9.41 w, "price" in **bright red FF0000**; slide 30: "price" red +
underlined. (5) Slide 33: "Bertrand Equilibrium" label → (6.07,
4.916), arrow → (6.88, 5.17)→(6.96, 5.73). (6) **Concrete fact-check
(researched, sources in slide-31 notes):** US ready-mix is quoted per
CUBIC YARD (unit fixed on slide: "$50 per cubic yard"); realistic MC
today ≈ $115–125/yd³ vs. ~$180/yd³ price (NRMCA: materials alone
$89–96/yd³) — the $50 magnitude was KEPT because the PollEverywhere
answer bands (slide 32: <$50, $50, $51–55, …, >$100) anchor on it and
live in Nico's PollEv account. Recommendation pending: MC = $120 vs.
price ≈ $180 with re-created poll bands. Backups: `..._2026-08-05` +
`..._2026-08-05b`. Verified renders, swapped, reopens fine.

---

## 2026-08-05 (later) – Slides 22–25: FULL hand-edit re-port (colors!)

**One-line summary.** Nico flagged that pass #5 ported positions but
missed colors/arrow details. Did an exhaustive XML dump of slides
22–25 from `..._backup_2026-08-05.pptx` (every shape: geometry incl.
flipH/V, fill, line color/width/dash/ends, per-run text formatting,
OMML colors) plus side-by-side render comparison vs the backup. Ported
EVERYTHING: 22/23 — optimal-quantity x-ticks BOLD RED; assumption
callout rebuilt manually (white card, red 1.25 pt border, ALL-RED
14 pt text with a real Q_B subscript run); "units" labels at exact
positions; 23 firm label gets a WHITE FILL mask + thin blue
(accent1 4F81BD) pointer arrow to the demand line. 24 — reaction
points/guides/tick labels/RP labels in **accent6@75% ≈ B97034**
(sysDot guides), pointer ticks at his exact endpoints (flipV decoded:
arrowheads at the lower ends), x-title in **accent3@50% ≈ 4D5D2C**
olive. 25 — Firm B label + equation + x-title olive, 2.5 pt olive
pointer with OPEN 'arrow' head to B's line, Firm A's equation BLACK,
"Starting point… Q_A = 15" with a real subscript run, navy
equilibrium arrow at his exact endpoints (start hidden under the gold
box); staircase verified identical. Theme colors resolved from
theme1.xml (accent3 9BBB59, accent6 F79646; lumMod = channel scaling).
Verified via 1100-px side-by-side sheet, swapped over canonical.

---

## 2026-08-05 – Video-box v2 (gradient, corner default), 18–25 hand-edits

**One-line summary.** Side-path pass #5. Slide 18: portraits were
swapped — image27 (oval) is COURNOT, image26 (long-haired) is BERTRAND
(script fixed accordingly). Slide 19: header boxes → rounded + shadow.
Slide 20: hand layout ported (text block 0.656/1.72, 9.804×4.982 with
a 3-blank-line gap after bullet 1; Vale logo 2.48/2.243 and TSINGSHAN
card 5.466/2.243 side by side in the gap). Slides 22/23: hand label
positions ported (22: market label 3.358/3.0, firm 1.804/4.089 3-line,
MRA 2.996/6.427; 23: market 6.401/4.563, firm 2.028/2.783 SHORTENED to
2 lines "P = 80 − QA", MRA 4.135/6.438, new red pointer line
2.27/3.17→2.56/3.56); assumption callout = his grouped geometry
(9.28/2.33, 3.47×1.0); "units" labels just below the gap arrows.
Slide 24: y-title rotated at 1.033/3.315, x-title 5.45/6.16, reaction
label 7.62/4.525, RP labels + 2 gold pointer ticks, card 10.04×5.04.
Slide 25: titles/labels at hand positions, green pointer line to B's
curve, **equilibrium callout now GOLD rounded box + OMML overlay with
shade** (5.16/3.78, 3.35×0.98 — group in phase 3). **Practice-Video
box v2**: vertical gray gradient (bg1 lumMod 65→95→65, lin 90°,
scaled), gold border, DEFAULT position bottom-right (6.92/6.83,
5.85×0.58) drawn AFTER the footer so it overlays rules/page number;
applied on 24/35/40/45; slide 26 keeps centered large boxes (with
gradient). **Codified in Teaching CLAUDE.md** (Layout patterns →
"Practice-Video link box"). Gotcha: replacing the video-box helper by
slicing to the next "# ---" banner swallowed the adjacent
`_dash_shape_line` helper — restored. Backups: `..._2026-08-04c` +
`..._2026-08-05`. Verified renders, swapped, reopens fine.

---

## 2026-08-04 (night) – Cournot section polish: hand-edits 11–18, 19–27 rework

**One-line summary.** Side-path pass #4. Hand-edits ported (12: full
diagram re-geometry — wider/shallower LAC, flatter industry supply,
tick names P/q/Q_Cartel / _Comp / _Dev, subscript-run bullets with
"P_cartel = LMC"; 13: ADM logo to top-right corner + large Informant
still; 17: 26/40/24 pt runs, bullet char dropped; 18: portrait
geometry; 21: underlined "identical products"/"quantity" + shadow).
Slide 19: white rounded shadowed cards behind both bullet columns.
Slide 20: Vale logo + TSINGSHAN typographic lockup (no freely licensed
logo exists — site blocks fetches, nothing on Commons) + supply
equation moved to its own line. Slides 22–25 re-laid out to the
ORIGINAL geometry: near-full-width charts, market demand navy vs. firm
demand/MR in FIRM_A_RED, the "50/20 units" gap arrow between demand
curves, assumption callout top-right (red border) + red optimal-
response line (22/23); slide 24/25 with rotated red y-axis titles,
green x-axis titles, gold reaction points/guides (24), red/green
reaction functions + top-right callout boxes + arrowed equilibrium
label (25). **New deck-standard `_add_video_link_box`** (no Module 3
precedent existed): white rounded card, gold border 1.75, soft shadow,
gold ▶ glyph + navy bold label — applied on 24, 26×2, 35, 40, 45.
Slides 30/39 assumption cards got the same soft shadow as 21.
**Slide 27: native pie chart** (python-pptx PIE chart part, Carlito,
legend right, % labels, palette slice colors, no title): US crude
steel production shares 2024 — Nucor 26, Cleveland-Cliffs 20,
U.S. Steel (Nippon Steel) 13, Steel Dynamics 13, CMC 6, Others 22
(worldsteel totals + U.S. Steel 10-K segments; CMC approximate;
full sourcing in speaker notes). Backup kept: `..._2026-08-04c` (+
`..._2026-08-04b`). Verified renders (2 fix rounds), swapped over
canonical, reopens fine.

**Bash gotcha:** the harness's Bash layer eats one backslash level in
command strings — heredoc Python needing literal `\\n`-in-source must
use `chr(92)` construction or run from a Written .py file.

**Open:** propose mirroring slide-21's underline emphasis on slide 30
("identical products" / "price") — awaiting Nico.

---

## 2026-08-04 (evening) – PHASE 1 COMPLETE: all 87 slides built

**One-line summary.** Slide 11 restructured (outer white rounded frame
titled "Definition of Collusion" at Nico's hand-set 24 pt gold, navy
hero box nested inside). Then, on Nico's go-ahead, ALL remaining stub
slides were built — the deck now has real content on every slide
except the 5 PollEv + 3 video positions (phase-3 splices). Deck opens
in PowerPoint, 87 slides, 79/80 hidden, all renders verified through
three fix rounds. Backups: `..._2026-08-04.pptx` + `..._2026-08-04b`.

### What was built (batches A–F)

- **12 Cartels**: two-panel native diagram (SimpleFig panels; LAC =
  2-segment Bézier via new `_add_bezier_curve` custGeom helper; LMC,
  P1/PC price lines, pale-gold "Profit with cartel" + green "Extra
  Profit when deviating" regions, D/S/MR industry panel, OMML tick
  labels via `_fig_math_tick`, unicode-subscript bottom bullets).
- **13 ADM** (hyperlinked runs via `run.hyperlink.address`: lysine /
  Podcast / Movie / The Poop Cartel; Informant still + caption) ·
  **15** lysine chart (EMF→PNG via PowerPoint COM insert-export,
  `_source_images/_adm_lysine_chart.png`, AEA source line).
- **29–36 Bertrand**: bullets/assumptions; 31 concrete-bid poll intro
  (Poll Break badge via `_add_discussion_break(text="Poll Break")`);
  33 undercutting chart (red Lost-by-A strip stylized taller at y
  35–40 so it stays visible, green Gained-by-A); 34 three-equilibria
  comparison chart; 35 worked math (OMML, answers in deep red);
  36 overview two-column.
- **38–46 differentiated goods**: firm colors A=red `#C00000`,
  B=green `#1B5E20` (from source); 40 setup with colored OMML demands
  QA=400−2PA+PB etc. + "substitutes" annotation; 41/42 parametrized
  `_airbus_reaction_slide` (PB=200→PA=250; PB=400→PA=300, previous
  demand in red); 43 price-reaction equilibrium (267,267) with
  staircase from PA=150; 44 BruinLearn button; 45 take-aways; 46
  discussion.
- **48–53 game theory**: 48 Nash photo + Nobel medal + Beautiful Mind
  poster; 49 warm-up (Zidane); 50 beauty contest (Keynes, underlined
  "you"); 51 numerical + Poll badge; 53 key concepts (concept-blue
  bold for Best response / Dominant strategy / Nash equilibrium).
- **55–84 classic games**: PD sequence 55–58 (`_pd_matrix`; dim='row'/
  'col' fades the other player's payoffs; `_br_arrow` + `_br_circle`
  best-response marks; Nash oval + callout finale); 59 ways to solve;
  60 podcast; 61 fair/dirty play; 63 commons (image46); 64 chicken
  (2 NE, Rebel-Without-a-Cause link); 65 M.A.D. (image48=tall Titan
  launch LEFT, image47=wide truck RIGHT — names misleading!);
  66 group work (EMPTY matrix cells, Strangelove + Documentary
  links, Discussion badge); 67 commitment (red-X impossible cells);
  68/70 penalty kicks (no pure NE); 71/73 Coke-Pepsi; 75 Split or
  Steal (£55/£5/£100/£0, 2 NE); 79 hidden Steven-commits (X'd row);
  80 hidden Nick/Ibrahim (£60*, p×£50 + footnote); 81–84 entry game
  ((0,1) NE solid + (1,0) NE dashed oval on 82; subsidy variants).
- **85–87**: two-column recap (18/16/14 sizing to fit); StubHub
  screenshots; 87 has the question line + navy "← Back" pill with
  `ppaction://hlinkshowjump?jump=lastslideviewed`.

### Render-fix rounds (what was wrong; all fixed + verified)

Round 1 (15 fixes): s12 bottom bullets overlapped footer → panels
shrunk + `_add_bulleted_list`; s15 EMF image overflowed (aspect!) →
height-constrained + centered; s30 box widened; s31 image collision;
s42 demand label capped inside plot (`min(dem_int*1.42, 440)`);
s43 equilibrium callout moved right of chart; s61/71/73 interpretation
boxes to bottom-left (rotated row-player name collided); s63 commons
image height-capped; s65 image48 width-capped; s80 matrix below
bullets; s81 photos swapped (image58=777 top); s85 left column
18/16/14; s87 screenshot shrunk. Round 2 (5): s12 LAC label lift;
s31 mixer/tug are image33/image32 (SWAPPED names!); s43 callout 14 pt;
s65 photo swap (see above); s87 more shrink. Round 3: s87 final nudge.

### Phase-3 backlog (unchanged)

Poll splices (32, 52, 69, 72, 77), video splices (14, 74, 76+78 with
NULL rels), box/text + table/backing grouping pass, animations,
trailing-period sweep, steel-pie decision (slide 27, still awaiting
Nico), teleprompter notes later.

---

## 2026-08-04 (later) – Slide-6 ribbons restyled; 7–9 hand-edits; slide-11 cartel rewording

**One-line summary.** Side-path pass #3. Slide 6: the (Price Takers)/
(Price Searchers) orange ribbons are now single rounded shadowed boxes
(`_add_rounded_filled_box`, corner 0.35, white italic 24) at the same
positions; the oval-capped zone lines removed (helper `_capped_line`
kept, unused). Hand-edits ported from slides 7–9: red `#C00000`
"interdependent" run in the slide-7 table cell (per-run rewrite after
`_add_styled_table`); slide-8 bullets bumped 24/22 → **28/24** plus
italic "potential" (multi-run bullet form); slide-9 note card narrowed
and lowered to (0.276, 6.554, 10.965 × 0.470) — decoded from his
scaled group. Slide 11 reworded per Nico (verified: a **cartel** is
the extreme, explicit/formalized case of collusion — collusion itself
can be tacit): hero definition now ends at "…pricing decisions"; cream
box gains third bullet "Extreme case – a cartel: an explicit,
organized price-fixing agreement (e.g., OPEC)" (measured to fit one
line via PIL/Calibri); label nudged to 1.55; quote box moved to 5.45.
**Teaching CLAUDE.md updated** (Build mechanics gotchas): OMML/
AlternateContent phantom-deletion diff artifact; `~$` lock-file check
before COM (never kill POWERPNT while Nico works); scaled-group
decode + notes/table-format diff caveats; Wikimedia logo-sourcing
convention. Backups: `..._2026-08-03b` + `..._2026-08-04` (two
newest). Swapped over canonical, reopens fine.

---

## 2026-08-04 – Slide 5 red circle, slide-6 hand-edit port v2, slide-8 logos

**One-line summary.** Side-path pass #2. Slide 5: gold Oligopoly header
cell reverted to navy; red oval (`#C00000`, 2.5 pt) around the header
like the original slide. Slide 6: Nico's second hand-edit round ported
verbatim from the canonical XML — 28 pt Least/Most labels; (Price
Takers)/(Price Searchers) as white-italic-24 text on **accent6 orange
(`#ED7D31`) ribbon fills** with oval-capped orange zone lines
(`_capped_line` helper); slim red arrow (5.200/1.770, 2.350×0.250) +
"WE ARE HERE" 16 pt bold red label + tightened red box
(7.535/1.620, 1.675×5.050); adidas → **"Restaurants in WeHo" map**
(`_source_images/_s6_image6.png`, rendered 3.988/4.502, 2.387×2.267)
with 16 pt black caption; hand-tuned logo positions ported. (His
groups → ungrouped shapes at identical rendered coords; grouping is a
phase-3 pass anyway. Group chOff/chExt scale decoded for the map.)
Slide 8: bars now brand-colored (Verizon `#EE0000`, T-Mobile
`#E20074`, AT&T `#00A8E0`) each with a white rounded chip carrying the
company logo; mergers callout rebuilt as a manual cream card with two
logo lockups (T + Sprint, T + UScellular). Logos downloaded from
Wikimedia Commons / Wikipedia to `_source_images/_logo_*.png`
(verizon, tmobile [magenta T square], tmobile2 [old wordmark, unused],
att, sprint, uscellular). Backups: `..._2026-08-03.pptx` and
`..._2026-08-03b.pptx` (two newest kept). Test build verified,
swapped over canonical, reopens fine.

---

## 2026-08-03 – Hand-edit port + slide-3 arrow + slide-6 full redo

**One-line summary.** Side-path pass: ported Nico's PowerPoint hand-edits
into `_build_Module7.py` (slide 1: Dogopoly cartoon at exact EMUs
8248214/2189988, 3871436×4267181, text block moved up/left to
L=-168626, tops 1152144/2203704/3529584/4169664, gold strip
3846422/3163824; slide 5: bottom italic note deleted), made the 3→4
roadmap arrow gray (FADED 3.0, was navy 3.5), and REBUILT slide 6 to
mirror the original: double-headed navy spectrum arrow, RED
(`#C00000`) block arrow pointing into a red 3.5 pt box around the
whole Oligopoly column, red divider bars between the four structures,
and separate double-headed arrows under (Price Takers) / (Price
Searchers). New helper `_double_arrow` (headEnd inserted BEFORE
tailEnd — schema order). Backup kept:
`Module 7 - Revised_backup_2026-08-03.pptx`. Test build verified,
`mv -f` over canonical, reopens fine (87 slides).

### GOTCHA (important for every future hand-edit diff)

**After PowerPoint re-saves the deck, every textbox containing OMML
math gets wrapped in `mc:AlternateContent` — and python-pptx no longer
enumerates it as a shape.** A python-pptx shape diff then reports all
OMML labels/equations as "deleted" (slides 20, 22–25 showed phantom
deletions of the mixed-OMML boxes). They are still present and render
fine. When diffing a PowerPoint-saved canonical against a script
build, treat missing OMML textboxes as artifacts — confirm hand-edits
via renders or raw slide XML, not python-pptx enumeration alone.
Also: notes diff (all 87 slides) showed no hand-edited notes.

---

## 2026-07-29 (later still) – Batch 2: nickel decision, Cournot charts, trade war

**One-line summary.** Nico's decisions: nickel duopoly = **Vale +
Tsingshan** (Norilsk excluded with a sanctions note at the bottom of
slide 9); trade-war slide keeps the original stylized payoffs with
US/China labels and calibration details in the speaker notes. Built
slides 9, 20, 22–27, 62; deck verified in PowerPoint, renders checked
(3 layout fixes: slide-24 axis title vs. video box, slide-25 callout
wrap → 16 pt one-liner, slide-27 image spacing).

### New helpers in `_build_Module7.py`

- `SimpleFig(left, bottom, w, h, xmax, ymax)` – logical→EMU transform
  for native shape charts; `_fig_axes` (navy arrow axes), `_fig_xtick` /
  `_fig_ytick`, `_fig_point`. OMML shortcuts `_oQ/_oP/_oMR(sub)`.
- `_add_payoff_matrix(...)` – game-theory-convention 2×2 matrix: col
  player on top (gold), row player rotated 270° at left (concept blue
  `ROW_BLUE #0070C0`), single-shape cells (payoff runs live INSIDE the
  bordered rect — no box+text grouping needed), `cell_texts` override
  for "impossible cell" text, `nash_cells` gold ovals, colored caption
  "Payoffs to (Row, Col)". Returns anchors dict
  `{(r,c,'row'|'col'|'cell'): (x,y)}` for best-response overlays;
  `_br_circle` draws the circles (arrows still to add on PD slides).
- `_cournot_reaction_slide(...)` – parametrized builder for slides
  22/23 (residual demand, MR_A dashed gold, MC gray, dashed guide
  drops, gold action box with the optimal response).

### Content notes

- Slide 9: bullets updated (Vale / Tsingshan Holding – mines mainly in
  Indonesia); cream note "Norilsk Nickel (Russia)… excluded because of
  the sanctions against Russia"; notes carry USGS/Tsingshan facts.
  "Crucial for batteries" softened to "Key input for batteries" (LFP
  shift) — reported to Nico.
- Slide 62: payoffs 20,20/13,24/24,13/15,15 kept; players US (blue) /
  China (gold); PIIE tariff line on the slide (≈48% vs ≈32%); speaker
  notes carry the full calibration (IMF GDP $30.6T/$19.4T, PIIE
  tracker Nov 2025, Yale Budget Lab / Tax Foundation / EIU cost
  estimates). Source grammar typo fixed: "set a high tariffs" → "set
  high tariffs".
- Slide 27: steel market-share pie kept as source image for now —
  **flagged**: data outdated (AK Steel, ArcelorMittal USA absorbed by
  Cleveland-Cliffs 2020; U.S. Steel → Nippon Steel 2025). Proposed
  native rebuild with current data, awaiting Nico.

### Open for next session

- Batch 3 candidates: 12 (cartel two-panel diagram), 13 (ADM case,
  external hyperlinks), 15 (EMF→PNG via PowerPoint COM), Bertrand
  section 29–31, 33–36.
- Steel-pie decision from Nico (slide 27).
- Then: differentiated goods 38–46, game theory 48–61, 63–87; phase-3
  splices + grouping + animations at the end.

---

## 2026-07-29 (later) – Phase-1 scaffold built: batch 1 (17 slides + stubs)

**One-line summary.** All five outline decisions locked (dividers in
Module 3 style; Stackelberg dropped; all three examples to be updated;
EMF → PNG; 87-slide 1:1 plan). Assembled `_build_Module7.py` and built
`Module 7 - Revised.pptx` (87 slides): batch 1 = slides 1–8, 11,
17–19, 21 + the six dividers; the rest are positional stubs. Deck
opens cleanly in PowerPoint; renders verified by eye.

### How the script was assembled

- `_build_template_samples.py` copied from Module 3 into Module 7 with
  M7 `FOOTER_TEXT` and a new `MODULE_AGENDA` (Part 1 Oligopoly: Collusion
  and Cartels / Cournot / Bertrand / Differentiated Goods; Part 2 Game
  Theory: Strategic Thinking and Key Concepts / Prisoner's Dilemma…).
- `_build_Module7.py` = new header + Module 3 helper layer copied
  VERBATIM (`_build_Module3.py` lines 33–75, 77–472, 503–2313 — skips
  the M3 title slide and M3-specific formula shortcuts) + M7 appendix.
- **M7 additions:** `_draw_footer` override emitting a LIVE
  `<a:fld type="slidenum">` page number (deterministic uuid5 GUIDs);
  `_add_media_image` (loads `_source_images/imageNN.ext` by media name);
  `_add_styled_table` + `_set_cell_borders` (navy header, white/cream
  body, thin borders, backing-card shadow — **backing + table still need
  grouping in phase 3**); `make_module_outline` divider (cream rounded
  band + gold border highlights the current sub; navy/faded text);
  `make_stub` (holds position; poll/video stubs marked PHASE 3;
  slides 79/80 stubs carry `show="0"` = hidden).

### Batch-1 slides built

1 title · 2 recap M6 (Netflix typo fixed) · 3 course roadmap (part 4
navy + "we are here") · 4 full outline · 5 market-structures native
table (gold Oligopoly header cell) · 6 market-power spectrum (arrow +
4 stops, source logos flat) · 7 characteristics native table ·
8 concentration (UPDATED example: top-3 ≈ 97%, native gold bar chart
146/140/119M, "Mergers: T-Mobile absorbed Sprint (2020) and UScellular
(2025)" callout; notes rewritten accordingly) · 11 collusion (navy hero
definition + cream aim box + gold-bordered Adam Smith quote) · 17 Note
convention callout · 18 Cournot/Bertrand portraits + model boxes ·
19 two-column comparison · 21 Cournot assumptions card · dividers 10,
16, 28, 37, 47, 54.

### Research results (subagent, 2026-07-29, sources in chat)

- **Wireless:** Verizon ~146M / T-Mobile ~140M / AT&T ~119M retail subs
  late 2025; top 3 ≈ 96–98%. T-Mobile closed UScellular deal
  2025-08-01 (~$4.3B). (Slide 8 already updated.)
- **Nickel:** Indonesia ~2/3 of 2025 mine output (USGS via secondary);
  Tsingshan ~30% (1.12 Mt 2023); Vale ~160kt, Norilsk ~205kt — the old
  Vale/Norilsk "duopoly" no longer holds. LFP (nickel-free) batteries
  >50% of EV deployment in 2025. **Decision pending** (slides 9/20).
- **Trade war:** PIIE Nov 2025: US avg tariff on China 47.5%, China on
  US 31.9%. GDP 2025: US ~$30.6T, China ~$19.4T. Cost estimates ~0.5%
  of GDP order of magnitude each. **Framing proposal pending** (slide 62).

### Open for next session

- Nico's answers: nickel example (option A keep-as-stylized vs. B switch
  industry); trade-war payoff framing (recommend: keep original stylized
  payoffs, relabel US/China, add real-magnitude bullet).
- Batch 2 candidates: 9, 12 (cartel two-panel diagram), 13, 15 (EMF→PNG
  via PowerPoint COM), 20, 22–27 (Cournot charts), 29–31, 33–36.
- Phase 3 backlog: poll/video splices, table/backing grouping, animations.

### Commands

```powershell
# rebuild in place (close PowerPoint first)
$env:PYTHONIOENCODING = "utf-8"
python "_build_Module7.py"          # from the Module 7 folder
```

---

## 2026-07-29 – Kickoff: source-deck inventory + revised outline proposal

**One-line summary.** First session on Module 7. Extracted a full
slide-by-slide inventory of the old deck, extracted all media, and
drafted the reformat outline (`Module 7 - outline.md`) for Nico's
review. No .pptx built yet — awaiting answers to the outline's five
open questions.

### Source-deck facts (Module 7.pptx)

- **87 slides**, 4:3 (10 × 7.5"), **4 slide masters** → full conversion
  to widescreen single-master template needed.
- Content: Part 1 Oligopoly (market structures, collusion/cartels/ADM,
  Cournot, Bertrand, differentiated goods Airbus/Boeing), Part 2 Game
  Theory (concepts, prisoner's dilemma, chicken/M.A.D., penalty kicks,
  Split-or-Steal, Boeing–Airbus entry game), recap + 2 ticket-pricing
  backup slides.
- **5 PollEverywhere slides** (32, 52, 69, 72, 77), **3 video slides**:
  14 (`media1.mp4`, ADM FBI footage), 74 (`media2.mp4`, Split or Steal),
  76 + 78 (**both** embed `media3.mp4` via the MS media rel, with the
  legacy video rel `Target="NULL" TargetMode="External"` — the exact
  pattern python-pptx strips; phase-3 splice mandatory).
- **2 hidden slides** (79, 80) — rebuild and keep hidden.
- Slide 15 is a 2.8 MB **EMF** chart (lysine prices, AEA source).
- 6 recurring "OUTLINE of Module 7" slides (10, 16, 28, 37, 47, 54) +
  the master outline at 4.
- **Inconsistency flagged:** outline slides 47/54 list "Stackelberg
  Competition" but the deck has no Stackelberg content (and outline
  slides 4/10/16/28/37 don't list it).

### Files created

- `_source_inventory.md` — auto-extracted per-slide inventory (titles,
  body text, notes, pictures, videos, links, hidden flags), display
  order resolved via `presentation.xml` sldIdLst (not part filenames).
- `_source_images/` — all 70 media files extracted (incl. 3 mp4s).
- `Module 7 - outline.md` — the reformat proposal: 87 → 87 slides 1:1,
  dividers replace outline slides in place; tags [KEEP]/[NATIVE]/
  [DIVIDER]/[POLL]/[VIDEO]/[HIDDEN]; global treatments; build plan
  (phase 1 `_build_Module7.py` → freeze → phase 3 surgery).

### Open questions (in the outline, awaiting Nico)

1. Divider style (highlighted-outline divider à la Module 3 vs. concept
   map).
2. Drop the phantom "Stackelberg Competition" outline line?
3. Example refreshes: T-Mobile/Sprint 2020 merger; Vale/Norilsk nickel
   duopoly (market now Indonesia-dominated); trade-war payoff slide.
4. Slide 15 EMF: native rebuild vs. PNG conversion.
5. Confirm 1:1 87-slide plan.

### Next steps

- Get answers → adjust outline → write `_build_Module7.py` (phase 1),
  reusing Module 3's helper layer (`Module 3/_build_Module3.py`:
  palette, chrome, bullets, boxes, Fig/eqn engines, payoff-matrix
  conventions from Teaching CLAUDE.md).
- No typo fixes applied to the source deck; "Neflix" → "Netflix" will
  be fixed in the rebuild (obvious-spelling exception, will report).
