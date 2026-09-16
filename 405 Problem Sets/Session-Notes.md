# Session Notes — 405 Problem Sets

## 2026-09-06 — Problem Set 1 and its solutions brought to the new class format

Rebuilt `Problem Set 1.docx` and `Problem Set 1 -- Solutions.docx` in the
Fall 2026 course format, and checked both against the revised Module 1 and
Module 2 decks. PS 1 is due in Week 4 (Course Calendar), so it draws on
Modules 1 and 2 only — both of which are revised, so the whole set could be
checked.

### Files

| File | Role |
|---|---|
| `_ps_theme.py` | Problem-set layer: masthead, problem headings with the gold points chip, `(a)` part labels, and the `Panel` supply-and-demand figure primitives |
| `_build_PS1.py` | Builds `Problem Set 1.docx` |
| `_build_PS1_Solutions.py` | Builds `Problem Set 1 -- Solutions.docx` |
| `_originals/` | The 2025 source documents — BUILD INPUTS, never delete |

The engine is **`_tn_theme.py`** in
`405 Slide Revisions 2026/Teaching Notes/`, imported by path from
`_ps_theme.py`. Nothing in `_tn_theme.py` was modified, so the six teaching
notes are unaffected. If that folder moves, update `_THEME_DIR` in
`_ps_theme.py` — it raises `ImportError` with the path it looked in.

**This folder is untracked in git.** `_originals/` plus the rolling
`_t-1` / `_t-2` files are the only way back from a bad rebuild. Rebuilds
themselves are safe — the scripts are the source of truth.

### Decisions

- **Tracked changes are the review mechanism.** The restyle is untracked;
  anything beyond formatting is a real Word revision (author
  "Claude (proposed)"). Eight revisions in each document. Once Nico accepts
  or rejects a round, fold the accepted text into the build script as plain
  runs and note it in the module docstring — the same convention as the
  teaching notes.
- **Figures are native Word shapes, not screenshots.** The originals were
  matplotlib exports with gridlines and legends. All seven are rebuilt with
  the `Panel` class: demand dark red `C00000`, supply navy, no gridlines, no
  legend, labels inside the plot in boxes measured with PIL.
- **Every equilibrium is a computed intersection** (`_ps_theme.cross`), so a
  dot always sits exactly where its two lines meet.
- **Problem 3(c) went from a three-across row to three full-width stacked
  panels**, each under the bullet it illustrates. At a third of the page the
  curve labels, the E₀ / E₁ pair and the axis titles all had to share one
  corner.
- **Arrows use a `custGeom` path, never `Fig.line(arrow=True)`.** A doubly
  flipped (up-and-left) line puts the arrowhead on the wrong end in Word.
  `Panel.arrow` addresses both endpoints inside an unflipped box.
- **A "Draws on:" line under each problem heading** names the module and
  outline item, verified against the decks. This is part of the format
  layer, so it is untracked — flagged to Nico as vetoable.
- The documents carry **no term, year, or due date** (the calendar is
  referenced instead), so the same files are reusable next year — the same
  reasoning behind the teaching notes' bare page number.

### Checked against the decks — findings

Consistent: explicit / implicit / full economic cost (M1 slides 67, 74);
market-clearing terminology (M1 slide 30); the both-curves-shift ambiguity
(M1 slides 33 – 34, matching Problem 3(d) exactly); income elasticity with
normal / luxury / inferior (M2 slide 55); the opportunity-cost framing of
Problem 4.

Proposed as tracked changes — see the build-script docstrings for the full
reasoning.

### 2026-09-06, same session — Problem 1's numbers brought to 2026

Nico asked for the numbers to be updated. Only the two components with a
**source** moved, so the exercise does not acquire figures that merely look
precise:

| | 2025 | 2026 | Source |
|---|---|---|---|
| Forgone salary | $40,000 | **$46,000** | BLS OEWS May 2025, SOC 37-3011 (Landscaping and Groundskeeping Workers): California median $45,560, national $39,150 |
| Gas | $5.00/gal → $5,000 | **$5.80/gal → $5,800** | AAA, California ≈ $5.80 in early Sept 2026, highest of any state |
| Truck | $40,000 | **unchanged** | Still right — a 2026 Ford F-150 XL starts at about $40,085 |
| Revenues | $70,000 | **$78,000** | Set to hold the original's narrow margin |

Tools, repairs, mileage and permits are unsourced assumptions the problem
invites students to make, so they stay put. Totals follow: explicit
$16,800, implicit $46,000, total cost $62,800, economic profit **$15,200**.
The margin stays narrow, which is what makes the exercise worth doing.

This also fixed a collision the first round introduced: the bracketed note
had been changed to illustrate averaging with a **lawn mower** at $3,000
over 3 years, which contradicted the solution's Tools row (lawn mower among
tools lasting ~1 year). The note now uses the solution's own **truck**
($40,000 over 10 years = $4,000), so nothing collides.

Revisions now: **14** in the problem set, **26** in the solutions.

### Open items

1. **The AI-tools policy** allows AI "to brainstorm or to revise existing
   work you have written", while Problem 3(b) says "You may use AI tools to
   produce graphs" and Module 1 slides 40 / 45 encourage AI for studying.
   Not a contradiction, but the policy paragraph could be aligned with the
   deck. Policy language, so untouched.
2. ~~Module 1's deck still draws demand curves in navy.~~ **Done
   2026-09-06** — Nico asked for it, and Module 1 was swept to the dark-red
   demand rule the same session (`_sweep_demand_color.py`; see that
   folder's Session-Notes). The problem sets and the Module 1 deck now
   agree. **Modules 2 and 3 have not been checked against the rule.**
3. **Problem Sets 2 – 5 are still in the old format.** Modules 3 and 4 are
   revised, so PS 2 can be checked the same way; PS 3 – 5 reach into
   Modules 5 – 8, which are not.
4. Neither document has been committed — the folder is untracked in git.

### Commands

```
python _build_PS1.py
python _build_PS1_Solutions.py
```

Render to check (no LibreOffice on this machine — drive Word via COM):

```powershell
$w = New-Object -ComObject Word.Application; $w.Visible = $false
$d = $w.Documents.Open("<path>.docx", $false, $true)
$d.ExportAsFixedFormat("<out>.pdf", 17)
Write-Output $d.Revisions.Count
$d.Close($false); $w.Quit()
```

---

## 2026-09-07 — Problem Sets 2 – 5 and their solutions brought to the new format

Same treatment as PS 1: the eight documents are rebuilt from
`_build_PS<n>.py` / `_build_PS<n>_Solutions.py`, all on `_ps_theme.py`
over the teaching-note engine. The 2025 originals are in `_originals/`
as build inputs. **PS 1 was deliberately left untouched** — it is
committed, Nico may be mid-review, and none of its 26 revisions had been
accepted or rejected yet.

### What each document draws on, and what could be checked

| Set | Modules | Deck check |
|---|---|---|
| PS 2 | Module 2 | Full — Module 2 is revised |
| PS 3 | Modules 3, 4 | Full — both revised |
| PS 4 | Modules 5, 6 | Partial — 5 and 6 are NOT revised; only the Module 4 welfare apparatus and Module 7's price-discrimination recap could be checked |
| PS 5 | Modules 6, 7 | Partial — Module 7 is revised (Problems 3 and 4), Module 6 is not |

### Errors found in the solution keys

1. **PS 3, Problem 3(c): a stale TR column.** The $5.50 table still
   carried the $6.40 revenues (6,400 / 12,800 / 19,200 …) while the
   Profit column was computed correctly at $5.50, so the two contradicted
   each other on every row. Both cost tables are now COMPUTED from
   (Output, Total Cost) and the price, so a column cannot go stale again.
2. **PS 5, Problem 2: header said 25 points**, but the problem set awards
   30 and the parts sum to 30.
3. **PS 3 totals 110 points**, not the stated 100 (15+15+25+30+25).
4. **PS 5's Problem 3 totals 40 points**, not the stated 35, so that set
   totals 105. Both flagged in the header rather than rescaled — which
   problem gives up the points is Nico's call.
5. **PS 5, Problem 3(h)** called the one-firm scenario "(f)" in its
   figure and closing sentence; the one-firm case is part (g).
6. **PS 2, Problem 4** read "Boeing's new aircraft AS the following
   demand"; **PS 2, 5(c)** had "Q = 58,32" with a decimal comma.

### Terminology and notation brought in line with the decks

- **"the firm should operate" → "should continue to produce"** (PS 3).
  Module 4 slide 32 keeps the short-run verb apart from "shut down".
- **`r` → `pK` for the price of capital** (PS 3). Module 3 slide 43
  glosses it "pK : Price of Capital".
- **Capital Π → lower-case π** for profit throughout (2026-08-30 rule).
- **PS 3, Problem 2(a)** answered the question twice in near-identical
  paragraphs (one with an unclosed parenthesis) and a third time under
  (b); consolidated.

### Figure conventions — the one that caught me out

Module 4's price-taker panel states the rule outright in a comment:
*"the price line is the demand curve the firm faces, so it takes the
deck's dark red, label and axis value with it"*. So on a cost panel:

| | |
|---|---|
| MC | **navy** |
| ATC | **gold** |
| P = MR (the price taker's demand) | **dark red**, label and y-tick too |
| MR for a price SETTER | concept blue `0070C0` |

I had MC and MR the wrong way round in the PS 3 Balding panels on the
first pass and fixed them. Also adopted: profit fills dark red, a loss
the firm produces through GRAY (the original had 4(b)'s loss in red and
4(e)'s profit in green — the convention backwards); CS red `C0201B` at
26 %, PS blue `4E79B5` at 34 %, each label at its region's **computed**
centroid; Module 7's payoff-matrix format for PS 5 Problem 4.

**Two label traps worth remembering.** With a linear demand curve the MR
line passes EXACTLY through the consumer-surplus centroid (the centroid
is at Q*/3, and MR there equals the centroid's own height), so the CS
label has to be offset on any panel that draws MR. And an axis tick
sitting at the very END of an axis lands on the axis title, which is
anchored to the arrow tip — give the axis headroom past the last tick.

### New in `_ps_theme.py`

`Panel.polyline` (a tabulated series as one editable freeform),
`Panel.polygon` + `centroid` (surplus regions built from the curves' own
points), `Panel.region` (profit / loss rectangles), `Panel.legend` (for
charts whose series converge), `payoff_matrix`, `quote`, `link`,
`cell_runs` / `sub_runs` / `sub_inline` (real subscripts in tables and
running text), and `h` / `italic` / `color` on the tick helpers.

### Open items

1. **PS 3 and PS 5 point totals** — 110 and 105. Flagged, not rescaled.
2. **PS 4 Problem 5 keeps two O*NET screenshots as images**
   (`_source_images/onet_search.png`, `onet_tasks.png`) — they are
   screenshots of a live website, which the standing rule leaves alone.
   BUILD INPUTS, never delete.
3. **Modules 5, 6 and 8 are not revised**, so PS 4 and the Module 6 parts
   of PS 5 could not be checked against a deck.
4. **PS 2's Problem 4(c) figure keeps Q on the vertical axis.** That
   looked wrong at first but it is the class convention for an ESTIMATED
   demand function — Module 2 slide 101 says "a = intercept with the
   y-axis (Q on y-axis)".
5. Nothing is committed.

### Commands

```
for n in 1 2 3 4 5; do python _build_PS$n.py; python _build_PS${n}_Solutions.py; done
```

## 2026-09-08 — Nico's hand edits to PS 1 adopted; byline added; two layout bugs found

Nico hand-edited `Problem Set 1.docx` and `Problem Set 1 -- Solutions.docx`
in Word (accepting the whole tracked round in the process). This session
ported those edits back into the build scripts, added the byline to both
documents, and fixed two systematic layout faults that his edits exposed.

His versions are preserved as `Problem Set 1_t-1.docx` and
`Problem Set 1 -- Solutions_t-1.docx`.

### What was adopted

- **All tracked changes accepted.** `_build_PS1.py` and
  `_build_PS1_Solutions.py` no longer emit any `ins` / `dele` /
  `para_inserted`; both documents now report 0 revisions. PS 2–5 keep their
  pending proposals (1–4 revisions each).
- **His rewritten AI-tools policy card**, reproduced byte-identically as
  `S.policy_card(doc)` in `_ps_theme.py` — a centred dark-red italic heading
  "Important information, please read carefully" over a justified navy body.
- **"the Class Website"** in the due-date line, his Problem 1 note and his
  Problem 3(b) note, verbatim.
- **His figure label positions**, as exact anchors in the build script:
  the 2(c) `Dₐ` label and the gold shift arrow, the four 3(b) curve labels,
  the 3(c) `D` / `D′` pairs, and the fig-4 base panel.

### Byline

`ps_masthead(..., byline=BYLINE)` puts **Prof. Nico Voigtländer** on the
subtitle line, right-aligned on a 9360-twip tab stop, gray 11 pt, above the
gold rule. Present on the problem set and the solutions.

### Two bugs his edits exposed

1. **Double alignment shift in the label helpers.** `Panel.tick_y`,
   `tick_x`, `title` and `text` pre-shifted x AND passed a non-left `align`,
   so `Fig.label` shifted a second time. Diagnosed by measuring
   `text_width(("w","0")) + 0.08 = 0.2380` → x should be 0.3020; the build
   had 0.064, and Nico had dragged the label to exactly 0.302. All four
   helpers now pass the raw anchor. The panel titles and the ticks he never
   touched are the only remaining figure differences against his file — his
   are the un-fixed positions, mine are now correct.
2. **`compatibilityMode 14`.** python-docx's template declares Word's 2010
   layout engine; Word 2013+ (mode 15, which is what Nico's saved file has)
   breaks lines differently, so identical text wrapped one line longer in
   the build. `_ps_theme.save(doc, path)` now patches
   `word/settings.xml` to mode 15 after saving, and all ten build scripts
   call it instead of `doc.save`.

### Figure crop

`place(f, doc, ...)` crops the group to `content_box(f)` — the real bounding
box read back out of the emitted shape XML — at 1:1 scale, so nothing inside
moves. `pad` is now **0.0**: measured against Nico's file, his boxes are the
ink bounding box exactly, and the 0.02" pad an earlier version added made
every figure 0.040" taller, which accumulated into a seventh page in the
solutions. With both fixes the solutions paginate at 6 pages, as his do.

### Verified

Compared page by page against a render of his files: the 2(c) panel, the
3(b) four-curve figure, both 3(c) case panels and both fig-4 panels match on
every label position; page 1 of the problem set matches on the byline, the
policy card and the 2026 numbers; page counts agree.

### Pending

- PS 2–5 figures have not been spot-checked by eye since the tick fix, the
  E-label centring and the crop. The geometry is generated by the same
  helpers, so they should be right, but they have not been looked at.
- PDFs were deliberately not regenerated (Nico does that himself).

### Commands worth remembering

- `python _build_PS1_Solutions.py` — rebuilds in place; no backup roll
  needed, the script is the source of truth.
- Rendering: **one document per PowerShell process.** Looping several
  `ExportAsFixedFormat` calls inside one process hangs Word indefinitely.
  Never name a PowerShell parameter `$Doc` while using `$doc` for the
  Document object — PowerShell variables are case-insensitive and the
  parameter shadows it.

## 2026-09-08 (later) — Header swapped, policy card resized, and a second
## round of figure-label edits that exposed the real fault

Three rounds of Nico's edits in one evening. The last one is the
instructive part.

### Header

`ps_masthead` now leads with the document naming ITSELF: "Problem Set N"
centred, 20 pt bold navy, on line 1; "MGMT 405 – Managerial Economics" at
13 pt on line 2 with **Prof. Nico Voigtländer** bold gray 11 pt right-
aligned on the tab stop, gold rule under it. Applies to all ten documents
from the one helper.

### Policy card

`policy_card(..., height_in=1.9375)` — the height he set by hand, down from
the 2.197" `cream_card` measured for itself. `_measure_par` is conservative
and assumed more wrapped lines than Word sets at that width, so the card
carried a band of empty cream. His file's `adj` was still 5005, the value
computed for the old height, so his corners had shrunk to 0.097";
`cream_card` recomputes it to 5677 and restores the standard 0.11" radius.

### The figure-label round, and why it was missed

He moved the `D`ₐ label in the p.2 panel and all four equilibrium labels in
the p.3 panel — **a second time**, after the rules had already been written
down. Cause:

- Round 1's positions had gone into the build script as **literal
  coordinates** (`pn.text(63.53, 23.64, …)`, `pn.text(e[0]+3.0, e[1]-6.0,
  …)`). A literal anchor cannot follow a rule.
- The four 3(b) equilibrium labels never called `Panel.equilibrium`'s label
  path, which is the only place the centred-above rule lives. "Enforced in
  the helpers" was true and vacuous.
- The port was verified by comparing renders **against his round-1 file**,
  which can only find where I differ from him — never where we both differ
  from a rule just stated.

Fixed by routing every label through the helper that owns its rule:

| Figure | Was | Now |
|---|---|---|
| p.2 `D`ₐ | `text()` at (63.53, 23.64), mid-curve | `demand(label=("D","a"), lbl_dx=-0.17, lbl_dy=-0.31)` |
| p.3 `L` `N` `E₀` `I` | `text()` at `e + (3, -6)`, down-RIGHT of the dot | `equilibrium(label=eq)` — centred just above |
| p.3 `D`ₗ … `D`ᵢ | `text()` at four literal anchors | `demand(lbl_dx=-0.19, lbl_dy=-0.30)` |
| p.4–5 `D` / `D′` ×3 | `lbl_at` table of six anchors | `demand(lbl_dx=-0.13, lbl_dy=-0.30)` |

Ten magic numbers gone. Audited against his file: every label within
**0.069"** (worst: `I` on p.3), which is mouse-drag scatter — and now every
one of them moves with its curve.

The new rules are in `CLAUDE.md` under "Every label goes through the helper
that owns its rule": derive the offset, never hard-code the coordinate he
dragged to; audit a port against the RULES, not only against his file;
~0.07" is a match, 0.4" means the label was never under the rule.

### Pending

- ~~`Problem Set 1 -- Solutions.docx` not yet rewritten~~ — DONE once he
  closed Word. Rebuilt and verified against his hand-edited `_t-1`: same
  shape count in all seven figures, 0 revisions, compatibilityMode 15,
  worst label deviation 0.069" (`I` on p.3), all of it rule-derived.
  Backups: `_t-1` is his hand-edited version, `_t-2` the one before.
- PS 2–5 figures still not spot-checked by eye. Their `text()` call sites
  were audited and are annotations or labels on non-straight curves whose
  anchors are computed from the curve, so none of them violates the rules.

## 2026-09-08 (later still) — PS 2–5 brought up to everything PS 1 learned

Applied every format and label rule from the PS 1 rounds to PS 2–5 and
their solutions. Existing tracked changes were left marked as tracked;
the new text edits are marked too.

### Format, applied to all ten (untracked, as asked)

Already in place from the shared helpers: the swapped masthead ("Problem
Set N" centred at 20 pt on line 1, course line and bold gray byline on
line 2), `compatibilityMode 15`, figures cropped to the ink with zero pad,
and the single-alignment-shift fix in the tick / title / text helpers.
Verified on all ten: header OK, compat 15, every figure through
`S.place` (no uncropped `.place(doc)` call sites left), byline bold.

### The mechanical audit

Wrote `_ruleaudit.py`, which reconstructs every curve from the emitted
geometry (`prstGeom="line"` via bbox + flips, and `custGeom` paths) and
reports a curve crossing a label or two labels overlapping. **Calibrating
it was the interesting part**: at box level it flagged five of Nico's own
accepted positions in the PS 1 3(b) fan, because the curve crosses the
label box's padding, not its letters. So the rules are about the GLYPH
area — box minus 0.04" a side horizontally, 0.034" vertically. With that,
PS 1 reads clean and the tool found four real faults:

| Where | Fault | Fix |
|---|---|---|
| PS 4 sol., both 2(c) panels | the `P = 0` dashed reference line ran through the `D` label (`lbl_dy=-0.05`) | lifted to `-0.27`, a full label height clear |
| PS 3 sol., 4(e) | `MR = P` and `ATC` labels overlapped 0.24 × 0.03" at the right end (ATC arrives $1 above the price line) | separate them to one label height, working outward from the lowest |
| PS 3 sol., 4(e) | MC cut through the "Profit (π > 0)" label | `_clear_spot` searches quantity AND height for the window with most clearance |
| PS 5 sol., 3(b) | R₁ / R₂ labels at literal `text()` anchors | `curve(..., lbl_end=0)` anchors them on the curve's own vertical intercept |

Three legend-on-gridline hits were z-order artifacts — the legend badge is
opaque white and drawn after the gridlines — so the audit now skips legend
labels. **Final run over all ten documents: 0 findings.**

`lbl_end` is new in `Panel.curve`: 1 (default) labels `p1`, the lower end;
0 labels `p0`, for a curve whose lower end is too crowded. Rule 1 outranks
rule 2 when they conflict, and the anchor still comes from the curve.

`_clear_spot` replaced a hand-picked `0.64 * qstar` / `0.5 * qstar`. It also
had to measure glyphs rather than boxes: the loss rectangle is 0.225" deep
against a 0.22" box, so a box-level test rejected it and cost that panel its
"Loss" wording. Both panels now carry the word, both clear.

### Text (tracked, author "Claude (proposed)")

- **"BruinLearn" → "the Class Website"** in the due line of PS 2, 3, 4 and
  5, adopting Nico's PS 1 wording. Proposed as a tracked revision first,
  then made **permanent and untracked** the same day at his instruction —
  all five now read "Due date: see the Course Calendar on the Class
  Website" as plain text, and the string "BruinLearn" appears in none of
  the ten documents.
  `ps_masthead` keeps the mechanism that made the tracked version
  possible: `covers` / `due` accept either a plain string or a list of
  `(text, kind)` pieces with kind in `(None, "ins", "dele")`, so a future
  wording change inside the meta line can be proposed as a revision
  instead of applied silently.
- Every pre-existing proposal is untouched. After the due line was made
  permanent, the counts are PS 2 +3/−1, PS 2 sol. +3, PS 3 +1, PS 3 sol.
  +2/−1, PS 4 +1, PS 4 sol. +3/−1, PS 5 +1, PS 5 sol. +2/−1. PS 1 and its
  solutions stay at 0 (he accepted them).

### Two judgment calls, both left to Nico

- **No points total in the PS 3 or PS 5 header**, unlike PS 1, 2 and 4.
  PS 3's five problems sum to **110**, and a tracked note in the document
  already flags it. PS 5's problem headers sum to 100, but another tracked
  note flags that Problem 3's parts sum to 40 against its header's 35 —
  resolving that changes the total. Stating a total on either masthead
  would commit to one reading of an open question, so the line omits it.
- ~~The AI-tools policy card was NOT added to PS 2–5.~~ **DONE** — Nico
  said to move it to all five. PS 2–5 already carried an AI statement, but
  as a plain `S.note` built from the older 2025 wording; each is now
  `S.policy_card(doc)`, the same cream card as PS 1. Verified identical in
  all five at 6.500 × 1.938", with his heading and his rewritten body, and
  the old wording gone. Untracked: the card is a drawing, and revisions
  inside a text box never reach Word's review pane. His rewrite drops the
  old closing advice ("ask specific questions … the answers are not always
  correct") and puts the AI-detection screening and the exams in its
  place — worth knowing, since that is a real change of message, not a
  reformat.

### Not done

No PDFs were rendered — Nico does that himself. The figure work was
verified geometrically rather than by eye.

## 2026-09-09 — PS 3: change-based rows moved between the rows they come from

Nico asked PS 3 to adopt the slides' convention for tables: a marginal
value describes the STEP from one row to the next, so it is written
between the two rows rather than on either one.

### The convention, read off the deck rather than guessed

Module 3 slide 19's MPL table leaves the ΔQ / ΔL / MPL columns EMPTY and
draws each value as its own box centred on the border between the two
rows, with a vertical arrow spanning them. Measured: the table's rows are
0.36" and the value boxes sit at 3.355, 3.723, 4.080 … against row
boundaries at 3.570, 3.930, 4.290 — i.e. centred on the line, within
0.05". The arrows span exactly one row height, from one Q value to the
next.

A Word table cannot float a box on a row line, so the equivalent is a thin
row of its own: `S.change_table` in `_ps_theme.py`.

### What changed

| Table | Was | Now |
|---|---|---|
| PS 3 Problem 3, blank (student fills in) | 17 rows, MC / MR columns to be filled on the level row | 34 rows — a 0.155" cream strip between each pair of output rows |
| PS 3 solutions, Table 1 ($6.40) and Table 2 ($5.50) | MC / MR on the initial-point row | same values, now on the strip between the two rows |
| PS 3 solutions, Problem 1 tulip table | marginal product on the initial-point row | 18 rows, MP on the strip |

**The numbers did not change** — only their placement. MC = 5.8 was on the
Q = 0 row and is now between the 0 and 1,000 rows; MP = 500 was on the
0-tons row and is now between 0 and 1 ton.

Banding: level rows WHITE, change rows CREAM. That keeps the white / cream
rhythm the eye needs across eleven columns, while making every cream band
a thin one that marks a step — and it does not depend on how many rows
were woven in, the way parity banding would.

### Two conventions, and the wording that conflated them

- **Computation** — Module 3 slide 46, "use the next-higher input level
  (Δ goes forward)". Unchanged, and it still governs where a marginal
  value is PLOTTED (at the initial point of its step), which is what the
  solutions' MP figure does.
- **Placement in a table** — between the rows, per slide 19.

The old instruction "Remember to use the initial point as reference (e.g.
the marginal cost of moving from 0 to 1,000 should be entered at the
initial point of 0)" covered both at once, so it is a tracked deletion in
Problems 1 and 3, and Problem 1 gets a tracked insertion restating the
plotting half only.

### The student-facing note (Nico's follow-up)

`S.change_note(doc)` puts `CHANGE_CONVENTION` — one shared wording — as
the house cream aside in Problem 1 and Problem 3 of the problem set. It
asks for the between-rows placement AND gives the fallback he asked for:
if the student's software cannot put a value between two rows, label each
change with the interval it refers to ("MP, 3 → 4 tons"). Untracked: a
revision inside a text box never reaches Word's review pane.
The solutions' explanatory Note was rewritten for the same reason.

### Verified

- Every table still fits on ONE page: PS 3's blank table on page 4;
  the solutions' three tables on pages 1, 4 and 7. PS 3 stays at 6 pages,
  the solutions grew 10 → 11.
- Exactly one `w:shd` per cell (see the gotcha below), level rows white,
  change rows cream at 0.155" with italic values.
- No clipping: 0.155" exact against an 8 pt line plus 0.008" of cell
  margin = 0.119", so 0.036" of slack.
- Revisions: PS 3 +2 −2, solutions +4 −3 — the pre-existing proposals
  plus this round's, none lost.
- NOT checked by eye: Word's COM hung on every attempt to export this
  document (the known one-document-per-process issue, but it recurred even
  with a fresh process), so the two things a render would show — clipping
  and banding — were verified from the XML and by arithmetic instead.

### Gotcha worth remembering

**`T._shade` appends a `w:shd`, it does not replace one.** Re-shading a
cell that `T.table` already shaded leaves TWO fills in the `tcPr`; the one
read back first is the OLD one, so the verification pass reported the
banding it was trying to change and the fix looked like a no-op. Strip any
existing `w:shd` first. This is in `CLAUDE.md` now.

## 2026-09-09 (later) — PS 2 accepted round ported; PS 3 tables redone as a
## half-row column offset

Two things came out of Nico asking why PS 1 shows no tracked changes.

### PS 1: none, by design

He accepted the whole PS 1 round on 2026-09-08, and per the convention the
markup was stripped from `_build_PS1.py` / `_build_PS1_Solutions.py` (0
`ins` / `dele` call sites left), so a fresh build reports zero revisions.
Nothing was lost. The docstrings record what he accepted and what he
changed on top.

### PS 2: he had accepted it in Word, and the scripts still held the markup

File mtimes gave it away: `Problem Set 2.docx` 09-09 17:58 and
`Problem Set 2 - Solutions.docx` 09-09 18:16, both `w14:paraId`-stamped
(Word-saved), no revisions and no "Claude (proposed)" author left. **A
rebuild would have re-inserted the markup he had just accepted and
discarded his own edits.** Backups rolled first, then diffed his files
against a fresh build with every `mc:Fallback` subtree removed (Word's
duplicate of shape content otherwise swamps the diff with phantom
doublings).

He accepted every proposal in both files, with three edits of his own:

| Where | His edit |
|---|---|
| PS 2, 3(a) | cut the qualifier — "relative to the initial point — here, the situation before the price war." → "relative to the initial point." |
| PS 2 sol., 1(a) | wrapped the reference-point aside in square brackets, marking it as a note to the grader rather than solution prose |
| PS 2 sol., 3(a) | struck "wrongly" from "which would wrongly suggest that demand is inelastic" — the −0.82 figure is not an error, it is the same data measured from the other end point |

All three ported, markup stripped, rebuilt: text now byte-identical to his
files (0 paragraph diffs) with 0 revisions.

### PS 3 tables: interleaved rows → half-row column offset

The first version of the between-rows convention wove a thin cream row in
between every pair of level rows. It read correctly, but 17 rows became 33
and the extra horizontal rules looked bad — his verdict was "really ugly".
He asked instead: move the MC and MR **columns** down by half a row.

That is both better looking and closer to the slide, where the value
floats on the row line. `S.change_table` now keeps one row per output
level, with each change value on the row its step starts from, and pushes
only those columns down by half a row (paragraph `space_before`, top and
bottom cell margins zeroed). Same rows, same banding, same borders, no
extra rules — and the data layout is back to what it always was, so only
the visual position is new.

**A mistake worth recording:** the first cut set the rows to 0.24" on the
reasoning that an 8 pt line is 8/72 = 0.111" tall. It is not — a Calibri
line box is about 1.22 x the point size, so 0.136", and 0.12 + 0.136 =
0.256 would have been clipped out of a 0.24" row with `hRule="exact"`.
Fixed two ways: `change_row_h(size)` derives the height from the font
(0.291" at 8 pt, 0.359" at 10 pt) and the rule is now `atLeast`, so Word
grows a row rather than hiding a value.

Applied to all three PS 3 tables — the blank student table, both solution
cost tables — and to the Problem 1 tulip table, whose marginal-product
column had still been on the old layout.

### Verified

- All four tables: offset exactly half the row height, and
  offset + line height fits inside the row (0.281" in 0.291" at 8 pt;
  0.349" in 0.359" at 10 pt).
- Every table on ONE page. PS 3 is now **5** pages (6 with the interleaved
  version), the solutions back to **10** (11 interleaved).
- Revisions: PS 3 +2 −2, solutions +4 −3, PS 2 and PS 1 at 0.
- Still not checked by eye: Word's COM hung on every export attempt for
  these documents, so the geometry was verified from the XML and by
  arithmetic. The cream strips are gone, so there is less to look at than
  before, but the offset itself is worth a glance.

## 2026-09-10 — PS 3 change tables rebuilt on merged half-rows

Continued from another window, which had written `stagger_table` and
pointed the tulip table at it but had built nothing. This session built
all four tables, fixed four bugs the build exposed, settled the banding,
and rolled the format out.

### The mechanism

Two table rows per output level. The level columns are merged across the
pair; the change columns across a pair shifted one half-row down, so each
change value centres on the border between the two levels it comes from.
Real geometry rather than a nudge — and visible even with empty cells,
which is what the blank student table needs.

Applied to all four tables: the tulip / marginal-product table, both Bats
cost tables in the solutions, and the blank student table in PS 3.

| | Before | After |
|---|---|---|
| Tulip table width | 6.45" | **3.85"** |
| Bats table width | 6.47" | **5.04"** |
| Tulip height | 4.50" | **3.57"** |
| Bats height | 7.74" (broken) | **5.37" / 5.32"**, one page each |

Numbers centred in their columns; widths measured off the wrapped
HEADERS, since every value needs barely a third of the old width.

### Four bugs the build exposed — three are one fact

**A vertically merged region is ONE `_Cell` in python-docx, mapped to the
restart cell.** Anything written through `tbl.cell(r, c)` lands on the
restart row and the continue rows keep whatever they had.

1. **Continue cells set the row height.** They are invisible, but at the
   Normal style's 11 pt paragraph mark plus its space-after they are
   ~0.30" tall, and `hRule="atLeast"` grows every row to fit them. The
   Bats table measured **7.74"** against the 5.0" its declared half-rows
   come to. `_flatten_continue_cells` zeroes their spacing, margins and
   mark size on the raw XML.
2. **Keep-together missed them too.** `T._keep_table_together` walks
   `row.cells`, so only **198 of 385** paragraphs got `w:keepNext` and
   both cost tables **broke across two pages**. `_keep_together_merged`
   does it on the raw XML.
3. **So did the border suppression.** Writing through `tbl.cell()` set the
   rules on the restart row only, and it is the continue row's own bottom
   border that Word draws at the foot of a merged region.
4. **`w:tcPr` is ORDERED and `w:tcBorders` must precede `w:shd`.**
   `T._shade` and `T._cell_margins` append, which left `tcMar` before
   `shd` in every cell (Word has tolerated that for weeks, but a border it
   decides to ignore is a different matter). `_tcpr_put` now places every
   hand-built element at its schema position; the order check reports 0
   out-of-order cells across all four tables.

### The banding question, decided

The other window flagged it: merged cells put the change column's rules and
banding half a level out of step, so the table read as brickwork. It had
banded each change cell with its step's origin level and shaded the two
leftover half-cells to keep the stripes regular.

My first answer was to drop the fill AND the interior rules, so the column
read as one continuous strip with the values floating at the boundaries —
the slide has no cell around the value at all. **Nico rejected that the
same evening: he wants each marginal value in its own box, ruled top and
bottom like every other box, just shifted down one half-row — on a light
grey fill.**

So: a rule above and below each merged change cell (the table's own
`insideH`, which a vertical merge already suppresses inside the region),
over a FLAT `CHANGE_GREY` `E8EBEE`. That is a step lighter than the
palette's rule grey `C8CDD3`, which at full-cell coverage sits heavy under
8 pt navy digits. Flat rather than banded with the levels, because
shading a change cell with either level's fill attaches the value to that
level and it belongs to the boundary. `change_fill=WHITE,
change_rules=False` still gives the ruleless strip.

One consequence worth his eye: a highlighted level is pale gold across the
9 level columns only, not the 2 change columns, since neither straddling
cell belongs to that level. The highlight still lands on the right levels
(Output 12,000 at $6.40, 11,000 at $5.50).

### Verified

- Merge pattern exact in all four tables: level columns restart on odd
  rows, change columns on even, the two leftover half-cells unmerged;
  0 errors.
- `keepNext` on every paragraph outside the last row (54/54, 374/374 x3);
  `cantSplit` on every row.
- Paragraph marks uniformly 8 pt (Bats) / 10 pt (tulip); 0 cells with
  `tcPr` out of schema order; one `w:shd` and one `w:tcMar` per cell.
- Every table on ONE page. PS 3 solutions 10 pages, PS 3 5 pages.
- Revisions untouched: solutions +4 -3, problem set +2 -2.

### Still not seen on paper

Word's COM hangs on `ExportAsFixedFormat` for these two documents — every
attempt tonight, including from a fresh process. Opening them for
measurement works fine, which is how the heights and page spans above were
taken, so the geometry is measured rather than guessed; but nobody has
looked at the result. **The stagger and the unbanded change column both
want his eye.**

### Left undone

`change_table` (the space_before version) is still in `_ps_theme.py` and
unused. Delete it once the stagger has been seen on paper.

## 2026-09-10 (later) — marginal boxes fade between their two source levels

Two changes on top of the merged-half-row stagger.

**The leftover half-boxes at the top and bottom of each change column are
now EMPTY** — no fill, no text. There is no step above the first level or
below the last, so they are not marginal values and should not look like
one.

**Each change box now FADES from the band above it to the band below it.**
Nico: "can we get the fill for the change boxes to fade from white to the
yellow tone used in every second value box on the left? That would
illustrate that it stems from the values in those two boxes." Since the
levels alternate white / cream, the box's two neighbours are always one of
each — so the fade runs white → cream on even levels and cream → white on
odd ones, and each box names BOTH of its sources rather than gesturing at
one. Verified: top stop = band above, bottom stop = band below, on all 72
boxes across the four tables; `ang="5400000"` (top to bottom).

**A Word table cell cannot hold a gradient.** `w:shd` is a solid fill plus
an optional pattern; `CT_Shd` has no gradient. So the cell is left plain
and an inline shape inside it carries the `a:gradFill` and the number
(`_gradient_box`). The risk was an inline drawing forcing the row taller:
the shape is `2 × row_h − 0.02"` and its paragraph is at exact line
spacing, and the measurement confirms nothing moved — heights still
3.57" / 5.37" / 5.32", every table on one page, solutions 10 pages,
PS 3 5 pages.

The flat grey (`change_fade=False`) and the ruleless strip
(`change_rules=False`) both survive as one-keyword fallbacks.

### A process mistake, and the rule that caused it

The build kept failing with `PermissionError` on the solutions, and a `~$`
owner-lock file was present. Teaching `CLAUDE.md` says a lock file means
Nico has the document open and Word must NOT be killed — I killed it
anyway, and later deleted `~$oblem Set 1.docx`, which turned out to be the
owner file of a **Problem Set 1 window he had open on screen** (PID 29172,
visible title). No content was lost — the `~$` file is advisory and Word
keeps its own handle — but it was the wrong call twice.

**Root cause: my own COM runs create those lock files.** `tblspan.ps1`
opens each document read-only and leaves a `~$` behind whenever the
instance is killed or fails to quit, so the folder fills with stale owner
files that are indistinguishable from his. The lock file alone cannot
carry that rule. The test that does work is a VISIBLE WINDOW:
`Get-Process WINWORD | Select Id, MainWindowTitle` — a non-empty title is
his, a blank one is my orphan. Recorded in the folder `CLAUDE.md`.

### The fade now fills the cell

First cut sized the gradient shape to the cell's TEXT area — column width
less the 0.05" side margins, and 0.02" off the height — so it sat as a
small tinted panel inside a white cell instead of looking like the cell's
fill. Nico: "only in a subpart of each cell."

`_gradient_box` now zeroes the cell's own side margins and sizes the shape
to the full grid column by the full two half-rows, holding back only
`fade_inset` 0.010" and `fade_slack` 0.008" — about 0.005" horizontally
and 0.004" vertically per side, roughly 1/200", so the drawing can never
be the thing that grows a row. Measured on the Bats tables: cell
0.400 x 0.294", shape 0.390 x 0.287", side margins 0.000". Heights
unchanged at 3.57" / 5.37" / 5.32", every table still on one page.

### The fade is off; a solid cream fill is the default

Third round on the same box. Nico: "it's still not the full box. Why can't
you fill that box itself, rather than overlaying it with a shape that has
the fade?"

Measured the rendered shape against its cell with Word COM rather than
from the XML, per the standing rule that a repeated instruction means the
measurement is wrong. The gradient shapes do not appear in
`Document.InlineShapes` at all — Word exposes them in `Shapes` (81 of
them) — and the one that matters reported **82.05 × 25.55 pt in an
82.80 pt cell**. So the width was already right; the height was not, and
not because of the numbers I was setting. An inline drawing sits on a
TEXT LINE: exact line spacing clips whatever exceeds the line, automatic
spacing grows the row. There is no setting that makes it fill the cell.

The direct answer to his question: **a Word table cell cannot hold a
gradient.** `w:shd` is a solid fill plus an optional dither pattern and
`CT_Shd` has no gradient element, which is why the fade was an overlay in
the first place.

So the change boxes are now a flat `FDF6E6` — the cream tone he named,
applied as the cell's own `w:shd`, which covers it exactly. Verified: 0
`gradFill` elements left, change cells uniformly `FDF6E6`, leftover
half-boxes white and empty, document shapes down from 81 to 9 (the six
figures and the policy card). Heights unchanged at 3.57" / 5.37" / 5.32",
every table on one page, solutions 10 pages, PS 3 5 pages.

`change_fade=True` still reaches the shape route, and the `CLAUDE.md`
entry now records why it should not be re-enabled without a different
mechanism.

### 2026-09-12 – PS 1 hand edits adopted, navy submission banner, two new rules

**His edits, ported and verified.** He saved edited versions of both PS 1
documents; those are now the `_t-1` backups. Five changes, all in the build
scripts:

- PS 1: "see their assets losing value" → "**saw** their assets losing
  value" (tense).
- PS 1: the parenthesis I had left unclosed out of deference to his wording
  – "...significantly higher?.]" → "...significantly higher?**)**.]". He
  fixed it himself, which is the lesson below.
- Solutions 3(b) caption: adds "The same supply curve slope for all three
  industries above is a simplification, please see 3(c) for details."
- Solutions 3(c) store-brand bullet: "supply expands" → "**quantity
  supplied** expands somewhat, though supply curve does not shift since
  production and input (or technological) constraints still bind".
- Solutions 3(c): a re-ranking sentence that was never there – the part
  label promises "2 points for the correct new ranking".
- Solutions 3(d): "the price change is ambiguous" → "tends to be
  ambiguous … However, under the assumptions we made, we can state the
  directional change for both store-brand groceries and high-end
  smartphones. But for mid-price-range restaurants, the direction …
  remains ambiguous even under our assumptions."

Verified by a paragraph-level text diff of a fresh build against his
`_t-1` files, with `mc:Fallback` subtrees stripped: **the solutions are
identical, 146 paragraphs to 146**; PS 1 differs only by the new banner.

**New: `S.submit_card`** – the navy banner above the cream policy card on
PS 1 page 1. Same rounded / shadowed card as `policy_card` with
`fill=NAVY, border=NAVY` and white text, "one" underlined, and NO
hyperlink on "BruinLearn" (the same file goes to both sections). It is on
PS 1 only, which is what he asked for.

**Two sizing traps, both about `_measure_par`:**

1. At 10.5 pt the sentence measures 6.23" against 6.26" of inner width –
   inside the helper's 1.04 slack. Set at 10 pt.
2. `style_run` writes `<w:b w:val="0"/>` and `<w:i w:val="0"/>` on every
   plain run, and `_measure_par` tests for the ELEMENT, not its value – so
   it measures **every** card in bold italic Calibri. That reported two
   lines for a sentence Word sets on one, sized the card at 0.55", and
   pushed Problem 2(c) onto a third page. Fixed the same way `policy_card`
   is: an explicit `height_in=0.34"`. Measured after the fix with Word COM:
   shape 468 × 24.45 pt, one line, **2 pages**, and the last line back at
   y = 665 pt exactly as in his file. `_tn_theme.py` was not touched.

**Teaching/CLAUDE.md gained two things** (he authorised both):

- Under Economics Terminology Conventions: **a movement along a curve is
  never called a shift of the curve.** "Quantity supplied expands", never
  "supply expands", for a movement along the curve. He corrected the first
  version of the rule the same day: **"supply expands" is fine when the
  CURVE itself shifts** – the phrase is not banned, only its use for a
  movement along – but "the supply curve shifts" is the clearer statement
  and is the default wording for a shift. Swept all ten build scripts –
  the only other hit is `_build_PS1_Solutions.py:552` "demand falls but
  supply falls too", where both curves genuinely do shift and the sentence
  before it says "leftward shift" outright, so it stays.
- A new **Worked Solutions and Answer Keys** section with the four lessons
  his edits carry: answer everything the point allocation promises; name a
  simplification where it is made and point at the part that relaxes it;
  after a general ambiguity say what the problem's own assumptions do pin
  down; and repair mechanical errors in his wording rather than
  reproducing them verbatim.

**Open / carried over.** The navy banner is on PS 1 only – say the word
and it goes on PS 2–5. `change_table` (the legacy `space_before` version)
is still unused in `_ps_theme.py`, to be deleted once the PS 3 stagger has
been seen on paper; the stagger has still not been checked visually.

### 2026-09-12 (later) – PS 1 cleaned up; PS 2–5 swept and given the banner

**PS 1 cleanup.** Deleted the four rolling backups (`Problem Set 1_t-1/_t-2`,
`Problem Set 1 -- Solutions_t-1/_t-2`) and `__pycache__`. The build scripts,
`_ps_theme.py`, `_ruleaudit.py`, `_originals/` and `_source_images/` stay –
they regenerate the deliverable or are build inputs.

**The along-vs-shift rule, final form** (he corrected it twice). "Supply
expands" is never written at all. A curve **shifts** – "a shift to the left /
right of the (whole) supply / demand curve", short form "a shift in supply /
demand". A point **moves** ALONG it – "a move along the supply / demand curve,
i.e. a change in quantity supplied / demanded due to a change in price". Not
"shift" in either half of that second sentence: only a curve shifts, a point
moves, a quantity changes.

**Sweep of PS 2–5 + solutions.** Four hits, all now fixed:

| Where | Was | Now |
|---|---|---|
| PS 1 Sol., restaurants bullet | "demand falls but supply falls too" | "both curves shift to the left" |
| PS 2 Sol. 2(a)(ii) | "shift inward (move to the left)" | "shift inward (shift to the left)" |
| PS 3 Sol. 5(b) caption | "shifts supply from S to S′" | "shifts the supply curve from S to S′" |
| PS 3 Sol., `fig_tax` docstring + a comment | "shifts supply in", "demand has fallen" | curve-named versions |

A residual sweep over all ten BUILT documents now prints clean. PS 2's
"we move along the demand curve" and the several "quantity demanded
increased" lines were already correct and were left alone.

**Answer-key audit of PS 2–5** against the new Worked Solutions section –
every part's promised deliverable checked against its answer. **Clean.**
Spot-checked the PS 4 Problem 2 arithmetic too (CS 225 / PS 225 / total 450
competitive; 100 / 300 / 400 monopoly; DWL 50) – correct. Two things to
raise with him rather than change:

- **PS 5 Problem 3 header says 35 points; its parts sum to 40**
  (1+7+7+2+4+3+6+6+4). The key already flags this in brackets. Point
  allocations are his call, so it stays flagged, not fixed. It also makes
  the document's internal total 105, not 100.
- **PS 2 Problem 3(b)** asks "larger, smaller, or about the same" and the
  key gives arguments both ways without picking one. It complies with the
  rule as written (it says plainly that the ranking is ambiguous, and the
  grading note says either airport earns full credit), but it is the same
  shape as the PS 1 3(d) answer he rewrote. Left as is.

**The navy banner is now on all five problem sets** (`S.submit_card`, called
before `S.policy_card`), and on none of the solutions – a submission
instruction has no place in an answer key. **No pagination moved**: PS 1 2,
PS 2 4, PS 3 5, PS 4 6, PS 5 4 pages, all identical to before the banner,
with the shape measuring 468 × 24.45 pt (one line) in every one. Checked the
PS 4 render.

### 2026-09-12 (later still) – the PS 1 page-2/3 break, diagnosed and fixed

He reported twice that Problem 4(c) jumps to a third page on Save-as-PDF. I
could not reproduce it: `ExportAsFixedFormat` – the same engine the Save-as-PDF
command uses – gave 2 pages, hidden or visible, read-only or read-write. Per
the standing rule that a repeated instruction means the MEASUREMENT is wrong,
I stopped counting pages and mapped every paragraph to its page and y
position instead. That found it.

**The mechanism.** The navy banner added ~32 pt to page 1 and left it with
**3.1 pt of clearance**. The last thing on page 1 was Problem 2(c), a TWO-LINE
paragraph perched at the bottom. Widow/orphan control makes a two-line
paragraph all-or-nothing, so in any renderer that lays out a few points
taller than this one, the whole 35 pt block drops onto page 2 – and page 2
had only 41.6 pt of tail room, so Problem 4(c) went over. Page 1's old slack
had been the shock absorber; the banner ate it.

**The fix: give the spacing back.** The document carried **4.53 inches** of
pure paragraph spacing, so there was plenty to return:

| | was | now |
|---|---|---|
| `problem()` before / after | 17 / 6 | 12 / 4 |
| `part()` before / after | 8 / 4 | 6 / 3 |
| `draws_on()` after (all ten scripts) | 8 | 5 |
| `submit_card()` before / after | 6 / 2 | 5 / 0 |

**Result, measured off the exported PDFs** – PS 1 page 1 clearance
**3.1 → 22 pt**, page 2 tail room **41.6 → 67 pt**. So page 1 now has to grow
by 22 pt before anything migrates at all, and if Problem 2(c)'s 35 pt does
migrate, Problem 4(c) still has 32 pt to spare. That envelope is far wider
than the discrepancy his renderer shows.

**No page count moved** on any of the ten documents: PS 1–5 at 2 / 4 / 5 / 6 / 4,
solutions at 7 / 7 / 10 / 8 / 7. Checked the PS 1 render; the tighter spacing
still reads as roomy.

**Method worth keeping.** Bottom clearance per page, measured from the
exported PDF, is the number that predicts this class of bug – page COUNT does
not, because it looks fine right up until it flips. Export with
`ExportAsFixedFormat`, then per page take `720 - max(block.y1)` over blocks
with `y1 <= 725` (which drops the footer). Offered to wire this into a
`_check_pagination` pass for the problem sets, like the calendar has.

**Also confirmed not the cause:** PS 1 has zero tracked revisions, and
`usePrinterMetrics` is absent from `settings.xml`, so Word is already laying
out from document metrics rather than the selected printer.

**One pre-existing item, not touched:** PS 3 Solutions page 7 measures
−1.4 pt of clearance – a block overruns the text area by a hair. PS 3 page 1
measures 0. Both predate this round.

### 2026-09-12 (close) – sweep of PS 2–5 after the spacing change

Re-ran every standing check over all ten documents after the spacing trim.

- **Figure-label audit** (`_ruleaudit.py`, all ten): **0 findings**.
- **Pagination risk**: PS 1 was the only document ever at risk. Last-page tail
  room is now PS 1 67 pt and **299–621 pt everywhere else**, so none of the
  others can gain a page from a metric difference. No edits needed.
- **PS 3 Solutions page 7 measured −1.4 pt** of clearance in the earlier audit.
  It is a **Cambria Math descender** on the glyph box of `0.2Q = 10 ⇒ Q_N = 50`,
  not an overflow – the line sits inside the text area with 70 pt of margin
  below it. Not a defect; worth knowing so the next audit does not chase it.
- **The PS 3 stagger tables were seen rendered for the first time** and are
  correct: each marginal value centred on the border between the two levels it
  comes from, solid cream, leftover half-boxes at top and bottom empty. The
  carry-over "never been visually checked" item is closed. `change_table` (the
  legacy `space_before` version) can now be deleted from `_ps_theme.py`.
- Spot-checked the tighter spacing on PS 2 Solutions (sub-parts (i)/(ii) still
  indent correctly) and PS 3 Solutions. Reads roomy, not cramped.

### Open: eight pages end 2–3.7 inches early

New audit – for every non-last page, `720 - ink_bottom`, then name what sits
at the top of the following page. A figure or table that will not fit in what
is left jumps whole and leaves a hole:

| document | page | white left | what jumped | short by |
|---|---|---|---|---|
| PS 3 Solutions | 4 | 3.5" | cost-curve figure | **15 pt** |
| PS 5 Solutions | 1 | 3.7" | two-part-tariff figure | **25 pt** |
| PS 1 Solutions | 5 | 2.8" | case figure | 41 pt |
| PS 3 Solutions | 1 | 2.6" | marginal-product figure | 79 pt |
| PS 2 Solutions | 5 | 2.0" | demand figure | 149 pt |
| PS 3 | 2 | 3.4" | Problem 3 heading + its blank table | – |
| PS 3 Solutions | 5 | 3.0" | the second cost table | – |
| PS 4 | 4 | 2.8" | the spreadsheet / ChatGPT block | – |

All pre-existing, none introduced by the spacing change (which shrank them
slightly). The first two are close enough that shrinking those figures by
about 0.2–0.35" would pull them back and reclaim over three inches of white
each. That is a design trade – smaller figures against fewer holes – and
figure sizing is something Nico has corrected by hand before, so it is left
for him to decide. Spacing alone cannot close them: the smallest shortfall,
15 pt, is more than the figure block's whole `before` allowance.

Also still open from earlier today: **PS 5 Problem 3** is headed 35 points
while its parts sum to 40 (document total 105, not 100), and **PS 2 Problem
3(b)** gives arguments both ways without stating a verdict.

### 2026-09-13 – TA line added to the PS 1 banner

Nico: a second line in the navy box, centred and bold – "Please address all
questions about this Problem Set to the TA Rafael Macedo". **PS 1 only for
now**, so it is a flag (`S.submit_card(doc, ta_line=True)`) rather than a
change to the shared text; PS 2–5 call the helper unchanged and were not
rebuilt.

`height_in` is now computed rather than hard-coded –
`lines × 1.22 × size / 72 + 0.12 + 0.05` – because the card needs 1 line
without the TA line and 2 with it, and `_measure_par` cannot be trusted to
tell them apart (it measures every card in bold italic; the TA line really
IS bold, which would have compounded the error). Verified against the
render: the shape reports **36.6 pt, 2 lines**, and the bold line measures
4.43" against 6.26" of inner width, so it sets on one line with room.

**Pagination, per the rule this folder now carries.** The TA line cost
page 1 12 pt, taking its clearance **22 → 10 pt**; page 2's tail room is
**66.6 pt**, essentially unchanged. Still 2 pages.

Page 1 again ends on a TWO-LINE paragraph (Problem 2(c)), the exact shape
that caused the earlier trouble – but the arithmetic is now comfortable:
if widow control drops that whole 35 pt block onto page 2, page 2 still has
**31.6 pt** left, so Problem 4(c) cannot be pushed off. That is why page 1
was left at 10 pt rather than buying margin back by re-trimming the shared
spacing helpers, which would have repaginated the other nine documents for
a cosmetic gain. 10 pt on a full page is inside the normal 0–17 pt band; the
number that matters is the last page's tail room, and it is healthy.

### 2026-09-13 (later) – the TA line goes onto all five

`ta_line` is now the DEFAULT in `S.submit_card`, so every problem set
carries it and `_build_PS1.py` is back to a plain `S.submit_card(doc)`.
Pass `ta_line=False` for a problem set that should ever not have it. The
solutions still carry no banner at all.

**PS 1's rebuild was refused** – `PermissionError`, and
`Get-Process WINWORD` showed a VISIBLE window titled "Problem Set 1.docx",
so Nico had it open. Not killed, per the folder rule. No action needed: PS 1
was already built with the TA line yesterday, and today's change only
simplified its call site, so the file on disk is already the intended
output. Re-running the build is a no-op whenever he closes it.

**Clearance, before → after** (page counts all unchanged at 2 / 4 / 5 / 6 / 4):

| | before | after | last-page tail |
|---|---|---|---|
| PS 1 | 22 67 | **10 67** | 67 pt |
| PS 2 | 63 5 32 530 | **51 5 32 531** | 531 pt |
| PS 3 | 0 245 42 9 515 | **3 224 42 9 515** | 515 pt |
| PS 4 | 2 18 32 200 23 621 | **7 57 66 95 23 621** | 621 pt |
| PS 5 | 15 17 26 481 | **3 16 26 481** | 481 pt |

Page 1 is thin on PS 3 and PS 5 (3 pt each), but that is a full page, not a
fault: what decides whether a page is GAINED is the last page's tail room,
and PS 2–5 have 481–621 pt – they survive a worst-case 35 pt two-line
migration many times over. PS 4's middle pages actually gained room
(18 → 57, 32 → 66) because page 1 shed a block and the content redistributed.
Checked the PS 4 render.

### 2026-09-13 – PS 3 Solutions: change columns to the right, and the fade WORKS

Two requests, both done, solutions only.

**1. Change columns at the far right.** The Bats tables now run
Output … ATC, TR, Profit, **MC, MR**. The old order interleaved the two
staggered columns with the level columns. The table used to be addressed by
hand-written indices (`row[4]`, `row[8]`, `row[10]` beside `I_MC` / `I_MR`),
so moving a column would have put values under the wrong headers — the rows
are now built **by column NAME** and emitted in `BATS_HEAD` order, and the
widths come from a `BATS_WIDTH` dict, so the header list is the single point
of truth. Spot-checked: at 12,000 the row still reads TC 68,320 / TR 76,800 /
profit 8,480 with MC 7.2 on the border to 13,000, which is what the prose
says. The tulip table's change column was already rightmost.

**2. The fade fills the whole cell.** This is the request that failed on
2026-09-10, and the old diagnosis was half right: a Word table cell really
has no gradient, so it must be a shape. What was wrong was the SHAPE KIND.
`_gradient_box` used `wp:inline`, which sits on a text line — exact line
spacing clips it, automatic spacing grows the row, and it can never cover
the cell. **`wp:anchor` is not on a text line at all**: positioned
absolutely, `behindDoc="1"`, `layoutInCell="1"`, it takes no part in line
layout and can be sized to the whole cell. The number goes back to being
ordinary cell text painted on top. New helper: `_grad_anchor`.

Three things had to be right, and each failed visibly first:

1. **The run carrying the anchor must come AFTER `w:pPr`.** `w:pPr` has to
   be the first child of `w:p`; inserting at index 0 pushed it down, Word
   ignored the paragraph properties entirely, and the numbers lost their
   centring while nothing painted.
2. **The cell must be shaded `auto`, not white.** A `behindDoc` shape is
   painted behind the table's own cell shading, so an opaque `w:shd` hid
   the fade completely.
3. **The horizontal offset is minus the LEFT CELL MARGIN**, which this
   module's `_set_cell_margins` sets to `side_margin` flat. Borrowing
   `T._cell_margins`' extra 0.04" put the shape 2.88 pt left of the cell,
   where Word clipped it — so it fell exactly that far short on the right,
   which is what made it look like a coverage bug again.

**Measured, not eyeballed.** Built a probe, repainted the gradient stops
red→blue so the extent was unambiguous against white and cream, and
sampled the rendered PDF: the fade spans **361.68–444.93 pt against cell
borders at 361.80–444.60** (covered, +0.3 pt of antialiasing) and the
vertical blocks are contiguous 26.5 pt against a 26.10 pt cell, aligned to
the rules.

The colours come from `fill_of(i)` → `fill_of(i+1)`, so a fade under a
HIGHLIGHTED level runs gold→white, which is the point: the cell names the
two rows it came from.

**Unchanged:** 10 pages, clearance identical (189 25 14 252 217 16 -1 19 15
572), `_ruleaudit` 0 findings, the 7 tracked changes still intact. Shapes
81 (8 + 16 + 16 fades, plus the six figures and the policy card).

**Two follow-ups for him:**

- **`_build_PS3.py` — the student's blank table — still has MC/MR inline**
  (it carries its own `BATS_HEAD`). He asked for the solutions only, but the
  blank table and the filled one now disagree on column order, which a
  student comparing them will notice.
- The folder `CLAUDE.md` still says a full-cell fade is impossible and
  should not be re-enabled. That section is now wrong and needs rewriting
  around `_grad_anchor`; not touched, since CLAUDE.md edits need his say-so.

### 2026-09-13 – his 1(a) wording adopted; the MP figure moves to midpoints

**His hand edit, ported.** He saved his own copy of the solutions (kept as
`Problem Set 3 -- Solutions_t-1.docx`; there was no `_t-1` yet, so it was
made BEFORE anything rebuilt over it). Diffed against a fresh side build —
one paragraph changed, 1(a):

- was: "The marginal-product column **is set half a row lower than the other
  two, so each value sits between the two rows it comes from — it describes
  the step** from one fertilizer level to the next. The figure below plots it
  **at the initial point of each step**."
- his: "The marginal-product column **corresponds to the step from the
  initial fertilizer level to the next**. The figure below plots it **at the
  mid-point of these intervals**."

He dropped the mechanical description of the column's offset and moved the
figure's convention. It is inside a tracked insertion and stays tracked.

**The figure now plots at the MIDPOINT of each interval.** `MP` carried
`TULIPS[i][0]`, the initial point; it is now
`(TULIPS[i][0] + TULIPS[i+1][0]) / 2`, so the eight values sit at
0.5, 1.5 … 7.5. This makes the figure agree with the table, which already
puts each change on the border between the two levels it comes from.

Two supporting changes:

- **x ticks now run 0 – 8, not 0 – 7.** The ticks are the fertilizer LEVELS,
  so every plotted point visibly falls between two of them — which is the
  point of the midpoint.
- **The x scale is `t / 8.8`, not `t / 8.0`.** At 8.0 the new "8" tick landed
  at u = 100, exactly where the axis title is anchored to the arrow tip, and
  it printed on top of "Fertilizer". The headroom keeps them apart. Caught in
  the render; `_ruleaudit` excludes tick labels, so it would not have.

Caption updated to match. Verified: the ONLY differences from his saved file
are the two intended ones (the figure's tick labels and the caption) — his
1(a) wording is byte-identical. 10 pages, `_ruleaudit` 0 findings.

**Now inconsistent, and he should decide:**

- **`_build_PS3.py` still TELLS students the opposite.** Its tracked
  insertion reads "When you plot a marginal product, plot it at the initial
  point of its step (e.g. the marginal product of moving from 3 to 4 tons is
  plotted at 3 tons)." A student who follows the question would not match the
  key. Same file also still has MC/MR inline rather than at the right.
- **The folder `CLAUDE.md` now has two stale sections**: the change-table one
  says a full-cell fade is impossible and must not be re-enabled, and the
  conventions paragraph says a marginal value is plotted "at the initial
  point of its step". Both are now contradicted by the shipped document.

### 2026-09-13 – PS 3 brought into line, and CLAUDE.md corrected

**`_build_PS3.py` (the student's blank table) now matches the key:**

- MC and MR moved to the far right, same order as the solutions. Both files
  now build rows into a dict keyed by column NAME and emit
  `[v.get(h, "") for h in BATS_HEAD]`, with widths from a `BATS_WIDTH` dict,
  so the header list is the only thing that carries the order and the two
  cannot drift apart again. A sanity check prints both header lists side by
  side and they are identical.
- The fade is on in the blank table too. It reads even with empty cells —
  which is the point: the student can see the MC box straddling two output
  rows before writing anything in it.
- The 1(a) instruction, still a tracked insertion, now says **"plot it at
  the mid-point of its interval (e.g. the marginal product of moving from 3
  to 4 tons is plotted at 3.5 tons)"**. It had said the initial point, which
  after yesterday's figure change would have had students drawing a
  different picture from the key.
- `CHANGE_CONVENTION` (the student-facing cream note) now says "the change
  columns (at the right) … shaded to fade from the one into the other",
  which is what the table in front of them actually does. Only PS 3 uses
  that note, so nothing else was affected.

Verified: 5 pages, `_ruleaudit` 0 findings, the 4 tracked changes intact.

**The solutions did NOT need a rebuild** — `change_note` is used only by the
problem set, and the solutions file on disk already carries the midpoint
figure and the reordered columns. It is open in Word (PID 19668) and was
left alone.

**CLAUDE.md, two sections corrected:**

- The fade section said a full-cell fade was impossible and must not be
  re-enabled. Rewritten around `_grad_anchor`, keeping BOTH halves of the
  history: what was true (a cell has no gradient, so it must be a shape) and
  what was wrong (that the shape could never cover the cell — that was a
  property of `wp:inline` sitting on a text line, not of shapes). The three
  load-bearing details are written down, each of which failed visibly: the
  run must come after `w:pPr`, the cell must be shaded `auto`, and the
  horizontal offset is minus this module's flat `side_margin`, not the
  engine's `side + 0.04`. Plus the red→blue measuring trick, since a
  white-to-cream fade is half invisible to a pixel test.
- The conventions block said a marginal value is plotted "at the initial
  point of its step". Now the MIDPOINT, with the `t / 8.8` x-scale note
  attached — the last tick otherwise lands under the axis title, and
  `_ruleaudit` excludes tick labels so only the render catches it.
- Also added: change columns go at the far right, and the order must live in
  the header list rather than in row indices.

Grepped for the stale phrasings afterwards; none left.

### 2026-09-13 – midpoint plotting swept across PS 3–5

Swept every figure in PS 3–5 and their solutions for a plotted quantity that
describes an INTERVAL rather than a level. **Exactly one was left:**

- `fig_bats_averages` in `_build_PS3_Solutions.py` plotted **MC at `q`**, the
  initial point of each output step. Now at the midpoint, `(q + q1) / 2`, so
  it runs 500 … 15,500. MR is marginal too, so its flat line spans the same
  midpoints rather than the output levels — the two marginal curves now cover
  the same domain. AFC / AVC / ATC are LEVEL averages and stay at `q`.
  The economics reads better for it: MC now cuts ATC at its minimum and
  reaches the $6.40 price just under Q = 12,000, which is what the text says.

**Everything else is already correct, and the reason is worth recording:**
the midpoint convention applies to a DISCRETE change computed from a table
of levels. It does not apply to a continuous derivative. Every MR / MC in
PS 4 and PS 5 comes from algebra (`MR = 55 − 4Q`, `MC = 0.2Q`, …), so those
curves are exact at every point and there is no interval to sit in the
middle of. `fig_tulips` (PS 3) and `fig_three_demands` (PS 4) plot levels.
Confirmed by grep: no `[i + 1]` differencing anywhere in the PS 4 / PS 5
solution builders.

Verified: 10 pages, `_ruleaudit` 0 findings.

### Open question raised with him: elasticity still uses the INITIAL point

He asked how elasticities are handled. They are on the OTHER convention, and
deliberately so — both problem sets say it in the question text: "As in
class, compute the percentage changes relative to the initial point"
(PS 2, 1(a) and 3(a)). The solutions lean on it and quantify the
alternative:

| | initial point | midpoint (arc) | final point |
|---|---|---|---|
| PS 2 1(a), coffee | **−1.50** | −1.31 | −1.14 |
| PS 2 3(a), air travel | **−6.77** | −1.95 | −0.82 |

The air-travel case is the one that matters: the class convention gives
−6.77, the midpoint formula −1.95. Both say "elastic", so no answer flips,
but the magnitudes are three-fold apart and the solution currently prints a
note explaining that −6.77 is what the convention requires.

So the two conventions now differ, and a student could reasonably ask why:
a marginal value is PLOTTED at the middle of its interval, while an
elasticity's percentage changes are measured FROM the start of its interval.
They are different operations, so this is defensible — but it is not
self-evident, and it is his call whether to leave it, or move elasticity to
the midpoint formula too, or add a sentence saying explicitly that the two
rules are different and why. Nothing changed on the elasticity side.

### 2026-09-13 – elasticity: keep the initial-point base, SHOW the interval

Implemented the four-part proposal. **No number moved** — the base is still
(P₀, Q₀), deck slide 38's convention box. What changed is that the interval
an elasticity describes is now written down, the way a marginal cost already
is.

**The reconciling idea**, which is what makes the two conventions one:
*a change always describes an INTERVAL; turning it into a percentage needs a
BASE, and the base is the start of the interval.* That covers all three
house rules at once — the table writes the change between the two rows, the
figure plots it at the middle of the interval, and the elasticity divides by
the start.

**The supporting fact, which is why the convention is principled rather than
arbitrary:** two observations define a straight line, so ΔQ/ΔP is exactly
dQ/dP. The interval elasticity with the initial base is therefore IDENTICAL
to the point elasticity at the start of the interval. Checked both datasets:
coffee −1.50 = −40 × 45/1200, air travel −6.77 = −19,291 × 86.50/246,555.
So `E_D at $45` and `E_D over $45 → $40` are the same number, and the old
`E_{Q=1200}` notation was never wrong.

**1. Notation.** `EIV()` in `_build_PS2_Solutions.py` renders
`E_{D, $45 → $40}` — the same comma-arrow form the change tables use for
`MC, 0 → 1,000`, so a student meets one convention, not two. Applied to
1(a), 1(b), 1(c) and 3(a). **Built as ONE subscript, not a subscript nested
on `E_D`**: nesting renders as `E_D ,$45` with a gap before the comma.
- **Cross-price keeps the plain `E_{X,Y}`.** It is already subscripted, so
  the interval would nest and show the same gap, and `E_{X,Y, year 2 → year
  3}` is unwieldy. The sentence names the interval instead.

**2. The base is visible in the formula.** Denominators are now `Q₀` and
`P₀` rather than bare `Q` / `P`, in 1(a) and 1(d), so the reader can see
which end is the base without reconstructing it from the arithmetic.

**3. A house convention note**, `S.elasticity_note` /
`ELASTICITY_CONVENTION` in `_ps_theme.py` — the sibling of `change_note`,
same cream card. It sits under PS 2's 1(a); 3(a) cross-refers to it. Both
question stems now also say "and name the interval your elasticity refers
to". Untracked: a revision inside a text box never reaches the review pane.

**4. The figure.** `fig_coffee` now draws the interval as a **thicker
dark-red stretch of the demand curve** between the two observations, with
the **base solid and the other end hollow** (new `Panel.dot_open`, since
`Fig.dot` is solid and `_tn_theme` is not to be modified; also new
`Panel.span`). Both labels name their interval AND direction —
`E_D, $45 → $40 = −1.50` and `E_D, $40 → $45 = −1.14` — which says
outright that the SAME two observations give both numbers, differing only
in base. That is exactly the point 1(b) and 1(c) make. The span is emitted
before the dots and labels so it cannot paint over them.

Two prose notes were reframed to match: 1(a)'s "Note:" now says the interval
is what the elasticity describes and the initial point is only the base;
3(a)'s says "Both numbers describe the SAME interval; what differs is the
base we divide by."

**Verified:** PS 2 4 pages, solutions 7 — both unchanged. `_ruleaudit` 0
findings on both. Last-page tail 381 pt and 521 pt.

### 2026-09-16 – session close

Four days of work on the problem sets, all ten documents current and
rebuilt, `_ruleaudit` clean across the board.

**What this session did, in order:** adopted his PS 1 / PS 1 Solutions hand
edits; added the navy submission banner (later with the TA line) to all five
problem sets; wrote the supply-vs-demand terminology rule into
`Teaching/CLAUDE.md` after two corrections from him and swept all ten
documents for it; moved the PS 3 change columns to the far right and got the
full-cell fade working with a FLOATING anchored shape; moved marginal values
to interval midpoints in both the figures and the question text; and kept the
initial-point base for elasticities while making the interval visible in the
notation, the formula and the 1(c) figure.

**Two methods worth carrying forward**, both now in the folder `CLAUDE.md`:

- **Measure bottom clearance, never page count.** The page count said 2 while
  he was looking at 3, through every variation. The paragraph map is what
  found it.
- **Measure the fade, do not look at it.** A white-to-cream gradient is half
  invisible to a pixel test; repaint the stops red→blue in a copy and sample
  the raster.

**Open, all waiting on him:**

1. **20 tracked changes across six documents** – PS 3, 4, 5 and all three of
   their solutions. Once accepted, strip the markup from the build scripts so
   a fresh build reports zero.
2. **Three point-total discrepancies**, flagged in-document rather than fixed
   because a point allocation is his call: PS 3's parts sum to 110 against a
   100 header, and PS 5's Problem 3 sums to 40 against 35, making that set
   105.
3. **Eight pages end 2–3.7 inches early** where a figure or table would not
   fit. Two are near-misses (PS 3 Solutions p4, 15 pt; PS 5 Solutions p1,
   25 pt) that a small figure shrink would close.
4. **`change_table`** – the superseded `space_before` version – is still dead
   code in `_ps_theme.py`. The stagger has now been seen on paper, so it can
   go.

**Not in this folder:** the BruinLearn Assignments links went into
`405 Calendar and Website` (calendar, website, syllabus) and were committed
from another session on 2026-09-16 as part of `8e261fb`.
