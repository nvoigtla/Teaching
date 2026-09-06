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
