# Session Notes – Teaching Notes (MGMT 405, Fall 2026)

Six teaching notes, each as a `.docx` (the source Nico edits) and a `.pdf`
(what students get):

| module | note |
|---|---|
| 2 | Marginal Revenue |
| 2 | Demand Elasticity and Total Revenue |
| 2 | Regressions |
| 3 | Bang for Buck Rule |
| 3 | Hiring Decisions in the Short Run |
| 5 | MR = MC |

## How they reach the students

The **PDFs** are published next to both course websites and linked from the
module and week pages. Nothing here needs to be copied by hand:

- `../../405 Calendar and Website/Course Website/_deploy.py` lists them in
  `DOCS`, with the published (space-free) name each one takes -- e.g.
  `MGMT-405-Teaching-Note-Marginal-Revenue.pdf`.
- The calendar's content module carries the matching links, so the week and
  module pages point at the published copies.
- `python _publish.py` (in `Course Website/`) checks every local link
  resolves before it publishes, so a note that is linked but missing fails
  the build rather than 404-ing live.

## When a note changes

Edit the `.docx`, **re-export the `.pdf` into this folder under the same
name**, then run `python _publish.py` from `Course Website/`. The published
name is derived from the file name, so renaming a file here means updating
`DOCS` in `_deploy.py` to match.

**Export the PDF with Word's `ExportAsFixedFormat`, having set
`ActivePrinter` to a local printer first** -- see the note in
`../../405 Calendar and Website/Course Calendar/Session-Notes.md`. Do NOT
print to the "Microsoft Print to PDF" driver: that flattens the file and
strips its hyperlinks.

## Where this stands (2026-09-06)

All six notes are current and live on both the EMBA and FEMBA sites.
