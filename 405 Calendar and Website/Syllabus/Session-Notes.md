# Session Notes – Course Syllabus (MGMT 405, Fall 2026)

Deliverable: `Course Syllabus - 405 <SECTION> Fall 2026.docx` (+ `.md`, and a
`.pdf` exported by hand), built by `_build_syllabus.py`.

```
python _build_syllabus.py                     # EMBA .docx
python _build_syllabus.py --md                # ... and the Markdown draft
python _build_syllabus.py --section femba --md
```

## The PDF export: ExportAsFixedFormat, with ActivePrinter set first

**Superseded advice below.** An earlier version of this note recommended
printing the document to the "Microsoft Print to PDF" driver, because
`ExportAsFixedFormat` was hanging. That workaround FLATTENS the file: the
calendar came out with 0 hyperlinks instead of 90, losing every video,
podcast and slide-deck link, because those sit inside drawn text boxes. It
was published in that state before being caught.

**What actually works:** set `$w.ActivePrinter` to a local printer BEFORE
calling `ExportAsFixedFormat`. Word wants printer metrics to lay a PDF out
and stalls when the default is unavailable -- that was the hang all along.
With that one line it returns in seconds and keeps every link.
`python _publish.py` does this, and `verify_pdfs()` fails any PDF with zero
hyperlinks so a flattened file cannot ship again.

The original note is kept below for the mechanics of driving Word over COM.

### Original note (do NOT use PrintOut for the real export)

`ExportAsFixedFormat` (and `SaveAs2` to wdFormatPDF) HANGS on this machine --
Word opens, stays Responding, burns CPU and writes nothing, with no modal
dialog visible in an `EnumWindows` dump. It failed on ~8 attempts across
three sessions, read-only and read-write, visible and invisible, to the
canonical path and to a scratch path, before and after killing WINWORD and
clearing the Resiliency key.

**What works: printing to the "Microsoft Print to PDF" driver** (2026-09-06).
A different code path, and it returns in seconds:

```powershell
$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
$w.ActivePrinter = "Microsoft Print to PDF"
$d = $w.Documents.Open($src, $false, $true)
$m = [Type]::Missing
# PrintOut(Background, Append, Range, OutputFileName, From, To, Item,
#          Copies, Pages, PageType, PrintToFile, Collate)
$d.GetType().InvokeMember("PrintOut", "InvokeMethod", $null, $d,
  @($false, $false, 0, $out, $m, $m, 0, 1, $m, 0, $true, $true))
$d.Close(0); $w.Quit()
```

Delete the target first (the driver appends to an existing file), and give
the print spooler a few seconds after `Quit()` before checking for the file.

**Verify the result with PyMuPDF, not a regex.** These PDFs use subset fonts,
so scanning the raw bytes or the inflated streams for "G305" returns False
even when the page plainly shows it -- the same trap that made an earlier
syllabus check look like a failure. `fitz.open(path)` and `page.get_text()`
read it correctly. PyMuPDF and `pdftoppm` (MiKTeX) are both installed;
LibreOffice is not.

## How it fits the rest of the folder

The syllabus is NOT a standalone document. It reads two things from the
course calendar so nothing is typed twice:

- **Content and addresses** from `../Course Calendar/_calendar_content.py` —
  the term, the room, the class times, the TA's name, and every URL in
  `LINKS`. The section table there also supplies the output filename and the
  subtitle, which is why `--section femba` is all it takes to produce the
  other section.
- **Layout helpers** are IMPORTED from `../Course Calendar/_build_calendar.py`
  (palette, `rounded_card`, `add_hyperlink`, the table helpers), so the
  syllabus, the calendar and the website read as one family. That module does
  all its work inside `main()`, so importing it is safe.

`--section` is parsed into `os.environ["MGMT405_SECTION"]` in a loop placed
**above** the imports. `_calendar_content` reads that variable at import
time, so moving the loop below an import would make the flag silently do
nothing. Same pattern in all four builders.

## Two standing content decisions (2026-09-04, Nico)

1. **No e-mail addresses anywhere in this document.** The PDF is published on
   the public course website, and the website deliberately obfuscates both
   addresses against harvesters; printing them in a public PDF would undo
   that. The syllabus points at the website's "Class and Contact" box and at
   Bruin Learn instead.
   - Re-confirmed 2026-09-06. When Nico sent the two section-specific TA
     addresses and asked to use them "throughout", the addresses were added
     here and then **reverted** — he meant the places we had already agreed
     on (calendars and websites), not the syllabus. Ask before changing this.
2. **Achieve is gone.** Practice exercises are the TA's own site, which is
   what the calendar and the website link.

## Wording that follows the section

`MEETING_SENTENCE`, `MEETING_HEADING` and `FIRST_MEETING_DAY` are derived
from `C.MEETINGS`, so the prose follows the meeting pattern rather than
repeating it:

| | EMBA | FEMBA |
|---|---|---|
| heading | On-campus weekends | On-campus sessions |
| sentence | Fridays 4:00 – 5:30 pm and Saturdays 9:00 am – 12:30 pm | Saturdays 2:00 – 8:00 pm |
| video deadline | "…or the **Friday** class" | "…or the **Saturday** class" |

Both sections meet in **G305** (2026-09-06; was A301 / G-402).

## Where this stands (2026-09-06)

Both syllabi are current and correct: EMBA and FEMBA, `.docx` and `.md`.

**Open: the PDFs.** Word's COM `ExportAsFixedFormat` has hung on every
attempt across three sessions — Word starts, stays Responding, burns CPU and
writes nothing, with no modal dialog visible in an `EnumWindows` dump.
Killing `WINWORD` and exporting to a scratch path does not help either. Nico
exports them by hand (2026-09-06: "do one attempt to generate it. If that's
not working let me know and i'll generate it myself"). The FEMBA syllabus PDF
does not exist yet, which is why the FEMBA website is published with
`_deploy.py --skip-docs` and its syllabus link 404s until the file is added.

One known cause of a hang, worth ruling out first: the target PDF being open
in a viewer. Word then blocks on an invisible overwrite prompt.

## Problem sets move to BruinLearn (2026-09-12)

Nico: "the problem sets will be downloadable only from the BL site." Every
mention of PS 1–5 now points at that section's **Assignments** page, and the
same link serves the download and the upload.

**One derived constant, not two literals.** `_calendar_content.py` gained

```python
"bruinlearn_assignments": SEC["bruinlearn_course"] + "/assignments",
```

next to `"bruinlearn_course"`. Hardcoding the pair he sent (EMBA
237825 / FEMBA 237860) would have re-created the 2026-09-06 bug recorded a
few lines above it, where a literal EMBA address made every FEMBA build
*print* the FEMBA address and *link* the EMBA site.

**Three call sites:**

- `Course Calendar/_build_calendar.py` – the due card's "Upload one solution
  per group on BruinLearn", was linking the course root.
- `Course Website/_build_site.py` – the same line on the week cards, plus a
  new `BRUINLEARN_ASSIGNMENTS` beside `BRUINLEARN_COURSE`.
- `Syllabus/_build_syllabus.py` – the sentence that used to say slides,
  problem sets and solutions were all on the class website, which the new
  policy made false. It now reads, in his words: "Electronic copies of all
  our slides are on the class website. You find the Problem Sets and Problem
  Set Solutions on BruinLearn under “Assignments.”" – with
  `BruinLearn under “Assignments.”` carrying the link.

The BruinLearn *class-site* panel on the website still points at the course
root, which is right – it is a link to the site itself, not to the problem
sets.

**Verified, both sections:** 5 assignments links per site (weeks 3, 5, 7, 9,
10 – the five problem sets), 5 per calendar `.docx`, 1 per syllabus; EMBA
outputs carry 237825 and FEMBA 237860 with no cross-contamination. PDFs
re-exported and the links survive into them. `_check_pagination.ps1` PASSES
at 14 pages, every week on one page. Syllabi 5 pages each.

**Not done:** `_publish.py` / `_deploy.py` were NOT run – nothing is live yet.


## Attendance card, recordings, the page-1 BruinLearn block (2026-09-15/16)

**"Do the readings…" split into three bullets**, and a **dark-red card**
(`missed_class_card()`, a new `("redcard", …)` block kind) now wraps "If you
have to miss an on-campus class:" and its five bullets. One of those bullets
links that section's class recordings — `LINKS["bruinlearn_recordings"]`,
built from the course URL plus `/external_tools/` and the section's tool id
(EMBA 10996, FEMBA 10995).

**Problem-set submission** points at the Assignments page; the general
BruinLearn references still point at the course root.

**The page-1 yellow card** gained a BruinLearn entry, over three rounds of
his edits. It now reads: Course Website, the address, the one-line
description, then the *Prefer a PDF…* caption — which belongs to the website
half — then a 12 pt gap, then **BruinLearn Course Site**, the full address
`bruinlearn.ucla.edu/courses/<id>`, and "On the BruinLearn site, you can
download and submit the problem sets and find the recordings of the
on-campus classes." The address comes from `C.BRUINLEARN_TEXT`, derived from
the link, so the two cannot drift. This is also the syllabus's only link to
the course ROOT — both other BruinLearn links are problem-set ones.

**A bug worth remembering.** `("grades", [list(r) for r …])` broke
`write_md`, because `"| %s | %s |" % r` unpacks a tuple but not a list. The
Markdown build had been failing **silently** since the `GRADE_WEIGHTS`
refactor, and an earlier "syllabus text unchanged" check was invalid: stderr
was suppressed and the grep ran against a stale `.md`. `GRADE_WEIGHTS` rows
stay **tuples**.

**Verified:** both syllabi 5 pages, 16 links, correct section — the extra
link over the previous 15 is the new course-root one.
