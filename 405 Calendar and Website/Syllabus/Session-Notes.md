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
