# Session Notes – Course Syllabus (MGMT 405, Fall 2026)

Deliverable: `Course Syllabus - 405 <SECTION> Fall 2026.docx` (+ `.md`, and a
`.pdf` exported by hand), built by `_build_syllabus.py`.

```
python _build_syllabus.py                     # EMBA .docx
python _build_syllabus.py --md                # ... and the Markdown draft
python _build_syllabus.py --section femba --md
```

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
