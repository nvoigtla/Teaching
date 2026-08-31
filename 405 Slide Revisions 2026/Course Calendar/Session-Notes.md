# Session Notes – Course Calendar (MGMT 405 EMBA Hybrid, Fall 2026)

Deliverable: `Calendar EMBA Hybrid -- Fall 2026.docx` (+ `.pdf`), built by
`_build_calendar.py` from `_calendar_content.py`. Content file is the single
source of truth for dates, links, video/podcast lists; the build script owns
layout, palette and chrome.

---

## 2026-08-28 – recolor, podcasts, video names, one-page weeks

Reworked the calendar to match the slide decks' color coding, gave every
module its two podcasts, adopted the module 1 – 3 video names from the module
folders, and made "every week on one page" an enforced invariant.

### Files worked on

| File | Change |
|---|---|
| `_build_calendar.py` | category colors, dark-gray agenda header + white gap row, `"Video content"` label, `"p"` podcast item kind, `(++)` video lengths, week-page spacing constants, band-label width guard |
| `_calendar_content.py` | `PODCASTS` registry + `podcast_items()`, podcast groups in 8 weeks, module 1 – 3 video names and lengths, new module 1 Panopto links |
| `_podcast_minutes.py` | NEW – reads podcast durations off Dropbox |
| `_video_minutes.py` | NEW – reads Panopto durations via `DeliveryInfo.aspx` |
| `_check_pagination.ps1` | NEW – asserts one page per week (Word COM) |

### What changed in the document

1. **Category colors, matching the decks.** In class = navy `0B2B4E`
   (white text), videos = light yellow `F6E8C9`, exams = darker yellow
   `E09F3E`. Applies to the page-1 agenda rows, the legend, the Videos card
   body, the on-campus card (navy header + pale blue `E7EDF4` body) and the
   exam card header.
2. **Agenda header row** is dark gray `555B66` at 12 pt with an 8 pt white
   spacer row beneath it, so it reads apart from Week 1's navy row.
3. **"Video deadline (suggested)" retired** in favour of **"Video content"**
   in the legend and `KIND_META`. The week band keeps the date and reads
   `Video content  ·  suggested: Sun, <date>`.
4. **Podcasts: two per module**, replacing "Posted on Bruin Learn" –
   `Podcast: Intro to Module X` and `Podcast: Wrap-Up of Module X`.
5. **Module 1 – 3 class video names** taken from the module folders. Module 3
   was retaped as **7** videos (was 6).
6. **Every week fits on one page**; the deck is 14 pages.

### Decisions

- **Week band stays navy on every page.** It is page chrome, like a slide's
  top bar, not a category marker. Only the content cards carry the coding.
- **The band label keeps its date.** "Video content" alone would drop the
  information, so the wording leads with the new term and keeps the date.
- **Practice videos and module recaps were NOT touched** by the video-name
  and `(++)` pass. Only the videos that teach the module itself changed.
  Recaps still read 7 / 8 / 7 min.
- **Rounding** for all durations is `round(seconds / 60)`.
- **Podcast module to week mapping** (the week that covers that module):

  | Week | 1 | 2 | 3 | 4 | 5 | 7 | 8 | 9 |
  |---|---|---|---|---|---|---|---|---|
  | Module | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 |

### Open items

- **Module 2 (3 links) and Module 3 (6 links) still point at last year's
  Panopto sessions.** Module 1's were re-recorded as new sessions
  (`b4b3…` ids), so modules 2 and 3 almost certainly need new URLs too.
- **Module 3's link keys are shifted one topic down** (`m3v1` now sits on
  "Video 2: The Production Function", and so on), because the new
  "Introduction to Module 3" joined at the front. That video has **no link
  at all** yet. Re-check the mapping when the new URLs arrive.
- **Module 2 and 3 video lengths are `(++)`**. `_video_minutes.py` fills
  them in, but only for sessions shared publicly.
- **Podcast links for modules 4 – 8** are placeholders (plain text, no link,
  no duration).
- **Week 1 has only 0.22 in of slack**; every other week has 1.15 in or more.
  Anything added to week 1 will push it to a second page. Run
  `_check_pagination.ps1` after any content change.
- **`.gitignore` ignores `*_t-1.pptx` / `*_t-2.pptx` but not the `.docx`
  equivalents**, so this folder's rolling backups show up as untracked.
- **Panopto session titles vs. deck names**: Nico renamed module 1's videos 2
  and 3 on Panopto so they now match the calendar. Worth checking the same
  for modules 2 and 3 when their links land.

### Commands

```
python _build_calendar.py                    # rebuild the .docx
python _podcast_minutes.py                   # podcast durations (Dropbox)
python _video_minutes.py            # all linked videos (Panopto)
python _video_minutes.py m2 m3      # just those modules
powershell -File _check_pagination.ps1       # one page per week?
```

PDF is exported through Word COM (no LibreOffice on these machines):

```powershell
$w = New-Object -ComObject Word.Application; $w.Visible = $false
$d = $w.Documents.Open("<abs path>.docx", $false, $true)
$d.ExportAsFixedFormat("<abs path>.pdf", 17); $d.Close(0); $w.Quit()
```

### Context worth keeping

- **`OUT` is now derived from the script's own folder** (overridable with
  the `CALENDAR_OUT` environment variable). It used to be a hardcoded `c:\`
  path, so the build could not run from the H: mapping at all.
- **`python-docx` had to be installed** on this machine; it was not present.
- **Durations are read, not typed.** Podcasts: a range request for the
  MP4/M4A `moov/mvhd` box, so nothing large is downloaded. Panopto: the
  viewer's own `DeliveryInfo.aspx` JSON, which answers anonymously only
  while a session is shared publicly and otherwise reports `LOCKED`. Both
  scripts re-check the stored numbers and flag disagreements, so they double
  as verification after editing.
- **Card heights are computed by `_measure_par` from the same spacing
  constants that lay out the text.** Changing a spacing constant moves the
  drawn box with the text. The one number to leave alone is the 0.18 in
  cushion inside `rounded_card`: it is what stops the approximation from
  clipping text out of a card.
- **This repo has concurrent activity.** Several commits appeared during the
  session that were not made from it, including the Course Calendar commit
  itself and a large Module 4 rebuild. Check `git log` before assuming the
  working tree is yours alone.

---

## 2026-08-30 – Module 4 re-split into five videos

**One-line summary:** Module 4's deck was converted for taping with a
different split than the calendar carried, so the Session-4 prep block was
updated to list five videos with placeholder lengths.

### What changed
`_calendar_content.py`, the Session-4 `"video"` group for Module 4 (Part I):

    Video 1: Introduction to Market Structures      (++)
    Video 2: Perfect Competition                    (++)
    Video 3: Profit Maximization of a Price Taker – Short Run  (++)
    Video 4: Firm-Level and Market Supply           (++)
    Video 5: Long-Run Competitive Equilibrium       (++)

Was four videos with real running times (13 / 47 / 6 / 11 min).

### Decisions
- **Lengths are `None`, which the builder already prints as `(++)`** — the
  same placeholder Modules 2 and 3 use. It also suppresses the
  "≈ N min of video in total" line, which would be meaningless.
- **Links are `None` too.** The four Panopto URLs (`m4v1` – `m4v4`) point at
  the OLD four-video cut, where Video 2 was the profit-max video and there
  was no separate Perfect Competition video — linking them would send
  students to the wrong recording. The URLs are still in `VIDEO_LINKS`,
  untouched, ready to be re-pointed.

### Open
1. Re-record the five videos, then put the real minutes in as the fourth
   element of each tuple (~line 313) and re-point `m4v1` – `m4v5`.
   `_video_minutes.py` can read the lengths off Panopto.
2. **The PDF was not regenerated** — it still shows the old four-video
   version. Export it from Word once the deck is final.
