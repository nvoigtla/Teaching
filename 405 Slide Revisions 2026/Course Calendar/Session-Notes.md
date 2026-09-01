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


## 2026-08-31 - practice-exercise quizzes, podcast label, exam boxes rebuilt

Added the TA's practice site to every week that follows a module, relabelled
the podcast groups, and rebuilt both exam weeks as a single dark-yellow box.

### Files worked on

| File | Change |
|---|---|
| `_calendar_content.py` | 10 `prac_*` links; a `"practice"` group in 7 weeks; podcast label; both exam dicts rewritten (title, window, lines) |
| `_build_calendar.py` | `"practice"` category card; `exam_window()` helper; exam card moved into the topics-card slot |
| `Calendar EMBA Hybrid -- Fall 2026.docx` / `.pdf` | rebuilt and re-exported |

### What changed in the document

1. **Podcast group label** is now `About Class Material:` (was `On class
   material:`) in all 8 weeks. Nico's capitalisation, kept verbatim; note the
   sibling label `Other podcasts:` is still sentence case.
2. **New card `Suggested Additional Practice Exercises`**, a fourth category
   after Suggested Reading, glyph **U+270E** (the decks' problem-set mark),
   navy header / white body. Registered in the card tuple in
   `_build_calendar.py`, so it renders wherever a week has a `practice` group.
3. **Ten links**, all live (HTTP 200 checked), labelled
   `Online quiz on Module N: <topic>`. Source:
   `https://rafaelrubiao.github.io/mgmt405-practice/` - `module-N.html`, with
   `-part-1` / `-part-2` for Modules 4 and 7.
4. **Placement rule (Nico's call, this session):** the week AFTER the module's
   own TEACHING week. Modules 3 and 6 follow their VIDEO week, not the later
   "Applications" week, so Module 3's set is available before the Week 6
   midterm that covers through Module 3.

   | Week | 2 | 4 | 5 | 7 | 8 | 9 | 10 |
   |---|---|---|---|---|---|---|---|
   | Modules | 1, 2 | 3 | 4-I | 4-II, 5 | 6 | 7-I | 7-II, 8 |

5. **Exam weeks rebuilt.** Weeks 6 and 12 previously showed three separate
   boxes (a top banner added earlier the same day, a plain topics card, and a
   detail card at the bottom). They are now ONE box in the topics-card slot:
   dark-yellow `E09F3E` header carrying the name only (`Midterm Exam
   (online)` / `Final Exam (online)`), cream body, bullets below. The topics
   card and the bottom exam card are suppressed on exam weeks.
6. **Exam windows moved:** midterm Thu-Sun -> **Fri-Sat, Oct 30 - 31**; final
   Fri-Sun -> **Sat-Sun, Dec 12 - 13**. Page 1 reads `Midterm window` /
   `Final Exam window`.
7. **Exam wording:** exact time window to be determined and announced in
   class; both exams online with proctoring software; both `Open book, open
   notes. Calculator allowed.`

### Decisions

- **One source for the exam dates.** They had been hard-coded in three places
  (p.1's Due/Exams column, `band_right_label`, the exam card). All three now
  call the new `exam_window(wk)`, which reads the exam's own `window` tuple -
  change a date once in `_calendar_content.py` and every surface follows.
- **Practice links carry the CALENDAR's module names**, not the site's bare
  "Module 1" link text, so they match the wording students see elsewhere.
- **The redundant sentence was dropped** from the final's question-count
  bullet ("The exam will cover all material.") because the new first bullet
  now says it. Flagged to Nico; one word restores it.
- **Week 12's own span stays Dec 11 - 13** even though the exam window is now
  Dec 12 - 13: Week 11 ends Dec 10, so shortening Week 12 would leave Dec 11
  in no week at all.

### Open items

1. **Module 8's practice page is titled just "Auctions"** (11 questions),
   while the calendar's Module 8 is "Asymmetric Information; Auctions". The
   link uses the calendar's wording. A draft email to Rafael asking whether
   the asymmetric-information half is covered was written in chat this
   session but NOT saved to a file - rewrite it if still needed.
2. **The midterm box states no dates** (only the week band and p.1 do), while
   the final's box spells its window out. Deliberate for now - the midterm's
   time is TBD - but the asymmetry is worth a look.
3. **No practice card in Week 11 (exam prep).** A consolidated all-modules
   card was offered and not taken; Week 12 cannot host one (no prep section).
4. Everything from the 2026-08-30 entry still stands: Module 4's five video
   links and running times are still placeholders.

### Commands worth remembering

- Rebuild: `python _build_calendar.py` (writes the .docx in place).
- One-page-per-week invariant: `_check_pagination.ps1` (Word COM). Run after
  ANY content change; it still passes at 14 pages with all seven new cards.
- PDF: Word COM `ExportAsFixedFormat($path, 17)`; there is no pandoc step.
- Eyeball a page without opening Word:
  `pdftoppm -f N -l N -r 100 -png "<pdf>" out` (MiKTeX ships pdftoppm; there
  is no PyMuPDF / pdf2image on this machine).
- Measure a label before trusting it to fit: PIL `ImageFont.truetype`
  on `C:/Windows/Fonts/calibrib.ttf`; practice-card text width is 5.566 in.

---

## 2026-09-01 - Module 4 podcast links, new Module 1 Video 1 session

Pasted in Module 4's two podcast links, read their running times off Dropbox,
and swapped Module 1's Video 1 for the re-uploaded Panopto session.

### Files worked on

| File | Change |
|---|---|
| `_calendar_content.py` | `PODCASTS[4]` intro + wrap links and minutes; `LINKS["m1v1"]` new session id; Week 1 `m1v1` length |
| `Calendar EMBA Hybrid -- Fall 2026.docx` | rebuilt three times (intro only, then both podcasts, then the new video link) |

### What changed in the document

1. **Module 4 podcasts are live links**, replacing the plain-text
   placeholders: intro **3:35 -> 4 min**, wrap-up **21:47 -> 22 min**, both
   read by `_podcast_minutes.py` straight off the Dropbox files. Modules 5 - 8
   are still placeholders.
2. **Module 1, Video 1: Introduction** points at the re-uploaded session
   `id=b7207fac-c4f6-45fd-be9c-b4b800fc0e3f` (was `45ebea5c-...a768`), length
   **8:53 -> 9 min**. Week 1's four videos now read 9 / 10 / 8 / 7 min.

### Decisions

- **The wrap-up file is named `Module-4-Video-Wrap-Up.m4a`**, not
  `Module-4-Podcast-Wrap-Up.m4a` like modules 1 - 3. It is an `.m4a` and Nico
  supplied it as the wrap-up podcast, so it is wired to the
  *Podcast: Wrap-Up of Module 4* bullet as-is. Renaming it on Dropbox would
  change the share URL, so the naming inconsistency stays.
- **A re-recorded video loses its old running time.** When the new m1v1
  session was still private, the length was set to `None` (prints `(++)`)
  rather than carrying the old 9 min forward - the old number described a
  different file. It happened to come back as 9 min once the session was
  shared.
- **No PDF this round** (Nico: "you can create only the word file"), so the
  committed `.pdf` now lags the `.docx`.
- Backups were rolled before each content-changing rebuild, but NOT before
  the one rebuild that only regenerated identical content - rolling there
  would have pushed a good version off the end of the chain.

### Open items

1. **The `.pdf` is stale.** Word's COM `ExportAsFixedFormat` hung twice
   (Word started, stayed Responding, wrote nothing; no modal dialog was
   visible in an `EnumWindows` dump). Killing `WINWORD` and exporting to a
   scratch path did not help either. Retry, or Save-As from Word by hand.
2. **Dropbox share URLs cannot be derived from the file name** - the `fi` id
   and `rlkey` are random per file. Two of the links pasted for Module 4 were
   byte-identical, so the wrap-up had to be requested separately. Worth
   checking the two URLs differ before building.
3. **`(++)` remains on 15 videos** - Module 2 (3), Module 3 (7), Module 4 (5).
   Their Panopto sessions are not shared publicly.
   `python _video_minutes.py m2 m3 m4` fills them in as each is shared.

### Commands worth remembering

- Podcast lengths: `python _podcast_minutes.py` (range-GETs the Dropbox
  `.m4a` and parses `moov/mvhd`; prints the minutes to paste into `PODCASTS`).
- Video lengths: `python _video_minutes.py m1` - `LOCKED (not shared
  publicly)` means the session needs its share setting changed, not that the
  link is wrong. Panopto's `oembed` endpoint 404s here and `Embed.aspx`
  carries no duration, so `DeliveryInfo.aspx` is the only route.
- Confirm a link actually landed in the built file:
  read `word/_rels/document.xml.rels` out of the `.docx` with `zipfile` and
  grep the session id / file name.
