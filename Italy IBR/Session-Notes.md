# Italy IBR (MGMTEX 421, Class 1 – "The Italian Economy") – Session Notes

Continuity log for the International Business Residential (Italy) deck. The
course is **MGMTEX 421 – International Business Residential in Italy**; this
deck is **Class 1**, the online big-picture class (first class held online;
the trip to Milan & Turin is in September). Audience wording everywhere:
"executives pursuing an MBA at UCLA Anderson."

Canonical deck: **`Class 1 - Revised.pptx`** (widescreen, 405 visual language).
Original source (untouched): **`Class 1.pptx`** (4:3, 123 slides).

---

## Resume here (start of next session)

- **Status:** deck is **140 slides**, opens cleanly, all work committed and
  pushed (last session 2026-07-26). This folder lives at `Teaching/Italy IBR/`.
  The single design/build rulebook is `Teaching/CLAUDE.md`.
- **Current display map:** 21 = hand-inserted Caesar slide; 22 = Empire-length
  poll; 25 = Golden Period; 37 = hand-inserted social-capital bridge (leads to
  Guiso et al. at 38–39); 52 = Featured-Research divider (Giorcelli–Moser),
  53 = copyright-question content slide (user swapped these); 102–106 = case
  block (Lavazza 102/103, GBU 105/106); 110 = Thank You; 111–140 = backups
  (113 = Presentation Topics, Lavazza chip struck out in red).
- **Workflow reminders proven this stretch:** whenever the user inserts /
  swaps slides by hand, shift/swap the display-keyed config in BOTH
  `_notes_data.py` and `_animate.py` (NOTES/BACKUP/ANECDOTE keys + SKIP sets +
  FIG_GROUP); before any `_add_notes.py apply`, run the full deck-vs-data
  sync check and ADOPT hand-edited notes into `_notes_data.py` (caught user
  edits on notes 15, 20, 51, 106 this way); a replaced picture usually drops
  out of its animation beat — rebuild that slide's timing with
  `AN.timing_xml` (picture + caption + its bullet on one beat).
- **The deck is the source of truth** (edited in place via OOXML; never
  round-trip through python-pptx — it strips poll `tags` + null links). Only
  the **reusable engine scripts** remain in the folder (the one-off build
  scripts were deleted — recoverable from git history):
  - `_animate.py` — rebuild animations after any structural change:
    `python _animate.py all apply` (or specific slides). Holds `FIG_GROUP`
    (figure→bullet) and `PIC_BULLET` (per-picture→bullet, e.g. slide 6).
  - `_group_shades.py` / `_group_boxes.py` — group figure+shade / colored
    box+text; **re-run `_animate.py` afterward** (grouping wipes timing).
  - `_resize_bullets.py apply` — re-fit bullet sizes.
  - `_add_notes.py apply` + `_notes_data.py` — edit speaker-note text in
    `_notes_data.py`, then re-apply (edits notesSlides in place; poll notes
    are protected).
- **Open items:**
  - **Commit + push the session-3 edits** (deck + `_notes_data.py` +
    `_animate.py` + this file). User to confirm.
  - Delete the empty stale folder `405 Slide Revisions 2026/Italy IBR/` — it
    was locked while the prior session was running; remove it now that this
    session is closed.
  - **Slide 132** (poll "Italian-Speakers in 1861?") is an orphan backup that
    nothing links to → it got **no** Back button. Add one → slide 61 if wanted.
  - Podcast **Parts I–IV** emails are drafted (link placeholders) — paste the
    NotebookLM links and double-check the per-part summaries.
  - **Slide 45's header line** ("Milan Cathedral, Marciana…") is partly hidden
    behind the two top pictures — offer to fix the layout.
  - Optional, previously flagged: slide **20** has no figure (too text-dense —
    could split); slide **47**'s plague image is Naples 1656 (captioned so);
    slides **49 / 100 / 101** reveal their image clusters on a single click.

---

## 2026-07-26 – case answers, new slides 21 + 37, hand-edit integration, notes voice pass

- **Case-discussion answers** added to notes 103 (Lavazza: stay-Italian vs.
  adapt with the Lavazza-Professional 2-brand middle path; Starbucks/Milan
  mirror-image = cultural embeddedness; Factory 1895 question candidates) and
  105 (GBU: good/bad/ugly sorting from case p. 1; deepest-roots vs.
  government-fixable split, Renzi referendum epilogue). Matched the user's
  reworded GBU question 2 ("what could a government realistically fix").
- **New slide 21 (Caesar, user-inserted):** styled both artworks (Camuccini
  1806 painting — user re-pasted the correct image; Ferrucci bust c. 1512–14,
  Met), captions beneath each, detailed assassination + rise-of-Caesar note.
  Display shift +1 from 21 in both engine scripts. "Ceasar" typo fixed after
  user approved a standing CLAUDE.md exception for unambiguous misspellings.
- **New slide 37 (social-capital bridge, user-inserted):** picture styled,
  note ties the 3 panels (regulation/taxation/courts → social capital) to the
  Guiso–Sapienza–Zingales paper next. Display shift +1 from 37. Page-number
  scan: all footer numbers are live slidenum fields (only 1, 2, Thank-You
  unnumbered, intentional).
- **Notes content:** 17 Etruscans-vs-Latins (incl. Posth et al. 2021 aDNA);
  20 Republic institutions (Senate seats: censors, life tenure, de-facto
  hereditary nobiles, "new man"); 23 Octavian vs. Antony & Cleopatra
  (Philippi, Actium, Egypt annexed); 24 Pantheon (Agrippa→Hadrian, oculus,
  609 AD church, Raphael + first 2 kings); 26 publicani expanded (societates,
  partes, lex Claudia 218 BC, NT "publican"); 32 Normans + top-down South
  (Roger II 1130, template through 1806) and long-run divide (Putnam →
  Guiso et al.).
- **Voice rule:** all "Let me walk you…"/"I want you to…" lecture-management
  phrasing purged from every note (8 instances); "No lecture-management
  phrases" rule added to Teaching/CLAUDE.md.
- **Slide edits:** 5 + 44 + 45 + 3 + 93 new/replaced pics styled (rounded +
  shadow; 44/45/93 rebeat so pic + caption + bullet co-reveal; 45 got
  "Venetian Arsenal" caption); 16 "Passed on…" demoted to 20 pt sub-bullet;
  32 North/South point moved into a gold takeaway bar (timing rebuilt, bar =
  final click); 34 new Venice pic styled + new one-line bullet "Milan a
  leading commune – Turin under Savoy rule" (verified: Milan yes, Turin
  Savoy); 62 emigration photo styled + "Italian emigration" caption,
  co-revealing with the last bullet; 113 Lavazza chip struck out with 2 red
  C00000 lines; 52↔53 user swap mirrored in notes keys + SKIP_FEATURED.
- **Hand-edits adopted into `_notes_data.py`:** notes 15 (caveat wording +
  "Twenty-8"→"28"), 20 (user's institutions note), 51 (new Napoleon-at-Lonato
  anecdote), 106 (paragraph split). One incident: a chained apply briefly
  overwrote note 15 before the adopt — restored from the same-minute backup;
  lesson: run the sync check and STOP before applying.
- **Slide 101 sourcing:** 156,000 Italian emigrants 2024 verified against
  ISTAT "Migrazioni…Anni 2023/2024" (June 2025 release) — full citation in
  the sourcing block above the 2026-07-24 section.

## 2026-07-24 – summary-slide "why" bullets, slide-45 header fix, check-first rule

- **Slide 103:** added two bullets resolving the paradox ("Brands: design,
  craft, and heritage…" / "Productivity: family firms promote on loyalty, not
  merit — and missed the IT payoff"); animations rebuilt (7 beats); note
  extended accordingly.
- **Slide 45:** header now visible above the pictures (pictures moved down)
  and names all three buildings incl. St. Peter's.
- **Slide 115 verified:** title and note both correctly say **San Leo** (the
  2026-07-23 fix went old-note-"San Marino" → "San Leo", matching the slide);
  removed the "near San Marino" phrase from the note to avoid confusion.
- **Slides 54/59/60 (co-reveal + caption proximity):** 54 — map, source
  line, and legend now fade in together on ONE click (hand-built timing);
  59 — combined bottom caption split into per-picture captions, each
  revealing WITH its picture (3 clicks); 60 — caption moved right under
  the Fattori painting. `_animate.py` got `SKIP_CUSTOM = {54, 59}` so
  "all apply" preserves the hand-built timing. Three rules added to
  `Teaching/CLAUDE.md` (figure+label same click; one caption per picture;
  captions right underneath). **Flagged:** 59's third painting (bottom)
  has no description — looks like the Breach of Porta Pia (Rome, 1870);
  ask Nico before captioning.
- **GBU discussion slide + PNRR note (2026-07-25) — deck now 138 slides.**
  New display **105 "Case Discussion: The Good, the Bad, and the Ugly"**
  (before the ten-years-on slide, which is now 106): 2 questions
  (good/bad/ugly sorting from the case; deepest historical roots vs. what
  Renzi's 1,000 days could fix), the Sergio Leone film-poster artwork Nico
  pasted (matched + downloaded from blacklight.com product image, 660×1000,
  `image137.jpg`; credit line "Film poster: United Artists, 1966"; rounded +
  shaded), Discussion-Break badge + Case-Study pill (template = Lavazza
  discussion slide). **PNRR explained in slide 106's note** (Piano Nazionale
  di Ripresa e Resilienza; NextGenerationEU; ~€194bn; installments tied to
  reform milestones = commitment device). Displays ≥105 shifted +1
  (Summary 107, Thank-You 108, backups 109–138, poll backup 137;
  `SKIP_CUSTOM` ten-years-on now 106); notesSlide number derived from the
  zip this time (121).
- **Case-slide polish (2026-07-25):** Nico's hand-swapped pics on 102/103
  rounded + shaded (102's re-bound into its animation; 103's static by
  design); navy **"Case Study" pills** (Back-button style, x 11.55/y 0.60)
  on 102/103/105. **Slide 105 got a 5-PM timeline strip** under the question
  card: Renzi 2014–16 · Gentiloni 2016–18 · Conte 2018–21 · Draghi 2021–22 ·
  Meloni 2022– (portraits from Wikimedia/governo.it, CC BY / attribution;
  `image129–133.jpg`; rounded + shaded, static like the question card);
  bullets moved down at 20 pt; source line credits the portraits; note names
  the 5 PMs. **Gotcha logged:** scratch script `numbers.py` shadowed the
  stdlib `numbers` module and self-executed on import (renamed to
  `numbers_pass.py`; its stray re-edit of _notes_data reverted).
- **Case slides inserted — deck now 137 slides.** Two cases uploaded (Lavazza,
  Ivey W27255; "Italy: The Good, the Bad and the Ugly", HBS 9-716-029) are
  integrated:
  - **102 "Lavazza: Italian Capitalism in One Company"** (facts: Turin 1895,
    4th-gen family firm, 140 countries / 70% abroad, 30 yrs US → <1.4% share /
    $110.8m, ~€1bn bank loans; Vegas Caffè-Lavazza photo, S. Stierch CC BY
    4.0, `image128.jpg`; takeaway bar "One firm, all the course themes").
  - **103 "Case Discussion: Lavazza in the US"** (3 questions: convert vs
    adapt; Starbucks-Milan mirror; question for Factory 1895) + gold
    "Discussion Break" parallelogram badge.
  - **105 "Ten Years On: The Bad Improved — the Ugly Remains"** (after the
    Paradox): cream question card (static) + 6 verified changes on clicks —
    2016 referendum/Renzi resigned; youth unemployment 42.7%→20.3% (Eurostat
    une_rt_a); debt 134.8%→154.4% (2020)→134.7% (2024) / 137.1% (2025)
    (Eurostat sdg_17_40); PNRR €194.4bn = €71.8 grants + €122.6 loans (EC RRF
    page); 2025 upgrades S&P BBB+ (Apr 11), Fitch BBB+ (Sep 19), Moody's Baa2
    (Nov) (Il Sole 24 Ore/S&P); productivity ~flat since mid-1990s (OECD/IMF/
    EC consensus). Custom timing; slide in `SKIP_CUSTOM` (now {54,59,61,72,
    105}).
  - Displays ≥ old-102 shifted +3 (Paradox→104, Summary→106, Thank-You→107,
    backups 108-137, poll backup 136); `_notes_data.py` and `_animate.py`
    shifted accordingly; teleprompter notes written for all 3 new slides.
  - **Corruption incident + fix:** the insert initially broke the deck
    (0x80CB8001) — `notesSlide117.xml` already existed (PowerPoint created it
    for part slide134 during Nico's save), so the script's overwrite + a
    duplicate Content-Types override made the package invalid. Fixed by
    restoring the original 117, moving the new note to `notesSlide120.xml`,
    and rebuilding Content-Types from the backup's. Lesson: **derive the next
    free notesSlide number from the zip, never assume** (PowerPoint creates
    notesSlides on save).
- **Digits pass over all notes:** 76 spelled-out numbers converted to digits
  deck-wide ("twenty-eight"→28, "fifty thousand"→50,000, "sixteenth
  century"→16th century, …), with guardrails: "one" kept as a word,
  discourse ordinals kept, "Two Sicilies" protected; 5 mixed forms cleaned
  by hand ("1 to 5 to 28", "2 to 1", "11th to 13th centuries"). Preference
  added to `Teaching/CLAUDE.md`. Note the slide-15 citation detail also
  lives in the slide's speaker notes now (not just this log).
- **Slide 15 sourcing — the ~28× land-vs-sea figure.**
  - On the slide: source line "Source: Diocletian's Price Edict (AD 301), as
    computed by Duncan-Jones (1974)" under the takeaway bar, co-revealing
    with the "Rugged terrain" click. In the note: "…a ratio the economic
    historian Richard Duncan-Jones computed from the edict's freight
    schedules."
  - **Full citation:** Richard Duncan-Jones, *The Economy of the Roman
    Empire: Quantitative Studies*, Cambridge University Press, 1974 (2nd ed.
    1982), chapter on transport costs (the transport-price calculations from
    the Edict, ~pp. 366–369).
  - **Where the number comes from:** Diocletian's Edict on Maximum Prices
    (AD 301) sets maximum freight charges for specific sea routes and for
    wagon transport per mile. Duncan-Jones computed relative costs per
    ton-mile from those schedules and arrived at the much-quoted ratio of
    roughly **sea : inland waterway : road ≈ 1 : 5 : 28**. The same edict
    underlies A.H.M. Jones's companion formulation (*The Later Roman
    Empire*, 1964) that it was cheaper to ship grain the length of the
    Mediterranean than to cart it about 75 miles inland; Stanford's ORBIS
    transport-cost model (Scheidel & Meeks) works with ratios of the same
    order of magnitude.
  - **Caveats:** the ratio is derived from a price-ceiling document, not
    observed market transactions, and depending on cargo, route, and wagon
    vs. pack-animal assumptions, scholars' land-to-sea multipliers range
    from roughly 20× to 60×. So "~28×" is the defensible textbook number,
    best presented as "land cost more per mile by an order of magnitude,
    roughly thirty times" — which is how the slide uses it.
- **Slide 101 sourcing — the "156,000 Italians emigrated in 2024" figure.**
  - **Full citation:** ISTAT, "Migrazioni interne e internazionali della
    popolazione residente – Anni 2023/2024" (report + press release,
    published June 2025). Expatriations of **Italian citizens** in 2024:
    **156,000**, up 36.5% from 114,000 in 2023 — the highest value recorded
    since 2000, above even the Great-Recession peaks. Total cancellations
    for abroad (incl. foreign citizens) were ~191,000 in 2024. Top
    destinations: Germany 12.8%, Spain 12.1%, UK 11.9%.
    https://www.istat.it/comunicato-stampa/migrazioni-interne-e-internazionali-della-popolazione-residente-anni-2023-2024/
  - The slide's source line "ISTAT (2024)" refers to the data year; the
    release itself is 2025.
- **Full notes revision pass (all 109 notes):** extracted the notes from the
  DECK (so hand edits were the base), revised via subagent, wrote back into
  `_notes_data.py` (NOTES/BACKUP dicts rebuilt) and re-applied. 37 notes
  changed: ~25 stage directions removed/converted to speakable words ("Point
  to the map…" → "Look at the map with me…"), 13 substantive expansions
  (Etruscan absorption, Republic breakdown, coin debasement, Norman conquest,
  Cape-route shift, Code civil, Congress of Vienna, IRI origin, biennio
  rosso, Salò, etc.), live-moment words made time-neutral. Poll notes,
  Source lines, and slide 17's Celts line untouched. Two new rules added to
  `Teaching/CLAUDE.md` (verbatim speech only / substance over gloss).
- **Deck now 134 slides** — Nico hand-inserted a second Matera backup at 111
  (Matera sequence 110 "1/2" → 111 "2/2" + Back → 14); all displays ≥111
  shifted +1 (poll backup now 133, Roman-roads now 134). `_notes_data.py`
  BACKUP keys shifted, new 111 entry added, `SKIP_POLL` 132→133 in both
  `_notes_data.py` and `_animate.py`.
- **Batch (14/15/110/111):** 14 — Valcamonica background appended to the note
  (UNESCO 1979, Camunni, and the "ancient astronaut" reading vs. the
  halo/headdress reading). 15 — new sub-bullet "Roman price data: moving
  goods by land cost ~28× more per mile than by sea" (Diocletian's Price
  Edict ratio, Duncan-Jones); note appended; animation rebuilt (4 beats);
  body box shrunk to clear the takeaway bar. 110/111 — "Sassi di Matera"
  headers added (title+rule+strip cloned from 115); 110's photo moved below
  the rule; 111 got the two requested bullets (caves inhabited until the
  1950s; film stand-in for Jerusalem) and its two hand-added photos received
  deck-standard rounding/shadow; new note written for 111. Hand layouts
  (pager boxes, photo placement) preserved.
- **Batch (59-note/61/99/132):** 59 — note expanded (surgically, only that
  notesSlide + mirrored in `_notes_data.py`): why Garibaldi still mattered
  after 1860 (Aspromonte 1862, Mentana 1867) and why Rome fell only in 1870
  (French garrison withdrawn for the Franco-Prussian War; Porta Pia taken by
  regular bersaglieri, Garibaldi then fighting for France). 61 — gray
  "Poll: Italian-speakers in 1861?" pill → 132; custom build: D'Azeglio card
  + poll link show WITH the slide, all text bullets animate afterwards
  (3 clicks; 61 added to `SKIP_CUSTOM`, now {54, 59, 61, 72}). 132 — Back
  button → 61, placed just above its POLL badge (which occupies the standard
  corner); poll embed untouched. 99 — "New research" bullet now carries the
  inline cite "(Anelli et al. 2023)", matching the existing source line.
  **Hand edits preserved:** deck edited in place only on the named shapes;
  notes NOT re-applied deck-wide; no script rebuilds.
- **Batch (59/71/72/76/97):** 59 — third painting identified from its
  signature ("…Quaedvlieg, Roma, 20 Sett. 1870") as the **Breach of Porta
  Pia**, captioned "The Breach of Porta Pia, 1870 (Quaedvlieg)", co-revealing
  with its picture. 71 — text box narrowed to 6.7", the two ruins photos
  enlarged + rounded/shaded, ERP shield enlarged but kept flat (logo
  exception). 72 — collage re-laid out below the header, per-figure captions
  added (cartoon / West-Berlin billboard / Italian ERP poster), custom
  timing (added to `SKIP_CUSTOM`, now {54, 59, 72}); **flagged: the Italian
  poster carries a Bridgeman watermark** — consider sourcing a clean copy.
  76 — caption moved right under the Fiat photo. 97 — added Cremona
  luthier's-workshop photo (Hildegard Dodel, public domain, Wikimedia;
  `image127.jpg`), rounded/shaded + caption, text narrowed to 8.6",
  animations rebuilt (photo rides with the "family firms" beat).
- **Slide 120:** the single combined caption (bottom-right, clipped by the
  Back button) replaced with a caption under each figure; both figures got
  rounded corners + drop shadow; the Lucia libretto moved up (top-aligned
  with the playbill) so its caption clears the Back button.
- **New rule in `Teaching/CLAUDE.md`:** never change slide titles or "correct"
  factual content (slides, captions, notes) without checking first — flag as a
  proposal in chat instead.
- **Slide 54:** added a grouped legend card (white, thin navy border, soft
  shade) right of the map with **"Treatment: Early Copyright" / "Control:
  Late Copyright" headers** and a
  per-region swatch (sampled from the map) + copyright year: Lombardy — 1801,
  Venetia — 1801; Papal State — 1826, Two Sicilies — 1828, Sardinia
  (Piedmont) — 1840, Parma & Modena — 1840 (merged: Modena is too small on
  the map for its own swatch; both joined the 1840 treaty), Tuscany — 1840.
  Map moved 0.75" left, given rounded corners + drop shadow, caption
  re-centered under it. **Tuscany recolored** in the map image
  (mint green → tan, `image48.png` edited in place; source line notes
  "Tuscany recolored for clarity") so green = treated with no exceptions.
  Timing verified against Giorcelli–Moser (NBER w26885 / JPE 2020, Sec. 1):
  Lombardy AND Venetia under Cisalpine copyright law May 1801 (Venetia
  despite formal Austrian rule until 1805, fn. 11); controls under French
  rule 1804–1812 with code civil but NO copyright (Sardinia 1804, Parma
  1805, Tuscany 1809, Naples/Papal 1812); controls adopt copyright 1826
  (Papal), 1828 (Two Sicilies), 1840 (Sardinia–Austria treaty; Tuscany,
  Modena, Parma join). Flagged, not changed: slide 53's "Rest of Italy
  annexed in 1805" compresses the staggered 1804–1812 annexations.

---

## 2026-07-23 (session 3) – art notes, featured-research chrome, Roman-roads backup, summary slide

- **Pager boxes:** backups **121–130** now show "1/11"…"10/11" (navy pill,
  1.0" wide, replaces Back; slide-57 links removed); **131** has "11/11" above
  its Back button. **Slide 106** got a "Villa Campari →" pill above the Monday
  "Aperitivo & welcome" cell → 108; **108** got a Back button → 106.
- **Art identification + teleprompter detail** added to notes on slides
  **16** (Sarcophagus of the Spouses), **19** (Capitoline Wolf – also new
  caption under the picture), **22** (Augustus of Prima Porta), **23**
  (Pantheon/Colosseum), **34** (Lorenzetti), **42** (Brunelleschi's dome),
  **44** (Leonardo pieces incl. Last Supper), **45** (now names all three
  buildings incl. St. Peter's), **47** (Micco Spadaro, Naples 1656), **52**
  (all three = La Scala). **114** identified as the **Chimera of Arezzo** —
  got a title, caption, and an expanded backup note. Backup note for **115
  corrected: San Leo, not San Marino**.
- **Featured research:** journal+year added under authors on **36** (JEEA
  2016), **51** (JPE 2020), **78** (AER 2019) — 88/96 already had lines
  (96 stays "NBER WP 23964"; Crossref shows no journal publication).
  Animations **removed** on 36/51/78/88/96 (`SKIP_FEATURED` added to
  `_animate.py`). Gold **"Featured Research" badges** (x>11.5" so the
  animation engine treats them as chrome) on member slides **37–40, 52–55,
  79–87, 89, 97**.
- **New backup slide 133:** Roman road network map (Andrei Nacu "Roman Empire
  125", Wikimedia PD; `image126.png`), Back → 24; **slide 24** got a gray
  link pill "Map: the Roman road network" → 133.
- **Slide 117 orphan resolved:** it's the year-1100 History-of-Italy map; the
  word **"Normans" on slide 31** is now a text link → 117.
- **Pictures:** **49** Napoleonic-map stack enlarged ×1.55 + rounded/shaded;
  **64** (hand-swapped Lingotto photo kept) rounded/shaded; **66/67** got
  name captions ("Benito Mussolini" / "Mussolini and Hitler").
- **Slide 103** rebuilt: Quiz → **"Summary and Take-Aways"** (5 bullets +
  gold takeaway bar; section tag "Italy IBR · Wrap-Up"; new teleprompter
  note; animations rebuilt: 5 beats). Old quiz wording is in git history.
- Animations re-run on 66/67/103 only; notes re-applied deck-wide via
  `_add_notes.py apply` (108 slides).
- **Continuity note:** the chat transcript does *not* carry into a session
  opened from this folder — this file is the handoff.

---

## 2026-07-23 (session 2) – map, box grouping, back buttons, folder move, cleanup

- **Slide 15:** replaced relief map rounded + shadowed; normalized its messy
  embedded media name (`image8.jpg&w=828&q=100` → `relief15.jpg`).
- **Colored box + text grouping:** grouped each takeaway bar / card with its
  text into one `<p:grpSp>` on slides **15, 25, 29, 43, 64, 90, 91, 102** (fixes
  "box appears before its text"); updated `_animate.py` to detect the groups
  (`TakeawayGroup` = final click; `CardGroup` = column beats) and re-ran
  animations. Added the rule to `Teaching/CLAUDE.md`.
- **Slide 6:** each hotel photo now fades in with its bullet (top → "Hotel in
  Milan", second → "Hotel in Turin") via the new `PIC_BULLET` map.
- **Back buttons:** added to backups **115–117 → 31, 119 → 53, 120 → 55,
  121–131 → 57** (cloned from slide 114). **Deleted empty slides 133–135**
  (deck 135 → 132).
- **Structure:** consolidated the course-layer CLAUDE.md into `Teaching/CLAUDE.md`
  and moved `Italy IBR/` up to `Teaching/`; Session-Notes are now per-subfolder.

---

## 2026-07-23 – Consolidated log for the Class 1 build (covers 07-18 → 07-23)

**One-line summary.** Rebuilt the old "Class 1" deck into the 405 visual
language (135 slides), enriched the economic-history content with new slides
and figures, refreshed the presentation-topics slide from Excel, wrote
teleprompter-style speaker notes for the whole deck, added Fade/on-click
animations, and grouped table shades into their figures. Deck committed and
pushed once (`e0fa35d`); later edits (shade grouping, CLAUDE.md consolidation,
this notes restructure) are still local.

### Current deck state
- **`Class 1 - Revised.pptx` — 132 slides** (was 135; 3 empty trailing slides
  deleted 2026-07-23), opens cleanly in PowerPoint.
- Preserves the **7 live PollEverywhere embeds** (slides 11, 13, 21, 35, 48,
  69, 132) and all **internal slide-jump links** (e.g. slide-4 "Schedule" →
  trip itinerary; company/topic → backup detail slides; "← Back" buttons).
- Footer page numbers are **live `slidenum` fields**.

### What was done (grouped)
1. **Reformat.** Rebuilt every script-buildable slide in the 405 style
   (navy/gold chrome, Calibri, three-level section tags, live footer fields,
   native tables). Kept original slide titles and bullet wording; converted
   4:3 → widescreen 13.33×7.5.
2. **Economic-history enrichment** (from the podcast research): new slides —
   geography/rugged-map (15), Rome "First Integrated Market" (25), Empire at
   Height map (26), the Fall/fragmentation (28–29), **Renaissance Economy
   (43)**, **Late Industrialization & IRI (64)**, **Economic Miracle (90)**,
   **Third Italy (91)**, **The Paradox (102)**; plus enrichment bullets on
   31, 33, 46, 48, 59, 60, 97 (Arab Sicily, Legnano/Peace of Constance,
   maritime republics, Columbus/Atlantic shift, refeudalization, 1630 plague,
   Napoleon reforms, 1887 tariff, emigration, fourth capitalism).
3. **Slide 107 presentation topics** rebuilt from `Presentation Topics --
   Italy.xlsx`: 6 debate topics (Topic / Team A / Team B) + a single
   full-width **"Company Option" block with no A/B split** (company pills;
   #6 a dashed "TBD (Wed 11am)" placeholder). Back button → slide 8.
4. **Figures.** Added 11 CC-licensed illustrative figures to text-only High +
   Medium slides (16, 22, 24, 27, 33, 34, 40, 42, 47, 89 — slide 31 excluded
   by request, slide 20 dropped as too text-dense). Each has a sourced italic
   caption, rounded corners, soft shadow. Sourced via a research subagent from
   Wikimedia/PD/CC. **Slide 47's plague image is Micco Spadaro's Naples 1656
   painting** (no clean Milan-1630 image existed) — captioned honestly.
5. **Speaker notes** on every content slide: teleprompter-style narration for
   a general reader (first person, read-aloud, extra background), time-neutral,
   no dash-asides, each pointing at the posted picture. ~1 anecdote in 3, set
   off as `ANECDOTE — …` with a verified `Source:` link. Logistics (3–9),
   agenda/dividers, and the 7 poll slides get **no** notes (poll notes are
   protected); backups get one brief line. Slide 17's original Celts link kept.
6. **Animations** (Fade, ~0.5 s, on click): 85 slides built. Content slides
   build one top-level bullet per click (first bullet shows with the slide;
   sub-bullets ride with their parent). **Side figures fade in on the click of
   the bullet they illustrate** (default = first built bullet; overrides:
   16 → "Advanced art", 90 → "Icons/Fiat 500", 47 → "1630 plague"). Galleries
   reveal one picture per click; two-column 102 = left card → right card →
   takeaway; featured-research = paper then authors; takeaway bars last.
   Logistics 3–9 also animated (per-main-bullet). Skipped: title (1–2),
   agenda/dividers, poll slides, backups (105–135).
7. **Shade grouping.** Slides 7 and 107 had the table's shadow on a separate
   backing rectangle; grouped each backing + table into one `<p:grpSp>` so the
   shade travels/animates with the figure, then re-ran animations.
8. **Emails** drafted for students (not in the deck): IBR Custom Podcast
   Parts I–IV announcements (link placeholders; Nico pastes NotebookLM links).

### Key decisions / conventions
- **Canonical deck edited IN PLACE via OOXML surgery** (zip + lxml); **never
  round-trip through python-pptx** — it strips poll `tags` rels and NULL-target
  internal links. New slides are built in a temp deck then spliced.
- **Display→part mapping** via `ppt/presentation.xml` `<p:sldIdLst>` → rels
  (never assume `slideN.xml` = display N).
- **Render/verify via PowerPoint COM** (no LibreOffice); open in PowerPoint as
  the real integrity check after any structural edit.
- `_manifest.json` keys the 7 poll GUIDs + internal links to build functions so
  splices are insertion-robust.
- Figures: pictures carry their own shadow; tables/charts put the shade on a
  backing rect and **group it with the figure** (now a CLAUDE.md rule).
- Backups: keep only the two newest + the original; **gitignored** (see root
  `.gitignore`, `*_backup_*.pptx`).

### Build scripts (source of truth, in `Italy IBR/`)
`_build_Italy_Class1.py` (STALE scaffold; imports Module 3 template helpers),
`_splice_polls.py`, `_wire_links.py`, `_manifest.json`, `_resize_bullets.py`
(28/24 or 24/22 fit, measured with PIL), `_audit_fix.py`, `_build_new_slides.py`
/ `_add_slides.py` / `_slide_polish.py` / `_enrich_more.py`, `_build_more_slides.py`
/ `_add_more.py` (the 5 new slides), `_topics107.py`, `_add_figures.py`
(+ `_notes_data.py` / `_add_notes.py` for speaker notes), `_animate.py`
(animation engine; `FIG_GROUP` maps figure→bullet), `_group_shades.py`.
Rerun animations with `python _animate.py all apply`.

### Pending / open items
- **Commit + push the local changes:** shade grouping (deck), `_group_shades.py`,
  the consolidated `Teaching/CLAUDE.md` (+ deleted `405 .../CLAUDE.md`), and this
  Session-Notes restructure. (User to confirm the commit.)
- **Podcast links** — Nico pastes the four NotebookLM links into the drafted
  emails; verify the Part II/III/IV content summaries match his actual splits.
- **Slide 20 (Roman Republic)** has no figure (too text-dense) — offer to split
  into two slides if a figure is wanted there.
- **Slide 47** plague image is Naples 1656, not Milan 1630 — swap if preferred.
- **Slides 49 / 100 / 101** reveal their image clusters on one click (they have
  ≥2 bullets); switch to one-image-per-click if desired.

### Useful commands
```powershell
# Render a slide to PNG + integrity-check open (kill stale POWERPNT first).
Get-Process POWERPNT -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
# ... COM: $pp.Presentations.Open(deck,$true,$false,$false); $slide.Export(png,"PNG",w,h)

# Re-run all animations after any structural change (grouping, splices):
python _animate.py all apply        # or: python _animate.py 16 90 apply  (specific slides)

# Re-apply speaker notes (edits notesSlides in place; poll notes untouched):
python _add_notes.py apply
```
