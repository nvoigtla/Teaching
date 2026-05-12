# 405 Slide Revisions 2026 – Session Notes

## 2026-05-12 – Clean deck rebuild on the "405 Slides Layout" (single layout)

**One-line summary.** Started a fresh `Module 3_clean.pptx` built from
scratch on the 6-type template system, with the PowerPoint Layout dropdown
stripped down to a single `405 Slides Layout`. Completed 42 of 78 slides
in three batches: front matter (1-6), §1.1 Short Run (7-22), §1.1b Wage
Searchers + §1.2 Long Run + Part 2 section divider (23-42). Added native
OMML (Cambria Math) equation rendering for `MP_K / p_K = MP_L / w` style
formulas, and a distinct "external-document reference" visual for
Teaching Note callouts (cream + dashed navy border + page icon + gold
`SEE TEACHING NOTE →` label).

### Final deck state

- [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) – **42 slides**, single layout (`405 Slides Layout`), audit clean (0 broken rels, 0 missing parts, 0 non-integer EMUs).
- Front matter (1-6): title cover, Zoom logistics, M2 recap, course-roadmap diagram, "every executive decision is a production-and-cost decision" diagram, hierarchical outline.
- §1.1 Short Run (7-22): section divider, production function with Karl Marx book, short-vs-long-run pair, rebuilt 9×9 production-function table, output curve, MPL with "plot the slope" callout, Black Death with half-page parchment textbox + chart, Tesla hiring setup, MRPL concept (with Teaching Note callout), MRPL detail, MRPL example, **poll slide using the source PollEv screenshot**, solution, MRPL > wage rule, optimum L\* (with "Revenue per car net of material cost" callout).
- §1.1b Wage Searchers (23-29): caution slide, big-employers-bid-wages-up, salary comparison chart, AI-researcher poaching example with Hassabis photo, poll, solution = **$8M**, UC wage search tool box.
- §1.2 Long Run (30-41): section divider (Part 1.2), Rivian Georgia plant context, optimal-input-mix concept with OMML formula, bang-for-the-buck rule (54pt OMML stacked-fraction headline), 5-step recipe with inline OMML inequalities, Rivian Georgia example, optimality check (production function), analysis with OMML inequality callout, poll, solution table, comparative statics (robot tax | union wages), grocery-shopping intuition reinforcer.
- Part 2 section divider (42): full Module 3 agenda with Part 2 now navy, Part 1 faded.

### Files added or modified this session

| File | Status | Notes |
|---|---|---|
| [Module 3/_build_clean_deck.py](Module 3/_build_clean_deck.py) | **New (main script)** | ~2000 lines. Builds the entire `Module 3_clean.pptx` from scratch. Imports primitives from `_build_template_samples.py`. Has builders for slides 1-42, the layout-stripping post-process, OMML helpers, and the new takeaway/Teaching Note/discussion-break/callout helpers. |
| [Module 3/Module 3_clean.pptx](Module 3/Module 3_clean.pptx) | **New (output)** | 42-slide draft. |
| Module 3/_source_images/ | New | Pictures extracted from `Module 3_NEW_draft.pptx` (slides 8-22 and 23-42), reused for image embedding in the clean deck. ~5 MB. |
| [Module 3/_zoom_logo.png](Module 3/_zoom_logo.png) | New | Extracted from source slide 2; used by `slide_2()`. |
| [Module 3/_batch2_dump.txt](Module 3/_batch2_dump.txt), [Module 3/_batch3_dump.txt](Module 3/_batch3_dump.txt) | New | Scratch reference dumps of source slides 7-22 and 23-42 with all shape positions, used as input for writing the per-slide builders. Kept for repeatability. |

### Decisions made this session

1. **Approach: per-slide builders, NOT wholesale XML surgery.** First considered lifting source slide XML wholesale and stripping/replacing chrome (the `_rebuild_via_zip.py` pattern), but landed on per-slide builders that explicitly add each source element (pictures, callouts, half-page textboxes, Teaching Notes). Reason: more readable, easier to iterate on per-slide visual issues, and we don't need to preserve animation timing XML (user can re-add in PowerPoint).

2. **Single-layout deck.** Stripped 10 of the 11 default python-pptx layouts via `strip_unused_layouts()`. Kept only the Blank layout, renamed to `405 Slides Layout`. The PowerPoint Layout dropdown now shows exactly one entry.

3. **Font sizes bumped for 2-up handout legibility.** Bullet text 32 pt / sub 28 pt where it fits; 24-26 pt in diagram boxes; section tag 16 pt; footer 12 pt. Slide 6 outline (13 lines) had to stay at 24/20 to fit; flagged as a candidate for splitting if you ever need bigger.

4. **Title case for all chrome.** Removed the `.upper()` from the top-bar helper. Section tags now read "Module 3 · Part 1 · Production" not "MODULE 3 · PART 1 · PRODUCTION". Action titles also title-cased.

5. **Lift-and-faithful-preserve over rewrite.** Every source element the user flagged got added back: Karl Marx book (slide 8), "plot the slope" callout (13), half-page parchment textbox (14), both Tesla images (15), Teaching Note + Optimal Hiring major-concept box (16), MRPL formula box (17), Discussion break parallelogram (18), full PollEv screenshot (19, 27, 38), Revenue-per-car callout (22), Teaching Note (32), Recipe-for-Exams Teaching Note (34), comparative-statics two-column (40), grocery pictures (41).

6. **Teaching Note redesigned as external-doc reference.** Previously rendered as a solid navy bar with gold arrow. Now: cream parchment fill, dashed navy border, navy folded-corner page icon on the left, gold `SEE TEACHING NOTE →` label, italic navy title. Visually distinct from in-slide takeaway/major-concept callouts.

7. **OMML formulas for TeX-feel rendering.** Used Office Math Markup Language (the same engine Insert > Equation uses) for `MP_K / p_K = MP_L / w` on slides 32, 33, 34, 37, 39 (inequality variant), and MRPL formula on slide 17. Native stacked fractions, italic Cambria Math variables, proper subscripts. NOT a hand-styled Unicode hack.

8. **Tesla → Rivian image correction on slide 35.** Source slide 35 still showed a Tesla plant despite the 2026 Tesla→Rivian currency revision. Replaced with `_rivian.jpg` (Rivian R1T, CC BY-SA Wikimedia) and corrected the caption.

### Gotchas — READ BEFORE CONTINUING ON HOME COMPUTER

1. **Non-integer EMU values silently break the file in PowerPoint.** Any
   `<a:off>` or `<a:ext>` with a decimal value (e.g. `3855567.5`) causes
   PowerPoint to refuse to open the deck. Cause: `int / 2` returns a
   float; `Inches(0.7) / 2` returns a float. Defense: always wrap shape
   coordinates with `int()` at the helper boundary, or use `//` for
   integer division. All five shape-primitive helpers
   (`_add_filled_box`, `_add_outlined_box`, `_add_arrow_shape`,
   `_add_arrow`, `_add_half_textbox`) now cast inputs defensively.
   **Test:** `python -c "..."` scanning for `(?:cx|cy|x|y)="-?\d+\.\d+"`
   in slide XML – expect 0 hits.

2. **OMML in a textbox needs `<a14:m>` wrapper to render.** Without it
   PowerPoint shows empty boxes instead of the equation. Required
   structure:
   ```xml
   <a:p>
     <a14:m>
       <m:oMathPara><m:oMath>…</m:oMath></m:oMathPara>
     </a14:m>
     <a:endParaRPr/>
   </a:p>
   ```
   `xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"`
   must be declared on (or above) the `<a:p>` element.

3. **For italic math variables, DO NOT emit `<m:rPr><m:sty m:val="p"/></m:rPr>`.**
   `m:sty="p"` forces "plain" (upright) style which silently overrides
   the italic property in `<a:rPr i="1">`. `_omml_run()` (italic var) is
   stripped of m:rPr; only `_omml_text()` (operators, numbers, acronyms)
   includes `m:sty="p"`.

4. **Animations (click-to-reveal) are NOT preserved on rebuild.**
   python-pptx has no clean API for `<p:timing>`. All the "appear on
   click" build animations from the original deck are flat in the clean
   deck. They have to be re-added in PowerPoint (select shape →
   Animations tab → Appear/Fade on click). Specific slides where this
   matters: 13 ("plot the slope" callout), 16 (Teaching Note), 32
   (Teaching Note), and any other animated callouts.

5. **The Teaching Note style is reserved for external-document references
   only.** Don't mix it with in-slide major-concept boxes. Teaching Note:
   cream + dashed navy border + page icon + gold `SEE TEACHING NOTE →`
   label. In-slide takeaway/major-concept: solid navy or gold fill,
   bold text. The visual separation is intentional.

6. **Source deck has stale Tesla images on Rivian-themed slides.** The
   2026 currency-revision pass swapped Tesla text → Rivian text but
   missed some images. Slide 35 was the one found and fixed; **watch
   for similar issues on slides 53-57** (cost-function chain that was
   also Tesla→Rivian-renamed) when working through batch 4 / §2.1.

7. **Numbers on slide 28 ($8M) and slide 39 (illustrative MP values) need
   user verification.** Slide 28: speaker notes say $8M ($5M + 2×$1.5M);
   stale source body had $3M – I went with $8M. Slide 39: MP_K ≈ 4 cars,
   MP_L ≈ 0.1 cars are illustrative numbers that produce the right
   directional conclusion – swap in actual classroom values when ready.

8. **Slide 32 and 34 content was built from speaker notes.** Source had
   only a title placeholder + Teaching Note callout on slide 32, and
   only a title on slide 34. I generated the major-concept content and
   the 5-step recipe from the bang-for-the-buck logic. Sanity check
   against your teaching notes.

### Pending – next session(s)

- [ ] **Batch 4: §2.1 Cost Concepts (slides 43-62, 20 slides).** Sunk costs, Waterworld decision tree, Apple Car opportunity-cost example, cost dictionary, Ross Stores annual report, ChatGPT subscription tier MC ≠ AC, Burn60-equivalent calculations, MC in finance, Rivian Georgia weekly cost function, iPhone teardown poll, naïve vs. complex cost functions. Several slides have **tables and complex group shapes** (Waterworld decision tree especially) that will need bespoke per-slide builders.
- [ ] **Batch 5: §2.2 Scale & Scope (slides 63-74, 12 slides).** Long-run AC envelope, economies of scale technological reasons, Embraer ERJ-145 vs Boeing 787 (also AI-training scale slides per session 2), diseconomies, scope (Airbus A380 / A318), Amazon scale-or-scope, Shark Tank case + two PollEv slides.
- [ ] **Batch 6: Closing synthesis (slide 75, 1 slide).** Use Layout 6 (the closing synthesis template) – top half Module 3 recap, bottom half Module 4 preview.
- [ ] **Animation pass** in PowerPoint after the deck is structurally final – add Appear/Fade on click for "plot the slope", Teaching Notes, takeaway bars where appropriate.
- [ ] **Verify slide 28 ($8M)**, **slide 39 numbers**, **slide 34 recipe**, **slide 32 content** with classroom material.
- [ ] **Slide 6 outline** (13 lines @ 24/20 with tight spacing) – decide if it should be split into two slides to allow bigger type for handouts.
- [ ] **Re-add manual styling** to slide 19 / 27 / 38 poll slides if you want the cleaner Layout 5 design (A/B/C/D auto-numbered options + POLL pill) instead of the source PollEv screenshot picture.
- [ ] **Stale Tesla images** – sweep slides 53-57 for the same Tesla→Rivian issue I caught on slide 35.

### Commands worth remembering

- **Rebuild the clean deck from scratch:**
  ```powershell
  cd "h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3"
  Remove-Item "Module 3_clean.pptx" -ErrorAction SilentlyContinue
  $env:PYTHONIOENCODING = "utf-8"
  python _build_clean_deck.py
  ```
- **Audit the deck (broken rels, missing parts, declared-but-missing):**
  ```powershell
  python _audit_package.py "Module 3_clean.pptx"
  ```
- **Scan for non-integer EMU values (should always be 0):**
  ```powershell
  python -c "import zipfile, re; z=zipfile.ZipFile('Module 3_clean.pptx'); bad=0
  for n in z.namelist():
      if n.startswith('ppt/slides/') and n.endswith('.xml'):
          x = z.read(n).decode('utf-8')
          for m in re.finditer(r'(?:cx|cy|x|y)=\"(\-?\d+\.\d+)\"', x):
              print(f'{n}: bad {m.group(0)}'); bad += 1
  print(f'Non-integer EMU values: {bad}')"
  ```
- **Extract source images for the next batch:**
  ```powershell
  # Edit slide range at top of the extraction script, then:
  python -c "import zipfile, re; from pathlib import Path
  OUT = Path('_source_images'); OUT.mkdir(exist_ok=True)
  src = zipfile.ZipFile('Module 3_NEW_draft.pptx')
  for sno in range(43, 63):  # next batch range
      try: rels = src.read(f'ppt/slides/_rels/slide{sno}.xml.rels').decode('utf-8')
      except KeyError: continue
      for m in re.finditer(r'<Relationship Id=\"(rId\d+)\" Type=\"[^\"]*image\" Target=\"([^\"]+)\"', rels):
          rid, target = m.group(1), m.group(2)
          media = 'ppt/' + target[3:] if target.startswith('../') else target
          try: data = src.read(media)
          except KeyError: continue
          (OUT / f'slide{sno}_{rid}.{media.rsplit(chr(46),1)[-1]}').write_bytes(data)"
  ```
- **Dump source slide content for a batch (input for writing builders):**
  ```powershell
  # Inside python:
  from pptx import Presentation
  from pptx.enum.shapes import MSO_SHAPE_TYPE
  # ...iterate slides, dump shapes with positions, see _batch2_dump.txt / _batch3_dump.txt for format
  ```

### Useful context for resuming

- Working directory is the network share `H:\Claude Code\Teaching\` on this machine. On the home computer it will be a different path – grep for hard-coded paths if any builds fail (there shouldn't be – the script uses `Path(__file__).parent`).
- `python-pptx 1.0.2` on Python 3.14 confirmed working. Use `lxml` for namespace-aware XML manipulation.
- The git remote is `https://github.com/nvoigtla/Teaching.git`, branch `main`.
- The 405 Slides Template definition is in [Module 3/_build_template_samples.py](Module 3/_build_template_samples.py) (single source of truth for visual constants `NAVY`, `GOLD`, `FADED`, `RULE`, `GRAY`, `WHITE`, `MARGIN`, `RULE_W`, `GOLD_W`, `SLIDE_W`, `SLIDE_H` and the chrome helpers `_add_text`, `_add_rect`, `_draw_action_title`, `_add_bulleted_list`, `_set_bullet_char`, `_blank_slide`, `MODULE_AGENDA`, `FOOTER_TEXT`).
- The clean-deck script `_build_clean_deck.py` adds: `_draw_top_bar_tc` (title-cased top bar, replacing the all-caps default), `_draw_footer` (with bigger 12pt type), `_add_filled_box` / `_add_outlined_box` / `_add_arrow_shape` / `_add_arrow` (diagram primitives), `_add_takeaway_bar` / `_add_teaching_note` / `_add_discussion_break` / `_add_callout_box` / `_add_half_textbox` (callout helpers), `_omml_run` / `_omml_text` / `_omml_sub` / `_omml_frac` / `_add_math_equation` / `_formula_bang_for_buck` / `_formula_mp_ratio` (OMML formula helpers), `strip_unused_layouts` (post-process layout pruning).
- **Critical visual constants** in `_build_template_samples.py`: NAVY=`#0B2B4E`, GOLD=`#E09F3E`, RULE=`#C8CDD3` (light grey horizontal rule), GRAY=`#555B66` (body text grey), FADED=`#B0B5BC` (inactive agenda items), Calibri throughout, 0.7 cm side margins (MARGIN ≈ 0.276" ≈ 251999 EMU).

---

## 2026-05-11 (session 2) – Module 3 rebuild + 2026 currency revision

**One-line summary.** Locked the open outline decisions, cycled the
template through three rounds of visual feedback, rebuilt
`Module 3.pptx` end-to-end into a 78-slide `Module 3_NEW_draft.pptx`
via direct zip+lxml surgery (python-pptx round-trip corrupts this
deck — see gotcha below), then ran a 2026-currency revision pass
(Tesla → Rivian, three new AI examples, Meta Reality Labs + AI training
scale inserts) and replaced the three placeholder logos with real
photographs.

### Final deck state

- [Module 3/Module 3_NEW_draft.pptx](Module 3/Module 3_NEW_draft.pptx) – **78 slides, 24.5 MB**, package audit clean.
- Title slide (#1) – cover-style white background, "Production and Costs / Module 3 / Management 405 EMBA / Prof. Nico Voigtlaender · UCLA Anderson".
- All 76 original notesSlides + 2 new ones populated with MBA-friendly speaker notes (3-6 sentences each, ~14k chars total).
- All character bullets normalized to navy squares (▪) with level-based sizing (lvl 0: 150%, lvl 1: 110%, lvl 2: 85%, lvl 3+: 70%).
- Chrome page numbers updated to match new positions.

### Files worked on (this session)

| File | Status | Notes |
|---|---|---|
| [Module 3/Module 3.pptx](Module 3/Module 3.pptx) | **Untouched** | Original 81-slide deck, never modified. |
| [Module 3/Module 3_NEW_draft.pptx](Module 3/Module 3_NEW_draft.pptx) | New (final output) | 78-slide rebuild with 2026 examples + real photos. |
| [Module 3/Module 3 - outline.md](Module 3/Module 3 - outline.md) | Modified | "Decisions locked" block added; all 5 open questions answered (section dividers in 405 template style; Tesla→Apple Car for opportunity cost; slide #5 kept; closing synthesis added; speaker-notes plan confirmed). |
| [Module 3/405 Slides Template.pptx](Module 3/405 Slides Template.pptx) | Modified | Now a 6-layout reference deck (title cover, section header / Agenda, content bulleted, two-column, poll, closing synthesis). Three rounds of feedback applied. |
| [Module 3/_build_template_samples.py](Module 3/_build_template_samples.py) | Modified | All 6 template layouts plus native PPT bullet helpers (`_set_bullet_char`, `_set_bullet_autonum`, `_add_bulleted_list`). |
| [Module 3/_rebuild_via_zip.py](Module 3/_rebuild_via_zip.py) | New | Main rebuild: 10 cuts, 5 NEW slides via sidecar, reorder, REVISE titles, Apple Car content swap, chrome overlay. **Direct zip+lxml surgery, never round-trips through python-pptx.** |
| [Module 3/_normalize_bullets.py](Module 3/_normalize_bullets.py) | New | Replaces every `<a:buChar>` with ▪ in Calibri (slides + masters). |
| [Module 3/_apply_bullet_hierarchy.py](Module 3/_apply_bullet_hierarchy.py) | New | Sets navy color + level-based `buSzPct` on every bullet pPr. |
| [Module 3/_polish_v2.py](Module 3/_polish_v2.py) | New | Forces white background on slide 1; first Apple Car (Apple Park image) build; adds/replaces notes on all 76 slides. |
| [Module 3/_replace_title_slide.py](Module 3/_replace_title_slide.py) | New | Clean title-slide swap (delete-and-reinject via sidecar) – the working fix after the background hack didn't take. |
| [Module 3/_revise_v1.py](Module 3/_revise_v1.py) | New | 2026 currency revision: Tesla→Rivian text swap, designer→AI researcher, Burn60→ChatGPT, iPhone 11→17, Airbus→Alphabet, + 2 NEW slides (Meta Reality Labs, AI training scale). Updates page numbers and all notes after structural changes. |
| [Module 3/_add_images.py](Module 3/_add_images.py) | New | Rebuilds slides 47 + 69 with bullets-left/picture-right layout; overlays images on 26, 31, 51, 72. First image pass (used logos for 3 slides). |
| [Module 3/_replace_logos_with_photos.py](Module 3/_replace_logos_with_photos.py) | New | Second image pass: removed the Anthropic / ChatGPT / Alphabet logos and replaced with real photographs (Demis Hassabis, smartphone running ChatGPT, Googleplex). |
| [Module 3/_speaker_notes.py](Module 3/_speaker_notes.py) | New | 78 entries, the single source of truth for speaker notes – edit here, re-run `_polish_v2.py`'s `add_notes_to_slide` to re-inject. |
| [Module 3/_audit_package.py](Module 3/_audit_package.py) | New | Sanity-checks declared parts vs zip files, plus broken `.rels` targets. Useful after every edit. |
| [Module 3/_extract_outline.py](Module 3/_extract_outline.py) | Carry-over | Dumps slide titles/body/notes from any .pptx; produced `_outline_dump.txt`. |
| Module 3/_*.jpg, _*.png (7 files) | New | Image assets: `_apple_car_concept.jpg`, `_chatgpt_phone.jpg`, `_googleplex.jpg`, `_hassabis.jpg`, `_meta_quest.jpg`, `_nvidia_h100.png`, `_rivian.jpg`. All CC BY-SA / PD from Wikimedia Commons. |

### Decisions made

1. **All 5 open questions resolved** (see top of `Module 3 - outline.md`): 405 template style for dividers; Tesla → Apple Car on opportunity-cost slide; old slide #5 kept (you'll draft agenda copy yourself); closing synthesis added; speaker-notes plan confirmed (new/merged slides get fresh notes; revised-title slides keep existing notes unless retitle changes framing).

2. **Implementation strategy: Hybrid** – restructure (cuts/inserts/reorders) + restyle (chrome overlay) on the original deck. Three options were offered (skeleton-only, restructure-only, hybrid); you picked hybrid.

3. **Template visual system** – navy `#0B2B4E`, gold `#E09F3E`, light-gray rule `#C8CDD3`, body gray `#555B66`, faded grey `#B0B5BC`; Calibri throughout; 0.7 cm side margins. Native PPT bullets ▪ in navy with level-based sizing.

4. **2026 currency revision scope (you chose "aggressive")** – Tesla → Rivian on long-run + cost-function half (9 slides); 3 AI examples added (AI researcher poaching, ChatGPT subscription tiers, AI training costs); Waterworld kept alongside Meta Reality Labs parallel; both Embraer/Boeing and AI scale slides kept (didn't replace one with the other).

5. **Slide-1 title fix** – first tried adding `<p:bg>` white override (didn't render in PowerPoint); the working fix was to delete slide 1 entirely and inject a fresh template-built one via sidecar.

6. **Image policy** – swapped the three placeholder logos (Anthropic, ChatGPT, Alphabet) for real photos after you flagged that they didn't add visual content. Final image set: Rivian R1T, Meta Quest 3, NVIDIA H100, Demis Hassabis (Nobel 2024), smartphone with ChatGPT, Googleplex Mountain View, plus the existing Vanarama Apple Car render.

### The big gotcha: python-pptx round-trip corrupts this deck

The first rebuild via `_rebuild_module_3.py` (now deleted) produced a file PowerPoint repaired to empty. Bisection showed even a no-op round-trip (open, save) of `Module 3.pptx` via python-pptx breaks the deck — diffing the rels files revealed python-pptx silently **strips `Target="NULL"` image rels** from slides 10, 11, 16, 17, 25 (and a couple others). The slide XML keeps the `r:id` references, so PowerPoint sees dangling pointers and bails.

**Workaround used everywhere now:** all edits to `Module 3_NEW_draft.pptx` are done via direct zip + lxml surgery (`_rebuild_via_zip.py`, `_polish_v2.py`, `_revise_v1.py`, `_add_images.py`, `_replace_logos_with_photos.py`). python-pptx is used only on freshly-built clean sidecar decks (no NULL rels to strip), whose slide XML is then extracted and injected.

### Open / pending for next session

- [ ] **Review the 78-slide deck visually** in PowerPoint and flag any layout issues. Specific things to scrutinize:
  - Slides 26, 31, 51, 72 – picture overlays on existing-content slides; might overlap with original elements at certain zoom levels.
  - Slide 4 ("Agenda for the class") – content is still a stub; you said you'd draft this yourself.
  - The original page-number placeholders from the master may show TWO numbers on some slides (the master's plus the chrome overlay's). My chrome page number is the correct one; the master's may be stale.

- [ ] **Speaker notes for affected slides** – they're plausible defaults, but a few use placeholder numbers (e.g., slide 60 iPhone 17 AVC ≈ $580 is my estimate, not a real teardown). Update with actual figures if you have them.

- [ ] **Rebuild Module 4** – when ready, the same pipeline applies. The scripts here are reusable (just change the cuts/inserts/replacements).

### Commands / workflows worth remembering

- **Rebuild from scratch** (re-runs entire pipeline on `Module 3.pptx`):
  ```powershell
  cd "d:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3"
  python _rebuild_via_zip.py
  python _normalize_bullets.py
  python _apply_bullet_hierarchy.py
  python _polish_v2.py
  python _replace_title_slide.py
  python _revise_v1.py
  python _add_images.py
  python _replace_logos_with_photos.py
  ```
- **Set UTF-8 stdout** when scripts print special characters (▪, →, ·):
  ```powershell
  $env:PYTHONIOENCODING = "utf-8"
  ```
- **Audit deck consistency** after any edit:
  ```powershell
  python _audit_package.py "Module 3_NEW_draft.pptx"
  ```
- **Dump deck content** for review:
  ```powershell
  python _extract_outline.py     # writes _outline_dump.txt from Module 3.pptx
  ```
- **Edit speaker notes** – modify `_speaker_notes.py` then re-run the notes step inside `_polish_v2.py` (or call its `add_notes_to_slide` directly).
- **If PowerPoint locks the file during `.tmp` rename**: the temp file is written successfully; just close PowerPoint and `mv "*.pptx.tmp" "*.pptx"`.
- **When a deck refuses to open after a python-pptx save**: don't try to repair via python-pptx; use direct zip+lxml and preserve the original rels (especially NULL-target image rels).

### Useful context for resuming

- Working directory is now `D:\Claude Code\Teaching\` (different machine from the May 4 session which was on the `H:\` network share). The OneDrive corruption issue from session 1 is gone here.
- `python-pptx 1.0.2` is installed on this machine (Python 3.14).
- The git remote is `https://github.com/nvoigtla/Teaching.git`, branch `main`.
- Folder is now clean: 4 deliverable files, 7 image assets, 12 reusable scripts, 1 outline doc, 1 reference dump. ~50 MB total.

---

## 2026-05-11 (session 1) – Module 3 structural review and template design

**One-line summary.** Reviewed the existing 81-slide Module 3 deck, drafted a revised 73-slide outline, designed a custom MBA-style PowerPoint template ("405 Slides Template"), and committed everything to GitHub.

### Files worked on

| File | Status | Notes |
|---|---|---|
| [Module 3/Module 3.pptx](Module 3/Module 3.pptx) | Untouched (restored from backup) | Original May 4 version, 23 MB, 81 slides. Was corrupted twice during the session by OneDrive / network-share sync (file size jumped to 42 MB and zip central directory broke). Restored from backup both times. |
| [Module 3/Module 3_backup_2026-05-11.pptx](Module 3/Module 3_backup_2026-05-11.pptx) | New | Timestamped backup of the May 4 deck, made before any planned edits. Keep until the revised deck is final. |
| [Module 3/Module 3 - outline.md](Module 3/Module 3 - outline.md) | New | Revised 73-slide outline with explicit Part 1/Part 2 structure, section-header dividers replacing the seven redundant "OUTLINE OF Module 3" slides, takeaway-style titles for ~20 slides, and a closing synthesis slide. |
| [Module 3/405 Slides Template.pptx](Module 3/405 Slides Template.pptx) | New | Custom MBA-style template – navy top bar with white section tag, action-title format, two gold left-edge segments (under title, above footer) creating a parallel rhythm, all main elements within 0.7 cm side margins. Built from scratch with python-pptx, no third-party assets, no attribution required. |
| [Module 3/_build_template_samples.py](Module 3/_build_template_samples.py) | New | Python-pptx script that regenerates the 405 Slides Template. Useful for future tweaks (colors, spacing, accent length). |
| [Module 3/_extract_outline.py](Module 3/_extract_outline.py) | New | Helper script that dumps the slide titles, bodies, and speaker notes from any .pptx to a UTF-8 text file. |
| [Module 3/_outline_dump.txt](Module 3/_outline_dump.txt) | New | Full text dump of the original 81-slide deck. Reference for the rebuild. |

### Decisions made

1. **Top-level structure for Module 3** – two-part, four-section: Part 1 Production (§1.1 Short Run, §1.2 Long Run) and Part 2 Costs (§2.1 Cost concepts, §2.2 Scale & scope). Cuts the seven redundant outline slides; replaces each with a true section-header divider.
2. **Front matter** – condensed from 9 slides to 5 (consolidated logistics+announcements, dropped the orphan "agenda for the class" stub and the duplicate outline, beefed up the "Big Picture" slide).
3. **Slide-title convention** – takeaway-first (per Teaching CLAUDE.md), e.g., "Cost types" → "Sunk costs should never drive decisions"; "Wage searchers" → "Big employers bid their own wages up".
4. **Template direction** – consulting/boardroom style as the base (closest to McKinsey/BCG aesthetic that EMBAs already read at work), with gold accents from an academic-modern style for visual rhythm.
5. **Template specifics** – navy `#0B2B4E`, gold `#E09F3E`, light-gray rule `#C8CDD3`, body gray `#555B66`. 0.7 cm side margins for all main horizontal elements. Two gold left-segments of 2.2" each, sitting on top of the title divider and the footer rule. Title sits close to the top bar (y=0.55") with the divider at y=1.25".
6. **Truly-free template sources** – verified none of the template-design work uses anything with hidden attribution or watermarks. Microsoft Create, SlidesCarnival, and certain GitHub repos identified as the only no-strings sources.

### Open / pending for next session

- [ ] **Build the full slide-master family** in the 405 Slides Template style: title slide, section-header layout, content layout (this), two-column variant, poll-question layout, closing synthesis layout.
- [ ] **Rebuild Module 3.pptx** from the 73-slide outline in `Module 3 - outline.md` using these layouts. Wait for explicit approval on the outline before doing this.
- [ ] **Answer the 5 open questions** at the bottom of `Module 3 - outline.md`:
  1. Section-divider visual style preference.
  2. Whether to swap a Tesla example (currently old slide #53 / new slide #46) for a non-Tesla company for variety.
  3. Confirm that the orphan slide #5 ("agenda for the class") is a leftover and can be cut.
  4. Whether to add a closing synthesis slide (currently in the plan as new slide #75).
  5. Speaker-notes treatment for retitled vs. merged vs. new slides.
- [ ] **Investigate the OneDrive / network-share corruption issue.** `Module 3.pptx` was corrupted twice today (file size jumped from 23 MB to 42 MB, ZIP end-of-central-directory bytes became UTF-8 replacement characters). Suspect OneDrive Files-on-Demand or network-share sync on `\\FSC1\Faculty_Personal\`. Workarounds used: restore from version history, work from a local copy, keep the timestamped backup.

### Commands / workflows worth remembering

- **Extract slide content for review** without opening PowerPoint:
  ```powershell
  cd "h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3"
  python _extract_outline.py
  # writes _outline_dump.txt (UTF-8 safe)
  ```
- **Rebuild the template after tweaks**:
  ```powershell
  cd "h:/Claude Code/Teaching/405 Slide Revisions 2026/Module 3"
  python _build_template_samples.py
  # writes "405 Slides Template.pptx"
  ```
- **If a .pptx becomes corrupted again** – check zip header/footer:
  ```bash
  head -c 8 "Module 3.pptx" | xxd   # should start with "PK\x03\x04"
  tail -c 32 "Module 3.pptx" | xxd  # should end cleanly with "PK\x05\x06" then small fields
  ```
  If the footer is followed by `EF BF BD` (UTF-8 replacement bytes), restore from backup or OneDrive version history.

### Useful context for resuming

- The Teaching folder lives on a network share (`\\FSC1\Faculty_Personal\nvoigtla\Claude Code\Teaching`), mapped to `H:\`. Git operations work but binary file sync is unreliable for large .pptx files. Consider working from a local copy and pushing to git rather than relying on the share to be the source of truth.
- python-pptx 1.0.2 is installed and works fine for reading and building decks.
- The git remote is `https://github.com/nvoigtla/Teaching.git` (branch `main`).
