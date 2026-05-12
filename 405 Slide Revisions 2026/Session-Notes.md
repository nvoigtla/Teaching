# 405 Slide Revisions 2026 – Session Notes

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
