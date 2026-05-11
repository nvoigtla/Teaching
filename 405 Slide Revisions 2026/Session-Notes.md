# 405 Slide Revisions 2026 – Session Notes

## 2026-05-11 – Module 3 structural review and template design

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
