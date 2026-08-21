# 405 Slide Revisions 2026 — project CLAUDE.md

Project-specific conventions for the module-by-module rebuild of my
Management 405 decks. The universal and Teaching CLAUDE.md layers
apply in full; this file only adds what is specific to this project.

## Comparison-Deck Adoption Protocol
Each module rebuild starts from MY old deck and consults a colleague's
parallel deck for innovations (CT's deck for Module 2, MW's for
Module 1). Rules:
- Propose imports item by item in the outline file; I approve or
  reject each. Never sweep in colleague content on your own.
- Tag adoptions in the outline as `[NEW – MW]` / `[NEW – CT]` with the
  source slide number.
- My existing application / mini-case slides keep their place and
  wording; colleague material only ADDS (or replaces after my
  explicit approval).
- Colleague polls run on THEIR PollEverywhere accounts — never splice
  those. When an adopted example changes what my own poll should ask,
  flag that I need to reword the activity in my PollEv account (the
  embed URL stays valid; the static screenshot keeps the old wording
  until I re-capture it).

## File Naming (this project)
- Canonical deck: `Module X - Revised.pptx` (backups per the Teaching
  rules — two newest only).
- Example-candidates review deck: `Module X - Example Candidates.pptx`
  with build script `_build_MX_candidates.py`, which imports the
  module build script's helper layer.
- Research dumps (build inputs, keep): `_source_inventory.md`
  (per-deck text / notes / media), `_runfmt_dump.md` (run-level
  emphasis), `_assets_manifest.md` (extracted images + positions).
- Web-fetched photos: `_source_images/web_*.jpg`, fetched and reviewed
  via `_fetch_web_images.py`.

## Pipeline (per module — the Module 7 rerunnable pattern)
`_build_ModuleX.py` → `_splice_media.py` (polls / live content,
verbatim with notes + tags) → `_group_pass.py` (box+text, shade+frame,
pic+caption groups) → `_animate.py all apply` (fade builds per
per-slide plans). Verify with the COM click-count check
(`_verify_anim.ps1`) and at least one full-screen slideshow probe
before handing a deck over. Hand-edits are surfaced with a
member-level geometry diff (`_diff_slides.py`) against a side-path
build and ported into the build script with a dated comment.
