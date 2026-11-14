# PPTX Skill (huawei-pptx)

Codex skill for creating, editing, templating, and QAing PowerPoint presentations.

This package includes:

- `SKILL.md`: trigger rules and workflow overview
- `pptxgenjs.md`: create-from-scratch guidance using PptxGenJS
- `editing.md`: template-based editing workflow
- `assets/huawei_ppt_master.pptx`: bundled Huawei-style master (slim shell: one master chain, ~36KB)
- `assets/huawei_layout_catalog.json`: 58 real Huawei slides distilled into layout fingerprints (bands/columns/density) for layout selection
- `scripts/`: helper scripts for templating (`apply_huawei_template.py`), template pruning (`prune_template.py`), packing, validation, thumbnails, and preview export
- `references/`: style guidance, plus `huawei_layouts.md` — six layout archetypes and hard density targets distilled from real decks
- `requirements.txt`: Python dependencies (`defusedxml`, `markitdown[pptx]`, `Pillow`)

Default behavior creates dense Huawei-style executive slides when no other brand or template is specified. The shell script migrates images, charts, hyperlinks, and slide jumps from the content deck into the final file automatically.
