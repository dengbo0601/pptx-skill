---
name: huawei-pptx
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes creating, reading, editing, modifying, combining, splitting, templating, and QA for PowerPoint presentations. Trigger whenever the user mentions \"deck,\" \"slides,\" \"presentation,\" or references a .pptx filename. Default to Huawei-style output when no other brand or visual style is specified, using the built-in Huawei path for dense executive slides."
license: Proprietary. LICENSE.txt has complete terms
---

# PPTX Skill

## Local Default

When creating a new presentation and the user does not specify another brand, template, or visual style, default to Huawei-style output. This `pptx` skill now contains the Huawei path directly: template asset, shelling script, typography rules, and QA workflow are all inside this skill folder.

## Quick Reference

| Task | Guide |
|------|-------|
| Read/analyze content | `python -m markitdown presentation.pptx` |
| Edit or create from template | Read [editing.md](editing.md) |
| Create from scratch | Read [pptxgenjs.md](pptxgenjs.md) |
| Huawei-style default path | Read [references/huawei_style.md](references/huawei_style.md) |

---

## Reading Content

```bash
# Text extraction
python -m markitdown presentation.pptx

# Visual overview
python scripts/thumbnail.py presentation.pptx

# Raw XML
python scripts/office/unpack.py presentation.pptx unpacked/
```

---

## Editing Workflow

**Read [editing.md](editing.md) for full details.**

1. Analyze template with `thumbnail.py`
2. Unpack → manipulate slides → edit content → clean → pack

---

## Creating from Scratch

**Read [pptxgenjs.md](pptxgenjs.md) for full details.**

Use when no template or reference presentation is available.

For Huawei-style output, generate the editable content deck first, then apply the bundled Huawei master shell. See [references/huawei_style.md](references/huawei_style.md).

---

## Design Ideas

**Don't create boring slides.** Plain bullets on a white background won't impress anyone. Consider ideas from this list for each slide.

### Before Starting

- **Pick a bold, content-informed color palette**: The palette should feel designed for THIS topic. If swapping your colors into a completely different presentation would still "work," you haven't made specific enough choices.
- **Dominance over equality**: One color should dominate (60-70% visual weight), with 1-2 supporting tones and one sharp accent. Never give all colors equal weight.
- **Dark/light contrast**: Dark backgrounds for title + conclusion slides, light for content ("sandwich" structure). Or commit to dark throughout for a premium feel.
- **Commit to a visual motif**: Pick ONE distinctive element and repeat it — rounded image frames, icons in colored circles, thick single-side borders. Carry it across every slide.

### Color Palettes

Choose colors that match your topic — don't default to generic blue. Use these palettes as inspiration:

| Theme | Primary | Secondary | Accent |
|-------|---------|-----------|--------|
| **Midnight Executive** | `1E2761` (navy) | `CADCFC` (ice blue) | `FFFFFF` (white) |
| **Forest & Moss** | `2C5F2D` (forest) | `97BC62` (moss) | `F5F5F5` (cream) |
| **Coral Energy** | `F96167` (coral) | `F9E795` (gold) | `2F3C7E` (navy) |
| **Warm Terracotta** | `B85042` (terracotta) | `E7E8D1` (sand) | `A7BEAE` (sage) |
| **Ocean Gradient** | `065A82` (deep blue) | `1C7293` (teal) | `21295C` (midnight) |
| **Charcoal Minimal** | `36454F` (charcoal) | `F2F2F2` (off-white) | `212121` (black) |
| **Teal Trust** | `028090` (teal) | `00A896` (seafoam) | `02C39A` (mint) |
| **Berry & Cream** | `6D2E46` (berry) | `A26769` (dusty rose) | `ECE2D0` (cream) |
| **Sage Calm** | `84B59F` (sage) | `69A297` (eucalyptus) | `50808E` (slate) |
| **Cherry Bold** | `990011` (cherry) | `FCF6F5` (off-white) | `2F3C7E` (navy) |

### Huawei Style

Use this style when the user asks for a Huawei look, references Huawei-branded slides, or does not specify another visual style:

- Format: `16:9`
- Bundled default template: `assets/huawei_ppt_master.pptx`
- If a user provides a different Huawei template, use the user-provided file instead of the bundled asset.
- Background: white
- Primary text: black
- Structural emphasis: restrained dark red
- Highest-priority emphasis: restrained red, bold, and sparse
- Red tones: keep them restrained but slightly brighter than deep maroon, never bright or neon
- Font: Microsoft YaHei for all slide text
- Information density: high by default
- Typography: use the largest font size that fits inside each control without clipping; this is mandatory for body text, labels, table cells, flow nodes, captions, legends, callout bands, and conclusion bars, not just titles

Default generation workflow:

1. Generate a content-only slide deck with `pptxgenjs`.
   - The content deck owns layout, text, shapes, KPI cards, matrices, flows, charts, and other editable content.
   - Name it like `*_content_pptxgenjs.pptx`.
   - Keep it free of template-specific footer/logo elements.
2. Place the generated slide content into the Huawei template shell.
   - Use the bundled `assets/huawei_ppt_master.pptx` unless the user provides another template.
   - Run `python <skill_dir>/scripts/apply_huawei_template.py *_content_pptxgenjs.pptx final_output.pptx`.
   - The template shelling script must preserve every generated slide; multi-slide decks should remain multi-slide after templating.
   - The script migrates images, charts, embedded workbooks, hyperlinks, and slide jumps into the shell automatically; content decks may safely use all of these.
   - Preserve the template's master, footer, page number, confidentiality mark, and logo where appropriate.
   - Output the final presentation as the user-facing file.

Practical guidance:

- Use red and bold for the top 1-3 most important title phrases, numbers, or decision-critical facts.
- Use dark red for structural emphasis, supporting headers, borders, and secondary callouts.
- Keep the rest of the slide restrained: white space, black body text, and simple alignment.
- Avoid decorative gradients, extra bright accent colors, and noisy backgrounds.
- If a slide uses red, make sure it is consistent across the deck and visually subdued.
- Make fonts as large as possible within their text boxes. Start larger than seems necessary, then shrink only enough to avoid clipping after rendered QA.
- Prefer shortening wording, widening the control, increasing row height, or simplifying the layout before reducing font size.
- Do not leave obvious unused vertical space inside a text box while the text is small; tiny body text is a QA defect even when it does not overflow.
- In dense executive layouts, table cells, matrix rows, flow nodes, labels, legends, and conclusion bars must remain readable at full-slide view.
- Do not add a top red rule unless it exists in the requested template or the user asks for one.
- Do not add production metadata to slide content: no "Huawei Style / One Page", no source filename, no "user-provided text", no generation method, no draft labels unless requested.
- When generating a Huawei-style deck, start from the default template shell above unless the user explicitly provides another template.
- Prefer the two-stage `pptxgenjs` content deck + Huawei template shell workflow for stable generation.
- Default to dense executive-report pages rather than sparse marketing-style slides.
- Choose the slide structure from the source material: use matrices for comparisons, timelines for sequences, process flows for methods, callout bands for conclusions, and compact cards for parallel points.
- Preserve substantive distinctions from the source material, such as causes, evidence, constraints, risks, tradeoffs, decisions, milestones, and recommendations; compress wording before dropping important meaning.
- Use compact but large-readable text blocks, tight grids, and restrained dividers to fit more content without turning the slide into plain bullet lists.

### For Each Slide

**Every slide needs a visual element** — image, chart, icon, or shape. Text-only slides are forgettable.

**Layout options:**
- Two-column (text left, illustration on right)
- Icon + text rows (icon in colored circle, bold header, description below)
- 2x2 or 2x3 grid (image on one side, grid of content blocks on other)
- Half-bleed image (full left or right side) with content overlay

**Data display:**
- Large stat callouts (big numbers 60-72pt with small labels below)
- Comparison columns (before/after, pros/cons, side-by-side options)
- Timeline or process flow (numbered steps, arrows)

**Visual polish:**
- Icons in small colored circles next to section headers
- Italic accent text for key stats or taglines

### Typography

**Choose an interesting font pairing** — don't default to Arial. Pick a header font with personality and pair it with a clean body font.

| Header Font | Body Font |
|-------------|-----------|
| Georgia | Calibri |
| Arial Black | Arial |
| Calibri | Calibri Light |
| Cambria | Calibri |
| Trebuchet MS | Calibri |
| Impact | Arial |
| Palatino | Garamond |
| Consolas | Calibri |
| Microsoft YaHei | Microsoft YaHei |

| Element | Size |
|---------|------|
| Slide title | 36-44pt bold |
| Section header | 20-24pt bold |
| Body text | Use the largest size that fits; normally 15-20pt on sparse slides and 9-13pt only for dense tables/matrices |
| Captions | Use the largest size that fits; normally 11-13pt, smaller only when the caption is secondary and still readable |

For generated slides, the size table is a starting point, not a cap. If a text box has spare space, increase the font size or reduce the box height until the text fills the control cleanly.

### Spacing

- 0.5" minimum margins
- 0.3-0.5" between content blocks
- Leave breathing room—don't fill every inch

### Avoid (Common Mistakes)

- **Don't repeat the same layout** — vary columns, cards, and callouts across slides
- **Don't center body text** — left-align paragraphs and lists; center only titles
- **Don't skimp on size contrast** — titles need 36pt+ to stand out from 14-16pt body
- **Don't default to blue** — pick colors that reflect the specific topic
- **Don't mix spacing randomly** — choose 0.3" or 0.5" gaps and use consistently
- **Don't style one slide and leave the rest plain** — commit fully or keep it simple throughout
- **Don't create text-only slides** — add images, icons, charts, or visual elements; avoid plain title + bullets
- **Don't forget text box padding** — when aligning lines or shapes with text edges, set `margin: 0` on the text box or offset the shape to account for padding
- **Don't use low-contrast elements** — icons AND text need strong contrast against the background; avoid light text on light backgrounds or dark text on dark backgrounds
- **NEVER use accent lines under titles** — these are a hallmark of AI-generated slides; use whitespace or background color instead

---

## QA (Required)

**Assume there are problems. Your job is to find them.**

Your first render is almost never correct. Approach QA as a bug hunt, not a confirmation step. If you found zero issues on first inspection, you weren't looking hard enough.

### Content QA

```bash
python -m markitdown output.pptx
```

Check for missing content, typos, wrong order.

**When using templates, check for leftover placeholder text:**

```bash
python -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum|this.*(page|slide).*layout"
```

If grep returns results, fix them before declaring success.

### Visual QA

**⚠️ USE SUBAGENTS** — even for 2-3 slides. You've been staring at the code and will see what you expect, not what's there. Subagents have fresh eyes.

Convert slides to images (see [Converting to Images](#converting-to-images)), then use this prompt:

```
Visually inspect these slides. Assume there are issues — find them.

Look for:
- Overlapping elements (text through shapes, lines through words, stacked elements)
- Text overflow or cut off at edges/box boundaries
- Decorative lines positioned for single-line text but title wrapped to two lines
- Source citations or footers colliding with content above
- Elements too close (< 0.3" gaps) or cards/sections nearly touching
- Uneven gaps (large empty area in one place, cramped in another)
- Insufficient margin from slide edges (< 0.5")
- Columns or similar elements not aligned consistently
- Low-contrast text (e.g., light gray text on cream-colored background)
- Low-contrast icons (e.g., dark icons on dark backgrounds without a contrasting circle)
- Text boxes too narrow causing excessive wrapping
- Leftover placeholder content

For each slide, list issues or areas of concern, even if minor.

Read and analyze these images:
1. /path/to/slide-01.jpg (Expected: [brief description])
2. /path/to/slide-02.jpg (Expected: [brief description])

Report ALL issues found, including minor ones.
```

### Verification Loop

1. Generate slides → Convert to images → Inspect
2. **List issues found** (if none found, look again more critically)
3. Fix issues
4. **Re-verify affected slides** — one fix often creates another problem
5. Repeat until a full pass reveals no new issues

**Do not declare success until you've completed at least one fix-and-verify cycle.**

---

## Converting to Images

Convert presentations to individual slide images for visual inspection.

**Windows with PowerPoint installed** (highest fidelity — renders exactly like the end user will see it):

```powershell
powershell -File scripts/export_preview_windows.ps1 -PptxPath output.pptx -OutDir slides/
```

Produces `Slide1.JPG`, `Slide2.JPG`, ... at 1920x1080. Requires PowerPoint (COM automation); run it only when no other PowerPoint window has unsaved work, since the script quits the PowerPoint instance it creates.

**LibreOffice + Poppler** (cross-platform default):

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

This creates `slide-01.jpg`, `slide-02.jpg`, etc.

To re-render specific slides after fixes:

```bash
pdftoppm -jpeg -r 150 -f N -l N output.pdf slide-fixed
```

---

## Dependencies

- `pip install -r requirements.txt` — installs `defusedxml` (XML parsing in all scripts), `markitdown[pptx]` (text extraction), and `Pillow` (thumbnail grids)
- `npm install pptxgenjs` — creating from scratch (local install preferred over `-g`; run node scripts from the directory that has `node_modules`)
- LibreOffice (`soffice`) - PDF conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- Poppler (`pdftoppm`) - PDF to images (skip on Windows; use `scripts/export_preview_windows.ps1` instead)
