# Huawei Style Path

## Use this path

Use this path when the user wants Huawei-style slides or when no other visual style is specified.

## Core rules

- Use `16:9`.
- Use the bundled `assets/huawei_ppt_master.pptx`.
- Use Microsoft YaHei for all text.
- Keep the slide background white.
- Use black for ordinary text.
- Use restrained red for structure and emphasis.
- Use red and bold for the most important title phrases and key facts.

## Typography

- Make text as large as possible while staying inside the box.
- Increase font size until the control nearly fills the available area.
- Shrink only when needed to avoid clipping.
- Prefer fewer lines and tighter copy over smaller type.
- Apply this to body text, labels, and section headers too, not just the title.

## Content structure

- Use a strong headline.
- Use one judgment band near the top.
- Use compact cards, matrices, flows, or timelines for the body.
- End with a concise conclusion bar or flywheel strip.

## Workflow

1. Build the content deck with `pptxgenjs`.
2. Keep it free of template-specific footer or logo elements.
3. Apply `scripts/apply_huawei_template.py`.
4. Run `python -m markitdown`.
5. Render to images and inspect the slide visually.

## Visual emphasis

- Use red and bold for 1 to 3 highest-priority title phrases or numbers.
- Do not overuse red; the slide should still read as Huawei-style, not a color-heavy poster.

## Shared use

- Copy the entire `pptx` folder to move this workflow to another agent.
- Keep the template asset and scripts inside the folder so the skill is self-contained.
