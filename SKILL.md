---
name: pptx-studio
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file; editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Also use when: (1) the user wants to apply or swap a template style — phrases like 'apply template', 'use this template', 'fit into template', 'change my PPT style', 'switch template', 'replace template', 'change the template', 'switch the template', 'use a different template', '套用模板', '换模板', '更换模板', '换个模板', '替换模板', '把PPT套入模板', '套入指定模板', '模板转换', '换一套模板', '更改模板', '切换模板'; (2) the user wants to beautify or redesign a PPT — phrases like 'beautify', 'redesign', 'make it look better', 'improve design', 'modernize slides', '美化PPT', '优化设计', '让PPT更好看', '重新设计风格', '改造PPT外观', '美化幻灯片'; (3) the user wants to generate or add speaker notes — phrases like 'add speaker notes', 'generate notes', 'write notes for presenter', 'add talking points', 'create presentation notes', 'notes for each slide', '加备注', '生成演讲者备注', '写备注', '添加演讲提示', '为每页写备注', '自动生成备注', '演讲者视图备注'. Trigger whenever the user mentions 'deck', 'slides', 'presentation', or references a .pptx filename."
license: Proprietary. LICENSE.txt has complete terms
---

# PPTX Studio

Handle all .pptx tasks — create from scratch, edit existing files, apply templates, and beautify designs.

## Decision Flow

**Start here every time.** Before doing anything, figure out which mode to use:

```
User provides TWO .pptx files?
  → Mode 1: Template Apply (apply second file's style to first file's content)
  ⚠️  Warn first if source file likely has animations:
      "apply_template 会重建幻灯片结构，所有动画和切换效果将丢失。
       如果源文件有动画，请确认是否继续。"
      Then proceed only after user confirms.

User provides ONE .pptx file + says "beautify / redesign / make it look better"?
  → Mode 2: Style Beautify

User provides ONE .pptx file + says "edit / update / change content / add slides"?
  → If it's a single text fix (typo, title, number): patch_slide.py (fast path)
  → If it's structural (add/delete/reorder slides): Editing Workflow (unpack → edit XML → pack)

User provides ONE .pptx file + says "read / extract / summarize / what's in this"?
  → Reading Content (markitdown)

User provides ONE .pptx file + says "add speaker notes / generate notes / write notes for me / 加备注 / 生成演讲者备注 / 写备注"?
  → Speaker Notes Workflow

User provides NO file + wants a new presentation?
  → Creating from Scratch (pptxgenjs)

Still unclear?
  → Ask: "您想对PPT做什么？套用模板、美化风格、编辑内容，还是从头创建？"
```

**Quick decision table:**

| User says... | Mode |
|--------------|------|
| 给两个pptx，套/换/应用模板 | Mode 1: Template Apply |
| 美化、优化、让它更好看 | Mode 2: Style Beautify |
| 更换模板、切换模板、换个风格 | Mode 1（有模板文件）或 Mode 2（无模板文件） |
| 修改内容、调整文字、增减页 | patch_slide（单点文字）或 Editing Workflow（结构性修改） |
| 读取、提取、总结内容 | Reading Content |
| 加备注、写演讲提示、生成 Speaker Notes | Speaker Notes Workflow |
| 做一个新PPT | Creating from Scratch |

---

## Quick Reference

| Task | Guide |
|------|-------|
| Read/analyze content | `python -m markitdown presentation.pptx` |
| Edit existing PPT | Editing Workflow section below |
| Create from scratch | Creating from Scratch section below |
| Apply template to PPT | Mode 1 below |
| Beautify PPT style | Mode 2 below |
| Generate speaker notes | Speaker Notes Workflow below |
| QA check output file | `python scripts/qa_check.py output.pptx` |

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

Edit existing presentations by working directly with XML.

**Steps:**

1. **Analyze** — `python scripts/thumbnail.py input.pptx` to see all layouts visually
2. **Unpack** — `python scripts/office/unpack.py input.pptx unpacked/`
3. **Restructure** (if needed) — delete/reorder/add slides in `ppt/presentation.xml → <p:sldIdLst>` before editing content
   - Add slide: `python scripts/add_slide.py unpacked/ slide2.xml` (duplicate existing)
   - Delete slide: remove its `<p:sldId>` from `<p:sldIdLst>`, then run clean
4. **Edit content** — update text in each `unpacked/ppt/slides/slideN.xml`
   - ⚠️ Use **subagents** for multi-slide edits — slides are independent XML files, perfect for parallelism
   - Use the Edit tool, never sed or Python scripts
   - Bold headers: `b="1"` on `<a:rPr>`; never use unicode bullets `•`
5. **Clean** — `python scripts/clean.py unpacked/`
6. **Pack** — `python scripts/office/pack.py unpacked/ output.pptx --original input.pptx`

**Key pitfalls:**
- Longer text may overflow text boxes — always run visual QA after edits
- Template slots ≠ source items: if template has 4 items but you have 3, delete the entire 4th element group (not just the text)
- Multi-item content: use separate `<a:p>` elements, never concatenate into one string

> Full reference (XML snippets, smart quotes, formatting rules): [editing.md](editing.md)

---

## Creating from Scratch

Use when no template or reference file is available. Uses **PptxGenJS** (Node.js).

**Prerequisites — check environment first:**

```bash
node --version   # must be v14+
npm --version
```

If `node` is not found, install before proceeding:

```bash
# macOS (Homebrew)
brew install node

# Or download from https://nodejs.org/en/download
```

If Node.js is unavailable and cannot be installed, consider the Editing Workflow instead:
unpack a blank template PPTX, write slide XML directly, then pack.

**Setup:**
```bash
npm install -g pptxgenjs react react-dom react-icons sharp
```

**Basic structure:**
```javascript
const pptxgen = require("pptxgenjs");
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';  // 10" × 5.625"

let slide = pres.addSlide();
slide.background = { color: "1E2761" };
slide.addText("Title", { x: 1, y: 2, w: 8, h: 1.5, fontSize: 44, bold: true, color: "FFFFFF" });

pres.writeFile({ fileName: "output.pptx" });
```

**Critical rules (file corruption if violated):**
- ❌ NEVER use `#` prefix in hex colors: `"FF0000"` ✅ / `"#FF0000"` ❌
- ❌ NEVER use 8-char hex for opacity: use `opacity: 0.15` property instead
- ❌ NEVER reuse option objects across multiple `addShape`/`addText` calls (PptxGenJS mutates them in-place)
- ❌ NEVER use unicode bullets `•`: use `bullet: true` option

**Common elements:**
```javascript
// Multi-line text with bullets
slide.addText([
  { text: "Point 1", options: { bullet: true, breakLine: true } },
  { text: "Point 2", options: { bullet: true } }
], { x: 0.5, y: 1, w: 9, h: 3, fontSize: 18 });

// Accent shape (left border bar pattern)
slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625,
  fill: { color: "028090" } });

// Image
slide.addImage({ path: "image.png", x: 5, y: 1, w: 4, h: 3 });

// Shadow (always use fresh object per call)
const makeShadow = () => ({ type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.15 });
slide.addShape(pres.shapes.RECTANGLE, { fill: { color: "FFFFFF" }, shadow: makeShadow(), x:1, y:1, w:3, h:2 });
```

> Full reference (charts, tables, icons, shadows, slide masters): [pptxgenjs.md](pptxgenjs.md)

---

## Transform Presentations

Two powerful transformation modes:

| Mode | Description |
|------|-------------|
| **Template Apply** | Extract content from a source PPT and reflow it into a target template's layouts and visual identity |
| **Style Beautify** | Analyze PPT content and redesign the visual style — colors, fonts, layouts — without changing the content |

Transform scripts quick reference:

| Task | Script |
|------|--------|
| Extract content from PPT | `python scripts/extract_content.py source.pptx` |
| Analyze template layouts | `python scripts/thumbnail.py template.pptx` |
| Apply template to PPT | `python scripts/apply_template.py source.pptx template.pptx output.pptx` |
| Beautify PPT | `python scripts/beautify_ppt.py source.pptx output.pptx [--theme THEME]` |
| Generate speaker notes | `python scripts/generate_notes.py source.pptx output.pptx [--mode MODE]` |
| QA check output file | `python scripts/qa_check.py output.pptx` |
| Patch text in a PPT | `python scripts/patch_slide.py deck.pptx --find "X" --replace "Y"` |
| Unpack for editing | `python scripts/office/unpack.py file.pptx unpacked/` |
| Pack after editing | `python scripts/office/pack.py unpacked/ output.pptx --original source.pptx` |

---

## Mode 1: Template Apply

Apply an existing template's visual identity to your content.

### When to use

- User provides a source PPT (their content) and a template PPT (the desired look)
- User says: "apply this template", "make my PPT look like this", "use this style"

### Workflow

**Step 0 — Animation warning (always do this first)**

Run a quick check to see if the source file has animations:

```bash
python -c "
from pptx import Presentation
from pptx.oxml.ns import qn
prs = Presentation('source.pptx')
has_anim = any(
    slide._element.find('.//' + qn('p:timing')) is not None
    for slide in prs.slides
)
print('HAS ANIMATIONS:', has_anim)
"
```

If `HAS ANIMATIONS: True`, tell the user:

> ⚠️ **注意：源文件包含动画效果。** `apply_template` 会重建每页幻灯片的结构，所有动画和切换效果（入场、强调、退场、路径）都将在输出文件中消失。如需保留动画，请在模板套用完成后在 PowerPoint 中手动重新添加。是否继续？

Only proceed after the user confirms.

**Step 1 — Analyze both files**

```bash
# Extract and display source content
python scripts/extract_content.py source.pptx

# View template layouts
python scripts/thumbnail.py template.pptx

# Extract template structure details
python -m markitdown template.pptx
```

Read the thumbnail grid to understand template slide layouts. Read the extract_content output to understand source slide structure.

**Step 2 — Plan the mapping**

For each source slide:
1. Identify the slide type (title, content, two-column, image+text, conclusion, etc.)
2. Find the best matching layout in the template
3. Note any content that won't fit (too many bullet points, missing images, etc.)

Create a mapping table like:
```
Source slide 1 (title) → Template slide 1 (title layout)
Source slide 2 (agenda) → Template slide 3 (list layout)
Source slide 3 (two columns) → Template slide 5 (two-column layout)
...
```

**Step 3 — Run the apply script**

```bash
python scripts/apply_template.py source.pptx template.pptx output.pptx
```

This script:
1. Extracts all text, images, charts from source slides
2. Unpacks the template
3. Duplicates the appropriate template slides for each source slide
4. Populates content into the template's placeholders
5. Packs the result

**Step 4 — Manual fine-tuning (if needed)**

For slides with imperfect automated mapping:

```bash
python scripts/office/unpack.py output.pptx unpacked_output/
# Edit specific slide XML files
python scripts/office/pack.py unpacked_output/ output_final.pptx --original template.pptx
```

**Step 5 — QA**

```bash
# Check content
python -m markitdown output.pptx

# Visual check
python scripts/thumbnail.py output.pptx output_thumb
```

---

## Mode 2: Style Beautify

Redesign the visual appearance of a PPT while preserving all content.

### When to use

- User only provides one PPT and wants it to look better
- User says: "beautify", "redesign", "make it professional", "modernize"

### Workflow

**Step 1 — Analyze the source**

```bash
# Extract content and structure
python scripts/extract_content.py source.pptx

# Visual overview
python scripts/thumbnail.py source.pptx
```

**Step 2 — Choose a theme**

Based on the content topic, select a theme:

| Theme | 中文 | Use when content is about... | 适用场景 |
|-------|------|------------------------------|---------|
| `executive` | 行政深蓝 | Corporate, finance, strategy | 企业战略、财务报告、管理层汇报 |
| `tech` | 科技深色 | Technology, software, AI, data | AI/科技产品、数据分析、投资人路演 |
| `creative` | 创意珊瑚 | Design, marketing, branding, arts | 品牌营销、创意方案、活动策划 |
| `warm` | 温暖陶土 | Education, community, wellness | 教育培训、社区活动、健康wellness |
| `minimal` | 极简炭灰 | Clean reports, academic, research | 学术报告、研究成果、简洁商务 |
| `bold` | 大胆酒红 | Sales pitches, product launches | 销售提案、产品发布、强视觉冲击 |
| `nature` | 自然森绿 | Environment, agriculture, sustainability | 环保、农业、可持续发展 |
| `ocean` | 海洋蓝绿 | Healthcare, science, trust-based topics | 医疗健康、科学研究、信任导向主题 |

**Step 3 — Run the beautify script**

```bash
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech
```

Flags:
- `--theme NAME` — theme name from table above (default: auto-detect from content)
- `--keep-images` — preserve original images (default: True)
- `--font-pair PAIR` — override font pairing (e.g., "georgia-calibri")
- `--no-restructure` — skip layout enrichment (only change colors/fonts)

**Step 4 — Review and iterate**

```bash
python scripts/thumbnail.py output.pptx beautified_thumb
```

Look for:
- Color consistency across slides
- Font size hierarchy (title > header > body)
- Visual variety between slide layouts
- Adequate white space

**Step 5 — Fine-tune if needed**

Use the XML editing approach from editing.md to adjust specific slides.

---

## Design Principles (Applied Automatically by Scripts)

Both transform scripts enforce these rules:

**Color System**
- One dominant color (60-70% of visual weight)
- One supporting tone
- One sharp accent
- Never more than 3 main colors

**Typography Hierarchy**
- Titles: 36-44pt bold
- Section headers: 20-24pt bold  
- Body text: 14-16pt regular
- Captions: 10-12pt muted color

**Layout Variety**
Every 3 consecutive slides must have different layouts. Monotonous presentations are automatically flagged.

**Visual Elements**
Every content slide has at least one non-text element (shape, icon, color block).

**Anti-Patterns Enforced**
- No accent lines under titles
- No centered body text
- No text-only slides
- No more than 6 bullet points per slide
- No low-contrast text

---

## Available Themes Detail

### Executive — 行政深蓝（企业/财务/战略）
```
Primary: #1E2761 (navy)
Secondary: #CADCFC (ice blue)  
Accent: #C9A84C (gold)
Background: #FFFFFF / #F8F9FF
Font: Cambria + Calibri
```

### Tech — 科技深色（AI/软件/数据/路演）
```
Primary: #028090 (teal)
Secondary: #1C2541 (dark navy)
Accent: #02C39A (mint)
Background: #0B0C10 / #FFFFFF
Font: Trebuchet MS + Calibri
```

### Creative — 创意珊瑚（品牌/营销/设计）
```
Primary: #F96167 (coral)
Secondary: #2F3C7E (navy)
Accent: #F9E795 (gold)
Background: #FFFFFF / #FFF8F5
Font: Georgia + Calibri
```

### Warm — 温暖陶土（教育/社区/wellness）
```
Primary: #B85042 (terracotta)
Secondary: #ECE2D0 (sand)
Accent: #84B59F (sage)
Background: #FFFDF9 / #F5F0E8
Font: Palatino + Calibri
```

### Minimal — 极简炭灰（学术/研究/简洁商务）
```
Primary: #36454F (charcoal)
Secondary: #F2F2F2 (off-white)
Accent: #212121 (black)
Background: #FFFFFF
Font: Arial Black + Arial
```

### Bold — 大胆酒红（销售/产品发布/强视觉）
```
Primary: #990011 (cherry)
Secondary: #2F3C7E (navy)
Accent: #FCF6F5 (near white)
Background: #1A1A2E / #FFFFFF
Font: Impact + Arial
```

### Nature — 自然森绿（环保/农业/可持续）
```
Primary: #2C5F2D (forest)
Secondary: #97BC62 (moss)
Accent: #F5F5F5 (cream)
Background: #FAFFF5 / #FFFFFF
Font: Georgia + Calibri
```

### Ocean — 海洋蓝绿（医疗/科学/信任导向）
```
Primary: #065A82 (deep blue)
Secondary: #1C7293 (teal)
Accent: #9FFFCB (mint)
Background: #FFFFFF / #F0F8FF
Font: Calibri Bold + Calibri
```

---

## Scripts Reference

### extract_content.py

Extracts all content from a PPTX into a structured JSON format.

```bash
python scripts/extract_content.py source.pptx [--output content.json]
```

Output format:
```json
{
  "slides": [
    {
      "index": 1,
      "type": "title",  // auto-detected: title|section|content|image|quote|end
      "title": "...",
      "body": ["bullet 1", "bullet 2"],
      "notes": "...",
      "has_images": false,
      "image_paths": [],
      "has_charts": false,
      "layout_hint": "title_slide"
    }
  ],
  "total_slides": 12,
  "topic_keywords": ["AI", "strategy", "2024"]
}
```

### apply_template.py

Applies a template's visual style to a source PPT's content.

```bash
python scripts/apply_template.py source.pptx template.pptx output.pptx [options]

Options:
  --mapping FILE     JSON file with manual slide mapping (overrides auto-mapping)
  --dry-run          Print mapping plan only — no output file written
  --save-mapping F   Save auto-generated mapping to JSON (pair with --dry-run)
  --keep-notes       Preserve speaker notes (default: True)
  --verbose          Show detailed mapping decisions
```

**Recommended workflow:**

```bash
# Step 1: Review the mapping plan before committing
python scripts/apply_template.py source.pptx template.pptx out.pptx --dry-run

# Step 2: (Optional) Save and edit the mapping if auto-mapping is wrong
python scripts/apply_template.py source.pptx template.pptx out.pptx --dry-run --save-mapping mapping.json
# Edit mapping.json to fix slide assignments
python scripts/apply_template.py source.pptx template.pptx out.pptx --mapping mapping.json

# Step 3: Execute
python scripts/apply_template.py source.pptx template.pptx out.pptx
```

### patch_slide.py

Lightweight text patcher — finds and replaces text without unpacking.
Use for single-point fixes: typos, title changes, number updates.

```bash
# Preview first (default — no file written)
python scripts/patch_slide.py deck.pptx --find "Draft" --replace "Final"

# Search only (no replacement specified)
python scripts/patch_slide.py deck.pptx --find "revenue"

# Apply in place
python scripts/patch_slide.py deck.pptx --find "Draft" --replace "Final" --confirm

# Apply to specific slides only
python scripts/patch_slide.py deck.pptx --find "TBD" --replace "Q2 2026" --slides 3,5 --confirm

# Write to new file (safer — original untouched)
python scripts/patch_slide.py deck.pptx --find "CEO" --replace "CFO" --output fixed.pptx --confirm

# Batch from JSON file
python scripts/patch_slide.py deck.pptx --patch-file patches.json --confirm
```

Patch file format:
```json
[
  {"find": "DRAFT", "replace": "FINAL"},
  {"find": "Q1", "replace": "Q2", "slides": [2, 4, 6]}
]
```

**Use `patch_slide.py` when:** fixing a word, updating a date, correcting a number  
**Use Editing Workflow when:** changing layout structure, adding/deleting slides, repositioning elements

### beautify_ppt.py

Redesigns a PPT's visual style while preserving content.

```bash
python scripts/beautify_ppt.py source.pptx output.pptx [options]

Options:
  --theme NAME         Theme name (default: auto-detect)
  --keep-images        Preserve original images (default: True)
  --font-pair PAIR     Font pair: georgia-calibri, arial-calibri, etc.
  --dark-mode          Use dark background variant
  --no-restructure     Skip layout restructuring (colors/fonts only)
```

**When to use `--no-restructure`:**

| 场景 | 建议 |
|------|------|
| 源 PPT 是纯文字，布局简单 | 不加（默认开启，脚本自动丰富布局） |
| 源 PPT 有精心排版（SmartArt、多列、自定义占位符） | 加上，防止脚本破坏原有结构 |
| 只想换配色和字体，不改版式 | 加上 |
| 输出必须与源文件版式完全一致 | 加上 |

Example — only change colors/fonts, preserve original layouts:
```bash
python scripts/beautify_ppt.py source.pptx output.pptx --theme executive --no-restructure
```

**Layout restructuring** (enabled by default): For content slides, the script
automatically varies the visual layout across slides to prevent monotony:

| Variant | Description |
|---------|-------------|
| `accent_bar` | Thin vertical color bar on the left (default) |
| `numbered_list` | Colored number circles alongside each bullet |
| `stat_highlight` | First item promoted to large rounded callout box |
| `two_tone` | Left 35% colored panel + right 65% content |
| `header_band` | Full-width colored top band above content |

The variant is chosen automatically to avoid 3+ consecutive slides using
the same layout. Use `--no-restructure` to disable if the source PPT already
has complex custom layouts.

---

## Manual XML Editing Reference

When scripts produce imperfect results, edit XML directly.

**Find text in a slide:**
```bash
grep -n "placeholder_text" unpacked/ppt/slides/slide3.xml
```

**Replace a color throughout the presentation:**
```bash
# Find all color references
grep -rn "1E2761" unpacked/ppt/
```

**Change slide background:**
```xml
<!-- In slide XML, inside <p:cSld> -->
<p:bg>
  <p:bgPr>
    <a:solidFill>
      <a:srgbClr val="1E2761"/>
    </a:solidFill>
  </p:bgPr>
</p:bg>
```

**Change text color in a run:**
```xml
<a:rPr lang="en-US" sz="2800" b="1" dirty="0">
  <a:solidFill>
    <a:srgbClr val="FFFFFF"/>
  </a:solidFill>
</a:rPr>
```

---

## QA (Required)

**Assume there are problems. Your job is to find them.**

Your first render is almost never correct. Approach QA as a bug hunt, not a confirmation step.

### Step 1 — Automated QA (always run first)

```bash
python scripts/qa_check.py output.pptx
```

This runs 10 structural checks without requiring LibreOffice:

| Check | What it catches |
|-------|----------------|
| `overflow` | Text likely overflowing its bounding box |
| `contrast` | Low text/background contrast (WCAG AA) |
| `empty` | Slides with no text or images |
| `placeholder` | Leftover "Click to edit" / lorem ipsum text |
| `bullets` | More than 6 bullet points on a slide |
| `fontsize` | Body text <12pt or titles <20pt |
| `offslide` | Shapes positioned outside slide boundaries |
| `duplicates` | Consecutive slides with the same title |
| `titles` | Content slides missing a title |
| `monotony` | 3+ consecutive slides with the same layout |

Run targeted checks when you know what you're fixing:

```bash
# Only check for critical issues
python scripts/qa_check.py output.pptx --min-severity warning

# Only check specific things (e.g., after a beautify)
python scripts/qa_check.py output.pptx --only overflow,contrast,placeholder

# Save JSON report for programmatic processing
python scripts/qa_check.py output.pptx --output qa_report.json

# Exit code 1 if any warnings/errors found (useful in workflows)
python scripts/qa_check.py output.pptx --strict
```

Fix all `error` severity issues before proceeding. Fix `warning` issues unless there's a specific reason not to.

### Step 2 — Content QA

```bash
python -m markitdown output.pptx
```

Check for missing content, typos, wrong order. The `placeholder` check above catches most leftover text automatically, but scan manually for content that looks semantically wrong.

### Step 3 — Visual QA (for beautify / template-apply outputs)

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
```

### Verification Loop

1. **Run `qa_check.py`** → fix all errors and warnings
2. **Visual inspect** affected slides → list issues found
3. Fix issues
4. **Re-run `qa_check.py`** on the updated file — one fix often creates another problem
5. Repeat until a full automated + visual pass reveals no new issues

**Do not declare success until `qa_check.py` returns zero errors and zero warnings.**

---

## Converting to Images

Convert presentations to individual slide images for visual inspection:

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

- `pip install "markitdown[pptx]"` — text extraction
- `pip install defusedxml Pillow python-pptx` — XML handling, thumbnails, PPTX manipulation
- `npm install -g pptxgenjs react react-dom react-icons sharp` — creating from scratch
- LibreOffice (`soffice`) — PDF/image conversion (auto-configured via `scripts/office/soffice.py`)
- Poppler (`pdftoppm`) — PDF to images

Install all Python deps:
```bash
pip install "markitdown[pptx]" defusedxml Pillow python-pptx
```

---

## Speaker Notes Workflow

自动为每页幻灯片生成演讲者备注（Speaker Notes）。脚本读取每页的标题和正文，生成口头展开内容并写入 notes 字段，演讲者打开 PPT 后可直接在演讲者视图中看到提示。

### 快速开始

```bash
# 最简单用法——无需 API，规则生成（推荐首选）
python scripts/generate_notes.py input.pptx output_with_notes.pptx

# 指定备注风格
python scripts/generate_notes.py input.pptx output.pptx --mode speaker
python scripts/generate_notes.py input.pptx output.pptx --mode coach
python scripts/generate_notes.py input.pptx output.pptx --mode summary

# 先预览再写入
python scripts/generate_notes.py input.pptx output.pptx --dry-run
```

### 三种备注风格（--mode）

| 模式 | 说明 | 适合场景 |
|------|------|---------|
| `speaker`（默认）| 演讲提示：过渡句 + 每点口头展开 + 收尾句 | 正式演讲、路演、汇报 |
| `coach` | 演讲教练：内容密度评估 + 用时建议 + 过渡技巧 | 提高演讲质量 |
| `summary` | 极简摘要：一到两句话概括本页核心 | 快速打印备忘 |

### 三种生成后端（--backend）

| 后端 | 质量 | 速度 | 需要 | 推荐度 |
|------|------|------|------|--------|
| `simple`（默认）| 中 | 极快 | 无 | ★★★★ 首选 |
| `openai` | 高 | 中 | API Key | ★★★ 有 key 时用 |
| `ollama` | 高 | 中 | 本地 Ollama | ★★★ 本地 LLM 时用 |

```bash
# OpenAI 后端
python scripts/generate_notes.py deck.pptx out.pptx --backend openai --api-key sk-xxx

# 或通过环境变量
OPENAI_API_KEY=sk-xxx python scripts/generate_notes.py deck.pptx out.pptx --backend openai

# 本地 Ollama 后端（需先启动 ollama serve）
python scripts/generate_notes.py deck.pptx out.pptx --backend ollama --model llama3
```

### 常用选项

```bash
# 指定语言（默认 auto 自动检测）
python scripts/generate_notes.py deck.pptx out.pptx --language zh  # 强制中文
python scripts/generate_notes.py deck.pptx out.pptx --language en  # 强制英文

# 跳过已有备注的页（不覆盖）
python scripts/generate_notes.py deck.pptx out.pptx --no-overwrite

# 先 dry-run 预览，确认效果后再写入
python scripts/generate_notes.py deck.pptx out.pptx --dry-run
python scripts/generate_notes.py deck.pptx out.pptx           # 确认后正式写入
```

**`--language auto` 检测规则：** 逐页统计标题 + 正文中 Unicode 范围 `\u4e00–\u9fff` 的汉字占比，超过 10% 则判定为中文，否则为英文。**中英混排的页面**（如英文标题 + 中文正文）汉字占比通常仍会超过阈值，会被判定为中文。若整份 PPT 语言需要统一，请显式传 `--language zh` 或 `--language en`。

### 完整工作流

**Step 1 — 预览现有内容**

```bash
python -m markitdown input.pptx
```

了解每页标题和正文，以便选择合适的 mode 和 language。

**Step 2 — dry-run 预览备注**

```bash
python scripts/generate_notes.py input.pptx output.pptx --dry-run --mode speaker
```

检查生成的备注质量，确认风格和语言是否符合预期。

**Step 3 — 正式生成**

```bash
python scripts/generate_notes.py input.pptx output.pptx --mode speaker
```

**Step 4 — 验证**

```bash
# 用 markitdown 检查备注是否写入
python -m markitdown output.pptx
# 备注内容会出现在每页的 Notes 部分
```

用 PowerPoint 或 Keynote 打开 output.pptx，切换到"演讲者视图"确认备注可见。

### 典型输出示例

**speaker 模式（中文）：**
```
接下来我们来看……「市场竞争分析」

关键要点提示：
  1. 当前市场份额 35% — 展开说明背景或数据支撑。
  2. 竞争对手数量增加 20% — 展开说明背景或数据支撑。
  3. 核心差异化优势 — 展开说明背景或数据支撑。

记住这个核心信息：市场竞争分析。
```

**coach 模式（中文）：**
```
【演讲教练提示】
本页主题：市场竞争分析
内容量适中（3 条），正常节奏即可。

建议用时：约 1-2 分钟。
转场建议：结束本页时给听众留 3 秒停顿，再进入下一页。
```

**summary 模式（中文）：**
```
本页核心：市场竞争分析
要点：当前市场份额 35%；竞争对手数量增加 20%；核心差异化优势。
```

### 注意事项

- `simple` 后端生成规则化的提示框架，不"理解"内容；如需更自然的备注，用 `openai` 或 `ollama` 后端
- 备注会覆盖现有内容，如需保留原备注请加 `--no-overwrite`
- 图表页（无文字）会生成通用的图表说明提示
- 生成后建议在演讲者视图中人工审阅，根据实际讲解习惯调整

---

## Known Limitations

| Scenario | Issue | Workaround |
|----------|-------|------------|
| `apply_template.py` with complex charts | Charts are extracted as images, not live chart objects | Manually re-insert charts after applying template |
| `apply_template.py` when slide counts differ greatly | Auto-mapping may repeat or skip template layouts | Use `--mapping` flag to provide a manual mapping JSON |
| `beautify_ppt.py` on slides with custom positioned elements | Script may reposition elements to fit the new theme grid | Run visual QA and manually adjust offending slides via XML |
| Very long bullet text | May overflow text boxes after theme change | Shorten content or adjust font size in XML after beautify |
| Embedded fonts in source PPT | Fonts may not transfer to output | Install the same fonts locally, or substitute with theme fonts |
| LibreOffice not installed | `soffice` conversion for image QA will fail | Install via `brew install libreoffice` (macOS) |
