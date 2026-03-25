---
name: pptx-studio
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file; editing, modifying, or updating existing presentations; combining or splitting slide files; merging multiple PPTX files into one; working with templates, layouts, speaker notes, or comments. Also use when: (1) the user wants to apply or swap a template style — phrases like 'apply template', 'use this template', 'fit into template', 'change my PPT style', 'switch template', 'replace template', 'change the template', 'switch the template', 'use a different template', '套用模板', '换模板', '更换模板', '换个模板', '替换模板', '把PPT套入模板', '套入指定模板', '模板转换', '换一套模板', '更改模板', '切换模板'; (2) the user wants to beautify or redesign a PPT — phrases like 'beautify', 'redesign', 'make it look better', 'improve design', 'modernize slides', '美化PPT', '优化设计', '让PPT更好看', '重新设计风格', '改造PPT外观', '美化幻灯片'; (3) the user wants to generate or add speaker notes — phrases like 'add speaker notes', 'generate notes', 'write notes for presenter', 'add talking points', 'create presentation notes', 'notes for each slide', '加备注', '生成演讲者备注', '写备注', '添加演讲提示', '为每页写备注', '自动生成备注', '演讲者视图备注'; (4) the user wants to merge or combine multiple PPTX files — phrases like 'merge pptx', 'combine presentations', 'join slides', 'concatenate pptx', '合并PPT', '合并幻灯片', '拼接PPT', '把多个PPT合并', '合在一起', '将两个PPT合成一个'. Trigger whenever the user mentions 'deck', 'slides', 'presentation', or references a .pptx filename."
---

# PPTX Studio

Handle all .pptx tasks — create from scratch, edit existing files, apply templates, and beautify designs.

📦 **Source Code**: https://github.com/chouraycn/PPTX-Studio

## Decision Flow

**Start here every time.** Before doing anything, figure out which mode to use:

```
User provides TWO .pptx files?
  → Mode 1: Template Apply (三阶段：AI审校内容 → 套入模板 → AI逐页美化)
  ℹ️  Animation note: apply_template 会自动迁移源 PPT 的动画时间轴结构，
      但因形状 ID 重新分配，目标形状可能需在 PowerPoint 中手动重新绑定。
      备注（Speaker Notes）完整保留。

User provides ONE .pptx file + says "beautify / redesign / make it look better"?
  → Mode 2: Style Beautify

User provides ONE .pptx file + says "edit / update / change content / add slides"?
  → If it's a single text fix (typo, title, number): patch_slide.py (fast path)
  → If it's structural (add/delete/reorder slides): Editing Workflow (unpack → edit XML → pack)

User provides ONE .pptx file + says "read / extract / summarize / what's in this"?
  → Reading Content (markitdown)

User provides ONE .pptx file + says "add speaker notes / generate notes / write notes for me / 加备注 / 生成演讲者备注 / 写备注"?
  → Speaker Notes Workflow

User provides ONE .pptx file + says "change color / 换颜色 / 改颜色 / 橙色换蓝色 / 整体换色"?
  → Global Color Replacement (color_replacement.py)

User provides TWO OR MORE .pptx files + says "merge / combine / join / concatenate / 合并 / 拼接 / 合在一起"?
  → Mode 6: Merge PPT

User provides NO file + wants a new presentation?
  → Creating from Scratch (pptxgenjs)

Still unclear?
  → Ask: "您想对PPT做什么？套用模板、美化风格、编辑内容、合并文件，还是从头创建？"
```

**Quick decision table:**

| User says... | Mode |
|--------------|------|
| 给两个pptx，套/换/应用模板 | Mode 1: Template Apply |
| 美化、优化、让它更好看 | Mode 2: Style Beautify |
| 编辑、更新、修改内容 | Mode 3: Editing Workflow |
| 读取、提取、总结内容 | Mode 4: Reading Content (markitdown) |
| 加备注、生成演讲者备注、写备注 | Mode 5: Speaker Notes Workflow |
| 换颜色、改颜色、整体换色、橙色换蓝色 | **Global Color Replacement** ⭐ |
| 合并、拼接、合在一起 | Mode 6: Merge PPT |
| 从零创建、新建PPT | Mode 7: Creating from Scratch (pptxgenjs) |
| 更换模板、切换模板、换个风格 | Mode 1（有模板文件）或 Mode 2（无模板文件） |
| 修改内容、调整文字、增减页 | patch_slide（单点文字）或 Editing Workflow（结构性修改） |
| 读取、提取、总结内容 | Reading Content |
| 加备注、写演讲提示、生成 Speaker Notes | Speaker Notes Workflow |
| 合并多个PPT、拼接、合在一起 | Mode 6: Merge PPT |
| 做一个新PPT | Creating from Scratch |

---

## Quick Reference

| Task | Guide |
|------|-------|
| Read/analyze content | `python -m markitdown presentation.pptx` |
| Edit existing PPT | Editing Workflow section below |
| Create from scratch | Creating from Scratch section below |
| Apply template (with AI review + beautify) | Mode 1 below — 三阶段完整流程 |
| AI content review only | Mode 1 阶段 A |
| AI per-slide beautify only | Mode 1 阶段 C |
| Beautify PPT style | Mode 2 below |
| Generate speaker notes | Speaker Notes Workflow below |
| Merge multiple PPTXs | Mode 6: Merge PPT below |
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

Two powerful transformation modes — both now include **AI-powered stages** and **enhanced visual features**:

| Mode | Description |
|------|-------------|
| **Template Apply** | 三阶段：① AI审校内容（纠错/精简/补全） → ② 脚本套入模板（配色/字体/布局） → ③ AI逐页美化（排版/字数控制/空行清理） |
| **Style Beautify** | Redesign visual style with **12 themes**, **10 layout variants**, **smart image enhancement**, and **auto icon insertion** |

### Enhanced Visual Features (New)

**12 Professional Themes:**
- Classic: `executive`, `tech`, `creative`, `warm`, `minimal`, `bold`, `nature`, `ocean`
- New: `elegant` (深灰蓝+珊瑚红), `modern` (紫罗兰+粉色), `sunset` (暖橙+金黄), `forest` (深林绿+薄荷)

**10 Layout Variants (Auto-rotated for visual variety):**
1. `accent_bar` — 左侧强调条（默认）
2. `numbered_list` — 数字圆圈列表
3. `stat_highlight` — 首条数据高亮卡片
4. `two_tone` — 双色分栏布局
5. `header_band` — 顶部色带标题
6. `card_grid` — 卡片网格（2×2 自动排列）
7. `timeline` — 时间轴布局
8. `split_diagonal` — 对角分割
9. `image_focus` — 图片聚焦框架
10. `quote_block` — 引用块布局

**Smart Image Enhancement:**
- 自动圆角（roundRect）
- 柔和阴影（outer shadow）
- 主题色边框

**Smart Table Enhancement (AI-powered):**
- 表头行样式：主题色背景 + 白色文字
- 斑马纹行：交替行背景色，提升可读性
- 统一边框：主题色边框，宽度一致
- 单元格边距优化：内边距标准化
- 字体统一：应用主题字体

**Smart Typography:**
- 动态字号（根据内容密度自动调整）
- 1.2倍行距优化
- 智能图标插入（基于关键词匹配 50+ 图标）

**Gradient Backgrounds:**
- 标题页/章节页可使用渐变色背景
- 基于主题色的自动渐变

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

Apply an existing template's visual identity to your content — with **two AI-powered stages** to ensure content accuracy and per-slide visual quality before and after the template is applied.

### When to use

- User provides a source PPT (their content) and a template PPT (the desired look)
- User says: "apply this template", "make my PPT look like this", "use this style"

### Overview — 三阶段流程

```
阶段 A: 内容提取 + AI 审校纠错
  ↓ AI 检查内容完整性、纠正错误、调整结构
阶段 B: 套入模板（脚本自动执行）
  ↓ apply_template.py 将审校后内容注入模板
  ↓ 自动替换形状/表格颜色 + 自动化排版美化
  ↓ [--beautify] AI 完整美化（布局/配色/字体/表格）← 新增
阶段 C: AI 逐页美化调整 + 备注增强
  ↓ AI 排版优化 + 保留原备注并追加 AI 摘要
最终输出：内容准确 + 视觉精良 + 备注完善的 PPT
```

> ℹ️ **动画保留说明：** `apply_template` 会迁移源 PPT 中所有幻灯片的动画时间轴结构（入场顺序、延迟、触发方式），但由于幻灯片形状 ID 在模板套用过程中会重新分配，部分动画的目标形状可能需要在 PowerPoint 里手动重新绑定。切换效果（Transitions）和备注（Speaker Notes）均完整保留。

---

### 阶段 A：内容提取 + AI 审校纠错

**Step A1 — 提取源 PPT 内容**

```bash
python scripts/extract_content.py source.pptx --output content.json
```

同时获取缩略图以便理解原始视觉结构：

```bash
python scripts/thumbnail.py source.pptx source_thumb/
```

**Step A2 — AI 内容审校（必做步骤）**

读取 `content.json` 后，作为 AI 你必须对每一页内容执行以下审校：

**审校任务列表：**

1. **标题完整性** — 检查每页 `title` 是否为空或过短（< 3字）；若为空则根据 body 内容推断补写
2. **正文内容纠错** — 检查 `body` 文本中的：
   - 明显错别字、标点错误
   - 截断/不完整的句子（常见于 OCR 识别源或复制粘贴）
   - 重复词语或重复要点
3. **页面类型判断** — 结合 `title`/`body`/`subtitle` 判断每页是否被正确分类（`type` 字段）；如 section 页被误识别为 content，应修正
4. **要点精简** — 若某页 `body` 超过 6 条要点，分析是否可以合并；超过 8 条要点时必须合并
5. **内容一致性** — 跨页检查：前后页提到相同概念时表述是否一致（如产品名、数字、缩写）
6. **缺失内容补全** — 若第一页无标题/副标题，根据全文主题推断补写；若结语页内容为空，补写一句话

**AI 审校输出格式：**

审校完成后，将修改后的内容以**对比形式**展示给用户：

```
📋 内容审校报告

第 2 页（原类型: content）
  标题: "产品介绍" → 无变化
  正文修改:
    原: "用户可以通过以下方式访问系功能..."
    改: "用户可以通过以下方式访问系统功能..." [纠正错别字]
    原: 8条要点（超出限制）
    改: 合并为6条 [精简]

第 5 页（原类型: section）
  发现问题: type 被识别为 content，实为 section 过渡页
  修正: type → "section"

[若无修改] 第 3 页 — 无问题 ✓
...

共审校 N 页，发现 M 处问题，已修正。是否继续套入模板？
```

**展示审校报告后，等待用户确认（"继续" / 或指出不同意见）再进入阶段 B。**

用户如有不同意见，按用户指示调整后再继续。

---

### 阶段 B：套入模板（脚本执行）

**Step B1 — 分析模板**

```bash
# 查看模板布局缩略图
python scripts/thumbnail.py template.pptx template_thumb/

# 查看模板文字结构
python -m markitdown template.pptx
```

**Step B2 — 执行套模板**

```bash
python scripts/apply_template.py source.pptx template.pptx output.pptx
```

**进阶：添加最终 AI 美化**
```bash
# 在套模板完成后，对输出进行完整的设计美化（重新设计布局、配色、字体、表格等）
python scripts/apply_template.py source.pptx template.pptx output.pptx --beautify

# 指定美化主题
python scripts/apply_template.py source.pptx template.pptx output.pptx --beautify --beautify-theme tech
```

脚本自动完成：
1. 从源 PPT 提取文字、图片、格式（bold/italic/size）
2. 解包模板，为每页源幻灯片找到最匹配的模板布局
3. 将内容注入模板占位符，使用模板配色/字体
4. **自动替换自定义形状和表格颜色** — 所有非占位符元素（装饰形状、表格、SmartArt等）自动使用模板配色
5. **自动美化排版** — 标题截短（>20字）、要点合并（>6条）、段落行距优化
6. 迁移动画时间轴结构和 Speaker Notes
7. 打包输出文件

> **如需覆盖自动映射方案：**
> ```bash
> python scripts/apply_template.py source.pptx template.pptx output.pptx --dry-run
> # 查看映射方案，保存到文件
> python scripts/apply_template.py source.pptx template.pptx output.pptx --dry-run --save-mapping mapping.json
> # 编辑 mapping.json 后使用自定义映射执行
> python scripts/apply_template.py source.pptx template.pptx output.pptx --mapping mapping.json
> ```

**Step B3 — 生成输出缩略图，准备 AI 美化**

```bash
python scripts/thumbnail.py output.pptx output_thumb/
```

---

### 阶段 C：AI 逐页美化调整

套模板脚本完成后，AI 需要对输出 PPT **逐页进行内容感知的排版优化**。这是纯文字/XML 层面的精调，不调用外部渲染工具。

**Step C1 — 解包输出文件**

```bash
python scripts/office/unpack.py output.pptx output_unpacked/
```

**Step C2 — AI 逐页分析 + 优化**

针对每一页幻灯片（`output_unpacked/ppt/slides/slideN.xml`），执行以下判断和处理：

#### 美化规则 — 按页面类型

**封面页 / 标题页（type: title）**
- 检查 title 字数：若超过 20 字，建议截短主标题，将剩余部分移入 subtitle（用 `patch_slide.py` 修改）
- 检查 subtitle 是否存在：若无 subtitle 但 body 有内容，将 body 第一条提升为 subtitle
- 标题文字应 bold，字号 ≥ 36pt；若小于此值，使用 patch_slide.py 调整

**内容页（type: content）**
- 要点数量：4-6 条最佳；若 ≤ 3 条，可保持不变或加一句引导语；若超 6 条，合并相似条目
- 每条要点字数：建议 ≤ 25 字；若某条超过 40 字，拆分成两条
- 检查是否有「孤行」（body 第一行是大标题式文字而非要点）—— 若有，将其提升为幻灯片 title（若当前 title 为空）或作为独立 subtitle
- 若 `has_images: true`，保留图片；检查正文是否为图注格式（短句），若是则字号不低于 14pt

**章节页（type: section）**
- body 内容通常应为空或一句话；若 body 超过 2 条，将其移到下一页（新建 content 页）并清空当前页 body
- title 居中对齐时检查 `<a:pPr algn="ctr">`；若不居中，不强制修改（保持模板默认）

**结语页（type: end / conclusion）**
- 若 title 为"谢谢"/"Thank You"/"END"等，body 若为空则补充一句简短结语（可使用 AI 生成）
- 若 body 有多余联系方式信息，移动到 notes（Speaker Notes）而非幻灯片正文

**所有页面通用检查**
- 连续 3 页布局类型相同：第 3 页建议切换为不同的模板 layout（通过 `--mapping` 参数在下次运行时调整；或在 XML 中修改 layout 关联）
- 空行清理：删除 body 中首尾多余的空行（`<a:p><a:endParaRPr/></a:p>`）
- 若某页 body 和 title 均为空：标记为"疑似多余页"，提示用户确认是否保留

**Step C3 — 应用优化**

使用 `patch_slide.py` 批量修改文字：

```bash
# 修改特定页面文字
python scripts/patch_slide.py output.pptx --find "原文字" --replace "新文字"
```

对于结构性调整（调整段落数量、移动内容）使用 XML 直接编辑：

```bash
python scripts/office/unpack.py output.pptx output_unpacked/
# 编辑 output_unpacked/ppt/slides/slideN.xml
python scripts/office/pack.py output_unpacked/ output_final.pptx --original template.pptx
```

**Step C-Notes — 备注增强（保留原备注 + 追加 AI 摘要）**

套模板后，每页备注需要在**保留源 PPT 原有备注**的基础上，追加 AI 生成的内容摘要，方便演讲者快速查看当页核心。

```bash
# 最简用法（规则摘要，无需 API）
python scripts/generate_notes.py output_final.pptx output_final.pptx --append-summary

# 使用 OpenAI 生成更高质量摘要
python scripts/generate_notes.py output_final.pptx output_final.pptx --append-summary --backend openai --api-key sk-xxx

# 使用本地 Ollama
python scripts/generate_notes.py output_final.pptx output_final.pptx --append-summary --backend ollama --model llama3
```

生成后每页备注结构如下：

```
[原有备注内容（从源 PPT 迁移的演讲稿、说明等，原样保留）]

────────────────────────
【AI 摘要】

本页核心：产品核心功能介绍
要点：多租户架构；弹性扩缩容；一键部署；安全审计。
（含图表）
```

> ℹ️ **智能去重：** 若已运行过 `--append-summary`，再次执行时脚本会自动检测备注中是否已含「AI 摘要」标记，跳过已处理的页面，避免重复追加。

**Step C4 — AI 美化报告**

完成逐页调整后，向用户展示：

```
✨ 逐页美化报告

第 1 页（封面）
  ✓ 标题 "产品年度发布会2024" → 已截短为 "产品年度发布会" + subtitle "2024"
  ✓ 去除首部2个空段落

第 3 页（内容）
  ✓ 将8条要点合并为5条（删除重复的第6、7条）
  ✓ 第2条要点拆分（原超55字）

第 6 页（章节）
  ✓ 移除多余 body 内容（已移至第7页开头）

第 8 页（结语）
  ✓ 补充结语：「感谢您的关注，期待与您的合作。」

共处理 N 页，N 处调整已完成。
```

**Step C5 — 最终 QA**

```bash
# 内容完整性检查
python -m markitdown output_final.pptx

# 视觉检查缩略图
python scripts/thumbnail.py output_final.pptx final_thumb/

# 质量评分
python scripts/qa_check.py output_final.pptx
```

---

### 完整命令速查

```bash
# 阶段 A: 提取内容
python scripts/extract_content.py source.pptx --output content.json
python scripts/thumbnail.py source.pptx source_thumb/
# → AI 审校 content.json，向用户展示修改报告，等待确认

# 阶段 B: 套入模板
python scripts/thumbnail.py template.pptx template_thumb/
python scripts/apply_template.py source.pptx template.pptx output.pptx
python scripts/thumbnail.py output.pptx output_thumb/

# 阶段 C: 逐页美化 + 备注增强
python scripts/office/unpack.py output.pptx output_unpacked/
# → AI 逐页分析 XML，执行排版优化
python scripts/office/pack.py output_unpacked/ output_final.pptx --original template.pptx
# → 保留原备注，追加 AI 摘要（智能去重，可安全重复执行）
python scripts/generate_notes.py output_final.pptx output_final.pptx --append-summary
python scripts/qa_check.py output_final.pptx
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

Based on the content topic, select from 12 professional themes:

| Theme | Style | Best for |
|-------|-------|----------|
| `executive` | 商务深蓝+金色 | 企业汇报、投资路演 |
| `tech` | 科技青绿+深灰 | 技术分享、产品发布 |
| `creative` | 活力珊瑚+海军蓝 | 创意提案、设计展示 |
| `warm` | 暖陶土+鼠尾草绿 | 教育、非营利组织 |
| `minimal` | 极简炭灰+纯白 | 学术报告、简洁风格 |
| `bold` | 大胆樱桃红+深蓝 | 冲击力强的演讲 |
| `nature` | 森林绿+苔藓 | 环保、健康、农业 |
| `ocean` | 深海蓝+薄荷绿 | 旅游、海洋、科技 |
| `elegant` | 深灰蓝+珊瑚红 | 高端商务、奢侈品 |
| `modern` | 紫罗兰+粉色渐变 | 互联网、时尚、创新 |
| `sunset` | 暖橙+金黄渐变 | 能源、餐饮、温暖主题 |
| `forest` | 深林绿+薄荷 | 可持续发展、有机产品 |

**Enhanced Options:**

```bash
# Basic beautify
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech

# With gradient backgrounds for title slides
python scripts/beautify_ppt.py source.pptx output.pptx --theme modern --gradient-bg

# With smart icons (auto-insert based on keywords)
python scripts/beautify_ppt.py source.pptx output.pptx --theme creative --smart-icons

# Full enhancement (gradient + icons + image enhancement)
python scripts/beautify_ppt.py source.pptx output.pptx --theme elegant --gradient-bg --smart-icons

# Skip layout restructuring (only colors/fonts)
python scripts/beautify_ppt.py source.pptx output.pptx --theme minimal --no-restructure

# Skip image enhancement (if you want original image styling)
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech --no-image-enhance
```

**Layout Variants (Auto-applied for variety):**

The beautifier automatically rotates through 10 layout variants to prevent visual monotony:

1. **accent_bar** — 左侧垂直强调条（最常用）
2. **numbered_list** — 大号数字圆圈列表（适合 4+ 要点）
3. **stat_highlight** — 首条数据高亮卡片（适合含数字的要点）
4. **two_tone** — 双色分栏（每第4页自动应用）
5. **header_band** — 顶部色带标题
6. **card_grid** — 卡片网格 2×2（自动排列 2-4 个要点）
7. **timeline** — 时间轴（适合流程性内容）
8. **split_diagonal** — 对角分割（视觉冲击）
9. **image_focus** — 图片聚焦框架（含装饰边框）
10. **quote_block** — 引用块（带引号和左侧强调条）

> 💡 **Tip:** The system automatically selects the best layout based on content type and avoids using the same layout 3 times in a row.

**Smart Features:**

- **Dynamic Font Sizing:** Automatically adjusts body text size based on content density (dense content → slightly smaller, sparse content → slightly larger)
- **Image Enhancement:** All images get rounded corners, subtle shadows, and theme-colored borders
- **Smart Icons:** When using `--smart-icons`, keywords like "growth", "team", "technology" automatically get relevant emoji icons prepended
- **Paragraph Spacing:** 1.2x line spacing applied for better readability
- **Gradient Backgrounds:** Title/section slides can have beautiful gradients when using `--gradient-bg`

**Step 3 — Run the beautify script**

```bash
# Basic usage
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech

# Full enhancement with all features
python scripts/beautify_ppt.py source.pptx output.pptx --theme elegant --gradient-bg --smart-icons

# Conservative approach (colors/fonts only)
python scripts/beautify_ppt.py source.pptx output.pptx --theme minimal --no-restructure
```

> **判断是否加 `--no-restructure`：** 源文件已有精心排版（SmartArt、多列、自定义占位符）或只想换配色/字体而不改版式时，加上此 flag；纯文字简单布局时不加（默认自动丰富）。详见 [beautify_ppt.py 说明](#beautify_pptpy)。

**All Available Flags:**
| Flag | Description |
|------|-------------|
| `--theme NAME` | Theme name (default: auto-detect from content) |
| `--dark-mode` | Force dark background on all slides |
| `--keep-images` | Preserve original images (default: True) |
| `--font-pair PAIR` | Override font pairing (e.g., "georgia-calibri") |
| `--no-restructure` | Skip layout enrichment (only change colors/fonts) |
| `--no-image-enhance` | Skip image enhancement (rounded corners, shadows, borders) |
| `--gradient-bg` | Use gradient backgrounds for title/section slides |
| `--smart-icons` | Auto-insert icons based on content keywords |
| `--verbose` | Show detailed processing information |
| `--list-themes` | Display all available themes and exit |

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
Secondary: #84B59F (sage)
Accent: #ECE2D0 (sand)
Background: #FFFDF9 / #F5F0E8
Font: Palatino Linotype + Calibri
```

### Minimal — 极简炭灰（学术/研究/简洁商务）
```
Primary: #36454F (charcoal)
Secondary: #F2F2F2 (off-white)
Accent: #212121 (black)
Background: #FFFFFF
Font: Calibri + Calibri
```

### Bold — 大胆酒红（销售/产品发布/强视觉）
```
Primary: #990011 (cherry)
Secondary: #2F3C7E (navy)
Accent: #FCF6F5 (near white)
Background: #1A1A2E / #FFFFFF
Font: Arial Black + Arial
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

### Elegant — 优雅深灰蓝（高端商务/咨询）
```
Primary: #2C3E50 (deep slate blue)
Secondary: #E8E8E8 (light silver)
Accent: #E74C3C (coral red)
Background: #FAFAFA / #1A1A2E
Font: Georgia + Calibri
```

### Modern — 现代紫罗兰（互联网/时尚/创新）
```
Primary: #6C5CE7 (soft purple)
Secondary: #A29BFE (light lavender)
Accent: #FD79A8 (pink)
Background: #F8F9FA / #2D3436
Font: Segoe UI + Segoe UI
```

### Sunset — 日落暖橙（能源/餐饮/生活方式）
```
Primary: #E17055 (burnt orange)
Secondary: #FDCB6E (warm yellow)
Accent: #D63031 (deep red)
Background: #FFF9F0 / #2D142C
Font: Georgia + Calibri
```

### Forest — 森林深绿（可持续发展/环保）
```
Primary: #1B4332 (deep forest)
Secondary: #52B788 (sage green)
Accent: #D8F3DC (pale mint)
Background: #F1F8E9 / #081C15
Font: Cambria + Calibri
```

### color_replacement.py ⭐

**整体换色 + AI 色彩阶梯** — 全局替换 PPT 中的所有颜色（文字、形状、背景、渐变、填充），支持 AI 智能生成多级色彩阶梯。

#### 核心能力

1. **整体换色** — 将整个 PPT 的配色方案从一种颜色换成另一种（如橙色→蓝色）
2. **AI 色彩阶梯** — 基于品牌色自动生成 3-10 级渐变色，智能应用到所有元素
3. **主题间换色** — 在 12 个预设主题之间快速切换配色
4. **预览模式** — 先预览再应用，避免误操作

#### 使用场景

| 场景 | 命令 |
|------|------|
| 橙色换蓝色（整体换色） | `python scripts/color_replacement.py input.pptx output.pptx --replace-primary F96167 0284C7` |
| AI 生成蓝色阶梯（深度5级） | `python scripts/color_replacement.py input.pptx output.pptx --ai-ladder 0284C7 --depth 5` |
| 主题切换（暖色→科技） | `python scripts/color_replacement.py input.pptx output.pptx --theme-from warm --theme-to tech` |
| 预览换色效果 | `python scripts/color_replacement.py input.pptx output.pptx --replace-primary F96167 0284C7 --preview` |

#### 命令参数

**单色替换**：
- `--replace-primary OLD NEW` — 替换主色（如 `F96167 0284C7`）
- `--replace-secondary OLD NEW` — 替换次要色
- `--replace-accent OLD NEW` — 替换强调色

**AI 色彩阶梯**：
- `--ai-ladder BASE_COLOR` — 基于基准色生成色彩阶梯
- `--ladder-depth N` — 阶梯级数（默认 5，范围 3-10）
- `--ladder-strategy STRATEGY` — 策略：`lightness`（明度渐变）、`saturation`（饱和度渐变）、`complementary`（互补色渐变）

**主题换色**：
- `--theme-from THEME` — 源主题（12 种主题名）
- `--theme-to THEME` — 目标主题

**其他选项**：
- `--color-map-file FILE` — 使用自定义 JSON 颜色映射文件
- `--preview` — 预览模式（不实际修改）
- `--verbose, -v` — 显示详细信息

#### 色彩阶梯策略说明

**Lightness（明度渐变，默认）** — 从暗到亮的 5 级渐变：
```
Level 0: 最深色 - 浅背景上的文字、深色元素
Level 1: 较深色 - 次要元素、强调边框
Level 2: 中等色 - 主要内容区域
Level 3: 较浅色 - 第三级元素、浅色背景
Level 4: 最浅色 - 深背景上的文字、高亮元素
```

**Saturation（饱和度渐变）** — 从灰暗到鲜艳的 5 级渐变：
```
Level 0: 灰暗（10% 饱和度）
Level 1: 柔和（32% 饱和度）
Level 2: 平衡（55% 饱和度）
Level 3: 鲜艳（77% 饱和度）
Level 4: 浓烈（100% 饱和度）
```

**Complementary（互补色渐变）** — 从基色跨越到互补色的 5 级渐变：
```
Level 0: 基色（如橙色）
Level 1: 过渡色 1（橙红→红橙）
Level 2: 中性过渡
Level 3: 过渡色 2（蓝绿→青蓝）
Level 4: 互补色（如蓝色）
```

#### 可用主题

| 主题 | 主色 | 次色 | 强调色 | 适用场景 |
|------|------|------|--------|----------|
| executive | 深蓝 | 冰蓝 | 金色 | 商务汇报、投资路演 |
| tech | 青绿 | 深海军蓝 | 薄荷绿 | 技术分享、产品发布 |
| creative | 珊瑚色 | 海军蓝 | 金色 | 创意提案、设计展示 |
| warm | 陶土色 | 鼠尾草绿 | 沙色 | 教育、非营利组织 |
| minimal | 炭灰 | 薰荷色 | 白色 | 学术报告、简洁风格 |
| bold | 樱桃红 | 深蓝 | 金色 | 冲击力强的演讲 |
| nature | 森林绿 | 琥珀色 | 天空蓝 | 环保、健康、农业 |
| ocean | 深海蓝 | 青色 | 浅蓝 | 旅游、海洋、科技 |
| elegant | 深灰蓝 | 橄榄色 | 珊瑚红 | 高端商务、奢侈品 |
| modern | 紫罗兰 | 粉红 | 洋红 | 互联网、时尚、创新 |
| sunset | 暖橙 | 金黄 | 浅黄 | 能源、餐饮、温暖主题 |
| forest | 深林绿 | 翡翠绿 | 薄荷绿 | 可持续发展、有机产品 |

#### 实战示例

**示例 1：品牌色集成**
你的品牌色是 `#0066CC`（蓝色），应用到现有 PPT：

```bash
# 生成 5 级蓝色阶梯并替换所有颜色
python scripts/color_replacement.py presentation.pptx output.pptx \
    --ai-ladder 0066CC \
    --depth 5 \
    --ladder-strategy lightness \
    --verbose
```

**示例 2：主题快速切换**
从暖色主题切换到科技主题：

```bash
# 预览效果
python scripts/color_replacement.py presentation.pptx output.pptx \
    --theme-from warm \
    --theme-to tech \
    --preview

# 应用切换
python scripts/color_replacement.py presentation.pptx output.pptx \
    --theme-from warm \
    --theme-to tech
```

**示例 3：季节性换色**
将秋季（橙色）更新为冬季（蓝色）：

```bash
python scripts/color_replacement.py presentation.pptx output.pptx \
    --replace-primary F96167 0077B6 \
    --replace-secondary F97316 00B4D8 \
    --replace-accent FBBF24 0284C7
```

**示例 4：多品牌 PPT 统一**
将不同品牌的 PPT 统一为科技主题：

```bash
# 品牌 A 的演示文稿
python scripts/color_replacement.py brand_a.pptx unified_a.pptx \
    --theme-from creative \
    --theme-to tech

# 品牌 B 的演示文稿
python scripts/color_replacement.py brand_b.pptx unified_b.pptx \
    --theme-from warm \
    --theme-to tech
```

#### 高级用法

**自定义颜色映射文件**：
创建 JSON 文件（如 `my_colors.json`）：
```json
{
  "F96167": "0284C7",
  "F97316": "0077B6",
  "FBBF24": "00A8E8",
  "B85042": "003366",
  "84B59F": "00D9FF"
}
```

应用：
```bash
python scripts/color_replacement.py input.pptx output.pptx \
    --color-map-file my_colors.json
```

**查看更多示例**：参见 `examples/README.md` 和 `examples/color_maps/` 目录。

> 💡 **Tip**: 始终先用 `--preview` 预览效果，确认后再应用。`--verbose` 会显示详细的颜色统计信息。

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
      "slide_file": "slide1.xml",
      "type": "title",  // auto-detected: title|section|content|image|quote|end
      "title": "...",
      "subtitle": "...",
      "body": ["bullet 1", "bullet 2"],
      "body_rich": [{"text": "...", "bold": true, "size": 28}],
      "notes": "...",
      "has_images": false,
      "image_count": 0,
      "has_charts": false,
      "has_tables": false,
      "table_data": [],
      "layout_file": "slideLayout2.xml",
      "layout_name": "Title Slide",
      "layout_hint": "title_slide",
      "shape_count": 3,
      "background_color": ""
    }
  ],
  "total_slides": 12,
  "topic_keywords": ["AI", "strategy", "2024"],
  "detected_theme": "minimal"
}
```

> **字段说明补充：**
> - `slide_file` — 幻灯片 XML 文件名（如 `slide1.xml`），可直接用于 `patch_slide.py` 或 `add_slide.py` 定位文件
> - `layout_file` — 关联的布局文件名（如 `slideLayout2.xml`），可传给 `add_slide.py` 的 source 参数
> - `layout_hint` — 建议模板布局类型（`title_slide` / `content_slide` / `two_column` 等），供 `apply_template.py` 自动映射使用
> - `body_rich` — 带层级结构的正文（含 bold、italic、size、color 格式信息）
> - `shape_count` — 当前页的形状（`<p:sp>`）总数，可用于判断页面复杂度
> - `background_color` — 幻灯片背景色十六进制值（如 `1A1A2E`），无背景色时为空字符串
>
> 处理有表格、需要按文件名定位幻灯片或精确版式匹配的场景时，请直接读取完整 JSON 输出。

### apply_template.py

Applies a template's visual style to a source PPT's content.

```bash
python scripts/apply_template.py source.pptx template.pptx output.pptx [options]

Options:
  --mapping FILE     JSON file with manual slide mapping (overrides auto-mapping)
  --dry-run          Print mapping plan only — no output file written
  --save-mapping F   Save auto-generated mapping to JSON (pair with --dry-run)
  --no-notes         Do NOT preserve speaker notes (default: notes are kept)
  --verbose, -v      Show detailed mapping decisions and layout analysis
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
  --dark-mode          Force dark background on all slides
  --no-restructure     Skip layout restructuring (colors/fonts only)
  --verbose, -v        Show detailed processing information
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
| `card_grid` | 2×2 card grid layout for multiple items |
| `timeline` | Horizontal timeline with nodes and labels |
| `split_diagonal` | Diagonal split layout for visual impact |
| `image_focus` | Large image area with decorative frame |
| `quote_block` | Centered quote with accent bar and quote marks |

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
soffice --headless --convert-to pdf output.pptx
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
- LibreOffice (`soffice`) — PDF/image conversion (path auto-resolved via `scripts/office/soffice.py` helper; call `soffice` directly, not the helper script)
- Poppler (`pdftoppm`) — PDF to images

Install all Python deps:
```bash
pip install "markitdown[pptx]" defusedxml Pillow python-pptx
```

---

## Mode 6: Merge PPT

将两个或多个 PPTX 文件的幻灯片按顺序合并为一个输出文件。输出文件的尺寸以第一个文件为准。

### 快速开始

```bash
# 最简用法：将 a.pptx 和 b.pptx 全部幻灯片合并（默认顺序 A1,A2,...,B1,B2,...）
python scripts/merge_pptx.py a.pptx b.pptx -o merged.pptx

# 合并三个文件
python scripts/merge_pptx.py a.pptx b.pptx c.pptx -o merged.pptx
```

### 选取指定幻灯片（--slides）

```bash
# a.pptx 取第 1-5 页，b.pptx 取第 2、3、7 页
python scripts/merge_pptx.py a.pptx b.pptx -o merged.pptx --slides "1-5" "2,3,7"

# 所有文件都只取第 1-3 页（单个 range 全局生效）
python scripts/merge_pptx.py a.pptx b.pptx c.pptx -o merged.pptx --slides "1-3"
```

Range 格式：`"1-5"`（连续）或 `"1,3,5"`（离散）或混合 `"1-3,7,9"`。页码从 **1** 开始。

### 自定义排序合并（--order）⭐

`--order` 让你按任意顺序交叉混排多个文件的页面，每个 token 格式为 `<文件序号>:<页码>`（均从 1 开始）。

```bash
# A1 → B1 → A2 → B2 → A3（交错穿插）
python scripts/merge_pptx.py a.pptx b.pptx -o merged.pptx \
    --order 1:1 2:1 1:2 2:2 1:3

# 完全自定义：先 B2，再 A1，再 B1，最后 A3
python scripts/merge_pptx.py a.pptx b.pptx -o merged.pptx \
    --order 2:2 1:1 2:1 1:3

# 三文件排序示例：C1, A2, B3, A1
python scripts/merge_pptx.py a.pptx b.pptx c.pptx -o merged.pptx \
    --order 3:1 1:2 2:3 1:1

# 先 dry-run 确认排序对不对
python scripts/merge_pptx.py a.pptx b.pptx -o merged.pptx \
    --order 1:1 2:1 1:2 2:2 --dry-run
```

> **注意**：`--order` 与 `--slides` 互斥，使用 `--order` 时 `--slides` 会被忽略。每个 token 都可以重复使用同一页（例如 `1:1 1:1` 会输出 A1 两次）。

### 其他选项

```bash
# 不复制演讲者备注
python scripts/merge_pptx.py a.pptx b.pptx -o merged.pptx --ignore-notes

# 预览合并计划（不写文件）
python scripts/merge_pptx.py a.pptx b.pptx -o merged.pptx --dry-run
```

### 完整参数说明

| 参数 | 说明 |
|------|------|
| `FILE ...` | 两个或更多 `.pptx` 输入文件（按顺序）|
| `-o / --output` | 输出文件路径（必填）|
| `--slides RANGE [RANGE ...]` | 每个文件的幻灯片范围；一个值全局生效，多个值按文件顺序匹配；`--order` 存在时忽略 |
| `--order FILE:SLIDE [...]` | 自定义跨文件排序，每个 token 为 `<文件序号>:<页码>`（均 1-based），输出按此顺序排列 |
| `--ignore-notes` | 不复制演讲者备注（默认：复制）|
| `--dry-run` | 仅打印合并计划，不写出文件 |

### 典型工作流

**Step 1 — 确认每个文件的幻灯片数量**

```bash
python -m markitdown a.pptx | grep "^## Slide"
python -m markitdown b.pptx | grep "^## Slide"
```

或用缩略图快速预览：

```bash
python scripts/thumbnail.py a.pptx
python scripts/thumbnail.py b.pptx
```

**Step 2 — dry-run 确认合并计划**

```bash
python scripts/merge_pptx.py a.pptx b.pptx -o merged.pptx --dry-run
# 输出示例：
# [Dry run] Merging 2 files → merged.pptx
#   a.pptx: slides [1, 2, 3] (3 slides)
#   b.pptx: slides [1, 2, 3, 4, 5] (5 slides)
# Total: 8 slides
```

**Step 3 — 正式合并**

```bash
python scripts/merge_pptx.py a.pptx b.pptx -o merged.pptx
```

**Step 4 — QA 验证**

```bash
python scripts/qa_check.py merged.pptx
python -m markitdown merged.pptx
```

### 注意事项

- 输出文件的**幻灯片尺寸**以第一个输入文件为准；若各文件尺寸不同，第二个文件起的幻灯片内容比例可能变形，建议预先统一尺寸。
- **动画和切换效果**：幻灯片内的动画（`<p:timing>`）会被原样保留；但复制后绑定 ID 可能冲突，在 PowerPoint 中打开时部分动画可能需重新触发。
- **嵌入媒体**（图片、音视频）：通过 python-pptx 关系层自动重新映射，通常无需手动处理。
- 合并完成后建议在 PowerPoint 中打开验证，尤其是有复杂嵌入元素的文件。

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

# 自定义 API 端点（兼容 OpenAI API 的本地服务）
python scripts/generate_notes.py deck.pptx out.pptx --backend openai \
    --api-key sk-xxx --base-url http://localhost:11434/v1
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
| `apply_template.py` rich text with mixed bold/plain | bold/italic/size are preserved per run; font face is always the template font | Use XML editing to adjust individual runs after apply |
| `apply_template.py` hyperlinks | Links are not migrated (text is extracted without URL) | Re-add hyperlinks manually in PowerPoint after apply |
| `beautify_ppt.py` on slides with custom positioned elements | Script may reposition elements to fit the new theme grid | Run visual QA and manually adjust offending slides via XML |
| Very long bullet text | May overflow text boxes after theme change | Shorten content or adjust font size in XML after beautify |
| Embedded fonts in source PPT | Fonts may not transfer to output | Install the same fonts locally, or substitute with theme fonts |
| LibreOffice not installed | `soffice` conversion for image QA will fail | Install via `brew install libreoffice` (macOS) |
