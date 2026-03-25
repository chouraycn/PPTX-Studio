# PPTX Studio — Long-Term Memory

## Project Overview
PPTX Studio Skill (`SKILL.md`) — 全功能 PPTX 处理 Skill，覆盖6种模式：Template Apply / Style Beautify / Editing / Reading / Speaker Notes / Creating from Scratch。脚本目录：`scripts/`（10个 Python 脚本 + `office/` 子目录）。

## SKILL.md Improvements (2026-03-25)
分析并执行了以下四项改进：
1. **动画丢失警告** — Decision Flow 和 Mode 1 Step 0 加入动画检测 + 用户确认流程
2. **`--no-restructure` 对比简表** — `beautify_ppt.py` 脚本说明区加入4行对比表，说明何时需要加该 flag
3. **Node.js 环境检测** — Creating from Scratch 区加入 `node --version` 检查步骤及安装指引
4. **`--language auto` 检测规则说明** — Speaker Notes 常用选项区加入检测逻辑（汉字占比 >10% 判中文）

## SKILL.md & Script Fixes (2026-03-25 第二轮)
执行了三项修复：
1. **`patch_slide.py` 注释修正** — 文档注释第9行 "Max 10 replacements" 改为 "Max 20 replacements"，与实际常量 `MAX_REPLACEMENTS_PER_RUN = 20` 一致
2. **Mode 2 Step 3 加 `--no-restructure` 引导** — 在 `beautify_ppt.py` 命令下方加了一行判断提示，用户按流程走时不用翻到脚本说明区
3. **`extract_content.py` 完整字段说明** — JSON 示例后加了"完整字段说明"注释块，列出 `subtitle`、`body_rich`、`image_count`、`has_tables`、`table_data`、`layout_name`、`detected_theme` 七个字段

## SKILL.md & Script Fixes (2026-03-25 第三轮)
执行了三项修复：
1. **`thumbnail.py` 运行时 bug 修复** — 在脚本顶部加入 `sys.path.insert(0, str(Path(__file__).parent))`，修复从项目根目录执行 `python scripts/thumbnail.py` 时 `ModuleNotFoundError: No module named 'office'` 的问题
2. **`add_slide.py` 头注释路径修正** — 将调用示例从 `python add_slide.py` 改为 `python scripts/add_slide.py`，与 SKILL.md 文档保持一致
3. **`extract_content.py` 完整字段说明补全** — 在上轮7字段基础上，新增 `slide_file`、`layout_file`、`shape_count`、`background_color` 四个字段说明；同时移除了脚本实际不输出的 `detected_theme` 错误字段

## SKILL.md & Script Fixes (2026-03-25 第四轮 深度检查)
执行了五项修复：
1. **SKILL.md Minimal 主题字体描述错误** — 文档写的是 Bold 主题的字体（Arial Black+Arial），改为实际代码中的 Calibri+Calibri
2. **SKILL.md `apply_template.py` 选项说明错误** — `--keep-notes` 这个 flag 不存在，实际是 `--no-notes`（取反语义）；文档已修正
3. **SKILL.md `beautify_ppt.py` Flags 列表漏掉 `--dark-mode`** — 补充进去，描述改为 "Force dark background on all slides"
4. **SKILL.md Converting to Images 命令错误** — `python scripts/office/soffice.py` 是辅助模块不是 CLI，改为正确的 `soffice --headless --convert-to pdf`
5. **`beautify_ppt.py` `_update_text_colors` 函数内 noop bug 修复** — 第658行 `xml = re.sub(r'<a:rPr[^>]*/>', xml, xml)` 参数顺序错误（repl 和 string 写反），会把每个 `<a:rPr ... />` 替换成整个 xml 字符串导致文件膨胀和数据污染；已删除该行

## merge_pptx.py --order 自定义排序支持 (2026-03-25)
在 `scripts/merge_pptx.py` 中新增 `--order` 参数，支持按 `<文件序号>:<页码>` token 序列自由交叉排序多个 PPTX 文件的幻灯片。新增 `parse_order_spec()` 函数，`merge()` 函数增加 `order_specs` 参数，分两路：Mode A（--order 自定义排序）和 Mode B（原 --slides 顺序合并）。同步更新 SKILL.md Mode 6 文档：新增"自定义排序合并（--order）⭐"章节，含交错穿插、完全自定义、三文件示例、dry-run 建议等；参数表新增 `--order` 一行；描述 `--slides` 与 `--order` 互斥关系。

## P2-B 功能补全：合并 PPT (2026-03-25)
新增 `scripts/merge_pptx.py`，使用 python-pptx 直接操作合并多个 PPTX 文件。支持 `--slides`（指定每个文件的页范围）、`--ignore-notes`、`--dry-run`。同时更新 SKILL.md：Decision Flow 加 Mode 6 分支、Quick Reference 加一行、Quick decision table 加一行、description 补充合并触发词（合并PPT/合并幻灯片/拼接PPT/merge pptx 等），并新增完整的 "Mode 6: Merge PPT" 文档章节。这解决了 P2-B 问题（description 承诺 combining/splitting 功能但文档和脚本均缺失）。

## SKILL.md 全面深度检查 (2026-03-25 第七轮)
系统性对比所有脚本实现 vs SKILL.md 文档，修复6项问题：
1. **`extract_content.py` JSON 示例补全** — 示例只有少数字段，现已更新为包含所有实际字段（slide_file、subtitle、body_rich、layout_file、shape_count、background_color），删除了下方冗余的"完整字段说明"注释块（改为简要备注）
2. **`apply_template.py` 参数说明补 `--verbose/-v`** — CLI 支持但文档遗漏
3. **`beautify_ppt.py` 参数说明补 `--verbose/-v`** — CLI 支持但文档遗漏
4. **`generate_notes.py` 后端示例补 `--base-url`** — 脚本支持该参数，示例没有体现
5. **Warm 主题字体名修正** — `Palatino` 改为 `Palatino Linotype`（与代码 `header_font: "Palatino Linotype"` 一致）
6. **QA 验证** — 确认所有 qa_check.py 的 10 个 check 名称、所有主题颜色均与代码一致，无其他漏洞

## SKILL.md & Script Fixes (2026-03-25 第六轮 冲突修复)
执行了五项修复（基于系统性冲突分析）：
1. **`thumbnail.py` 头注释路径修正** — Usage/Examples 中的调用示例从 `python thumbnail.py` 改为 `python scripts/thumbnail.py`，与其他脚本风格统一
2. **SKILL.md `extract_content.py` JSON 示例补充 `detected_theme`** — 顶层示例末尾新增 `"detected_theme": "minimal"` 字段，避免用户误以为该字段不存在（`beautify_ppt.py` 依赖此字段做自动检测）
3. **SKILL.md Bold 主题字体修正** — `Impact + Arial` 改为实际代码的 `Arial Black + Arial`
4. **SKILL.md Warm 主题 secondary/accent 颜色互换** — 文档中 secondary 和 accent 写反，修正为与代码一致：`Secondary: #84B59F (sage)`, `Accent: #ECE2D0 (sand)`
5. **`generate_notes.py` `write_notes_to_slide` 段落残留 bug 修复** — 原实现只清空文字不删节点，多次执行后空白 `<a:p>` 堆积；改为先删除 txBody 中所有旧 `<a:p>` 节点，再逐行重写，实现真正覆盖

## SKILL.md & Script Fixes (2026-03-25 第五轮 深度检查)
执行了四项修复：
1. **SKILL.md `extract_content.py` JSON 示例含幽灵字段 `image_paths`** — 脚本实际输出的是 `image_count`，不存在 `image_paths` 字段；JSON 示例已修正
2. **SKILL.md Dependencies 里 `soffice.py` 描述误导性** — 原文"auto-configured via scripts/office/soffice.py"暗示可直接调用；改为说明它是路径辅助模块，应直接调用 `soffice` 命令
3. **`scripts/office/unpack.py` 头注释调用示例缺路径前缀** — `python unpack.py` 改为 `python scripts/office/unpack.py`
4. **`scripts/clean.py` 头注释和运行时错误提示调用示例缺路径前缀** — `python clean.py` 改为 `python scripts/clean.py`（头注释和 `__main__` 里的两处 print 均已修正）
