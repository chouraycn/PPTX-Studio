# PPTX Studio — Long-Term Memory

## Project Overview
PPTX Studio Skill (`SKILL.md`) — 全功能 PPTX 处理 Skill，覆盖6种模式：Template Apply / Style Beautify / Editing / Reading / Speaker Notes / Creating from Scratch。脚本目录：`scripts/`（10个 Python 脚本 + `office/` 子目录）。

## SKILL.md Improvements (2026-03-25)
分析并执行了以下四项改进：
1. **动画丢失警告** — Decision Flow 和 Mode 1 Step 0 加入动画检测 + 用户确认流程
2. **`--no-restructure` 对比简表** — `beautify_ppt.py` 脚本说明区加入4行对比表，说明何时需要加该 flag
3. **Node.js 环境检测** — Creating from Scratch 区加入 `node --version` 检查步骤及安装指引
4. **`--language auto` 检测规则说明** — Speaker Notes 常用选项区加入检测逻辑（汉字占比 >10% 判中文）
