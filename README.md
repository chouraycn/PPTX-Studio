# PPTX Studio

一个功能强大的PPT处理工具集，支持PPTX文件的创建、编辑、模板套用、风格美化和演讲者备注生成。

📦 **Source Code**: https://github.com/chouraycn/PPTX-Studio

## 项目简介

PPTX Studio 是一个专门用于处理Microsoft PowerPoint (.pptx)文件的工具集合。它提供了多种功能，包括：
- 从PPTX文件中提取和分析内容
- 套用模板样式
- 美化PPT风格
- 自动生成演讲者备注
- 创建和编辑PPT文件

## 功能特性

### 1. 内容提取与分析
- 从PPTX文件中提取结构化内容（JSON格式）
- 自动识别幻灯片类型（标题、章节、内容、图表、表格等）
- 提取关键词，推断主题风格

### 2. 模板套用
- 将源PPT内容套入指定模板的视觉风格
- 自动映射幻灯片类型到模板布局
- 支持自定义映射规则

### 3. 风格美化
- 提供8种预定义主题风格：
  - **executive** - 商务风格：深蓝色/金色配色
  - **tech** - 科技风格：青绿色/绿色配色
  - **creative** - 创意风格：珊瑚色/黄色配色
  - **warm** - 温暖风格：砖红色/米色配色
  - **minimal** - 极简风格：炭灰色/深灰色配色
  - **bold** - 大胆风格：深红色/浅色配色
  - **nature** - 自然风格：深绿色/浅灰色配色
  - **ocean** - 海洋风格：深蓝色/淡绿色配色

### 4. 演讲者备注生成
- 三种备注风格：`speaker`（演讲提示）、`coach`（演讲教练）、`summary`（摘要）
- 三种生成后端：`simple`（规则驱动，无需API）、`openai`、`ollama`
- 支持中英文自动检测
- 支持dry-run预览和保留现有备注

### 5. PPT编辑与创建
- 基于XML的PPT编辑（解包→编辑→打包）
- 使用PptxGenJS从头创建PPT
- 支持添加、删除、重新排序幻灯片

## 快速开始

### 安装要求
- Python 3.9.6+
- 依赖包：可通过运行脚本自动安装

### 基本用法

#### 提取PPT内容
```bash
python3 scripts/extract_content.py source.pptx --print-summary
```

#### 套用模板
```bash
python3 scripts/apply_template.py source.pptx template.pptx output.pptx
```

#### 美化风格
```bash
# 查看可用主题
python3 scripts/beautify_ppt.py dummy dummy --list-themes

# 应用特定主题
python3 scripts/beautify_ppt.py source.pptx output.pptx --theme tech
```

#### 生成演讲者备注
```bash
# 最简单用法（规则生成，无需API）
python3 scripts/generate_notes.py source.pptx output.pptx

# 预览备注效果（不写入）
python3 scripts/generate_notes.py source.pptx output.pptx --dry-run

# 跳过已有备注的幻灯片
python3 scripts/generate_notes.py source.pptx output.pptx --no-overwrite
```

#### 解包/打包PPT进行编辑
```bash
# 解包PPT
python3 scripts/office/unpack.py source.pptx unpacked/

# 编辑解包后的XML文件
# ...

# 打包回PPT
python3 scripts/office/pack.py unpacked/ output.pptx --original source.pptx
```

## 详细使用指南

### 1. 内容提取与分析
`extract_content.py`脚本可以从PPTX文件中提取结构化信息，包括：
- 幻灯片标题和内容
- 幻灯片类型（标题页、章节页、内容页等）
- 关键词提取
- 主题风格推断

```bash
# 提取内容并保存为JSON
python3 scripts/extract_content.py presentation.pptx -o content.json

# 打印人类可读的摘要
python3 scripts/extract_content.py presentation.pptx --print-summary
```

### 2. 模板套用工作流
当您有一个内容PPT和一个样式PPT时，可以使用`apply_template.py`将内容套用到模板样式中：

1. 首先分析两个文件：
```bash
# 查看源内容
python3 scripts/extract_content.py source.pptx --print-summary

# 查看模板布局
python3 scripts/thumbnail.py template.pptx
```

2. 应用模板：
```bash
python3 scripts/apply_template.py source.pptx template.pptx output.pptx
```

3. 可选：使用自定义映射文件：
```bash
python3 scripts/apply_template.py source.pptx template.pptx output.pptx --mapping mapping.json
```

### 3. 风格美化选项
`beautify_ppt.py`提供了8种预定义主题：

| 主题 | 配色方案 | 字体 | 适用场景 |
|------|---------|------|---------|
| executive | #1E2761 / #C9A84C | Cambria + Calibri | 商务报告、正式会议 |
| tech | #028090 / #02C39A | Trebuchet MS + Calibri | 技术演示、产品发布 |
| creative | #F96167 / #F9E795 | Georgia + Calibri | 创意展示、设计提案 |
| warm | #B85042 / #ECE2D0 | Palatino Linotype + Calibri | 品牌故事、人文主题 |
| minimal | #36454F / #212121 | Calibri + Calibri | 简约设计、数据报告 |
| bold | #990011 / #FCF6F5 | Arial Black + Arial | 营销材料、引人注目的演讲 |
| nature | #2C5F2D / #F5F5F5 | Georgia + Calibri | 环保主题、可持续发展 |
| ocean | #065A82 / #9FFFCB | Calibri + Calibri | 海洋相关、清新风格 |

### 4. 演讲者备注生成
`generate_notes.py`支持多种生成模式：

#### 备注模式
- **speaker**（默认）：提供演讲提示，帮助演讲者流畅表达
- **coach**：提供演讲教练建议，包括语气、节奏和肢体语言
- **summary**：生成幻灯片内容摘要

#### 生成后端
- **simple**：基于规则的生成，无需API调用
- **openai**：使用OpenAI API生成
- **ollama**：使用本地Ollama模型生成

#### 使用示例
```bash
# 使用OpenAI生成
python3 scripts/generate_notes.py deck.pptx out.pptx --backend openai --api-key sk-xxx

# 使用本地Ollama
python3 scripts/generate_notes.py deck.pptx out.pptx --backend ollama --model llama3

# 生成中文备注
python3 scripts/generate_notes.py deck.pptx out.pptx --language zh

# 生成教练风格的备注
python3 scripts/generate_notes.py deck.pptx out.pptx --mode coach
```

### 5. PPT编辑工作流
对于需要深度编辑的场景，可以使用解包/打包工作流：

1. 解包PPT：
```bash
python3 scripts/office/unpack.py input.pptx unpacked/
```

2. 编辑XML文件：
   - 幻灯片内容：`unpacked/ppt/slides/slideN.xml`
   - 演示文稿结构：`unpacked/ppt/presentation.xml`
   - 使用`scripts/add_slide.py`添加幻灯片
   - 使用`scripts/clean.py`清理临时文件

3. 打包回PPT：
```bash
python3 scripts/office/pack.py unpacked/ output.pptx --original input.pptx
```

## 技术说明

### Python版本兼容性
- 系统Python版本：3.9.6
- 使用`Optional[str]`而非`str | None`（Python 3.10+语法）
- `office/pack.py`和`office/unpack.py`需要Python 3.10+，因此新脚本通过subprocess调用

### 文件结构
```
PPTX Studio/
├── scripts/                    # 主要脚本文件
│   ├── extract_content.py     # 提取PPT内容
│   ├── apply_template.py      # 套用模板
│   ├── beautify_ppt.py        # 美化风格
│   ├── generate_notes.py      # 生成演讲者备注
│   ├── thumbnail.py           # 生成缩略图
│   ├── add_slide.py           # 添加幻灯片
│   ├── clean.py               # 清理临时文件
│   └── office/                # PPT解包/打包工具
├── SKILL.md                   # 完整技能文档
├── editing.md                 # 编辑指南
├── pptxgenjs.md               # PptxGenJS使用指南
├── LICENSE.txt               # 许可证文件
└── README.md                 # 本文件
```

### 脚本参数说明
所有脚本都支持`--help`参数查看详细使用说明：
```bash
python3 scripts/<script_name>.py --help
```

## 常见问题

### Q: 如何选择合适的模板或主题？
A: 根据您的演讲内容和受众选择：
- **正式场合**：使用`executive`或`minimal`主题
- **技术演示**：使用`tech`主题
- **创意展示**：使用`creative`或`warm`主题
- **营销材料**：使用`bold`主题

### Q: 生成的备注质量如何？
A: 备注质量取决于：
1. **simple模式**：基于规则生成，适合简单的演讲提示
2. **openai/ollama模式**：使用AI生成，质量更高，但需要API访问

### Q: 如何处理中英文混合的PPT？
A: 使用`--language auto`参数，脚本会自动检测每页幻灯片的主要语言并生成相应语言的备注。

### Q: 如何确保编辑后的PPT格式正确？
A: 建议：
1. 使用`scripts/clean.py`清理解包后的文件
2. 使用`--original`参数保持与原文件的兼容性
3. 在打包前检查XML文件的结构

## 贡献

## 贡献

欢迎提交问题和拉取请求。请确保：
1. 遵循现有的代码风格
2. 添加适当的测试
3. 更新相关文档

## 联系

如有问题或建议，请通过项目仓库的问题跟踪器提交。

---

**注意**：本项目仍在积极开发中，功能可能会有所变化。建议定期查看更新。