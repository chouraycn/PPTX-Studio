# AI 智能阶梯全面集成方案

## 目标
将 `color_replacement.py` 中的 AI 色彩阶梯能力整合到 PPTX Studio Skill 的各个模式中，使得在套用模板、美化 PPT、合并 PPT 等场景中都能自动应用智能色彩渐变。

## 集成策略

### 1. `beautify_ppt.py` 集成 AI 阶梯（优先级：高）

**当前状态**：
- 12 个预设主题，每个主题只有固定的 3-4 种颜色（primary, secondary, accent）
- 主题定义在 `THEMES` 字典中，静态配置

**目标状态**：
- 为每个主题添加 **AI 色彩阶梯** 配置
- 用户可选择：使用传统固定配色 vs. 使用 AI 生成的多级渐变色
- 支持 3 种阶梯策略：lightness（明度）、saturation（饱和度）、complementary（互补色）

**实现方案**：

#### 1.1 扩展主题定义

在 `THEMES` 字典中为每个主题添加阶梯配置：

```python
THEMES = {
    "tech": {
        "name": "Tech",
        "primary": "028090",
        "secondary": "1C2541",
        "accent": "02C39A",
        # 新增：AI 色彩阶梯配置
        "ai_ladder": {
            "enabled": True,
            "base_color": "028090",  # 基于主色生成阶梯
            "depth": 5,              # 默认 5 级
            "strategy": "lightness",   # 默认明度渐变
            "generate_on_demand": True  # 运行时动态生成
        },
        # ... 其他原有字段
    },
    # ... 其他主题
}
```

#### 1.2 新增命令行参数

```bash
# 使用传统固定配色（原有行为）
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech

# 启用 AI 色彩阶梯（新增）
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech --ai-ladder

# 自定义阶梯深度和策略
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech \
    --ai-ladder --ladder-depth 7 --ladder-strategy saturation

# 使用品牌色生成阶梯（覆盖主题主色）
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech \
    --ai-ladder --brand-color 0066CC
```

#### 1.3 集成 `color_replacement.py` 模块

在 `beautify_ppt.py` 中导入色彩阶梯生成函数：

```python
# 导入色彩阶梯模块
from color_replacement import generate_ai_ladder, apply_color_ladder

def generate_theme_ladder(theme: dict, depth: int = 5, 
                        strategy: str = "lightness") -> Dict[str, str]:
    """为主题生成 AI 色彩阶梯"""
    base_color = theme.get("primary")
    ladder = generate_ai_ladder(base_color, depth, strategy)
    
    # 将阶梯映射到主题颜色
    theme["ladder"] = ladder
    theme["ladder_level_0"] = ladder[0]  # 最深色
    theme["ladder_level_1"] = ladder[1]  # 较深色
    theme["ladder_level_2"] = ladder[2]  # 中等色
    theme["ladder_level_3"] = ladder[3]  # 较浅色
    theme["ladder_level_4"] = ladder[4]  # 最浅色
    
    return theme
```

#### 1.4 应用阶梯到幻灯片

在 `_apply_theme_to_slide()` 函数中使用阶梯色：

```python
def _apply_theme_to_slide(xml: str, theme: dict, use_dark: bool,
                        ladder_enabled: bool = False) -> str:
    """应用主题到单个幻灯片"""
    
    if ladder_enabled and "ladder" in theme:
        # 使用 AI 阶梯色
        primary = theme["ladder_level_2"]  # 中等色作为主色
        secondary = theme["ladder_level_1"]  # 较深色作为次色
        accent = theme["ladder_level_3"]     # 较浅色作为强调色
    else:
        # 使用传统固定配色
        primary = theme["primary"]
        secondary = theme["secondary"]
        accent = theme["accent"]
    
    # 应用颜色到形状、文字、背景
    xml = _update_text_colors(xml, theme, use_dark, ladder_enabled)
    xml = _update_shape_colors(xml, theme, use_dark, ladder_enabled)
    # ...
```

#### 1.5 版式变体适配

更新 10 个版式变体函数，支持阶梯色：

```python
def _add_two_tone_panel(xml: str, primary: str, use_dark: bool, 
                       theme: dict, ladder_enabled: bool = False) -> str:
    """添加双色面板（支持阶梯色）"""
    
    if ladder_enabled:
        # 使用阶梯色中的两个色阶
        dark_color = theme["ladder_level_1"]
        light_color = theme["ladder_level_3"]
    else:
        dark_color = primary
        light_color = theme["bg_light"]
    
    # 生成双色面板 XML
    # ...
```

---

### 2. `apply_template.py` 集成 AI 阶梯（优先级：中）

**当前状态**：
- 套用模板时，仅迁移源 PPT 的内容到模板
- 使用模板的固定配色，不提供换色选项

**目标状态**：
- 套用模板后，可选应用 AI 色彩阶梯
- 支持基于品牌色自定义阶梯

**实现方案**：

#### 2.1 新增命令行参数

```bash
# 基础套用模板（原有行为）
python scripts/apply_template.py source.pptx template.pptx output.pptx

# 套用模板后应用 AI 阶梯（新增）
python scripts/apply_template.py source.pptx template.pptx output.pptx \
    --ai-ladder --base-color 0066CC

# 使用主题预设阶梯
python scripts/apply_template.py source.pptx template.pptx output.pptx \
    --ai-ladder --ladder-theme tech
```

#### 2.2 后处理流程

在套用模板完成后，调用 `color_replacement.py`：

```python
def apply_template(source: str, template: str, output: str,
                ai_ladder: bool = False, base_color: str = None,
                ladder_theme: str = None, **kwargs):
    """应用模板到源 PPT"""
    
    # 原有逻辑：提取内容、注入模板
    # ...
    
    # 新增：AI 色彩阶梯后处理
    if ai_ladder:
        temp_output = output
        
        if ladder_theme:
            # 使用主题预设阶梯
            temp_output = apply_theme_ladder(output, ladder_theme, temp_output)
        elif base_color:
            # 使用品牌色生成阶梯
            temp_output = apply_brand_ladder(output, base_color, temp_output)
        else:
            # 自动从模板提取主色
            primary = extract_primary_color(template)
            temp_output = apply_brand_ladder(output, primary, temp_output)
        
        # 替换原始输出
        shutil.move(temp_output, output)
```

---

### 3. `merge_pptx.py` 集成 AI 阶梯（优先级：中）

**当前状态**：
- 合并多个 PPT 文件
- 不处理配色统一

**目标状态**：
- 合并后统一配色
- 支持 AI 色彩阶梯统一多个 PPT

**实现方案**：

#### 3.1 新增命令行参数

```bash
# 基础合并（原有行为）
python scripts/merge_pptx.py file1.pptx file2.pptx output.pptx

# 合并后应用 AI 阶梯统一配色（新增）
python scripts/merge_pptx.py file1.pptx file2.pptx output.pptx \
    --ai-ladder --base-color 0284C7

# 合并后切换到主题阶梯
python scripts/merge_pptx.py file1.pptx file2.pptx output.pptx \
    --ai-ladder --ladder-theme tech
```

#### 3.2 后处理流程

```python
def merge(file_list: List[str], output: str,
          ai_ladder: bool = False, base_color: str = None,
          ladder_theme: str = None, **kwargs):
    """合并多个 PPT 文件"""
    
    # 原有逻辑：合并幻灯片
    # ...
    
    # 新增：AI 色彩阶梯统一配色
    if ai_ladder:
        if ladder_theme:
            output = apply_theme_ladder(output, ladder_theme, output)
        elif base_color:
            output = apply_brand_ladder(output, base_color, output)
        else:
            # 自动检测主色
            primary = auto_detect_primary(output)
            output = apply_brand_ladder(output, primary, output)
```

---

### 4. 创建统一 API 模块（优先级：高）

为了避免代码重复，创建统一的色彩阶梯 API 模块：

#### 4.1 新建 `scripts/color_ladder.py`

```python
"""
Color Ladder API - 统一的 AI 色彩阶梯接口

提供以下功能：
1. 生成色彩阶梯（基于 color_replacement.py）
2. 应用阶梯到 PPT（调用 color_replacement.py）
3. 主题预设阶梯映射
4. 品牌色阶梯生成
"""

from typing import Dict, Optional
from color_replacement import generate_ai_ladder, apply_color_ladder

# 主题预设阶梯配置
THEME_LADDERS = {
    "tech": {
        "base_color": "028090",
        "depth": 5,
        "strategy": "lightness"
    },
    "executive": {
        "base_color": "1E2761",
        "depth": 5,
        "strategy": "lightness"
    },
    # ... 其他主题
}

def get_theme_ladder(theme_name: str, depth: Optional[int] = None,
                   strategy: Optional[str] = None) -> Dict[str, str]:
    """获取主题预设阶梯"""
    config = THEME_LADDERS.get(theme_name, THEME_LADDERS["minimal"])
    return generate_ai_ladder(
        base_color=config["base_color"],
        depth=depth or config["depth"],
        strategy=strategy or config["strategy"]
    )

def apply_theme_ladder(pptx_path: str, theme_name: str, 
                      output_path: str, **kwargs) -> str:
    """应用主题阶梯到 PPT"""
    ladder = get_theme_ladder(theme_name)
    return apply_color_ladder(pptx_path, ladder, output_path, **kwargs)

def apply_brand_ladder(pptx_path: str, brand_color: str,
                       output_path: str, depth: int = 5,
                       strategy: str = "lightness", **kwargs) -> str:
    """应用品牌色阶梯到 PPT"""
    ladder = generate_ai_ladder(brand_color, depth, strategy)
    return apply_color_ladder(pptx_path, ladder, output_path, **kwargs)
```

#### 4.2 更新 `beautify_ppt.py`、`apply_template.py`、`merge_pptx.py`

统一使用 `color_ladder.py` 接口：

```python
# beautify_ppt.py
from color_ladder import get_theme_ladder, apply_theme_ladder

# apply_template.py
from color_ladder import apply_theme_ladder, apply_brand_ladder

# merge_pptx.py
from color_ladder import apply_theme_ladder, apply_brand_ladder
```

---

## 实施计划

### Phase 1: 核心模块开发（2-3 小时）

1. **创建 `color_ladder.py`**（1 小时）
   - 实现 `get_theme_ladder()`、`apply_theme_ladder()`、`apply_brand_ladder()`
   - 为 12 个主题添加阶梯配置

2. **扩展 `beautify_ppt.py`**（1.5 小时）
   - 添加 `--ai-ladder`、`--ladder-depth`、`--ladder-strategy` 参数
   - 修改主题定义，添加阶梯配置
   - 更新 10 个版式变体函数支持阶梯色

3. **单元测试**（0.5 小时）
   - 测试 `color_ladder.py` 的阶梯生成
   - 测试 `beautify_ppt.py` 的阶梯应用

### Phase 2: 模式集成（2 小时）

1. **更新 `apply_template.py`**（1 小时）
   - 添加 `--ai-ladder`、`--base-color`、`--ladder-theme` 参数
   - 实现后处理流程

2. **更新 `merge_pptx.py`**（1 小时）
   - 添加 `--ai-ladder`、`--base-color`、`--ladder-theme` 参数
   - 实现后处理流程

### Phase 3: 文档更新（1 小时）

1. **更新 SKILL.md**
   - 更新 Mode 1、Mode 2、Mode 6 的命令示例
   - 添加 AI 阶梯使用说明

2. **创建示例文档**
   - `examples/ai_ladder_beautify.md`
   - `examples/ai_ladder_template.md`
   - `examples/ai_ladder_merge.md`

### Phase 4: 测试验证（1 小时）

1. **集成测试**
   - 端到端测试 Mode 1 + AI 阶梯
   - 端到端测试 Mode 2 + AI 阶梯
   - 端到端测试 Mode 6 + AI 阶梯

2. **性能测试**
   - 测量阶梯生成时间
   - 测量 PPT 处理时间

---

## 预期效果

### 用户体验

1. **Mode 2: Style Beautify**
```bash
# 原有：使用固定配色
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech

# 新增：使用 AI 阶梯（更丰富的色彩层次）
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech --ai-ladder

# 自定义品牌色阶梯
python scripts/beautify_ppt.py source.pptx output.pptx --theme tech \
    --ai-ladder --brand-color 0066CC --ladder-depth 7
```

2. **Mode 1: Template Apply**
```bash
# 原有：套用模板
python scripts/apply_template.py source.pptx template.pptx output.pptx

# 新增：套用模板 + AI 阶梯
python scripts/apply_template.py source.pptx template.pptx output.pptx \
    --ai-ladder --base-color 0066CC
```

3. **Mode 6: Merge PPT**
```bash
# 原有：合并 PPT
python scripts/merge_pptx.py file1.pptx file2.pptx output.pptx

# 新增：合并 + 统一配色（AI 阶梯）
python scripts/merge_pptx.py file1.pptx file2.pptx output.pptx \
    --ai-ladder --ladder-theme tech
```

### 技术收益

1. **代码复用**：统一 API 模块，避免重复代码
2. **扩展性强**：轻松添加新的阶梯策略或主题
3. **性能优化**：阶梯生成和应用的独立优化
4. **易维护性**：清晰的模块边界和接口

---

## 文件清单

### 新增文件

1. **scripts/color_ladder.py** - 统一的色彩阶梯 API 模块（~300 行）

### 修改文件

1. **scripts/beautify_ppt.py**
   - 扩展主题定义
   - 添加 AI 阶梯参数
   - 更新版式变体函数

2. **scripts/apply_template.py**
   - 添加 AI 阶梯参数
   - 实现后处理流程

3. **scripts/merge_pptx.py**
   - 添加 AI 阶梯参数
   - 实现后处理流程

4. **SKILL.md**
   - 更新 Mode 1、Mode 2、Mode 6 文档

5. **examples/** - 新增示例文档

---

## 风险与缓解

### 风险 1：向后兼容性

**风险**：新参数可能影响现有用户工作流

**缓解**：
- `--ai-ladder` 默认为 `False`
- 不指定时保持原有行为
- 文档明确标注新增功能

### 风险 2：性能影响

**风险**：阶梯生成和替换可能增加处理时间

**缓解**：
- 阶梯生成预缓存（主题预设）
- 正则替换优化
- 提供性能基准测试

### 风险 3：色彩协调性

**风险**：AI 生成的阶梯可能与内容不协调

**缓解**：
- 提供 `--preview` 模式
- 优化默认阶梯策略（lightness）
- 支持自定义阶梯配置
