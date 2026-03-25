# PPTX Studio Color Replacement Examples

This directory contains example workflows and color maps for the `color_replacement.py` script.

## Quick Start Examples

### 1. Simple Color Replacement
Replace a specific color (e.g., change orange to blue):

```bash
# Preview first
python scripts/color_replacement.py presentation.pptx output.pptx \
    --replace-primary F96167 0284C7 \
    --preview

# Apply changes
python scripts/color_replacement.py presentation.pptx output.pptx \
    --replace-primary F96167 0284C7
```

### 2. AI Color Ladder (Multi-Level Gradients)
Generate 5-level color ladder and automatically replace all colors:

```bash
# Lightness-based ladder (dark to light)
python scripts/color_replacement.py presentation.pptx output.pptx \
    --ai-ladder F96167 \
    --depth 5 \
    --ladder-strategy lightness \
    --verbose

# Saturation-based ladder (dull to vivid)
python scripts/color_replacement.py presentation.pptx output.pptx \
    --ai-ladder 0284C7 \
    --depth 7 \
    --ladder-strategy saturation

# Complementary ladder (base to complementary)
python scripts/color_replacement.py presentation.pptx output.pptx \
    --ai-ladder F96167 \
    --depth 5 \
    --ladder-strategy complementary
```

### 3. Theme-Based Replacement
Replace entire theme (e.g., Warm theme → Tech theme):

```bash
# Preview theme change
python scripts/color_replacement.py presentation.pptx output.pptx \
    --theme-from warm \
    --theme-to tech \
    --preview

# Apply theme change
python scripts/color_replacement.py presentation.pptx output.pptx \
    --theme-from warm \
    --theme-to tech
```

### 4. Custom Color Map
Use a JSON file with custom color mappings:

```bash
python scripts/color_replacement.py presentation.pptx output.pptx \
    --color-map-file examples/color_maps/warm_to_blue.json \
    --verbose
```

## Color Maps

See the `color_maps/` directory for predefined color mapping examples:

- `warm_to_blue.json` - Convert Warm theme to blue-based
- `orange_to_green.json` - Orange to green gradient
- `dark_to_light.json` - Dark theme to light theme

## Color Ladder Strategies

### Lightness Strategy (Default)
Generates a gradient from dark to light:
```
Level 0: Darkest - Text on light backgrounds
Level 1: Darker - Secondary elements
Level 2: Medium - Primary content
Level 3: Lighter - Tertiary elements
Level 4: Lightest - Text on dark backgrounds
```

Best for: Creating depth and hierarchy

### Saturation Strategy
Generates a gradient from dull to vivid:
```
Level 0: Muted (10% saturation)
Level 1: Subtle (32% saturation)
Level 2: Balanced (55% saturation)
Level 3: Vivid (77% saturation)
Level 4: Bold (100% saturation)
```

Best for: Creating emphasis and visual interest

### Complementary Strategy
Generates a gradient crossing to the complementary color:
```
Level 0: Base color (e.g., Orange)
Level 1: Orange-red transition
Level 2: Neutral transition
Level 3: Blue-green transition
Level 4: Complementary color (e.g., Blue)
```

Best for: Creating color harmony and contrast

## Available Themes

| Theme | Primary | Secondary | Accent | Best For |
|-------|---------|-----------|--------|----------|
| executive | Navy | Ice blue | Gold | Business, finance |
| tech | Teal | Dark navy | Mint | Technology, startups |
| creative | Coral | Navy | Gold | Design, proposals |
| warm | Terracotta | Sage | Sand | Education, non-profit |
| minimal | Charcoal | Lavender | White | Academic, clean |
| bold | Red | Navy | Gold | High-impact presentations |
| nature | Forest | Amber | Sky blue | Environmental, health |
| ocean | Deep blue | Cyan | Light blue | Travel, marine |
| elegant | Dark gray | Olive | Coral | Luxury, fashion |
| modern | Purple | Pink | Magenta | Internet, fashion |
| sunset | Orange | Gold | Yellow | Energy, food |
| forest | Green | Emerald | Mint | Sustainability, organic |

## Tips

1. **Always preview first**: Use `--preview` to see what will change
2. **Use verbose mode**: `--verbose` shows detailed color counts
3. **Start with theme**: Use `--theme-from/--theme-to` for complete makeovers
4. **Fine-tune with AI ladder**: Generate custom gradients from your brand color
5. **Backup your files**: Original PPT is not modified unless you specify `--color-map-file`

## Advanced Workflow

```bash
# 1. Analyze current colors
python scripts/color_replacement.py presentation.pptx output.pptx --preview

# 2. Generate AI ladder from brand color
python scripts/color_replacement.py presentation.pptx output.pptx \
    --ai-ladder YOUR_BRAND_COLOR \
    --depth 5 \
    --ladder-strategy lightness \
    --preview

# 3. Apply if satisfied
python scripts/color_replacement.py presentation.pptx output.pptx \
    --ai-ladder YOUR_BRAND_COLOR \
    --depth 5 \
    --ladder-strategy lightness

# 4. Review thumbnails
python scripts/thumbnail.py output.pptx output_thumb
```

## Common Use Cases

### Use Case 1: Brand Color Integration
Your brand color is #0066CC (blue). Apply it to an existing presentation:

```bash
python scripts/color_replacement.py presentation.pptx output.pptx \
    --ai-ladder 0066CC \
    --depth 5 \
    --ladder-strategy lightness
```

### Use Case 2: Dark to Light Theme
Convert a dark-themed presentation to light:

```bash
python scripts/color_replacement.py presentation.pptx output.pptx \
    --theme-from bold \
    --theme-to minimal
```

### Use Case 3: Seasonal Color Update
Update presentation from autumn (orange) to winter (blue):

```bash
python scripts/color_replacement.py presentation.pptx output.pptx \
    --replace-primary F96167 0077B6 \
    --replace-secondary F97316 00B4D8
```

### Use Case 4: Multi-Brand Presentation
Unify different brand presentations to a common theme:

```bash
# Brand A's presentation
python scripts/color_replacement.py brand_a.pptx unified_a.pptx \
    --theme-from creative \
    --theme-to tech

# Brand B's presentation
python scripts/color_replacement.py brand_b.pptx unified_b.pptx \
    --theme-from warm \
    --theme-to tech
```
