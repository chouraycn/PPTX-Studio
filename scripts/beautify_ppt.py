"""Beautify a PPT by redesigning its visual style while preserving content.

Analyzes the source presentation and applies a professional theme with
enhanced visual features including 12 themes, 10 layout variants,
smart image enhancement, and auto icon insertion.

Features:
- 12 professional themes with cohesive color palettes
- 10 layout variants (auto-rotated for visual variety)
- Smart image enhancement (rounded corners, shadows, borders)
- Dynamic font sizing based on content density
- Gradient backgrounds for title slides
- Smart icon insertion based on keywords
- Paragraph spacing optimization

Usage:
    python scripts/beautify_ppt.py source.pptx output.pptx
    python scripts/beautify_ppt.py source.pptx output.pptx --theme tech
    python scripts/beautify_ppt.py source.pptx output.pptx --theme elegant --gradient-bg
    python scripts/beautify_ppt.py source.pptx output.pptx --theme modern --smart-icons
    python scripts/beautify_ppt.py source.pptx output.pptx --theme sunset --gradient-bg --smart-icons
    python scripts/beautify_ppt.py source.pptx output.pptx --no-restructure  # skip layout changes

Available themes: executive, tech, creative, warm, minimal, bold, nature, ocean,
                  elegant, modern, sunset, forest

Layout variants: accent_bar, numbered_list, stat_highlight, two_tone, header_band,
                 card_grid, timeline, split_diagonal, image_focus, quote_block
"""

import argparse
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Optional, List, Dict

import subprocess

# Add scripts dir to path for extract_content (Python 3.9 compatible)
sys.path.insert(0, str(Path(__file__).parent))

from extract_content import extract_content, _detect_theme


def _run_unpack(input_file: str, output_dir: str) -> None:
    """Run unpack.py as a subprocess (handles Python version compatibility)."""
    scripts_dir = Path(__file__).parent
    result = subprocess.run(
        [sys.executable, str(scripts_dir / "office" / "unpack.py"), input_file, output_dir],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"Unpack error: {result.stderr}", file=sys.stderr)
        sys.exit(1)
    print(f"  {result.stdout.strip()}")


def _run_pack(input_dir: str, output_file: str, original: Optional[str] = None) -> None:
    """Run pack.py as a subprocess (handles Python version compatibility)."""
    scripts_dir = Path(__file__).parent
    cmd = [sys.executable, str(scripts_dir / "office" / "pack.py"),
           input_dir, output_file, "--validate", "false"]
    if original:
        cmd += ["--original", original]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Pack error: {result.stderr}", file=sys.stderr)
        sys.exit(1)
    print(f"  {result.stdout.strip()}")


# ─────────────────────────────────────────────────────────────────────────────
# THEME DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────

THEMES = {
    "executive": {
        "name": "Executive",
        "primary": "1E2761",      # Navy
        "secondary": "CADCFC",    # Ice blue
        "accent": "C9A84C",       # Gold
        "bg_light": "FFFFFF",     # White
        "bg_dark": "1E2761",      # Navy
        "text_on_dark": "FFFFFF",
        "text_on_light": "1E2761",
        "text_muted": "6B7280",
        "header_font": "Cambria",
        "body_font": "Calibri",
        "title_bold": True,
        "title_size": 4000,       # 40pt in hundredths
        "body_size": 1800,        # 18pt
        "caption_size": 1200,     # 12pt
        "gradient_start": "1E2761",
        "gradient_end": "2A3A7C",
    },
    "tech": {
        "name": "Tech",
        "primary": "028090",      # Teal
        "secondary": "1C2541",    # Dark navy
        "accent": "02C39A",       # Mint
        "bg_light": "F0FFFE",     # Near white with teal tint
        "bg_dark": "0B0C10",      # Very dark
        "text_on_dark": "FFFFFF",
        "text_on_light": "1C2541",
        "text_muted": "4B5563",
        "header_font": "Trebuchet MS",
        "body_font": "Calibri",
        "title_bold": True,
        "title_size": 3800,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "028090",
        "gradient_end": "02C39A",
    },
    "creative": {
        "name": "Creative",
        "primary": "F96167",      # Coral
        "secondary": "2F3C7E",    # Navy
        "accent": "F9E795",       # Gold
        "bg_light": "FFFDF9",     # Warm white
        "bg_dark": "2F3C7E",      # Navy
        "text_on_dark": "FFFFFF",
        "text_on_light": "2F3C7E",
        "text_muted": "6B7280",
        "header_font": "Georgia",
        "body_font": "Calibri",
        "title_bold": True,
        "title_size": 4000,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "F96167",
        "gradient_end": "F9E795",
    },
    "warm": {
        "name": "Warm",
        "primary": "B85042",      # Terracotta
        "secondary": "84B59F",    # Sage
        "accent": "ECE2D0",       # Sand
        "bg_light": "FFFDF9",     # Cream
        "bg_dark": "B85042",      # Terracotta
        "text_on_dark": "FFFFFF",
        "text_on_light": "3D2B1F",
        "text_muted": "78716C",
        "header_font": "Palatino Linotype",
        "body_font": "Calibri",
        "title_bold": True,
        "title_size": 3800,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "B85042",
        "gradient_end": "D4685A",
    },
    "minimal": {
        "name": "Minimal",
        "primary": "36454F",      # Charcoal
        "secondary": "F2F2F2",    # Off-white
        "accent": "212121",       # Near black
        "bg_light": "FFFFFF",
        "bg_dark": "36454F",      # Charcoal
        "text_on_dark": "FFFFFF",
        "text_on_light": "36454F",
        "text_muted": "9CA3AF",
        "header_font": "Calibri",
        "body_font": "Calibri",
        "title_bold": True,
        "title_size": 4000,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "36454F",
        "gradient_end": "4A5A65",
    },
    "bold": {
        "name": "Bold",
        "primary": "990011",      # Cherry
        "secondary": "2F3C7E",    # Navy
        "accent": "FCF6F5",       # Near white
        "bg_light": "FFFFFF",
        "bg_dark": "1A1A2E",      # Very dark navy
        "text_on_dark": "FFFFFF",
        "text_on_light": "1A1A2E",
        "text_muted": "6B7280",
        "header_font": "Arial Black",
        "body_font": "Arial",
        "title_bold": True,
        "title_size": 4400,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "990011",
        "gradient_end": "B81A2C",
    },
    "nature": {
        "name": "Nature",
        "primary": "2C5F2D",      # Forest
        "secondary": "97BC62",    # Moss
        "accent": "F5F5F5",       # Cream
        "bg_light": "FAFFF5",     # Very light green
        "bg_dark": "2C5F2D",      # Forest
        "text_on_dark": "FFFFFF",
        "text_on_light": "1A2E1B",
        "text_muted": "6B7280",
        "header_font": "Georgia",
        "body_font": "Calibri",
        "title_bold": True,
        "title_size": 4000,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "2C5F2D",
        "gradient_end": "4A7F4B",
    },
    "ocean": {
        "name": "Ocean",
        "primary": "065A82",      # Deep blue
        "secondary": "1C7293",    # Teal
        "accent": "9FFFCB",       # Mint
        "bg_light": "F0F8FF",     # Alice blue
        "bg_dark": "02364A",      # Dark ocean
        "text_on_dark": "FFFFFF",
        "text_on_light": "02364A",
        "text_muted": "64748B",
        "header_font": "Calibri",
        "body_font": "Calibri",
        "title_bold": True,
        "title_size": 4000,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "065A82",
        "gradient_end": "1C7293",
    },
    # ═══════════════════════════════════════════════════════════════════════
    # NEW THEMES - Enhanced Visual Appeal
    # ═══════════════════════════════════════════════════════════════════════
    "elegant": {
        "name": "Elegant",
        "primary": "2C3E50",      # Deep slate blue
        "secondary": "E8E8E8",    # Light silver
        "accent": "E74C3C",       # Coral red
        "bg_light": "FAFAFA",     # Soft white
        "bg_dark": "1A1A2E",      # Midnight
        "text_on_dark": "FFFFFF",
        "text_on_light": "2C3E50",
        "text_muted": "7F8C8D",
        "header_font": "Georgia",
        "body_font": "Calibri",
        "title_bold": True,
        "title_size": 4200,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "2C3E50",
        "gradient_end": "34495E",
    },
    "modern": {
        "name": "Modern",
        "primary": "6C5CE7",      # Soft purple
        "secondary": "A29BFE",    # Light lavender
        "accent": "FD79A8",       # Pink
        "bg_light": "F8F9FA",     # Very light gray
        "bg_dark": "2D3436",      # Dark charcoal
        "text_on_dark": "FFFFFF",
        "text_on_light": "2D3436",
        "text_muted": "636E72",
        "header_font": "Segoe UI",
        "body_font": "Segoe UI",
        "title_bold": True,
        "title_size": 4000,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "6C5CE7",
        "gradient_end": "A29BFE",
    },
    "sunset": {
        "name": "Sunset",
        "primary": "E17055",      # Burnt orange
        "secondary": "FDCB6E",    # Warm yellow
        "accent": "D63031",       # Deep red
        "bg_light": "FFF9F0",     # Creamy white
        "bg_dark": "2D142C",      # Deep plum
        "text_on_dark": "FFFFFF",
        "text_on_light": "2D142C",
        "text_muted": "8B7355",
        "header_font": "Georgia",
        "body_font": "Calibri",
        "title_bold": True,
        "title_size": 4000,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "E17055",
        "gradient_end": "FDCB6E",
    },
    "forest": {
        "name": "Forest",
        "primary": "1B4332",      # Deep forest
        "secondary": "52B788",    # Sage green
        "accent": "D8F3DC",       # Pale mint
        "bg_light": "F1F8E9",     # Very pale green
        "bg_dark": "081C15",      # Deep jungle
        "text_on_dark": "FFFFFF",
        "text_on_light": "1B4332",
        "text_muted": "52796F",
        "header_font": "Cambria",
        "body_font": "Calibri",
        "title_bold": True,
        "title_size": 4000,
        "body_size": 1800,
        "caption_size": 1200,
        "gradient_start": "1B4332",
        "gradient_end": "2D6A4F",
    },
}

# Which slide types get dark backgrounds (title slides, section headers, conclusions)
DARK_BG_TYPES = {"title", "section", "conclusion"}


def beautify_ppt(
    source_pptx: str,
    output_pptx: str,
    theme_name: Optional[str] = None,
    dark_mode: bool = False,
    keep_images: bool = True,
    font_pair: Optional[str] = None,
    restructure: bool = True,
    verbose: bool = False,
    enhance_images: bool = True,
    use_gradient: bool = False,
    smart_icons: bool = False,
    ai_ladder: bool = False,
    ladder_depth: int = 5,
    ladder_strategy: str = "lightness",
    brand_color: Optional[str] = None,
) -> None:
    """Redesign a PPT's visual style while preserving content.
    
    Args:
        ai_ladder: Enable AI color ladder (overrides theme colors with gradients)
        ladder_depth: Number of ladder levels (3-10)
        ladder_strategy: Gradient strategy ("lightness"/"saturation"/"complementary")
        brand_color: Custom brand color for ladder generation (overrides theme primary)
    """

    source_path = Path(source_pptx)
    output_path = Path(output_pptx)

    if not source_path.exists():
        print(f"Error: Source file not found: {source_pptx}", file=sys.stderr)
        sys.exit(1)

    print(f"Analyzing source presentation: {source_pptx}")
    content = extract_content(str(source_path), print_summary=verbose)

    if not content["slides"]:
        print("Error: No slides found in source file", file=sys.stderr)
        sys.exit(1)

    # Select theme
    if not theme_name:
        theme_name = content.get("detected_theme", "minimal")
        print(f"Auto-detected theme: {theme_name}")
    else:
        print(f"Using theme: {theme_name}")

    if theme_name not in THEMES:
        print(f"Warning: Unknown theme '{theme_name}', falling back to 'minimal'")
        theme_name = "minimal"

    theme = THEMES[theme_name]

    # AI Ladder Integration
    ladder_enabled = False
    if ai_ladder:
        print("\n🎨 AI Color Ladder Enabled")
        ladder_enabled = True
        
        # Import color_ladder module
        try:
            from color_ladder import get_theme_ladder, apply_brand_ladder
            
            if brand_color:
                print(f"  Using brand color: #{brand_color}")
                ladder_dict = apply_brand_ladder(
                    str(source_path),
                    brand_color,
                    str(tmp_path / "temp_ladder.pptx"),
                    depth=ladder_depth,
                    strategy=ladder_strategy,
                    preview=False,
                    verbose=verbose
                )
                # Extract ladder colors from generated file
                ladder = get_theme_ladder(theme_name, ladder_depth, ladder_strategy)
            else:
                print(f"  Generating ladder for theme: {theme_name}")
                ladder = get_theme_ladder(theme_name, ladder_depth, ladder_strategy)
            
            # Update theme with ladder colors
            theme["ladder"] = ladder
            theme["ladder_enabled"] = True
            print(f"  Ladder depth: {ladder_depth} levels")
            print(f"  Ladder strategy: {ladder_strategy}")
            print(f"  Level 0 (darkest): #{ladder['level_0']}")
            print(f"  Level 2 (middle): #{ladder['level_2']}")
            print(f"  Level 4 (lightest): #{ladder['level_4']}")
        except ImportError as e:
            print(f"Warning: Could not import color_ladder module: {e}")
            print("Falling back to standard theme colors")
            ladder_enabled = False

    # Override font pair if specified
    if font_pair:
        parts = font_pair.split("-")
        if len(parts) >= 2:
            theme = dict(theme)
            theme["header_font"] = parts[0].replace("_", " ").title()
            theme["body_font"] = parts[1].replace("_", " ").title()

    print(f"Theme: {theme['name']}")
    if ladder_enabled:
        print(f"Colors: AI Ladder (Level 0-4)")
    else:
        print(f"Colors: Primary #{theme['primary']}, Accent #{theme['accent']}")
    print(f"Fonts: {theme['header_font']} / {theme['body_font']}")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        unpacked_dir = tmp_path / "working"

        print(f"\nUnpacking presentation...")
        _run_unpack(str(source_path), str(unpacked_dir))

        slides_dir = unpacked_dir / "ppt" / "slides"
        slide_files = sorted(
            slides_dir.glob("slide*.xml"),
            key=lambda f: int(re.match(r"slide(\d+)\.xml", f.name).group(1))
        )

        print(f"\nApplying theme to {len(slide_files)} slides...")

        # Track layout usage for monotony prevention
        layout_streak: List[str] = []

        # Process each slide
        for i, slide_path in enumerate(slide_files):
            # Find matching content data
            slide_content = None
            for s in content["slides"]:
                if s.get("slide_file") == slide_path.name:
                    slide_content = s
                    break

            slide_type = slide_content.get("type", "content") if slide_content else "content"
            body_items = slide_content.get("body", []) if slide_content else []
            use_dark = (slide_type in DARK_BG_TYPES) or dark_mode

            if verbose:
                print(f"  Processing {slide_path.name} [{slide_type}] "
                      f"({'dark' if use_dark else 'light'} bg)")

            # Decide layout variant to avoid monotony
            layout_variant = _pick_layout_variant(
                slide_type, body_items, layout_streak, i, theme
            )
            layout_streak.append(layout_variant)

            # Smart icon insertion based on content
            if smart_icons and body_items:
                body_items = _insert_smart_icons(body_items, theme)

            _beautify_slide(
                slide_path, theme, use_dark, slide_type, dark_mode,
                layout_variant=layout_variant,
                body_items=body_items,
                restructure=restructure,
                verbose=verbose,
                enhance_images=enhance_images,
                use_gradient=use_gradient,
            )

        # Apply theme to slide master/layouts
        _apply_theme_to_master(unpacked_dir, theme, verbose)

        # Fix theme XML
        _update_theme_xml(unpacked_dir, theme, verbose)

        print(f"\nPacking output to {output_pptx}...")
        _run_pack(str(unpacked_dir), str(output_path))

    print(f"\nDone! Beautified presentation saved to: {output_pptx}")
    print(f"Run QA check with:")
    print(f"  python scripts/qa_check.py {output_pptx}")
    print(f"Run visual QA with:")
    print(f"  python scripts/thumbnail.py {output_pptx}")


# ─────────────────────────────────────────────────────────────────────────────
# LAYOUT INTELLIGENCE
# ─────────────────────────────────────────────────────────────────────────────

# Layout variants that can be applied to content slides
LAYOUT_VARIANTS = [
    "accent_bar",       # Left vertical accent bar + text (default, most common)
    "numbered_list",    # Large auto-numbered circles for each bullet point
    "stat_highlight",   # First bullet promoted to large stat callout
    "two_tone",         # Left colored panel (40%) + right content (60%)
    "header_band",      # Thick colored top band with title, clean body below
    "card_grid",        # Cards layout for multiple items
    "timeline",         # Timeline layout for sequential content
    "split_diagonal",   # Diagonal split layout
    "image_focus",      # Large image area with text overlay
    "quote_block",      # Quote/callout centered layout
]


def _pick_layout_variant(
    slide_type: str,
    body_items: List[str],
    layout_streak: List[str],
    slide_index: int,
    theme: dict,
) -> str:
    """Choose a layout variant to maximize visual variety."""
    if slide_type not in ("content", "list_content", "agenda"):
        return "none"

    # Count recent variant usage
    recent = layout_streak[-3:] if len(layout_streak) >= 3 else layout_streak
    recent_set = set(recent)

    # Prefer variants not recently used
    candidates = [v for v in LAYOUT_VARIANTS if v not in recent_set]
    if not candidates:
        candidates = LAYOUT_VARIANTS[:]

    # Heuristic rules for best variant given content shape
    if len(body_items) >= 4:
        # Many items → numbered list works well
        if "numbered_list" in candidates:
            return "numbered_list"
    elif len(body_items) <= 2 and any(
        any(c.isdigit() for c in item) for item in body_items
    ):
        # Short items with numbers → stat highlight
        if "stat_highlight" in candidates:
            return "stat_highlight"
    elif slide_index % 4 == 3:
        # Every 4th slide: two-tone for variety
        if "two_tone" in candidates:
            return "two_tone"

    # Default rotation
    return candidates[slide_index % len(candidates)]


def _restructure_slide(xml: str, theme: dict, layout_variant: str,
                       body_items: List[str], use_dark: bool) -> str:
    """Apply structural layout changes to a content slide.

    This operates at the XML level to add visual elements that change
    the perceived layout — without touching the text content itself.
    """
    if layout_variant == "none" or not layout_variant:
        return xml

    # Don't restructure slides that already have complex custom shapes
    existing_sp_count = xml.count("<p:sp>")
    if existing_sp_count > 6:
        # Too many custom shapes — likely already complex, leave it
        return xml

    primary = theme["primary"]
    secondary = theme["secondary"]
    accent = theme["accent"]
    bg_color = theme["bg_dark"] if use_dark else theme["bg_light"]

    if layout_variant == "two_tone":
        xml = _add_two_tone_panel(xml, primary, use_dark, theme)
    elif layout_variant == "header_band":
        xml = _add_header_band(xml, primary, use_dark, theme)
    elif layout_variant == "numbered_list":
        xml = _add_numbered_circles(xml, primary, body_items, theme)
    elif layout_variant == "stat_highlight":
        xml = _add_stat_highlight(xml, primary, accent, body_items, theme)
    elif layout_variant == "card_grid":
        xml = _add_card_grid(xml, primary, accent, body_items, theme, use_dark)
    elif layout_variant == "timeline":
        xml = _add_timeline(xml, primary, accent, body_items, theme)
    elif layout_variant == "split_diagonal":
        xml = _add_split_diagonal(xml, primary, secondary, use_dark, theme)
    elif layout_variant == "image_focus":
        xml = _add_image_focus_frame(xml, primary, theme)
    elif layout_variant == "quote_block":
        xml = _add_quote_block(xml, primary, accent, body_items, theme, use_dark)
    # accent_bar is handled by _add_accent_bar (already exists)

    return xml


def _add_two_tone_panel(xml: str, primary: str, use_dark: bool, theme: dict) -> str:
    """Add a colored left panel occupying ~35% of slide width.
    Creates a two-tone layout: dark panel left, light content right."""
    panel_color = primary
    # Panel: x=0, y=0, w=3.5", h=5.625" (full slide height)
    # Slide is 10" wide × 5.625" tall = 9144000 EMU × 5143500 EMU
    panel_xml = (
        '\n<p:sp>'
        '<p:nvSpPr>'
        '<p:cNvPr id="9010" name="TwoTonePanel"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr/>'
        '</p:nvSpPr>'
        '<p:spPr>'
        '<a:xfrm>'
        '<a:off x="0" y="0"/>'
        '<a:ext cx="3200400" cy="5143500"/>'   # 3.5" × 5.625"
        '</a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{panel_color}"><a:alpha val="92000"/></a:srgbClr></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:sp>'
    )
    if "</p:spTree>" in xml and "TwoTonePanel" not in xml:
        # Insert at the beginning of spTree (so it goes behind text)
        xml = xml.replace("<p:spTree>", "<p:spTree>" + panel_xml, 1)
    return xml


def _add_header_band(xml: str, primary: str, use_dark: bool, theme: dict) -> str:
    """Add a thick colored band at the top of the slide (title area)."""
    band_xml = (
        '\n<p:sp>'
        '<p:nvSpPr>'
        '<p:cNvPr id="9011" name="HeaderBand"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr/>'
        '</p:nvSpPr>'
        '<p:spPr>'
        '<a:xfrm>'
        '<a:off x="0" y="0"/>'
        '<a:ext cx="9144000" cy="1143000"/>'   # Full width × 1.25"
        '</a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:sp>'
    )
    if "</p:spTree>" in xml and "HeaderBand" not in xml:
        xml = xml.replace("<p:spTree>", "<p:spTree>" + band_xml, 1)
    return xml


def _add_numbered_circles(xml: str, primary: str, body_items: List[str],
                          theme: dict) -> str:
    """Add colored number circles alongside the first N bullet items."""
    if not body_items:
        return xml

    count = min(len(body_items), 6)
    # Start y at 1.3" (below title), each row is ~0.75"
    # Circles: diameter 0.45", positioned at x=0.35"
    circle_size = 411480    # 0.45" in EMU
    start_y = 1188000       # 1.3"
    row_height = 685800     # 0.75"
    circle_x = 320040       # 0.35"

    circles_xml = ""
    for idx in range(count):
        cy = start_y + idx * row_height
        num = str(idx + 1)
        circles_xml += (
            f'\n<p:sp>'
            f'<p:nvSpPr>'
            f'<p:cNvPr id="{9020 + idx}" name="NumCircle{idx}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr/>'
            f'</p:nvSpPr>'
            f'<p:spPr>'
            f'<a:xfrm><a:off x="{circle_x}" y="{cy}"/>'
            f'<a:ext cx="{circle_size}" cy="{circle_size}"/></a:xfrm>'
            f'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln>'
            f'</p:spPr>'
            f'<p:txBody>'
            f'<a:bodyPr anchor="ctr"/>'
            f'<a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="en-US" sz="1400" b="1" dirty="0">'
            f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            f'</a:rPr><a:t>{num}</a:t></a:r>'
            f'</a:p>'
            f'</p:txBody>'
            f'</p:sp>'
        )

    if "</p:spTree>" in xml and "NumCircle0" not in xml:
        xml = xml.replace("</p:spTree>", circles_xml + "</p:spTree>")

    return xml


def _add_stat_highlight(xml: str, primary: str, accent: str,
                        body_items: List[str], theme: dict) -> str:
    """Promote the first bullet point to a large stat callout box."""
    if not body_items:
        return xml

    first_item = body_items[0][:50] if body_items else ""

    # Large callout box: x=0.5", y=1.4", w=4", h=2.5"
    stat_xml = (
        '\n<p:sp>'
        '<p:nvSpPr>'
        '<p:cNvPr id="9030" name="StatCallout"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr/>'
        '</p:nvSpPr>'
        '<p:spPr>'
        '<a:xfrm>'
        '<a:off x="457200" y="1280160"/>'    # 0.5", 1.4"
        '<a:ext cx="3657600" cy="2286000"/>' # 4" × 2.5"
        '</a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr>'
        '<p:txBody>'
        '<a:bodyPr anchor="ctr" wrap="square"/>'
        '<a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/>'
        f'<a:r><a:rPr lang="zh-CN" sz="2400" b="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        f'</a:rPr><a:t>{first_item}</a:t></a:r>'
        '</a:p>'
        '</p:txBody>'
        '</p:sp>'
    )

    if "</p:spTree>" in xml and "StatCallout" not in xml:
        xml = xml.replace("</p:spTree>", stat_xml + "</p:spTree>")

    return xml


def _add_card_grid(xml: str, primary: str, accent: str,
                   body_items: List[str], theme: dict, use_dark: bool) -> str:
    """Add card-style grid layout for multiple content items."""
    if not body_items or len(body_items) < 2:
        return xml

    # Limit to 4 cards max for visual clarity
    count = min(len(body_items), 4)
    cards_per_row = 2 if count > 2 else count
    rows = (count + cards_per_row - 1) // cards_per_row

    # Card dimensions
    card_width = 3200400    # 3.5"
    card_height = 1371600   # 1.5"
    gap = 228600            # 0.25"
    start_x = 914400        # 1"
    start_y = 1371600       # 1.5"

    bg_color = theme.get("bg_light", "FFFFFF") if not use_dark else theme.get("bg_dark", "1A1A2E")
    text_color = theme.get("text_on_light", "000000") if not use_dark else theme.get("text_on_dark", "FFFFFF")

    cards_xml = ""
    for idx in range(count):
        row = idx // cards_per_row
        col = idx % cards_per_row
        cx = start_x + col * (card_width + gap)
        cy = start_y + row * (card_height + gap)

        # Alternate colors for visual interest
        card_color = primary if idx % 2 == 0 else accent
        # If accent is too light for dark mode, use primary
        if use_dark and card_color == accent:
            card_color = primary

        item_text = body_items[idx][:40] if idx < len(body_items) else ""

        cards_xml += (
            f'\n<p:sp>'
            f'<p:nvSpPr>'
            f'<p:cNvPr id="{9040 + idx}" name="Card{idx}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr/>'
            f'</p:nvSpPr>'
            f'<p:spPr>'
            f'<a:xfrm><a:off x="{cx}" y="{cy}"/>'
            f'<a:ext cx="{card_width}" cy="{card_height}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{card_color}"><a:alpha val="15000"/></a:srgbClr></a:solidFill>'
            f'<a:ln><a:solidFill><a:srgbClr val="{card_color}"/></a:solidFill></a:ln>'
            f'</p:spPr>'
            f'<p:txBody>'
            f'<a:bodyPr anchor="ctr" wrap="square"/>'
            f'<a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="zh-CN" sz="1600" b="1" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{text_color}"/></a:solidFill>'
            f'</a:rPr><a:t>{item_text}</a:t></a:r>'
            f'</a:p>'
            f'</p:txBody>'
            f'</p:sp>'
        )

    if "</p:spTree>" in xml and "Card0" not in xml:
        xml = xml.replace("</p:spTree>", cards_xml + "</p:spTree>")

    return xml


def _add_timeline(xml: str, primary: str, accent: str,
                  body_items: List[str], theme: dict) -> str:
    """Add timeline layout for sequential content."""
    if not body_items or len(body_items) < 2:
        return xml

    count = min(len(body_items), 5)
    timeline_y = 2286000    # 2.5"
    start_x = 685800        # 0.75"
    end_x = 8458200         # 9.25"
    step = (end_x - start_x) // (count - 1) if count > 1 else 0

    timeline_xml = (
        # Horizontal line
        f'\n<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="9050" name="TimelineLine"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{start_x}" y="{timeline_y}"/>'
        f'<a:ext cx="{end_x - start_x}" cy="76200"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
    )

    # Timeline nodes
    for idx in range(count):
        cx = start_x + idx * step
        item_text = body_items[idx][:25] if idx < len(body_items) else ""

        timeline_xml += (
            # Node circle
            f'\n<p:sp>'
            f'<p:nvSpPr>'
            f'<p:cNvPr id="{9060 + idx}" name="TimelineNode{idx}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr/>'
            f'</p:nvSpPr>'
            f'<p:spPr>'
            f'<a:xfrm><a:off x="{cx - 228600}" y="{timeline_y - 190500}"/>'
            f'<a:ext cx="457200" cy="457200"/></a:xfrm>'
            f'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{accent}"/></a:solidFill>'
            f'<a:ln><a:solidFill><a:srgbClr val="{primary}"/></a:solidFill></a:ln>'
            f'</p:spPr>'
            f'<p:txBody>'
            f'<a:bodyPr anchor="ctr"/>'
            f'<a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="en-US" sz="1400" b="1" dirty="0">'
            f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            f'</a:rPr><a:t>{idx + 1}</a:t></a:r>'
            f'</a:p>'
            f'</p:txBody>'
            f'</p:sp>'
            # Label below
            f'\n<p:sp>'
            f'<p:nvSpPr>'
            f'<p:cNvPr id="{9070 + idx}" name="TimelineLabel{idx}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr/>'
            f'</p:nvSpPr>'
            f'<p:spPr>'
            f'<a:xfrm><a:off x="{cx - 457200}" y="{timeline_y + 685800}"/>'
            f'<a:ext cx="914400" cy="457200"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'<a:noFill/>'
            f'<a:ln><a:noFill/></a:ln>'
            f'</p:spPr>'
            f'<p:txBody>'
            f'<a:bodyPr anchor="ctr" wrap="square"/>'
            f'<a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="zh-CN" sz="1200" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
            f'</a:rPr><a:t>{item_text}</a:t></a:r>'
            f'</a:p>'
            f'</p:txBody>'
            f'</p:sp>'
        )

    if "</p:spTree>" in xml and "TimelineLine" not in xml:
        xml = xml.replace("</p:spTree>", timeline_xml + "</p:spTree>")

    return xml


def _add_split_diagonal(xml: str, primary: str, secondary: str,
                        use_dark: bool, theme: dict) -> str:
    """Add diagonal split layout."""
    # Create a diagonal shape using a polygon
    diagonal_xml = (
        '\n<p:sp>'
        '<p:nvSpPr>'
        '<p:cNvPr id="9080" name="DiagonalSplit"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr/>'
        '</p:nvSpPr>'
        '<p:spPr>'
        '<a:xfrm><a:off x="0" y="0"/>'
        '<a:ext cx="9144000" cy="5143500"/></a:xfrm>'
        '<a:custGeom>'
        '<a:avLst/>'
        '<a:gdLst>'
        '<a:gd name="x1" fmla="val 0"/>'
        '<a:gd name="y1" fmla="val 0"/>'
        '<a:gd name="x2" fmla="val 4572000"/>'
        '<a:gd name="y2" fmla="val 0"/>'
        '<a:gd name="x3" fmla="val 9144000"/>'
        '<a:gd name="y3" fmla="val 5143500"/>'
        '<a:gd name="x4" fmla="val 0"/>'
        '<a:gd name="y4" fmla="val 5143500"/>'
        '</a:gdLst>'
        '<a:pathLst>'
        '<a:path w="9144000" h="5143500">'
        '<a:moveTo><a:pt x="x1" y="y1"/></a:moveTo>'
        '<a:lnTo><a:pt x="x2" y="y2"/></a:lnTo>'
        '<a:lnTo><a:pt x="x3" y="y3"/></a:lnTo>'
        '<a:lnTo><a:pt x="x4" y="y4"/></a:lnTo>'
        '<a:close/>'
        '</a:path>'
        '</a:pathLst>'
        '</a:custGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"><a:alpha val="20000"/></a:srgbClr></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:sp>'
    )

    if "</p:spTree>" in xml and "DiagonalSplit" not in xml:
        xml = xml.replace("<p:spTree>", "<p:spTree>" + diagonal_xml, 1)

    return xml


def _add_image_focus_frame(xml: str, primary: str, theme: dict) -> str:
    """Add decorative frame for image-focused slides."""
    # Create a subtle border frame
    frame_xml = (
        '\n<p:sp>'
        '<p:nvSpPr>'
        '<p:cNvPr id="9090" name="ImageFrame"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr/>'
        '</p:nvSpPr>'
        '<p:spPr>'
        '<a:xfrm><a:off x="457200" y="457200"/>'
        '<a:ext cx="8229600" cy="4229100"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        f'<a:ln w="76200"><a:solidFill><a:srgbClr val="{primary}"/></a:solidFill></a:ln>'
        '</p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:sp>'
        # Corner accents
        f'\n<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="9091" name="CornerTL"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="457200" y="457200"/>'
        f'<a:ext cx="228600" cy="228600"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
        f'\n<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="9092" name="CornerBR"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="8458200" y="4229100"/>'
        f'<a:ext cx="228600" cy="228600"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
    )

    if "</p:spTree>" in xml and "ImageFrame" not in xml:
        xml = xml.replace("</p:spTree>", frame_xml + "</p:spTree>")

    return xml


def _add_quote_block(xml: str, primary: str, accent: str,
                     body_items: List[str], theme: dict, use_dark: bool) -> str:
    """Add decorative quote/callout block layout."""
    if not body_items:
        return xml

    quote_text = body_items[0][:80] if body_items else ""
    bg_color = theme.get("bg_light", "FFFFFF") if not use_dark else theme.get("bg_dark", "1A1A2E")
    text_color = theme.get("text_on_light", "000000") if not use_dark else theme.get("text_on_dark", "FFFFFF")

    quote_xml = (
        # Left accent bar
        f'\n<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="9100" name="QuoteBar"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="1371600" y="1371600"/>'
        f'<a:ext cx="114300" cy="2746380"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{accent}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
        # Quote mark
        f'\n<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="9101" name="QuoteMark"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="914400" y="1143000"/>'
        f'<a:ext cx="457200" cy="457200"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr anchor="ctr"/>'
        f'<a:lstStyle/>'
        f'<a:p><a:pPr algn="ctr"/>'
        f'<a:r><a:rPr lang="en-US" sz="4800" i="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        f'</a:rPr><a:t>"</a:t></a:r>'
        f'</a:p>'
        f'</p:txBody>'
        f'</p:sp>'
        # Quote text box
        f'\n<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="9102" name="QuoteText"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="1600200" y="1600200"/>'
        f'<a:ext cx="5943600" cy="2286000"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr anchor="ctr" wrap="square"/>'
        f'<a:lstStyle/>'
        f'<a:p><a:pPr algn="l"/>'
        f'<a:r><a:rPr lang="zh-CN" sz="2000" i="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="{text_color}"/></a:solidFill>'
        f'</a:rPr><a:t>{quote_text}</a:t></a:r>'
        f'</a:p>'
        f'</p:txBody>'
        f'</p:sp>'
    )

    if "</p:spTree>" in xml and "QuoteBar" not in xml:
        xml = xml.replace("</p:spTree>", quote_xml + "</p:spTree>")

    return xml


def _beautify_slide(
    slide_path: Path,
    theme: dict,
    use_dark: bool,
    slide_type: str,
    dark_mode: bool,
    layout_variant: str = "accent_bar",
    body_items: Optional[List[str]] = None,
    restructure: bool = True,
    verbose: bool = False,
    enhance_images: bool = True,
    use_gradient: bool = False,
) -> None:
    """Apply theme to a single slide XML file."""
    xml = slide_path.read_text(encoding="utf-8")

    # Check if AI ladder is enabled
    ladder_enabled = theme.get("ladder_enabled", False)

    # 1. Set background (with optional gradient)
    if use_gradient and slide_type in ("title", "section", "conclusion"):
        xml = _set_gradient_background(xml, theme, use_dark, ladder_enabled)
    else:
        xml = _set_background(xml, theme, use_dark, ladder_enabled)

    # 2. Update text colors for contrast
    xml = _update_text_colors(xml, theme, use_dark, ladder_enabled)

    # 3. Update shape fill colors
    xml = _update_shape_colors(xml, theme, use_dark, ladder_enabled)

    # 4. Update font faces
    xml = _update_fonts(xml, theme)

    # 5. Update font sizes for hierarchy (with smart scaling)
    xml = _update_font_sizes_smart(xml, theme, slide_type, body_items)

    # 6. Remove anti-patterns
    xml = _remove_antipatterns(xml)

    # 7. Enhance images (rounded corners, shadows, borders)
    if enhance_images:
        xml = _enhance_images(xml, theme, use_dark)

    # 8. Structural layout enrichment
    if restructure and slide_type in ("content", "list_content", "agenda"):
        if layout_variant and layout_variant not in ("none", "accent_bar"):
            xml = _restructure_slide(xml, theme, layout_variant, body_items or [], use_dark)

    # 9. Add visual accent for content slides
    if slide_type in ("content", "list_content", "agenda") and not use_dark:
        if layout_variant in ("accent_bar", "none") or not restructure:
            xml = _add_accent_bar(xml, theme)

    # 10. Optimize paragraph spacing
    xml = _optimize_paragraph_spacing(xml, theme)

    slide_path.write_text(xml, encoding="utf-8")


def _set_background(xml: str, theme: dict, use_dark: bool) -> str:
    """Set the slide background color."""
    bg_color = theme["bg_dark"] if use_dark else theme["bg_light"]

    # Check if background already defined
    if "<p:bg>" in xml:
        # Replace existing background color
        xml = re.sub(
            r'(<p:bg>.*?<a:solidFill>.*?<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
            lambda m: m.group(1) + bg_color + m.group(2),
            xml,
            flags=re.DOTALL,
        )
    else:
        # Add background after <p:cSld ... > opening
        bg_xml = (
            f'\n  <p:bg>'
            f'<p:bgPr>'
            f'<a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill>'
            f'<a:effectLst/>'
            f'</p:bgPr>'
            f'</p:bg>'
        )
        xml = re.sub(r'(<p:cSld[^>]*>)', r'\1' + bg_xml, xml, count=1)

    return xml


def _set_gradient_background(xml: str, theme: dict, use_dark: bool) -> str:
    """Set a gradient background for visual interest."""
    start_color = theme.get("gradient_start", theme["primary"])
    end_color = theme.get("gradient_end", theme["secondary"])

    # Create gradient fill
    gradient_xml = (
        f'\n  <p:bg>'
        f'<p:bgPr>'
        f'<a:gradFill rotWithShape="1">'
        f'<a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="{start_color}"/></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="{end_color}"/></a:gs>'
        f'</a:gsLst>'
        f'<a:lin ang="2700000" scaled="1"/>'  # Top to bottom gradient
        f'<a:tileRect/>'
        f'</a:gradFill>'
        f'<a:effectLst/>'
        f'</p:bgPr>'
        f'</p:bg>'
    )

    if "<p:bg>" in xml:
        # Replace existing background
        xml = re.sub(r'<p:bg>.*?</p:bg>', gradient_xml.strip(), xml, flags=re.DOTALL)
    else:
        xml = re.sub(r'(<p:cSld[^>]*>)', r'\1' + gradient_xml, xml, count=1)

    return xml


def _enhance_images(xml: str, theme: dict, use_dark: bool) -> str:
    """Enhance images with rounded corners, shadows, and borders."""
    primary = theme["primary"]

    def enhance_pic(m):
        pic_xml = m.group(0)

        # Check if this is actually an image (blip fill)
        if "<a:blip" not in pic_xml:
            return pic_xml

        # Add rounded corners to the shape
        if '<a:prstGeom prst="rect"' in pic_xml:
            # Replace rect with roundRect
            pic_xml = pic_xml.replace(
                '<a:prstGeom prst="rect"><a:avLst/>',
                '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst>'
            )

        # Add subtle shadow effect
        if '<a:effectLst>' not in pic_xml and '<p:spPr>' in pic_xml:
            shadow_xml = (
                '<a:effectLst>'
                '<a:outerShdw blurRad="63500" dist="50800" dir="2700000" algn="bl">'
                '<a:srgbClr val="000000"><a:alpha val="25000"/></a:srgbClr>'
                '</a:outerShdw>'
                '</a:effectLst>'
            )
            # Insert before closing spPr
            pic_xml = pic_xml.replace('</p:spPr>', shadow_xml + '</p:spPr>')

        # Add subtle border
        if '<a:ln>' not in pic_xml and '<p:spPr>' in pic_xml:
            border_xml = (
                f'<a:ln w="12700">'
                f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
                f'</a:ln>'
            )
            pic_xml = pic_xml.replace('</p:spPr>', border_xml + '</p:spPr>')

        return pic_xml

    xml = re.sub(r'<p:pic>.*?</p:pic>', enhance_pic, xml, flags=re.DOTALL)
    return xml


def _update_font_sizes_smart(xml: str, theme: dict, slide_type: str,
                              body_items: Optional[List[str]]) -> str:
    """Smart font sizing based on content length and slide type."""
    # First apply basic font size updates
    xml = _update_font_sizes(xml, theme, slide_type)

    # Calculate content density for dynamic sizing
    content_length = sum(len(item) for item in (body_items or []))
    item_count = len(body_items) if body_items else 0

    # Adjust body font size based on content density
    if content_length > 300 or item_count > 6:
        # High density - reduce font size slightly
        target_size = theme["body_size"] - 200  # 2pt smaller
    elif content_length < 100 and item_count <= 3:
        # Low density - can increase font size
        target_size = theme["body_size"] + 100  # 1pt larger
    else:
        target_size = theme["body_size"]

    def adjust_body_size(m):
        sp_xml = m.group(0)
        # Skip title placeholders
        if re.search(r'<p:ph[^>]*type="(?:title|ctrTitle)"', sp_xml):
            return sp_xml

        def fix_sz(sm):
            sz = int(sm.group(1))
            # Only adjust if within body text range
            if 1400 <= sz <= 2200:
                return f'sz="{target_size}"'
            return sm.group(0)

        return re.sub(r'sz="(\d+)"', fix_sz, sp_xml)

    xml = re.sub(r'<p:sp\b.*?</p:sp>', adjust_body_size, xml, flags=re.DOTALL)
    return xml


def _optimize_paragraph_spacing(xml: str, theme: dict) -> str:
    """Optimize paragraph spacing for better readability."""
    # Add line spacing to paragraphs
    def add_spacing(m):
        p_xml = m.group(0)

        # Skip if already has spacing
        if '<a:lnSpc' in p_xml:
            return p_xml

        # Add 1.2 line spacing (120%)
        spacing_xml = '<a:lnSpc><a:spcPct val="120000"/></a:lnSpc>'

        # Insert after pPr or at the beginning of the paragraph
        if '<a:pPr' in p_xml:
            p_xml = p_xml.replace('</a:pPr>', spacing_xml + '</a:pPr>')
        else:
            p_xml = p_xml.replace('<a:p>', '<a:p><a:pPr>' + spacing_xml + '</a:pPr>')

        return p_xml

    # Only apply to body text paragraphs (not titles)
    xml = re.sub(r'<a:p>.*?</a:p>', add_spacing, xml, flags=re.DOTALL)
    return xml


# ═════════════════════════════════════════════════════════════════════════════
# SMART ICON SYSTEM
# ═════════════════════════════════════════════════════════════════════════════

# Keyword to icon mapping (using Unicode symbols as fallback)
ICON_KEYWORDS = {
    # Business & Strategy
    "growth": "📈",
    "increase": "📈",
    "revenue": "💰",
    "profit": "💵",
    "money": "💵",
    "cost": "💸",
    "budget": "📊",
    "finance": "📊",
    "investment": "📈",
    "market": "🏪",
    "customer": "👥",
    "client": "👤",
    "user": "👤",
    "team": "👥",
    "people": "👥",
    "partner": "🤝",
    "collaboration": "🤝",
    # Technology
    "technology": "💻",
    "tech": "💻",
    "digital": "🔌",
    "software": "💿",
    "app": "📱",
    "mobile": "📱",
    "data": "📊",
    "analytics": "📈",
    "ai": "🤖",
    "automation": "⚙️",
    "cloud": "☁️",
    "security": "🔒",
    "privacy": "🔐",
    # Goals & Success
    "goal": "🎯",
    "target": "🎯",
    "success": "🏆",
    "achievement": "🏆",
    "win": "🏆",
    "milestone": "🚩",
    "launch": "🚀",
    "start": "🚀",
    "begin": "🚀",
    # Time & Process
    "time": "⏰",
    "deadline": "⏰",
    "schedule": "📅",
    "plan": "📋",
    "process": "🔄",
    "workflow": "🔄",
    "step": "👣",
    "phase": "🔄",
    # Quality & Innovation
    "quality": "✨",
    "innovation": "💡",
    "idea": "💡",
    "creative": "🎨",
    "design": "🎨",
    "solution": "🔧",
    "problem": "⚠️",
    "risk": "⚠️",
    "warning": "⚠️",
    # Communication
    "communication": "💬",
    "message": "💬",
    "email": "📧",
    "phone": "📞",
    "call": "📞",
    "meeting": "🤝",
    "presentation": "📊",
    # Environment
    "environment": "🌱",
    "sustainability": "♻️",
    "green": "🌿",
    "eco": "🌍",
    # Health & Wellness
    "health": "❤️",
    "wellness": "🧘",
    "fitness": "💪",
    # Education
    "education": "📚",
    "learning": "📖",
    "training": "🎓",
    "knowledge": "🧠",
}


def _insert_smart_icons(body_items: List[str], theme: dict) -> List[str]:
    """Insert relevant icons based on content keywords."""
    enhanced_items = []

    for item in body_items:
        enhanced_item = item
        item_lower = item.lower()

        # Find matching icon
        matched_icon = None
        for keyword, icon in ICON_KEYWORDS.items():
            if keyword in item_lower:
                matched_icon = icon
                break

        # Add icon if found and not already present
        if matched_icon and matched_icon not in item:
            enhanced_item = f"{matched_icon} {item}"

        enhanced_items.append(enhanced_item)

    return enhanced_items


def _update_text_colors(xml: str, theme: dict, use_dark: bool) -> str:
    """Update text run colors to match the theme palette.

    Strategy:
    - Title placeholders (type="title" / "ctrTitle") → theme primary (light bg)
      or text_on_dark (dark bg)
    - All other text runs that carry an explicit solidFill → theme body color
    - Runs that inherit color (no solidFill) are left untouched so the master
      theme inheritance chain works correctly
    """
    title_color = theme["text_on_dark"] if use_dark else theme["primary"]
    body_color  = theme["text_on_dark"] if use_dark else theme["text_on_light"]

    def recolor_sp(m):
        sp_xml = m.group(0)
        # Determine if this shape is a title placeholder
        is_title = bool(re.search(r'<p:ph[^>]*type="(?:title|ctrTitle)"', sp_xml))
        color = title_color if is_title else body_color

        def recolor_rpr(rm):
            rpr = rm.group(0)
            # Only recolor runs that already carry an explicit solidFill color;
            # leave scheme-color references and purely inherited runs alone.
            if "<a:solidFill>" not in rpr:
                return rpr
            # Strip any existing srgbClr / sysClr inside solidFill, replace with theme color
            new_rpr = re.sub(
                r'<a:solidFill>.*?</a:solidFill>',
                f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>',
                rpr,
                flags=re.DOTALL,
            )
            return new_rpr

        return re.sub(r'<a:rPr\b.*?</a:rPr>', recolor_rpr, sp_xml, flags=re.DOTALL)

    xml = re.sub(r'<p:sp\b.*?</p:sp>', recolor_sp, xml, flags=re.DOTALL)
    return xml


def _update_shape_colors(xml: str, theme: dict, use_dark: bool) -> str:
    """Update accent shape fill colors to theme accent/primary.

    Skips shapes that serve as functional backgrounds (very light or very dark).
    Replaces all other solid-fill shapes with the theme accent color so that
    decorative elements automatically adopt the new palette.
    """
    accent_color = theme["accent"] if use_dark else theme["primary"]
    bg_light = theme["bg_light"].upper()
    bg_dark  = theme["bg_dark"].upper()

    def replace_accent(m):
        color = m.group(1).upper()
        r = int(color[0:2], 16)
        g = int(color[2:4], 16)
        b = int(color[4:6], 16)
        luminance = (r + g + b) / 3

        # Skip pure white / near-white (backgrounds / light fills)
        if luminance > 235:
            return m.group(0)
        # Skip pure black / near-black (text shadows, etc.)
        if luminance < 20:
            return m.group(0)
        # Skip the template's own bg colors to avoid color-flipping backgrounds
        if color in (bg_light, bg_dark):
            return m.group(0)
        # Replace everything else with the theme accent
        return f'<a:solidFill><a:srgbClr val="{accent_color}"/></a:solidFill>'

    # Only update spPr fills (shape properties), not text run fills
    def process_sppr(m):
        sppr = m.group(0)
        sppr = re.sub(
            r'<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/></a:solidFill>',
            replace_accent,
            sppr,
        )
        return sppr

    xml = re.sub(r'<p:spPr>.*?</p:spPr>', process_sppr, xml, flags=re.DOTALL)
    return xml


def _update_fonts(xml: str, theme: dict) -> str:
    """Update font faces to match theme."""
    header_font = theme["header_font"]
    body_font = theme["body_font"]

    # Replace common default fonts
    common_fonts = [
        "Arial", "Helvetica", "Times New Roman", "Times", "Calibri",
        "Cambria", "Verdana", "Tahoma", "Trebuchet MS", "Georgia",
        "Palatino", "Garamond", "Comic Sans MS", "Impact",
        "Calibri Light", "Century Gothic",
    ]

    def replace_font(m):
        typeface = m.group(1)
        if any(f.lower() == typeface.lower() for f in common_fonts):
            # For now, use body font as a safe default
            # Title fonts are harder to detect from raw XML runs
            return f'<a:latin typeface="{body_font}"'
        return m.group(0)

    xml = re.sub(r'<a:latin typeface="([^"]+)"', replace_font, xml)
    return xml


def _update_font_sizes(xml: str, theme: dict, slide_type: str) -> str:
    """Enforce font size hierarchy."""
    # This is a light touch - only fix obviously wrong sizes
    # Title sizes: ensure they're at least 3600 (36pt)
    # Body sizes: ensure they're between 1400-2000

    def fix_title_size(m):
        sp_xml = m.group(0)
        if not re.search(r'<p:ph[^>]*type="(?:title|ctrTitle)"', sp_xml):
            return sp_xml
        # Update font sizes in title
        def fix_sz(sm):
            sz = int(sm.group(1))
            if sz < 3200:  # Too small for a title
                return f'sz="{theme["title_size"]}"'
            return sm.group(0)
        return re.sub(r'sz="(\d+)"', fix_sz, sp_xml)

    xml = re.sub(r'<p:sp\b.*?</p:sp>', fix_title_size, xml, flags=re.DOTALL)
    return xml


def _remove_antipatterns(xml: str) -> str:
    """Remove common visual anti-patterns."""
    # Remove thin horizontal lines that look like title underlines
    # These are shapes with height < 0.05" (45720 EMU) positioned right after a title
    # We detect them heuristically: LINE shapes or very thin RECTANGLE shapes

    # NOTE: remove_thin_lines function removed - was dead code (only had 'pass')
    # Anti-pattern removal is now handled entirely by remove_accent_underlines

    # Remove "accent lines under titles" - thin rectangles positioned as title decorators
    # These typically have h ~3600 EMU (0.05"), full width
    def remove_accent_underlines(m):
        sp_xml = m.group(0)
        if '<p:ph' in sp_xml:
            return sp_xml
        # Very thin horizontal shape with large width
        if re.search(r'cy="[1-9][0-9]{3}"', sp_xml) and re.search(r'cx="[5-9][0-9]{5,}"', sp_xml):
            # This matches thin horizontal elements wider than ~0.5" - might be accent lines
            # Only remove if they're solid fills (decorative)
            if '<a:solidFill>' in sp_xml and '<p:ph' not in sp_xml:
                return ''  # Remove the element
        return sp_xml

    # Apply cleanup
    xml = re.sub(r'<p:sp\b.*?</p:sp>', remove_accent_underlines, xml, flags=re.DOTALL)

    return xml


def _add_accent_bar(xml: str, theme: dict) -> str:
    """Add a subtle left accent bar to content slides for visual interest."""
    # Add a thin colored bar on the left side to make slides more dynamic
    # This is a 0.05" wide, 4.5" tall bar at x=0.3", y=0.7"

    accent_bar = (
        '\n<p:sp>'
        '<p:nvSpPr>'
        '<p:cNvPr id="9001" name="AccentBar"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr/>'
        '</p:nvSpPr>'
        '<p:spPr>'
        '<a:xfrm>'
        '<a:off x="274638" y="640080"/>'  # 0.3", 0.7"
        '<a:ext cx="45720" cy="4115040"/>'  # 0.05" wide, 4.5" tall
        '</a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{theme["primary"]}"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr>'
        '<p:txBody>'
        '<a:bodyPr/><a:lstStyle/><a:p/>'
        '</p:txBody>'
        '</p:sp>'
    )

    # Insert before </p:spTree>
    if "</p:spTree>" in xml and "AccentBar" not in xml:
        xml = xml.replace("</p:spTree>", accent_bar + "</p:spTree>")

    return xml


def _apply_theme_to_master(unpacked_dir: Path, theme: dict, verbose: bool) -> None:
    """Update the slide master with theme colors and fonts."""
    master_dir = unpacked_dir / "ppt" / "slideMasters"
    if not master_dir.exists():
        return

    for master_path in master_dir.glob("slideMaster*.xml"):
        if verbose:
            print(f"  Updating master: {master_path.name}")
        xml = master_path.read_text(encoding="utf-8")

        # Update master color scheme
        xml = _update_master_colors(xml, theme)
        xml = _update_fonts(xml, theme)

        master_path.write_text(xml, encoding="utf-8")


def _update_master_colors(xml: str, theme: dict) -> str:
    """Update color references in master XML."""
    # Update dk1 (dark 1 / text) color
    xml = re.sub(
        r'(<a:dk1>.*?<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
        lambda m: m.group(1) + theme["text_on_light"] + m.group(2),
        xml, flags=re.DOTALL
    )
    # Update lt1 (light 1 / background) color
    xml = re.sub(
        r'(<a:lt1>.*?<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
        lambda m: m.group(1) + theme["bg_light"] + m.group(2),
        xml, flags=re.DOTALL
    )
    # Update accent1 color
    xml = re.sub(
        r'(<a:accent1>.*?<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
        lambda m: m.group(1) + theme["primary"] + m.group(2),
        xml, flags=re.DOTALL
    )
    # Update accent2 color
    xml = re.sub(
        r'(<a:accent2>.*?<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
        lambda m: m.group(1) + theme["secondary"] + m.group(2),
        xml, flags=re.DOTALL
    )
    return xml


def _update_theme_xml(unpacked_dir: Path, theme: dict, verbose: bool) -> None:
    """Update the theme XML file with new colors."""
    theme_dir = unpacked_dir / "ppt" / "theme"
    if not theme_dir.exists():
        return

    for theme_path in theme_dir.glob("theme*.xml"):
        if verbose:
            print(f"  Updating theme: {theme_path.name}")
        xml = theme_path.read_text(encoding="utf-8")

        # Update the color scheme in theme XML
        color_replacements = [
            ("dk1", theme["text_on_light"]),
            ("lt1", theme["bg_light"]),
            ("dk2", theme["secondary"]),
            ("lt2", "EEECE1"),  # Keep neutral
            ("accent1", theme["primary"]),
            ("accent2", theme["secondary"]),
            ("accent3", theme["accent"]),
            ("accent4", theme["primary"]),
            ("accent5", theme["secondary"]),
            ("accent6", theme["accent"]),
        ]

        for elem_name, color in color_replacements:
            xml = re.sub(
                rf'(<a:{elem_name}>.*?<a:srgbClr val=")[0-9A-Fa-f]{{6}}(")',
                lambda m, c=color: m.group(1) + c + m.group(2),
                xml, flags=re.DOTALL
            )

        # Update font scheme
        xml = re.sub(
            r'(<a:latin typeface=")[^"]+(" pitchFamily=")',
            lambda m: m.group(1) + theme["header_font"] + m.group(2),
            xml, count=2  # Update first two (major and minor fonts)
        )

        theme_path.write_text(xml, encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Beautify a PPT by redesigning its visual style"
    )
    parser.add_argument(
        "--theme",
        choices=list(THEMES.keys()),
        help=f"Theme to apply. Available: {', '.join(THEMES.keys())}. Default: auto-detect",
    )
    parser.add_argument(
        "--dark-mode",
        action="store_true",
        help="Force dark background on all slides",
    )
    parser.add_argument(
        "--keep-images",
        action="store_true",
        default=True,
        help="Preserve original images (default: True)",
    )
    parser.add_argument(
        "--font-pair",
        help="Override font pair, e.g. 'georgia-calibri' or 'arial_black-arial'",
    )
    parser.add_argument(
        "--no-restructure",
        action="store_true",
        help="Skip layout restructuring (only change colors and fonts)",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Verbose output",
    )
    parser.add_argument(
        "source",
        nargs="?",
        help="Source PPTX file",
    )
    parser.add_argument(
        "output",
        nargs="?",
        help="Output PPTX file",
    )
    parser.add_argument(
        "--list-themes",
        action="store_true",
        help="List available themes and exit",
    )
    parser.add_argument(
        "--no-image-enhance",
        action="store_true",
        help="Skip image enhancement (rounded corners, shadows, borders)",
    )
    parser.add_argument(
        "--gradient-bg",
        action="store_true",
        help="Use gradient backgrounds for title/section slides",
    )
    parser.add_argument(
        "--smart-icons",
        action="store_true",
        help="Auto-insert icons based on content keywords",
    )
    parser.add_argument(
        "--ai-ladder",
        action="store_true",
        help="Enable AI color ladder (overrides theme colors with multi-level gradients)",
    )
    parser.add_argument(
        "--ladder-depth",
        type=int,
        default=5,
        choices=range(3, 11),
        metavar="N",
        help="Number of ladder levels (3-10, default: 5)",
    )
    parser.add_argument(
        "--ladder-strategy",
        choices=["lightness", "saturation", "complementary"],
        default="lightness",
        help="Gradient strategy: lightness (dark→light), saturation (dull→vivid), complementary (color→opposite)",
    )
    parser.add_argument(
        "--brand-color",
        metavar="HEX",
        help="Custom brand color for ladder generation (e.g., '0066CC')",
    )
    args = parser.parse_args()

    if args.list_themes:
        print("Available themes:")
        for name, t in THEMES.items():
            print(f"  {name:12s} — {t['name']}: #{t['primary']} / #{t['accent']}")
            print(f"               Fonts: {t['header_font']} + {t['body_font']}")
        sys.exit(0)

    # Ensure source and output are provided for normal operation
    if not args.source or not args.output:
        parser.error("the following arguments are required: source, output")

    beautify_ppt(
        args.source,
        args.output,
        theme_name=args.theme,
        dark_mode=args.dark_mode,
        keep_images=args.keep_images,
        font_pair=args.font_pair,
        restructure=not args.no_restructure,
        verbose=args.verbose,
        enhance_images=not args.no_image_enhance,
        use_gradient=args.gradient_bg,
        smart_icons=args.smart_icons,
        ai_ladder=args.ai_ladder,
        ladder_depth=args.ladder_depth,
        ladder_strategy=args.ladder_strategy,
        brand_color=args.brand_color,
    )
