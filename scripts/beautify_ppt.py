"""Beautify a PPT by redesigning its visual style while preserving content.

Analyzes the source presentation and applies a professional theme:
- Replaces colors with a cohesive palette
- Improves font hierarchy
- Adds visual structure (background shapes, accent bars)
- Varies slide layouts for visual interest
- Removes common anti-patterns (accent lines under titles, etc.)
- Restructures text-heavy single-column slides into visually richer layouts

Usage:
    python scripts/beautify_ppt.py source.pptx output.pptx
    python scripts/beautify_ppt.py source.pptx output.pptx --theme tech
    python scripts/beautify_ppt.py source.pptx output.pptx --theme executive --dark-mode
    python scripts/beautify_ppt.py source.pptx output.pptx --theme creative --verbose
    python scripts/beautify_ppt.py source.pptx output.pptx --no-restructure  # skip layout changes

Available themes: executive, tech, creative, warm, minimal, bold, nature, ocean
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
) -> None:
    """Redesign a PPT's visual style while preserving content."""

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

    # Override font pair if specified
    if font_pair:
        parts = font_pair.split("-")
        if len(parts) >= 2:
            theme = dict(theme)
            theme["header_font"] = parts[0].replace("_", " ").title()
            theme["body_font"] = parts[1].replace("_", " ").title()

    print(f"Theme: {theme['name']}")
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

            _beautify_slide(
                slide_path, theme, use_dark, slide_type, dark_mode,
                layout_variant=layout_variant,
                body_items=body_items,
                restructure=restructure,
                verbose=verbose,
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
) -> None:
    """Apply theme to a single slide XML file."""
    xml = slide_path.read_text(encoding="utf-8")

    # 1. Set background
    xml = _set_background(xml, theme, use_dark)

    # 2. Update text colors for contrast
    xml = _update_text_colors(xml, theme, use_dark)

    # 3. Update shape fill colors
    xml = _update_shape_colors(xml, theme, use_dark)

    # 4. Update font faces
    xml = _update_fonts(xml, theme)

    # 5. Update font sizes for hierarchy
    xml = _update_font_sizes(xml, theme, slide_type)

    # 6. Remove anti-patterns
    xml = _remove_antipatterns(xml)

    # 7. Structural layout enrichment (NEW)
    if restructure and slide_type in ("content", "list_content", "agenda"):
        if layout_variant and layout_variant not in ("none", "accent_bar"):
            xml = _restructure_slide(xml, theme, layout_variant, body_items or [], use_dark)

    # 8. Add visual accent for content slides
    if slide_type in ("content", "list_content", "agenda") and not use_dark:
        if layout_variant in ("accent_bar", "none") or not restructure:
            xml = _add_accent_bar(xml, theme)

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


def _update_text_colors(xml: str, theme: dict, use_dark: bool) -> str:
    """Update text run colors to match theme."""
    # Determine colors based on background
    title_color = theme["text_on_dark"] if use_dark else theme["primary"]
    body_color = theme["text_on_dark"] if use_dark else theme["text_on_light"]

    def update_rpr(m):
        rpr = m.group(0)
        # Only update if solidFill is present (explicit color), not inherited
        if "<a:solidFill>" not in rpr and "<a:schemeClr" not in rpr:
            return rpr
        # Replace with theme color
        rpr = re.sub(
            r'<a:solidFill><a:srgbClr val="[0-9A-Fa-f]{6}"/></a:solidFill>',
            f'<a:solidFill><a:srgbClr val="{body_color}"/></a:solidFill>',
            rpr,
        )
        return rpr

    # Update all run properties that have explicit colors
    xml = re.sub(r'<a:rPr[^>]*/>', xml, xml)  # noop to warm up
    # This is a simplified approach - in real scenarios we'd parse XML properly
    # Replace obvious old colors that are too far from theme
    bad_colors = ["FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
                  "808080", "C0C0C0", "000080", "008000", "800000"]
    for bad in bad_colors:
        xml = xml.replace(
            f'<a:srgbClr val="{bad}"/>',
            f'<a:srgbClr val="{body_color}"/>'
        )

    return xml


def _update_shape_colors(xml: str, theme: dict, use_dark: bool) -> str:
    """Update accent shape fill colors to theme accent."""
    # Replace common accent colors with theme primary/accent
    # This targets shapes that likely serve as design accents
    # Avoid changing very light or very dark fills (those are often functional)

    accent_color = theme["accent"] if use_dark else theme["primary"]

    # Pattern: standalone shape fills that are solid and colorful
    def replace_accent(m):
        color = m.group(1).upper()
        # Skip very light colors (backgrounds), very dark (shadows), and white/black
        r = int(color[0:2], 16)
        g = int(color[2:4], 16)
        b = int(color[4:6], 16)
        luminance = (r + g + b) / 3
        if luminance > 220 or luminance < 30:
            return m.group(0)
        # Replace with theme accent
        return f'<a:solidFill><a:srgbClr val="{accent_color}"/></a:solidFill>'

    # Only update spPr fills (shape properties), not text fills
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

    def remove_thin_lines(m):
        sp_xml = m.group(0)
        # Check if it's a line shape
        if '<p:ph' in sp_xml:
            return sp_xml  # Don't remove placeholder shapes

        # Check for LINE geometry or very flat rectangle
        is_line = (
            'preset="line"' in sp_xml or
            'prst="line"' in sp_xml or
            re.search(r'cy="[0-9]{1,4}"', sp_xml) and  # height < 10000 EMU (< 0.11")
            not re.search(r'cy="[5-9][0-9]{4,}"', sp_xml)  # not >= 50000 EMU
        )

        if is_line:
            # Check if this is a decorative line (not a chart line)
            if '<p:pic' not in sp_xml and 'chart' not in sp_xml.lower():
                # Remove thin horizontal line shapes
                pass  # We'll keep it but could remove

        return sp_xml

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
    parser.add_argument("source", help="Source PPTX file")
    parser.add_argument("output", help="Output PPTX file")
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
        "--list-themes",
        action="store_true",
        help="List available themes and exit",
    )
    args = parser.parse_args()

    if args.list_themes:
        print("Available themes:")
        for name, t in THEMES.items():
            print(f"  {name:12s} — {t['name']}: #{t['primary']} / #{t['accent']}")
            print(f"               Fonts: {t['header_font']} + {t['body_font']}")
        sys.exit(0)

    beautify_ppt(
        args.source,
        args.output,
        theme_name=args.theme,
        dark_mode=args.dark_mode,
        keep_images=args.keep_images,
        font_pair=args.font_pair,
        restructure=not args.no_restructure,
        verbose=args.verbose,
    )
