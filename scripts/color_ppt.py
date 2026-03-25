#!/usr/bin/env python3
"""
color_ppt.py — AI-powered color replacement for PPTX files.

Directly manipulates PPTX XML to replace colors in shapes, tables, SmartArt,
and other graphic elements. Uses intelligent color classification to maintain
visual hierarchy while applying a new color scheme.

Usage:
    python scripts/color_ppt.py input.pptx output.pptx [--theme THEME] [--verbose]

Themes: executive | tech | creative | warm | minimal | bold | nature |
        ocean | elegant | modern | sunset | forest
"""

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Set, Tuple

# ─────────────────────────────────────────────────────────────────────────────
# THEME DEFINITIONS (same as beautify_ppt.py)
# ─────────────────────────────────────────────────────────────────────────────

THEMES = {
    "executive": {
        "primary": "#1A1A2E",
        "secondary": "#16213E",
        "accent": "#0F3460",
        "text_on_light": "#333333",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#1A1A2E",
        "light_fill": "#F5F5F5",
    },
    "tech": {
        "primary": "#667EEA",
        "secondary": "#764BA2",
        "accent": "#F093FB",
        "text_on_light": "#2D3748",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#1A1A2E",
        "light_fill": "#F7FAFC",
    },
    "creative": {
        "primary": "#FF6B6B",
        "secondary": "#4ECDC4",
        "accent": "#FFE66D",
        "text_on_light": "#2D3436",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#2D3436",
        "light_fill": "#FFFEF0",
    },
    "warm": {
        "primary": "#E07A5F",
        "secondary": "#81B29A",
        "accent": "#ECE2D0",
        "text_on_light": "#3D405B",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#3D405B",
        "light_fill": "#F4F1DE",
    },
    "minimal": {
        "primary": "#2C3E50",
        "secondary": "#34495E",
        "accent": "#BDC3C7",
        "text_on_light": "#2C3E50",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#2C3E50",
        "light_fill": "#ECF0F1",
    },
    "bold": {
        "primary": "#E74C3C",
        "secondary": "#3498DB",
        "accent": "#F1C40F",
        "text_on_light": "#2C3E50",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#1A1A2E",
        "light_fill": "#FFFFFF",
    },
    "nature": {
        "primary": "#27AE60",
        "secondary": "#2980B9",
        "accent": "#F39C12",
        "text_on_light": "#2C3E50",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#1E3A2F",
        "light_fill": "#E8F5E9",
    },
    "ocean": {
        "primary": "#2E86AB",
        "secondary": "#A23B72",
        "accent": "#F18F01",
        "text_on_light": "#2D3436",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#1B3A4B",
        "light_fill": "#E3F2FD",
    },
    "elegant": {
        "primary": "#1A1A2E",
        "secondary": "#4A4E69",
        "accent": "#9A8C98",
        "text_on_light": "#333333",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#1A1A2E",
        "light_fill": "#F8F7F4",
    },
    "modern": {
        "primary": "#6366F1",
        "secondary": "#EC4899",
        "accent": "#10B981",
        "text_on_light": "#1F2937",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#111827",
        "light_fill": "#F9FAFB",
    },
    "sunset": {
        "primary": "#FF6B35",
        "secondary": "#F7C59F",
        "accent": "#2E294E",
        "text_on_light": "#333333",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#2E294E",
        "light_fill": "#FFF8F0",
    },
    "forest": {
        "primary": "#2D5A27",
        "secondary": "#8B4513",
        "accent": "#D4A574",
        "text_on_light": "#333333",
        "text_on_dark": "FFFFFF",
        "dark_fill": "#1A3A16",
        "light_fill": "#F5F5DC",
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# COLOR UTILITIES
# ─────────────────────────────────────────────────────────────────────────────

def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """Convert hex color to RGB tuple."""
    hex_color = hex_color.lstrip('#')
    return (
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16),
    )


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Convert RGB tuple to hex color."""
    return f"{r:02X}{g:02X}{b:02X}"


def classify_color(hex_color: str) -> str:
    """Classify color by luminance: dark, mid, or light."""
    r, g, b = hex_to_rgb(hex_color)
    
    def to_linear(c):
        c = c / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    
    lum = 0.2126 * to_linear(r) + 0.7152 * to_linear(g) + 0.0722 * to_linear(b)
    
    if lum < 0.3:
        return "dark"
    elif lum > 0.7:
        return "light"
    else:
        return "mid"


def build_replacement_map(source_colors: Set[str], theme: Dict) -> Dict[str, str]:
    """Build mapping from source colors to theme colors based on luminance."""
    replacements = {}
    
    dark_sources = []
    light_sources = []
    mid_sources = []
    
    theme_colors = set(theme.values())
    
    for color in source_colors:
        if color in theme_colors:
            continue  # Skip if it's already a theme color
        cls = classify_color(color)
        if cls == "dark":
            dark_sources.append(color)
        elif cls == "light":
            light_sources.append(color)
        else:
            mid_sources.append(color)
    
    # Map dark sources to dark_fill
    dark_fill = theme.get("dark_fill", theme["primary"])
    for color in dark_sources:
        replacements[color] = dark_fill
    
    # Map light sources to light_fill
    light_fill = theme.get("light_fill", theme.get("secondary", "FFFFFF"))
    for color in light_sources:
        replacements[color] = light_fill
    
    # Map mid sources to accent
    accent = theme.get("accent", theme["secondary"])
    for color in mid_sources:
        replacements[color] = accent
    
    return replacements


# ─────────────────────────────────────────────────────────────────────────────
# XML COLOR REPLACEMENT
# ─────────────────────────────────────────────────────────────────────────────

def extract_all_colors(xml: str) -> Set[str]:
    """Extract all srgbClr colors from XML."""
    colors = set()
    
    # Match <a:srgbClr val="RRGGBB">
    for match in re.finditer(r'<a:srgbClr val="([A-Fa-f0-9]{6})"', xml):
        colors.add(match.group(1))
    
    # Match srgbVal="RRGGBB" (older format)
    for match in re.finditer(r'srgbVal="([A-Fa-f0-9]{6})"', xml):
        colors.add(match.group(1))
    
    return colors


def replace_colors_in_xml(xml: str, replacements: Dict[str, str]) -> Tuple[str, int]:
    """Replace colors in XML. Returns (modified_xml, count)."""
    if not replacements:
        return xml, 0
    
    modified = xml
    count = 0
    
    for old_color, new_color in replacements.items():
        old_upper = old_color.upper()
        old_lower = old_color.lower()
        
        # Replace in srgbClr val="..."
        for pattern in [
            rf'(<a:srgbClr val="){old_upper}(")',
            rf'(<a:srgbClr val="){old_lower}(")',
        ]:
            new_xml, n = re.subn(
                pattern,
                lambda m: m.group(1) + new_color + m.group(2),
                modified,
                flags=re.DOTALL
            )
            if new_xml != modified:
                modified = new_xml
                count += n
        
        # Replace in gradient stops
        for pattern in [
            rf'(<a:gradStop[^>]*clrType="srgb"[^>]*val="){old_upper}(")',
            rf'(<a:srgbClr[^>]*val="){old_upper}(")',
        ]:
            new_xml, n = re.subn(pattern, lambda m: m.group(1) + new_color + m.group(2), modified)
            if new_xml != modified:
                modified = new_xml
                count += n
        
        # Replace in line elements <a:ln ...>
        for pattern in [
            rf'(<a:ln[^>]*>.*?<a:solidFill>.*?<a:srgbClr[^>]*val="){old_upper}(")',
        ]:
            new_xml, n = re.subn(pattern, lambda m: m.group(1) + new_color + m.group(2), modified, flags=re.DOTALL)
            if new_xml != modified:
                modified = new_xml
                count += n
    
    return modified, count


# ─────────────────────────────────────────────────────────────────────────────
# PPTX PROCESSING
# ─────────────────────────────────────────────────────────────────────────────

def process_pptx(
    input_path: str,
    output_path: str,
    theme_name: str = "tech",
    verbose: bool = False,
) -> None:
    """Process PPTX file to replace colors in all slides."""
    
    if theme_name not in THEMES:
        print(f"Error: Unknown theme '{theme_name}'")
        print(f"Available themes: {', '.join(THEMES.keys())}")
        sys.exit(1)
    
    theme = THEMES[theme_name]
    
    if verbose:
        print(f"Theme: {theme_name}")
        print(f"  Primary: #{theme['primary']}")
        print(f"  Secondary: #{theme['secondary']}")
        print(f"  Accent: #{theme['accent']}")
        print()
    
    input_path = Path(input_path)
    output_path = Path(output_path)
    
    # Create temp directory
    temp_dir = output_path.parent / f".color_ppt_temp_{output_path.stem}"
    temp_dir.mkdir(exist_ok=True)
    
    try:
        # Extract PPTX
        with zipfile.ZipFile(input_path, 'r') as zf:
            zf.extractall(temp_dir)
        
        # Find all slide XML files
        slides_dir = temp_dir / "ppt" / "slides"
        if not slides_dir.exists():
            print("Error: No slides directory found in PPTX")
            sys.exit(1)
        
        slide_files = sorted(slides_dir.glob("slide*.xml"))
        
        total_replacements = 0
        slides_processed = 0
        
        for slide_file in slide_files:
            xml = slide_file.read_text(encoding="utf-8")
            original_xml = xml
            
            # Extract colors from this slide
            colors = extract_all_colors(xml)
            
            if not colors:
                continue
            
            # Build replacement map
            replacements = build_replacement_map(colors, theme)
            
            if not replacements:
                continue
            
            # Replace colors
            xml, count = replace_colors_in_xml(xml, replacements)
            
            if xml != original_xml:
                slide_file.write_text(xml, encoding="utf-8")
                slides_processed += 1
                total_replacements += count
                
                if verbose:
                    print(f"  {slide_file.name}: {count} replacements")
        
        # Also process slide layouts and masters
        for xml_file in (temp_dir / "ppt").glob("**/*.xml"):
            if "slideLayout" in str(xml_file) or "slideMaster" in str(xml_file):
                xml = xml_file.read_text(encoding="utf-8")
                colors = extract_all_colors(xml)
                replacements = build_replacement_map(colors, theme)
                if replacements:
                    xml, count = replace_colors_in_xml(xml, replacements)
                    if xml != original_xml if 'original_xml' in dir() else True:
                        xml_file.write_text(xml, encoding="utf-8")
                        total_replacements += count
        
        # Repack PPTX
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for file_path in temp_dir.rglob("*"):
                if file_path.is_file():
                    arcname = file_path.relative_to(temp_dir)
                    zf.write(file_path, arcname)
        
        print(f"\nDone! Processed {slides_processed} slides, {total_replacements} color replacements")
        print(f"Output: {output_path}")
        
    finally:
        # Clean up temp directory
        shutil.rmtree(temp_dir, ignore_errors=True)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="AI-powered color replacement for PPTX files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python scripts/color_ppt.py input.pptx output.pptx --theme tech
  python scripts/color_ppt.py input.pptx output.pptx --theme warm --verbose

Available themes:
  executive - Deep blue professional (default)
  tech      - Modern purple gradient
  creative  - Vibrant coral and teal
  warm      - Warm terracotta and sage
  minimal   - Clean gray tones
  bold      - Energetic red and blue
  nature    - Fresh green and blue
  ocean     - Deep blue and coral
  elegant   - Sophisticated dark tones
  modern    - Trendy pink and green
  sunset    - Warm orange and purple
  forest    - Natural green and brown
        """
    )
    
    parser.add_argument("input", help="Input PPTX file")
    parser.add_argument("output", help="Output PPTX file")
    parser.add_argument(
        "--theme", "-t",
        default="executive",
        help=f"Color theme to apply (default: executive)"
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Show detailed progress"
    )
    
    args = parser.parse_args()
    
    if not Path(args.input).exists():
        print(f"Error: Input file '{args.input}' not found")
        sys.exit(1)
    
    process_pptx(args.input, args.output, args.theme, args.verbose)


if __name__ == "__main__":
    main()
