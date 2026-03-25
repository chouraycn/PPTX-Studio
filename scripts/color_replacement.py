"""Global PPTX Color Replacement & AI Color Ladder

Features:
- Replace all colors in PPT (text, shapes, backgrounds, gradients, fills)
- AI-powered color ladder generation (5-level depth gradients)
- Intelligent color harmony preservation
- Support for custom color targets and theme-based replacements
- Preview mode before applying changes

Usage:
    # Replace primary color (e.g., orange → blue)
    python scripts/color_replacement.py input.pptx output.pptx --replace-primary F96167 0284C7
    
    # Replace full color palette with AI-generated ladder
    python scripts/color_replacement.py input.pptx output.pptx --ai-ladder F96167 --depth 5
    
    # Replace based on theme
    python scripts/color_replacement.py input.pptx output.pptx --theme-from warm --theme-to tech
    
    # Preview changes (dry-run)
    python scripts/color_replacement.py input.pptx output.pptx --replace-primary F96167 0284C7 --preview
"""

import argparse
import json
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Optional, Dict, List, Tuple
import subprocess

# Add scripts dir to path for imports (Python 3.9 compatible)
sys.path.insert(0, str(Path(__file__).parent))

from extract_content import extract_content


def _run_unpack(input_file: str, output_dir: str) -> None:
    """Run unpack.py as a subprocess."""
    scripts_dir = Path(__file__).parent
    result = subprocess.run(
        [sys.executable, str(scripts_dir / "office" / "unpack.py"), input_file, output_dir],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"Unpack error: {result.stderr}", file=sys.stderr)
        sys.exit(1)


def _run_pack(input_dir: str, output_file: str, original: Optional[str] = None) -> None:
    """Run pack.py as a subprocess."""
    scripts_dir = Path(__file__).parent
    cmd = [sys.executable, str(scripts_dir / "office" / "pack.py"),
           input_dir, output_file, "--validate", "false"]
    if original:
        cmd += ["--original", original]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Pack error: {result.stderr}", file=sys.stderr)
        sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# AI COLOR LADDER GENERATION
# ─────────────────────────────────────────────────────────────────────────────

def _hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """Convert hex color to RGB."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb: Tuple[int, int, int]) -> str:
    """Convert RGB to hex color."""
    return ''.join(f'{c:02X}' for c in rgb)


def _hsv_to_rgb(h: float, s: float, v: float) -> Tuple[int, int, int]:
    """Convert HSV to RGB."""
    c = v * s
    x = c * (1 - abs((h / 60) % 2 - 1))
    m = v - c
    
    if 0 <= h < 60:
        r, g, b = c, x, 0
    elif 60 <= h < 120:
        r, g, b = x, c, 0
    elif 120 <= h < 180:
        r, g, b = 0, c, x
    elif 180 <= h < 240:
        r, g, b = 0, x, c
    elif 240 <= h < 300:
        r, g, b = x, 0, c
    else:
        r, g, b = c, 0, x
    
    return (int((r + m) * 255), int((g + m) * 255), int((b + m) * 255))


def _rgb_to_hsv(rgb: Tuple[int, int, int]) -> Tuple[float, float, float]:
    """Convert RGB to HSV."""
    r, g, b = [c / 255 for c in rgb]
    cmax = max(r, g, b)
    cmin = min(r, g, b)
    delta = cmax - cmin
    
    if delta == 0:
        h = 0
    elif cmax == r:
        h = 60 * (((g - b) / delta) % 6)
    elif cmax == g:
        h = 60 * (((b - r) / delta) + 2)
    else:
        h = 60 * (((r - g) / delta) + 4)
    
    if cmax == 0:
        s = 0
    else:
        s = delta / cmax
    
    return (h % 360, s, cmax)


def _generate_color_ladder(
    base_color: str,
    depth: int = 5,
    strategy: str = "lightness"
) -> List[Dict[str, any]]:
    """Generate AI-powered color ladder (multi-level gradients).
    
    Args:
        base_color: Starting hex color (e.g., "F96167")
        depth: Number of levels in ladder (3-10)
        strategy: "lightness" (dark→light), "saturation" (dull→vivid), "complementary"
    
    Returns:
        List of color dicts with hex, h, s, v values and usage hints
    """
    rgb = _hex_to_rgb(base_color)
    h, s, v = _rgb_to_hsv(rgb)
    
    ladder = []
    
    if strategy == "lightness":
        # Generate gradient from dark to light
        for i in range(depth):
            # Lightness varies from 0.2 to 0.95
            new_v = 0.2 + (0.75 * i / (depth - 1))
            new_rgb = _hsv_to_rgb(h, s, new_v)
            new_hex = _rgb_to_hex(new_rgb)
            
            # Usage hint based on lightness
            if i == 0:
                usage = "darkest - text on light backgrounds, dark elements"
            elif i == depth - 1:
                usage = "lightest - text on dark backgrounds, highlights"
            elif i < depth // 2:
                usage = "darker - secondary elements, accents"
            else:
                usage = "lighter - tertiary elements, backgrounds"
            
            ladder.append({
                "index": i,
                "hex": new_hex,
                "h": round(h, 2),
                "s": round(s, 2),
                "v": round(new_v, 2),
                "usage": usage
            })
    
    elif strategy == "saturation":
        # Generate gradient from dull to vivid
        for i in range(depth):
            # Saturation varies from 0.1 to 1.0
            new_s = 0.1 + (0.9 * i / (depth - 1))
            new_rgb = _hsv_to_rgb(h, new_s, v)
            new_hex = _rgb_to_hex(new_rgb)
            
            usage = "muted" if i < depth // 2 else "vivid"
            ladder.append({
                "index": i,
                "hex": new_hex,
                "h": round(h, 2),
                "s": round(new_s, 2),
                "v": round(v, 2),
                "usage": f"{usage} - accent elements"
            })
    
    elif strategy == "complementary":
        # Generate ladder crossing to complementary color
        h_comp = (h + 180) % 360
        for i in range(depth):
            # Hue interpolates from base to complementary
            new_h = h + (180 * i / (depth - 1))
            new_rgb = _hsv_to_rgb(new_h % 360, s, v)
            new_hex = _rgb_to_hex(new_rgb)
            
            if i == 0:
                usage = "base color"
            elif i == depth - 1:
                usage = "complementary color"
            else:
                usage = "transition - gradients, harmony"
            
            ladder.append({
                "index": i,
                "hex": new_hex,
                "h": round(new_h % 360, 2),
                "s": round(s, 2),
                "v": round(v, 2),
                "usage": usage
            })
    
    else:
        raise ValueError(f"Unknown strategy: {strategy}")
    
    return ladder


def _calculate_color_distance(color1: str, color2: str) -> float:
    """Calculate Euclidean distance between two colors in RGB space."""
    rgb1 = _hex_to_rgb(color1)
    rgb2 = _hex_to_rgb(color2)
    return sum((a - b) ** 2 for a, b in zip(rgb1, rgb2)) ** 0.5


def _find_closest_theme_color(
    target_color: str,
    theme_colors: Dict[str, str]
) -> Tuple[str, float]:
    """Find the closest theme color to a target color."""
    closest_name = None
    min_distance = float('inf')
    
    for name, color in theme_colors.items():
        distance = _calculate_color_distance(target_color, color)
        if distance < min_distance:
            min_distance = distance
            closest_name = name
    
    return closest_name, min_distance


# ─────────────────────────────────────────────────────────────────────────────
# THEME DEFINITIONS (for --theme-from/--theme-to)
# ─────────────────────────────────────────────────────────────────────────────

THEME_PALETTES = {
    "executive": {
        "primary": "1E2761",
        "secondary": "CADCFC",
        "accent": "C9A84C",
    },
    "tech": {
        "primary": "028090",
        "secondary": "1C2541",
        "accent": "02C39A",
    },
    "creative": {
        "primary": "F96167",
        "secondary": "2F3C7E",
        "accent": "F9E795",
    },
    "warm": {
        "primary": "B85042",
        "secondary": "84B59F",
        "accent": "ECE2D0",
    },
    "minimal": {
        "primary": "374151",
        "secondary": "9CA3AF",
        "accent": "FFFFFF",
    },
    "bold": {
        "primary": "FF5757",
        "secondary": "1E2761",
        "accent": "FFD700",
    },
    "nature": {
        "primary": "2D6A4F",
        "secondary": "F9A826",
        "accent": "B4E4FF",
    },
    "ocean": {
        "primary": "0077B6",
        "secondary": "00B4D8",
        "accent": "CAF0F8",
    },
    "elegant": {
        "primary": "1F2937",
        "secondary": "737373",
        "accent": "FF7F50",
    },
    "modern": {
        "primary": "7C3AED",
        "secondary": "DB2777",
        "accent": "F472B6",
    },
    "sunset": {
        "primary": "F97316",
        "secondary": "FBBF24",
        "accent": "FCD34D",
    },
    "forest": {
        "primary": "065F46",
        "secondary": "059669",
        "accent": "10B981",
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# COLOR REPLACEMENT ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def _replace_color_in_xml(
    xml_content: str,
    color_map: Dict[str, str],
    preview_mode: bool = False,
    verbose: bool = False
) -> Tuple[str, Dict[str, int]]:
    """Replace colors in XML content.
    
    Args:
        xml_content: XML string
        color_map: Dict mapping old colors to new colors
        preview_mode: If True, don't modify, just count
        verbose: Print detailed stats
    
    Returns:
        Tuple of (modified_xml, stats_dict)
    """
    stats = {
        "total_replacements": 0,
        "color_counts": {},
    }
    
    for old_color, new_color in color_map.items():
        # Pattern matches hex colors in various formats:
        # srgbVal="F96167", schemeClr val="F96167", fill="F96167", etc.
        pattern = re.compile(
            rf'(srgbVal=|schemeClr val=|fill=|stroke=|color=|bgClr=|fgClr=|a:srgbClr val=)["\']?([A-Fa-f0-9]{{6}})["\']?',
            re.IGNORECASE
        )
        
        count = len(pattern.findall(xml_content))
        if count > 0:
            stats["total_replacements"] += count
            stats["color_counts"][old_color] = count
            
            if verbose and not preview_mode:
                print(f"    {old_color} → {new_color}: {count} occurrences")
    
    if preview_mode:
        return xml_content, stats
    
    # Perform replacements
    # PPTX XML format: <a:srgbClr val="RRGGBB"/> or val='RRGGBB'
    # Also handles: <a:srgbClr val="RRGGBB" other="..."/>
    # Pattern: quotes around the color value, preceded by attr name like val=
    for old_color, new_color in color_map.items():
        # Match color value inside quotes, with flexible whitespace
        pattern = rf'val\s*=\s*["\']?({re.escape(old_color)})["\']]?(?=\s|/>|>)'
        def make_repl(nc=new_color):
            def replacer(m):
                # Replace the hex color, preserve surrounding context
                return 'val="' + nc + '"'
            return replacer
        xml_content = re.sub(pattern, make_repl(new_color), xml_content, flags=re.IGNORECASE)
    
    return xml_content, stats


def _extract_colors_from_pptx(unpacked_dir: Path) -> Dict[str, int]:
    """Extract all unique colors and their counts from PPTX.
    
    Returns:
        Dict mapping hex colors to occurrence counts
    """
    color_counts = {}
    
    # Search for colors in all XML files
    for xml_file in unpacked_dir.rglob("*.xml"):
        content = xml_file.read_text(encoding="utf-8", errors="ignore")
        
        # Find all hex colors
        colors = re.findall(r'(?i)(?:srgbVal|schemeClr val|fill|stroke|color|bgClr|fgClr)\s*=\s*["\']?([A-Fa-f0-9]{6})["\']?', content)
        
        for color in colors:
            upper_color = color.upper()
            color_counts[upper_color] = color_counts.get(upper_color, 0) + 1
    
    return color_counts


def color_replacement(
    input_pptx: str,
    output_pptx: str,
    replace_primary: Optional[Tuple[str, str]] = None,
    replace_secondary: Optional[Tuple[str, str]] = None,
    replace_accent: Optional[Tuple[str, str]] = None,
    ai_ladder: Optional[str] = None,
    ladder_depth: int = 5,
    ladder_strategy: str = "lightness",
    theme_from: Optional[str] = None,
    theme_to: Optional[str] = None,
    color_map_file: Optional[str] = None,
    preview: bool = False,
    verbose: bool = False,
) -> None:
    """Main color replacement function."""
    
    # Unpack PPTX
    print(f"📦 Unpacking {input_pptx}...")
    with tempfile.TemporaryDirectory() as temp_dir:
        unpacked_dir = Path(temp_dir) / "unpacked"
        unpacked_dir.mkdir(parents=True, exist_ok=True)
        _run_unpack(input_pptx, str(unpacked_dir))
        
        # Extract existing colors
        print("\n🔍 Analyzing existing colors...")
        color_counts = _extract_colors_from_pptx(unpacked_dir)
        total_colors = sum(color_counts.values())
        
        if verbose:
            print(f"  Found {len(color_counts)} unique colors ({total_colors} total occurrences)")
            print("  Top 10 most used colors:")
            for color, count in sorted(color_counts.items(), key=lambda x: -x[1])[:10]:
                print(f"    {color}: {count} occurrences")
        
        # Build color replacement map
        color_map = {}
        
        if theme_from and theme_to:
            # Replace based on theme
            print(f"\n🎨 Replacing {theme_from} theme with {theme_to} theme...")
            from_palette = THEME_PALETTES.get(theme_from)
            to_palette = THEME_PALETTES.get(theme_to)
            
            if not from_palette or not to_palette:
                print(f"Error: Unknown theme(s)", file=sys.stderr)
                print(f"  Available themes: {', '.join(THEME_PALETTES.keys())}")
                sys.exit(1)
            
            # Map each color in from_palette to to_palette
            for key in from_palette:
                old_color = from_palette[key]
                new_color = to_palette.get(key, old_color)
                color_map[old_color] = new_color
        
        elif ai_ladder:
            # Generate AI color ladder and replace based on distance
            print(f"\n🤖 Generating AI color ladder (depth={ladder_depth}, strategy={ladder_strategy})...")
            ladder = _generate_color_ladder(ai_ladder, ladder_depth, ladder_strategy)
            
            if verbose:
                print(f"  Generated {len(ladder)} color levels:")
                for level in ladder:
                    print(f"    {level['hex']} - {level['usage']}")
            
            # Replace existing colors based on closest ladder color
            print(f"\n🔄 Mapping existing colors to ladder...")
            for old_color in color_counts.keys():
                closest_level = min(ladder, key=lambda x: _calculate_color_distance(old_color, x["hex"]))
                new_color = closest_level["hex"]
                color_map[old_color] = new_color
        
        elif color_map_file:
            # Load custom color map from JSON
            print(f"\n📋 Loading color map from {color_map_file}...")
            with open(color_map_file, 'r', encoding='utf-8') as f:
                loaded_map = json.load(f)
                color_map.update(loaded_map)
        
        else:
            # Single color replacements
            if replace_primary:
                old_color, new_color = replace_primary
                print(f"\n🎯 Replacing primary color: {old_color} → {new_color}")
                color_map[old_color] = new_color
            
            if replace_secondary:
                old_color, new_color = replace_secondary
                print(f"🎯 Replacing secondary color: {old_color} → {new_color}")
                color_map[old_color] = new_color
            
            if replace_accent:
                old_color, new_color = replace_accent
                print(f"🎯 Replacing accent color: {old_color} → {new_color}")
                color_map[old_color] = new_color
        
        if not color_map:
            print("Error: No color replacements specified", file=sys.stderr)
            print("  Use --replace-primary, --replace-secondary, --replace-accent,")
            print("  --ai-ladder, --theme-from/--theme-to, or --color-map-file")
            sys.exit(1)
        
        if preview:
            print("\n📊 PREVIEW MODE - No changes will be made")
            print("=" * 60)
            
            # Preview replacements
            total_replacements = 0
            for old_color, new_color in color_map.items():
                count = color_counts.get(old_color, 0)
                if count > 0:
                    total_replacements += count
                    print(f"  {old_color} → {new_color}: {count} occurrences")
            
            print("=" * 60)
            print(f"Total replacements: {total_replacements}")
            return
        
        # Apply color replacements
        print("\n🔄 Applying color replacements...")
        
        xml_files = list(unpacked_dir.rglob("*.xml"))
        total_xml_replacements = 0
        
        for xml_file in xml_files:
            content = xml_file.read_text(encoding="utf-8", errors="ignore")
            
            modified_content, stats = _replace_color_in_xml(
                content, color_map, preview_mode=False, verbose=False
            )
            
            if stats["total_replacements"] > 0:
                total_xml_replacements += stats["total_replacements"]
                xml_file.write_text(modified_content, encoding="utf-8")
        
        if verbose:
            print(f"  Modified {len(xml_files)} XML files")
        
        # Pack output
        print(f"\n📦 Packing {output_pptx}...")
        _run_pack(str(unpacked_dir), output_pptx, input_pptx)
    
    print(f"\n✅ Color replacement complete!")
    print(f"   Total replacements: {total_xml_replacements}")


def main():
    parser = argparse.ArgumentParser(
        description="Global PPTX Color Replacement & AI Color Ladder",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Replace primary color (e.g., orange → blue)
  python scripts/color_replacement.py input.pptx output.pptx --replace-primary F96167 0284C7
  
  # Generate AI color ladder and replace all colors
  python scripts/color_replacement.py input.pptx output.pptx --ai-ladder F96167 --depth 5
  
  # Replace based on theme
  python scripts/color_replacement.py input.pptx output.pptx --theme-from warm --theme-to tech
  
  # Preview changes before applying
  python scripts/color_replacement.py input.pptx output.pptx --replace-primary F96167 0284C7 --preview
  
  # Use custom color map from JSON file
  python scripts/color_replacement.py input.pptx output.pptx --color-map-file my_colors.json
        """
    )
    
    parser.add_argument("input", help="Input PPTX file")
    parser.add_argument("output", help="Output PPTX file")
    
    # Single color replacements
    parser.add_argument(
        "--replace-primary",
        nargs=2,
        metavar=("OLD_COLOR", "NEW_COLOR"),
        help="Replace primary color (e.g., F96167 0284C7)"
    )
    parser.add_argument(
        "--replace-secondary",
        nargs=2,
        metavar=("OLD_COLOR", "NEW_COLOR"),
        help="Replace secondary color"
    )
    parser.add_argument(
        "--replace-accent",
        nargs=2,
        metavar=("OLD_COLOR", "NEW_COLOR"),
        help="Replace accent color"
    )
    
    # AI ladder
    parser.add_argument(
        "--ai-ladder",
        metavar="BASE_COLOR",
        help="Generate AI-powered color ladder from base color and replace all colors"
    )
    parser.add_argument(
        "--ladder-depth",
        type=int,
        default=5,
        help="Number of levels in color ladder (default: 5, range: 3-10)"
    )
    parser.add_argument(
        "--ladder-strategy",
        choices=["lightness", "saturation", "complementary"],
        default="lightness",
        help="Color ladder strategy (default: lightness)"
    )
    
    # Theme-based replacement
    parser.add_argument(
        "--theme-from",
        choices=list(THEME_PALETTES.keys()),
        help="Source theme name"
    )
    parser.add_argument(
        "--theme-to",
        choices=list(THEME_PALETTES.keys()),
        help="Target theme name"
    )
    
    # Custom color map
    parser.add_argument(
        "--color-map-file",
        metavar="FILE",
        help="JSON file with custom color mappings (format: {\"OLD\": \"NEW\"})"
    )
    
    # Options
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Preview mode: show what will be changed without applying"
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Show detailed processing information"
    )
    
    args = parser.parse_args()
    
    # Validate depth range
    if args.ai_ladder and (args.ladder_depth < 3 or args.ladder_depth > 10):
        print("Error: --ladder-depth must be between 3 and 10", file=sys.stderr)
        sys.exit(1)
    
    color_replacement(
        args.input,
        args.output,
        replace_primary=tuple(args.replace_primary) if args.replace_primary else None,
        replace_secondary=tuple(args.replace_secondary) if args.replace_secondary else None,
        replace_accent=tuple(args.replace_accent) if args.replace_accent else None,
        ai_ladder=args.ai_ladder,
        ladder_depth=args.ladder_depth,
        ladder_strategy=args.ladder_strategy,
        theme_from=args.theme_from,
        theme_to=args.theme_to,
        color_map_file=args.color_map_file,
        preview=args.preview,
        verbose=args.verbose,
    )


if __name__ == "__main__":
    main()
