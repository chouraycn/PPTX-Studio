"""Color Ladder API - Unified AI Color Ladder Interface for PPTX Studio

Provides unified interface for AI-powered color ladder integration across all modes:
- Mode 1: Template Apply
- Mode 2: Style Beautify
- Mode 6: Merge PPT

Features:
1. Theme-based preset ladders (12 themes)
2. Brand color ladder generation
3. Unified ladder application API
4. Strategy selection (lightness/saturation/complementary)
"""

import sys
from pathlib import Path
from typing import Dict, Optional, Tuple

# Add scripts dir to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from color_replacement import generate_ai_ladder, apply_color_ladder


# ─────────────────────────────────────────────────────────────────────────────
# THEME LADDER CONFIGURATIONS
# ─────────────────────────────────────────────────────────────────────────────

THEME_LADDERS = {
    "executive": {
        "name": "Executive Ladder",
        "base_color": "1E2761",  # Navy
        "depth": 5,
        "strategy": "lightness",
        "description": "Professional navy-based gradient, ideal for business and finance"
    },
    "tech": {
        "name": "Tech Ladder",
        "base_color": "028090",  # Teal
        "depth": 5,
        "strategy": "lightness",
        "description": "Modern teal gradient, perfect for technology and startups"
    },
    "creative": {
        "name": "Creative Ladder",
        "base_color": "F96167",  # Coral
        "depth": 5,
        "strategy": "lightness",
        "description": "Warm coral gradient, great for design and creative agencies"
    },
    "warm": {
        "name": "Warm Ladder",
        "base_color": "B85042",  # Terracotta
        "depth": 5,
        "strategy": "lightness",
        "description": "Earthy terracotta gradient, suitable for education and nonprofits"
    },
    "minimal": {
        "name": "Minimal Ladder",
        "base_color": "36454F",  # Charcoal
        "depth": 5,
        "strategy": "lightness",
        "description": "Clean charcoal gradient, ideal for academic and minimalist designs"
    },
    "bold": {
        "name": "Bold Ladder",
        "base_color": "990011",  # Cherry
        "depth": 5,
        "strategy": "lightness",
        "description": "Impactful cherry gradient, designed for high-energy presentations"
    },
    "nature": {
        "name": "Nature Ladder",
        "base_color": "2C5F2D",  # Forest
        "depth": 5,
        "strategy": "lightness",
        "description": "Natural forest gradient, perfect for environmental and health topics"
    },
    "ocean": {
        "name": "Ocean Ladder",
        "base_color": "0B3D91",  # Deep ocean
        "depth": 5,
        "strategy": "lightness",
        "description": "Calm ocean gradient, great for travel and maritime themes"
    },
    "elegant": {
        "name": "Elegant Ladder",
        "base_color": "4A5A6A",  # Slate blue
        "depth": 5,
        "strategy": "lightness",
        "description": "Sophisticated slate gradient, ideal for luxury and high-end brands"
    },
    "modern": {
        "name": "Modern Ladder",
        "base_color": "8B5CF6",  # Violet
        "depth": 5,
        "strategy": "lightness",
        "description": "Trendy violet gradient, perfect for internet and fashion"
    },
    "sunset": {
        "name": "Sunset Ladder",
        "base_color": "F97316",  # Orange
        "depth": 5,
        "strategy": "lightness",
        "description": "Warm sunset gradient, great for energy and hospitality"
    },
    "forest": {
        "name": "Forest Ladder",
        "base_color": "2D6A4F",  # Deep forest
        "depth": 5,
        "strategy": "lightness",
        "description": "Rich forest gradient, ideal for sustainability and organic products"
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# CORE FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def get_theme_ladder(
    theme_name: str,
    depth: Optional[int] = None,
    strategy: Optional[str] = None
) -> Dict[str, str]:
    """Generate color ladder from theme preset.
    
    Args:
        theme_name: Theme name (e.g., "tech", "executive")
        depth: Custom depth (3-10), defaults to theme preset
        strategy: Custom strategy ("lightness"/"saturation"/"complementary"), 
                 defaults to theme preset
    
    Returns:
        Dict mapping ladder levels to hex colors:
        {
            "level_0": "darkest_color",
            "level_1": "darker_color",
            "level_2": "middle_color",
            "level_3": "lighter_color",
            "level_4": "lightest_color"
        }
    """
    if theme_name not in THEME_LADDERS:
        print(f"Warning: Unknown theme '{theme_name}', falling back to 'minimal'")
        theme_name = "minimal"
    
    config = THEME_LADDERS[theme_name]
    base_color = config["base_color"]
    final_depth = depth or config["depth"]
    final_strategy = strategy or config["strategy"]
    
    # Generate ladder
    ladder_colors = generate_ai_ladder(base_color, final_depth, final_strategy)
    
    # Map to level keys
    ladder_dict = {}
    for i, color_info in enumerate(ladder_colors):
        ladder_dict[f"level_{i}"] = color_info["hex"]
    
    return ladder_dict


def apply_theme_ladder(
    pptx_path: str,
    theme_name: str,
    output_path: str,
    depth: Optional[int] = None,
    strategy: Optional[str] = None,
    preview: bool = False,
    verbose: bool = False
) -> str:
    """Apply theme preset color ladder to PPTX.
    
    Args:
        pptx_path: Input PPTX file path
        theme_name: Theme name (e.g., "tech", "executive")
        output_path: Output PPTX file path
        depth: Custom ladder depth (3-10)
        strategy: Custom ladder strategy
        preview: Preview mode (dry-run)
        verbose: Show detailed statistics
    
    Returns:
        Path to output file (or original if preview mode)
    """
    # Generate theme ladder
    ladder_dict = get_theme_ladder(theme_name, depth, strategy)
    
    # Build color map from theme's traditional colors to ladder
    # This ensures all existing colors are mapped to appropriate ladder levels
    import tempfile
    import json
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        # Map theme colors to ladder levels
        color_map = {
            "1E2761": ladder_dict["level_2"] if theme_name == "executive" else ladder_dict.get("level_0"),
            "028090": ladder_dict["level_2"] if theme_name == "tech" else ladder_dict.get("level_0"),
            "F96167": ladder_dict["level_2"] if theme_name == "creative" else ladder_dict.get("level_0"),
            "B85042": ladder_dict["level_2"] if theme_name == "warm" else ladder_dict.get("level_0"),
            "36454F": ladder_dict["level_2"] if theme_name == "minimal" else ladder_dict.get("level_0"),
            "990011": ladder_dict["level_2"] if theme_name == "bold" else ladder_dict.get("level_0"),
            "2C5F2D": ladder_dict["level_2"] if theme_name == "nature" else ladder_dict.get("level_0"),
            "0B3D91": ladder_dict["level_2"] if theme_name == "ocean" else ladder_dict.get("level_0"),
            "4A5A6A": ladder_dict["level_2"] if theme_name == "elegant" else ladder_dict.get("level_0"),
            "8B5CF6": ladder_dict["level_2"] if theme_name == "modern" else ladder_dict.get("level_0"),
            "F97316": ladder_dict["level_2"] if theme_name == "sunset" else ladder_dict.get("level_0"),
            "2D6A4F": ladder_dict["level_2"] if theme_name == "forest" else ladder_dict.get("level_0"),
        }
        json.dump(color_map, f)
        temp_file = f.name
    
    try:
        # Apply color map using color_replacement.py
        from color_replacement import main as color_replacement_main
        
        # Build command args
        args = [
            pptx_path,
            output_path,
            "--color-map-file", temp_file
        ]
        if preview:
            args.append("--preview")
        if verbose:
            args.append("--verbose")
        
        # Parse args
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("input_file")
        parser.add_argument("output_file")
        parser.add_argument("--color-map-file")
        parser.add_argument("--preview", action="store_true")
        parser.add_argument("--verbose", "-v", action="store_true")
        parsed_args = parser.parse_args(args)
        
        # Execute
        return color_replacement_main(parsed_args)
    finally:
        import os
        if os.path.exists(temp_file):
            os.remove(temp_file)


def apply_brand_ladder(
    pptx_path: str,
    brand_color: str,
    output_path: str,
    depth: int = 5,
    strategy: str = "lightness",
    preview: bool = False,
    verbose: bool = False
) -> str:
    """Apply brand color ladder to PPTX.
    
    Args:
        pptx_path: Input PPTX file path
        brand_color: Brand hex color (e.g., "0066CC")
        output_path: Output PPTX file path
        depth: Ladder depth (3-10, default 5)
        strategy: Ladder strategy ("lightness"/"saturation"/"complementary")
        preview: Preview mode (dry-run)
        verbose: Show detailed statistics
    
    Returns:
        Path to output file (or original if preview mode)
    """
    # Generate brand ladder
    ladder_colors = generate_ai_ladder(brand_color, depth, strategy)
    
    # Apply ladder to PPTX
    return apply_color_ladder(
        pptx_path,
        ladder_colors,
        output_path,
        preview=preview,
        verbose=verbose
    )


def auto_detect_primary_color(pptx_path: str) -> str:
    """Auto-detect primary color from PPTX.
    
    Extracts all colors from PPTX and returns the most frequently used color.
    
    Args:
        pptx_path: Input PPTX file path
    
    Returns:
        Most frequent hex color
    """
    from color_replacement import _extract_colors_from_pptx
    
    # Extract colors
    with tempfile.TemporaryDirectory() as temp_dir:
        color_counts = _extract_colors_from_pptx(pptx_path, temp_dir)
    
    # Find most frequent color (excluding white, black, gray)
    exclude_colors = {"FFFFFF", "000000", "808080", "C0C0C0", "00000000"}
    sorted_colors = sorted(
        [(color, count) for color, count in color_counts.items() 
         if color not in exclude_colors and len(color) == 6],
        key=lambda x: x[1],
        reverse=True
    )
    
    if sorted_colors:
        return sorted_colors[0][0]
    else:
        return "36454F"  # Default to charcoal


def list_theme_ladders() -> None:
    """List all available theme ladders with descriptions."""
    print("Available Theme Ladders:")
    print("=" * 60)
    for theme_name, config in THEME_LADDERS.items():
        print(f"\n📋 {theme_name.title()}")
        print(f"   Base Color: #{config['base_color']}")
        print(f"   Depth: {config['depth']} levels")
        print(f"   Strategy: {config['strategy']}")
        print(f"   Description: {config['description']}")


# ─────────────────────────────────────────────────────────────────────────────
# EXPORTS FOR OTHER MODULES
# ─────────────────────────────────────────────────────────────────────────────

__all__ = [
    "get_theme_ladder",
    "apply_theme_ladder",
    "apply_brand_ladder",
    "auto_detect_primary_color",
    "list_theme_ladders",
    "THEME_LADDERS",
]
