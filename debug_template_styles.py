#!/usr/bin/env python3
"""Debug script for apply_template.py style application"""
import sys
from pathlib import Path

# Add scripts directory to path
sys.path.insert(0, str(Path(__file__).parent / "scripts"))

from apply_template import (
    _extract_template_colors,
    _extract_template_fonts,
    _analyze_template_layouts,
    _extract_layout_placeholder_styles,
    _detect_lang,
)
from extract_content import extract_content

def debug_template(template_path: str):
    """Debug template extraction"""
    print("=" * 80)
    print(f"DEBUG: Analyzing template {template_path}")
    print("=" * 80)

    # 1. Extract colors
    print("\n1. Extracting template colors...")
    try:
        template_colors = _extract_template_colors(template_path)
        print(f"   ✓ Primary color: {template_colors.get('primary', 'NOT FOUND')}")
        print(f"   ✓ Text on light: {template_colors.get('text_on_light', 'NOT FOUND')}")
        print(f"   ✓ Text on dark: {template_colors.get('text_on_dark', 'NOT FOUND')}")
        print(f"   ✓ Total colors extracted: {len(template_colors)}")
    except Exception as e:
        print(f"   ✗ Error extracting colors: {e}")
        template_colors = {}

    # 2. Extract fonts
    print("\n2. Extracting template fonts...")
    try:
        template_fonts = _extract_template_fonts(template_path)
        print(f"   ✓ Major Latin: {template_fonts.get('major_latin', 'NOT FOUND')}")
        print(f"   ✓ Minor Latin: {template_fonts.get('minor_latin', 'NOT FOUND')}")
        print(f"   ✓ Major EA: {template_fonts.get('major_ea', 'NOT FOUND')}")
        print(f"   ✓ Minor EA: {template_fonts.get('minor_ea', 'NOT FOUND')}")
        print(f"   ✓ Total fonts extracted: {len(template_fonts)}")
    except Exception as e:
        print(f"   ✗ Error extracting fonts: {e}")
        template_fonts = {}

    # 3. Analyze layouts
    print("\n3. Analyzing template layouts...")
    try:
        template_layouts = _analyze_template_layouts(template_path)
        print(f"   ✓ Total layouts: {len(template_layouts)}")
        print(f"   Layout files: {list(template_layouts.keys())[:5]}")
    except Exception as e:
        print(f"   ✗ Error analyzing layouts: {e}")
        template_layouts = {}

    # 4. Extract placeholder styles for first layout
    if template_layouts:
        first_layout = list(template_layouts.keys())[0]
        print(f"\n4. Analyzing placeholder styles in layout: {first_layout}")
        try:
            ph_styles = _extract_layout_placeholder_styles(template_path, first_layout)
            print(f"   ✓ Placeholders with styles: {len(ph_styles)}")
            for ph_idx, ph_style in list(ph_styles.items())[:3]:
                print(f"\n      Placeholder {ph_idx}:")
                if ph_style.get('bodyPr_attrs'):
                    print(f"        Has bodyPr attrs: {len(ph_style['bodyPr_attrs'])} attributes")
                else:
                    print(f"        No bodyPr attrs")
                if ph_style.get('defPPr_xml'):
                    print(f"        Has defPPr XML: {len(ph_style['defPPr_xml'])} chars")
                else:
                    print(f"        No defPPr XML")
                if ph_style.get('defRPr_xml'):
                    print(f"        Has defRPr XML: {len(ph_style['defRPr_xml'])} chars")
                else:
                    print(f"        No defRPr XML")
                if ph_style.get('default_sz'):
                    print(f"        Default size: {ph_style['default_sz']}")
                else:
                    print(f"        No default size")
        except Exception as e:
            print(f"   ✗ Error extracting placeholder styles: {e}")

    return template_colors, template_fonts, template_layouts

def debug_source(source_path: str):
    """Debug source content extraction"""
    print("\n" + "=" * 80)
    print(f"DEBUG: Analyzing source {source_path}")
    print("=" * 80)

    try:
        # Extract content
        content = extract_content(source_path)

        print(f"\n1. Total slides: {len(content.get('slides', []))}")
        print(f"   Detected theme: {content.get('detected_theme', 'UNKNOWN')}")

        # Show first 2 slides
        for i, slide in enumerate(content.get('slides', [])[:2]):
            print(f"\n   Slide {i+1}:")
            title = slide.get('title', 'NO TITLE')
            print(f"      Title: {title[:60]}{'...' if len(title) > 60 else ''}")
            print(f"      Layout file: {slide.get('layout_file', 'UNKNOWN')}")

            body_rich = slide.get('body_rich', [])
            print(f"      Body lines (rich format): {len(body_rich)}")

            # Show body_rich format info
            for j, line in enumerate(body_rich[:3]):
                print(f"\n      Line {j+1}: {line.get('text', '')[:40]}{'...' if len(line.get('text', '')) > 40 else ''}")
                print(f"        Has bold: {bool(line.get('bold'))}")
                print(f"        Has italic: {bool(line.get('italic'))}")
                print(f"        Has size: {bool(line.get('size'))} ({line.get('size') if line.get('size') else 'not set'})")
                print(f"        Has color: {bool(line.get('color'))}")
                print(f"        Language detected: {_detect_lang(line.get('text', ''))}")

        return content
    except Exception as e:
        print(f"\n✗ Error extracting source content: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Debug apply_template style application")
    parser.add_argument('template', help='Path to template PPTX')
    parser.add_argument('source', help='Path to source PPTX', nargs='?', default='')
    args = parser.parse_args()

    if not Path(args.template).exists():
        print(f"Error: Template file not found: {args.template}")
        sys.exit(1)

    template_colors, template_fonts, template_layouts = debug_template(args.template)

    if args.source:
        if not Path(args.source).exists():
            print(f"Error: Source file not found: {args.source}")
            sys.exit(1)
        debug_source(args.source)
