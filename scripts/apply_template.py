"""Apply a template's visual style to a source presentation's content.

Extracts content from source.pptx and reflows it into the template's layouts.
The output preserves the template's visual identity (colors, fonts, backgrounds)
while using the source's actual text, images, and data.

Usage:
    python scripts/apply_template.py source.pptx template.pptx output.pptx
    python scripts/apply_template.py source.pptx template.pptx output.pptx --dry-run
    python scripts/apply_template.py source.pptx template.pptx output.pptx --verbose
    python scripts/apply_template.py source.pptx template.pptx output.pptx --mapping mapping.json

--dry-run:
    Print the slide mapping plan and exit. No output file is created.
    Use this to verify the auto-mapping before committing to execution.
    The mapping can be saved to a JSON file with --save-mapping and used
    with --mapping to override the auto-mapping on the next run.

How it works:
    1. Extract content from source PPT
    2. Analyze template slide layouts
    3. Auto-map each source slide to the best template layout
    4. [dry-run stops here and prints the plan]
    5. For each source slide:
       a. Duplicate the appropriate template slide
       b. Replace placeholder text with source content
       c. Copy images if the layout supports them
    6. Remove unused template slides
    7. Pack result

Note: This script handles common layout types. For complex slides (custom shapes,
      overlapping elements), manual XML editing may still be needed afterward.
"""

import argparse
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, List, Dict, Tuple

import subprocess

# Namespace definitions for PPTX XML
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
# Register namespaces to avoid ns0, ns1 prefixes
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)
# Also register common namespaces found in PPTX
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
ET.register_namespace("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram")

# Add scripts dir to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from extract_content import extract_content
from animation_migration import migrate_animations_with_id_mapping
from hyperlink_migration import extract_hyperlinks, inject_hyperlinks


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
    print(f"  {result.stdout.strip()}")


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
    print(f"  {result.stdout.strip()}")


# Layout type priority order for matching
LAYOUT_TYPE_KEYWORDS = {
    "title_slide": [
        "title slide", "title, subtitle", "title only", "ctrtitle",
        "首页", "封面", "标题幻灯片", "title",
    ],
    "section_header": [
        "section header", "section", "divider", "blank",
        "章节", "节标题", "过渡页", "section",
    ],
    "content_slide": [
        "title and content", "content", "object",
        "标题和内容", "内容", "正文",
    ],
    "two_column": [
        "two content", "comparison", "2 column", "two column",
        "两栏", "双栏", "对比", "比较",
    ],
    "image_text": [
        "picture with caption", "picture", "image", "photo",
        "图片", "图文", "图像",
    ],
    "list_content": [
        "title and content", "content", "bulleted list",
        "列表", "要点", "bullet",
    ],
    "chart_content": ["title and content", "content", "图表", "chart"],
    "table_content": ["title and content", "content", "表格", "table"],
    "quote_slide": ["blank", "title only", "quote", "引用", "金句"],
    "conclusion": ["blank", "title only", "title slide", "结语", "总结", "谢谢"],
    "full_image": ["blank", "picture", "全图", "大图"],
}


def apply_template(
    source_pptx: str,
    template_pptx: str,
    output_pptx: str,
    mapping_file: Optional[str] = None,
    save_mapping: Optional[str] = None,
    dry_run: bool = False,
    verbose: bool = False,
    keep_notes: bool = True,
    skip_animations: bool = False,
    interactive: bool = False,
    beautify: bool = False,
    beautify_theme: Optional[str] = None,
    dark_mode: bool = False,
) -> None:
    """Apply template to source PPT content."""

    source_path = Path(source_pptx)
    template_path = Path(template_pptx)
    output_path = Path(output_pptx)

    for p, name in [(source_path, "Source"), (template_path, "Template")]:
        if not p.exists():
            print(f"Error: {name} file not found: {p}", file=sys.stderr)
            sys.exit(1)

    print(f"Extracting content from source: {source_pptx}")
    source_content = extract_content(str(source_path), print_summary=verbose)

    if not source_content["slides"]:
        print("Error: No slides found in source file", file=sys.stderr)
        sys.exit(1)

    print(f"Source: {source_content['total_slides']} slides")
    print(f"Analyzing template: {template_pptx}")

    template_layouts = _analyze_template_layouts(template_pptx)
    # Extract template color palette for content injection
    template_colors = _extract_template_colors(template_pptx)
    # Extract template font scheme (majorFont / minorFont)
    template_fonts  = _extract_template_fonts(template_pptx)
    if verbose:
        print(f"Template colors: primary=#{template_colors['primary']}, "
              f"accent=#{template_colors['accent']}")
        print(f"Template fonts: major={template_fonts.get('major_latin','?')}, "
              f"minor={template_fonts.get('minor_latin','?')}, "
              f"major_ea={template_fonts.get('major_ea','')}, "
              f"minor_ea={template_fonts.get('minor_ea','')}")
        print(f"Template has {len(template_layouts)} layouts:")
        for l in template_layouts:
            print(f"  {l['layout_file']}: {l['layout_name']} ({l['detected_type']})")

    # Load or create slide mapping
    if mapping_file and Path(mapping_file).exists():
        with open(mapping_file) as f:
            slide_mapping = json.load(f)
        print(f"Using manual mapping from {mapping_file}")
    else:
        slide_mapping = _auto_map_slides(
            source_content["slides"], template_layouts, verbose, interactive
        )

    # ── DRY-RUN: print mapping plan and stop ─────────────────────────────────
    _print_mapping_plan(slide_mapping, source_content["slides"])

    if save_mapping:
        with open(save_mapping, "w", encoding="utf-8") as f:
            json.dump(slide_mapping, f, indent=2, ensure_ascii=False)
        print(f"\nMapping saved to: {save_mapping}")
        print("Edit this file and pass it with --mapping to override the auto-mapping.")

    if dry_run:
        print("\n[dry-run] No output file written. "
              "Re-run without --dry-run to execute the apply.")
        return
    # ─────────────────────────────────────────────────────────────────────────

    # Work in temp directory
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        unpacked_dir = tmp_path / "working"

        print(f"\nUnpacking template...")
        _run_unpack(str(template_path), str(unpacked_dir))

        # Ensure slideMasters and themes are preserved from template
        _ensure_slide_masters_preserved(unpacked_dir, verbose)

        # Get template slide structure
        template_slide_order = _get_presentation_slide_order(unpacked_dir)
        if verbose:
            print(f"Template has {len(template_slide_order)} slides in sldIdLst")

        # Extract source images to temp dir
        source_images_dir = tmp_path / "source_images"
        source_images_dir.mkdir()
        _extract_source_images(source_pptx, source_images_dir)

        # Unpack source PPT into a separate dir so we can read animations & notes
        source_unpacked_dir = tmp_path / "source"
        print(f"Unpacking source for animation/notes extraction...")
        _run_unpack(str(source_path), str(source_unpacked_dir))

        # Build slide-index → source slide file mapping (1-based index from extraction)
        source_slide_file_map = _build_source_slide_file_map(
            source_unpacked_dir, source_content["slides"]
        )

        # Build template theme (template colors + beautify design system)
        template_theme = _build_template_theme(template_colors, template_fonts, beautify_theme)
        if verbose and beautify:
            print(f"  Beautify theme: {template_theme['name']}")
            print(f"  Colors: primary=#{template_theme['primary']}, accent=#{template_theme['accent']}")

        # Build new slide list
        new_slides = _build_new_slides(
            source_content["slides"],
            slide_mapping,
            template_layouts,
            unpacked_dir,
            source_images_dir,
            keep_notes,
            template_colors,
            template_fonts,
            source_unpacked_dir,
            source_slide_file_map,
            verbose,
            skip_animations,
            beautify=beautify,
            template_theme=template_theme,
            dark_mode=dark_mode,
        )

        # Update presentation.xml with new slide order
        _update_presentation_order(unpacked_dir, new_slides)

        # Extract and migrate hyperlinks from source to target
        if verbose:
            print("\nMigrating hyperlinks...")
        source_hyperlinks = extract_hyperlinks(str(source_unpacked_dir))
        # Build source index -> target slide file mapping
        source_index_to_target = {}
        for idx, target_slide in enumerate(new_slides, 1):
            source_index_to_target[idx] = target_slide
        hyperlinks_injected = inject_hyperlinks(
            str(unpacked_dir),
            source_hyperlinks,
            source_index_to_target
        )
        if verbose and hyperlinks_injected > 0:
            print(f"  + Migrated {hyperlinks_injected} hyperlinks")

        # Clean up unreferenced files
        print("\nCleaning up...")
        sys.path.insert(0, str(Path(__file__).parent))
        from clean import clean_unused_files
        removed = clean_unused_files(unpacked_dir)
        if removed and verbose:
            print(f"  Removed {len(removed)} orphaned files")

        # Pack result
        print(f"\nPacking output to {output_pptx}...")
        _run_pack(str(unpacked_dir), str(output_path), original=str(template_path))

    # Note: Deep beautification (backgrounds, layout variants, fonts, images, tables)
    # is now applied per-slide during slide building when --beautify is enabled.
    # The template theme colors/fonts are preserved throughout.

    print(f"\nDone! Output saved to: {output_pptx}")
    print("Run QA check with:")
    print(f"  python scripts/qa_check.py {output_pptx}")
    print("Run visual QA with:")
    print(f"  python scripts/thumbnail.py {output_pptx}")


def _print_mapping_plan(slide_mapping: List[dict], source_slides: List[dict]) -> None:
    """Print a human-readable mapping plan."""
    # Build a lookup for source slide titles
    source_titles = {s["index"]: s.get("title", "") for s in source_slides}

    print("\n" + "─" * 60)
    print("  Slide Mapping Plan")
    print("─" * 60)
    print(f"  {'#':<4} {'Source Title':<30} {'Source Type':<16} {'→ Template Layout'}")
    print("─" * 60)
    
    auto_toc_count = sum(1 for sm in slide_mapping if sm.get("auto_generated_toc"))
    regular_count = len(slide_mapping) - auto_toc_count
    
    for sm in slide_mapping:
        idx = sm["source_index"]
        
        # Handle auto-generated TOC slides
        if sm.get("auto_generated_toc"):
            title = "[AI生成目录]"
            src_type = "auto-toc"
            tmpl = sm.get("template_layout", "?")
            tmpl_type = sm.get("template_type", "")
            print(f"  {'*':<4} {title:<30} {src_type:<16}   {tmpl} [{tmpl_type}]")
        else:
            title = source_titles.get(idx, "")[:28]
            src_type = sm.get("source_type", "?")[:14]
            tmpl = sm.get("template_layout", "?")
            tmpl_type = sm.get("template_type", "")
            print(f"  {idx:<4} {title:<30} {src_type:<16}   {tmpl} [{tmpl_type}]")
    
    print("─" * 60)
    print(f"  Total: {regular_count} source slides", end="")
    if auto_toc_count:
        print(f" + {auto_toc_count} AI-generated TOC")
    print()
    print("─" * 60 + "\n")


def _analyze_template_layouts(template_pptx: str) -> List[dict]:
    """Analyze which layout types are available in the template.

    Now also extracts per-placeholder styling (bodyPr, lstStyle, defRPr)
    from each layout so that injected content can inherit the layout's
    intended spacing, margins, and default font sizes.
    """
    layouts = []
    with zipfile.ZipFile(template_pptx, "r") as zf:
        layout_files = [n for n in zf.namelist()
                       if n.startswith("ppt/slideLayouts/") and n.endswith(".xml")
                       and "_rels" not in n]
        layout_files.sort()

        for layout_path in layout_files:
            layout_file = Path(layout_path).name
            xml = zf.read(layout_path).decode("utf-8")
            name_m = re.search(r'<p:cSld[^>]*name="([^"]*)"', xml)
            layout_name = name_m.group(1) if name_m else layout_file.replace(".xml", "")
            detected_type = _detect_layout_type(layout_name, xml)
            placeholder_types = re.findall(r'<p:ph[^>]*type="([^"]*)"', xml)

            # Extract per-placeholder styling from this layout
            ph_styles = _extract_layout_placeholder_styles(xml)

            layouts.append({
                "layout_file": layout_file,
                "layout_name": layout_name,
                "detected_type": detected_type,
                "placeholder_types": placeholder_types,
                "has_body": "body" in placeholder_types or "obj" in placeholder_types,
                "has_title": "title" in placeholder_types or "ctrTitle" in placeholder_types,
                "ph_styles": ph_styles,   # NEW: per-placeholder style info
            })

    return layouts


def _extract_template_colors(template_pptx: str) -> Dict[str, str]:
    """Extract the primary/accent/background colors from a template's theme XML.

    Returns a dict with keys: primary, secondary, accent, bg_light, bg_dark,
    text_on_light, text_on_dark.  Falls back to neutral values if not found.
    """
    defaults = {
        "primary":       "1E2761",
        "secondary":     "CADCFC",
        "accent":        "C9A84C",
        "bg_light":      "FFFFFF",
        "bg_dark":       "1E2761",
        "text_on_light": "1E2761",
        "text_on_dark":  "FFFFFF",
    }

    try:
        with zipfile.ZipFile(template_pptx, "r") as zf:
            theme_files = [n for n in zf.namelist()
                          if n.startswith("ppt/theme/") and n.endswith(".xml")]
            if not theme_files:
                return defaults

            xml = zf.read(theme_files[0]).decode("utf-8")

            def _get_color(tag: str) -> Optional[str]:
                m = re.search(
                    rf'<a:{tag}>\s*(?:<a:srgbClr val="([0-9A-Fa-f]{{6}})"'
                    r'|<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})")',
                    xml,
                )
                if m:
                    return (m.group(1) or m.group(2)).upper()
                return None

            dk1      = _get_color("dk1")   or defaults["text_on_light"]
            lt1      = _get_color("lt1")   or defaults["bg_light"]
            accent1  = _get_color("accent1") or defaults["primary"]
            accent2  = _get_color("accent2") or defaults["secondary"]
            accent3  = _get_color("accent3") or defaults["accent"]
            dk2      = _get_color("dk2")   or defaults["bg_dark"]

            return {
                "primary":       accent1,
                "secondary":     accent2,
                "accent":        accent3,
                "bg_light":      lt1,
                "bg_dark":       dk2,
                "text_on_light": dk1,
                "text_on_dark":  lt1,
            }
    except Exception:
        return defaults


def _extract_template_fonts(template_pptx: str) -> Dict[str, str]:
    """Extract the font scheme (majorFont / minorFont) from the template's theme XML.

    OOXML font scheme has two slots:
    - majorFont (a:majorFont) → used for headings / titles
    - minorFont (a:minorFont) → used for body text

    Each slot can define:
    - latin typeface  → Latin/Western scripts
    - ea    typeface  → East Asian (CJK) scripts
    - cs    typeface  → Complex scripts (Arabic, Hebrew, …)

    PowerPoint also supports the special magic typefaces "+mj-lt" (major latin)
    and "+mn-lt" (minor latin) as placeholders that resolve to the theme fonts.
    We never hard-code those; instead we extract the real names so we can write
    them explicitly into injected runs when needed.

    Returns a dict with keys:
        major_latin, major_ea, minor_latin, minor_ea
    Falls back to empty strings (= "inherit from template") when not found.
    """
    result = {
        "major_latin": "",
        "major_ea": "",
        "minor_latin": "",
        "minor_ea": "",
    }
    try:
        with zipfile.ZipFile(template_pptx, "r") as zf:
            theme_files = [n for n in zf.namelist()
                           if n.startswith("ppt/theme/") and n.endswith(".xml")]
            if not theme_files:
                return result

            xml = zf.read(theme_files[0]).decode("utf-8")

            # Extract majorFont block
            major_m = re.search(r'<a:majorFont>(.*?)</a:majorFont>', xml, re.DOTALL)
            if major_m:
                block = major_m.group(1)
                lat = re.search(r'<a:latin\b[^>]*typeface="([^"]+)"', block)
                ea  = re.search(r'<a:ea\b[^>]*typeface="([^"]+)"', block)
                if lat:
                    result["major_latin"] = lat.group(1)
                if ea and ea.group(1) not in ("", " "):
                    result["major_ea"] = ea.group(1)

            # Extract minorFont block
            minor_m = re.search(r'<a:minorFont>(.*?)</a:minorFont>', xml, re.DOTALL)
            if minor_m:
                block = minor_m.group(1)
                lat = re.search(r'<a:latin\b[^>]*typeface="([^"]+)"', block)
                ea  = re.search(r'<a:ea\b[^>]*typeface="([^"]+)"', block)
                if lat:
                    result["minor_latin"] = lat.group(1)
                if ea and ea.group(1) not in ("", " "):
                    result["minor_ea"] = ea.group(1)

    except Exception:
        pass

    return result


def _extract_layout_placeholder_styles(layout_xml: str) -> Dict[str, dict]:
    """Extract per-placeholder styling from a slideLayout XML.

    For each placeholder (identified by type), extracts:
    - bodyPr XML attributes (wrap, rtlCol, anchor, etc.) → to preserve layout bodyPr
    - defPPr / defRPr attributes (default paragraph/run props) → for reference
    - The full bodyPr element text → will be re-injected to preserve spacing/margins
    - spPr XML (shape properties) → to preserve fills, borders, effects

    Returns a dict keyed by placeholder type string (e.g. "title", "body", "subTitle").
    """
    styles: Dict[str, dict] = {}

    # Find all <p:sp> elements in the layout (simplified: find ph type + txBody pair)
    for sp_m in re.finditer(r'<p:sp\b.*?</p:sp>', layout_xml, re.DOTALL):
        sp_xml = sp_m.group(0)

        # Get placeholder type
        ph_m = re.search(r'<p:ph\b[^>]*type="([^"]+)"', sp_xml)
        if not ph_m:
            # body placeholder may have no explicit type attribute (idx-only)
            # Check for idx="1" which is the standard body placeholder
            ph_idx_m = re.search(r'<p:ph\b[^>]*idx="(\d+)"', sp_xml)
            ph_type = f"_idx{ph_idx_m.group(1)}" if ph_idx_m else None
        else:
            ph_type = ph_m.group(1)

        if not ph_type:
            continue

        info: dict = {}

        # Extract spPr (shape properties) for decorative styling (fills, borders, etc.)
        sppr_m = re.search(r'<p:spPr\b[^>]*>(.*?)</p:spPr>', sp_xml, re.DOTALL)
        if sppr_m:
            info["spPr_xml"] = sppr_m.group(0)

        # Extract the full <a:bodyPr ...> opening tag (or self-closing)
        body_pr_m = re.search(r'<a:bodyPr(\s[^>]*)?>|<a:bodyPr\s*/>', sp_xml)
        if body_pr_m:
            info["bodyPr_attrs"] = body_pr_m.group(1) or ""

        # Extract lstStyle / defPPr / defRPr for reference (may be multi-level)
        lst_m = re.search(r'<a:lstStyle>(.*?)</a:lstStyle>', sp_xml, re.DOTALL)
        if lst_m:
            lst_xml = lst_m.group(1)
            # Get defPPr (default paragraph props)
            def_ppr_m = re.search(r'<a:defPPr>(.*?)</a:defPPr>', lst_xml, re.DOTALL)
            if def_ppr_m:
                info["defPPr_xml"] = def_ppr_m.group(0)
            # Get first lvl1pPr (level-1 paragraph props)
            lvl1_m = re.search(r'<a:lvl1pPr\b[^>]*(?:/>|>.*?</a:lvl1pPr>)', lst_xml, re.DOTALL)
            if lvl1_m:
                info["lvl1pPr_xml"] = lvl1_m.group(0)
                # Extract default run props inside lvl1
                def_rpr_m = re.search(r'<a:defRPr\b[^>]*(?:/>|>.*?</a:defRPr>)', lvl1_m.group(0), re.DOTALL)
                if def_rpr_m:
                    info["defRPr_xml"] = def_rpr_m.group(0)
                    # Extract sz attribute from defRPr
                    sz_m = re.search(r'\bsz="(\d+)"', def_rpr_m.group(0))
                    if sz_m:
                        info["default_sz"] = int(sz_m.group(1))

        styles[ph_type] = info

    return styles


def _detect_layout_type(layout_name: str, layout_xml: str) -> str:
    """Classify a template layout type.

    Uses layout name first (case-insensitive, supports Chinese names),
    then falls back to placeholder type analysis.
    """
    name_lower = layout_name.lower()

    for layout_type, keywords in LAYOUT_TYPE_KEYWORDS.items():
        if any(kw in name_lower for kw in keywords):
            return layout_type

    # Fallback: analyze placeholder types in the layout XML
    ph_types = re.findall(r'<p:ph[^>]*type="([^"]*)"', layout_xml)
    ph_count = len(re.findall(r'<p:ph\b', layout_xml))

    if "ctrTitle" in ph_types:
        return "title_slide"
    if "subTitle" in ph_types:
        return "title_slide"
    # Two body/obj placeholders → two-column
    body_count = sum(1 for t in ph_types if t in ("body", "obj"))
    if body_count >= 2:
        return "two_column"
    if "body" in ph_types or "obj" in ph_types:
        return "content_slide"
    # No body placeholder at all → section header or blank
    if ph_count <= 1:
        return "section_header"
    return "content_slide"


def _detect_template_has_toc(template_layouts: List[dict]) -> bool:
    """Check if template has a table of contents / agenda layout."""
    toc_keywords = ["agenda", "contents", "outline", "目录", "大纲", "内容", "index"]
    for layout in template_layouts:
        layout_name = layout.get("layout_name", "").lower()
        if any(kw in layout_name for kw in toc_keywords):
            return True
    return False


def _detect_source_has_toc(source_slides: List[dict]) -> Optional[dict]:
    """Check if source PPT already has a table of contents slide."""
    toc_keywords = ["agenda", "contents", "outline", "目录", "大纲", "content"]
    for slide in source_slides:
        title = slide.get("title", "").lower()
        if any(kw in title for kw in toc_keywords):
            return slide
        # Also check if it's classified as agenda type
        if slide.get("type") == "agenda":
            return slide
    return None


def _generate_toc_content(source_slides: List[dict]) -> dict:
    """Generate table of contents content from source slide titles."""
    toc_items = []
    section_slides = []
    
    for slide in source_slides:
        slide_type = slide.get("type", "content")
        title = slide.get("title", "").strip()
        
        # Skip title slide and already identified TOC slide
        if slide_type == "title" or slide.get("is_toc", False):
            continue
            
        # Collect section headers and important slides
        if slide_type == "section" and title:
            section_slides.append({"title": title, "type": "section"})
        elif title and len(title) < 50:  # Reasonable title length
            # Avoid duplicates
            if not any(item["title"] == title for item in section_slides):
                section_slides.append({"title": title, "type": "content"})
    
    # Limit to 6-7 items for visual clarity
    if len(section_slides) > 7:
        section_slides = section_slides[:7]
    
    return {
        "title": "目录",
        "subtitle": "Contents",
        "body": [item["title"] for item in section_slides],
        "type": "agenda",
        "layout_hint": "list_content",
        "is_auto_generated": True,
    }


def _calculate_mapping_confidence(
    source_slide: dict,
    chosen_layout: dict,
    layouts_by_type: Dict[str, List[dict]],
) -> Dict[str, any]:
    """Calculate confidence score for a slide-to-layout mapping.

    Returns dict with:
    - score: 0-100 confidence score
    - reason: string explaining the score
    - risk_level: 'low' | 'medium' | 'high'
    """
    hint = source_slide.get("layout_hint", "content_slide")
    source_type = source_slide.get("type", "content")
    layout_type = chosen_layout.get("detected_type", "unknown")
    layout_name = chosen_layout.get("layout_name", "").lower()

    score = 100
    reasons = []
    risk_factors = []

    # Exact match bonus
    if hint in layouts_by_type and hint == layout_type:
        score += 10
        reasons.append("Exact layout hint match")

    # Source type match bonus
    elif source_type == layout_type:
        score += 5
        reasons.append("Source type match")

    # Penalize mismatches
    if hint and hint != layout_type:
        score -= 20
        risk_factors.append(f"Layout hint mismatch: {hint} != {layout_type}")

    # Check for high-risk mappings
    # Section mapped to content slide (loss of visual hierarchy)
    if source_type == "section" and layout_type == "content_slide":
        score -= 30
        risk_factors.append("Section slide mapped to content layout (visual hierarchy loss)")

    # Content mapped to section header (content overflow risk)
    if source_type in ("content", "title") and layout_type == "section_header":
        score -= 25
        risk_factors.append("Content slide mapped to section layout (overflow risk)")

    # Content mapped to title slide (unlikely)
    if source_type == "content" and layout_type == "title_slide":
        score -= 40
        risk_factors.append("Content slide mapped to title layout (severe mismatch)")

    # Check content complexity vs layout capacity
    body_length = len(" ".join(source_slide.get("body", [])))
    has_images = source_slide.get("has_images", False)
    has_tables = source_slide.get("has_tables", False)

    # Too much content for simple layouts
    if body_length > 500 and layout_type in ("section_header", "title_slide"):
        score -= 20
        risk_factors.append("Excessive content for simple layout")

    # Images mapped to text-only layout
    if has_images and layout_type == "section_header":
        score -= 15
        risk_factors.append("Images in section layout (may not display)")

    # Determine risk level
    risk_level = "low"
    if score < 50:
        risk_level = "high"
    elif score < 75:
        risk_level = "medium"

    # Clamp score
    score = max(0, min(100, score))

    return {
        "score": score,
        "reason": "; ".join(reasons) if reasons else "Default matching",
        "risk_factors": risk_factors,
        "risk_level": risk_level,
    }


def _auto_map_slides(source_slides: List[dict], template_layouts: List[dict],
                    verbose: bool, interactive: bool = False) -> List[dict]:
    """Automatically map source slides to template layouts with intelligent cycling.
    
    When source has more slides than template layouts, this function intelligently
    cycles through available layouts to ensure visual variety and avoid monotony.
    
    Additionally, if the template has a TOC layout but the source doesn't have
    a TOC slide, this function will auto-generate one.
    """
    mapping = []

    # Build layout lookup by type
    layouts_by_type: Dict[str, List[dict]] = {}
    for layout in template_layouts:
        t = layout["detected_type"]
        if t not in layouts_by_type:
            layouts_by_type[t] = []
        layouts_by_type[t].append(layout)

    # Fallback order when preferred type not found
    fallback_order = ["content_slide", "list_content", "title_slide", "section_header"]
    
    # Track layout usage for cycling when source slides > template layouts
    layout_usage_count: Dict[str, int] = {}
    last_used_layout: Dict[str, str] = {}  # Track last used layout per type
    
    # Check for TOC situation
    template_has_toc = _detect_template_has_toc(template_layouts)
    source_has_toc = _detect_source_has_toc(source_slides)
    auto_toc_inserted = False
    
    if verbose:
        if template_has_toc:
            print("  Template has TOC layout detected")
        if source_has_toc:
            print(f"  Source has TOC slide: {source_has_toc.get('title', 'Untitled')}")

    for slide in source_slides:
        if "error" in slide:
            mapping.append({
                "source_index": slide["index"],
                "source_type": "unknown",
                "template_layout": template_layouts[0]["layout_file"] if template_layouts else "",
                "template_type": "unknown",
            })
            continue

        hint = slide.get("layout_hint", "content_slide")
        source_type = slide.get("type", "content")
        
        # Check if we should insert auto-generated TOC before first content slide
        if (template_has_toc and not source_has_toc and not auto_toc_inserted 
            and source_type not in ("title", "section") and slide["index"] > 1):
            # Insert auto-generated TOC slide
            toc_content = _generate_toc_content(source_slides)
            toc_layout = None
            
            # Find best TOC layout
            for layout_type in ["list_content", "content_slide", "section_header"]:
                if layout_type in layouts_by_type and layouts_by_type[layout_type]:
                    toc_layout = layouts_by_type[layout_type][0]
                    break
            
            if toc_layout:
                if verbose:
                    print(f"  [AUTO-TOC] Inserting auto-generated table of contents "
                          f"using layout: {toc_layout['layout_file']}")
                
                mapping.append({
                    "source_index": 0,  # Special marker for auto-generated
                    "source_type": "agenda",
                    "template_layout": toc_layout["layout_file"],
                    "template_type": toc_layout["detected_type"],
                    "layout_name": toc_layout["layout_name"],
                    "auto_generated_toc": True,
                    "toc_content": toc_content,
                })
                auto_toc_inserted = True

        # Find best matching layout with cycling support
        chosen_layout = None
        candidate_layouts = []

        # Try exact type match
        if hint in layouts_by_type and layouts_by_type[hint]:
            candidate_layouts = layouts_by_type[hint]
        # Try source_type if hint didn't work
        elif source_type in layouts_by_type:
            candidate_layouts = layouts_by_type[source_type]
        # Try fallbacks
        else:
            for fb in fallback_order:
                if fb in layouts_by_type:
                    candidate_layouts = layouts_by_type[fb]
                    break
        
        # Ultimate fallback: all layouts
        if not candidate_layouts and template_layouts:
            candidate_layouts = template_layouts

        # Choose layout with intelligent cycling
        if candidate_layouts:
            if len(candidate_layouts) == 1:
                chosen_layout = candidate_layouts[0]
            else:
                # Find least recently used layout for this type to avoid repetition
                min_usage = float('inf')
                for layout in candidate_layouts:
                    layout_file = layout["layout_file"]
                    usage = layout_usage_count.get(layout_file, 0)
                    if usage < min_usage:
                        min_usage = usage
                        chosen_layout = layout
                
                # If all have equal usage, pick one that's different from last used
                if chosen_layout and min_usage > 0:
                    last_layout = last_used_layout.get(hint or source_type)
                    for layout in candidate_layouts:
                        if layout["layout_file"] != last_layout:
                            chosen_layout = layout
                            break

        if chosen_layout:
            layout_file = chosen_layout["layout_file"]
            layout_usage_count[layout_file] = layout_usage_count.get(layout_file, 0) + 1
            last_used_layout[hint or source_type] = layout_file

            # Calculate mapping confidence
            confidence = _calculate_mapping_confidence(
                slide, chosen_layout, layouts_by_type
            )

            mapping_entry = {
                "source_index": slide["index"],
                "source_type": source_type,
                "template_layout": layout_file,
                "template_type": chosen_layout["detected_type"],
                "layout_name": chosen_layout["layout_name"],
                "confidence_score": confidence["score"],
                "confidence_reason": confidence["reason"],
                "risk_level": confidence["risk_level"],
            }

            if confidence["risk_factors"]:
                mapping_entry["risk_factors"] = confidence["risk_factors"]

            mapping.append(mapping_entry)

            if verbose:
                risk_indicator = {
                    "low": "✓",
                    "medium": "⚠️",
                    "high": "❌",
                }.get(confidence["risk_level"], "")

                print(f"  Slide {slide['index']}: {risk_indicator} mapped to {layout_file} "
                      f"(confidence: {confidence['score']}%, {confidence['risk_level']})")
                if confidence["risk_factors"]:
                    for factor in confidence["risk_factors"]:
                        print(f"    - {factor}")
            
            if verbose and len(source_slides) > len(template_layouts):
                print(f"  Slide {slide['index']}: mapped to {layout_file} "
                      f"(usage: {layout_usage_count[layout_file]})")
        else:
            print(f"Warning: No layout found for slide {slide['index']}", file=sys.stderr)

    # Print mapping summary and warnings
    if verbose:
        print("\n" + "=" * 60)
        print("MAPPING SUMMARY")
        print("=" * 60)

        high_risk = [m for m in mapping if m.get("risk_level") == "high"]
        medium_risk = [m for m in mapping if m.get("risk_level") == "medium"]

        if high_risk:
            print(f"\n❌ HIGH RISK MAPPINGS ({len(high_risk)}):")
            for m in high_risk:
                print(f"  Slide {m['source_index']}: {m['source_type']} → {m['layout_name']}")
                for factor in m.get("risk_factors", []):
                    print(f"    - {factor}")
            print("\n⚠️  These mappings may result in visual issues.")
            print("    Consider using --save-mapping to manually adjust.")

        if medium_risk:
            print(f"\n⚠️  MEDIUM RISK MAPPINGS ({len(medium_risk)}):")
            for m in medium_risk[:5]:  # Show first 5
                print(f"  Slide {m['source_index']}: {m['source_type']} → {m['layout_name']}")
                for factor in m.get("risk_factors", []):
                    print(f"    - {factor}")
            if len(medium_risk) > 5:
                print(f"  ... and {len(medium_risk) - 5} more")

        if not high_risk and not medium_risk:
            print("\n✓ All mappings are low risk.")

        print("=" * 60)

    # Interactive confirmation for high-risk mappings
    if interactive and high_risk:
        print(f"\n⚠️  {len(high_risk)} high-risk mapping(s) detected!")
        print("\nOptions:")
        print("  1. Continue anyway (auto-fix high-risk)")
        print("  2. Save mapping and exit (manual edit)")
        print("  3. Cancel operation")

        try:
            choice = input("\nYour choice [1-3]: ").strip()
            if choice == "2":
                # Save mapping for manual editing
                mapping_file = "mapping_review.json"
                with open(mapping_file, "w", encoding="utf-8") as f:
                    json.dump(mapping, f, indent=2, ensure_ascii=False)
                print(f"\n✓ Mapping saved to: {mapping_file}")
                print("  Edit the file and re-run with: --mapping mapping_review.json")
                sys.exit(0)
            elif choice == "3":
                print("\n✗ Operation cancelled.")
                sys.exit(0)
            # choice == "1" or any other input: continue
        except (KeyboardInterrupt, EOFError):
            print("\n✗ Operation cancelled.")
            sys.exit(0)

    return mapping


def _get_presentation_slide_order(unpacked_dir: Path) -> List[Tuple[str, str]]:
    """Get (slide_file, rId) list from presentation.xml."""
    pres_path = unpacked_dir / "ppt" / "presentation.xml"
    rels_path = unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels"

    rels_content = rels_path.read_text(encoding="utf-8")
    rid_to_file = {}
    for m in re.finditer(r'<Relationship[^>]+Id="([^"]+)"[^>]+Target="slides/([^"]+)"', rels_content):
        rid_to_file[m.group(1)] = m.group(2)

    pres_content = pres_path.read_text(encoding="utf-8")
    order = []
    for m in re.finditer(r'<p:sldId[^>]+id="(\d+)"[^>]+r:id="([^"]+)"', pres_content):
        rid = m.group(2)
        if rid in rid_to_file:
            order.append((rid_to_file[rid], rid))

    return order


def _extract_source_images(source_pptx: str, output_dir: Path) -> Dict[str, Path]:
    """Extract images from source PPTX to output directory."""
    image_map = {}
    with zipfile.ZipFile(source_pptx, "r") as zf:
        for name in zf.namelist():
            if name.startswith("ppt/media/"):
                fname = Path(name).name
                dest = output_dir / fname
                dest.write_bytes(zf.read(name))
                image_map[fname] = dest
    return image_map


def _build_source_slide_file_map(
    source_unpacked_dir: Path, source_slides: List[dict]
) -> Dict[int, str]:
    """Map source slide index (1-based) → slide filename in the unpacked source dir.

    Uses the slide_file field from extract_content if available, otherwise falls
    back to enumerating slides from presentation.xml order.
    """
    mapping: Dict[int, str] = {}

    # First try: use the slide_file field extracted by extract_content
    for s in source_slides:
        idx = s.get("index")
        sf = s.get("slide_file")
        if idx is not None and sf:
            mapping[idx] = sf

    if mapping:
        return mapping

    # Fallback: read presentation.xml sldIdLst order
    try:
        order = _get_presentation_slide_order(source_unpacked_dir)
        for i, (slide_file, _) in enumerate(order, start=1):
            mapping[i] = slide_file
    except Exception:
        pass

    return mapping


def _get_layout_rids(unpacked_dir: Path) -> Dict[str, str]:
    """Get layout file -> rId mapping from slideLayouts relationships."""
    rels_path = unpacked_dir / "ppt" / "slideLayouts"
    result = {}
    # Get from presentation rels
    pres_rels = (unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels").read_text(encoding="utf-8")
    for m in re.finditer(r'<Relationship[^>]+Id="([^"]+)"[^>]+Target="slideLayouts/([^"]+)"', pres_rels):
        result[m.group(2)] = m.group(1)
    return result


def _ensure_slide_masters_preserved(unpacked_dir: Path, verbose: bool = False) -> None:
    """Ensure slideMasters and their relationships are preserved from template.
    
    In PowerPoint's OOXML structure:
    - slide -> references slideLayout
    - slideLayout -> references slideMaster
    - slideMaster -> contains theme (colors, fonts, effects)
    
    This function ensures all slideMasters referenced by slideLayouts are
    preserved in the output, along with their themes and relationships.
    """
    import shutil
    
    layouts_dir = unpacked_dir / "ppt" / "slideLayouts"
    masters_dir = unpacked_dir / "ppt" / "slideMasters"
    masters_rels_dir = masters_dir / "_rels"
    themes_dir = unpacked_dir / "ppt" / "theme"
    
    # Collect all slideMasters referenced by layouts
    referenced_masters = set()
    layout_master_map = {}  # layout_file -> master_file
    
    if layouts_dir.exists():
        for rels_file in layouts_dir.glob("_rels/*.xml.rels"):
            rels_content = rels_file.read_text(encoding="utf-8")
            # Find slideMaster reference
            master_m = re.search(r'Target="(?:\.\/)?slideMasters/([^"]+)"', rels_content)
            if master_m:
                master_file = master_m.group(1)
                layout_file = rels_file.name.replace(".xml.rels", "")
                referenced_masters.add(master_file)
                layout_master_map[layout_file] = master_file
    
    if not referenced_masters:
        if verbose:
            print("  No slideMasters referenced by layouts")
        return
    
    if verbose:
        print(f"  Found {len(referenced_masters)} slideMasters referenced by layouts")
    
    # Ensure directories exist
    masters_dir.mkdir(exist_ok=True)
    masters_rels_dir.mkdir(exist_ok=True)
    
    # Check which masters are missing and copy them from template if needed
    # (They should already be there from unpack, but verify)
    for master_file in referenced_masters:
        master_path = masters_dir / master_file
        if not master_path.exists():
            if verbose:
                print(f"  Warning: slideMaster {master_file} not found in unpacked dir")
            continue
        
        # Ensure master rels file exists
        master_rels_path = masters_rels_dir / f"{master_file}.rels"
        if not master_rels_path.exists():
            if verbose:
                print(f"  Warning: slideMaster rels for {master_file} not found")
    
    # Ensure all themes referenced by masters are preserved
    if masters_dir.exists():
        for master_file in masters_dir.glob("*.xml"):
            master_rels_path = masters_rels_dir / f"{master_file.name}.rels"
            if master_rels_path.exists():
                rels_content = master_rels_path.read_text(encoding="utf-8")
                # Find theme references
                for theme_m in re.finditer(r'Target="(?:\.\/)?theme/([^"]+)"', rels_content):
                    theme_file = theme_m.group(1)
                    theme_path = themes_dir / theme_file
                    if not theme_path.exists() and verbose:
                        print(f"  Warning: theme {theme_file} referenced by {master_file.name} not found")
    
    # Ensure [Content_Types].xml includes slideMasters and themes
    ct_path = unpacked_dir / "[Content_Types].xml"
    if ct_path.exists():
        ct_content = ct_path.read_text(encoding="utf-8")
        modified = False
        
        # Add slideMaster content type if missing
        master_ct = '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        if "slideMasters/slideMaster" not in ct_content:
            # Add generic slideMaster content type
            ct_content = ct_content.replace(
                "</Types>",
                '  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>\n</Types>'
            )
            modified = True
        
        if modified:
            ct_path.write_text(ct_content, encoding="utf-8")
            if verbose:
                print("  Updated [Content_Types].xml with slideMaster entries")
    
    if verbose:
        print(f"  Preserved {len(referenced_masters)} slideMasters with themes")


def _extract_layout_placeholders(layout_xml: str) -> List[dict]:
    """Extract placeholder shape definitions from a slideLayout XML.

    For each <p:sp> in the layout that contains a <p:ph>, we extract:
    - ph_type:  the 'type' attribute of <p:ph> (e.g. 'title', 'body', 'subTitle')
                or '' for the implicit body placeholder (idx only, no type attr)
    - ph_idx:   the 'idx' attribute of <p:ph> (integer, default 0)
    - sp_id:    shape id (int)
    - name:     shape name string
    - xfrm:     dict with x, y, cx, cy (EMU integers) — used for positioning
                Falls back to full-slide defaults when the layout doesn't define it
                (common for layouts that inherit position from the slide master).

    This lets _create_slide_from_layout build a slide that already contains the
    correct placeholder shapes — so that _replace_placeholder_text /
    _replace_placeholder_content can find them by type/idx.
    """
    placeholders = []
    shape_id = 2  # start at 2; id=1 is reserved for the spTree group

    # Default slide dimensions (widescreen 16:9 in EMU)
    SLIDE_W = 9144000
    SLIDE_H = 5143500

    for sp_m in re.finditer(r'<p:sp\b.*?</p:sp>', layout_xml, re.DOTALL):
        sp_xml = sp_m.group(0)

        # Must have a placeholder tag
        ph_m = re.search(r'<p:ph\b([^>]*)/?>',  sp_xml)
        if not ph_m:
            continue

        ph_attrs = ph_m.group(1)
        ph_type_m = re.search(r'\btype="([^"]+)"', ph_attrs)
        ph_idx_m  = re.search(r'\bidx="(\d+)"',   ph_attrs)
        ph_type = ph_type_m.group(1) if ph_type_m else ""
        ph_idx  = int(ph_idx_m.group(1)) if ph_idx_m else 0

        # Skip decoration-only placeholders (footer, date, slide number)
        if ph_type in ("ftr", "dt", "sldNum"):
            continue

        # Extract shape name from nvPr
        name_m = re.search(r'<p:cNvPr[^>]*name="([^"]*)"', sp_xml)
        sp_name = name_m.group(1) if name_m else f"Placeholder {ph_idx}"

        # Extract xfrm (position / size) — may be absent if inherited from master
        xfrm_m = re.search(r'<a:xfrm\b.*?</a:xfrm>|<a:xfrm\b[^>]*/>', sp_xml, re.DOTALL)
        if xfrm_m:
            xfrm_xml = xfrm_m.group(0)
            off_m = re.search(r'<a:off\b[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"', xfrm_xml)
            ext_m = re.search(r'<a:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"', xfrm_xml)
            x  = int(off_m.group(1)) if off_m else 0
            y  = int(off_m.group(2)) if off_m else 0
            cx = int(ext_m.group(1)) if ext_m else SLIDE_W
            cy = int(ext_m.group(2)) if ext_m else SLIDE_H
        else:
            # Reasonable defaults when layout inherits from master
            if ph_type in ("title", "ctrTitle"):
                x, y, cx, cy = 457200, 274638, 8229600, 1143000
            elif ph_type == "subTitle":
                x, y, cx, cy = 1371600, 1600200, 6400800, 1828800
            elif ph_type == "body":
                x, y, cx, cy = 457200, 1600200, 8229600, 3200400
            else:
                x, y, cx, cy = 457200, 457200, 8229600, 4114800

        placeholders.append({
            "ph_type": ph_type,
            "ph_idx":  ph_idx,
            "sp_id":   shape_id,
            "name":    sp_name,
            "x": x, "y": y, "cx": cx, "cy": cy,
        })
        shape_id += 1

    return placeholders


def _build_placeholder_sp_xml(ph: dict, spPr_xml: str = "") -> str:
    """Build a <p:sp> XML element for a content placeholder.

    The shape has the correct <p:ph type="..."> tag so that
    _replace_placeholder_text / _replace_placeholder_content can find it.
    The txBody is empty (just bodyPr + lstStyle) — content is injected later.

    If spPr_xml is provided (from the layout), it's used instead of the default
    minimal spPr, preserving decorative styling like fills, borders, and effects.
    """
    ph_type = ph["ph_type"]
    ph_idx  = ph["ph_idx"]
    sp_id   = ph["sp_id"]
    name    = ph["name"]
    x, y, cx, cy = ph["x"], ph["y"], ph["cx"], ph["cy"]

    # Build the <p:ph> tag
    if ph_type:
        ph_tag = f'<p:ph type="{ph_type}"'
    else:
        ph_tag = '<p:ph'
    if ph_idx > 0:
        ph_tag += f' idx="{ph_idx}"'
    ph_tag += '/>'

    # Use layout's spPr if available, otherwise build minimal one
    if spPr_xml:
        # Extract just the inner content and update position
        sppr_inner = re.search(r'<p:spPr\b[^>]*>(.*)</p:spPr>', spPr_xml, re.DOTALL)
        if sppr_inner:
            # Update xfrm with our position but keep other properties
            inner = sppr_inner.group(1)
            # Replace xfrm section with our values
            inner = re.sub(
                r'<a:xfrm\b[^>]*>.*?</a:xfrm>|<a:xfrm\b[^>]*/>',
                f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>',
                inner,
                flags=re.DOTALL
            )
            sppr_content = f'<p:spPr>{inner}</p:spPr>'
        else:
            sppr_content = (
                f'<p:spPr>'
                f'<a:xfrm>'
                f'<a:off x="{x}" y="{y}"/>'
                f'<a:ext cx="{cx}" cy="{cy}"/>'
                f'</a:xfrm>'
                f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                f'</p:spPr>'
            )
    else:
        sppr_content = (
            f'<p:spPr>'
            f'<a:xfrm>'
            f'<a:off x="{x}" y="{y}"/>'
            f'<a:ext cx="{cx}" cy="{cy}"/>'
            f'</a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'</p:spPr>'
        )

    return (
        f'<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="{name}"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr>{ph_tag}</p:nvPr>'
        f'</p:nvSpPr>'
        f'{sppr_content}'
        f'<p:txBody>'
        f'<a:bodyPr/>'
        f'<a:lstStyle/>'
        f'<a:p><a:endParaRPr lang="zh-CN" dirty="0"/></a:p>'
        f'</p:txBody>'
        f'</p:sp>'
    )


def _create_slide_from_layout(unpacked_dir: Path, layout_file: str) -> Tuple[str, str]:
    """Create a new slide using a layout, with placeholder shapes extracted from the layout.

    Unlike the old implementation which created a completely empty spTree,
    this version reads the layout XML to find all placeholder shapes (title,
    body, subTitle, etc.) and pre-populates the slide's spTree with
    corresponding <p:sp> elements.  This is critical: without placeholder
    shapes in the slide XML, _replace_placeholder_text and
    _replace_placeholder_content will find nothing to replace, resulting in
    a visually blank slide.

    Additionally:
    - Placeholder shapes inherit spPr (fill, border, effects) from the layout
    - The slide background is copied from the layout

    Returns (slide_file, rId).
    """
    slides_dir = unpacked_dir / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"
    rels_dir.mkdir(exist_ok=True)

    # Get next slide number
    existing = [int(m.group(1)) for f in slides_dir.glob("slide*.xml")
                if (m := re.match(r"slide(\d+)\.xml", f.name))]
    next_num = max(existing) + 1 if existing else 1
    slide_file = f"slide{next_num}.xml"

    # Read layout XML so we can extract placeholder shapes and background
    layout_path = unpacked_dir / "ppt" / "slideLayouts" / layout_file
    layout_xml = ""
    if layout_path.exists():
        layout_xml = layout_path.read_text(encoding="utf-8")

    # Extract placeholder definitions from the layout
    placeholders = _extract_layout_placeholders(layout_xml)

    # Extract layout styles for spPr inheritance
    layout_styles = _extract_layout_placeholder_styles(layout_xml)

    # Build placeholder shape XML strings with inherited spPr styling
    ph_shapes = []
    for ph in placeholders:
        ph_type = ph["ph_type"] or "_idx1"
        spPr_xml = layout_styles.get(ph_type, {}).get("spPr_xml", "")
        ph_shapes.append(_build_placeholder_sp_xml(ph, spPr_xml))

    ph_shapes_xml = "\n      ".join(ph_shapes)
    if ph_shapes_xml:
        ph_shapes_xml = "\n      " + ph_shapes_xml

    # Extract slide background from layout (if present)
    bg_xml = ""
    bg_m = re.search(r'<p:bg\b.*?</p:bg>', layout_xml, re.DOTALL)
    if bg_m:
        bg_xml = f"\n    {bg_m.group(0)}"

    # Build slide XML with placeholder shapes and background
    slide_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>{bg_xml}
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>{ph_shapes_xml}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>'''

    (slides_dir / slide_file).write_text(slide_xml, encoding="utf-8")

    # Create rels file pointing to the layout
    rels_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/{layout_file}"/>
</Relationships>'''
    (rels_dir / f"{slide_file}.rels").write_text(rels_xml, encoding="utf-8")

    # Add to Content_Types.xml
    ct_path = unpacked_dir / "[Content_Types].xml"
    ct = ct_path.read_text(encoding="utf-8")
    override = f'<Override PartName="/ppt/slides/{slide_file}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    if f"/ppt/slides/{slide_file}" not in ct:
        ct = ct.replace("</Types>", f"  {override}\n</Types>")
        ct_path.write_text(ct, encoding="utf-8")

    # Add to presentation.xml.rels
    pres_rels_path = unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels"
    pres_rels = pres_rels_path.read_text(encoding="utf-8")
    rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', pres_rels)]
    next_rid_num = max(rids) + 1 if rids else 1
    rid = f"rId{next_rid_num}"
    new_rel = f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/{slide_file}"/>'
    pres_rels = pres_rels.replace("</Relationships>", f"  {new_rel}\n</Relationships>")
    pres_rels_path.write_text(pres_rels, encoding="utf-8")

    return slide_file, rid


def _duplicate_template_slide(unpacked_dir: Path, source_slide: str) -> Tuple[str, str]:
    """Duplicate an existing template slide. Returns (new_slide_file, rId)."""
    slides_dir = unpacked_dir / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"

    source_path = slides_dir / source_slide
    source_rels_path = rels_dir / f"{source_slide}.rels"

    # Get next slide number
    existing = [int(m.group(1)) for f in slides_dir.glob("slide*.xml")
                if (m := re.match(r"slide(\d+)\.xml", f.name))]
    next_num = max(existing) + 1 if existing else 1
    dest_file = f"slide{next_num}.xml"

    shutil.copy2(source_path, slides_dir / dest_file)
    if source_rels_path.exists():
        dest_rels = rels_dir / f"{dest_file}.rels"
        rels_content = source_rels_path.read_text(encoding="utf-8")
        # Remove notes references
        rels_content = re.sub(r'\s*<Relationship[^>]*notesSlide[^>]*/>\s*', '\n', rels_content)
        dest_rels.write_text(rels_content, encoding="utf-8")

    # Add to Content_Types.xml
    ct_path = unpacked_dir / "[Content_Types].xml"
    ct = ct_path.read_text(encoding="utf-8")
    override = f'<Override PartName="/ppt/slides/{dest_file}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
    if f"/ppt/slides/{dest_file}" not in ct:
        ct = ct.replace("</Types>", f"  {override}\n</Types>")
        ct_path.write_text(ct, encoding="utf-8")

    # Add to presentation.xml.rels
    pres_rels_path = unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels"
    pres_rels = pres_rels_path.read_text(encoding="utf-8")
    rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', pres_rels)]
    next_rid_num = max(rids) + 1 if rids else 1
    rid = f"rId{next_rid_num}"
    new_rel = f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/{dest_file}"/>'
    pres_rels = pres_rels.replace("</Relationships>", f"  {new_rel}\n</Relationships>")
    pres_rels_path.write_text(pres_rels, encoding="utf-8")

    return dest_file, rid


def _find_template_slide_for_layout(unpacked_dir: Path, layout_file: str) -> Optional[str]:
    """Find the first template slide that uses a given layout file."""
    slides_dir = unpacked_dir / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"

    for rels_path in sorted(rels_dir.glob("slide*.rels")):
        rels_content = rels_path.read_text(encoding="utf-8")
        if layout_file in rels_content:
            return rels_path.name.replace(".rels", "")
    return None


def _detect_lang(text: str) -> str:
    """Detect language tag for a text run.

    Returns 'zh-CN' when >10% of characters are CJK, otherwise 'en-US'.
    This prevents English content from being tagged as Chinese (which breaks
    spell-check and some font-substitution rules in PowerPoint).
    """
    if not text:
        return "zh-CN"
    cjk = sum(
        1 for ch in text
        if "\u4e00" <= ch <= "\u9fff"
        or "\u3400" <= ch <= "\u4dbf"
        or "\u20000" <= ch <= "\u2a6df"
    )
    return "zh-CN" if cjk / len(text) > 0.10 else "en-US"


def _analyze_template_color_semantics(template_colors: Dict[str, str]) -> Dict[str, str]:
    """Analyze template colors to understand their semantic usage.
    
    Returns a mapping from color role to hex value based on luminance analysis.
    This helps determine which colors to use for fills vs text vs accents.
    """
    colors = template_colors
    result = {}
    
    # Get luminance values for all colors
    def luminance(hex_color: str) -> float:
        """Calculate relative luminance of a hex color."""
        hex_color = hex_color.lstrip('#')
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        r, g, b = r/255, g/255, b/255
        r = r/12.92 if r <= 0.03928 else ((r+0.055)/1.055) ** 2.4
        g = g/12.92 if g <= 0.03928 else ((g+0.055)/1.055) ** 2.4
        b = b/12.92 if b <= 0.03928 else ((b+0.055)/1.055) ** 2.4
        return 0.2126 * r + 0.7152 * g + 0.0722 * b
    
    color_roles = {
        "primary": colors.get("primary", "0070C0"),
        "secondary": colors.get("secondary", "4472C4"),
        "accent": colors.get("accent", "ED7D31"),
        "text_on_light": colors.get("text_on_light", "000000"),
        "text_on_dark": colors.get("text_on_dark", "FFFFFF"),
        "bg_light": colors.get("bg_light", "FFFFFF"),
        "bg_dark": colors.get("bg_dark", "1A1A2E"),
    }
    
    # Classify each color by luminance
    dark_colors = []
    light_colors = []
    mid_colors = []
    
    for role, hex_val in color_roles.items():
        lum = luminance(hex_val)
        if lum < 0.3:
            dark_colors.append((role, hex_val, lum))
        elif lum > 0.7:
            light_colors.append((role, hex_val, lum))
        else:
            mid_colors.append((role, hex_val, lum))
    
    # Sort by luminance
    dark_colors.sort(key=lambda x: x[2])
    light_colors.sort(key=lambda x: x[2], reverse=True)
    mid_colors.sort(key=lambda x: x[2])
    
    # Assign semantic roles
    result["dark_fill"] = dark_colors[0][1] if dark_colors else "1A1A2E"
    result["light_fill"] = light_colors[0][1] if light_colors else "FFFFFF"
    result["accent_fill"] = mid_colors[-1][1] if mid_colors else colors.get("accent", "ED7D31")
    result["text_primary"] = color_roles["text_on_light"]
    result["text_secondary"] = "666666"  # Medium gray for secondary text
    
    # Keep original colors for direct mapping
    result.update(color_roles)
    
    return result


def _classify_source_color(hex_color: str) -> str:
    """Classify a source color by its luminance to match with template semantics."""
    hex_color = hex_color.lstrip('#')
    r, g, b = int(hex_color[0:2], 16)/255, int(hex_color[2:4], 16)/255, int(hex_color[4:6], 16)/255
    
    # Calculate relative luminance
    def to_linear(c):
        return c/12.92 if c <= 0.03928 else ((c+0.055)/1.055) ** 2.4
    lum = 0.2126 * to_linear(r) + 0.7152 * to_linear(g) + 0.0722 * to_linear(b)
    
    if lum < 0.3:
        return "dark"
    elif lum > 0.7:
        return "light"
    else:
        return "mid"


def _build_color_replacement_map(source_colors: set, template_semantics: dict) -> dict:
    """Build a mapping from source colors to template colors based on semantics.
    
    Source dark colors -> template dark colors
    Source light colors -> template light colors
    Source mid colors -> template accent/mid colors
    """
    # Classify source colors
    dark_sources = []
    light_sources = []
    mid_sources = []
    
    for color in source_colors:
        if color in (template_semantics.get("text_on_light", ""), 
                     template_semantics.get("text_on_dark", ""),
                     template_semantics.get("primary", ""),
                     template_semantics.get("secondary", ""),
                     template_semantics.get("accent", "")):
            continue  # Skip if it's already a template color
        cls = _classify_source_color(color)
        if cls == "dark":
            dark_sources.append(color)
        elif cls == "light":
            light_sources.append(color)
        else:
            mid_sources.append(color)
    
    # Build replacement map
    replacements = {}
    
    # Dark source colors -> dark template fill
    dark_template = template_semantics.get("dark_fill", "1A1A2E")
    for src in dark_sources:
        replacements[src] = dark_template
    
    # Light source colors -> light template fill
    light_template = template_semantics.get("light_fill", "FFFFFF")
    for src in light_sources:
        replacements[src] = light_template
    
    # Mid source colors -> accent fill
    accent_template = template_semantics.get("accent_fill", "ED7D31")
    for src in mid_sources:
        replacements[src] = accent_template
    
    return replacements


def _extract_colors_from_slide(slide_path: Path) -> set:
    """Extract all hard-coded srgbClr colors from a slide XML file."""
    colors = set()
    if not slide_path.exists():
        return colors
    
    xml = slide_path.read_text(encoding="utf-8")
    
    # Match srgbClr val="RRGGBB"
    for match in re.finditer(r'<a:srgbClr val="([A-Fa-f0-9]{6})"', xml):
        colors.add(match.group(1))
    
    # Also match srgbVal (older format)
    for match in re.finditer(r'srgbVal="([A-Fa-f0-9]{6})"', xml):
        colors.add(match.group(1))
    
    return colors


def _apply_template_colors_to_slide(
    slide_path: Path,
    template_colors: Dict[str, str],
    verbose: bool = False,
) -> None:
    """Replace hard-coded colors in a slide with template color scheme.
    
    This handles:
    - Custom shapes (non-placeholder <p:sp> elements)
    - Table cells and borders
    - Decorative elements (lines, backgrounds)
    - SmartArt and graphicFrame elements
    
    Text colors are already handled by _inject_content_into_slide.
    This function focuses on non-text element colors.
    """
    if not slide_path.exists():
        return
    
    xml = slide_path.read_text(encoding="utf-8")
    original_xml = xml
    
    # Analyze template color semantics
    template_semantics = _analyze_template_color_semantics(template_colors)
    
    # Extract colors used in this slide
    slide_colors = _extract_colors_from_slide(slide_path)
    
    # Build color replacement map
    replacements = _build_color_replacement_map(slide_colors, template_semantics)
    
    if not replacements:
        return
    
    # Apply replacements to all srgbClr elements
    # This includes shapes, tables, and other non-text elements
    for old_color, new_color in replacements.items():
        # Replace in solidFill elements (shapes, backgrounds)
        xml = re.sub(
            rf'(<a:solidFill>.*?<a:srgbClr val="){old_color}(")',
            lambda m: m.group(1) + new_color + m.group(2),
            xml,
            flags=re.DOTALL
        )
        
        # Replace in gradientFill stops
        xml = re.sub(
            rf'(<a:gradStop[^>]*clrType="srgb"[^>]*val="){old_color}(")',
            lambda m: m.group(1) + new_color + m.group(2),
            xml,
            flags=re.DOTALL
        )
        
        # Replace in ln (line/border) elements
        xml = re.sub(
            rf'(<a:ln[^>]*w="[^"]*"[^>]*>.*?<a:solidFill>.*?<a:srgbClr val="){old_color}(")',
            lambda m: m.group(1) + new_color + m.group(2),
            xml,
            flags=re.DOTALL
        )
        
        # Replace in table cells
        xml = re.sub(
            rf'(<a:tc>.*?<a:srgbClr val="){old_color}(")',
            lambda m: m.group(1) + new_color + m.group(2),
            xml,
            flags=re.DOTALL
        )
        
        # Replace in tblPr (table properties)
        xml = re.sub(
            rf'(<a:tblPr>.*?<a:srgbClr val="){old_color}(")',
            lambda m: m.group(1) + new_color + m.group(2),
            xml,
            flags=re.DOTALL
        )
    
    if xml != original_xml:
        slide_path.write_text(xml, encoding="utf-8")
        if verbose:
            print(f"  Replaced {len(replacements)} colors in {slide_path.name} to match template")


# ═════════════════════════════════════════════════════════════════════════════
# CUSTOM SHAPE HANDLING
# ═════════════════════════════════════════════════════════════════════════════

def _extract_text_from_custom_shapes(source_slide_path: Path) -> Dict[str, any]:
    """Extract text from custom shapes (non-placeholder <p:sp> elements).
    
    When source PPT uses custom shapes instead of standard placeholders,
    this function extracts all text content from those shapes.
    
    Returns:
        Dict with 'title', 'subtitle', 'body' keys containing extracted text.
    """
    if not source_slide_path.exists():
        return {"title": "", "subtitle": "", "body": []}
    
    xml = source_slide_path.read_text(encoding="utf-8")
    
    # Extract all text runs from the slide
    all_text_runs = re.findall(r'<a:t>([^<]*)</a:t>', xml)
    
    # Filter out empty runs and keep track of their positions
    text_items = []
    for i, match in enumerate(re.finditer(r'<a:t>([^<]*)</a:t>', xml)):
        text = match.group(1).strip()
        if text:
            text_items.append({
                "text": text,
                "pos": match.start(),
                "len": len(text)
            })
    
    if not text_items:
        return {"title": "", "subtitle": "", "body": []}
    
    # Analyze text items to classify them
    # Heuristic: first substantial text is likely title, rest are body
    result = {
        "title": "",
        "subtitle": "",
        "body": []
    }
    
    # Classify by position and characteristics
    # Check if there's a clear title pattern (short, bold, or at top)
    for item in text_items:
        text = item["text"]
        # Skip very short strings that might be labels
        if len(text) <= 2:
            continue
        # Skip common decorative text
        if text in ("©", "®", "™", "·", "-", "—", "–"):
            continue
        
        if not result["title"]:
            # First substantial text = title
            result["title"] = text
        elif len(text) <= 30 and not result["subtitle"]:
            # Short text after title = subtitle
            result["subtitle"] = text
        else:
            # Rest = body content
            result["body"].append(text)
    
    return result


def _merge_custom_shape_content(
    source_slide: dict,
    source_unpacked_dir: Path,
    source_slide_file: str
) -> dict:
    """Merge content from custom shapes into source_slide dict.
    
    If the source_slide dict is empty (no title/body), try to extract
    content directly from the source slide XML.
    """
    # Check if source_slide already has content
    has_title = bool(source_slide.get("title", "").strip())
    has_body = bool(source_slide.get("body", []))
    
    if has_title and has_body:
        return source_slide  # Already has content
    
    # Try to extract from source slide XML
    source_slide_path = source_unpacked_dir / "ppt" / "slides" / source_slide_file
    extracted = _extract_text_from_custom_shapes(source_slide_path)
    
    # Merge extracted content with existing source_slide
    merged = dict(source_slide)
    
    if not has_title and extracted["title"]:
        merged["title"] = extracted["title"]
        if verbose_print := False:  # Will be set by caller if needed
            pass
    
    if not has_body and extracted["body"]:
        # Preserve existing body if it's not empty, otherwise use extracted
        existing_body = source_slide.get("body", [])
        if not existing_body:
            merged["body"] = extracted["body"]
        else:
            # Merge: append extracted body items not already in existing
            existing_set = set(existing_body)
            for item in extracted["body"]:
                if item not in existing_set:
                    existing_body.append(item)
            merged["body"] = existing_body
    
    has_subtitle = bool(source_slide.get("subtitle", "").strip())
    if not has_subtitle:
        if extracted["subtitle"]:
            merged["subtitle"] = extracted["subtitle"]
    
    return merged


# ═════════════════════════════════════════════════════════════════════════════
# BEAUTIFY PPT — INTEGRATED FROM beautify_ppt.py
# ═════════════════════════════════════════════════════════════════════════════
# These functions are embedded here to enable deep integration where
# beautification ENHANCES the template's style (colors, fonts) rather than
# replacing it with a completely different theme.
#
# Key difference from standalone beautify_ppt.py:
# - Uses template's own colors/fonts as the theme base
# - Applied per-slide during slide building (not post-processing)
# - Adds 10 layout variants, gradient backgrounds, visual accents

THEMES = {
    "executive": {
        "name": "Executive",
        "primary": "1E2761", "secondary": "CADCFC", "accent": "C9A84C",
        "bg_light": "FFFFFF", "bg_dark": "1E2761",
        "text_on_dark": "FFFFFF", "text_on_light": "1E2761",
        "text_muted": "6B7280",
        "header_font": "Cambria", "body_font": "Calibri",
        "title_bold": True, "title_size": 4000, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "1E2761", "gradient_end": "2A3A7C",
    },
    "tech": {
        "name": "Tech",
        "primary": "028090", "secondary": "1C2541", "accent": "02C39A",
        "bg_light": "F0FFFE", "bg_dark": "0B0C10",
        "text_on_dark": "FFFFFF", "text_on_light": "1C2541",
        "text_muted": "4B5563",
        "header_font": "Trebuchet MS", "body_font": "Calibri",
        "title_bold": True, "title_size": 3800, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "028090", "gradient_end": "02C39A",
    },
    "creative": {
        "name": "Creative",
        "primary": "F96167", "secondary": "2F3C7E", "accent": "F9E795",
        "bg_light": "FFFDF9", "bg_dark": "2F3C7E",
        "text_on_dark": "FFFFFF", "text_on_light": "2F3C7E",
        "text_muted": "6B7280",
        "header_font": "Georgia", "body_font": "Calibri",
        "title_bold": True, "title_size": 4000, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "F96167", "gradient_end": "F9E795",
    },
    "warm": {
        "name": "Warm",
        "primary": "B85042", "secondary": "84B59F", "accent": "ECE2D0",
        "bg_light": "FFFDF9", "bg_dark": "B85042",
        "text_on_dark": "FFFFFF", "text_on_light": "3D2B1F",
        "text_muted": "78716C",
        "header_font": "Palatino Linotype", "body_font": "Calibri",
        "title_bold": True, "title_size": 3800, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "B85042", "gradient_end": "D4685A",
    },
    "minimal": {
        "name": "Minimal",
        "primary": "36454F", "secondary": "F2F2F2", "accent": "212121",
        "bg_light": "FFFFFF", "bg_dark": "36454F",
        "text_on_dark": "FFFFFF", "text_on_light": "36454F",
        "text_muted": "9CA3AF",
        "header_font": "Calibri", "body_font": "Calibri",
        "title_bold": True, "title_size": 4000, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "36454F", "gradient_end": "4A5A65",
    },
    "bold": {
        "name": "Bold",
        "primary": "990011", "secondary": "2F3C7E", "accent": "FCF6F5",
        "bg_light": "FFFFFF", "bg_dark": "1A1A2E",
        "text_on_dark": "FFFFFF", "text_on_light": "1A1A2E",
        "text_muted": "6B7280",
        "header_font": "Arial Black", "body_font": "Arial",
        "title_bold": True, "title_size": 4400, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "990011", "gradient_end": "B81A2C",
    },
    "nature": {
        "name": "Nature",
        "primary": "2C5F2D", "secondary": "97BC62", "accent": "F5F5F5",
        "bg_light": "FAFFF5", "bg_dark": "2C5F2D",
        "text_on_dark": "FFFFFF", "text_on_light": "1A2E1B",
        "text_muted": "6B7280",
        "header_font": "Georgia", "body_font": "Calibri",
        "title_bold": True, "title_size": 4000, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "2C5F2D", "gradient_end": "4A7F4B",
    },
    "ocean": {
        "name": "Ocean",
        "primary": "065A82", "secondary": "1C7293", "accent": "9FFFCB",
        "bg_light": "F0F8FF", "bg_dark": "02364A",
        "text_on_dark": "FFFFFF", "text_on_light": "02364A",
        "text_muted": "64748B",
        "header_font": "Calibri", "body_font": "Calibri",
        "title_bold": True, "title_size": 4000, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "065A82", "gradient_end": "1C7293",
    },
    "elegant": {
        "name": "Elegant",
        "primary": "2C3E50", "secondary": "E8E8E8", "accent": "E74C3C",
        "bg_light": "FAFAFA", "bg_dark": "1A1A2E",
        "text_on_dark": "FFFFFF", "text_on_light": "2C3E50",
        "text_muted": "7F8C8D",
        "header_font": "Georgia", "body_font": "Calibri",
        "title_bold": True, "title_size": 4200, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "2C3E50", "gradient_end": "34495E",
    },
    "modern": {
        "name": "Modern",
        "primary": "6C5CE7", "secondary": "A29BFE", "accent": "FD79A8",
        "bg_light": "F8F9FA", "bg_dark": "2D3436",
        "text_on_dark": "FFFFFF", "text_on_light": "2D3436",
        "text_muted": "636E72",
        "header_font": "Segoe UI", "body_font": "Segoe UI",
        "title_bold": True, "title_size": 4000, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "6C5CE7", "gradient_end": "A29BFE",
    },
    "sunset": {
        "name": "Sunset",
        "primary": "E17055", "secondary": "FDCB6E", "accent": "D63031",
        "bg_light": "FFF9F0", "bg_dark": "2D142C",
        "text_on_dark": "FFFFFF", "text_on_light": "2D142C",
        "text_muted": "8B7355",
        "header_font": "Georgia", "body_font": "Calibri",
        "title_bold": True, "title_size": 4000, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "E17055", "gradient_end": "FDCB6E",
    },
    "forest": {
        "name": "Forest",
        "primary": "1B4332", "secondary": "52B788", "accent": "D8F3DC",
        "bg_light": "F1F8E9", "bg_dark": "081C15",
        "text_on_dark": "FFFFFF", "text_on_light": "1B4332",
        "text_muted": "52796F",
        "header_font": "Cambria", "body_font": "Calibri",
        "title_bold": True, "title_size": 4000, "body_size": 1800, "caption_size": 1200,
        "gradient_start": "1B4332", "gradient_end": "2D6A4F",
    },
}

DARK_BG_TYPES = {"title", "section", "conclusion"}

LAYOUT_VARIANTS = [
    "accent_bar", "accent_bar", "numbered_list", "stat_highlight",
    "two_tone", "header_band", "card_grid", "timeline",
    "split_diagonal", "image_focus", "quote_block",
]


def _build_template_theme(
    template_colors: Dict[str, str],
    template_fonts: Dict[str, str],
    beautify_theme: Optional[str] = None,
) -> dict:
    """Build a beautify-compatible theme from template colors and fonts.

    If beautify_theme is specified, use that theme's design system (fonts, sizes, gradients)
    but replace the colors with the template's actual palette. This gives us the best of
    both worlds: the template's authentic colors + beautify's professional design system.
    """
    base = None
    if beautify_theme and beautify_theme in THEMES:
        base = dict(THEMES[beautify_theme])
    else:
        # Default to minimal for fonts/sizes, use template for colors
        base = dict(THEMES["minimal"])

    # Override colors with template's actual palette
    base["primary"]       = template_colors.get("primary", base["primary"])
    base["secondary"]     = template_colors.get("secondary", base["secondary"])
    base["accent"]         = template_colors.get("accent", base["accent"])
    base["bg_light"]       = template_colors.get("bg_light", base["bg_light"])
    base["bg_dark"]        = template_colors.get("bg_dark", base["bg_dark"])
    base["text_on_light"]  = template_colors.get("text_on_light", base["text_on_light"])
    base["text_on_dark"]   = template_colors.get("text_on_dark", base["text_on_dark"])

    # Use template fonts if available
    if template_fonts.get("major_latin"):
        base["header_font"] = template_fonts["major_latin"]
    if template_fonts.get("minor_latin"):
        base["body_font"] = template_fonts["minor_latin"]

    # Derive gradient from template primary/secondary
    base["gradient_start"] = base["primary"]
    base["gradient_end"]   = base["secondary"]

    return base


def _pick_layout_variant(
    slide_type: str,
    body_items: list,
    layout_streak: list,
    slide_index: int,
) -> str:
    """Choose a layout variant to maximize visual variety."""
    if slide_type not in ("content", "list_content", "agenda"):
        return "none"

    recent = layout_streak[-3:] if len(layout_streak) >= 3 else layout_streak
    recent_set = set(recent)
    candidates = [v for v in LAYOUT_VARIANTS if v not in recent_set]
    if not candidates:
        candidates = LAYOUT_VARIANTS[:]

    if len(body_items) >= 4 and "numbered_list" in candidates:
        return "numbered_list"
    elif len(body_items) <= 2 and any(c.isdigit() for c in " ".join(body_items)):
        if "stat_highlight" in candidates:
            return "stat_highlight"
    elif slide_index % 4 == 3 and "two_tone" in candidates:
        return "two_tone"

    return candidates[slide_index % len(candidates)]


def _restructure_slide(
    xml: str, theme: dict, layout_variant: str,
    body_items: list, use_dark: bool,
) -> str:
    """Apply structural layout changes using template theme colors."""
    if layout_variant == "none" or not layout_variant:
        return xml
    existing_sp_count = xml.count("<p:sp>")
    if existing_sp_count > 6:
        return xml

    primary = theme["primary"]
    secondary = theme["secondary"]
    accent = theme["accent"]

    if layout_variant == "two_tone":
        xml = _add_two_tone_panel(xml, primary)
    elif layout_variant == "header_band":
        xml = _add_header_band(xml, primary)
    elif layout_variant == "numbered_list":
        xml = _add_numbered_circles(xml, primary, body_items)
    elif layout_variant == "stat_highlight":
        xml = _add_stat_highlight(xml, primary, accent, body_items)
    elif layout_variant == "card_grid":
        xml = _add_card_grid(xml, primary, accent, body_items, use_dark)
    elif layout_variant == "timeline":
        xml = _add_timeline(xml, primary, accent, body_items)
    elif layout_variant == "split_diagonal":
        xml = _add_split_diagonal(xml, primary, secondary, use_dark)
    elif layout_variant == "image_focus":
        xml = _add_image_focus_frame(xml, primary)
    elif layout_variant == "quote_block":
        xml = _add_quote_block(xml, primary, accent, body_items, use_dark)

    return xml


def _add_two_tone_panel(xml: str, primary: str) -> str:
    panel_xml = (
        '\n<p:sp><p:nvSpPr>'
        '<p:cNvPr id="9010" name="TwoTonePanel"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        '</p:nvSpPr><p:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="3200400" cy="5143500"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"><a:alpha val="92000"/></a:srgbClr></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:sp>'
    )
    if "</p:spTree>" in xml and "TwoTonePanel" not in xml:
        xml = xml.replace("<p:spTree>", "<p:spTree>" + panel_xml, 1)
    return xml


def _add_header_band(xml: str, primary: str) -> str:
    band_xml = (
        '\n<p:sp><p:nvSpPr>'
        '<p:cNvPr id="9011" name="HeaderBand"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        '</p:nvSpPr><p:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1143000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:sp>'
    )
    if "</p:spTree>" in xml and "HeaderBand" not in xml:
        xml = xml.replace("<p:spTree>", "<p:spTree>" + band_xml, 1)
    return xml


def _add_numbered_circles(xml: str, primary: str, body_items: list) -> str:
    if not body_items:
        return xml
    count = min(len(body_items), 6)
    circle_size = 411480
    start_y = 1188000
    row_height = 685800
    circle_x = 320040
    circles_xml = ""
    for idx in range(count):
        cy = start_y + idx * row_height
        circles_xml += (
            f'\n<p:sp><p:nvSpPr>'
            f'<p:cNvPr id="{9020 + idx}" name="NumCircle{idx}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
            f'</p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{circle_x}" y="{cy}"/>'
            f'<a:ext cx="{circle_size}" cy="{circle_size}"/></a:xfrm>'
            f'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln>'
            f'</p:spPr><p:txBody>'
            f'<a:bodyPr anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="en-US" sz="1400" b="1" dirty="0">'
            f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            f'</a:rPr><a:t>{idx + 1}</a:t></a:r>'
            f'</a:p></p:txBody></p:sp>'
        )
    if "</p:spTree>" in xml and "NumCircle0" not in xml:
        xml = xml.replace("</p:spTree>", circles_xml + "</p:spTree>")
    return xml


def _add_stat_highlight(xml: str, primary: str, accent: str, body_items: list) -> str:
    if not body_items:
        return xml
    first_item = body_items[0][:50]
    stat_xml = (
        '\n<p:sp><p:nvSpPr>'
        '<p:cNvPr id="9030" name="StatCallout"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        '</p:nvSpPr><p:spPr>'
        '<a:xfrm><a:off x="457200" y="1280160"/><a:ext cx="3657600" cy="2286000"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr><p:txBody>'
        '<a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/>'
        f'<a:r><a:rPr lang="zh-CN" sz="2400" b="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        f'</a:rPr><a:t>{first_item}</a:t></a:r>'
        '</a:p></p:txBody></p:sp>'
    )
    if "</p:spTree>" in xml and "StatCallout" not in xml:
        xml = xml.replace("</p:spTree>", stat_xml + "</p:spTree>")
    return xml


def _add_card_grid(xml: str, primary: str, accent: str, body_items: list, use_dark: bool) -> str:
    if not body_items or len(body_items) < 2:
        return xml
    count = min(len(body_items), 4)
    cards_per_row = 2 if count > 2 else count
    card_width = 3200400
    card_height = 1371600
    gap = 228600
    start_x = 914400
    start_y = 1371600
    bg_color = "FFFFFF" if not use_dark else "1A1A2E"
    text_color = "000000" if not use_dark else "FFFFFF"
    cards_xml = ""
    for idx in range(count):
        row = idx // cards_per_row
        col = idx % cards_per_row
        cx = start_x + col * (card_width + gap)
        cy = start_y + row * (card_height + gap)
        card_color = primary if idx % 2 == 0 else accent
        if use_dark and card_color == accent:
            card_color = primary
        item_text = body_items[idx][:40] if idx < len(body_items) else ""
        cards_xml += (
            f'\n<p:sp><p:nvSpPr>'
            f'<p:cNvPr id="{9040 + idx}" name="Card{idx}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
            f'</p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{cx}" y="{cy}"/>'
            f'<a:ext cx="{card_width}" cy="{card_height}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{card_color}"><a:alpha val="15000"/></a:srgbClr></a:solidFill>'
            f'<a:ln><a:solidFill><a:srgbClr val="{card_color}"/></a:solidFill></a:ln>'
            f'</p:spPr><p:txBody>'
            f'<a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="zh-CN" sz="1600" b="1" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{text_color}"/></a:solidFill>'
            f'</a:rPr><a:t>{item_text}</a:t></a:r>'
            f'</a:p></p:txBody></p:sp>'
        )
    if "</p:spTree>" in xml and "Card0" not in xml:
        xml = xml.replace("</p:spTree>", cards_xml + "</p:spTree>")
    return xml


def _add_timeline(xml: str, primary: str, accent: str, body_items: list) -> str:
    if not body_items or len(body_items) < 2:
        return xml
    count = min(len(body_items), 5)
    timeline_y = 2286000
    start_x = 685800
    end_x = 8458200
    step = (end_x - start_x) // (count - 1) if count > 1 else 0
    timeline_xml = (
        f'\n<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="9050" name="TimelineLine"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        f'</p:nvSpPr><p:spPr>'
        f'<a:xfrm><a:off x="{start_x}" y="{timeline_y}"/>'
        f'<a:ext cx="{end_x - start_x}" cy="76200"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
    )
    for idx in range(count):
        cx = start_x + idx * step
        item_text = body_items[idx][:25] if idx < len(body_items) else ""
        timeline_xml += (
            f'\n<p:sp><p:nvSpPr>'
            f'<p:cNvPr id="{9060 + idx}" name="TimelineNode{idx}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
            f'</p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{cx - 228600}" y="{timeline_y - 190500}"/>'
            f'<a:ext cx="457200" cy="457200"/></a:xfrm>'
            f'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{accent}"/></a:solidFill>'
            f'<a:ln><a:solidFill><a:srgbClr val="{primary}"/></a:solidFill></a:ln>'
            f'</p:spPr><p:txBody>'
            f'<a:bodyPr anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="en-US" sz="1400" b="1" dirty="0">'
            f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            f'</a:rPr><a:t>{idx + 1}</a:t></a:r>'
            f'</a:p></p:txBody></p:sp>'
            f'\n<p:sp><p:nvSpPr>'
            f'<p:cNvPr id="{9070 + idx}" name="TimelineLabel{idx}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
            f'</p:nvSpPr><p:spPr>'
            f'<a:xfrm><a:off x="{cx - 457200}" y="{timeline_y + 685800}"/>'
            f'<a:ext cx="914400" cy="457200"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'<a:noFill/><a:ln><a:noFill/></a:ln>'
            f'</p:spPr><p:txBody>'
            f'<a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="zh-CN" sz="1200" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
            f'</a:rPr><a:t>{item_text}</a:t></a:r>'
            f'</a:p></p:txBody></p:sp>'
        )
    if "</p:spTree>" in xml and "TimelineLine" not in xml:
        xml = xml.replace("</p:spTree>", timeline_xml + "</p:spTree>")
    return xml


def _add_split_diagonal(xml: str, primary: str, secondary: str, use_dark: bool) -> str:
    diagonal_xml = (
        '\n<p:sp><p:nvSpPr>'
        '<p:cNvPr id="9080" name="DiagonalSplit"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        '</p:nvSpPr><p:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="5143500"/></a:xfrm>'
        '<a:custGeom><a:avLst/>'
        '<a:gdLst>'
        '<a:gd name="x1" fmla="val 0"/><a:gd name="y1" fmla="val 0"/>'
        '<a:gd name="x2" fmla="val 4572000"/><a:gd name="y2" fmla="val 0"/>'
        '<a:gd name="x3" fmla="val 9144000"/><a:gd name="y3" fmla="val 5143500"/>'
        '<a:gd name="x4" fmla="val 0"/><a:gd name="y4" fmla="val 5143500"/>'
        '</a:gdLst>'
        '<a:pathLst><a:path w="9144000" h="5143500">'
        '<a:moveTo><a:pt x="x1" y="y1"/></a:moveTo>'
        '<a:lnTo><a:pt x="x2" y="y2"/></a:lnTo>'
        '<a:lnTo><a:pt x="x3" y="y3"/></a:lnTo>'
        '<a:lnTo><a:pt x="x4" y="y4"/></a:lnTo>'
        '<a:close/></a:path></a:pathLst>'
        '</a:custGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"><a:alpha val="20000"/></a:srgbClr></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:sp>'
    )
    if "</p:spTree>" in xml and "DiagonalSplit" not in xml:
        xml = xml.replace("<p:spTree>", "<p:spTree>" + diagonal_xml, 1)
    return xml


def _add_image_focus_frame(xml: str, primary: str) -> str:
    frame_xml = (
        '\n<p:sp><p:nvSpPr>'
        '<p:cNvPr id="9090" name="ImageFrame"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        '</p:nvSpPr><p:spPr>'
        '<a:xfrm><a:off x="457200" y="457200"/><a:ext cx="8229600" cy="4229100"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        f'<a:ln w="76200"><a:solidFill><a:srgbClr val="{primary}"/></a:solidFill></a:ln>'
        '</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:sp>'
        f'\n<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="9091" name="CornerTL"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        f'</p:nvSpPr><p:spPr>'
        f'<a:xfrm><a:off x="457200" y="457200"/><a:ext cx="228600" cy="228600"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
        f'\n<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="9092" name="CornerBR"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        f'</p:nvSpPr><p:spPr>'
        f'<a:xfrm><a:off x="8458200" y="4229100"/><a:ext cx="228600" cy="228600"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
    )
    if "</p:spTree>" in xml and "ImageFrame" not in xml:
        xml = xml.replace("</p:spTree>", frame_xml + "</p:spTree>")
    return xml


def _add_quote_block(xml: str, primary: str, accent: str, body_items: list, use_dark: bool) -> str:
    if not body_items:
        return xml
    quote_text = body_items[0][:80]
    text_color = "000000" if not use_dark else "FFFFFF"
    quote_xml = (
        f'\n<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="9100" name="QuoteBar"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        f'</p:nvSpPr><p:spPr>'
        f'<a:xfrm><a:off x="1371600" y="1371600"/><a:ext cx="114300" cy="2746380"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{accent}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
        f'\n<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="9101" name="QuoteMark"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        f'</p:nvSpPr><p:spPr>'
        f'<a:xfrm><a:off x="914400" y="1143000"/><a:ext cx="457200" cy="457200"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/><a:ln><a:noFill/></a:ln>'
        f'</p:spPr><p:txBody>'
        f'<a:bodyPr anchor="ctr"/><a:lstStyle/>'
        f'<a:p><a:pPr algn="ctr"/>'
        f'<a:r><a:rPr lang="en-US" sz="4800" i="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="{primary}"/></a:solidFill>'
        f'</a:rPr><a:t>"</a:t></a:r>'
        f'</a:p></p:txBody></p:sp>'
        f'\n<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="9102" name="QuoteText"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        f'</p:nvSpPr><p:spPr>'
        f'<a:xfrm><a:off x="1600200" y="1600200"/><a:ext cx="5943600" cy="2286000"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/><a:ln><a:noFill/></a:ln>'
        f'</p:spPr><p:txBody>'
        f'<a:bodyPr anchor="ctr" wrap="square"/><a:lstStyle/>'
        f'<a:p><a:pPr algn="l"/>'
        f'<a:r><a:rPr lang="zh-CN" sz="2000" i="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="{text_color}"/></a:solidFill>'
        f'</a:rPr><a:t>{quote_text}</a:t></a:r>'
        f'</a:p></p:txBody></p:sp>'
    )
    if "</p:spTree>" in xml and "QuoteBar" not in xml:
        xml = xml.replace("</p:spTree>", quote_xml + "</p:spTree>")
    return xml


def _set_background(xml: str, theme: dict, use_dark: bool) -> str:
    bg_color = theme["bg_dark"] if use_dark else theme["bg_light"]
    if "<p:bg>" in xml:
        xml = re.sub(
            r'(<p:bg>.*?<a:solidFill>.*?<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
            lambda m: m.group(1) + bg_color + m.group(2),
            xml, flags=re.DOTALL,
        )
    else:
        bg_xml = (
            f'\n  <p:bg><p:bgPr>'
            f'<a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill>'
            f'<a:effectLst/></p:bgPr></p:bg>'
        )
        xml = re.sub(r'(<p:cSld[^>]*>)', r'\1' + bg_xml, xml, count=1)
    return xml


def _set_gradient_background(xml: str, theme: dict, use_dark: bool) -> str:
    start_color = theme.get("gradient_start", theme["primary"])
    end_color = theme.get("gradient_end", theme["secondary"])
    gradient_xml = (
        f'\n  <p:bg><p:bgPr>'
        f'<a:gradFill rotWithShape="1">'
        f'<a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="{start_color}"/></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="{end_color}"/></a:gs>'
        f'</a:gsLst>'
        f'<a:lin ang="2700000" scaled="1"/>'
        f'<a:tileRect/></a:gradFill>'
        f'<a:effectLst/></p:bgPr></p:bg>'
    )
    if "<p:bg>" in xml:
        xml = re.sub(r'<p:bg>.*?</p:bg>', gradient_xml.strip(), xml, flags=re.DOTALL)
    else:
        xml = re.sub(r'(<p:cSld[^>]*>)', r'\1' + gradient_xml, xml, count=1)
    return xml


def _update_text_colors(xml: str, theme: dict, use_dark: bool) -> str:
    title_color = theme["text_on_dark"] if use_dark else theme["primary"]
    body_color  = theme["text_on_dark"] if use_dark else theme["text_on_light"]

    def recolor_sp(m):
        sp_xml = m.group(0)
        is_title = bool(re.search(r'<p:ph[^>]*type="(?:title|ctrTitle)"', sp_xml))
        color = title_color if is_title else body_color

        def recolor_rpr(rm):
            rpr = rm.group(0)
            if "<a:solidFill>" not in rpr:
                return rpr
            new_rpr = re.sub(
                r'<a:solidFill>.*?</a:solidFill>',
                f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>',
                rpr, flags=re.DOTALL,
            )
            return new_rpr

        return re.sub(r'<a:rPr\b.*?</a:rPr>', recolor_rpr, sp_xml, flags=re.DOTALL)

    xml = re.sub(r'<p:sp\b.*?</p:sp>', recolor_sp, xml, flags=re.DOTALL)
    return xml


def _update_shape_colors(xml: str, theme: dict, use_dark: bool) -> str:
    accent_color = theme["accent"] if use_dark else theme["primary"]
    bg_light = theme["bg_light"].upper()
    bg_dark  = theme["bg_dark"].upper()

    def replace_accent(m):
        color = m.group(1).upper()
        r = int(color[0:2], 16)
        g = int(color[2:4], 16)
        b = int(color[4:6], 16)
        luminance = (r + g + b) / 3
        if luminance > 235 or luminance < 20 or color in (bg_light, bg_dark):
            return m.group(0)
        return f'<a:solidFill><a:srgbClr val="{accent_color}"/></a:solidFill>'

    def process_sppr(m):
        sppr = m.group(0)
        sppr = re.sub(
            r'<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/></a:solidFill>',
            replace_accent, sppr,
        )
        return sppr

    xml = re.sub(r'<p:spPr>.*?</p:spPr>', process_sppr, xml, flags=re.DOTALL)
    return xml


def _update_fonts(xml: str, theme: dict) -> str:
    header_font = theme.get("header_font", "Calibri")
    body_font = theme.get("body_font", "Calibri")
    common_fonts = [
        "Arial", "Helvetica", "Times New Roman", "Times", "Calibri",
        "Cambria", "Verdana", "Tahoma", "Trebuchet MS", "Georgia",
        "Palatino", "Garamond", "Comic Sans MS", "Impact",
        "Calibri Light", "Century Gothic",
    ]

    def replace_font(m):
        typeface = m.group(1)
        if any(f.lower() == typeface.lower() for f in common_fonts):
            return f'<a:latin typeface="{body_font}"'
        return m.group(0)

    xml = re.sub(r'<a:latin typeface="([^"]+)"', replace_font, xml)
    return xml


def _enhance_images(xml: str, theme: dict, use_dark: bool) -> str:
    primary = theme["primary"]

    def enhance_pic(m):
        pic_xml = m.group(0)
        if "<a:blip" not in pic_xml:
            return pic_xml
        if '<a:prstGeom prst="rect"' in pic_xml:
            pic_xml = pic_xml.replace(
                '<a:prstGeom prst="rect"><a:avLst/>',
                '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst>',
            )
        if '<a:effectLst>' not in pic_xml and '<p:spPr>' in pic_xml:
            shadow_xml = (
                '<a:effectLst>'
                '<a:outerShdw blurRad="63500" dist="50800" dir="2700000" algn="bl">'
                '<a:srgbClr val="000000"><a:alpha val="25000"/></a:srgbClr>'
                '</a:outerShdw></a:effectLst>'
            )
            pic_xml = pic_xml.replace('</p:spPr>', shadow_xml + '</p:spPr>')
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


def _beautify_tables(xml: str, theme: dict, use_dark: bool) -> str:
    if '<a:tbl>' not in xml:
        return xml
    primary = theme["primary"]
    accent = theme["accent"]
    header_bg = primary
    header_text = "FFFFFF"
    stripe_light = "FFFFFF"
    stripe_dark = _lighten_color(primary, 92)
    body_text = theme.get("text_on_light", "333333")

    def process_table(tbl_match):
        tbl_xml = tbl_match.group(0)
        row_pattern = re.compile(r'<a:tr[^>]*>(.*?)</a:tr>', re.DOTALL)
        rows = list(row_pattern.finditer(tbl_xml))
        if not rows:
            return tbl_xml

        for i, row_match in enumerate(rows):
            row_xml = row_match.group(0)
            is_header = (i == 0)
            is_stripe = (i % 2 == 0)
            if is_header:
                row_bg = f'<a:solidFill><a:srgbClr val="{header_bg}"/></a:solidFill>'
            elif is_stripe:
                row_bg = f'<a:solidFill><a:srgbClr val="{stripe_light}"/></a:solidFill>'
            else:
                row_bg = f'<a:solidFill><a:srgbClr val="{stripe_dark}"/></a:solidFill>'

            def process_cell(cell_match):
                cell_xml = cell_match.group(0)
                tcpr_match = re.search(r'<a:tcPr[^>]*>(.*?)</a:tcPr>', cell_xml, re.DOTALL)
                if tcpr_match:
                    tcpr = tcpr_match.group(0)
                    tcpr = re.sub(r'<a:solidFill>.*?</a:solidFill>', '', tcpr, flags=re.DOTALL)
                    tcpr = re.sub(r'(<a:tcPr[^>]*>)', r'\1' + row_bg, tcpr, count=1)
                    cell_xml = cell_xml.replace(tcpr_match.group(0), tcpr)
                else:
                    cell_xml = re.sub(
                        r'(<a:tc[^>]*>)',
                        r'\1<a:tcPr>' + row_bg +
                        f'<a:lnL w="12700"><a:solidFill><a:srgbClr val="{accent}"/></a:solidFill></a:lnL>'
                        f'<a:lnR w="12700"><a:solidFill><a:srgbClr val="{accent}"/></a:solidFill></a:lnR>'
                        f'<a:lnT w="12700"><a:solidFill><a:srgbClr val="{accent}"/></a:solidFill></a:lnT>'
                        f'<a:lnB w="12700"><a:solidFill><a:srgbClr val="{accent}"/></a:solidFill></a:lnB>'
                        '<a:cellMar><a:l mar="91440"/><a:r mar="91440"/><a:t mar="45720"/><a:b mar="45720"/></a:cellMar>',
                        cell_xml, count=1
                    )
                if is_header:
                    cell_xml = re.sub(
                        r'<a:rPr([^>]*)>',
                        r'<a:rPr\1><a:solidFill><a:srgbClr val="' + header_text + '"/></a:solidFill>',
                        cell_xml
                    )
                    cell_xml = re.sub(
                        r'<a:rPr([^>]*)(/>|><a:)',
                        r'<a:rPr\1><a:solidFill><a:srgbClr val="' + header_text + '"/></a:solidFill><a:',
                        cell_xml
                    )
                return cell_xml

            row_xml = re.sub(r'<a:tc>.*?</a:tc>', process_cell, row_xml, flags=re.DOTALL)
            tbl_xml = tbl_xml[:row_match.start()] + row_xml + tbl_xml[row_match.end():]
            row_pattern = re.compile(r'<a:tr[^>]*>(.*?)</a:tr>', re.DOTALL)
            rows = list(row_pattern.finditer(tbl_xml))

        return tbl_xml

    xml = re.sub(r'<a:tbl>.*?</a:tbl>', process_table, xml, flags=re.DOTALL)
    return xml


def _lighten_color(hex_color: str, percent: int) -> str:
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    factor = percent / 100
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"{r:02X}{g:02X}{b:02X}"


def _optimize_paragraph_spacing(xml: str, theme: dict) -> str:
    def add_spacing(m):
        p_xml = m.group(0)
        if '<a:lnSpc' in p_xml:
            return p_xml
        spacing_xml = '<a:lnSpc><a:spcPct val="120000"/></a:lnSpc>'
        if '<a:pPr' in p_xml:
            p_xml = p_xml.replace('</a:pPr>', spacing_xml + '</a:pPr>')
        else:
            p_xml = p_xml.replace('<a:p>', '<a:p><a:pPr>' + spacing_xml + '</a:pPr>')
        return p_xml

    xml = re.sub(r'<a:p>.*?</a:p>', add_spacing, xml, flags=re.DOTALL)
    return xml


def _remove_antipatterns(xml: str) -> str:
    def remove_accent_underlines(m):
        sp_xml = m.group(0)
        if '<p:ph' in sp_xml:
            return sp_xml
        if re.search(r'cy="[1-9][0-9]{3}"', sp_xml) and re.search(r'cx="[5-9][0-9]{5,}"', sp_xml):
            if '<a:solidFill>' in sp_xml and '<p:ph' not in sp_xml:
                return ''
        return sp_xml

    xml = re.sub(r'<p:sp\b.*?</p:sp>', remove_accent_underlines, xml, flags=re.DOTALL)
    return xml


def _add_accent_bar(xml: str, theme: dict) -> str:
    accent_bar = (
        '\n<p:sp><p:nvSpPr>'
        '<p:cNvPr id="9001" name="AccentBar"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/>'
        '</p:nvSpPr><p:spPr>'
        '<a:xfrm><a:off x="274638" y="640080"/><a:ext cx="45720" cy="4115040"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{theme["primary"]}"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:sp>'
    )
    if "</p:spTree>" in xml and "AccentBar" not in xml:
        xml = xml.replace("</p:spTree>", accent_bar + "</p:spTree>")
    return xml


def _apply_theme_to_master(unpacked_dir: Path, theme: dict, verbose: bool) -> None:
    """Update the slide master with template/theme colors and fonts."""
    master_dir = unpacked_dir / "ppt" / "slideMasters"
    if not master_dir.exists():
        return
    for master_path in master_dir.glob("slideMaster*.xml"):
        if verbose:
            print(f"  Updating master: {master_path.name}")
        xml = master_path.read_text(encoding="utf-8")
        # Update master color scheme
        xml = re.sub(
            r'(<a:dk1>.*?<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
            lambda m: m.group(1) + theme.get("text_on_light", "1E2761") + m.group(2),
            xml, flags=re.DOTALL
        )
        xml = re.sub(
            r'(<a:lt1>.*?<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
            lambda m: m.group(1) + theme.get("bg_light", "FFFFFF") + m.group(2),
            xml, flags=re.DOTALL
        )
        xml = _update_fonts(xml, theme)
        master_path.write_text(xml, encoding="utf-8")


def _update_theme_xml(unpacked_dir: Path, theme: dict, verbose: bool) -> None:
    """Update theme XML with the template's actual color palette."""
    theme_dir = unpacked_dir / "ppt" / "theme"
    if not theme_dir.exists():
        return
    for theme_path in theme_dir.glob("theme*.xml"):
        if verbose:
            print(f"  Updating theme: {theme_path.name}")
        xml = theme_path.read_text(encoding="utf-8")

        replacements = {
            "dk1": theme.get("text_on_light", "1E2761"),
            "lt1": theme.get("bg_light", "FFFFFF"),
            "accent1": theme.get("primary", "1E2761"),
            "accent2": theme.get("secondary", "CADCFC"),
            "accent3": theme.get("accent", "C9A84C"),
            "dk2": theme.get("bg_dark", "1E2761"),
        }
        for tag, color in replacements.items():
            xml = re.sub(
                rf'(<a:{tag}>\s*(?:<a:srgbClr val=")[0-9A-Fa-f]{{6}}")',
                f'<a:{tag}><a:srgbClr val="{color}"',
                xml,
            )
        theme_path.write_text(xml, encoding="utf-8")


def _beautify_slide_with_template_theme(
    slide_path: Path,
    theme: dict,
    slide_type: str,
    body_items: list,
    layout_streak: list,
    slide_index: int,
    dark_mode: bool = False,
    restructure: bool = True,
    use_gradient: bool = False,
    verbose: bool = False,
) -> dict:
    """Full beautification pipeline using template theme.

    This is the deep integration point: applies beautify_ppt's full styling
    system (backgrounds, colors, fonts, layout variants, images, tables)
    while preserving the template's authentic color palette and font choices.

    Returns beautification notes dict.
    """
    if not slide_path.exists():
        return {}

    xml = slide_path.read_text(encoding="utf-8")
    original_xml = xml
    notes = {}

    # Determine background type
    use_dark = (slide_type in DARK_BG_TYPES) or dark_mode
    if slide_type in ("title", "section", "conclusion") and use_gradient:
        xml = _set_gradient_background(xml, theme, use_dark)
    else:
        xml = _set_background(xml, theme, use_dark)

    # Update text colors
    xml = _update_text_colors(xml, theme, use_dark)

    # Update accent shape colors
    xml = _update_shape_colors(xml, theme, use_dark)

    # Update fonts
    xml = _update_fonts(xml, theme)

    # Enhance images (rounded corners, shadows, borders)
    xml = _enhance_images(xml, theme, use_dark)

    # Beautify tables
    xml = _beautify_tables(xml, theme, use_dark)

    # Remove anti-patterns
    xml = _remove_antipatterns(xml)

    # Structural layout enrichment
    if restructure and slide_type in ("content", "list_content", "agenda"):
        layout_variant = _pick_layout_variant(
            slide_type, body_items, layout_streak, slide_index
        )
        layout_streak.append(layout_variant)
        if layout_variant and layout_variant not in ("none", "accent_bar"):
            xml = _restructure_slide(xml, theme, layout_variant, body_items, use_dark)

    # Add visual accent bar for content slides
    if slide_type in ("content", "list_content", "agenda") and not use_dark:
        xml = _add_accent_bar(xml, theme)

    # Optimize paragraph spacing
    xml = _optimize_paragraph_spacing(xml, theme)

    # Title truncation (existing logic preserved)
    xml = _truncate_long_titles_in_xml(xml, notes)

    # Bullet point merging (existing logic preserved)
    xml = _merge_excess_bullets_in_xml(xml, notes)

    # Check for empty slides
    text_content = re.findall(r'<a:t>([^<]+)</a:t>', xml)
    non_empty = [t.strip() for t in text_content if t.strip() and len(t.strip()) > 1]
    if not non_empty:
        notes["empty_slide"] = "检测到空白幻灯片"

    if xml != original_xml:
        slide_path.write_text(xml, encoding="utf-8")
        if verbose:
            for key, value in notes.items():
                print(f"    美化: {value}")

    return notes


def _truncate_long_titles_in_xml(xml: str, notes: dict) -> str:
    def truncate_titles(m):
        p_xml = m.group(0)
        if '<p:ph type="title"' not in p_xml and '<p:ph type="ctrTitle"' not in p_xml:
            return p_xml

        def truncate_runs(m2):
            run_xml = m2.group(0)
            text_match = re.search(r'<a:t>([^<]{20,})</a:t>', run_xml)
            if text_match:
                long_text = text_match.group(1)
                truncated = long_text[:18] + "…"
                run_xml = run_xml.replace(f'<a:t>{long_text}</a:t>', f'<a:t>{truncated}</a:t>')
                notes["title_truncated"] = f"标题截短: \"{long_text[:30]}...\" → \"{truncated}\""
            return run_xml

        return re.sub(r'<a:r>.*?</a:r>', truncate_runs, p_xml, flags=re.DOTALL)

    return re.sub(r'<a:p>.*?</a:p>', truncate_titles, xml, flags=re.DOTALL)


def _merge_excess_bullets_in_xml(xml: str, notes: dict) -> str:
    def merge_bullets(m):
        p_xml = m.group(0)
        if '<p:ph type="body"' not in p_xml and '<p:ph type="obj"' not in p_xml:
            return p_xml

        bullet_count = len(re.findall(r'<a:p>.*?<a:pPr.*?<a:buChar', p_xml, flags=re.DOTALL))
        if bullet_count <= 6:
            return p_xml

        paragraphs = list(re.finditer(r'<a:p>.*?</a:p>', p_xml, flags=re.DOTALL))
        if len(paragraphs) <= 6:
            return p_xml

        merged_paragraphs = paragraphs[:6]
        to_merge = paragraphs[6:]

        merged_texts = []
        for p in to_merge:
            text_match = re.search(r'<a:t>([^<]*)</a:t>', p.group(0))
            if text_match:
                merged_texts.append(text_match.group(1).strip())

        if merged_texts:
            last_p = merged_paragraphs[-1].group(0)
            merged_text = "；".join(merged_texts)
            merged_text_xml = f'<a:r><a:rPr lang="zh-CN"/><a:t>（含{len(to_merge)}条合并：{merged_text}）</a:t></a:r>'
            last_p = last_p.replace('</a:p>', merged_text_xml + '</a:p>')
            p_xml = p_xml[:merged_paragraphs[-1].start()] + last_p + p_xml[merged_paragraphs[-1].end():]
            for p in reversed(to_merge):
                p_xml = p_xml[:p.start()] + p_xml[p.end():]
            notes["bullets_merged"] = f"要点合并: {bullet_count}条 → 6条"

        return p_xml

    return re.sub(r'<a:p>.*?</a:p>', merge_bullets, xml, flags=re.DOTALL)


# ═════════════════════════════════════════════════════════════════════════════
# DEPRECATED: Old per-slide beautification (kept for reference, superseded)
# ═════════════════════════════════════════════════════════════════════════════

def _beautify_slide_layout(slide_path: Path, verbose: bool = False) -> Dict[str, any]:
    """Apply AI beautification to a slide's layout and content.
    
    This function performs the following optimizations:
    1. Truncate over-long titles (>20 characters)
    2. Merge bullet points that exceed 6 items
    3. Split over-long body text (>40 chars per line)
    4. Detect and flag empty slides
    5. Optimize paragraph spacing
    
    Returns a dict of beautification notes for reporting.
    """
    if not slide_path.exists():
        return {}
    
    xml = slide_path.read_text(encoding="utf-8")
    original_xml = xml
    notes = {}
    
    # 1. Truncate over-long titles
    def truncate_long_titles(p_xml):
        if '<p:ph type="title"' not in p_xml and '<p:ph type="ctrTitle"' not in p_xml:
            return p_xml
        
        # Find text runs in title placeholder
        def truncate_runs(m):
            run_xml = m.group(0)
            # Check if this is inside a title placeholder
            text_match = re.search(r'<a:t>([^<]{20,})</a:t>', run_xml)
            if text_match:
                long_text = text_match.group(1)
                # Truncate to 18 chars and add ellipsis
                truncated = long_text[:18] + "…"
                run_xml = run_xml.replace(f'<a:t>{long_text}</a:t>', f'<a:t>{truncated}</a:t>')
                notes["title_truncated"] = f"标题截短: \"{long_text[:30]}...\" → \"{truncated}\""
            return run_xml
        
        xml = re.sub(r'<a:r>.*?</a:r>', truncate_runs, p_xml, flags=re.DOTALL)
        return xml
    
    # 2. Merge excess bullet points (more than 6)
    def merge_excess_bullets(p_xml):
        # Only process body placeholders
        if '<p:ph type="body"' not in p_xml and '<p:ph type="obj"' not in p_xml:
            return p_xml
        
        # Count bullet paragraphs
        bullet_count = len(re.findall(r'<a:p>.*?<a:pPr.*?<a:buChar', p_xml, flags=re.DOTALL))
        
        if bullet_count > 6:
            # Merge paragraphs after the 6th bullet
            paragraphs = list(re.finditer(r'<a:p>.*?</a:p>', p_xml, flags=re.DOTALL))
            if len(paragraphs) > 6:
                # Find the 7th bullet and merge it with the 6th
                merged_paragraphs = paragraphs[:6]
                to_merge = paragraphs[6:]
                
                # Get text from paragraphs to merge
                merged_texts = []
                for p in to_merge:
                    text_match = re.search(r'<a:t>([^<]*)</a:t>', p.group(0))
                    if text_match:
                        merged_texts.append(text_match.group(1).strip())
                
                if merged_texts:
                    # Append merged text to the last kept paragraph
                    last_p = merged_paragraphs[-1].group(0)
                    # Find the last </a:p> and insert the merged text before it
                    merged_text = "；".join(merged_texts)
                    merged_text_xml = f'<a:r><a:rPr lang="zh-CN"/><a:t>（含{len(to_merge)}条合并内容：{merged_text}）</a:t></a:r>'
                    last_p = last_p.replace('</a:p>', merged_text_xml + '</a:p>')
                    
                    # Replace the last kept paragraph
                    p_xml = p_xml[:merged_paragraphs[-1].start()] + last_p + p_xml[merged_paragraphs[-1].end():]
                    
                    # Remove the merged paragraphs
                    for p in reversed(to_merge):
                        p_xml = p_xml[:p.start()] + p_xml[p.end():]
                    
                    notes["bullets_merged"] = f"要点合并: {bullet_count}条 → 6条（合并了{len(to_merge)}条）"
        
        return p_xml
    
    # 3. Optimize paragraph spacing
    def optimize_spacing(p_xml):
        # Skip if already has spacing
        if '<a:lnSpc' in p_xml:
            return p_xml
        
        # Add 1.2 line spacing (120%)
        spacing_xml = '<a:lnSpc><a:spcPct val="120000"/></a:lnSpc>'
        
        # Insert after pPr or at the beginning
        if '<a:pPr' in p_xml:
            p_xml = p_xml.replace('</a:pPr>', spacing_xml + '</a:pPr>')
        else:
            p_xml = p_xml.replace('<a:p>', '<a:p><a:pPr>' + spacing_xml + '</a:pPr>')
        
        return p_xml
    
    # Apply optimizations
    xml = re.sub(r'<a:p>.*?</a:p>', truncate_long_titles, xml, flags=re.DOTALL)
    xml = re.sub(r'<a:p>.*?</a:p>', merge_excess_bullets, xml, flags=re.DOTALL)
    
    # Only add spacing to body text, not titles
    def add_spacing_conditional(p_xml):
        if '<p:ph type="body"' in p_xml or '<p:ph type="obj"' in p_xml:
            return optimize_spacing(p_xml)
        return p_xml
    
    xml = re.sub(r'<a:p>.*?</a:p>', add_spacing_conditional, xml, flags=re.DOTALL)
    
    # 4. Check for empty slides
    text_content = re.findall(r'<a:t>([^<]+)</a:t>', xml)
    non_empty_text = [t.strip() for t in text_content if t.strip() and len(t.strip()) > 1]
    if not non_empty_text:
        notes["empty_slide"] = "检测到空白幻灯片"
    
    if xml != original_xml:
        slide_path.write_text(xml, encoding="utf-8")
        if verbose:
            for key, value in notes.items():
                print(f"  美化: {value}")
    
    return notes


def _find_and_replace_text_in_element(
    elem: ET.Element,
    ph_types: set[str],
    new_text: str,
    rpr_attrs: str,
    rpr_children: str,
) -> bool:
    """Find placeholder element by type and replace its text content using ET.

    Returns True if replacement was made, False otherwise.
    """
    # Check if this element is a placeholder with matching type
    ph = elem.find("p:nvSpPr/p:nvPr/p:ph", NS)
    if ph is not None:
        ph_type = ph.get("type", "")
        if ph_type in ph_types:
            # Found matching placeholder, replace text content
            txBody = elem.find("p:txBody", NS)
            if txBody is not None:
                # Clear existing paragraphs and add new one
                for p in list(txBody):
                    if p.tag.endswith("}p") and p.tag != f"{{{NS['p']}}}p":
                        continue
                    if p.tag == f"{{{NS['a']}}}p":
                        txBody.remove(p)

                # Build new paragraph with text run
                new_p = ET.SubElement(txBody, f"{{{NS['a']}}}p")

                # Handle rPr
                if rpr_attrs.strip():
                    # Parse rPr attributes
                    rpr_elem = ET.SubElement(new_p, f"{{{NS['a']}}}rPr")
                    for attr_pair in rpr_attrs.strip().split():
                        if attr_pair.startswith("lang="):
                            rpr_elem.set("lang", attr_pair.split("=")[1].strip('"'))
                        elif "=" in attr_pair:
                            key, val = attr_pair.split("=", 1)
                            rpr_elem.set(key.strip(), val.strip('"'))

                    # Add rPr children (color, fonts)
                    for child_xml in rpr_children.split("</a:rPr>")[0].split("<a:rPr")[1:]:
                        if child_xml.strip():
                            child_str = f"<a:rPr{child_xml}</a:rPr>"
                            child_elem = ET.fromstring(child_str)
                            rpr_elem.append(child_elem)
                else:
                    rpr_elem = ET.SubElement(new_p, f"{{{NS['a']}}}rPr")
                    rpr_elem.set("lang", "zh-CN")
                    rpr_elem.set("dirty", "0")

                # Add text run
                r = ET.SubElement(new_p, f"{{{NS['a']}}}r")
                t = ET.SubElement(r, f"{{{NS['a']}}}t")
                t.text = new_text
                return True

    # Recursively search children
    for child in elem:
        if _find_and_replace_text_in_element(child, ph_types, new_text, rpr_attrs, rpr_children):
            return True
    return False


def _inject_content_using_et(
    slide_xml: str,
    title: str,
    subtitle: str,
    body_xml: str,
    title_color: str,
    body_color: str,
    title_latin_font: str,
    title_ea_font: str,
    body_latin_font: str,
    body_ea_font: str,
    verbose: bool,
) -> Tuple[str, bool, bool, bool]:
    """Inject content into slide XML using xml.etree.ElementTree.

    Returns (modified_xml, title_injected, subtitle_injected, body_injected).
    """
    try:
        # Parse XML with namespace handling
        # First, add namespace declarations to make parsing reliable
        nsmap = {
            "p": NS["p"],
            "a": NS["a"],
            "r": NS["r"],
        }

        root = ET.fromstring(slide_xml)

        title_injected = False
        subtitle_injected = False
        body_injected = False

        # Find all sp elements (shapes)
        for sp in root.iter(f"{{{NS['p']}}}sp"):
            # Check if this is a placeholder
            nvSpPr = sp.find("p:nvSpPr", NS)
            if nvSpPr is None:
                continue

            nvPr = nvSpPr.find("p:nvPr", NS)
            if nvPr is None:
                continue

            ph = nvPr.find("p:ph", NS)
            if ph is None:
                continue

            ph_type = ph.get("type", "")
            txBody = sp.find("p:txBody", NS)

            # Title injection
            if not title_injected and ph_type in ("title", "ctrTitle"):
                if verbose:
                    print(f"    [ET] Found title placeholder: type={ph_type}")

                # Clear existing paragraphs
                if txBody is not None:
                    for p in list(txBody):
                        p_tag = p.tag.split("}")[-1] if "}" in p.tag else p.tag
                        if p_tag == "p":
                            txBody.remove(p)

                    # Build title text run
                    new_p = ET.SubElement(txBody, f"{{{NS['a']}}}p")
                    rpr_elem = ET.SubElement(new_p, f"{{{NS['a']}}}rPr")
                    rpr_elem.set("lang", _detect_lang(title))
                    rpr_elem.set("dirty", "0")

                    if title_color:
                        fill = ET.SubElement(rpr_elem, f"{{{NS['a']}}}solidFill")
                        clr = ET.SubElement(fill, f"{{{NS['a']}}}srgbClr")
                        clr.set("val", title_color)

                    if title_latin_font:
                        latin = ET.SubElement(rpr_elem, f"{{{NS['a']}}}latin")
                        latin.set("typeface", title_latin_font)
                    if title_ea_font:
                        ea = ET.SubElement(rpr_elem, f"{{{NS['a']}}}ea")
                        ea.set("typeface", title_ea_font)

                    r = ET.SubElement(new_p, f"{{{NS['a']}}}r")
                    t = ET.SubElement(r, f"{{{NS['a']}}}t")
                    t.text = title
                    title_injected = True
                    if verbose:
                        print(f"    [ET] Injected title: {title[:50]}...")

            # Subtitle injection
            elif not subtitle_injected and ph_type == "subTitle":
                if verbose:
                    print(f"    [ET] Found subtitle placeholder: type={ph_type}")

                if txBody is not None:
                    for p in list(txBody):
                        p_tag = p.tag.split("}")[-1] if "}" in p.tag else p.tag
                        if p_tag == "p":
                            txBody.remove(p)

                    new_p = ET.SubElement(txBody, f"{{{NS['a']}}}p")
                    rpr_elem = ET.SubElement(new_p, f"{{{NS['a']}}}rPr")
                    rpr_elem.set("lang", _detect_lang(subtitle))
                    rpr_elem.set("dirty", "0")

                    if body_color:
                        fill = ET.SubElement(rpr_elem, f"{{{NS['a']}}}solidFill")
                        clr = ET.SubElement(fill, f"{{{NS['a']}}}srgbClr")
                        clr.set("val", body_color)

                    r = ET.SubElement(new_p, f"{{{NS['a']}}}r")
                    t = ET.SubElement(r, f"{{{NS['a']}}}t")
                    t.text = subtitle
                    subtitle_injected = True
                    if verbose:
                        print(f"    [ET] Injected subtitle: {subtitle[:50]}...")

            # Body injection
            elif not body_injected and ph_type in ("body", "obj"):
                if verbose:
                    print(f"    [ET] Found body placeholder: type={ph_type}")

                if txBody is not None:
                    # Parse body_xml and inject paragraphs
                    body_root = ET.fromstring(f"<root xmlns:a='{NS['a']}'>{body_xml}</root>")
                    body_paragraphs = list(body_root)

                    # Remove existing paragraphs
                    for p in list(txBody):
                        p_tag = p.tag.split("}")[-1] if "}" in p.tag else p.tag
                        if p_tag == "p":
                            txBody.remove(p)

                    # Add new paragraphs
                    for p_elem in body_paragraphs:
                        txBody.append(p_elem)

                    body_injected = True
                    if verbose:
                        print(f"    [ET] Injected {len(body_paragraphs)} body paragraphs")

        # Serialize back to string
        modified_xml = ET.tostring(root, encoding="unicode")
        return modified_xml, title_injected, subtitle_injected, body_injected

    except ET.ParseError as e:
        if verbose:
            print(f"    [ET] XML parse error: {e}, falling back to regex")
        return slide_xml, False, False, False
    except Exception as e:
        if verbose:
            print(f"    [ET] Error: {e}, falling back to regex")
        return slide_xml, False, False, False


def _inject_content_into_slide(
    unpacked_dir: Path,
    slide_file: str,
    source_slide: dict,
    template_colors: Dict[str, str],
    template_fonts: Dict[str, str],
    layout_ph_styles: Dict[str, dict],
    verbose: bool,
) -> None:
    """Replace placeholder content in a template slide with source content.

    Text formatting rules:
    - Colors: always use the template's color palette (primary/text_on_light/text_on_dark)
    - Font face (typeface): use the template's majorFont for titles, minorFont for body
      (extracted from the theme fontScheme).  If the theme font is empty, don't set
      typeface so PowerPoint inherits it from the layout/master.
    - Bold / italic: preserved from source slide's body_rich field
    - Font size: preserved from source slide's body_rich field when available;
      falls back to layout's default size (defRPr sz) so spacing matches template design
    - Language tag: auto-detected per run (zh-CN if >10% CJK, else en-US)
    - bodyPr: preserved from the layout placeholder (not overwritten) so margins,
      word wrap, and auto-fit settings remain as the template designer intended

    Fallback when placeholder replacement produces no visible change:
    - If title was given but no matching placeholder was found, attempt to inject
      into ANY text-capable placeholder that has no content yet (best-effort)
    """
    slide_path = unpacked_dir / "ppt" / "slides" / slide_file
    slide_xml = slide_path.read_text(encoding="utf-8")

    if verbose:
        print(f"  [INJECT] Processing slide: {slide_file}")
        print(f"    Source content: title={repr(source_slide.get('title', '')[:30])}, "
              f"subtitle={repr(source_slide.get('subtitle', '')[:20])}, "
              f"body_items={len(source_slide.get('body', []))}")

    # Infer background type from existing slide background
    use_dark = _slide_has_dark_bg(slide_xml, template_colors)
    title_color = template_colors["text_on_dark"] if use_dark else template_colors["primary"]
    body_color  = template_colors["text_on_dark"] if use_dark else template_colors["text_on_light"]

    if verbose:
        print(f"    Colors: title={title_color}, body={body_color}, dark_bg={use_dark}")

    # Font faces from theme (empty string = don't set, inherit from layout/master)
    title_latin_font = template_fonts.get("major_latin", "")
    title_ea_font    = template_fonts.get("major_ea", "")
    body_latin_font  = template_fonts.get("minor_latin", "")
    body_ea_font     = template_fonts.get("minor_ea", "")

    title = source_slide.get("title", "")
    subtitle = source_slide.get("subtitle", "")
    body = source_slide.get("body", [])
    # body_rich carries per-run formatting: [{text, bold, italic, size, color}, ...]
    body_rich: List[dict] = source_slide.get("body_rich", [])

    # ── Build body XML ─────────────────────────────────────────────────────────
    body_xml = ""
    if body:
        body_index_offset = 0
        effective_title = title
        if not title and body:
            effective_title = body[0]
            body_index_offset = 1

        if source_slide.get("type") == "title":
            body_to_use = body[1:] if subtitle else body
            rich_to_use = body_rich[1:] if (subtitle and body_rich) else body_rich
        else:
            body_to_use = body[body_index_offset:]
            rich_to_use = body_rich[body_index_offset:] if body_rich else []

        if body_to_use:
            body_xml = _build_body_xml(
                body_to_use,
                color=body_color,
                rich=rich_to_use,
                latin_font=body_latin_font,
                ea_font=body_ea_font,
                layout_ph_styles=layout_ph_styles,
            )
            if verbose:
                print(f"    Built body XML with {len(body_to_use)} items")

    # ── Try ET-based injection first (more reliable) ─────────────────────────
    modified, et_title_ok, et_subtitle_ok, et_body_ok = _inject_content_using_et(
        slide_xml, title, subtitle, body_xml,
        title_color, body_color,
        title_latin_font, title_ea_font,
        body_latin_font, body_ea_font,
        verbose,
    )

    if et_title_ok and et_subtitle_ok and (not body_xml or et_body_ok):
        if verbose:
            print(f"  [INJECT] ET injection successful: title={et_title_ok}, subtitle={et_subtitle_ok}, body={et_body_ok}")
        slide_path.write_text(modified, encoding="utf-8")
        return

    if verbose:
        print(f"  [INJECT] ET injection incomplete, falling back to regex")

    # ── Fallback: regex-based injection ──────────────────────────────────────
    modified = slide_xml

    # If title is empty but body has content, use body[0] as title
    effective_title = title
    body_index_offset = 0  # How many items to skip from body when injecting content
    if not title and body:
        effective_title = body[0]
        body_index_offset = 1  # Skip body[0] since it's now the title

    if effective_title:
        modified_after, injected = _replace_placeholder_text_with_log(
            modified, ["title", "ctrTitle"], effective_title,
            color=title_color,
            latin_font=title_latin_font,
            ea_font=title_ea_font,
            verbose=verbose,
            label="title",
        )
        modified = modified_after
        if injected and not et_title_ok:
            et_title_ok = True

    if subtitle:
        modified, injected = _replace_placeholder_text_with_log(
            modified, ["subTitle"], subtitle,
            color=body_color,
            latin_font=body_latin_font,
            ea_font=body_ea_font,
            verbose=verbose,
            label="subtitle",
        )
        if injected and not et_subtitle_ok:
            et_subtitle_ok = True

    if body_xml:
        modified, injected = _replace_placeholder_content_with_log(
            modified, ["body", "obj"], body_xml, verbose=verbose, label="body"
        )
        if injected and not et_body_ok:
            et_body_ok = True

    # ── Fallback: if the slide still looks empty, try any remaining placeholder ─
    # This handles edge cases where:
    # (a) A title-only slide with title that went into a layout lacking "title" ph type
    # (b) body content that couldn't find body/obj ph — try injecting into first
    #     available text placeholder
    if effective_title and modified == slide_xml:
        # Title didn't get placed; try any remaining placeholder types
        for fallback_types in (["body", "obj"], ["subTitle"], ["pic"]):
            result, injected = _replace_placeholder_text_with_log(
                modified, fallback_types, effective_title,
                color=title_color,
                latin_font=title_latin_font,
                ea_font=title_ea_font,
                verbose=verbose,
                label=f"title-fallback-{fallback_types[0]}",
            )
            if injected:
                modified = result
                if verbose:
                    print(f"    [FALLBACK] Injected title into {fallback_types} placeholder")
                break

    # Summary log
    if verbose:
        content_was_injected = (modified != slide_xml)
        print(f"  [INJECT] Summary: {'INJECTED' if content_was_injected else 'NO CHANGE'}, "
              f"title={et_title_ok or injected}, body={et_body_ok}")

    slide_path.write_text(modified, encoding="utf-8")


def _replace_placeholder_text_with_log(
    slide_xml: str,
    ph_types: List[str],
    new_text: str,
    color: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    size: Optional[int] = None,
    latin_font: str = "",
    ea_font: str = "",
    verbose: bool = False,
    label: str = "",
) -> Tuple[str, bool]:
    """Replace text content in a placeholder with logging. Returns (modified_xml, was_injected)."""
    if verbose:
        print(f"    [REGEX] Trying to inject {label}: {repr(new_text[:30])}...")

    original_xml = slide_xml
    modified = _replace_placeholder_text(
        slide_xml, ph_types, new_text,
        color=color, bold=bold, italic=italic, size=size,
        latin_font=latin_font, ea_font=ea_font,
    )

    was_injected = modified != slide_xml
    if verbose:
        if was_injected:
            print(f"    [REGEX] ✓ Injected {label}")
        else:
            print(f"    [REGEX] ✗ Failed to inject {label} (no matching placeholder)")

    return modified, was_injected


def _replace_placeholder_content_with_log(
    slide_xml: str,
    ph_types: List[str],
    new_content_xml: str,
    verbose: bool = False,
    label: str = "",
) -> Tuple[str, bool]:
    """Replace body content in a placeholder with logging. Returns (modified_xml, was_injected)."""
    if verbose:
        print(f"    [REGEX] Trying to inject {label} body content...")

    original_xml = slide_xml
    modified = _replace_placeholder_content(slide_xml, ph_types, new_content_xml)

    was_injected = modified != slide_xml
    if verbose:
        if was_injected:
            print(f"    [REGEX] ✓ Injected {label} body content")
        else:
            print(f"    [REGEX] ✗ Failed to inject {label} body (no matching placeholder)")

    return modified, was_injected


def _slide_has_dark_bg(slide_xml: str, template_colors: Dict[str, str]) -> bool:
    """Heuristically determine if a slide has a dark background.

    Detection order:
    1. Explicit solidFill hex colour inside <p:bg> — most reliable
    2. schemeClr reference inside <p:bg> → map via template_colors (dk2 → bg_dark)
    3. bgRef idx attribute — OOXML indices 1001+ are filled (not blank); we assume
       high-contrast templates use dark fills for idx ≥ 1002
    4. bg_dark colour present anywhere in the slide XML (loose fallback)
    Falls back to False (assume light) when undetermined.
    """
    # ── 1. Explicit srgbClr inside <p:bg> ────────────────────────────────────
    bg_m = re.search(
        r'<p:bg\b.*?<a:srgbClr val="([0-9A-Fa-f]{6})"',
        slide_xml, re.DOTALL
    )
    if bg_m:
        hex_color = bg_m.group(1).upper()
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        luminance = (r + g + b) / 3
        return luminance < 128

    # ── 2. schemeClr reference inside <p:bg> ─────────────────────────────────
    scheme_m = re.search(
        r'<p:bg\b.*?<a:schemeClr val="([^"]+)"',
        slide_xml, re.DOTALL
    )
    if scheme_m:
        scheme_val = scheme_m.group(1).lower()
        # dk1/dk2 are typically dark; lt1/lt2 are light; accent varies
        if scheme_val in ("dk1", "dk2"):
            return True
        if scheme_val in ("lt1", "lt2"):
            return False
        # For accent colours, check template_colors bg_dark
        bg_dark = template_colors.get("bg_dark", "").upper()
        if bg_dark:
            r = int(bg_dark[0:2], 16)
            g = int(bg_dark[2:4], 16)
            b = int(bg_dark[4:6], 16)
            return (r + g + b) / 3 < 128

    # ── 3. bgRef idx — template slides often use reference fills ─────────────
    bgref_m = re.search(r'<p:bgRef\b[^>]*idx="(\d+)"', slide_xml)
    if bgref_m:
        idx = int(bgref_m.group(1))
        # idx 1001 = no fill; 1002+ = solid fill from format scheme
        # We can't resolve the actual colour without the theme, but we know
        # that if the template's bg_dark is truly dark and idx ≥ 1002 we
        # probably have a filled (potentially dark) background.
        # For safety, also check the schemeClr child of bgRef.
        bgref_scheme_m = re.search(
            r'<p:bgRef\b.*?<a:schemeClr val="([^"]+)"',
            slide_xml, re.DOTALL
        )
        if bgref_scheme_m:
            sv = bgref_scheme_m.group(1).lower()
            if sv in ("dk1", "dk2"):
                return True
            if sv in ("lt1", "lt2"):
                return False

    # ── 4. bg_dark colour present anywhere in the slide XML ──────────────────
    bg_dark = template_colors.get("bg_dark", "000000").upper()
    if bg_dark and bg_dark in slide_xml.upper():
        return True

    return False


def _replace_placeholder_text(
    slide_xml: str,
    ph_types: List[str],
    new_text: str,
    color: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    size: Optional[int] = None,
    latin_font: str = "",
    ea_font: str = "",
) -> str:
    """Replace text content in a placeholder while applying the template color and fonts.

    Styling rules applied:
    - Color:      always set to *color* (the template palette value)
    - Latin font: set when *latin_font* is non-empty (from theme majorFont/minorFont)
    - EA font:    set when *ea_font* is non-empty (from theme majorFont/minorFont)
    - Bold / italic / size: carried forward from the source slide when provided
    - Language tag: auto-detected per run (zh-CN when >10% CJK, en-US otherwise)
    - bodyPr:     the original <a:bodyPr> from the template placeholder is preserved;
                  only the <a:p> paragraphs are replaced — NOT bodyPr or lstStyle
    """
    type_pattern = "|".join(ph_types)
    lang = _detect_lang(new_text)

    def replace_sp(m):
        sp_xml = m.group(0)
        if not re.search(rf'<p:ph[^>]*type="(?:{type_pattern})"', sp_xml):
            return sp_xml

        # Build rPr attributes
        rpr_attrs = f'lang="{lang}" dirty="0"'
        if bold:
            rpr_attrs += ' b="1"'
        if italic:
            rpr_attrs += ' i="1"'
        if size:
            # OOXML sz is in hundredths of a point (e.g. 2400 = 24pt)
            sz_val = size * 100 if size < 1000 else size
            rpr_attrs += f' sz="{sz_val}"'

        # Build rPr children: colour first, then font declarations
        rpr_children = ""
        if color:
            rpr_children += f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        if latin_font:
            rpr_children += f'<a:latin typeface="{latin_font}"/>'
        if ea_font:
            rpr_children += f'<a:ea typeface="{ea_font}"/>'

        def clear_txbody(tbm):
            # group(1): everything up to and including </a:lstStyle>
            # group(2): closing </a:txBody>
            # → replace ONLY the paragraph content; bodyPr and lstStyle are kept intact
            before = tbm.group(1)
            after  = tbm.group(2)
            new_para = (
                f'<a:p>'
                f'<a:r><a:rPr {rpr_attrs}>{rpr_children}</a:rPr>'
                f'<a:t>{_escape_xml(new_text)}</a:t></a:r>'
                f'</a:p>'
            )
            return before + new_para + after

        new_sp = re.sub(
            r'(<a:txBody>.*?</a:lstStyle>)(.*?)(</a:txBody>)',
            lambda tbm: clear_txbody(tbm),
            sp_xml,
            count=1,
            flags=re.DOTALL,
        )

        if new_sp == sp_xml:
            # Fallback: txBody has no lstStyle — inject a minimal structure BUT
            # PRESERVE the existing <a:bodyPr> if present (don't replace with empty one)
            existing_bodypr_m = re.search(r'<a:bodyPr\b[^>]*(?:/>|>.*?</a:bodyPr>)', sp_xml, re.DOTALL)
            bodypr_xml = existing_bodypr_m.group(0) if existing_bodypr_m else "<a:bodyPr/>"
            new_sp = re.sub(
                r'<a:txBody>.*?</a:txBody>',
                lambda _: (
                    f'<a:txBody>'
                    f'{bodypr_xml}<a:lstStyle/>'
                    f'<a:p><a:r><a:rPr {rpr_attrs}>{rpr_children}</a:rPr>'
                    f'<a:t>{_escape_xml(new_text)}</a:t></a:r></a:p>'
                    f'</a:txBody>'
                ),
                sp_xml,
                count=1,
                flags=re.DOTALL,
            )

        return new_sp

    return re.sub(r'<p:sp\b.*?</p:sp>', replace_sp, slide_xml, flags=re.DOTALL)


def _replace_placeholder_content(slide_xml: str, ph_types: List[str], new_content_xml: str) -> str:
    """Replace entire text body of a placeholder with new XML content.

    Handles two cases:
    1. Normal: txBody contains bodyPr + lstStyle → replace everything after lstStyle
    2. Fallback: txBody has no lstStyle (freshly created placeholder from layout
       extraction) → replace all <a:p> paragraphs while keeping bodyPr intact

    Also supports matching placeholders without an explicit type attribute but
    with idx≥1 (the implicit body placeholder that OOXML allows).
    """
    type_pattern = "|".join(ph_types)

    def replace_sp(m):
        sp_xml = m.group(0)
        # Primary match: explicit type attribute
        has_type_match = bool(re.search(rf'<p:ph[^>]*type="(?:{type_pattern})"', sp_xml))
        # Secondary match: no type attr but has idx≥1 (implicit body placeholder)
        # Only activate for body/obj ph_types
        has_implicit_body = False
        if not has_type_match and ("body" in ph_types or "obj" in ph_types):
            ph_no_type = re.search(r'<p:ph\b(?![^>]*type=)[^>]*/>', sp_xml)
            has_implicit_body = bool(ph_no_type)

        if not has_type_match and not has_implicit_body:
            return sp_xml

        # Case 1: txBody has lstStyle — replace content after it
        new_sp = re.sub(
            r'(<a:txBody>.*?</a:lstStyle>).*?(</a:txBody>)',
            lambda tm: tm.group(1) + new_content_xml + tm.group(2),
            sp_xml,
            count=1,
            flags=re.DOTALL,
        )
        if new_sp != sp_xml:
            return new_sp

        # Case 2: txBody has no lstStyle — preserve bodyPr, replace all paragraphs
        existing_bodypr_m = re.search(r'<a:bodyPr\b[^>]*(?:/>|>.*?</a:bodyPr>)', sp_xml, re.DOTALL)
        bodypr_xml = existing_bodypr_m.group(0) if existing_bodypr_m else "<a:bodyPr/>"
        new_sp = re.sub(
            r'<a:txBody>.*?</a:txBody>',
            lambda _: (
                f'<a:txBody>'
                f'{bodypr_xml}'
                f'<a:lstStyle/>'
                f'{new_content_xml}'
                f'</a:txBody>'
            ),
            sp_xml,
            count=1,
            flags=re.DOTALL,
        )
        return new_sp

    return re.sub(r'<p:sp\b.*?</p:sp>', replace_sp, slide_xml, flags=re.DOTALL)


def _build_body_xml(
    lines: List[str],
    color: Optional[str] = None,
    rich: Optional[List[dict]] = None,
    latin_font: str = "",
    ea_font: str = "",
    layout_ph_styles: Optional[Dict[str, dict]] = None,
) -> str:
    """Build XML paragraphs from a list of text lines.

    Each line becomes one paragraph.  Bullet marker is NOT hard-coded —
    we omit `<a:buChar>` so the placeholder inherits the template layout's
    list style (which may use numbers, custom bullets, or no bullets at all).
    Forcing &#x2022; here would override the template's intended style.

    Per-run formatting from *rich* (body_rich field) is applied when available:
    - bold   → <a:rPr b="1">
    - italic → <a:rPr i="1">
    - size   → <a:rPr sz="N"> (converted to hundredths-of-a-point)

    Font faces from the template's font scheme are applied when provided:
    - latin_font → <a:latin typeface="..."/>  (minorFont for body text)
    - ea_font    → <a:ea typeface="..."/>     (East Asian font for CJK text)
    Both are skipped when empty so PowerPoint inherits from the layout/master.

    Color is always overridden with the template palette value so text matches
    the template's visual identity regardless of the source PPT's colours.

    Language tag is auto-detected per run (zh-CN / en-US) to avoid mistagging
    English content with Chinese locale (which breaks spell-check in PowerPoint).

    Layout placeholder styles (from layout_ph_styles) are applied:
    - lvl1pPr_xml: default paragraph formatting (indent, alignment, spacing)
    - defRPr_xml: default run properties (font size, bold, etc.)
    - bodyPr_attrs: text body properties (wrap, anchor, etc.)
    """
    # Build a quick lookup: rich item index → formatting dict
    rich_lookup: Dict[int, dict] = {}
    if rich:
        for i, r in enumerate(rich):
            rich_lookup[i] = r

    # Extract layout style info for body placeholder
    body_style = (layout_ph_styles or {}).get("body", {})
    lvl1pPr_xml = body_style.get("lvl1pPr_xml", "")
    defRPr_xml = body_style.get("defRPr_xml", "")
    default_sz = body_style.get("default_sz", 0)
    bodyPr_attrs = body_style.get("bodyPr_attrs", "")

    # Pre-build shared XML fragments
    rpr_color_xml = (
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        if color else ""
    )
    font_xml = ""
    if latin_font:
        font_xml += f'<a:latin typeface="{latin_font}"/>'
    if ea_font:
        font_xml += f'<a:ea typeface="{ea_font}"/>'

    paragraphs = []
    for i, line in enumerate(lines):
        fmt = rich_lookup.get(i, {})
        bold   = fmt.get("bold", False)
        italic = fmt.get("italic", False)
        size   = fmt.get("size")

        lang = _detect_lang(line)

        # Build rPr attributes
        rpr_attrs = f'lang="{lang}" dirty="0"'
        if bold:
            rpr_attrs += ' b="1"'
        if italic:
            rpr_attrs += ' i="1"'
        if size:
            sz_val = size * 100 if size < 1000 else size
            rpr_attrs += f' sz="{sz_val}"'
        elif default_sz and not bold:
            # Use layout's default size if no explicit size from source
            rpr_attrs += f' sz="{default_sz}"'

        escaped = _escape_xml(line)
        rpr_children = rpr_color_xml + font_xml

        # Build paragraph with layout's default paragraph style if available
        # Clean up lvl1pPr to remove conflicting attributes that we'll set explicitly
        ppr_xml = ""
        if lvl1pPr_xml:
            # Remove buChar/buAutoNum from lvl1pPr to avoid double bullets
            clean_lvl1 = re.sub(r'<a:buChar[^/]*/>', '', lvl1pPr_xml)
            clean_lvl1 = re.sub(r'<a:buAutoNum[^/]*/>', '', clean_lvl1)
            ppr_xml = re.sub(r'<a:lvl1pPr\b', '<a:pPr', clean_lvl1)
            # Remove trailing /> or </a:lvl1pPr> to make it an opening tag
            ppr_xml = re.sub(r'\s*/>', '>', ppr_xml)
            ppr_xml = re.sub(r'</a:lvl1pPr>', '', ppr_xml)

        # No <a:pPr><a:buChar .../></a:pPr> — inherit template list style
        para_content = ""
        if ppr_xml:
            para_content = f'<a:p>{ppr_xml}'
        else:
            para_content = '<a:p>'

        paragraphs.append(
            f'{para_content}'
            f'<a:r>'
            f'<a:rPr {rpr_attrs}>{rpr_children}</a:rPr>'
            f'<a:t>{escaped}</a:t>'
            f'</a:r>'
            f'</a:p>'
        )
    return "\n".join(paragraphs)


def _escape_xml(text: str) -> str:
    """Escape special XML characters."""
    return (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def _migrate_animations(
    source_unpacked_dir: Path,
    source_slide_file: str,
    dest_unpacked_dir: Path,
    dest_slide_file: str,
    verbose: bool,
) -> None:
    """Copy <p:timing> animation block from source slide into destination slide.

    PPTX stores entrance/exit/emphasis animations inside a <p:timing> element
    at the end of each slide XML.  We extract that block from the source and
    inject it verbatim into the destination, replacing any existing <p:timing>.
    Shape IDs inside <p:timing> reference source shapes — since we clear all
    source shapes during content injection, the animation targets may become
    orphaned.  This is an acceptable trade-off: users get the animation timing
    structure and can re-link targets in PowerPoint if needed.
    """
    src_path = source_unpacked_dir / "ppt" / "slides" / source_slide_file
    dst_path = dest_unpacked_dir  / "ppt" / "slides" / dest_slide_file

    if not src_path.exists() or not dst_path.exists():
        return

    src_xml = src_path.read_text(encoding="utf-8")

    # Extract <p:timing>…</p:timing> block from source
    timing_m = re.search(r'<p:timing\b.*?</p:timing>', src_xml, re.DOTALL)
    if not timing_m:
        return  # No animations in this source slide

    timing_xml = timing_m.group(0)

    dst_xml = dst_path.read_text(encoding="utf-8")

    # Remove any existing <p:timing> in the destination
    dst_xml = re.sub(r'<p:timing\b.*?</p:timing>', '', dst_xml, flags=re.DOTALL)

    # Insert before the closing </p:sld> tag
    if '</p:sld>' in dst_xml:
        dst_xml = dst_xml.replace('</p:sld>', timing_xml + '\n</p:sld>')
        dst_path.write_text(dst_xml, encoding="utf-8")
        if verbose:
            print(f"    + Migrated animations to {dest_slide_file}")


def _migrate_notes(
    source_unpacked_dir: Path,
    source_slide_file: str,
    dest_unpacked_dir: Path,
    dest_slide_file: str,
    verbose: bool,
) -> None:
    """Copy the speaker notes from the source slide to the destination slide.

    Notes are stored in ppt/notesSlides/notesSlideN.xml, referenced by a
    Relationship entry in ppt/slides/_rels/slideN.xml.rels.
    Steps:
      1. Find the notesSlide file for the source slide via its .rels file.
      2. Copy that notesSlide XML into the destination unpacked dir.
      3. Add a Relationship entry in the destination slide's .rels file.
      4. Register the new notesSlide in [Content_Types].xml.
    """
    src_slides_dir    = source_unpacked_dir / "ppt" / "slides"
    src_notes_dir     = source_unpacked_dir / "ppt" / "notesSlides"
    dst_slides_dir    = dest_unpacked_dir   / "ppt" / "slides"
    dst_notes_dir     = dest_unpacked_dir   / "ppt" / "notesSlides"
    dst_rels_dir      = dst_slides_dir / "_rels"

    src_rels_path = src_slides_dir / "_rels" / f"{source_slide_file}.rels"
    if not src_rels_path.exists():
        return

    src_rels = src_rels_path.read_text(encoding="utf-8")

    # Find the notesSlide relationship in source .rels
    notes_m = re.search(
        r'<Relationship[^>]+Type="[^"]*notesSlide[^"]*"[^>]+Target="([^"]+)"',
        src_rels,
    )
    if not notes_m:
        return  # No notes for this slide

    notes_target = notes_m.group(1)  # e.g. "../notesSlides/notesSlide1.xml"
    notes_filename = Path(notes_target).name  # e.g. "notesSlide1.xml"
    src_notes_path = src_notes_dir / notes_filename

    if not src_notes_path.exists():
        return

    # Determine destination notes filename (avoid collision)
    dst_notes_dir.mkdir(exist_ok=True)
    dst_notes_rels_dir = dst_notes_dir / "_rels"
    dst_notes_rels_dir.mkdir(exist_ok=True)

    existing_notes = [
        int(m.group(1))
        for f in dst_notes_dir.glob("notesSlide*.xml")
        if (m := re.match(r"notesSlide(\d+)\.xml", f.name))
    ]
    next_notes_num = max(existing_notes) + 1 if existing_notes else 1
    dst_notes_filename = f"notesSlide{next_notes_num}.xml"
    dst_notes_path = dst_notes_dir / dst_notes_filename

    # Copy and fix the notesSlide XML (update the r:id slideRef if present)
    notes_xml = src_notes_path.read_text(encoding="utf-8")
    dst_notes_path.write_text(notes_xml, encoding="utf-8")

    # Copy notesSlide .rels file if it exists (references the slide and notesMaster)
    src_notes_rels = src_notes_dir / "_rels" / f"{notes_filename}.rels"
    if src_notes_rels.exists():
        src_notes_rels_xml = src_notes_rels.read_text(encoding="utf-8")
        # Point the slide relationship back to the new dest slide
        src_notes_rels_xml = re.sub(
            r'(Target="\.\./slides/)[^"]+(")',
            rf'\g<1>{dest_slide_file}\2',
            src_notes_rels_xml,
        )
        (dst_notes_rels_dir / f"{dst_notes_filename}.rels").write_text(
            src_notes_rels_xml, encoding="utf-8"
        )

    # Add notesSlide to [Content_Types].xml
    ct_path = dest_unpacked_dir / "[Content_Types].xml"
    ct = ct_path.read_text(encoding="utf-8")
    notes_ct = (
        f'<Override PartName="/ppt/notesSlides/{dst_notes_filename}" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
    )
    if f"/ppt/notesSlides/{dst_notes_filename}" not in ct:
        ct = ct.replace("</Types>", f"  {notes_ct}\n</Types>")
        ct_path.write_text(ct, encoding="utf-8")

    # Add Relationship in destination slide's .rels file
    dst_slide_rels_path = dst_rels_dir / f"{dest_slide_file}.rels"
    if dst_slide_rels_path.exists():
        dest_rels = dst_slide_rels_path.read_text(encoding="utf-8")
    else:
        dest_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            '</Relationships>'
        )

    # Remove any old notes relationship first
    dest_rels = re.sub(
        r'\s*<Relationship[^>]*notesSlide[^>]*/>\s*', '\n', dest_rels
    )

    # Find next rId
    rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', dest_rels)]
    next_rid_num = max(rids) + 1 if rids else 10
    notes_rid = f"rId{next_rid_num}"

    new_rel = (
        f'<Relationship Id="{notes_rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" '
        f'Target="../notesSlides/{dst_notes_filename}"/>'
    )
    dest_rels = dest_rels.replace("</Relationships>", f"  {new_rel}\n</Relationships>")
    dst_slide_rels_path.write_text(dest_rels, encoding="utf-8")

    if verbose:
        print(f"    + Migrated notes → {dst_notes_filename}")


def _build_new_slides(
    source_slides: List[dict],
    slide_mapping: List[dict],
    template_layouts: List[dict],
    unpacked_dir: Path,
    source_images_dir: Path,
    keep_notes: bool,
    template_colors: Dict[str, str],
    template_fonts: Dict[str, str],
    source_unpacked_dir: Path,
    source_slide_file_map: Dict[int, str],
    verbose: bool,
    skip_animations: bool = False,
    beautify: bool = False,
    template_theme: Optional[dict] = None,
    dark_mode: bool = False,
) -> List[Tuple[str, str]]:
    """Create all new slides by duplicating template slides and injecting content."""
    # Build lookup including auto-generated TOC slides (source_index=0)
    mapping_lookup = {}
    auto_toc_mappings = []
    for sm in slide_mapping:
        if sm.get("auto_generated_toc"):
            auto_toc_mappings.append(sm)
        else:
            mapping_lookup[sm["source_index"]] = sm
    
    layout_lookup = {l["layout_file"]: l for l in template_layouts}
    new_slides = []
    auto_toc_processed = False
    # Track layout variant usage for visual variety
    layout_streak: list = []
    slide_index = 0

    for source_slide in source_slides:
        idx = source_slide.get("index", 0)
        
        # Check if we need to insert auto-generated TOC before this slide
        if not auto_toc_processed and auto_toc_mappings:
            # Insert TOC before first non-title slide (index > 1)
            if idx > 1 and source_slide.get("type") not in ("title", "section"):
                for toc_mapping in auto_toc_mappings:
                    layout_file = toc_mapping["template_layout"]
                    layout_info = layout_lookup.get(layout_file, {})
                    layout_ph_styles = layout_info.get("ph_styles", {})
                    
                    # Create TOC slide
                    template_source_slide = _find_template_slide_for_layout(unpacked_dir, layout_file)
                    if template_source_slide:
                        new_slide_file, rid = _duplicate_template_slide(unpacked_dir, template_source_slide)
                    else:
                        new_slide_file, rid = _create_slide_from_layout(unpacked_dir, layout_file)
                    
                    # Inject TOC content
                    toc_content = toc_mapping.get("toc_content", {})
                    _inject_content_into_slide(
                        unpacked_dir, new_slide_file, toc_content,
                        template_colors, template_fonts, layout_ph_styles,
                        verbose,
                    )

                    # Apply deep beautification to TOC slides too
                    if beautify and template_theme:
                        toc_slide_path = unpacked_dir / "ppt" / "slides" / new_slide_file
                        _beautify_slide_with_template_theme(
                            slide_path=toc_slide_path,
                            theme=template_theme,
                            slide_type="content",
                            body_items=toc_content.get("body", []),
                            layout_streak=layout_streak,
                            slide_index=slide_index,
                            dark_mode=dark_mode,
                            restructure=True,
                            verbose=verbose,
                        )
                        slide_index += 1

                    if verbose:
                        print(f"  [AUTO-TOC] Created {new_slide_file} with auto-generated table of contents")
                    
                    new_slides.append((new_slide_file, rid))
                auto_toc_processed = True
        
        if "error" in source_slide:
            print(f"  Skipping slide {idx} (parse error)")
            continue

        sm = mapping_lookup.get(idx)
        if not sm:
            print(f"  Skipping slide {idx} (no mapping)")
            continue

        layout_file = sm["template_layout"]

        # Get layout info (includes ph_styles extracted from layout XML)
        layout_info = layout_lookup.get(layout_file, {})
        layout_ph_styles = layout_info.get("ph_styles", {})

        # Find or create slide using this layout
        template_source_slide = _find_template_slide_for_layout(unpacked_dir, layout_file)

        if template_source_slide:
            new_slide_file, rid = _duplicate_template_slide(unpacked_dir, template_source_slide)
        else:
            new_slide_file, rid = _create_slide_from_layout(unpacked_dir, layout_file)

        # Merge content from custom shapes if source_slide is missing content
        src_file = source_slide_file_map.get(idx)
        if src_file:
            source_slide = _merge_custom_shape_content(
                source_slide, source_unpacked_dir, src_file
            )

        # Inject content with full template style (colors + fonts + layout ph_styles)
        _inject_content_into_slide(
            unpacked_dir, new_slide_file, source_slide,
            template_colors, template_fonts, layout_ph_styles,
            verbose,
        )

        # Apply template colors to non-placeholder elements (shapes, tables, etc.)
        slide_path = unpacked_dir / "ppt" / "slides" / new_slide_file
        _apply_template_colors_to_slide(slide_path, template_colors, verbose)

        # Apply deep beautification (full pipeline: backgrounds, colors, fonts, layout variants, images, tables)
        if beautify and template_theme:
            body_items = source_slide.get("body", [])
            slide_type = source_slide.get("type", "content")
            beautify_notes = _beautify_slide_with_template_theme(
                slide_path=slide_path,
                theme=template_theme,
                slide_type=slide_type,
                body_items=body_items,
                layout_streak=layout_streak,
                slide_index=slide_index,
                dark_mode=dark_mode,
                restructure=True,
                verbose=verbose,
            )
            slide_index += 1
        else:
            # Fallback: basic layout beautification only
            beautify_notes = _beautify_slide_layout(slide_path, verbose)
            slide_index += 1

        # Migrate animations from source slide (with enhanced ID mapping)
        if src_file and not skip_animations:
            migration_result = migrate_animations_with_id_mapping(
                source_unpacked_dir, src_file,
                unpacked_dir, new_slide_file,
                id_mapping=None,  # Auto-detect ID mapping
                verbose=verbose,
            )
            if verbose and migration_result["animations_migrated"]:
                print(f"    ✓ {migration_result['updated_targets']} animation target(s) updated")
            # Migrate speaker notes if requested
            if keep_notes:
                _migrate_notes(
                    source_unpacked_dir, src_file,
                    unpacked_dir, new_slide_file,
                    verbose,
                )

        if verbose:
            print(f"  Slide {idx}: created {new_slide_file} using layout {layout_file}")

        new_slides.append((new_slide_file, rid))

    return new_slides


def _update_presentation_order(unpacked_dir: Path, new_slides: List[Tuple[str, str]]) -> None:
    """Update presentation.xml to only contain the new slides in order."""
    pres_path = unpacked_dir / "ppt" / "presentation.xml"
    pres_content = pres_path.read_text(encoding="utf-8")

    # Build new sldIdLst content
    # Find the next available slide ID
    existing_ids = [int(m) for m in re.findall(r'<p:sldId[^>]*id="(\d+)"', pres_content)]
    next_id = max(existing_ids) + 100 if existing_ids else 256

    new_sld_entries = []
    for i, (slide_file, rid) in enumerate(new_slides):
        slide_id = next_id + i
        new_sld_entries.append(f'      <p:sldId id="{slide_id}" r:id="{rid}"/>')

    new_sld_id_lst = "    <p:sldIdLst>\n" + "\n".join(new_sld_entries) + "\n    </p:sldIdLst>"

    # Replace existing sldIdLst
    pres_content = re.sub(
        r'<p:sldIdLst>.*?</p:sldIdLst>',
        new_sld_id_lst,
        pres_content,
        flags=re.DOTALL,
    )

    pres_path.write_text(pres_content, encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Apply a template's visual style to a source PPT's content"
    )
    parser.add_argument("source", help="Source PPTX file (your content)")
    parser.add_argument("template", help="Template PPTX file (your desired style)")
    parser.add_argument("output", help="Output PPTX file")
    parser.add_argument("--mapping", help="JSON file with manual slide mapping")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print mapping plan only — do not create output file",
    )
    parser.add_argument(
        "--save-mapping",
        metavar="FILE",
        help="Save auto-generated mapping to a JSON file (use with --dry-run to review first)",
    )
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")
    parser.add_argument("--no-notes", dest="keep_notes", action="store_false",
                       help="Don't preserve speaker notes")
    parser.add_argument(
        "--skip-animations",
        action="store_true",
        help="Skip animation migration (useful when animations are incompatible with new layout)",
    )
    parser.add_argument(
        "--interactive",
        action="store_true",
        help="Interactive mode: confirm when high-risk mappings are detected",
    )
    parser.add_argument(
        "--beautify",
        action="store_true",
        help="Apply AI beautification pass after template application (redesigns layout, colors, fonts)",
    )
    parser.add_argument(
        "--beautify-theme",
        metavar="THEME",
        help="Theme design system for beautification (executive/tech/creative/warm/minimal/bold/nature/ocean/elegant/modern/sunset/forest). "
             "Colors come from template; this sets fonts/sizes/layout style. Default: minimal",
    )
    parser.add_argument(
        "--dark-mode",
        action="store_true",
        help="Apply dark background to title/section slides during beautification",
    )
    args = parser.parse_args()

    apply_template(
        args.source,
        args.template,
        args.output,
        mapping_file=args.mapping,
        save_mapping=args.save_mapping,
        dry_run=args.dry_run,
        verbose=args.verbose,
        keep_notes=args.keep_notes,
        skip_animations=args.skip_animations,
        interactive=args.interactive,
        beautify=args.beautify,
        beautify_theme=args.beautify_theme,
        dark_mode=args.dark_mode,
    )
