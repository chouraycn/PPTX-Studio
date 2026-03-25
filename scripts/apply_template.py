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
from pathlib import Path
from typing import Optional, List, Dict, Tuple

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
    "title_slide": ["title slide", "title, subtitle", "title only", "ctrTitle"],
    "section_header": ["section header", "section", "divider", "blank"],
    "content_slide": ["title and content", "content", "object"],
    "two_column": ["two content", "comparison", "2 column", "two column"],
    "image_text": ["picture with caption", "picture", "image", "photo"],
    "list_content": ["title and content", "content", "bulleted list"],
    "chart_content": ["title and content", "content"],
    "table_content": ["title and content", "content"],
    "quote_slide": ["blank", "title only", "quote"],
    "conclusion": ["blank", "title only", "title slide"],
    "full_image": ["blank", "picture"],
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
    if verbose:
        print(f"Template has {len(template_layouts)} layouts:")
        for l in template_layouts:
            print(f"  {l['layout_file']}: {l['layout_name']} ({l['detected_type']})")

    # Load or create slide mapping
    if mapping_file and Path(mapping_file).exists():
        with open(mapping_file) as f:
            slide_mapping = json.load(f)
        print(f"Using manual mapping from {mapping_file}")
    else:
        slide_mapping = _auto_map_slides(source_content["slides"], template_layouts, verbose)

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

        # Get template slide structure
        template_slide_order = _get_presentation_slide_order(unpacked_dir)
        if verbose:
            print(f"Template has {len(template_slide_order)} slides in sldIdLst")

        # Extract source images to temp dir
        source_images_dir = tmp_path / "source_images"
        source_images_dir.mkdir()
        _extract_source_images(source_pptx, source_images_dir)

        # Build new slide list
        new_slides = _build_new_slides(
            source_content["slides"],
            slide_mapping,
            template_layouts,
            unpacked_dir,
            source_images_dir,
            keep_notes,
            verbose,
        )

        # Update presentation.xml with new slide order
        _update_presentation_order(unpacked_dir, new_slides)

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
    for sm in slide_mapping:
        idx = sm["source_index"]
        title = source_titles.get(idx, "")[:28]
        src_type = sm.get("source_type", "?")[:14]
        tmpl = sm.get("template_layout", "?")
        tmpl_type = sm.get("template_type", "")
        print(f"  {idx:<4} {title:<30} {src_type:<16}   {tmpl} [{tmpl_type}]")
    print("─" * 60)
    print(f"  Total: {len(slide_mapping)} slides")
    print("─" * 60 + "\n")


def _analyze_template_layouts(template_pptx: str) -> List[dict]:
    """Analyze which layout types are available in the template."""
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

            layouts.append({
                "layout_file": layout_file,
                "layout_name": layout_name,
                "detected_type": detected_type,
                "placeholder_types": placeholder_types,
                "has_body": "body" in placeholder_types or "obj" in placeholder_types,
                "has_title": "title" in placeholder_types or "ctrTitle" in placeholder_types,
            })

    return layouts


def _detect_layout_type(layout_name: str, layout_xml: str) -> str:
    """Classify a template layout type."""
    name_lower = layout_name.lower()

    for layout_type, keywords in LAYOUT_TYPE_KEYWORDS.items():
        if any(kw in name_lower for kw in keywords):
            return layout_type

    # Fallback: check placeholder count
    ph_types = re.findall(r'<p:ph[^>]*type="([^"]*)"', layout_xml)
    if "ctrTitle" in ph_types:
        return "title_slide"
    if "body" in ph_types or "obj" in ph_types:
        return "content_slide"
    return "content_slide"


def _auto_map_slides(source_slides: List[dict], template_layouts: List[dict],
                    verbose: bool) -> List[dict]:
    """Automatically map source slides to template layouts."""
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

        # Find best matching layout
        chosen_layout = None

        # Try exact type match
        if hint in layouts_by_type and layouts_by_type[hint]:
            chosen_layout = layouts_by_type[hint][0]

        # Try source_type if hint didn't work
        if not chosen_layout and source_type in layouts_by_type:
            chosen_layout = layouts_by_type[source_type][0]

        # Try fallbacks
        if not chosen_layout:
            for fb in fallback_order:
                if fb in layouts_by_type:
                    chosen_layout = layouts_by_type[fb][0]
                    break

        # Ultimate fallback: first layout
        if not chosen_layout and template_layouts:
            chosen_layout = template_layouts[0]

        if chosen_layout:
            mapping.append({
                "source_index": slide["index"],
                "source_type": source_type,
                "template_layout": chosen_layout["layout_file"],
                "template_type": chosen_layout["detected_type"],
                "layout_name": chosen_layout["layout_name"],
            })
        else:
            print(f"Warning: No layout found for slide {slide['index']}", file=sys.stderr)

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


def _get_layout_rids(unpacked_dir: Path) -> Dict[str, str]:
    """Get layout file -> rId mapping from slideLayouts relationships."""
    rels_path = unpacked_dir / "ppt" / "slideLayouts"
    result = {}
    # Get from presentation rels
    pres_rels = (unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels").read_text(encoding="utf-8")
    for m in re.finditer(r'<Relationship[^>]+Id="([^"]+)"[^>]+Target="slideLayouts/([^"]+)"', pres_rels):
        result[m.group(2)] = m.group(1)
    return result


def _create_slide_from_layout(unpacked_dir: Path, layout_file: str) -> Tuple[str, str]:
    """Create a new blank slide using a layout. Returns (slide_file, rId)."""
    slides_dir = unpacked_dir / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"
    rels_dir.mkdir(exist_ok=True)

    # Get next slide number
    existing = [int(m.group(1)) for f in slides_dir.glob("slide*.xml")
                if (m := re.match(r"slide(\d+)\.xml", f.name))]
    next_num = max(existing) + 1 if existing else 1
    slide_file = f"slide{next_num}.xml"

    # Create blank slide XML
    slide_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
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
      </p:grpSpPr>
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


def _inject_content_into_slide(
    unpacked_dir: Path,
    slide_file: str,
    source_slide: dict,
    verbose: bool,
) -> None:
    """Replace placeholder content in a template slide with source content."""
    slide_path = unpacked_dir / "ppt" / "slides" / slide_file
    slide_xml = slide_path.read_text(encoding="utf-8")

    title = source_slide.get("title", "")
    subtitle = source_slide.get("subtitle", "")
    body = source_slide.get("body", [])

    modified = slide_xml

    # Replace title placeholder
    if title:
        modified = _replace_placeholder_text(modified, ["title", "ctrTitle"], title)

    # Replace subtitle placeholder
    if subtitle:
        modified = _replace_placeholder_text(modified, ["subTitle"], subtitle)
    elif body and not subtitle:
        # Use first body line as subtitle if no explicit subtitle
        if source_slide.get("type") == "title":
            modified = _replace_placeholder_text(modified, ["subTitle"], body[0] if body else "")

    # Replace body/content placeholder with bullet list
    if body:
        if source_slide.get("type") == "title":
            body_to_use = body[1:] if subtitle else body  # Skip first if used as subtitle
        else:
            body_to_use = body

        if body_to_use:
            body_xml = _build_body_xml(body_to_use)
            modified = _replace_placeholder_content(modified, ["body", "obj"], body_xml)

    slide_path.write_text(modified, encoding="utf-8")


def _replace_placeholder_text(slide_xml: str, ph_types: list[str], new_text: str) -> str:
    """Replace text content in a placeholder while preserving formatting."""
    type_pattern = "|".join(ph_types)

    def replace_sp(m):
        sp_xml = m.group(0)
        if not re.search(rf'<p:ph[^>]*type="(?:{type_pattern})"', sp_xml):
            return sp_xml
        # Replace the text content of the first <a:t> in the placeholder
        # Keep the first <a:p> structure, replace all text
        new_sp = re.sub(
            r'(<a:txBody>.*?<a:p[^>]*>.*?<a:r[^>]*>.*?<a:t>)[^<]*(</a:t>)',
            lambda tm: tm.group(1) + _escape_xml(new_text) + tm.group(2),
            sp_xml,
            count=1,
            flags=re.DOTALL,
        )
        if new_sp == sp_xml:
            # If no <a:r> found, inject a basic text run
            new_sp = re.sub(
                r'(<a:txBody>.*?<a:bodyPr[^>]*/?>.*?<a:p[^>]*>)',
                lambda tm: tm.group(0) + f'<a:r><a:rPr lang="en-US" dirty="0"/><a:t>{_escape_xml(new_text)}</a:t></a:r>',
                sp_xml,
                count=1,
                flags=re.DOTALL,
            )
        return new_sp

    return re.sub(r'<p:sp\b.*?</p:sp>', replace_sp, slide_xml, flags=re.DOTALL)


def _replace_placeholder_content(slide_xml: str, ph_types: list[str], new_content_xml: str) -> str:
    """Replace entire text body of a placeholder with new XML content."""
    type_pattern = "|".join(ph_types)

    def replace_sp(m):
        sp_xml = m.group(0)
        if not re.search(rf'<p:ph[^>]*type="(?:{type_pattern})"', sp_xml):
            return sp_xml
        # Replace everything inside <a:txBody> after <a:bodyPr> and <a:lstStyle>
        new_sp = re.sub(
            r'(<a:txBody>.*?</a:lstStyle>).*?(</a:txBody>)',
            lambda tm: tm.group(1) + new_content_xml + tm.group(2),
            sp_xml,
            count=1,
            flags=re.DOTALL,
        )
        return new_sp

    return re.sub(r'<p:sp\b.*?</p:sp>', replace_sp, slide_xml, flags=re.DOTALL)


def _build_body_xml(lines: list[str]) -> str:
    """Build XML paragraphs from a list of text lines."""
    paragraphs = []
    for line in lines:
        escaped = _escape_xml(line)
        paragraphs.append(
            f'<a:p><a:pPr><a:buChar char="&#x2022;"/></a:pPr>'
            f'<a:r><a:rPr lang="en-US" dirty="0"/><a:t>{escaped}</a:t></a:r></a:p>'
        )
    return "\n".join(paragraphs)


def _escape_xml(text: str) -> str:
    """Escape special XML characters."""
    return (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def _build_new_slides(
    source_slides: List[dict],
    slide_mapping: List[dict],
    template_layouts: List[dict],
    unpacked_dir: Path,
    source_images_dir: Path,
    keep_notes: bool,
    verbose: bool,
) -> List[Tuple[str, str]]:
    """Create all new slides by duplicating template slides and injecting content."""
    mapping_lookup = {sm["source_index"]: sm for sm in slide_mapping}
    layout_lookup = {l["layout_file"]: l for l in template_layouts}
    new_slides = []

    for source_slide in source_slides:
        idx = source_slide.get("index", 0)
        if "error" in source_slide:
            print(f"  Skipping slide {idx} (parse error)")
            continue

        sm = mapping_lookup.get(idx)
        if not sm:
            print(f"  Skipping slide {idx} (no mapping)")
            continue

        layout_file = sm["template_layout"]

        # Find or create slide using this layout
        template_source_slide = _find_template_slide_for_layout(unpacked_dir, layout_file)

        if template_source_slide:
            new_slide_file, rid = _duplicate_template_slide(unpacked_dir, template_source_slide)
        else:
            new_slide_file, rid = _create_slide_from_layout(unpacked_dir, layout_file)

        # Inject content
        _inject_content_into_slide(unpacked_dir, new_slide_file, source_slide, verbose)

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
    )
