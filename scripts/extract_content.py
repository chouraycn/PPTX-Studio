"""Extract structured content from a PPTX file.

Outputs a JSON representation of all slide content, preserving structure for
later use in template application or style beautification.

Usage:
    python scripts/extract_content.py source.pptx
    python scripts/extract_content.py source.pptx --output content.json
    python scripts/extract_content.py source.pptx --print-summary

Output JSON format:
    {
      "slides": [
        {
          "index": 1,
          "slide_file": "slide1.xml",
          "type": "title",
          "title": "...",
          "subtitle": "...",
          "body": ["bullet 1", "bullet 2"],
          "body_rich": [{"text": "...", "bold": true, "size": 28}],
          "notes": "...",
          "has_images": false,
          "image_count": 0,
          "has_charts": false,
          "has_tables": false,
          "table_data": [],
          "layout_name": "Title Slide",
          "layout_file": "slideLayout1.xml",
          "layout_hint": "title_slide",
          "shape_count": 5,
          "background_color": "FFFFFF"
        }
      ],
      "total_slides": 12,
      "topic_keywords": ["AI", "strategy"],
      "detected_theme": "executive"
    }
"""

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple


def extract_content(pptx_path: str, output_path: Optional[str] = None, print_summary: bool = False) -> dict:
    """Extract all content from a PPTX file into a structured dict."""
    path = Path(pptx_path)
    if not path.exists():
        print(f"Error: {pptx_path} not found", file=sys.stderr)
        sys.exit(1)

    result = {"slides": [], "total_slides": 0, "topic_keywords": [], "detected_theme": "minimal"}

    with zipfile.ZipFile(path, "r") as zf:
        # Get slide order from presentation.xml
        slide_order = _get_slide_order(zf)

        # Get layout names
        layout_names = _get_layout_names(zf)

        # Extract each slide
        for idx, (slide_file, layout_file) in enumerate(slide_order, 1):
            try:
                slide_xml = zf.read(f"ppt/slides/{slide_file}").decode("utf-8")
                rels_path = f"ppt/slides/_rels/{slide_file}.rels"
                rels_xml = ""
                if rels_path in zf.namelist():
                    rels_xml = zf.read(rels_path).decode("utf-8")

                slide_data = _parse_slide(slide_xml, rels_xml, idx, slide_file, layout_file, layout_names, zf)
                result["slides"].append(slide_data)
            except Exception as e:
                result["slides"].append({
                    "index": idx,
                    "slide_file": slide_file,
                    "error": str(e),
                    "type": "unknown"
                })

    result["total_slides"] = len(result["slides"])
    result["topic_keywords"] = _extract_keywords(result["slides"])
    result["detected_theme"] = _detect_theme(result["slides"])

    if output_path:
        Path(output_path).write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"Content extracted to {output_path}")

    if print_summary:
        _print_summary(result)

    return result


def _get_slide_order(zf: zipfile.ZipFile) -> List[Tuple[str, str]]:
    """Return list of (slide_file, layout_file) in presentation order."""
    pres_xml = zf.read("ppt/presentation.xml").decode("utf-8")
    rels_xml = zf.read("ppt/_rels/presentation.xml.rels").decode("utf-8")

    # Build rId -> slide file mapping
    rid_to_file = {}
    for m in re.finditer(r'<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]*slide"[^>]+Target="([^"]+)"', rels_xml):
        rid, target = m.group(1), m.group(2)
        # target is like "slides/slide1.xml"
        rid_to_file[rid] = target.replace("slides/", "")

    # Get slide order
    order = []
    for m in re.finditer(r'<p:sldId[^>]+r:id="([^"]+)"', pres_xml):
        rid = m.group(1)
        if rid in rid_to_file:
            slide_file = rid_to_file[rid]
            # Find the layout for this slide
            layout_file = _get_slide_layout(zf, slide_file)
            order.append((slide_file, layout_file))

    return order


def _get_slide_layout(zf: zipfile.ZipFile, slide_file: str) -> str:
    """Get the layout file name for a slide."""
    rels_path = f"ppt/slides/_rels/{slide_file}.rels"
    if rels_path not in zf.namelist():
        return ""
    rels_xml = zf.read(rels_path).decode("utf-8")
    m = re.search(r'Type="[^"]*slideLayout"[^>]+Target="([^"]+)"', rels_xml)
    if m:
        return m.group(1).replace("../slideLayouts/", "")
    return ""


def _get_layout_names(zf: zipfile.ZipFile) -> Dict[str, str]:
    """Return mapping of layout file -> layout name."""
    names = {}
    for name in zf.namelist():
        if name.startswith("ppt/slideLayouts/") and name.endswith(".xml") and "_rels" not in name:
            layout_file = Path(name).name
            try:
                xml = zf.read(name).decode("utf-8")
                m = re.search(r'<p:cSld[^>]*name="([^"]*)"', xml)
                if m:
                    names[layout_file] = m.group(1)
                else:
                    names[layout_file] = layout_file.replace(".xml", "")
            except Exception:
                names[layout_file] = layout_file.replace(".xml", "")
    return names


def _parse_slide(slide_xml: str, rels_xml: str, idx: int, slide_file: str,
                 layout_file: str, layout_names: Dict[str, str], zf: zipfile.ZipFile) -> dict:
    """Parse a single slide XML into structured content."""
    data = {
        "index": idx,
        "slide_file": slide_file,
        "type": "content",
        "title": "",
        "subtitle": "",
        "body": [],
        "body_rich": [],
        "notes": "",
        "has_images": False,
        "image_count": 0,
        "has_charts": False,
        "has_tables": False,
        "table_data": [],
        "layout_file": layout_file,
        "layout_name": layout_names.get(layout_file, ""),
        "layout_hint": "",
        "shape_count": 0,
        "background_color": "",
    }

    # Extract title (from ph type="title" or ph type="ctrTitle")
    title_patterns = [
        r'<p:ph[^>]*type="(?:title|ctrTitle)"[^>]*/?>.*?</p:ph>.*?<a:t>([^<]+)</a:t>',
        r'<p:ph[^>]*type="(?:title|ctrTitle)"[^/]*/?>.*?<a:t>([^<]+)</a:t>',
    ]

    # More robust: find all text in title placeholders
    title_text = _extract_placeholder_text(slide_xml, ["title", "ctrTitle"])
    if title_text:
        data["title"] = title_text

    # Extract subtitle/body
    subtitle_text = _extract_placeholder_text(slide_xml, ["subTitle"])
    if subtitle_text:
        data["subtitle"] = subtitle_text

    # Extract body text (all text runs from body/object placeholders)
    body_lines = _extract_body_text(slide_xml)
    data["body"] = body_lines

    # Extract rich body text
    data["body_rich"] = _extract_rich_text(slide_xml)

    # Check for images
    image_count = len(re.findall(r'<p:pic\b', slide_xml))
    data["has_images"] = image_count > 0
    data["image_count"] = image_count

    # Check for charts
    data["has_charts"] = "<c:chart" in slide_xml or "chart" in rels_xml.lower()

    # Check for tables
    data["has_tables"] = "<a:tbl" in slide_xml
    if data["has_tables"]:
        data["table_data"] = _extract_table_data(slide_xml)

    # Count shapes (non-placeholder shapes)
    data["shape_count"] = len(re.findall(r'<p:sp\b', slide_xml))

    # Extract background color if any
    m = re.search(r'<p:bg>.*?<a:srgbClr val="([0-9A-Fa-f]{6})"', slide_xml, re.DOTALL)
    if m:
        data["background_color"] = m.group(1)

    # Extract notes
    notes_file = slide_file.replace("slide", "notesSlide")
    notes_path = f"ppt/notesSlides/{notes_file}"
    if notes_path in zf.namelist():
        notes_xml = zf.read(notes_path).decode("utf-8")
        notes_texts = re.findall(r'<a:t>([^<]+)</a:t>', notes_xml)
        data["notes"] = " ".join(t.strip() for t in notes_texts if t.strip())

    # Determine slide type
    data["type"] = _classify_slide_type(data)
    data["layout_hint"] = _get_layout_hint(data)

    return data


def _extract_placeholder_text(slide_xml: str, ph_types: List[str]) -> str:
    """Extract all text from specific placeholder types."""
    texts = []
    type_pattern = "|".join(ph_types)

    # Find sp elements containing the placeholder type
    for sp_match in re.finditer(r'<p:sp\b.*?</p:sp>', slide_xml, re.DOTALL):
        sp_xml = sp_match.group(0)
        if re.search(rf'<p:ph[^>]*type="(?:{type_pattern})"', sp_xml):
            # Extract all text runs
            t_texts = re.findall(r'<a:t>([^<]*)</a:t>', sp_xml)
            if t_texts:
                combined = "".join(t_texts).strip()
                if combined:
                    texts.append(combined)

    return " ".join(texts)


def _extract_body_text(slide_xml: str) -> List[str]:
    """Extract body/content text as a list of lines.

    Only pulls from genuine content placeholders (body, obj) and non-typed
    placeholders (idx-only, which are also treated as body in OOXML).

    Explicitly excludes:
    - title / ctrTitle  → handled by _extract_placeholder_text separately
    - subTitle          → has its own 'subtitle' field in the slide data
    - dt / ftr / sldNum → decoration (date, footer, slide number) — not content
    """
    lines = []
    # Only extract from actual body/content placeholder types
    body_types = ["body", "obj"]
    type_pattern = "|".join(body_types)
    # Types to explicitly skip (not content)
    skip_types = {"title", "ctrTitle", "subTitle", "dt", "ftr", "sldNum"}

    for sp_match in re.finditer(r'<p:sp\b.*?</p:sp>', slide_xml, re.DOTALL):
        sp_xml = sp_match.group(0)
        # Skip if this has a non-content placeholder type
        ph_type_m = re.search(r'<p:ph[^>]*type="([^"]+)"', sp_xml)
        if ph_type_m and ph_type_m.group(1) in skip_types:
            continue
        # Include body/obj/non-typed placeholders (idx-only = implicit body)
        if re.search(rf'<p:ph[^>]*type="(?:{type_pattern})"', sp_xml) or \
           (re.search(r'<p:ph\b', sp_xml) and not re.search(r'<p:ph[^>]*type=', sp_xml)):
            # Extract paragraph by paragraph
            for para in re.finditer(r'<a:p\b.*?</a:p>', sp_xml, re.DOTALL):
                para_xml = para.group(0)
                t_texts = re.findall(r'<a:t>([^<]*)</a:t>', para_xml)
                line = "".join(t_texts).strip()
                if line:
                    lines.append(line)

    # Also extract text from non-placeholder text boxes (free-floating)
    for sp_match in re.finditer(r'<p:sp\b.*?</p:sp>', slide_xml, re.DOTALL):
        sp_xml = sp_match.group(0)
        if not re.search(r'<p:ph\b', sp_xml):
            for para in re.finditer(r'<a:p\b.*?</a:p>', sp_xml, re.DOTALL):
                para_xml = para.group(0)
                t_texts = re.findall(r'<a:t>([^<]*)</a:t>', para_xml)
                line = "".join(t_texts).strip()
                if line and line not in lines:
                    lines.append(line)

    return lines


def _extract_rich_text(slide_xml: str) -> List[dict]:
    """Extract rich text with formatting info."""
    items = []
    for sp_match in re.finditer(r'<p:sp\b.*?</p:sp>', slide_xml, re.DOTALL):
        sp_xml = sp_match.group(0)
        if re.search(r'<p:ph[^>]*type="(?:title|ctrTitle)"', sp_xml):
            continue
        for run_match in re.finditer(r'<a:r\b.*?</a:r>', sp_xml, re.DOTALL):
            run_xml = run_match.group(0)
            t_match = re.search(r'<a:t>([^<]*)</a:t>', run_xml)
            if not t_match or not t_match.group(1).strip():
                continue
            item = {"text": t_match.group(1)}
            rpr = re.search(r'<a:rPr([^>]*)>', run_xml)
            if rpr:
                rpr_attrs = rpr.group(1)
                item["bold"] = 'b="1"' in rpr_attrs
                item["italic"] = 'i="1"' in rpr_attrs
                sz = re.search(r'sz="(\d+)"', rpr_attrs)
                if sz:
                    item["size"] = int(sz.group(1)) // 100  # Convert from hundredths of points
                color = re.search(r'val="([0-9A-Fa-f]{6})"', run_xml)
                if color:
                    item["color"] = color.group(1)
            items.append(item)
    return items


def _extract_table_data(slide_xml: str) -> List[List[str]]:
    """Extract table data as 2D array."""
    tables = []
    for tbl_match in re.finditer(r'<a:tbl\b.*?</a:tbl>', slide_xml, re.DOTALL):
        tbl_xml = tbl_match.group(0)
        rows = []
        for row_match in re.finditer(r'<a:tr\b.*?</a:tr>', tbl_xml, re.DOTALL):
            row_xml = row_match.group(0)
            cells = []
            for tc_match in re.finditer(r'<a:tc\b.*?</a:tc>', row_xml, re.DOTALL):
                tc_xml = tc_match.group(0)
                t_texts = re.findall(r'<a:t>([^<]*)</a:t>', tc_xml)
                cells.append("".join(t_texts).strip())
            rows.append(cells)
        tables.append(rows)
    return tables


def _classify_slide_type(data: dict) -> str:
    """Classify a slide's type based on its content."""
    title = data["title"].lower()
    body = data["body"]
    has_images = data["has_images"]
    has_charts = data["has_charts"]
    has_tables = data["has_tables"]
    layout_name = data["layout_name"].lower()

    # Check layout name first
    if any(kw in layout_name for kw in ["title slide", "title only", "title, content"]):
        if not body and not has_images:
            return "title"

    # Title slide: large title, subtitle, minimal content
    if not body and not has_images and not has_charts and data["title"]:
        if any(kw in title for kw in ["agenda", "目录", "contents", "outline"]):
            return "agenda"
        if any(kw in title for kw in ["thank", "end", "conclusion", "questions", "谢谢", "结束", "总结"]):
            return "conclusion"
        return "title"

    # Section header
    if not body and data["title"] and (len(data["title"]) < 40):
        if "section" in layout_name or "divider" in layout_name:
            return "section"

    # Image-heavy
    if has_images and len(body) <= 3:
        return "image"

    # Chart/data
    if has_charts:
        return "chart"

    # Table
    if has_tables:
        return "table"

    # Quote slide
    if len(body) == 1 and len(body[0]) > 80:
        return "quote"

    # Agenda
    if any(kw in title for kw in ["agenda", "outline", "contents", "目录", "大纲"]):
        return "agenda"

    # Conclusion
    if any(kw in title for kw in ["conclusion", "summary", "next steps", "questions", "总结", "结论", "谢谢"]):
        return "conclusion"

    return "content"


def _get_layout_hint(data: dict) -> str:
    """Suggest the best layout to use in a template."""
    slide_type = data["type"]
    body_count = len(data["body"])
    has_images = data["has_images"]

    if slide_type == "title":
        return "title_slide"
    if slide_type == "section":
        return "section_header"
    if slide_type == "conclusion":
        return "conclusion"
    if slide_type == "agenda":
        return "list_content"
    if slide_type == "chart":
        return "chart_content"
    if slide_type == "table":
        return "table_content"
    if slide_type == "quote":
        return "quote_slide"
    if slide_type == "image":
        if body_count > 0:
            return "image_text"
        return "full_image"
    if has_images and body_count > 0:
        return "image_text"
    if body_count > 6:
        return "list_content"
    if body_count > 0:
        return "content_slide"
    return "content_slide"


def _extract_keywords(slides: List[dict]) -> List[str]:
    """Extract topic keywords from all slide content."""
    all_text = []
    for slide in slides:
        if slide.get("title"):
            all_text.append(slide["title"])
        all_text.extend(slide.get("body", []))

    # Simple keyword extraction: frequent meaningful words
    text = " ".join(all_text).lower()
    words = re.findall(r'\b[a-z\u4e00-\u9fff]{3,}\b', text)

    # Filter common stop words
    stop_words = {"the", "and", "for", "with", "this", "that", "are", "from",
                  "our", "your", "will", "can", "has", "have", "its", "been",
                  "not", "all", "was", "they", "their", "which", "how", "what"}

    freq: Dict[str, int] = {}
    for word in words:
        if word not in stop_words and len(word) > 3:
            freq[word] = freq.get(word, 0) + 1

    sorted_words = sorted(freq.items(), key=lambda x: x[1], reverse=True)
    return [w for w, _ in sorted_words[:10]]


def _detect_theme(slides: List[dict]) -> str:
    """Auto-detect appropriate theme based on content keywords."""
    keywords = _extract_keywords(slides)
    text = " ".join(keywords).lower()

    if any(w in text for w in ["tech", "software", "ai", "data", "digital", "api", "code", "cloud", "platform"]):
        return "tech"
    if any(w in text for w in ["design", "creative", "brand", "market", "product", "user", "experience"]):
        return "creative"
    if any(w in text for w in ["finance", "revenue", "growth", "profit", "investment", "strategy", "business"]):
        return "executive"
    if any(w in text for w in ["education", "learn", "student", "health", "community", "people", "team"]):
        return "warm"
    if any(w in text for w in ["environment", "green", "nature", "sustainable", "ecology", "carbon"]):
        return "nature"
    if any(w in text for w in ["research", "study", "analysis", "report", "academic", "survey"]):
        return "minimal"
    if any(w in text for w in ["sales", "launch", "campaign", "promote", "offer", "deal"]):
        return "bold"
    return "minimal"


def _print_summary(result: dict) -> None:
    """Print a human-readable summary of extracted content."""
    print(f"\n{'='*60}")
    print(f"PPT Content Summary — {result['total_slides']} slides")
    print(f"Detected theme: {result['detected_theme']}")
    print(f"Keywords: {', '.join(result['topic_keywords'][:5])}")
    print(f"{'='*60}\n")

    for slide in result["slides"]:
        if "error" in slide:
            print(f"  Slide {slide['index']}: ERROR — {slide['error']}")
            continue
        print(f"  Slide {slide['index']} [{slide['type']}] — {slide['title'] or '(no title)'}")
        if slide["body"]:
            for i, line in enumerate(slide["body"][:3]):
                print(f"    • {line[:60]}{'...' if len(line) > 60 else ''}")
            if len(slide["body"]) > 3:
                print(f"    ... (+{len(slide['body']) - 3} more lines)")
        flags = []
        if slide["has_images"]:
            flags.append(f"{slide['image_count']} image(s)")
        if slide["has_charts"]:
            flags.append("chart")
        if slide["has_tables"]:
            flags.append("table")
        if flags:
            print(f"    [{', '.join(flags)}]")
        print()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract content from a PPTX file")
    parser.add_argument("input", help="Input PPTX file")
    parser.add_argument("--output", "-o", help="Output JSON file path")
    parser.add_argument("--print-summary", "-s", action="store_true",
                        help="Print human-readable summary to stdout")
    args = parser.parse_args()

    if not args.output and not args.print_summary:
        # Default: print summary if no output specified
        args.print_summary = True

    extract_content(args.input, args.output, args.print_summary)
