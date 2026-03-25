"""
Enhanced content extraction with table and SmartArt preservation.

This module provides functionality to:
1. Detect tables and SmartArt in source PPT
2. Extract complete XML structure for preservation
3. Re-inject tables/SmartArt into destination slides

Usage:
    from table_smartart import extract_tables_smartart, preserve_tables_smartart

    # Extract from source
    tables_smartart = extract_tables_smartart(source_unpacked_dir, slide_file)

    # Preserve during template apply
    preserve_tables_smartart(
        dest_unpacked_dir,
        dest_slide_file,
        tables_smartart,
        template_colors,
    )
"""

import re
from pathlib import Path
from typing import Optional, List, Dict, Tuple


def _extract_tables(slide_xml: str) -> List[Dict]:
    """Extract all tables from a slide XML.

    Returns:
        List of dict with table structure:
        {
            "id": table shape ID,
            "xml": complete table XML including <a:tbl> element,
            "pos": {"x": ..., "y": ...} (position coordinates),
            "rows": row count,
            "cols": col count,
        }
    """
    tables = []

    # Find all <a:tbl> elements
    for m in re.finditer(r'<a:tbl\b.*?</a:tbl>', slide_xml, re.DOTALL):
        table_xml = m.group(0)

        # Extract table ID (from parent <p:graphicFrame> or <p:sp>)
        table_id_match = re.search(
            r'(?:(?:<p:(?:graphicFrame|sp)\b[^>]*id="([^"]*)"[^>]*>).*?)(?=<a:tbl>)',
            slide_xml[:m.start() + 500]  # Look in context before table
        )
        table_id = table_id_match.group(1) if table_id_match else None

        # Extract position from <a:off>
        pos_match = re.search(r'<a:off x="(\d+)" y="(\d+)"', table_xml)
        pos = {
            "x": int(pos_match.group(1)) if pos_match else None,
            "y": int(pos_match.group(2)) if pos_match else None,
        }

        # Count rows and cols
        tr_count = len(re.findall(r'<a:tr\b', table_xml))
        max_cols = 0
        for tr_m in re.finditer(r'<a:tr\b.*?</a:tr>', table_xml):
            tc_count = len(re.findall(r'<a:tc\b', tr_m.group(0)))
            max_cols = max(max_cols, tc_count)

        tables.append({
            "id": table_id,
            "xml": table_xml,
            "pos": pos,
            "rows": tr_count,
            "cols": max_cols,
            "type": "table",
        })

    return tables


def _extract_smartart(slide_xml: str) -> List[Dict]:
    """Extract all SmartArt diagrams from a slide XML.

    SmartArt in PPTX is stored as <p:graphic> with dataModel references.
    We extract the complete structure including:
    - The diagram type (dataModel)
    - Layout information
    - Shape positions and connections

    Returns:
        List of dict with SmartArt structure:
        {
            "id": shape ID,
            "xml": complete graphic XML,
            "pos": {"x": ..., "y": ...},
            "dataModel": diagram type identifier,
        }
    """
    smartart = []

    # Find all <p:graphic> elements (SmartArt container)
    for m in re.finditer(r'<p:graphic\b.*?</p:graphic>', slide_xml, re.DOTALL):
        graphic_xml = m.group(0)

        # Extract graphic ID from parent <p:graphicFrame>
        id_context = slide_xml[:m.start() + 200]
        id_match = re.search(r'<p:(?:graphicFrame|sp)\b[^>]*id="([^"]*)"', id_context)
        smartart_id = id_match.group(1) if id_match else None

        # Extract dataModel (diagram type)
        dm_match = re.search(r'dataModel="([^"]*)"', graphic_xml)
        data_model = dm_match.group(1) if dm_match else None

        # Extract position
        pos_match = re.search(r'<a:off x="(\d+)" y="(\d+)"', graphic_xml)
        pos = {
            "x": int(pos_match.group(1)) if pos_match else None,
            "y": int(pos_match.group(2)) if pos_match else None,
        }

        smartart.append({
            "id": smartart_id,
            "xml": graphic_xml,
            "pos": pos,
            "dataModel": data_model,
            "type": "smartart",
        })

    return smartart


def extract_tables_smartart(
    unpacked_dir: Path,
    slide_file: str,
) -> List[Dict]:
    """Extract all tables and SmartArt from a slide.

    Args:
        unpacked_dir: Path to unpacked PPT
        slide_file: Slide filename (e.g., "slide1.xml")

    Returns:
        List of dict with tables and SmartArt structures
    """
    slide_path = unpacked_dir / "ppt" / "slides" / slide_file
    if not slide_path.exists():
        return []

    slide_xml = slide_path.read_text(encoding="utf-8")

    elements = []

    # Extract tables
    tables = _extract_tables(slide_xml)
    elements.extend(tables)

    # Extract SmartArt
    smartart = _extract_smartart(slide_xml)
    elements.extend(smartart)

    return elements


def _apply_template_colors_to_table(
    table_xml: str,
    template_colors: Dict[str, str],
) -> str:
    """Apply template colors to table XML.

    Updates fill colors in table cells to use template colors.
    Preserves table structure while applying new color scheme.

    Args:
        table_xml: Original table XML
        template_colors: Dict with "primary", "secondary", "accent" keys

    Returns:
        Updated table XML with template colors
    """
    updated_xml = table_xml

    # Apply primary color to header fills
    updated_xml = re.sub(
        r'<a:solidFill>.*?<a:srgbClr val="[^"]*"/></a:solidFill>',
        lambda m: m.group(0).replace(
            m.group(0).split('val="')[1].split('"')[0],
            template_colors.get("primary", "1E2761")
        ) if "header" in updated_xml[:m.start() + 100] else m.group(0),
        updated_xml,
        count=1,  # Only replace first occurrence (header)
    )

    # Apply accent color to borders
    updated_xml = re.sub(
        r'<a:solidFill>.*?<a:srgbClr val="[^"]*"/></a:solidFill>',
        lambda m: m.group(0).replace(
            m.group(0).split('val="')[1].split('"')[0],
            template_colors.get("accent", "C9A84C")
        ),
        updated_xml,
    )

    # Apply secondary color to alternating rows
    # This is simplified; full implementation would track row indices
    return updated_xml


def preserve_tables_smartart(
    dest_unpacked_dir: Path,
    dest_slide_file: str,
    tables_smartart: List[Dict],
    template_colors: Optional[Dict[str, str]] = None,
    verbose: bool = False,
) -> int:
    """Preserve tables and SmartArt from source in destination slide.

    This function:
    1. Identifies tables/SmartArt positions in destination
    2. Re-injects source tables/SmartArt at appropriate locations
    3. Applies template colors to tables (if template_colors provided)

    Args:
        dest_unpacked_dir: Path to unpacked destination PPT
        dest_slide_file: Destination slide filename
        tables_smartart: List of tables/SmartArt from extract_tables_smartart()
        template_colors: Optional template colors to apply to tables
        verbose: Print detailed progress

    Returns:
        Number of elements preserved
    """
    if not tables_smartart:
        return 0

    dest_path = dest_unpacked_dir / "ppt" / "slides" / dest_slide_file
    if not dest_path.exists():
        return 0

    dest_xml = dest_path.read_text(encoding="utf-8")
    original_xml = dest_xml

    preserved_count = 0

    for element in tables_smartart:
        elem_type = element.get("type", "unknown")
        elem_id = element.get("id")
        elem_xml = element["xml"]
        pos = element.get("pos", {})

        if elem_type == "table":
            # Apply template colors to table
            if template_colors:
                elem_xml = _apply_template_colors_to_table(elem_xml, template_colors)

            # Find appropriate insertion point in destination
            # For simplicity, we replace the first <p:graphicFrame> or add after title
            if "<p:graphicFrame" in dest_xml:
                # Replace first table/graphic
                dest_xml = re.sub(
                    r'<p:graphicFrame\b.*?</p:graphicFrame>',
                    elem_xml,
                    dest_xml,
                    count=1,
                )
                preserved_count += 1
                if verbose:
                    print(f"    ✓ Preserved table with {element['rows']}×{element['cols']} cells")

            else:
                # Insert after first paragraph or before closing </p:spTree>
                insert_pos = dest_xml.find("</a:p>")
                if insert_pos > 0:
                    dest_xml = dest_xml[:insert_pos] + "</a:p>" + elem_xml + dest_xml[insert_pos:]
                    preserved_count += 1
                    if verbose:
                        print(f"    ✓ Inserted table with {element['rows']}×{element['cols']} cells")

        elif elem_type == "smartart":
            # SmartArt needs to be injected carefully to preserve dataModel references
            if "<p:graphic>" in dest_xml:
                # Replace existing SmartArt (if any)
                dest_xml = re.sub(
                    r'<p:graphicFrame\b[^>]*>.*?<p:graphic\b.*?</p:graphic>.*?</p:graphicFrame>',
                    elem_xml,
                    dest_xml,
                    count=1,
                    flags=re.DOTALL,
                )
                preserved_count += 1
                if verbose:
                    dm = element.get("dataModel", "unknown")
                    print(f"    ✓ Preserved SmartArt diagram (type: {dm})")

            else:
                # Insert after title section
                insert_pos = dest_xml.find("</p:spTree>") - 11  # Before closing tag
                if insert_pos > 0:
                    dest_xml = dest_xml[:insert_pos] + elem_xml + dest_xml[insert_pos:]
                    preserved_count += 1
                    if verbose:
                        dm = element.get("dataModel", "unknown")
                        print(f"    ✓ Inserted SmartArt diagram (type: {dm})")

    # Write updated XML
    if dest_xml != original_xml:
        dest_path.write_text(dest_xml, encoding="utf-8")

    return preserved_count
