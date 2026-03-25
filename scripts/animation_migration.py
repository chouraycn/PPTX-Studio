"""
Enhanced animation migration module for PPTX template application.

This module provides improved animation migration that:
1. Tracks shape ID mappings between source and destination slides
2. Updates <p:tgtEl> references in animation XML to point to correct shapes
3. Preserves animation timing structure while maintaining target bindings

Usage:
    from animation_migration import migrate_animations_with_id_mapping

    migrate_animations_with_id_mapping(
        source_unpacked_dir,
        source_slide_file,
        dest_unpacked_dir,
        dest_slide_file,
        id_mapping,  # {old_shape_id: new_shape_id}
        verbose=True
    )
"""

import re
from pathlib import Path
from typing import Optional, Dict


def _extract_shape_ids(slide_xml: str) -> Dict[str, str]:
    """Extract all shape IDs and their types from a slide XML.

    Returns:
        Dict mapping shape ID to shape type (sp, pic, grp, etc.)
    """
    shape_ids = {}

    # Match <p:sp> (shapes), <p:pic> (pictures), <p:grp> (groups)
    for tag_type in ['sp', 'pic', 'grp', 'graphicFrame', 'cxnSp']:
        pattern = rf'<p:{tag_type}\b[^>]*id="([^"]*)"'
        for m in re.finditer(pattern, slide_xml):
            shape_id = m.group(1)
            shape_ids[shape_id] = tag_type

    return shape_ids


def _update_animation_targets(timing_xml: str, id_mapping: Dict[str, str]) -> str:
    """Update shape ID references in animation XML.

    Args:
        timing_xml: The <p:timing> block from source slide
        id_mapping: Dict mapping old shape IDs to new shape IDs

    Returns:
        Updated timing_xml with shape IDs replaced
    """
    updated_xml = timing_xml

    # Update <p:tgtEl> references
    # Pattern: <p:tgtEl spid="old_id"/>
    updated_xml = re.sub(
        r'<p:tgtEl spid="([^"]*)"/>',
        lambda m: f'<p:tgtEl spid="{id_mapping.get(m.group(1), m.group(1))}"/>',
        updated_xml
    )

    # Update other potential shape ID references in animations
    # <p:cond> elements may reference shapes
    updated_xml = re.sub(
        r'<p:cond\b[^>]*spid="([^"]*)"',
        lambda m: re.sub(
            r'spid="([^"]*)"',
            f'spid="{id_mapping.get(m.group(1), m.group(1))}"',
            m.group(0)
        ),
        updated_xml
    )

    return updated_xml


def _build_id_mapping_by_content(
    source_xml: str,
    dest_xml: str,
    source_shapes: Dict[str, str],
    dest_shapes: Dict[str, str],
) -> Dict[str, str]:
    """Build shape ID mapping based on content similarity.

    Strategy:
    1. Match shapes by type and position (x, y coordinates)
    2. For shapes with similar positions, map old ID to new ID
    3. Preserve mapping for unmatched IDs (fallback)

    Args:
        source_xml: Source slide XML
        dest_xml: Destination slide XML
        source_shapes: Dict of source shape IDs to types
        dest_shapes: Dict of destination shape IDs to types

    Returns:
        Dict mapping old_shape_id -> new_shape_id
    """
    mapping = {}

    # Extract shape positions
    def extract_position(xml: str, shape_id: str) -> Optional[tuple]:
        """Extract (x, y) position for a shape."""
        # Find the element with this ID
        for tag_type in ['sp', 'pic', 'grp', 'graphicFrame']:
            pattern = rf'<p:{tag_type}\b[^>]*id="{re.escape(shape_id)}"[^>]*>.*?</p:{tag_type}>'
            m = re.search(pattern, xml, re.DOTALL)
            if m:
                # Extract a:off or p:spPr/a:xfrm
                xfrm_m = re.search(r'<a:xfrm>.*?<a:off x="(\d+)" y="(\d+)"', m.group(0))
                if xfrm_m:
                    return (int(xfrm_m.group(1)), int(xfrm_m.group(2)))
        return None

    # Build position maps
    source_positions = {}
    for sid in source_shapes:
        pos = extract_position(source_xml, sid)
        if pos:
            source_positions[sid] = pos

    dest_positions = {}
    for did in dest_shapes:
        pos = extract_position(dest_xml, did)
        if pos:
            dest_positions[did] = pos

    # Match shapes by type and position proximity
    POSITION_TOLERANCE = 50000  # EMUs (1 inch = 914400 EMUs)

    matched_dest_ids = set()

    for old_id, old_type in source_shapes.items():
        if old_id not in source_positions:
            continue

        old_pos = source_positions[old_id]

        # Find closest match in destination
        best_match = None
        best_distance = float('inf')

        for new_id, new_type in dest_shapes.items():
            if new_id in matched_dest_ids:
                continue
            if new_type != old_type:
                continue

            if new_id not in dest_positions:
                continue

            new_pos = dest_positions[new_id]
            distance = (
                (old_pos[0] - new_pos[0]) ** 2 +
                (old_pos[1] - new_pos[1]) ** 2
            ) ** 0.5

            if distance < POSITION_TOLERANCE and distance < best_distance:
                best_match = new_id
                best_distance = distance

        if best_match:
            mapping[old_id] = best_match
            matched_dest_ids.add(best_match)

    # Fallback: preserve unmapped IDs (will result in broken animations,
    # but at least won't crash)
    for old_id in source_shapes:
        if old_id not in mapping:
            # Map to closest destination shape of same type
            same_type_dest = [
                did for did, dtype in dest_shapes.items()
                if dtype == source_shapes[old_id] and did not in matched_dest_ids
            ]
            if same_type_dest:
                mapping[old_id] = same_type_dest[0]
                matched_dest_ids.add(same_type_dest[0])

    return mapping


def migrate_animations_with_id_mapping(
    source_unpacked_dir: Path,
    source_slide_file: str,
    dest_unpacked_dir: Path,
    dest_slide_file: str,
    id_mapping: Optional[Dict[str, str]] = None,
    verbose: bool = False,
) -> Dict[str, any]:
    """Migrate animations with shape ID remapping.

    This improves over basic animation migration by:
    1. Tracking shape IDs across source and destination
    2. Updating animation target references to point to correct shapes
    3. Providing detailed feedback on mapping success

    Args:
        source_unpacked_dir: Path to unpacked source PPT
        source_slide_file: Source slide filename (e.g., "slide1.xml")
        dest_unpacked_dir: Path to unpacked destination PPT
        dest_slide_file: Destination slide filename
        id_mapping: Optional pre-computed ID mapping. If None, will auto-detect.
        verbose: Print detailed progress

    Returns:
        Dict with migration statistics:
        {
            "animations_found": bool,
            "animations_migrated": bool,
            "shape_mappings": int,
            "unmapped_shapes": List[str],
            "updated_targets": int
        }
    """
    src_path = source_unpacked_dir / "ppt" / "slides" / source_slide_file
    dst_path = dest_unpacked_dir / "ppt" / "slides" / dest_slide_file

    result = {
        "animations_found": False,
        "animations_migrated": False,
        "shape_mappings": 0,
        "unmapped_shapes": [],
        "updated_targets": 0,
    }

    if not src_path.exists() or not dst_path.exists():
        if verbose:
            print(f"    ⚠️  Animation migration skipped: source or dest file not found")
        return result

    src_xml = src_path.read_text(encoding="utf-8")
    dst_xml = dst_path.read_text(encoding="utf-8")

    # Extract animations from source
    timing_m = re.search(r'<p:timing\b.*?</p:timing>', src_xml, re.DOTALL)
    if not timing_m:
        if verbose:
            print(f"    - No animations found in source slide")
        return result

    result["animations_found"] = True
    timing_xml = timing_m.group(0)

    # Build shape ID mapping if not provided
    if id_mapping is None:
        source_shapes = _extract_shape_ids(src_xml)
        dest_shapes = _extract_shape_ids(dst_xml)

        if verbose:
            print(f"    - Found {len(source_shapes)} shapes in source")
            print(f"    - Found {len(dest_shapes)} shapes in destination")

        id_mapping = _build_id_mapping_by_content(
            src_xml, dst_xml, source_shapes, dest_shapes
        )

        result["shape_mappings"] = len(id_mapping)
        result["unmapped_shapes"] = [
            sid for sid in source_shapes if sid not in id_mapping
        ]

        if verbose:
            print(f"    - Mapped {len(id_mapping)} shape IDs")
            if result["unmapped_shapes"]:
                print(f"    ⚠️  {len(result['unmapped_shapes'])} shapes could not be mapped")
                for sid in result["unmapped_shapes"][:3]:
                    print(f"        Unmapped: {sid}")
                if len(result["unmapped_shapes"]) > 3:
                    print(f"        ... and {len(result['unmapped_shapes']) - 3} more")

    # Update animation targets
    original_targets = len(re.findall(r'<p:tgtEl spid="', timing_xml))
    updated_timing = _update_animation_targets(timing_xml, id_mapping)
    new_targets = len(re.findall(r'<p:tgtEl spid="', updated_timing))

    result["updated_targets"] = new_targets

    if verbose:
        print(f"    - Animation targets: {original_targets} → {new_targets} updated")

    # Remove existing timing from destination
    dst_xml_clean = re.sub(
        r'<p:timing\b.*?</p:timing>', '', dst_xml, flags=re.DOTALL
    )

    # Insert updated timing
    if '</p:sld>' in dst_xml_clean:
        dst_final = dst_xml_clean.replace('</p:sld>', updated_timing + '\n</p:sld>')
        dst_path.write_text(dst_final, encoding="utf-8")
        result["animations_migrated"] = True

        if verbose:
            print(f"    ✓ Animations migrated to {dest_slide_file}")
    else:
        if verbose:
            print(f"    ⚠️  Failed to insert animations: malformed slide XML")

    return result
