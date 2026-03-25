"""Extract and migrate hyperlinks from source PPT to target PPT.

This module provides functions to:
1. Extract all hyperlinks from a source PPT's slides
2. Inject those hyperlinks into target slides based on text matching

Usage in apply_template.py:
    from hyperlink_migration import extract_hyperlinks, inject_hyperlinks
    hyperlinks = extract_hyperlinks(source_unpacked_dir)
    inject_hyperlinks(unpacked_dir, hyperlinks)
"""

import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import zipfile


def _extract_slide_hyperlinks(slide_xml: str) -> List[Tuple[str, str]]:
    """Extract hyperlinks from a single slide's XML.

    Returns list of (text, url) tuples.
    """
    hyperlinks = []

    # Find all a:hlinkClick elements with r:id
    # Pattern: <a:hlinkClick r:id="rIdN"/>
    hlink_pattern = r'<a:hlinkClick\s+r:id="([^"]+)"\s*/>'

    # Also find hyperlinks in text runs with relationships
    # Pattern: <a:t>text</a:t> with <a:hlinkClick r:id="..."/>
    text_hlink_pattern = r'<a:t>([^<]+)</a:t>.*?<a:hlinkClick\s+r:id="([^"]+)"'

    # Extract hyperlinks with their r:ids
    for match in re.finditer(hlink_pattern, slide_xml, re.DOTALL):
        rid = match.group(1)
        # Find associated text before this hlinkClick
        before_text = slide_xml[:match.start()]
        # Look for the last <a:t> tag
        text_match = re.search(r'<a:t>([^<]+)</a:t>[^<]*$', before_text)
        if text_match:
            text = text_match.group(1).strip()
            hyperlinks.append((text, rid))

    return hyperlinks


def _get_relationship_urls(rels_xml: str) -> Dict[str, str]:
    """Extract relationship URLs from relationships XML.

    Returns dict mapping r:id to target URL.
    """
    relationships = {}

    # Pattern: <Relationship Id="rIdN" Target="..." Type="http://.../hyperlink"/>
    rel_pattern = r'<Relationship\s+Id="([^"]+)"\s+Target="([^"]+)"\s+Type="http://[^"]*/hyperlink"/>'

    for match in re.finditer(rel_pattern, rels_xml):
        rid = match.group(1)
        url = match.group(2)
        relationships[rid] = url

    return relationships


def extract_hyperlinks(unpacked_dir: str) -> Dict[str, List[Tuple[str, str]]]:
    """Extract all hyperlinks from unpacked source PPT.

    Returns dict mapping slide file name to list of (text, url) tuples.

    Args:
        unpacked_dir: Path to unpacked source PPT directory

    Returns:
        {
            "slide1.xml": [("click here", "https://example.com"), ...],
            "slide2.xml": [...],
            ...
        }
    """
    unpacked_path = Path(unpacked_dir)
    slides_dir = unpacked_path / "ppt" / "slides"
    rels_dir = unpacked_path / "ppt" / "slides" / "_rels"

    all_hyperlinks = {}

    # Get all slide files
    slide_files = sorted(slides_dir.glob("slide*.xml"))

    for slide_file in slide_files:
        slide_name = slide_file.name

        # Read slide XML
        with open(slide_file, 'r', encoding='utf-8') as f:
            slide_xml = f.read()

        # Extract hyperlinks with r:ids
        hyperlinks_with_rid = _extract_slide_hyperlinks(slide_xml)

        if not hyperlinks_with_rid:
            continue

        # Read relationships file
        rels_file = rels_dir / f"{slide_name}.rels"
        if rels_file.exists():
            with open(rels_file, 'r', encoding='utf-8') as f:
                rels_xml = f.read()

            # Map r:ids to URLs
            rid_to_url = _get_relationship_urls(rels_xml)

            # Replace r:ids with actual URLs
            hyperlinks = []
            for text, rid in hyperlinks_with_rid:
                if rid in rid_to_url:
                    hyperlinks.append((text, rid_to_url[rid]))
                else:
                    # Keep the r:id as a fallback (shouldn't happen)
                    hyperlinks.append((text, rid))

            if hyperlinks:
                all_hyperlinks[slide_name] = hyperlinks

    return all_hyperlinks


def inject_hyperlinks(
    unpacked_dir: str,
    hyperlinks: Dict[str, List[Tuple[str, str]]],
    source_to_target_map: Optional[Dict[int, str]] = None,
) -> int:
    """Inject hyperlinks into target slides based on text matching.

    Args:
        unpacked_dir: Path to unpacked target PPT directory
        hyperlinks: Dict from extract_hyperlinks(), mapping source slide files to hyperlinks
        source_to_target_map: Optional dict mapping source slide index (1-based) to target slide file name
                           If None, assumes target slides have same file names as source

    Returns:
        Number of hyperlinks injected
    """
    unpacked_path = Path(unpacked_dir)
    slides_dir = unpacked_path / "ppt" / "slides"
    rels_dir = unpacked_path / "ppt" / "slides" / "_rels"

    injected_count = 0

    # Get next available r:id for each target slide
    def get_next_rid(rels_xml: str) -> str:
        """Find the next available r:id number."""
        # Find all existing r:ids
        rids = re.findall(r'rId(\d+)', rels_xml)
        if rids:
            max_id = max(int(rid) for rid in rids)
            return f"rId{max_id + 1}"
        return "rId1"

    # For each source slide with hyperlinks, inject into corresponding target slide
    for source_slide_name, slide_hyperlinks in hyperlinks.items():
        # Determine target slide file
        if source_to_target_map:
            # Extract source slide index from filename (slide1.xml -> 1)
            source_index = int(re.search(r'slide(\d+)\.xml', source_slide_name).group(1))
            if source_index not in source_to_target_map:
                continue
            target_slide_name = source_to_target_map[source_index]
        else:
            # Assume same filename
            target_slide_name = source_slide_name

        target_slide_path = slides_dir / target_slide_name
        if not target_slide_path.exists():
            continue

        # Read target slide XML
        with open(target_slide_path, 'r', encoding='utf-8') as f:
            slide_xml = f.read()

        # Read or create target relationships file
        target_rels_path = rels_dir / f"{target_slide_name}.rels"
        if target_rels_path.exists():
            with open(target_rels_path, 'r', encoding='utf-8') as f:
                rels_xml = f.read()
        else:
            # Create minimal relationships file
            rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>'''

        # Track new relationships to add
        new_rels = []
        rid_mapping = {}

        # Find and replace text with hyperlinked text
        for text, url in slide_hyperlinks:
            # Escape special regex characters in text
            escaped_text = re.escape(text)

            # Find the text in a run
            # Pattern: <a:t>text</a:t>
            if re.search(f'<a:t>{escaped_text}</a:t>', slide_xml):
                # Check if this text already has a hyperlink
                context_before = slide_xml.find(f'<a:t>{text}</a:t>')
                context_after = slide_xml.find('</a:t>', context_before) + len('</a:t>')

                # Look ahead for existing hlinkClick
                if '<a:hlinkClick' in slide_xml[context_after:context_after+100]:
                    # Already has hyperlink, skip
                    continue

                # Get next r:id
                next_rid = get_next_rid(rels_xml)

                # Add to new relationships
                rel_entry = f'  <Relationship Id="{next_rid}" Target="{url}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"/>'
                new_rels.append(rel_entry)
                rid_mapping[text] = next_rid

                # Inject hlinkClick after the text run
                # Find the complete run: <a:r><a:t>text</a:t></a:r>
                run_pattern = f'(<a:r>.*?<a:t>{escaped_text}</a:t>)</a:r>'
                slide_xml = re.sub(
                    run_pattern,
                    rf'\1<a:hlinkClick r:id="{next_rid}"/></a:r>',
                    slide_xml,
                    flags=re.DOTALL
                )
                injected_count += 1

        # If we have new relationships, add them to the file
        if new_rels:
            # Insert before closing </Relationships> tag
            rels_xml = rels_xml.rstrip()
            if rels_xml.endswith('</Relationships>'):
                rels_xml = rels_xml[:-17]  # Remove </Relationships>

            # Add new relationships (reverse order to maintain increasing r:id)
            for rel in sorted(new_rels, reverse=True):
                rels_xml += '\n' + rel

            rels_xml += '\n</Relationships>'

            # Write back relationships file
            target_rels_path.parent.mkdir(parents=True, exist_ok=True)
            with open(target_rels_path, 'w', encoding='utf-8') as f:
                f.write(rels_xml)

        # Write back slide XML
        with open(target_slide_path, 'w', encoding='utf-8') as f:
            f.write(slide_xml)

    return injected_count


# CLI interface for testing
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Extract or inject hyperlinks")
    parser.add_argument("action", choices=["extract", "inject"], help="Action to perform")
    parser.add_argument("path", help="Path to unpacked PPT directory")
    parser.add_argument("--output", help="Output JSON file for extract action")
    parser.add_argument("--input", help="Input JSON file for inject action")

    args = parser.parse_args()

    if args.action == "extract":
        hyperlinks = extract_hyperlinks(args.path)
        if args.output:
            import json
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(hyperlinks, f, indent=2)
            print(f"Extracted {sum(len(v) for v in hyperlinks.values())} hyperlinks to {args.output}")
        else:
            print(f"Extracted {sum(len(v) for v in hyperlinks.values())} hyperlinks:")
            for slide, links in hyperlinks.items():
                print(f"  {slide}:")
                for text, url in links:
                    print(f"    {text} -> {url}")

    elif args.action == "inject":
        if not args.input:
            print("Error: --input is required for inject action", file=sys.stderr)
            sys.exit(1)

        import json
        with open(args.input, 'r', encoding='utf-8') as f:
            hyperlinks = json.load(f)

        injected = inject_hyperlinks(args.path, hyperlinks)
        print(f"Injected {injected} hyperlinks")
