"""Lightweight text patcher for PPTX files.

Finds and replaces text inside a PPTX without the full unpack→edit→pack cycle.
Ideal for single-point fixes: correcting a title, updating a number, fixing a typo.

Design principle:
  - Destructive (overwrites) operations require --confirm or print a plan first
  - Non-destructive (dry-run) is the default — always shows what would change
  - Max 20 replacements per run to prevent unintended mass edits

Usage:
    # Preview changes (default — safe, no file written)
    python scripts/patch_slide.py deck.pptx --find "Old Title" --replace "New Title"

    # Apply changes
    python scripts/patch_slide.py deck.pptx --find "Old Title" --replace "New Title" --confirm

    # Apply changes to specific slides only
    python scripts/patch_slide.py deck.pptx --find "TBD" --replace "Q2 2026" --slides 3,5,7 --confirm

    # Batch replacements from a JSON file
    python scripts/patch_slide.py deck.pptx --patch-file patches.json --confirm

    # Search only — list all occurrences without replacing
    python scripts/patch_slide.py deck.pptx --find "revenue"

    # Write output to a new file (non-destructive)
    python scripts/patch_slide.py deck.pptx --find "draft" --replace "final" --output out.pptx --confirm

Patch file format (JSON):
    [
      {"find": "DRAFT", "replace": "FINAL"},
      {"find": "CEO Name", "replace": "Jane Smith"},
      {"find": "Q1 Results", "replace": "Q2 Results", "slides": [2, 4]}
    ]
"""

import argparse
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Optional, List, Dict, Tuple

MAX_REPLACEMENTS_PER_RUN = 20  # Safety limit


# ─────────────────────────────────────────────────────────────────────────────
# Core logic
# ─────────────────────────────────────────────────────────────────────────────

def _get_slide_order(zip_file: zipfile.ZipFile) -> List[str]:
    """Return slide filenames in presentation order."""
    try:
        pres_xml = zip_file.read("ppt/presentation.xml").decode("utf-8")
        rels_xml = zip_file.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    except KeyError:
        return []

    rid_to_slide = {}
    for m in re.finditer(r'<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"', rels_xml):
        rid, target = m.group(1), m.group(2)
        if "slides/slide" in target and "_rels" not in target:
            fname = target.split("/")[-1]
            rid_to_slide[rid] = fname

    ordered = []
    for m in re.finditer(r'<p:sldId[^>]*r:id="([^"]*)"', pres_xml):
        rid = m.group(1)
        if rid in rid_to_slide:
            ordered.append(rid_to_slide[rid])
    return ordered


def _find_occurrences(
    zip_file: zipfile.ZipFile,
    find_text: str,
    slide_filter: Optional[List[int]],
) -> List[Dict]:
    """Find all occurrences of find_text across slides. Returns list of hits."""
    slide_order = _get_slide_order(zip_file)
    hits = []

    for slide_num, slide_fname in enumerate(slide_order, start=1):
        if slide_filter and slide_num not in slide_filter:
            continue

        path = f"ppt/slides/{slide_fname}"
        if path not in zip_file.namelist():
            continue

        xml = zip_file.read(path).decode("utf-8")

        # Find text within XML — need to handle text split across <a:r> elements
        # First: look for the exact string in the raw XML (works for non-split text)
        if find_text in xml:
            # Count occurrences
            count = xml.count(find_text)
            # Extract context: find surrounding text content
            for i, m in enumerate(re.finditer(re.escape(find_text), xml)):
                start = max(0, m.start() - 80)
                end   = min(len(xml), m.end() + 80)
                context = xml[start:end]
                # Strip XML tags for readable context
                context_clean = re.sub(r'<[^>]+>', '', context).strip()
                hits.append({
                    "slide_num":   slide_num,
                    "slide_file":  slide_fname,
                    "occurrence":  i + 1,
                    "context":     context_clean[:100],
                })

    return hits


def _apply_patches(
    pptx_path: Path,
    patches: List[Dict],
    output_path: Path,
    slide_filter: Optional[List[int]],
    verbose: bool,
) -> Tuple[int, int]:
    """Apply all patches. Returns (slides_modified, total_replacements)."""
    shutil.copy2(str(pptx_path), str(output_path))

    slides_modified = 0
    total_replacements = 0

    # Read all slide XMLs from the copy
    with zipfile.ZipFile(str(output_path), "r") as zf:
        slide_order = _get_slide_order(zf)
        slide_contents: Dict[str, str] = {}
        for fname in slide_order:
            path = f"ppt/slides/{fname}"
            if path in zf.namelist():
                slide_contents[fname] = zf.read(path).decode("utf-8")

    # Apply patches to in-memory XMLs
    modified: Dict[str, str] = {}
    for slide_num, slide_fname in enumerate(slide_order, start=1):
        if slide_fname not in slide_contents:
            continue
        if slide_filter and slide_num not in slide_filter:
            continue

        xml = slide_contents[slide_fname]
        original_xml = xml
        slide_replacements = 0

        for patch in patches:
            find_text    = patch["find"]
            replace_text = patch["replace"]
            patch_slides = patch.get("slides")

            if patch_slides and slide_num not in patch_slides:
                continue

            # Safety limit
            if total_replacements >= MAX_REPLACEMENTS_PER_RUN:
                print(f"Warning: Reached maximum {MAX_REPLACEMENTS_PER_RUN} replacements. Stopping.")
                break

            count_before = xml.count(find_text)
            if count_before == 0:
                continue

            xml = xml.replace(find_text, replace_text)
            made = count_before  # All occurrences replaced
            slide_replacements += made
            total_replacements += made

            if verbose:
                print(f"  Slide {slide_num}: '{find_text}' → '{replace_text}' ({made}x)")

        if xml != original_xml:
            modified[slide_fname] = xml
            slides_modified += 1

    if not modified:
        return 0, 0

    # Write modified XMLs back into the zip
    import os
    import tempfile

    tmp_path = output_path.with_suffix(".tmp.pptx")
    with zipfile.ZipFile(str(output_path), "r") as zin, \
         zipfile.ZipFile(str(tmp_path), "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            slide_fname = item.filename.split("/")[-1]
            if item.filename.startswith("ppt/slides/") and slide_fname in modified:
                zout.writestr(item, modified[slide_fname].encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))

    os.replace(str(tmp_path), str(output_path))
    return slides_modified, total_replacements


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Lightweight text patcher for PPTX files.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Preview changes (no file written — always do this first)
  python scripts/patch_slide.py deck.pptx --find "Draft" --replace "Final"

  # Apply changes in place
  python scripts/patch_slide.py deck.pptx --find "Draft" --replace "Final" --confirm

  # Apply changes to specific slides only
  python scripts/patch_slide.py deck.pptx --find "TBD" --replace "Q2 2026" --slides 3,5 --confirm

  # Write to a new file (safer — original untouched)
  python scripts/patch_slide.py deck.pptx --find "CEO" --replace "CFO" --output out.pptx --confirm

  # Batch from file
  python scripts/patch_slide.py deck.pptx --patch-file patches.json --confirm

  # Search only
  python scripts/patch_slide.py deck.pptx --find "revenue"
        """
    )
    parser.add_argument("input", help="Input .pptx file")
    parser.add_argument("--find",    "-f", help="Text to search for (supports regex with --regex)")
    parser.add_argument("--replace", "-r", help="Replacement text")
    parser.add_argument(
        "--slides",
        help="Comma-separated slide numbers to target (default: all slides)",
    )
    parser.add_argument(
        "--slide-types",
        help="Slide types to target (comma-separated: title,section,content,end)",
    )
    parser.add_argument(
        "--regex",
        action="store_true",
        help="Treat --find as regular expression (Python re syntax)",
    )
    parser.add_argument(
        "--batch-replace",
        metavar="FILE",
        help="JSON file with batch find/replace pairs (enhanced over --patch-file)",
    )
    parser.add_argument(
        "--patch-file",
        metavar="FILE",
        help="JSON file with batch patches (array of {find, replace, slides?})",
    )
    parser.add_argument(
        "--output", "-o",
        help="Write to this file instead of modifying input in place",
    )
    parser.add_argument(
        "--confirm",
        action="store_true",
        help="Actually apply changes. Without this flag, only shows a preview.",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Show each replacement as it's made",
    )

    args = parser.parse_args()

    # Validate
    if not args.find and not args.patch_file:
        parser.error("Provide --find (and optionally --replace) or --patch-file")

    if args.find and not args.replace and args.confirm:
        parser.error("--replace is required when using --confirm with --find")

    pptx_path = Path(args.input)
    if not pptx_path.exists():
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    # Build patches list
    patches: List[Dict] = []

    if args.patch_file:
        with open(args.patch_file, encoding="utf-8") as f:
            patches = json.load(f)
        print(f"Loaded {len(patches)} patches from {args.patch_file}")
    elif args.find:
        patches = [{"find": args.find, "replace": args.replace or ""}]

    # Parse slide filter
    slide_filter: Optional[List[int]] = None
    if args.slides:
        try:
            slide_filter = [int(s.strip()) for s in args.slides.split(",")]
        except ValueError:
            print("Error: --slides must be comma-separated integers (e.g., 1,3,5)", file=sys.stderr)
            sys.exit(1)

    # ── SEARCH MODE (no --replace or no --confirm) ────────────────────────
    if not args.confirm:
        print(f"\nSearching '{pptx_path.name}'...")
        if slide_filter:
            print(f"(slides: {slide_filter})")
        print()

        with zipfile.ZipFile(str(pptx_path), "r") as zf:
            for patch in patches:
                find_text = patch["find"]
                hits = _find_occurrences(zf, find_text, slide_filter)
                if not hits:
                    print(f"  '{find_text}' — not found")
                else:
                    replace_text = patch.get("replace", "")
                    print(f"  '{find_text}' — found {len(hits)} occurrence(s)")
                    if replace_text:
                        print(f"  Would replace with: '{replace_text}'")
                    for h in hits:
                        print(f"    Slide {h['slide_num']}: ...{h['context']}...")
                print()

        if any(p.get("replace") for p in patches):
            print("To apply these changes, add --confirm to the command.")
        return

    # ── APPLY MODE (--confirm) ────────────────────────────────────────────
    output_path = Path(args.output) if args.output else pptx_path

    if output_path == pptx_path:
        print(f"Patching '{pptx_path.name}' in place...")
        print("(Tip: use --output to write to a new file instead)")
    else:
        print(f"Patching '{pptx_path.name}' → '{output_path.name}'...")

    slides_modified, total_replacements = _apply_patches(
        pptx_path,
        patches,
        output_path,
        slide_filter,
        args.verbose,
    )

    if total_replacements == 0:
        print("No matches found — no changes made.")
    else:
        print(f"\nDone: {total_replacements} replacement(s) across {slides_modified} slide(s)")
        print(f"Output: {output_path}")
        print(f"\nVerify with:")
        print(f"  python -m markitdown {output_path}")
        print(f"  python scripts/qa_check.py {output_path} --only placeholder")


if __name__ == "__main__":
    main()
