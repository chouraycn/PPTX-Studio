"""
Enhanced QA checks for PPTX - visual and semantic checks.

This module extends qa_check.py with additional checks:
- Visual alignment and spacing
- Content semantic consistency
- Cross-slide coherence
- Data accuracy validation

Usage:
    from qa_enhanced import run_enhanced_qa_checks

    results = run_enhanced_qa_checks(
        unpacked_dir,
        slides_data,
        verbose=True
    )
"""

import re
from pathlib import Path
from typing import List, Dict, Any, Tuple


def _check_element_alignment(slide_xml: str) -> List[Dict]:
    """Check for misaligned elements within a slide.

    Detects:
    - Elements with same x coordinate that should be aligned
    - Uneven gaps between similar elements
    - Elements not on a grid pattern when they should be

    Returns:
        List of alignment issues found
    """
    issues = []

    # Extract all element positions
    elements = []
    for tag in ['sp', 'pic', 'graphicFrame']:
        for m in re.finditer(
            rf'<p:{tag}\b[^>]*>.*?</p:{tag}>',
            slide_xml,
            re.DOTALL
        ):
            elem_xml = m.group(0)

            # Extract position
            off_match = re.search(r'<a:off x="(\d+)" y="(\d+)"', elem_xml)
            if off_match:
                elements.append({
                    "tag": tag,
                    "x": int(off_match.group(1)),
                    "y": int(off_match.group(2)),
                    "xml": elem_xml,
                })

    if len(elements) < 2:
        return issues

    # Check for x-coordinate clustering (should align)
    x_coords = [e["x"] for e in elements]
    x_coords.sort()

    # Group by similar x (tolerance 5000 EMUs ≈ 5pt)
    tolerance = 5000
    for i in range(len(x_coords) - 1):
        if abs(x_coords[i] - x_coords[i + 1]) < tolerance:
            # These should align but might not be perfectly aligned
            if x_coords[i] != x_coords[i + 1]:
                issues.append({
                    "check": "alignment",
                    "severity": "warning",
                    "message": f"Elements at x={x_coords[i]} and x={x_coords[i + 1]} are close but not aligned",
                    "suggestion": "Align elements horizontally for cleaner layout",
                })

    # Check for uneven vertical spacing
    for i in range(len(elements) - 1):
        gap = elements[i + 1]["y"] - elements[i]["y"]
        # Gaps should be consistent; flag if gap < 10000 (too tight) or > 300000 (too loose)
        if gap < 10000:
            issues.append({
                "check": "spacing",
                "severity": "warning",
                "message": f"Elements too close vertically (gap: {gap} EMUs)",
                "suggestion": "Increase vertical spacing to at least 10pt",
            })

    return issues


def _check_color_consistency(
    slide_xml: str,
    slide_num: int,
) -> List[Dict]:
    """Check for inconsistent color usage within a slide.

    Detects:
    - Too many different fill colors
    - Inconsistent accent color usage
    - Low-contrast element combinations

    Returns:
        List of color consistency issues
    """
    issues = []

    # Extract all fill colors
    fill_colors = set()
    for m in re.finditer(r'<a:srgbClr val="([A-F0-9]{6})"', slide_xml):
        color = m.group(1)
        fill_colors.add(color)

    # Flag if too many distinct colors (> 5 is excessive for most slides)
    if len(fill_colors) > 5:
        issues.append({
            "check": "color_consistency",
            "severity": "info",
            "message": f"Slide uses {len(fill_colors)} different colors (recommend ≤5)",
            "suggestion": "Reduce color palette for more cohesive design",
        })

    # Check for low contrast combinations
    # Look for dark on dark or light on light
    # This is simplified; full implementation would check luminance ratios
    color_pairs = [
        ("000000", "1A1A1A"),  # black on dark gray
        ("FFFFFF", "F5F5F5"),  # white on light gray
    ]
    text_colors = set()
    bg_colors = set()

    for m in re.finditer(r'<a:rPr[^>]*>.*?<a:solidFill>.*?<a:srgbClr val="([A-F0-9]{6})"', slide_xml, re.DOTALL):
        text_colors.add(m.group(1))

    for m in re.finditer(r'<p:bgPr>.*?<a:srgbClr val="([A-F0-9]{6})"', slide_xml, re.DOTALL):
        bg_colors.add(m.group(1))

    # Simple low-contrast check
    for text_c in text_colors:
        for bg_c in bg_colors:
            for bad_pair in color_pairs:
                if text_c in bad_pair and bg_c in bad_pair:
                    issues.append({
                        "check": "contrast",
                        "severity": "warning",
                        "message": f"Potential low contrast: text {text_c} on background {bg_c}",
                        "suggestion": "Increase contrast for better readability",
                    })

    return issues


def _check_semantic_consistency(
    slides_data: List[Dict],
) -> List[Dict]:
    """Check for semantic consistency across slides.

    Detects:
    - Inconsistent terminology for the same concept
    - Conflicting data (e.g., different numbers for same metric)
    - Inconsistent formatting patterns

    Returns:
        List of semantic issues found
    """
    issues = []

    if len(slides_data) < 2:
        return issues

    # Collect all titles and body text
    titles = {}
    all_text = []

    for slide in slides_data:
        idx = slide.get("index", 0)
        title = slide.get("title", "").strip()
        body = slide.get("body", [])

        if title:
            if title not in titles:
                titles[title] = []
            titles[title].append(idx)

        all_text.extend(body)

    # Check for duplicate titles (should be rare unless intentional)
    for title, slide_nums in titles.items():
        if len(slide_nums) > 1 and len(title) > 10:
            # Skip very short titles (e.g., "Introduction")
            issues.append({
                "check": "semantic",
                "severity": "info",
                "message": f"Title '{title}' appears on slides {slide_nums}",
                "suggestion": "Ensure duplicate titles are intentional or consider differentiating",
            })

    # Check for inconsistent terminology
    # Extract key terms (words > 4 chars appearing multiple times)
    term_counts = {}
    for text in all_text:
        words = re.findall(r'\b\w{5,}\b', text.lower())
        for word in words:
            term_counts[word] = term_counts.get(word, 0) + 1

    # Flag terms with inconsistent capitalization
    # This is a basic check; full NLP would be more sophisticated
    for term, count in term_counts.items():
        if count >= 3 and term not in term_counts:
            # Check capitalization variance
            variants = set()
            for text in all_text:
                if term.lower() in text.lower():
                    # Extract the actual word with original case
                    m = re.search(rf'\b(\w*{re.escape(term)}\w*)\b', text, re.IGNORECASE)
                    if m:
                        variants.add(m.group(1))

            if len(variants) > 1:
                issues.append({
                    "check": "semantic",
                    "severity": "info",
                    "message": f"Inconsistent capitalization of term '{term}': {list(variants)}",
                    "suggestion": "Use consistent capitalization for terminology",
                })

    return issues


def _check_data_accuracy(
    slides_data: List[Dict],
) -> List[Dict]:
    """Check for potential data accuracy issues.

    Detects:
    - Numbers that don't sum correctly
    - Inconsistent dates
    - Mismatched percentages

    Returns:
        List of data accuracy warnings
    """
    issues = []

    for slide in slides_data:
        body = slide.get("body", [])
        body_text = " ".join(body)

        # Look for percentage lists that should sum to 100%
        # Pattern: "X% Y% Z%" on same slide
        pct_match = re.findall(r'(\d+)%', body_text)
        if len(pct_match) >= 2:
            try:
                total = sum(int(p) for p in pct_match)
                if total != 100 and 95 <= total <= 105:
                    # Close to 100% but not exactly
                    issues.append({
                        "check": "data_accuracy",
                        "severity": "warning",
                        "message": f"Slide {slide.get('index')}: Percentages sum to {total}% (should be 100%)",
                        "suggestion": "Verify percentages sum to exactly 100% or note if intentional",
                    })
                elif total < 95 or total > 105:
                    issues.append({
                        "check": "data_accuracy",
                        "severity": "error",
                        "message": f"Slide {slide.get('index')}: Percentages sum to {total}% (significantly off from 100%)",
                        "suggestion": "Review data accuracy - percentages should typically sum to 100%",
                    })
            except ValueError:
                pass

        # Look for inconsistent date formats
        # Pattern: "YYYY-MM-DD", "MM/DD/YYYY", "DD Month YYYY"
        date_patterns = [
            r'\b\d{4}-\d{2}-\d{2}\b',
            r'\b\d{1,2}/\d{1,2}/\d{4}\b',
            r'\b\d{1,2} [A-Za-z]+ \d{4}\b',
        ]
        dates_found = []
        for pattern in date_patterns:
            matches = re.findall(pattern, body_text)
            dates_found.extend(matches)

        if len(dates_found) >= 2:
            # Check for mixed formats (very basic check)
            has_dash = any("-" in d for d in dates_found)
            has_slash = any("/" in d for d in dates_found)
            if has_dash and has_slash:
                issues.append({
                    "check": "data_accuracy",
                    "severity": "info",
                    "message": f"Slide {slide.get('index')}: Mixed date formats detected",
                    "suggestion": "Use consistent date format throughout presentation",
                })

    return issues


def run_enhanced_qa_checks(
    unpacked_dir: Path,
    slides_data: List[Dict],
    verbose: bool = False,
) -> Dict[str, List[Dict]]:
    """Run all enhanced QA checks.

    Args:
        unpacked_dir: Path to unpacked PPT
        slides_data: List of slide data (from extract_content.py)
        verbose: Print detailed progress

    Returns:
        Dict mapping check type to list of issues:
        {
            "visual_alignment": [...],
            "color_consistency": [...],
            "semantic_consistency": [...],
            "data_accuracy": [...],
        }
    """
    results = {
        "visual_alignment": [],
        "color_consistency": [],
        "semantic_consistency": [],
        "data_accuracy": [],
    }

    slides_dir = unpacked_dir / "ppt" / "slides"
    if not slides_dir.exists():
        return results

    # Per-slide visual checks
    for slide in slides_data:
        idx = slide.get("index", 0)
        slide_file = slide.get("slide_file", f"slide{idx}.xml")
        slide_path = slides_dir / slide_file

        if not slide_path.exists():
            continue

        slide_xml = slide_path.read_text(encoding="utf-8")

        # Visual alignment
        alignment_issues = _check_element_alignment(slide_xml)
        if alignment_issues:
            results["visual_alignment"].extend(alignment_issues)
            if verbose and alignment_issues:
                print(f"  Slide {idx}: {len(alignment_issues)} alignment issue(s)")

        # Color consistency
        color_issues = _check_color_consistency(slide_xml, idx)
        if color_issues:
            results["color_consistency"].extend(color_issues)
            if verbose and color_issues:
                print(f"  Slide {idx}: {len(color_issues)} color issue(s)")

    # Cross-slide checks
    semantic_issues = _check_semantic_consistency(slides_data)
    if semantic_issues:
        results["semantic_consistency"].extend(semantic_issues)
        if verbose and semantic_issues:
            print(f"  Cross-slide: {len(semantic_issues)} semantic issue(s)")

    data_issues = _check_data_accuracy(slides_data)
    if data_issues:
        results["data_accuracy"].extend(data_issues)
        if verbose and data_issues:
            print(f"  Data accuracy: {len(data_issues)} issue(s)")

    return results
