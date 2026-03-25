#!/usr/bin/env python3
"""
AI Auto Resize Module — 自动调整不同尺寸 PPT 的内容布局

支持策略:
- smart: AI 智能重排（推荐）
- scale: 等比例缩放居中（保守）
- stretch: 拉伸填充（可能导致变形）
- crop: 裁剪超出部分（可能丢失内容）

作者: PPTX Studio
版本: 1.0.0
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Dict, Literal, Optional, Tuple, List
from dataclasses import dataclass
import xml.etree.ElementTree as ET

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Common slide sizes (EMU - English Metric Units)
# 1 inch = 914400 EMU
SLIDE_SIZES = {
    "16:9": (12192000, 6858000),  # 10" × 5.625"
    "4:3": (9144000, 6858000),   # 9.6" × 7.2"
    "16:10": (12192000, 7620000), # 10" × 6.25"
}

# Common font size adjustment factors based on aspect ratio change
FONT_SCALE_FACTORS = {
    "4:3_to_16:9": 1.0,   # No scaling needed (similar area)
    "16:9_to_4:3": 0.85,  # Reduce font size to fit narrower width
}


# ---------------------------------------------------------------------------
# Data Classes
# ---------------------------------------------------------------------------

@dataclass
class SlideSize:
    """Slide size information"""
    width: int      # EMU
    height: int     # EMU
    aspect_ratio: str  # e.g., "16:9", "4:3"

    @property
    def width_inches(self) -> float:
        return self.width / 914400

    @property
    def height_inches(self) -> float:
        return self.height / 914400

    @property
    def aspect_ratio_value(self) -> float:
        return self.width / self.height


@dataclass
class ResizeStrategy:
    """Resize strategy and parameters"""
    strategy: Literal["smart", "scale", "stretch", "crop"]
    scale_x: float
    scale_y: float
    offset_x: int  # EMU
    offset_y: int  # EMU
    font_scale: float
    warnings: List[str]


# ---------------------------------------------------------------------------
# Core Functions
# ---------------------------------------------------------------------------

def detect_slide_size(pptx_path: str) -> SlideSize:
    """检测 PPT 尺寸 (width, height, aspect_ratio)

    Args:
        pptx_path: Path to .pptx file

    Returns:
        SlideSize object with width, height, and aspect ratio
    """
    from pptx import Presentation

    prs = Presentation(pptx_path)
    width = prs.slide_width
    height = prs.slide_height

    # Detect aspect ratio
    aspect_ratio = detect_aspect_ratio(width, height)

    return SlideSize(width=width, height=height, aspect_ratio=aspect_ratio)


def detect_aspect_ratio(width: int, height: int) -> str:
    """从 width/height 检测宽高比

    Args:
        width: Slide width in EMU
        height: Slide height in EMU

    Returns:
        Aspect ratio string (e.g., "16:9", "4:3", "16:10")
    """
    ratio = width / height

    # Tolerance for aspect ratio matching
    tolerance = 0.05

    # Check against common sizes
    for name, (std_w, std_h) in SLIDE_SIZES.items():
        std_ratio = std_w / std_h
        if abs(ratio - std_ratio) < tolerance:
            return name

    # Fallback: calculate from current size
    simplified = simplify_ratio(int(ratio * 100), 100)
    return f"{simplified[0]}:{simplified[1]}"


def simplify_ratio(width: int, height: int) -> Tuple[int, int]:
    """简化宽高比

    Args:
        width: Width value
        height: Height value

    Returns:
        Simplified (width, height) tuple
    """
    from math import gcd

    divisor = gcd(width, height)
    return (width // divisor, height // divisor)


def calculate_resize_strategy(
    source_size: SlideSize,
    target_size: SlideSize,
    strategy: str = "smart"
) -> ResizeStrategy:
    """计算尺寸调整策略和参数

    Args:
        source_size: Source slide size
        target_size: Target slide size
        strategy: Resize strategy (smart/scale/stretch/crop)

    Returns:
        ResizeStrategy with scale factors and offsets
    """
    warnings = []

    # Calculate scale factors
    if strategy == "scale":
        # Maintain aspect ratio, scale to fit
        scale_x = min(target_size.width / source_size.width, target_size.height / source_size.height)
        scale_y = scale_x

        # Center the content
        offset_x = int((target_size.width - source_size.width * scale_x) / 2)
        offset_y = int((target_size.height - source_size.height * scale_y) / 2)

        # Font scaling based on aspect ratio change
        key = f"{source_size.aspect_ratio}_to_{target_size.aspect_ratio}"
        font_scale = FONT_SCALE_FACTORS.get(key, 1.0)

    elif strategy == "stretch":
        # Stretch to fill (may cause distortion)
        scale_x = target_size.width / source_size.width
        scale_y = target_size.height / source_size.height
        offset_x = 0
        offset_y = 0
        font_scale = min(scale_x, scale_y)

    elif strategy == "crop":
        # Crop to fill (may lose content)
        scale_x = max(target_size.width / source_size.width, target_size.height / source_size.height)
        scale_y = scale_x

        # Center the content (overflow will be cropped)
        offset_x = int((target_size.width - source_size.width * scale_x) / 2)
        offset_y = int((target_size.height - source_size.height * scale_y) / 2)
        font_scale = scale_x

    else:  # smart (default)
        # Smart: Use scale strategy for now
        # TODO: Implement AI layout analysis in Phase 2
        warnings.append("Smart resize not fully implemented yet, using scale strategy")
        return calculate_resize_strategy(source_size, target_size, "scale")

    # Add warnings if scale is extreme
    if scale_x > 1.5 or scale_x < 0.67:
        warnings.append(f"Significant scale change: {scale_x:.2f}x")
    if scale_y > 1.5 or scale_y < 0.67:
        warnings.append(f"Significant scale change: {scale_y:.2f}x")

    return ResizeStrategy(
        strategy=strategy,
        scale_x=scale_x,
        scale_y=scale_y,
        offset_x=offset_x,
        offset_y=offset_y,
        font_scale=font_scale,
        warnings=warnings
    )


def resize_slide_xml(
    slide_xml: str,
    source_size: SlideSize,
    target_size: SlideSize,
    strategy: str = "smart",
    verbose: bool = False
) -> str:
    """调整单页幻灯片尺寸

    Args:
        slide_xml: Slide XML content
        source_size: Source slide size
        target_size: Target slide size
        strategy: Resize strategy (smart/scale/stretch/crop)
        verbose: Print debug information

    Returns:
        Resized slide XML
    """
    # Calculate resize strategy
    resize_info = calculate_resize_strategy(source_size, target_size, strategy)

    if verbose:
        print(f"Resize strategy: {resize_info.strategy}")
        print(f"Scale: {resize_info.scale_x:.2f}x × {resize_info.scale_y:.2f}x")
        print(f"Offset: ({resize_info.offset_x}, {resize_info.offset_y}) EMU")
        if resize_info.warnings:
            for w in resize_info.warnings:
                print(f"Warning: {w}")

    # Apply transformations
    xml = slide_xml

    # 1. Resize shapes (position and size)
    xml = _resize_shapes(xml, resize_info)

    # 2. Adjust font sizes
    xml = _adjust_font_sizes(xml, resize_info)

    # 3. Adjust background elements
    xml = _adjust_background(xml, resize_info)

    return xml


def _resize_shapes(xml: str, resize_info: ResizeStrategy) -> str:
    """调整所有形状的位置和大小

    Args:
        xml: Slide XML
        resize_info: Resize strategy

    Returns:
        Updated XML
    """
    def transform_xfrm(m):
        xfrm_xml = m.group(0)
        
        # Extract current x and y
        x_match = re.search(r'<a:off[^>]*x="([^"]+)"', xfrm_xml)
        y_match = re.search(r'<a:off[^>]*y="([^"]+)"', xfrm_xml)
        
        # Extract current cx and cy
        cx_match = re.search(r'<a:ext[^>]*cx="([^"]+)"', xfrm_xml)
        cy_match = re.search(r'<a:ext[^>]*cy="([^"]+)"', xfrm_xml)
        
        if not any([x_match, y_match, cx_match, cy_match]):
            return xfrm_xml  # No transform, keep original
        
        # Parse and transform
        new_xml = xfrm_xml
        
        if x_match:
            x = int(x_match.group(1))
            new_x = int(x * resize_info.scale_x + resize_info.offset_x)
            new_xml = re.sub(
                r'x="[^"]+"',
                f'x="{new_x}"',
                new_xml,
                count=1
            )
        
        if y_match:
            y = int(y_match.group(1))
            new_y = int(y * resize_info.scale_y + resize_info.offset_y)
            new_xml = re.sub(
                r'y="[^"]+"',
                f'y="{new_y}"',
                new_xml,
                count=1
            )
        
        if cx_match:
            cx = int(cx_match.group(1))
            new_cx = int(cx * resize_info.scale_x)
            new_xml = re.sub(
                r'cx="[^"]+"',
                f'cx="{new_cx}"',
                new_xml,
                count=1
            )
        
        if cy_match:
            cy = int(cy_match.group(1))
            new_cy = int(cy * resize_info.scale_y)
            new_xml = re.sub(
                r'cy="[^"]+"',
                f'cy="{new_cy}"',
                new_xml,
                count=1
            )
        
        return new_xml
    
    # Apply to all xfrm elements
    xml = re.sub(r'<a:xfrm\b[^>]*>.*?</a:xfrm>', transform_xfrm, xml, flags=re.DOTALL)
    
    return xml


def _adjust_font_sizes(xml: str, resize_info: ResizeStrategy) -> str:
    """调整字体大小

    Args:
        xml: Slide XML
        resize_info: Resize strategy

    Returns:
        Updated XML
    """
    if abs(resize_info.font_scale - 1.0) < 0.01:
        return xml  # No significant change needed
    
    def transform_sz(m):
        rPr_xml = m.group(0)
        sz_match = re.search(r'<a:sz[^>]*val="([^"]+)"', rPr_xml)
        
        if not sz_match:
            return rPr_xml
        
        sz = int(sz_match.group(1))
        new_sz = int(sz * resize_info.font_scale)
        
        # Don't make fonts too small or too large
        new_sz = max(1000, min(new_sz, 72000))  # 8pt - 72pt range in half-points
        
        return re.sub(
            r'val="[^"]+"',
            f'val="{new_sz}"',
            rPr_xml,
            count=1
        )
    
    # Apply to all rPr elements with sz
    xml = re.sub(r'<a:rPr\b[^>]*>.*?</a:rPr>', transform_sz, xml, flags=re.DOTALL)
    
    return xml


def _adjust_background(xml: str, resize_info: ResizeStrategy) -> str:
    """调整背景元素

    Args:
        xml: Slide XML
        resize_info: Resize strategy

    Returns:
        Updated XML
    """
    # Extract background elements (p:bg)
    # For now, keep backgrounds as-is
    # TODO: Implement background resizing in Phase 2
    
    return xml


# ---------------------------------------------------------------------------
# Layout Analysis (Phase 2)
# ---------------------------------------------------------------------------

def analyze_slide_layout(slide_xml: str) -> Dict:
    """分析幻灯片布局（AI 辅助）

    Args:
        slide_xml: Slide XML content

    Returns:
        Layout analysis dict:
        {
            "layout_type": "title" | "content" | "two_column" | "image_center" | "chart",
            "elements": [...],
            "hierarchy": {...}
        }
    """
    # TODO: Implement AI layout analysis in Phase 2
    # For now, return basic layout detection
    
    # Detect title
    has_title = bool(re.search(r'<p:ph[^>]*type="title"', slide_xml))
    
    # Detect content
    has_body = bool(re.search(r'<p:ph[^>]*type="body"', slide_xml))
    
    # Determine layout type
    if has_title and not has_body:
        layout_type = "title"
    elif has_title and has_body:
        layout_type = "content"
    elif not has_title and has_body:
        layout_type = "content"
    else:
        layout_type = "image_center"
    
    return {
        "layout_type": layout_type,
        "elements": [],
        "hierarchy": {}
    }


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def parse_size_spec(size_spec: str) -> Tuple[int, int]:
    """解析尺寸规格

    Args:
        size_spec: Size spec (e.g., "16:9", "4:3")

    Returns:
        (width, height) tuple in EMU
    """
    if size_spec in SLIDE_SIZES:
        return SLIDE_SIZES[size_spec]
    
    # Try to parse as "W:H" ratio
    if ":" in size_spec:
        try:
            w_str, h_str = size_spec.split(":")
            w = int(w_str)
            h = int(h_str)
            
            # Calculate EMU based on ratio and standard 16:9 width
            std_width = SLIDE_SIZES["16:9"][0]
            scale = w / 16
            width = int(std_width * scale)
            height = int(width * h / w)
            
            return (width, height)
        except (ValueError, ZeroDivisionError):
            pass
    
    # Fallback to 16:9
    return SLIDE_SIZES["16:9"]


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="AI Auto Resize - Automatically adjust PPT layout for different slide sizes",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "source",
        help="Source PPTX file"
    )
    parser.add_argument(
        "output",
        help="Output PPTX file"
    )
    parser.add_argument(
        "--target-size",
        choices=["16:9", "4:3", "16:10", "auto"],
        default="auto",
        help="Target slide size (default: auto = use output file size)"
    )
    parser.add_argument(
        "--resize-strategy",
        choices=["smart", "scale", "stretch", "crop"],
        default="smart",
        help="Resize strategy (default: smart)"
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Print debug information"
    )

    args = parser.parse_args()

    # Detect sizes
    source_size = detect_slide_size(args.source)
    
    if args.target_size == "auto":
        # Check if output file exists
        if Path(args.output).exists():
            target_size = detect_slide_size(args.output)
        else:
            print("Warning: Output file does not exist, using 16:9 as target size")
            target_size = SlideSize(*SLIDE_SIZES["16:9"], "16:9")
    else:
        target_width, target_height = parse_size_spec(args.target_size)
        target_size = SlideSize(target_width, target_height, args.target_size)

    # Print summary
    print(f"Source size: {source_size.aspect_ratio} ({source_size.width_inches:.2f}\" × {source_size.height_inches:.2f}\")")
    print(f"Target size: {target_size.aspect_ratio} ({target_size.width_inches:.2f}\" × {target_size.height_inches:.2f}\")")

    # TODO: Implement full resize workflow
    print(f"\nResize strategy: {args.resize_strategy}")
    print("Note: Full resize workflow will be implemented in Phase 2")
