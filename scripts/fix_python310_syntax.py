#!/usr/bin/env python3
"""Fix Python 3.10+ syntax to be compatible with Python 3.9.6

Replaces:
- str | None -> Optional[str]
- int | None -> Optional[int]
- bool | None -> Optional[bool]
- list[T] -> List[T]
- dict[K, V] -> Dict[K, V]
- tuple[T1, T2] -> Tuple[T1, T2]
"""

import re
from pathlib import Path
from typing import Set

# Files to fix (exclude those already using from __future__ import annotations)
FILES_TO_FIX = [
    "scripts/extract_content.py",
    "scripts/apply_template.py",
    "scripts/beautify_ppt.py",
    "scripts/color_ladder.py",
    "scripts/color_ppt.py",
    "scripts/color_replacement.py",
    "scripts/generate_notes.py",
    "scripts/merge_pptx.py",
    "scripts/patch_slide.py",
    "scripts/qa_check.py",
    "scripts/qa_enhanced.py",
    "scripts/table_smartart.py",
    "scripts/thumbnail.py",
    "scripts/add_slide.py",
    "scripts/animation_migration.py",
    "scripts/auto_resize.py",
    "scripts/clean.py",
    "scripts/office/helpers/simplify_redlines.py",
]

# Skip these (already have from __future__ import annotations)
SKIP_FILES = {
    "scripts/office/soffice.py",
}


def fix_type_annotations(content: str) -> str:
    """Fix Python 3.10+ type annotations to Python 3.9 compatible."""
    lines = content.split('\n')
    in_type_annotation = False
    in_future_imports = False

    # Check if file already has from __future__ import annotations
    for line in lines[:20]:  # Check first 20 lines
        if 'from __future__ import annotations' in line:
            in_future_imports = True
            break

    # If already has the import, skip fixing
    if in_future_imports:
        return content

    # Fix type annotations
    result = []
    for line in lines:
        # Skip lines that are comments or strings
        stripped = line.strip()
        if stripped.startswith('#') or stripped.startswith('"""') or stripped.startswith("'''"):
            result.append(line)
            continue

        # Fix str | None, int | None, bool | None, etc.
        # Pattern: word | None
        line = re.sub(r'(\w+)\s*\|\s*None', r'Optional[\1]', line)

        # Fix list[T] -> List[T]
        line = re.sub(r'\blist\[', r'List[', line)

        # Fix dict[K, V] -> Dict[K, V]
        line = re.sub(r'\bdict\[', r'Dict[', line)

        # Fix tuple[T1, T2] -> Tuple[T1, T2]
        line = re.sub(r'\btuple\[', r'Tuple[', line)

        result.append(line)

    return '\n'.join(result)


def main():
    """Fix all Python files in the list."""
    base_dir = Path(__file__).parent.parent
    fixed_count = 0
    skipped_count = 0

    for file_path in FILES_TO_FIX:
        full_path = base_dir / file_path

        if not full_path.exists():
            print(f"  ⚠️  File not found: {file_path}")
            continue

        if file_path in SKIP_FILES:
            print(f"  ⊘ Skipping (already has future import): {file_path}")
            skipped_count += 1
            continue

        print(f"  → Fixing: {file_path}")

        # Read file
        with open(full_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Fix type annotations
        fixed_content = fix_type_annotations(content)

        # Write back
        with open(full_path, 'w', encoding='utf-8') as f:
            f.write(fixed_content)

        fixed_count += 1

    print(f"\n✅ Fixed {fixed_count} files, skipped {skipped_count} files")


if __name__ == "__main__":
    main()
