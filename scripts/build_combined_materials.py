#!/usr/bin/env python3
"""Build materials/combined_materials.md and .docx from legacy Word/PDF files."""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

os.chdir(ROOT)

from app import (  # noqa: E402
    _build_combined_materials_docx,
    _build_combined_materials_md,
    _export_feedback_prompt_md,
    COMBINED_MATERIALS_DOCX,
    COMBINED_MATERIALS_MD,
)


def main():
    if _build_combined_materials_md():
        print(f"Created {COMBINED_MATERIALS_MD}")
        _build_combined_materials_docx()
        if os.path.isfile(COMBINED_MATERIALS_DOCX):
            print(f"Created {COMBINED_MATERIALS_DOCX}")
    else:
        print("No source docx/pdf found. Place files in project root or materials/ and retry.")
    _export_feedback_prompt_md()
    print("Done.")


if __name__ == "__main__":
    main()
