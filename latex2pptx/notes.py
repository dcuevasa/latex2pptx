"""
Markdown speaker notes parser and PowerPoint notes integrator.
Supports slide headings, numeric ranges, HTML comment directives, and horizontal rule separators.
"""

import os
import re
from typing import Any, Dict, List, Optional


# Regex for explicit slide number headings:
# Matches:
#   # Slide 1, ## Slide 2: Title, ### Diapositiva 3 - Summary
#   # 1, ## 2. Introduction, # Slide 1-3, ## Slides 4 to 6
#   <!-- slide: 1 -->, <!-- slide: 1-3: Overview -->, <!-- slide 5 -->
HEADER_RE = re.compile(
    r"""^(?:
        [#]{1,4}\s*(?:(?:Slide|Slides|Diapositiva|Diapositivas|Diapo|Diapos|Page|Pages|Página|Páginas|Pág|Págs)\s*)?(\d+)(?:\s*(?:-|to|a)\s*(\d+))?(?:\s*[:.\-–—]\s*(.*))?
        |
        <!--\s*(?:slide|diapo|page)s?\s*[:\s]\s*(\d+)(?:\s*(?:-|to|a)\s*(\d+))?(?:\s*[:.\-–—]\s*(.*?))?\s*-->
    )$""",
    re.IGNORECASE | re.VERBOSE,
)

COMMENT_DELIM_RE = re.compile(r"^<!--\s*(?:slide|diapo|page)\s*-->$", re.IGNORECASE)
HR_RE = re.compile(r"^(?:-{3,}|\*{3,}|_{3,})$")


def parse_markdown_notes(text: str) -> Dict[int, str]:
    """
    Parse markdown notes into a dictionary mapping 1-indexed slide number to notes content.

    Supported syntax:
    1. Explicit Slide Headings / Ranges:
       # Slide 1: Welcome
       ## Slide 2-4: Key Concepts
       ### Diapositiva 5
       <!-- slide: 6 -->

    2. Horizontal Rule Separators (Marp / Slidev style):
       Slide 1 notes
       ---
       Slide 2 notes
       ---
       Slide 3 notes

    3. Generic Top-Level Headings (when no slide numbers are present):
       # First Topic -> Slide 1
       # Second Topic -> Slide 2
    """
    lines = text.splitlines()
    has_explicit_headers = any(bool(HEADER_RE.match(line.strip())) for line in lines)

    notes_by_slide: Dict[int, List[str]] = {}

    if has_explicit_headers:
        current_targets: List[int] = []
        current_buf: List[str] = []

        def flush() -> None:
            if current_targets and current_buf:
                content = "\n".join(current_buf).strip()
                if content:
                    for s in current_targets:
                        notes_by_slide.setdefault(s, []).append(content)

        for line in lines:
            m = HEADER_RE.match(line.strip())
            if m:
                flush()
                current_buf = []
                g = m.groups()
                start_str = g[0] or g[3]
                end_str = g[1] or g[4]
                title = g[2] or g[5]
                start = int(start_str)
                end = int(end_str) if end_str else start
                current_targets = list(range(start, end + 1))
                if title and title.strip():
                    current_buf.append(f"## {title.strip()}")
            else:
                current_buf.append(line)
        flush()
    else:
        # Sequential mode: split on HR (---) or <!-- slide -->
        sections: List[str] = []
        cur: List[str] = []
        for line in lines:
            stripped = line.strip()
            if HR_RE.match(stripped) or COMMENT_DELIM_RE.match(stripped):
                if cur:
                    sections.append("\n".join(cur).strip())
                    cur = []
            else:
                cur.append(line)
        if cur:
            sections.append("\n".join(cur).strip())

        # If no HRs found, check if there are top-level markdown headings (# Heading)
        if len(sections) <= 1 and any(l.startswith("# ") for l in lines):
            sections = []
            cur = []
            for line in lines:
                if line.startswith("# ") and cur:
                    sections.append("\n".join(cur).strip())
                    cur = [line]
                else:
                    cur.append(line)
            if cur:
                sections.append("\n".join(cur).strip())

        for idx, sec in enumerate(sections, start=1):
            if sec:
                notes_by_slide[idx] = [sec]

    return {s: "\n\n".join(chunks) for s, chunks in notes_by_slide.items()}


def load_and_parse_notes(notes_path: str) -> Dict[int, str]:
    """Read a markdown notes file and return a mapping of slide numbers to notes."""
    if not os.path.isfile(notes_path):
        raise FileNotFoundError(f"Speaker notes file not found: {notes_path}")
    with open(notes_path, "r", encoding="utf-8") as f:
        content = f.read()
    return parse_markdown_notes(content)


def apply_notes_to_slide(slide: Any, notes_text: str) -> None:
    """Apply notes text to a PowerPoint slide's notes frame."""
    if not notes_text or not notes_text.strip():
        return
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text.strip()


def apply_notes_to_presentation(prs: Any, notes_map: Dict[int, str]) -> int:
    """
    Apply notes from notes_map to the slides in prs.
    Returns the number of slides that received notes.
    """
    applied_count = 0
    total_slides = len(prs.slides)
    for idx, slide in enumerate(prs.slides):
        slide_num = idx + 1
        if slide_num in notes_map:
            apply_notes_to_slide(slide, notes_map[slide_num])
            applied_count += 1
    return applied_count
