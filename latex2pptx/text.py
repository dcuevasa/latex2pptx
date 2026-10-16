"""
Text extraction, paragraph grouping, alignment detection, and native PowerPoint text box generation.
"""

import math
import os
from typing import Any, Dict, List, Optional, Tuple
from PIL import ImageFont
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Emu, Pt

_FONT_PATHS = {
    ("Liberation Sans", False, False): "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ("Liberation Sans", True, False): "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    ("Liberation Sans", False, True): "/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf",
    ("Liberation Sans", True, True): "/usr/share/fonts/truetype/liberation/LiberationSans-BoldItalic.ttf",
    ("Liberation Mono", False, False): "/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf",
    ("Liberation Mono", True, False): "/usr/share/fonts/truetype/liberation/LiberationMono-Bold.ttf",
    ("Liberation Mono", False, True): "/usr/share/fonts/truetype/liberation/LiberationMono-Italic.ttf",
    ("Liberation Mono", True, True): "/usr/share/fonts/truetype/liberation/LiberationMono-BoldItalic.ttf",
}
_LOADED_FONTS: Dict[Tuple[str, bool, bool], Optional[Any]] = {}


def _get_pil_font(family: str, bold: bool, italic: bool) -> Optional[Any]:
    """Load and cache a 1000-pt PIL FreeType font for accurate advance width calculation."""
    key = (family, bold, italic)
    if key not in _LOADED_FONTS:
        path = _FONT_PATHS.get(key) or _FONT_PATHS.get((family, bold, False)) or _FONT_PATHS.get(("Liberation Sans", False, False))
        if path and os.path.exists(path):
            try:
                _LOADED_FONTS[key] = ImageFont.truetype(path, 1000)
            except Exception:
                _LOADED_FONTS[key] = None
        else:
            _LOADED_FONTS[key] = None
    return _LOADED_FONTS[key]


def measure_line_width_pt(ln: Dict[str, Any]) -> float:
    """Measure the unscaled width in points of a line when rendered in its target PowerPoint font."""
    total_w = 0.0
    for r in ln["runs"]:
        font = _get_pil_font(r["family"], r["bold"], r["italic"])
        if font is not None:
            total_w += font.getlength(r["text"]) * (r["size"] / 1000.0)
        else:
            # Fallback heuristic if system TTF is unavailable
            factor = 0.60 if r["family"] == "Liberation Mono" else (0.55 if r["bold"] else 0.52)
            total_w += len(r["text"]) * r["size"] * factor
    return total_w


def compute_paragraph_font_scale(p_lines: List[Dict[str, Any]], align: str) -> float:
    """
    Compute a uniform font size scale factor for a paragraph so that target fonts
    (e.g. Liberation Sans / Arial) match narrow LaTeX Sans (LMSans) x-height and
    never overflow the PDF line bounding boxes horizontally.
    """
    is_bold = any(r["bold"] for ln in p_lines for r in ln["runs"])
    is_mono = any(r["family"] == "Liberation Mono" for ln in p_lines for r in ln["runs"])
    if is_mono:
        base_scale = 0.96
    elif is_bold:
        base_scale = 0.925
    else:
        base_scale = 0.94

    max_pdf_w = max(ln["x1"] - ln["x0"] for ln in p_lines)
    scale = base_scale
    for ln in p_lines:
        pdf_w = ln["x1"] - ln["x0"]
        ft_w = measure_line_width_pt(ln)
        if ft_w > 1.0:
            slack = 1.5 if (align == "left" and pdf_w < 35.0) else 0.5
            avail_w = max(pdf_w + slack, 0.92 * max_pdf_w)
            scale = min(scale, avail_w / ft_w)
    return max(0.85, scale)



def rgb_from_int(val: int) -> RGBColor:
    """Convert a 24-bit integer 0xRRGGBB into an RGBColor."""
    return RGBColor((val >> 16) & 0xFF, (val >> 8) & 0xFF, val & 0xFF)


def is_icon_font(font_name: str) -> bool:
    """Check if font name represents an icon or math symbol font."""
    return any(k in font_name for k in ("FontAwesome", "CMSY", "CMEX", "Dingbats", "Symbol"))


def map_font_style(pdf_font_name: str, flags: int) -> Tuple[str, bool, bool]:
    """Map PDF subset font name and flags to system font family, bold, and italic."""
    base = pdf_font_name.split("+")[-1]
    is_bold = bool(flags & 16) or any(k in base for k in ("Bold", "Bx", "Demi", "Black", "Heavy"))
    is_italic = bool(flags & 2) or any(k in base for k in ("Italic", "Oblique", "Slant", "It"))
    if any(k in base for k in ("Mono", "TT", "Courier", "Typewriter", "Consolas")):
        family = "Liberation Mono"
    else:
        family = "Liberation Sans"
    return family, is_bold, is_italic


def extract_lines_and_icons(page: Any) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Extract raw text lines and icon glyphs from a PDF page with accurate
    character kerning, inter-word gap reconstruction, and run merging.
    """
    rdict = page.get_text("rawdict")
    raw_segments = []
    icons = []

    for b in rdict["blocks"]:
        if b["type"] != 0:
            continue
        for l in b["lines"]:
            cur_spans = []
            pre_icon_x0 = None
            for s in l["spans"]:
                if is_icon_font(s["font"]):
                    # Keep non-whitespace glyphs
                    chars = [c for c in s["chars"] if ord(c["c"]) != 32]
                    if chars:
                        x0 = min(c["bbox"][0] for c in chars)
                        y0 = min(c["bbox"][1] for c in chars)
                        x1 = max(c["bbox"][2] for c in chars)
                        y1 = max(c["bbox"][3] for c in chars)
                        icons.append({
                            "font": s["font"],
                            "size": s["size"],
                            "color": s["color"],
                            "bbox": (x0, y0, x1, y1),
                        })
                        if not cur_spans and pre_icon_x0 is None:
                            pre_icon_x0 = x0
                    if cur_spans:
                        raw_segments.append((l["dir"], cur_spans, pre_icon_x0))
                        cur_spans = []
                        pre_icon_x0 = None
                else:
                    chars = s["chars"]
                    if not chars:
                        continue
                    rebuilt = []
                    for i, c in enumerate(chars):
                        rebuilt.append(c["c"])
                        if i + 1 < len(chars):
                            c1, c2 = chars[i], chars[i + 1]
                            if c1["c"] != " " and c2["c"] != " ":
                                gap = (c2["bbox"][0] - c1["bbox"][2]) / max(1.0, s["size"])
                                if gap > 0.085:
                                    rebuilt.append(" ")
                    txt = "".join(rebuilt)
                    if not txt:
                        continue
                    # Check for large horizontal gap indicating separate column
                    if cur_spans:
                        prev_s = cur_spans[-1]
                        gap_pt = s["bbox"][0] - prev_s["bbox"][2]
                        if gap_pt > max(12.0, 1.3 * max(s["size"], prev_s["size"])):
                            raw_segments.append((l["dir"], cur_spans, pre_icon_x0))
                            cur_spans = []
                            pre_icon_x0 = None
                        elif gap_pt > 0.085 * max(s["size"], prev_s["size"]):
                            if not prev_s["text"].endswith(" ") and not txt.startswith(" "):
                                txt = " " + txt
                    s_copy = dict(s)
                    s_copy["text"] = txt
                    cur_spans.append(s_copy)
            if cur_spans:
                raw_segments.append((l["dir"], cur_spans, pre_icon_x0))

    clean_lines = []
    for l_dir, spans, pre_x0 in raw_segments:
        while spans and not spans[0]["text"].strip(" "):
            spans.pop(0)
        while spans and not spans[-1]["text"].strip(" "):
            spans.pop()
        if not spans:
            continue
        spans[0]["text"] = spans[0]["text"].lstrip(" ")
        spans[-1]["text"] = spans[-1]["text"].rstrip(" ")

        # Normalize LaTeX '\dots' ('. . .') -> '...'
        full_check = "".join(s["text"] for s in spans)
        if not full_check.strip(" "):
            continue
        for s in spans:
            s["text"] = s["text"].replace(". . .", "...")

        first_chars = [c for c in spans[0]["chars"] if c["c"] != " "]
        last_chars = [c for c in spans[-1]["chars"] if c["c"] != " "]
        x0 = first_chars[0]["bbox"][0] if first_chars else spans[0]["bbox"][0]
        x1 = last_chars[-1]["bbox"][2] if last_chars else spans[-1]["bbox"][2]
        y0 = min(s["bbox"][1] for s in spans)
        y1 = max(s["bbox"][3] for s in spans)
        origin_y = sum(s["origin"][1] * len(s["text"]) for s in spans) / max(1, sum(len(s["text"]) for s in spans))
        max_size = max(s["size"] for s in spans)

        # Merge adjacent runs with identical typography
        merged_runs = []
        for s in spans:
            fam, bold, italic = map_font_style(s["font"], s["flags"])
            color = rgb_from_int(s["color"])
            sz = round(s["size"], 2)
            style_key = (fam, bold, italic, sz, color)
            if merged_runs and merged_runs[-1]["style"] == style_key:
                merged_runs[-1]["text"] += s["text"]
            else:
                merged_runs.append({
                    "style": style_key,
                    "family": fam,
                    "bold": bold,
                    "italic": italic,
                    "size": s["size"],
                    "color": color,
                    "text": s["text"],
                })

        clean_lines.append({
            "dir": l_dir,
            "runs": merged_runs,
            "x0": x0,
            "x1": x1,
            "y0": y0,
            "y1": y1,
            "cx": 0.5 * (x0 + x1),
            "origin_y": origin_y,
            "size": max_size,
            "pre_icon_x0": pre_x0,
        })
    return clean_lines, icons


def group_lines_into_paragraphs(lines: List[Dict[str, Any]], shapes_rects: List[Any]) -> List[Dict[str, Any]]:
    """Group visual text lines into coherent paragraphs with alignment detection."""
    n = len(lines)
    used = [False] * n
    paragraphs = []

    order = sorted(range(n), key=lambda i: (round(lines[i]["y0"], 1), round(lines[i]["x0"], 1)))

    for idx in order:
        if used[idx]:
            continue
        used[idx] = True
        group = [lines[idx]]
        align_mode = None

        while True:
            last = group[-1]
            best_j = None
            best_dy = 1e9
            best_mode = None

            for j in order:
                if used[j]:
                    continue
                cand = lines[j]
                if abs(last["dir"][0] - cand["dir"][0]) > 0.02 or abs(last["dir"][1] - cand["dir"][1]) > 0.02:
                    continue
                if abs(last["size"] - cand["size"]) > 0.8:
                    continue
                dy_base = cand["origin_y"] - last["origin_y"]
                v_gap = cand["y0"] - last["y1"]
                if dy_base < 0.75 * last["size"] or v_gap > max(4.2, 0.42 * last["size"]):
                    continue
                if cand["pre_icon_x0"] is not None:
                    continue
                overlap = min(last["x1"], cand["x1"]) - max(last["x0"], cand["x0"])
                if overlap < -2.0:
                    continue

                left_match = abs(last["x0"] - cand["x0"]) <= 2.5
                center_match = abs(last["cx"] - cand["cx"]) <= 3.5
                right_match = abs(last["x1"] - cand["x1"]) <= 2.5

                mode = None
                if align_mode == "left" and left_match:
                    mode = "left"
                elif align_mode == "center" and center_match:
                    mode = "center"
                elif align_mode == "right" and right_match:
                    mode = "right"
                elif align_mode is None:
                    if left_match and (not center_match or abs(last["x0"] - cand["x0"]) <= abs(last["cx"] - cand["cx"])):
                        mode = "left"
                    elif center_match:
                        mode = "center"
                    elif right_match:
                        mode = "right"

                if mode is not None and dy_base < best_dy:
                    best_dy = dy_base
                    best_j = j
                    best_mode = mode

            if best_j is not None:
                used[best_j] = True
                group.append(lines[best_j])
                if align_mode is None:
                    align_mode = best_mode
            else:
                break

        paragraphs.append({"lines": group, "align": align_mode})

    # Determine alignment for single-line paragraphs
    for p in paragraphs:
        if p["align"] is not None:
            continue
        ln = p["lines"][0]
        # Never center lines that had an icon directly preceding them
        if ln["pre_icon_x0"] is not None:
            p["align"] = "left"
            continue

        # Check for right-aligned vertical columns (e.g. chart value labels)
        right_peers = [
            other["lines"][0]
            for other in paragraphs
            if len(other["lines"]) == 1
            and other["lines"][0] is not ln
            and abs(other["lines"][0]["size"] - ln["size"]) < 0.8
            and abs(other["lines"][0]["x1"] - ln["x1"]) <= 1.5
            and abs(other["lines"][0]["x0"] - ln["x0"]) > 3.0
        ]
        if len(right_peers) >= 2:
            p["align"] = "right"
            continue

        # Check if centered inside an enclosing shape
        enclosing = [
            r
            for r in shapes_rects
            if r.x0 - 2.0 <= ln["x0"]
            and ln["x1"] <= r.x1 + 2.0
            and r.y0 - 2.0 <= ln["y0"]
            and ln["y1"] <= r.y1 + 2.0
            and r.width < 400.0
        ]
        if enclosing:
            smallest = min(enclosing, key=lambda r: r.width * r.height)
            shape_cx = 0.5 * (smallest.x0 + smallest.x1)
            if abs(ln["cx"] - shape_cx) <= 3.0:
                p["align"] = "center"
                continue

        p["align"] = "left"

    return paragraphs


def add_paragraphs_to_slide(slide: Any, paragraphs: List[Dict[str, Any]], page_rect: Any) -> None:
    """Generate native PowerPoint text boxes with calibrated baseline alignment."""
    for p_info in paragraphs:
        p_lines = p_info["lines"]
        align = p_info["align"]
        first = p_lines[0]
        dx, dy = first["dir"]
        angle_deg = math.degrees(math.atan2(dy, dx))

        p_scale = compute_paragraph_font_scale(p_lines, align)
        scaled_first_size = first["size"] * p_scale

        x0 = min(ln["x0"] for ln in p_lines)
        x1 = max(ln["x1"] for ln in p_lines)
        raw_w = max(4.0, x1 - x0)

        if abs(angle_deg) > 0.5:
            # Rotated text box
            y0 = min(ln["y0"] for ln in p_lines)
            y1 = max(ln["y1"] for ln in p_lines)
            cx = 0.5 * (x0 + x1)
            cy = 0.5 * (y0 + y1)
            box_w = raw_w * 1.25 + 16.0
            box_h = scaled_first_size * 1.35
            left = cx - 0.5 * box_w
            top = cy - 0.45 * box_h
            align = "center"
        else:
            pad_w = max(14.0, raw_w * 0.14)
            if align == "center":
                cx = sum(ln["cx"] for ln in p_lines) / len(p_lines)
                box_w = raw_w + 2.0 * pad_w
                left = cx - 0.5 * box_w
            elif align == "right":
                box_w = raw_w + pad_w
                left = x1 - box_w
            else:
                left = x0
                box_w = min(page_rect.width - left + 10.0, raw_w + pad_w)

            # Calibrated baseline formula: top = first.origin_y - 1.0 * scaled_first_size
            top = first["origin_y"] - 1.0 * scaled_first_size
            box_h = (p_lines[-1]["origin_y"] - first["origin_y"]) + scaled_first_size * 1.35

        tb = slide.shapes.add_textbox(Pt(left), Pt(top), Pt(box_w), Pt(box_h))
        if abs(angle_deg) > 0.5:
            tb.rotation = angle_deg

        tf = tb.text_frame
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.word_wrap = False
        tf.margin_left = Emu(0)
        tf.margin_right = Emu(0)
        tf.margin_top = Emu(0)
        tf.margin_bottom = Emu(0)

        pp_align = (
            PP_ALIGN.CENTER
            if align == "center"
            else (PP_ALIGN.RIGHT if align == "right" else PP_ALIGN.LEFT)
        )

        for idx, ln in enumerate(p_lines):
            para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            para.alignment = pp_align
            para.space_before = Pt(0)
            para.space_after = Pt(0)
            if idx > 0:
                step_pt = ln["origin_y"] - p_lines[idx - 1]["origin_y"]
                para.line_spacing = Pt(max(1.0, step_pt))
            for r_info in ln["runs"]:
                run = para.add_run()
                run.text = r_info["text"]
                run.font.name = r_info["family"]
                run.font.size = Pt(r_info["size"] * p_scale)
                run.font.bold = r_info["bold"]
                run.font.italic = r_info["italic"]
                run.font.color.rgb = r_info["color"]
