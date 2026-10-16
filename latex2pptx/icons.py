"""
FontAwesome and math symbol glyph isolation, high-DPI rasterization, and tight alpha cropping.
"""

import io
import re
from typing import Any, Dict, List
from PIL import Image
import pymupdf
from pptx.util import Pt


def make_icons_only_doc(pdf_path: str) -> pymupdf.Document:
    """
    Produce an in-memory copy of the PDF where all vector drawings, images,
    and non-icon text glyphs are stripped, preserving only icon glyphs and
    their exact coordinate matrices.
    """
    doc = pymupdf.open(pdf_path)
    for page in doc:
        icon_fonts = {
            f[4]
            for f in page.get_fonts()
            if any(k in f[3] for k in ("FontAwesome", "CMSY", "CMEX", "Dingbats", "Symbol"))
        }
        for xref in page.get_contents():
            text = doc.xref_stream(xref).decode("latin1")
            parts = re.split(r"(BT\b.*?ET\b)", text, flags=re.DOTALL)
            out = []
            for i, part in enumerate(parts):
                if i % 2 == 1:
                    tokens = re.split(r"(/F\d+\s+[\d.]+\s+Tf)", part)
                    cur_is_icon = False
                    bt_out = []
                    for tok in tokens:
                        m = re.match(r"^/(F\d+)\s+[\d.]+\s+Tf$", tok)
                        if m:
                            cur_is_icon = m.group(1) in icon_fonts
                            bt_out.append(tok)
                        else:
                            if not cur_is_icon:
                                tok = re.sub(r"\[[^\]]*\]\s*TJ", "[]TJ", tok)
                                tok = re.sub(r"\([^)]*\)\s*Tj", "()Tj", tok)
                            bt_out.append(tok)
                    out.append("".join(bt_out))
                else:
                    p = re.sub(r"/\w+\s+Do\b", "", part)
                    p = re.sub(r"(?<![A-Za-z0-9_/])(f\*|F|f|S|s|B\*|B|b\*|b)(?![A-Za-z0-9_])", "n", p)
                    out.append(p)
            doc.update_stream(xref, "".join(out).encode("latin1"))
    return doc


def add_icons_to_slide(
    slide: Any,
    icon_page: Any,
    icons: List[Dict[str, Any]],
    page_rect: Any,
    supersample: float = 8.0,
) -> None:
    """Render and add isolated icon glyphs with sub-pixel tight alpha cropping."""
    S = supersample
    for ic in icons:
        ib = pymupdf.Rect(ic["bbox"])
        clip = pymupdf.Rect(ib.x0 - 1.5, ib.y0 - 1.5, ib.x1 + 1.5, ib.y1 + 1.5) & page_rect
        if clip.width <= 0.1 or clip.height <= 0.1:
            continue

        pix = icon_page.get_pixmap(clip=clip, matrix=pymupdf.Matrix(S, S), alpha=True)
        pil_img = Image.frombytes("RGBA", (pix.width, pix.height), pix.samples)
        tight = pil_img.getbbox()
        if not tight:
            continue

        cropped = pil_img.crop(tight)
        true_x0 = clip.x0 + tight[0] / S
        true_y0 = clip.y0 + tight[1] / S
        true_w = (tight[2] - tight[0]) / S
        true_h = (tight[3] - tight[1]) / S

        buf = io.BytesIO()
        cropped.save(buf, format="PNG")
        buf.seek(0)
        slide.shapes.add_picture(buf, Pt(true_x0), Pt(true_y0), Pt(true_w), Pt(true_h))
