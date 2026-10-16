"""
Image extraction, SMask transparency handling, and TikZ clip mask rasterization.
"""

import re
from typing import Any, Dict, List, Tuple
import pymupdf


def make_clipped_images_doc(pdf_path: str) -> pymupdf.Document:
    """
    Produce an in-memory copy of the PDF where text and vector shapes are stripped,
    leaving image XObjects and active PDF clipping paths (W n) intact.
    """
    doc = pymupdf.open(pdf_path)
    for page in doc:
        for xref in page.get_contents():
            text = doc.xref_stream(xref).decode("latin1")
            text = re.sub(r"BT\b.*?ET\b", "", text, flags=re.DOTALL)
            text = re.sub(r"(?<![A-Za-z0-9_/])(f\*|F|f|S|s|B\*|B|b\*|b)(?![A-Za-z0-9_])", "n", text)
            doc.update_stream(xref, text.encode("latin1"))
    return doc


def extract_image_bytes(
    doc: pymupdf.Document,
    clip_page: Any,
    img_info: Dict[str, Any],
    clip_rects: List[Any],
) -> Tuple[bytes, pymupdf.Rect]:
    """
    Extract image bytes (PNG) preserving either native SMask transparency
    or applying TikZ clipping masks.
    """
    bbox = pymupdf.Rect(img_info["bbox"])
    page_rect = doc[0].rect
    is_clipped = any(
        bbox.intersects(cr) and (cr.width < page_rect.width * 0.95 or cr.height < page_rect.height * 0.95)
        for cr in clip_rects
    )
    xref = img_info.get("xref", 0)

    if not is_clipped and xref > 0:
        try:
            base_pix = pymupdf.Pixmap(doc, xref)
            smask_info = doc.xref_get_key(xref, "SMask")
            if smask_info and smask_info[0] == "xref":
                smask_xref = int(smask_info[1].split()[0])
                mask_pix = pymupdf.Pixmap(doc, smask_xref)
                if base_pix.alpha:
                    base_pix = pymupdf.Pixmap(pymupdf.csRGB, base_pix)
                pix = pymupdf.Pixmap(base_pix, mask_pix)
            else:
                if base_pix.colorspace and base_pix.colorspace.n > 3:
                    pix = pymupdf.Pixmap(pymupdf.csRGB, base_pix)
                else:
                    pix = base_pix
            return pix.tobytes("png"), bbox
        except Exception:
            pass

    # Render clipped image using native PDF clip path
    pix = clip_page.get_pixmap(clip=bbox, matrix=pymupdf.Matrix(4, 4), alpha=True)
    return pix.tobytes("png"), bbox
