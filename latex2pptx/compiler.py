"""
LaTeX compilation pipeline with automatic asset discovery, placeholder generation, and multi-pass support.
"""

import os
import re
import shutil
import subprocess
import tempfile
from typing import List, Optional, Set
from PIL import Image, ImageDraw


class CompileError(Exception):
    """Raised when the LaTeX compiler fails."""
    pass


def find_missing_images(tex_content: str, search_dirs: List[str]) -> Set[str]:
    """Scan LaTeX source for referenced images and identify any that are missing."""
    # Find graphicspath declarations
    all_search_dirs = list(search_dirs)
    gp_matches = re.findall(r"\\graphicspath\s*\{((?:\{[^{}]*\}\s*)+)\}", tex_content)
    for m in gp_matches:
        subdirs = re.findall(r"\{([^{}]+)\}", m)
        for sd in subdirs:
            for base in search_dirs:
                p = os.path.normpath(os.path.join(base, sd.strip()))
                if os.path.isdir(p) and p not in all_search_dirs:
                    all_search_dirs.append(p)

    # Find explicit includegraphics calls
    raw_names = set(re.findall(r"\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}", tex_content))

    # Resolve macro parameters like 'bg#1' where \usebg{N} or \newcommand calls exist
    expanded_names = set()
    for name in raw_names:
        if "#1" in name:
            # Look for calls like \usebg{1}, \usebg{2}, etc.
            calls = re.findall(r"\\usebg\{(\d+)\}", tex_content)
            if calls:
                for num in calls:
                    expanded_names.add(name.replace("#1", num))
            else:
                for num in range(1, 16):
                    expanded_names.add(name.replace("#1", str(num)))
        else:
            expanded_names.add(name)

    missing = set()
    exts = ["", ".png", ".jpg", ".jpeg", ".pdf"]
    for img_name in expanded_names:
        found = False
        for sdir in all_search_dirs:
            for ext in exts:
                cand_base = img_name + ext if not img_name.endswith(ext) else img_name
                root, e = os.path.splitext(cand_base)
                for cand_variant in (cand_base, f"{root} (1){e}", f"{root} (2){e}"):
                    cand = os.path.join(sdir, cand_variant)
                    if os.path.isfile(cand):
                        found = True
                        break
                if found:
                    break
            if found:
                break
        if not found:
            missing.add(img_name)

    return missing


def create_placeholder_image(name: str, target_dir: str) -> str:
    """Generate a clean visual placeholder image when an asset is missing."""
    lower = name.lower()
    base, ext = os.path.splitext(name)
    if not ext:
        ext = ".png"

    dest_filename = base + ext
    dest_path = os.path.join(target_dir, dest_filename)

    if lower.startswith("bg") or "background" in lower:
        # 16:9 full slide background placeholder
        img = Image.new("RGB", (1600, 900), (247, 250, 255))
        draw = ImageDraw.Draw(img)
        # Subtle institutional bottom bar
        draw.rectangle([0, 870, 1600, 900], fill=(0, 75, 135))
        draw.rectangle([0, 0, 1600, 6], fill=(43, 108, 176))
    elif any(k in lower for k in ("logo", "gemini", "notebook", "icon", "gmail", "docs", "sheets")):
        img = Image.new("RGBA", (400, 150), (240, 245, 252, 230))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle([2, 2, 397, 147], radius=8, outline=(180, 205, 230), width=2)
    elif any(k in lower for k in ("foto", "photo", "avatar", "pepper")):
        img = Image.new("RGB", (600, 600), (218, 232, 245))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle([5, 5, 594, 594], radius=16, outline=(100, 150, 200), width=3)
    else:
        img = Image.new("RGB", (400, 300), (235, 240, 245))
        draw = ImageDraw.Draw(img)
        draw.rectangle([0, 0, 399, 299], outline=(180, 190, 200), width=2)

    img.save(dest_path)
    return dest_path


def compile_latex(
    tex_path: str,
    output_pdf_path: Optional[str] = None,
    engine: str = "pdflatex",
    passes: int = 2,
    image_dirs: Optional[List[str]] = None,
    allow_placeholders: bool = True,
    keep_temp: bool = False,
    verbose: bool = False,
) -> str:
    """
    Compile a LaTeX file to PDF using multi-pass compilation and automatic
    missing asset placeholder generation.
    """
    if not os.path.isfile(tex_path):
        raise FileNotFoundError(f"LaTeX file not found: {tex_path}")

    tex_dir = os.path.abspath(os.path.dirname(tex_path))
    tex_filename = os.path.basename(tex_path)
    tex_stem = os.path.splitext(tex_filename)[0]

    with open(tex_path, "r", encoding="utf-8", errors="replace") as f:
        tex_content = f.read()

    search_dirs = [tex_dir]
    gp_matches = re.findall(r"\\graphicspath\s*\{((?:\{[^{}]*\}\s*)+)\}", tex_content)
    for m in gp_matches:
        for sd in re.findall(r"\{([^{}]+)\}", m):
            p = os.path.normpath(os.path.join(tex_dir, sd.strip()))
            if os.path.isdir(p) and p not in search_dirs:
                search_dirs.append(p)
    if image_dirs:
        for d in image_dirs:
            if os.path.isdir(d):
                p = os.path.abspath(d)
                if p not in search_dirs:
                    search_dirs.append(p)

    # Create temporary build environment
    temp_dir = tempfile.mkdtemp(prefix="latex2pptx_build_")
    try:
        # Copy source file
        staged_tex = os.path.join(temp_dir, tex_filename)
        shutil.copy2(tex_path, staged_tex)

        # Check for missing images
        missing = find_missing_images(tex_content, search_dirs)
        if missing:
            if not allow_placeholders:
                raise CompileError(
                    f"Missing required image files: {', '.join(sorted(missing))}. "
                    "Use --placeholder-images to generate visual placeholders."
                )
            if verbose:
                print(f"[latex2pptx] Generating visual placeholders for missing assets: {', '.join(sorted(missing))}")
            for m in missing:
                create_placeholder_image(m, temp_dir)

        # Copy existing assets from search directories
        for sdir in search_dirs:
            rel_subdir = os.path.relpath(sdir, tex_dir)
            target_sub = temp_dir if rel_subdir == "." else os.path.join(temp_dir, rel_subdir)
            os.makedirs(target_sub, exist_ok=True)
            for item in os.listdir(sdir):
                src_item = os.path.join(sdir, item)
                dst_item = os.path.join(target_sub, item)
                dst_root = os.path.join(temp_dir, item)
                if os.path.isfile(src_item):
                    if not os.path.exists(dst_item):
                        shutil.copy2(src_item, dst_item)
                    # Also copy to root of temp_dir if not exists
                    if not os.path.exists(dst_root):
                        shutil.copy2(src_item, dst_root)
                    clean_name = re.sub(r"\s*\(\d+\)", "", item)
                    clean_dst = os.path.join(temp_dir, clean_name)
                    if not os.path.exists(clean_dst):
                        shutil.copy2(src_item, clean_dst)
                elif os.path.isdir(src_item) and item in ("img", "images", "figures"):
                    if not os.path.exists(os.path.join(temp_dir, item)):
                        shutil.copytree(src_item, os.path.join(temp_dir, item))

        # Run compilation passes
        cmd = [engine, "-interaction=nonstopmode", tex_filename]
        last_stdout = ""
        for pass_num in range(1, passes + 1):
            if verbose:
                print(f"[latex2pptx] Running {engine} (pass {pass_num}/{passes})...")
            res = subprocess.run(
                cmd,
                cwd=temp_dir,
                capture_output=True,
                text=True,
                errors="replace",
            )
            last_stdout = res.stdout
            if res.returncode != 0 and pass_num == passes:
                error_lines = [
                    line for line in last_stdout.splitlines()
                    if line.startswith("!") or "Error:" in line or "Fatal error" in line
                ]
                msg = "\n".join(error_lines[:10]) if error_lines else last_stdout[-1500:]
                raise CompileError(f"{engine} compilation failed (code {res.returncode}):\n{msg}")

        temp_pdf = os.path.join(temp_dir, f"{tex_stem}.pdf")
        if not os.path.isfile(temp_pdf):
            raise CompileError(f"Compilation finished but output PDF was not found at {temp_pdf}")

        final_pdf = output_pdf_path or os.path.join(tex_dir, f"{tex_stem}.pdf")
        shutil.copy2(temp_pdf, final_pdf)
        return final_pdf
    finally:
        if not keep_temp:
            shutil.rmtree(temp_dir, ignore_errors=True)
