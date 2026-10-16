# latex2pptx

A reliable, faithful command-line tool to convert complex LaTeX Beamer presentations into equivalent, editable PowerPoint (`.pptx`) presentations.

Developed and tested against complex, modern Beamer slides featuring extensive TikZ drawings, coordinate mathematics, custom macros, FontAwesome5 icons, multi-line paragraph groupings, and overlay steps (`\only<...>`, `\pause`, etc.).

---

## Features

- **Editable Native PowerPoint Elements**: Converts text, rectangles, rounded rectangles (with matched corner radius), circles/ovals, and custom vector curves into native PowerPoint shapes (`AUTO_SHAPE`, `TEXT_BOX`, and DrawingML `<a:custGeom>` Freeforms).
- **Faithful Visual Layout**: Preserves exact 16:9 widescreen proportions, colors, fonts, baselines, and geometric alignment.
- **Sub-Pixel Icon & Symbol Isolation**: Isolates FontAwesome5 icons and math symbols and rasterizes them with high-DPI supersampling and tight alpha cropping, preventing font missing/corrupted glyph issues across different operating systems.
- **Overlay Step Preservation**: Accurately maps each Beamer overlay step (`\only<2->`, `\pause`, etc.) to its corresponding PowerPoint slide.
- **Smart Asset & Placeholder Management**: Detects referenced images (including macro loops like `bg#1`) and automatically generates clean visual placeholders if source images are missing, ensuring compilation never crashes.
- **Dual Input Modes**: Directly compiles `.tex` files with multi-pass TeX engines (`pdflatex`, `xelatex`, `lualatex`), or accepts already-compiled `.pdf` files.

---

## Installation

### Prerequisites
- Python 3.9+
- A TeX Live distribution (`pdflatex`, `texlive-latex-recommended`, `texlive-pictures`, `texlive-fonts-extra`, etc.)

### Install via pip
```bash
git clone https://github.com/dcuevasa/latex2pptx.git
cd latex2pptx
pip install -e .
```

---

## Usage

### Basic Usage
Convert a `.tex` presentation directly to `.pptx`:
```bash
latex2pptx clase1.tex
```
This produces `clase1.pptx` in the same directory.

### Specify Output Path
```bash
latex2pptx clase1.tex -o presentation.pptx
```

### Convert Existing Compiled PDF
If you already have a compiled Beamer PDF:
```bash
latex2pptx clase1.pdf -o presentation.pptx
```

### Additional Options
```bash
# Search additional directories for images
latex2pptx clase1.tex -I ./assets -I ./img

# Use XeLaTeX or LuaLaTeX
latex2pptx clase1.tex --engine xelatex

# Specify compiler passes (default: 2)
latex2pptx clase1.tex --passes 2

# Fail if any image is missing (disable auto-placeholders)
latex2pptx clase1.tex --no-placeholder-images

# Verbose progress output
latex2pptx clase1.tex -v
```

---

## CLI Reference

```text
usage: latex2pptx [-h] [-o OUTPUT] [-e {pdflatex,xelatex,lualatex}] [-p PASSES]
                  [-I IMAGE_DIRS] [--placeholder-images] [--no-placeholder-images]
                  [--keep-temp] [-v] [--version]
                  input

positional arguments:
  input                 Input LaTeX (.tex) or compiled PDF (.pdf) file.

options:
  -h, --help            show this help message and exit
  -o, --output OUTPUT   Output PowerPoint (.pptx) file path (default: <input_stem>.pptx).
  -e, --engine          LaTeX compiler engine to use (default: pdflatex).
  -p, --passes PASSES   Number of LaTeX compiler passes (default: 2 for TikZ / remember picture).
  -I, --image-dir       Additional directory to search for referenced images (can be repeated).
  --placeholder-images  Automatically generate visual placeholder images for missing assets (default: enabled).
  --no-placeholder-images
                        Fail compilation if any referenced images are missing.
  --keep-temp           Keep the temporary compilation build directory for debugging.
  -v, --verbose         Enable verbose output.
  --version             show program's version number and exit
```

---

## Architecture

1. **Compiler Pipeline (`compiler.py`)**:
   Runs multi-pass compilation inside an isolated staging directory. Scans LaTeX macros and `\graphicspath` for external image dependencies and generates labeled vector placeholders if any asset is missing.
2. **Layer Separation & Vector Classifier (`shapes.py`)**:
   Inspects PDF vector streams (`seqno`), classifying geometry into native PowerPoint `RECTANGLE`, `ROUNDED_RECTANGLE`, `OVAL`, or custom Bézier geometry (`<a:custGeom>`) while removing default theme drop shadows.
3. **Typography & Paragraph Merger (`text.py`)**:
   Extracts characters from raw glyph dictionaries with kerning preservation, maps TeX fonts to system font families, and merges lines into native text boxes with calibrated baseline offsets (`top = origin_y - size`) and baseline line-spacing.
4. **Symbol & Icon Isolation (`icons.py`)**:
   Filters PDF content streams to isolate icon glyphs, supersamples them at 8x with anti-aliasing, and tightly crops alpha bounds for pixel-faithful rendering.
5. **Image & Clip Handling (`images.py`)**:
   Extracts embedded raster images, preserving SMask transparency and executing TikZ clipping paths (`\clip circle`, `\clip rounded rectangle`) as anti-aliased alpha channels.

---

## Running Tests

Run the test suite with:
```bash
python3 -m unittest discover -s tests -v
```

---

## License
MIT License