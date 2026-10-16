# latex2pptx

<div align="center">

[![CI](https://github.com/dcuevasa/latex2pptx/actions/workflows/ci.yml/badge.svg)](https://github.com/dcuevasa/latex2pptx/actions/workflows/ci.yml)
[![GitHub Release](https://img.shields.io/github/v/release/dcuevasa/latex2pptx?color=blue&logo=github)](https://github.com/dcuevasa/latex2pptx/releases)
[![Python Version](https://img.shields.io/badge/python-3.9%20%7C%203.10%20%7C%203.11%20%7C%203.12-blue?logo=python)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

**Convert complex LaTeX Beamer presentations into faithful, fully editable PowerPoint (`.pptx`) slides.**

[Key Features](#key-features) • [Installation](#installation) • [Quick Start](#quick-start) • [CLI Reference](#cli-reference) • [Architecture](#architecture) • [License](#license)

</div>

---

## Overview

**`latex2pptx`** is a robust command-line tool engineered to bridge the gap between LaTeX Beamer and Microsoft PowerPoint. Unlike tools that simply convert PDF pages into static slide images, `latex2pptx` extracts, reconstructs, and generates **native, fully editable PowerPoint objects**:

- **Native Text Boxes** with matching fonts, colors, baselines, and kerning.
- **Native AutoShapes** (rectangles, rounded rectangles with accurate corner radii, ovals/circles).
- **Native Vector Curves** via DrawingML `<a:custGeom>` custom cubic Bézier shapes.
- **Overlay Steps** (`\only<...>`, `\pause`, etc.) mapped 1-to-1 to successive PowerPoint slides.
- **High-DPI Isolated Graphics** for icons (e.g. `fontawesome5`) and complex math symbols with tight alpha bounds.
- **TikZ Soft Clipping** supporting circular avatars, rounded-corner image crops, and arbitrary clipping paths.

Tested and validated on complex, production-grade Beamer slides featuring coordinate mathematics, custom macro loops (`\usebg{1}`..`\usebg{11}`), multi-column cards, and mixed typography.

---

## Key Features

### 🎨 Native & Editable PowerPoint Elements
- Text boxes are fully selectable, editable, and copyable.
- Visual vector geometry is translated directly into native PowerPoint `AUTO_SHAPE` elements or `<a:custGeom>` vector paths. Default PowerPoint theme drop shadows are automatically stripped so slides remain clean and flat.

### 📐 Calibrated Baseline & Sub-Pixel Font Alignment
- LaTeX `LMSans` fonts have narrower character advances and smaller x-heights than standard system sans-serif fonts (`Liberation Sans` / `Arial`). `latex2pptx` measures each line with FreeType at the glyph level and applies calibrated font scaling and baseline positioning (`top = origin_y - size`), completely eliminating text overflow and vertical drift.

### ⏱️ Beamer Overlay & Animation Preservation
- Slides with `\only<1>`, `\only<2->`, or `\pause` are faithfully expanded into separate sequential slides, preserving the exact pacing and transitions of the compiled presentation.

### 🖼️ Smart Image & Clipping Path Handling
- High-resolution images are extracted with their underlying alpha masks (`SMask`).
- TikZ clipping masks (`\clip circle`, `\clip rounded rectangle`) are executed as anti-aliased transparency channels, ensuring circular profile photos and rounded badges appear smoothly in PowerPoint.
- Automatically handles filename variants and searches configured `\graphicspath` directories.

### 🛡️ Missing Asset Recovery & Placeholders
- Missing background images or logos referenced in LaTeX loops are automatically substituted with clean visual vector placeholders during intermediate compilation, preventing compilation halts.

### 📝 Speaker Notes from Markdown
- Attach speaker notes to each slide directly from a Markdown (`.md`) file using headings (`# Slide 1: Welcome`, `## Slide 2-4: Overlays`, `### Diapositiva 5`), HTML comment directives (`<!-- slide: 1 -->`), or horizontal rules (`---`). See [Speaker Notes Guide](docs/SPEAKER_NOTES.md).

---

## Installation

### Prerequisites
- Python 3.9 or higher
- A standard TeX Live distribution (`pdflatex`, `xelatex`, or `lualatex`):
  ```bash
  # Debian / Ubuntu
  sudo apt-get install -y texlive-latex-base texlive-latex-recommended texlive-latex-extra texlive-pictures fonts-liberation

  # macOS (Homebrew)
  brew install --cask mactex-no-gui
  ```

### Install with pip
```bash
git clone https://github.com/dcuevasa/latex2pptx.git
cd latex2pptx
pip install -e .
```

Verify your installation:
```bash
latex2pptx --version
```

---

## Quick Start

### 1. Convert a `.tex` Source File
Directly compile and convert a LaTeX Beamer source presentation:
```bash
latex2pptx presentation.tex
```
This produces `presentation.pptx` in the same directory.

### 2. Specify Output File
```bash
latex2pptx presentation.tex -o output_deck.pptx
```

### 3. Convert an Existing Compiled PDF
If you already have a compiled Beamer PDF:
```bash
latex2pptx presentation.pdf -o output_deck.pptx
```

### 4. Provide Additional Image Directories
If your LaTeX presentation loads assets from custom folders or subdirectories:
```bash
latex2pptx presentation.tex -I ./assets -I ./img
```

### 5. Attach Speaker Notes from Markdown
Attach speaking prompts, talking points, or bullet lists directly to each slide:
```bash
latex2pptx presentation.tex -n notes.md -o output_deck.pptx
```
*(Tip: If you name your file `<input_stem>.notes.md`, e.g. `presentation.notes.md`, `latex2pptx` automatically attaches it!)*

---

## Speaker Notes (.md)

You can write notes in Markdown using simple headings or delimiters. Full details and examples are documented in [docs/SPEAKER_NOTES.md](docs/SPEAKER_NOTES.md).

```markdown
# Slide 1: Welcome & Introductions
Welcome everyone to the session.
- Introduce speakers
- Mention session agenda

## Slide 2-4: Core Concepts
Notes that apply across multiple Beamer overlay steps.
- Key takeaways
- Audience check-in

### Diapositiva 5
Spanish keywords like Diapositiva / Diapo / Página are fully supported!

---
Or simply use horizontal rules to separate slides sequentially.
```

---

## CLI Reference

```text
usage: latex2pptx [-h] [-o OUTPUT] [-e {pdflatex,xelatex,lualatex}] [-p PASSES]
                  [-I IMAGE_DIRS] [--placeholder-images] [--no-placeholder-images]
                  [-n NOTES_FILE] [--keep-temp] [-v] [--version]
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
  -n, --notes           Path to a Markdown (.md) file containing speaker notes for slides.
  --keep-temp           Keep the temporary compilation build directory for debugging.
  -v, --verbose         Enable verbose output.
  --version             show program's version number and exit
```

---

## Architecture

`latex2pptx` operates through a multi-stage modular conversion pipeline:

```
┌─────────────────┐     ┌──────────────────────┐     ┌────────────────────────┐
│  clase1.tex     │ ──> │ compiler.py          │ ──> │ Multi-pass TeX Build   │
│  (Beamer+TikZ)  │     │ (Asset Recovery & Stg)│    │ (pdflatex / xelatex)   │
└─────────────────┘     └──────────────────────┘     └───────────┬────────────┘
                                                                 │ (PDF)
                                                                 ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│ converter.py (Global Z-Order Stream Parser)                                 │
│                                                                             │
│   ├── shapes.py: Vector geometry -> PPTX Native Shapes & <a:custGeom>      │
│   ├── text.py: FreeType metric scaling & calibrated baseline text boxes    │
│   ├── images.py: Raster extraction, SMask transparency & TikZ clip masks    │
│   ├── icons.py: FontAwesome & symbol isolation at 8x supersampling          │
│   └── notes.py: Markdown notes parser (headings, ranges, delimiters)        │
└──────────────────────────────────────┬──────────────────────────────────────┘
                                       │
                                       ▼
                             ┌───────────────────┐
                             │    output.pptx    │
                             │ (Native Editable) │
                             └───────────────────┘
```

1. **`compiler.py`**: Runs isolated multi-pass compilation. Analyzes `\graphicspath` and LaTeX macros for image dependencies and auto-generates geometry-matched placeholders for any missing assets.
2. **`converter.py`**: Interrogates the PDF drawing and clipping streams in strict z-order, synchronizing visual layers.
3. **`shapes.py`**: Classifies PDF vector paths (`rect`, `roundrect`, `circle`, `cubic bezier`), builds PowerPoint DrawingML geometry, and strips legacy PowerPoint drop shadows.
4. **`text.py`**: Performs character-level glyph kerning, normalizes whitespace and punctuation, detects alignment (left, center, right), and scales font sizes using FreeType metrics to prevent line overflows while locking baseline vertical coordinates.
5. **`images.py`**: Extracts embedded raster images, handles alpha transparency (`SMask`), and renders clipped paths with sub-pixel anti-aliasing.
6. **`icons.py`**: Identifies symbol/icon glyphs (FontAwesome5, Computer Modern symbols), supersamples them at 8x, and tightly crops alpha bounds.
7. **`notes.py`**: Parses structured Markdown notes files into 1-indexed slide mapping and populates PowerPoint speaker notes frames.

---

## Running Tests

The test suite covers CLI argument parsing, asset discovery, placeholder generation, font metric scaling, and end-to-end Beamer compilation:

```bash
python3 -m unittest discover -s tests -v
```

---

## Contributing

Contributions, bug reports, and feature suggestions are welcome! Feel free to:
1. Fork the repository.
2. Create a feature branch (`git checkout -b feature/my-feature`).
3. Commit your changes (`git commit -m "feat: Add support for ..."`).
4. Push to the branch (`git push origin feature/my-feature`).
5. Open a Pull Request.

---

## License

This project is licensed under the [MIT License](LICENSE).