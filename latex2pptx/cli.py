"""
Command line interface for latex2pptx.
"""

import argparse
import os
import sys
import time

from . import __version__
from .compiler import compile_latex, CompileError
from .converter import convert_pdf_to_pptx


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="latex2pptx",
        description="Convert LaTeX Beamer presentations into faithful, editable PowerPoint (.pptx) decks.",
    )
    parser.add_argument(
        "input",
        help="Input LaTeX (.tex) or compiled PDF (.pdf) file.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output PowerPoint (.pptx) file path (default: <input_stem>.pptx).",
    )
    parser.add_argument(
        "-e",
        "--engine",
        default="pdflatex",
        choices=["pdflatex", "xelatex", "lualatex"],
        help="LaTeX compiler engine to use (default: pdflatex).",
    )
    parser.add_argument(
        "-p",
        "--passes",
        type=int,
        default=2,
        help="Number of LaTeX compiler passes (default: 2 for TikZ / remember picture).",
    )
    parser.add_argument(
        "-I",
        "--image-dir",
        action="append",
        dest="image_dirs",
        default=[],
        help="Additional directory to search for referenced images (can be used multiple times).",
    )
    parser.add_argument(
        "--placeholder-images",
        dest="allow_placeholders",
        action="store_true",
        default=True,
        help="Automatically generate visual placeholder images for any missing \\includegraphics assets (default: enabled).",
    )
    parser.add_argument(
        "--no-placeholder-images",
        dest="allow_placeholders",
        action="store_false",
        help="Fail compilation if any referenced images are missing.",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep the temporary compilation build directory for debugging.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose output.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    return parser


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        return 1

    t0 = time.time()
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".tex":
        print(f"[latex2pptx] Compiling LaTeX source: {os.path.basename(input_path)}...")
        try:
            pdf_path = compile_latex(
                tex_path=input_path,
                engine=args.engine,
                passes=args.passes,
                image_dirs=args.image_dirs,
                allow_placeholders=args.allow_placeholders,
                keep_temp=args.keep_temp,
                verbose=args.verbose,
            )
        except CompileError as err:
            print(f"Error during LaTeX compilation:\n{err}", file=sys.stderr)
            return 2
        except Exception as err:
            print(f"Unexpected error during LaTeX compilation: {err}", file=sys.stderr)
            return 2
    elif ext == ".pdf":
        pdf_path = input_path
    else:
        print(f"Error: expected a .tex or .pdf file, got '{ext}'", file=sys.stderr)
        return 1

    output_pptx = args.output
    if not output_pptx:
        base = os.path.splitext(input_path)[0]
        output_pptx = f"{base}.pptx"
    output_pptx = os.path.abspath(output_pptx)

    print(f"[latex2pptx] Converting PDF to PowerPoint: {os.path.basename(output_pptx)}...")
    try:
        convert_pdf_to_pptx(
            pdf_path=pdf_path,
            output_pptx_path=output_pptx,
            verbose=args.verbose,
        )
    except Exception as err:
        print(f"Error during PPTX conversion: {err}", file=sys.stderr)
        return 3

    elapsed = time.time() - t0
    print(f"[latex2pptx] Successfully created {output_pptx} in {elapsed:.1f}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
