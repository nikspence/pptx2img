"""Command-line interface, including pptxtopdf-compatible input flags."""

import argparse
from pathlib import Path
import sys

from .core import convert


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert PowerPoint to images or PDF (Windows + PowerPoint).")
    parser.add_argument("input", nargs="?", help="PowerPoint file or folder")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--input-file", "--input_file", type=Path, help="Input .pptx or .ppt file")
    group.add_argument("--input-dir", "--input_dir", type=Path, help="Input folder (nonrecursive)")
    parser.add_argument("-o", "--output-dir", "--output_dir", type=Path, help="Output folder (default: input folder)")
    parser.add_argument("--dpi", type=int, default=150, help="Image resolution (default: 150)")
    parser.add_argument("--format", default="png", dest="extension", help="Pillow output extension, e.g. png, jpg, webp, tiff")
    parser.add_argument("--pdf-only", action="store_true", help="Save PDFs instead of images")
    parser.add_argument("--overwrite", action="store_true", help="Replace existing output files")
    args = parser.parse_args(argv)
    if args.input and (args.input_file is not None or args.input_dir is not None):
        parser.error("Use either positional input or --input-file/--input-dir")
    source = args.input or args.input_file or args.input_dir
    if source is None:
        parser.error("Provide a PowerPoint file or folder")
    if args.input_file is not None and not args.input_file.is_file():
        parser.error("--input-file must name an existing file")
    if args.input_dir is not None and not args.input_dir.is_dir():
        parser.error("--input-dir must name an existing directory")
    try:
        result = convert(source, args.output_dir, dpi=args.dpi, extension=args.extension,
                         pdf_only=args.pdf_only, overwrite=args.overwrite)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    for outputs in result.converted.values():
        for path in outputs:
            print(path)
    for path, error in result.failed.items():
        print(f"Error converting {path}: {error}", file=sys.stderr)
    print(f"Conversion completed: {len(result.converted)} files converted successfully, "
          f"{len(result.failed)} files failed.")
    return 1 if result.failed else 0
