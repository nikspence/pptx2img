"""Conversion API. Slide numbers are one-based PDF page numbers."""

from __future__ import annotations

from contextlib import contextmanager, suppress
from dataclasses import dataclass, field
from pathlib import Path
import os
import sys
from tempfile import TemporaryDirectory
from typing import Iterator

import pymupdf
from PIL import Image

PathLike = str | os.PathLike[str]


def _source(input_path: PathLike) -> Path:
    source = Path(input_path).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"PowerPoint file not found: {source}")
    if source.suffix.lower() not in {".ppt", ".pptx"}:
        raise ValueError("Input must be a .pptx or .ppt file")
    return source


def _dpi(dpi: int) -> int:
    if isinstance(dpi, bool) or not isinstance(dpi, int) or dpi <= 0:
        raise ValueError("dpi must be a positive integer")
    return dpi


@contextmanager
def _powerpoint() -> Iterator[object]:
    if sys.platform != "win32":
        raise OSError("PowerPoint conversion requires Windows and Microsoft PowerPoint")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ImportError("Install pywin32 to automate Microsoft PowerPoint") from exc

    # Balance initialization on this calling thread, including when COM fails.
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.DisplayAlerts = 1  # ppAlertsNone
        yield app
    finally:
        if app is not None:
            with suppress(Exception):
                app.Quit()
        app = None
        pythoncom.CoUninitialize()


def save_pdf(
    input_path: PathLike,
    output_path: PathLike | None = None,
    *,
    overwrite: bool = False,
) -> Path:
    """Save a presentation as PDF; return its absolute path.

    output_path defaults to the source path with a .pdf extension. Existing
    output files raise FileExistsError unless overwrite=True. PowerPoint opens
    the source read-only without a presentation window. Export is staged in a
    temporary directory so a failed export does not damage an existing PDF.
    """
    source = _source(input_path)
    target = Path(output_path).resolve() if output_path is not None else source.with_suffix(".pdf")
    if target.suffix.lower() != ".pdf":
        raise ValueError("PDF output_path must end in .pdf")
    if target.exists() and not overwrite:
        raise FileExistsError(f"Output already exists: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="pptx2img-", dir=target.parent) as folder:
        staged = Path(folder) / "presentation.pdf"
        with _powerpoint() as app:
            presentation = None
            try:
                # Positional arguments also work with late-bound pywin32 COM.
                # Open(FileName, ReadOnly, Untitled, WithWindow)
                presentation = app.Presentations.Open(str(source), True, False, False)
                presentation.SaveAs(str(staged), 32)  # ppSaveAsPDF
            finally:
                if presentation is not None:
                    with suppress(Exception):
                        presentation.Close()
                presentation = None
        if not staged.is_file() or staged.stat().st_size == 0:
            raise RuntimeError(f"PowerPoint did not produce a PDF for {source}")
        if overwrite:
            staged.replace(target)
        else:
            # On Windows rename refuses an existing destination, even if it
            # appeared after the initial check.
            staged.rename(target)
    return target


@contextmanager
def _pdf(input_path: PathLike) -> Iterator[pymupdf.Document]:
    with TemporaryDirectory(prefix="pptx2img-") as folder:
        pdf_path = save_pdf(input_path, Path(folder) / "presentation.pdf")
        with pymupdf.open(pdf_path) as document:
            yield document


def _render(page: pymupdf.Page, dpi: int) -> Image.Image:
    pixmap = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csRGB, alpha=False)
    # frombytes owns its pixels: no temporary image file or live PDF is needed.
    image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    image.info["dpi"] = (dpi, dpi)
    return image


def pil_img(input_path: PathLike, slide: int = 1, dpi: int = 150) -> Image.Image:
    """Return one slide as a materialized RGB Pillow image (one-based slide)."""
    _dpi(dpi)
    if isinstance(slide, bool) or not isinstance(slide, int) or slide < 1:
        raise ValueError("slide must be a positive, one-based integer")
    with _pdf(input_path) as document:
        if slide > len(document):
            raise ValueError(f"slide {slide} is out of range; PDF has {len(document)} pages")
        return _render(document[slide - 1], dpi)


def to_images(input_path: PathLike, dpi: int = 150) -> list[Image.Image]:
    """Return all exported slides as RGB Pillow images in presentation order.

    The intermediate PDF is deleted before return. The caller owns the images.
    For large presentations prefer save_images(), which renders one at a time.
    """
    _dpi(dpi)
    images: list[Image.Image] = []
    try:
        with _pdf(input_path) as document:
            for page in document:
                images.append(_render(page, dpi))
        return images
    except BaseException:
        for image in images:
            image.close()
        raise


def _image_format(extension: str) -> tuple[str, str]:
    extension = extension.lower().lstrip(".")
    Image.init()
    image_format = Image.registered_extensions().get("." + extension)
    if image_format is None or image_format not in Image.SAVE:
        raise ValueError(f"Pillow cannot save image extension: {extension!r}")
    return extension, image_format


def save_images(
    input_path: PathLike,
    output_dir: PathLike | None = None,
    *,
    dpi: int = 150,
    extension: str = "png",
    overwrite: bool = False,
    **save_options: object,
) -> list[Path]:
    """Save one image per slide and return absolute paths.

    Names are <presentation>_slide_001.<extension>, etc. extension selects a
    Pillow writer (png, jpg, webp, tiff, ...). Additional keyword arguments go
    to Image.save(), e.g. quality=95. RGB images are converted to palette or
    grayscale for Pillow writers that require it. All targets are checked
    before rendering; a later writer/I/O failure may leave earlier images.
    """
    source = _source(input_path)
    _dpi(dpi)
    extension, image_format = _image_format(extension)
    folder = Path(output_dir).resolve() if output_dir is not None else source.parent
    folder.mkdir(parents=True, exist_ok=True)
    with _pdf(source) as document:
        targets = [folder / f"{source.stem}_slide_{i:03d}.{extension}"
                   for i in range(1, len(document) + 1)]
        if not overwrite:
            for target in targets:
                if target.exists():
                    raise FileExistsError(f"Output already exists: {target}")
        for page, target in zip(document, targets):
            with _render(page, dpi) as image:
                # These writers do not accept RGB; most others do.
                mode = {"GIF": "P", "PALM": "P", "XBM": "1", "MSP": "1",
                        "WBMP": "1", "SPIDER": "F"}.get(image_format, "RGB")
                converted = image.convert(mode) if mode != image.mode else image
                try:
                    options = {"dpi": (dpi, dpi), **save_options}
                    # Exclusive create preserves the no-overwrite contract
                    # even if another writer creates a target after preflight.
                    with target.open("wb" if overwrite else "xb") as stream:
                        converted.save(stream, format=image_format, **options)
                finally:
                    if converted is not image:
                        converted.close()
    return targets


@dataclass
class ConversionResult:
    """Batch results keyed by absolute source paths; failures do not stop a batch."""

    converted: dict[Path, list[Path]] = field(default_factory=dict)
    failed: dict[Path, str] = field(default_factory=dict)


def convert(
    input_path: PathLike,
    output_folder_path: PathLike | None = None,
    *,
    dpi: int = 150,
    extension: str = "png",
    pdf_only: bool = False,
    overwrite: bool = False,
    **save_options: object,
) -> ConversionResult:
    """Convert one file or a nonrecursive directory of .ppt/.pptx files.

    This follows pptxtopdf's convert(input_path, output_folder_path) workflow,
    with image output by default. pdf_only=True saves PDFs instead. Temporary
    Office lock files (~$...) are ignored. Errors are returned in result.failed.
    """
    if not pdf_only:
        _dpi(dpi)
        _image_format(extension)
    source = Path(input_path).resolve()
    if source.is_dir():
        sources = sorted((p for p in source.iterdir() if p.is_file()
                          and p.suffix.lower() in {".ppt", ".pptx"}
                          and not p.name.startswith("~$")), key=lambda p: p.name.lower())
    else:
        sources = [_source(source)]
    folder = Path(output_folder_path).resolve() if output_folder_path is not None else (source if source.is_dir() else source.parent)
    result = ConversionResult()
    for presentation in sources:
        try:
            if pdf_only:
                paths = [save_pdf(presentation, folder / f"{presentation.stem}.pdf", overwrite=overwrite)]
            else:
                paths = save_images(presentation, folder, dpi=dpi, extension=extension,
                                    overwrite=overwrite, **save_options)
            result.converted[presentation] = paths
        except Exception as exc:
            result.failed[presentation] = str(exc)
    return result
