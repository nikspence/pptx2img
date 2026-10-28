# pptx2img

Convert PowerPoint presentations (`.pptx` and `.ppt`) to Pillow images or PDF.
PowerPoint exports the presentation to PDF using **pywin32**; **PyMuPDF** then
renders PDF pages into materialized RGB `PIL.Image.Image` objects. Default
resolution is **150 DPI**. No `comtypes` dependency is used.

## Requirements and installation

- Windows, Python 3.10+, and installed Microsoft PowerPoint.
- Run under a signed-in Windows user with working PowerPoint automation.
- Python dependencies: pywin32, PyMuPDF, Pillow (installed by pip).

From this directory:

```powershell
python -m pip install .
# For development instead:
python -m pip install -e .
```

PowerPoint must be able to open the source without password or interactive
prompts. This library does not implement an Office-free or server rendering
engine. PowerPoint COM automation is performed synchronously on the calling
thread; avoid concurrent PowerPoint conversions in one process.

## Python API

```python
from pptx2img import pil_img, to_images, save_images, save_pdf, convert

# One slide as a Pillow image; slides are numbered from 1.
image = pil_img("deck.pptx")                  # First slide, 150 DPI
image.save("first-slide.png")
image.close()

with pil_img("deck.pptx", slide=2, dpi=300) as image:
    image.save("second-slide.jpg", quality=95)

# All slides in PDF page order, held in memory.
images = to_images("deck.pptx", dpi=150)
try:
    for number, image in enumerate(images, 1):
        image.save(f"slide-{number}.webp", quality=90)
finally:
    for image in images:
        image.close()

# Save one image per slide, rendering one at a time to limit memory use.
paths = save_images("deck.pptx", "images", dpi=200, extension="jpg", quality=95)
# images/deck_slide_001.jpg, images/deck_slide_002.jpg, ...

# Save only the PDF. The default destination is deck.pdf next to the input.
pdf_path = save_pdf("deck.pptx")
pdf_path = save_pdf("deck.pptx", "exports/deck.pdf", overwrite=True)

# Single-file or nonrecursive folder conversion, following pptxtopdf's API.
result = convert("presentations", "images", dpi=150, extension="png")
for source, outputs in result.converted.items():
    print(source, outputs)
for source, error in result.failed.items():
    print(source, error)

result = convert("presentations", "pdfs", pdf_only=True)
```

| Function | Result |
| --- | --- |
| `pil_img(input_path, slide=1, dpi=150)` | One Pillow image |
| `to_images(input_path, dpi=150)` | List of Pillow images |
| `save_images(input_path, output_dir=None, *, dpi=150, extension="png", overwrite=False, **save_options)` | List of absolute image paths |
| `save_pdf(input_path, output_path=None, *, overwrite=False)` | Absolute PDF path |
| `convert(input_path, output_folder_path=None, *, dpi=150, extension="png", pdf_only=False, overwrite=False, **save_options)` | `ConversionResult` with `converted` and `failed` dictionaries |

`dpi` must be a positive integer. Higher DPI increases pixel dimensions and
memory use. `slide` refers to a one-based page of PowerPoint's exported PDF;
only pages included in that PDF are available. Image functions export a fresh
temporary PDF and delete it before returning. They do not write and reread
temporary image files. Call `save_pdf` separately to retain a PDF.

Image formats are selected by a Pillow-supported save extension, such as
`png`, `jpg`, `jpeg`, `bmp`, `tif`, `tiff`, `webp`, or `gif`. Available writers
depend on your Pillow build; some formats require writer-specific options.
Extra `save_images`/`convert` keywords pass through to `Image.save`. DPI metadata
is supplied, but whether it is retained depends on the image format. Each slide
is saved separately, including TIFF and GIF. Returned Pillow images can also be
saved directly with any compatible `Image.save` options.

Default output directories are the source file's parent or the input folder.
Output folders are created automatically. Existing files raise `FileExistsError`
unless `overwrite=True`. Direct conversion functions raise errors; `convert`
records per-presentation errors and continues with the next file. A batch with
no PowerPoint files returns empty dictionaries. `~$` Office lock files are
ignored. Presentations with the same stem share output names, so existing-file
protection applies. A writer/I/O failure can leave earlier slide images behind.

## CLI

After installation, use `pptx2img` or `python -m pptx2img`:

```powershell
pptx2img "deck.pptx"
pptx2img "deck.pptx" -o "images" --dpi 300 --format jpg
pptx2img --input_dir "presentations" --output_dir "images"
pptx2img --input_file "deck.pptx" --output_dir "pdfs" --pdf-only
python -m pptx2img "presentations" -o "images" --format webp --overwrite
```

Both underscore and hyphen forms of `--input_file`, `--input_dir`, and
`--output_dir` are supported. Positional input cannot be combined with an input
flag. The CLI prints output paths and a success/failure summary. Exit status is
0 for a successful batch (including an empty folder), 1 for conversion failures,
and 2 for argument errors. Without `--pdf-only`, it saves images and discards
the intermediate PDF.

## Tests

```powershell
python -m unittest discover -s tests -v
```

The unit tests exercise real PyMuPDF/Pillow rendering with synthetic PDFs and
mock PowerPoint automation to test cleanup and failure handling. The optional
live check creates a two-slide presentation with installed PowerPoint:

```powershell
python tests/live_powerpoint.py
```

Validation in the development environment: the unit suite covers real PDF
rendering and mock COM lifecycle handling. The live check could not start
PowerPoint because COM returned `-2147023584` ("A specified logon session does
not exist. It may already have been terminated."). End-to-end PowerPoint
export must therefore be verified in an interactive Windows session using
the command above.

## Attribution and licensing

The file/folder workflow and PDF export behavior are based on
[hlviones/pptxtopdf](https://github.com/hlviones/pptxtopdf), licensed under MIT;
its copyright notice is retained in `LICENSE`. This implementation replaces
comtypes with pywin32 and adds image rendering. Rendering follows the
[PyMuPDF DPI API](https://pymupdf.readthedocs.io/en/latest/recipes-images.html#how-to-increase-image-resolution).
This project's MIT license does not replace dependency licenses; PyMuPDF is
available under AGPL or a commercial license.
