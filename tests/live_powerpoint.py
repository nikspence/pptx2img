"""Optional real PowerPoint round trip; creates only temporary synthetic data."""

from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory

# Allow direct execution from a source checkout without installation.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pymupdf
from PIL import Image
from pptx2img import pil_img, save_images, save_pdf, to_images
from pptx2img.core import _powerpoint


def main():
    with TemporaryDirectory(prefix="pptx2img-live-") as folder:
        root = Path(folder)
        source = root / "test deck.pptx"
        with _powerpoint() as app:
            deck = app.Presentations.Add(False)
            try:
                deck.PageSetup.SlideWidth = 720
                deck.PageSetup.SlideHeight = 405
                for number in (1, 2):
                    slide = deck.Slides.Add(number, 12)  # ppLayoutBlank
                    shape = slide.Shapes.AddTextbox(1, 50, 50, 620, 100)
                    shape.TextFrame.TextRange.Text = f"pptx2img live test - slide {number}"
                    shape.TextFrame.TextRange.Font.Size = 32
                    shape = None
                    slide = None
                deck.SaveAs(str(source), 24)
            finally:
                deck.Close()
                deck = None
        pdf = save_pdf(source)
        with pymupdf.open(pdf) as document:
            assert len(document) == 2
            assert "slide 2" in document[1].get_text()
        with pil_img(source) as image:
            assert image.size == (1500, 844), image.size
            image.save(root / "first.png")
        images = to_images(source, dpi=72)
        try:
            assert len(images) == 2
            assert images[1].size == (720, 405)
        finally:
            for image in images:
                image.close()
        paths = save_images(source, root / "jpg", extension="jpg", dpi=72, quality=90)
        for path in paths:
            with Image.open(path) as image:
                assert image.size == (720, 405)
        for flags in ([str(source), "-o", str(root / "cli"), "--dpi", "72"],
                      ["--input_dir", str(root), "--output_dir", str(root / "pdfs"), "--pdf-only"]):
            subprocess.run([sys.executable, "-m", "pptx2img", *flags], check=True)
        print("Live PowerPoint -> PDF -> Pillow/API/CLI checks passed.")


if __name__ == "__main__":
    main()
