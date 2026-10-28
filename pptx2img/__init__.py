"""PowerPoint -> PDF -> Pillow images (Windows + Microsoft PowerPoint)."""

from .core import ConversionResult, convert, pil_img, save_images, save_pdf, to_images

__version__ = "0.1.0"
__all__ = ["ConversionResult", "convert", "pil_img", "save_images", "save_pdf", "to_images"]
