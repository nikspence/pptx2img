from contextlib import contextmanager
import io
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import MagicMock, patch

import pymupdf
from PIL import Image

from pptx2img import convert, pil_img, save_images, save_pdf, to_images
from pptx2img import core
from pptx2img.cli import main


def write_pdf(input_path, output_path, **kwargs):
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with pymupdf.open() as document:
        for color in ((1, 0, 0), (0, 0, 1)):
            page = document.new_page(width=144, height=72)
            page.draw_rect(page.rect, color=color, fill=color)
        document.save(output_path)
    return Path(output_path)


class RenderingTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "deck.pptx"
        self.source.touch()
        self.exporter = patch.object(core, "save_pdf", side_effect=write_pdf).start()
        self.addCleanup(patch.stopall)

    def test_images_own_pixels_after_pdf_cleanup_and_dpi_scales(self):
        images = to_images(self.source)
        try:
            self.assertEqual(len(images), 2)
            self.assertEqual(images[0].size, (300, 150))
            self.assertEqual(images[0].getpixel((100, 50)), (255, 0, 0))
            self.assertEqual(images[1].getpixel((100, 50)), (0, 0, 255))
            self.assertFalse(Path(self.exporter.call_args.args[1]).exists())
        finally:
            for image in images:
                image.close()
        with pil_img(self.source, slide=2, dpi=300) as image:
            self.assertEqual(image.size, (600, 300))
            self.assertEqual(image.info["dpi"], (300, 300))

    def test_file_writers_and_overwrite(self):
        for extension in ("png", "jpg", "webp", "tiff", "gif", "bmp", "xbm"):
            with self.subTest(extension=extension):
                folder = self.root / extension
                paths = save_images(self.source, folder, extension=extension)
                self.assertEqual(len(paths), 2)
                with Image.open(paths[0]) as image:
                    self.assertEqual(image.size, (300, 150))
                original = paths[0].read_bytes()
                with self.assertRaises(FileExistsError):
                    save_images(self.source, folder, extension=extension)
                self.assertEqual(paths[0].read_bytes(), original)
                save_images(self.source, folder, extension=extension, overwrite=True)

    def test_validation(self):
        for dpi in (0, -1, 1.5, True):
            with self.assertRaises(ValueError):
                pil_img(self.source, dpi=dpi)
        for slide in (0, -1, 3, True):
            with self.assertRaises(ValueError):
                pil_img(self.source, slide=slide)
        with self.assertRaises(ValueError):
            save_images(self.source, extension="unknown")
        with self.assertRaises(FileNotFoundError):
            save_images(self.root / "missing.pptx")

    def test_batch_continues_and_ignores_lock_files(self):
        (self.root / "bad.pptx").touch()
        (self.root / "~$deck.pptx").touch()
        (self.root / "ignore.txt").touch()
        def export(source, target, **kwargs):
            if Path(source).stem == "bad":
                raise RuntimeError("bad presentation")
            return write_pdf(source, target)
        self.exporter.side_effect = export
        result = convert(self.root, self.root / "out")
        self.assertEqual(list(result.converted), [self.source])
        self.assertEqual(list(result.failed), [self.root / "bad.pptx"])
        with patch("sys.stdout", new_callable=io.StringIO), patch("sys.stderr", new_callable=io.StringIO):
            self.assertEqual(main([str(self.root), "-o", str(self.root / "cli")]), 1)

    def test_empty_directory_and_pdf_only_batch(self):
        folder = self.root / "empty"
        folder.mkdir()
        self.assertEqual(convert(folder).converted, {})
        result = convert(self.source, self.root / "pdf", pdf_only=True)
        self.assertTrue(result.converted[self.source][0].is_file())

    def test_cli_success_and_pillow_save_options(self):
        with patch("sys.stdout", new_callable=io.StringIO) as output:
            self.assertEqual(main(["--input_file", str(self.source), "--output_dir",
                                   str(self.root / "cli"), "--dpi", "72"]), 0)
        self.assertIn("1 files converted successfully, 0 files failed", output.getvalue())
        paths = save_images(self.source, self.root / "quality", extension="jpg", quality=95)
        with Image.open(paths[0]) as image:
            self.assertEqual(image.info["dpi"], (150, 150))


class ExportTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.source = Path(self.temp.name) / "deck.pptx"
        self.source.touch()
        self.app = MagicMock()
        self.presentation = self.app.Presentations.Open.return_value
        self.presentation.SaveAs.side_effect = lambda target, fmt: Path(target).write_bytes(b"%PDF-test")
        @contextmanager
        def powerpoint():
            yield self.app
        patcher = patch.object(core, "_powerpoint", powerpoint)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_export_and_close(self):
        target = save_pdf(self.source)
        self.assertEqual(target.read_bytes(), b"%PDF-test")
        self.app.Presentations.Open.assert_called_once_with(str(self.source), True, False, False)
        self.presentation.Close.assert_called_once()
        with self.assertRaises(FileExistsError):
            save_pdf(self.source)
        save_pdf(self.source, overwrite=True)

    def test_failure_closes_deck_and_preserves_existing_pdf(self):
        target = self.source.with_suffix(".pdf")
        target.write_bytes(b"original")
        self.presentation.SaveAs.side_effect = RuntimeError("export failed")
        with self.assertRaisesRegex(RuntimeError, "export failed"):
            save_pdf(self.source, overwrite=True)
        self.presentation.Close.assert_called_once()
        self.assertEqual(target.read_bytes(), b"original")
        self.assertEqual(list(target.parent.glob("pptx2img-*")), [])


@unittest.skipUnless(core.sys.platform == "win32", "Windows COM lifecycle")
class ComLifecycleTests(unittest.TestCase):
    def test_com_is_balanced_on_export_failure_and_start_failure(self):
        import pythoncom
        import win32com.client
        for start_failure in (False, True):
            with self.subTest(start_failure=start_failure):
                app = MagicMock()
                with patch.object(pythoncom, "CoInitialize") as initialize, \
                     patch.object(pythoncom, "CoUninitialize") as uninitialize, \
                     patch.object(win32com.client, "DispatchEx", return_value=app,
                                  side_effect=RuntimeError("startup") if start_failure else None):
                    with self.assertRaises(RuntimeError):
                        with core._powerpoint():
                            raise RuntimeError("export")
                    initialize.assert_called_once()
                    uninitialize.assert_called_once()
                    self.assertEqual(app.Quit.call_count, 0 if start_failure else 1)


if __name__ == "__main__":
    unittest.main()
