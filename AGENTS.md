# Repository Guidelines

## Project Structure & Module Organization

`pptx2img/` contains the Python package: `core.py` handles PowerPoint COM export, PDF rendering, and batch conversion; `cli.py` defines command-line behavior; `__init__.py` exposes the public API; and `__main__.py` enables module execution. `tests/test_conversion.py` contains unit tests, while `tests/live_powerpoint.py` provides an optional integration check. Packaging and dependencies live in `pyproject.toml`; usage examples live in `README.md`. Tests generate synthetic documents in temporary directories; there is no dedicated assets directory.

## Build, Test, and Development Commands

Use Python 3.10+ and run these commands from the repository root:

- `python -m pip install -e .` installs the package for development.
- `python -m unittest discover -s tests -v` runs the unit suite.
- `python tests/live_powerpoint.py` checks real PowerPoint export, rendering, and CLI behavior in a signed-in Windows session with PowerPoint installed.
- `python -m pip wheel . --no-deps -w dist` builds a distributable wheel.
- `python -m pptx2img "deck.pptx" -o "images" --dpi 150` converts a presentation locally.

## Coding Style & Naming Conventions

Follow the existing four-space Python indentation and PEP 8 naming: `snake_case` functions and variables, `PascalCase` classes, and leading underscores for internal helpers. Add type annotations and concise docstrings to public functions. Prefer `pathlib.Path` and context managers for file and resource handling. No formatter or linter is currently configured. Keep API and CLI changes synchronized with `README.md`.

## Testing Guidelines

Use standard-library `unittest`, `unittest.mock`, and `test_*` method names in `tests/test_*.py`. Cover changed behavior, particularly overwrite protection, cleanup, validation, and batch failures. No numeric coverage threshold is configured. Unit tests use synthetic PDFs and mocked COM; report live PowerPoint validation separately.

## Commit & Pull Request Guidelines

Git history is unavailable in this checkout. Use concise, imperative commit subjects describing the change. PRs should explain the problem, resulting behavior, test commands and results, and relevant issues. Include sample output or images when rendering behavior changes, and disclose whether live COM testing ran.

## Resource Lifecycle

Preserve balanced COM initialization and cleanup. Avoid concurrent PowerPoint conversions in one process. Close returned Pillow images, preserve one-based slide numbering, and retain existing-file protection unless overwrite is explicitly requested.
