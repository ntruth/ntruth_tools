"""Utilities for accessing the built-in Excel template.

This module exposes helpers to decode the default template that ships with the
project.  The actual template bytes are stored as a Base64 string so that the
repository remains text-only and can be committed without binary diffs.
"""
from __future__ import annotations

import argparse
import base64
from io import BytesIO
from pathlib import Path
from typing import Optional

DEFAULT_TEMPLATE_FILENAME = "txt_to_excel_template.xlsx"

# Base64 representation of the built-in Excel template.  The template is a
# minimal workbook with a single worksheet and no extra styling.  It exists so
# that users always have a starting point that matches the behaviour described
# in the original requirements document.
DEFAULT_TEMPLATE_B64 = (
    "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"
)


def default_template_bytes() -> bytes:
    """Return the decoded bytes for the built-in template."""

    return base64.b64decode(DEFAULT_TEMPLATE_B64)


def default_template_stream() -> BytesIO:
    """Return a BytesIO handle for the built-in template."""

    return BytesIO(default_template_bytes())


def ensure_default_template_file(path: Optional[Path] = None) -> Path:
    """Ensure the default template exists on disk and return its path.

    The template is written lazily the first time this function is called so
    that users can still reference it by path (e.g. in the GUI file chooser)
    without shipping a binary file in the repository.
    """

    if path is None:
        path = Path(__file__).with_name(DEFAULT_TEMPLATE_FILENAME)

    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(default_template_bytes())

    return path


def export_default_template(destination: Path) -> Path:
    """Export the built-in template to an arbitrary location."""

    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(default_template_bytes())
    return destination


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Export the built-in Excel template")
    parser.add_argument(
        "destination",
        nargs="?",
        type=Path,
        default=Path.cwd() / DEFAULT_TEMPLATE_FILENAME,
        help="导出模板的目标路径，默认为当前目录",
    )
    return parser


def main(argv: Optional[list[str]] = None) -> None:
    parser = _build_parser()
    args = parser.parse_args(argv)
    path = export_default_template(args.destination)
    print(f"模板已导出到 {path}")


if __name__ == "__main__":  # pragma: no cover - 命令行工具
    main()
