#!/usr/bin/env python3
"""Utility to dump the current siaf_dashboard.py code for manual sharing.

This helper lets non-technical users obtain the complete source either
by writing it to a destination path or printing it to the terminal so it
can be copied from the console.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _resolve_source() -> Path:
    """Return the absolute path to siaf_dashboard.py, ensuring it exists."""

    repo_root = Path(__file__).resolve().parent
    source = repo_root / "siaf_dashboard.py"
    if not source.is_file():
        raise FileNotFoundError(
            "No se encontró siaf_dashboard.py en el directorio del proyecto."
        )
    return source


def dump_code(output: Path | None, to_stdout: bool) -> int:
    """Export the dashboard code either to a file or stdout.

    Parameters
    ----------
    output:
        Destination file path when ``to_stdout`` is ``False``.
    to_stdout:
        If ``True``, the file content is written to ``sys.stdout``.

    Returns
    -------
    int
        Zero on success, non-zero when an error occurs.
    """

    try:
        source = _resolve_source()
        code = source.read_text(encoding="utf-8")
    except Exception as exc:  # pragma: no cover - runtime safeguard
        print(f"Error al leer el archivo fuente: {exc}", file=sys.stderr)
        return 1

    if to_stdout:
        sys.stdout.write(code)
        return 0

    if output is None:
        print(
            "Debes indicar una ruta de destino cuando no usas --stdout",
            file=sys.stderr,
        )
        return 2

    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(code, encoding="utf-8")
    except Exception as exc:  # pragma: no cover - runtime safeguard
        print(f"No se pudo escribir el archivo destino: {exc}", file=sys.stderr)
        return 3

    print(f"Código copiado correctamente en: {output}")
    return 0


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Genera una copia de siaf_dashboard.py para compartirla manualmente."
        )
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("siaf_dashboard_copy.py"),
        help=(
            "Ruta donde se guardará la copia (por defecto: siaf_dashboard_copy.py "
            "en el directorio actual)."
        ),
    )
    parser.add_argument(
        "--stdout",
        action="store_true",
        help="Imprime el contenido completo en pantalla en lugar de guardarlo en un archivo.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    output_path = None if args.stdout else args.output.expanduser().resolve()
    return dump_code(output_path, args.stdout)


if __name__ == "__main__":
    raise SystemExit(main())
