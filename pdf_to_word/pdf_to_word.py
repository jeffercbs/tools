from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional


def convert_pdf_to_word(pdf_path: Path, word_path: Optional[Path] = None) -> Path:
    """Convierte un PDF a DOCX.

    Args:
        pdf_path: Ruta al archivo PDF.
        word_path: Ruta de salida del DOCX. Si no se especifica, usa el mismo nombre con .docx.

    Returns:
        Ruta del archivo DOCX generado.

    Raises:
        FileNotFoundError: Si el PDF no existe.
        ValueError: Si la ruta no es un PDF válido.
        ImportError: Si 'pdf2docx' no está instalado.
        Exception: Si ocurre un error durante la conversión.
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists() or not pdf_path.is_file():
        raise FileNotFoundError(f"El archivo no existe: {pdf_path}")
    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError(f"La entrada debe ser un PDF: {pdf_path}")

    word_path = Path(word_path) if word_path else pdf_path.with_suffix(".docx")
    word_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        from pdf2docx import Converter  # type: ignore
    except ImportError as exc:
        raise ImportError(
            "Falta la dependencia 'pdf2docx'. Instálala con: pip install pdf2docx"
        ) from exc

    cv = Converter(str(pdf_path))
    try:
        cv.convert(str(word_path))
    finally:
        cv.close()

    return word_path


def convert_multiple_pdfs(directory_path: Path) -> tuple[int, int]:
    """Convierte todos los PDFs en un directorio.

    Returns una tupla (exitosos, total).
    """
    directory = Path(directory_path)
    if not directory.exists() or not directory.is_dir():
        raise FileNotFoundError(f"El directorio no existe: {directory}")

    pdf_files = sorted(directory.glob("*.pdf"))
    total = len(pdf_files)
    exitosos = 0

    for pdf in pdf_files:
        try:
            print(f"Convirtiendo: {pdf.name}")
            convert_pdf_to_word(pdf)
            exitosos += 1
        except Exception as e:  # noqa: BLE001
            print(f"Error con {pdf.name}: {e}")

    return exitosos, total


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convertir archivos PDF a Word (.docx)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Ejemplos:\n"
            "  python pdf_to_word.py documento.pdf\n"
            "  python pdf_to_word.py documento.pdf -o resultado.docx\n"
            "  python pdf_to_word.py -d ./pdfs\n"
        ),
    )

    parser.add_argument("pdf_file", nargs="?", help="Archivo PDF a convertir")
    parser.add_argument(
        "-d",
        "--directory",
        help="Directorio con archivos PDF para conversión en lote",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Archivo de salida (.docx) para conversión individual",
    )

    args = parser.parse_args(argv)

    if bool(args.pdf_file) == bool(args.directory):
        parser.error("Debe especificar un archivo PDF o un directorio, pero no ambos.")
    if args.output and not args.pdf_file:
        parser.error("--output solo aplica cuando se especifica un archivo PDF.")

    try:
        if args.pdf_file:
            out = convert_pdf_to_word(
                Path(args.pdf_file), Path(args.output) if args.output else None
            )
            print(f"Listo: {out}")
        else:
            exitosos, total = convert_multiple_pdfs(Path(args.directory))
            print(f"Resumen: {exitosos}/{total} conversiones exitosas")
        return 0
    except ImportError as e:
        print(e)
        return 1
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")
        return 2
    except Exception as e:  # noqa: BLE001
        print(f"Error inesperado: {e}")
        return 3


if __name__ == "__main__":
    sys.exit(main())
