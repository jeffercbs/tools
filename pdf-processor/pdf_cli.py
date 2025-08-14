import sys
import argparse
import fitz
from pdf_core import PDFProcessor
from logger_config import setup_default_logging


def main():
    """Función principal del CLI"""
    setup_default_logging()

    parser = argparse.ArgumentParser(
        description="Procesador de PDF para separar soportes",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python pdf_cli.py "REPORTE PAGOS BRL (Mayo).pdf" --initial-excel datos.xlsx --search-column "Transaction Reference Number" --rename-column "Rename"
  python pdf_cli.py input.pdf -o ./soportes_output --initial-excel mapping.xlsx --search-column "Buscar" --rename-column "Renombrar"
  python pdf_cli.py input.pdf --extract-text --initial-excel datos.xlsx --search-column "Ref" --rename-column "Nombre"
    """,
    )

    parser.add_argument("input_pdf", help="Ruta al archivo PDF de entrada")

    parser.add_argument(
        "-o",
        "--output",
        help="Directorio de salida (por defecto: ./soportes_separados)",
        default=None,
    )

    parser.add_argument(
        "--extract-text",
        action="store_true",
        help="Extraer también el texto de cada página",
    )

    parser.add_argument(
        "--detailed-info",
        action="store_true",
        help="Extraer información detallada de cada soporte de pago (incluye CSV)",
    )

    parser.add_argument(
        "--export-format",
        choices=["csv", "xlsx"],
        default="csv",
        help="Formato de exportación de datos (csv o xlsx)",
    )

    parser.add_argument(
        "--initial-excel",
        required=True,
        help="Ruta a un Excel con datos para búsqueda y renombrado (obligatorio)",
    )

    parser.add_argument(
        "--search-column",
        required=True,
        help="Nombre de la columna del Excel para buscar dentro de cada soporte (obligatorio)",
    )

    parser.add_argument(
        "--rename-column",
        required=True,
        help="Nombre de la columna del Excel con el valor para renombrar el PDF generado (obligatorio)",
    )

    args = parser.parse_args()

    try:
        mapping = (args.search_column, args.rename_column)
        processor = PDFProcessor(
            args.input_pdf,
            args.output,
            export_format=args.export_format,
            initial_excel_path=args.initial_excel,
            mapping_columns=mapping,
        )

        processor.validate_input()

        print(f"Iniciando procesamiento de: {args.input_pdf}")

        created_files = processor.separate_pages()

        if args.extract_text:
            print("Extrayendo texto de las páginas...")
            processor.extract_text_from_pages()

        if args.detailed_info:
            print("Extrayendo información detallada de soportes...")
            pdf_doc = fitz.open(args.input_pdf)
            metadata = processor.extract_metadata(pdf_doc)
            pdf_doc.close()
            processor.create_detailed_summary_report(metadata, created_files)

        print("\n✅ Procesamiento completado!")
        print(f"📁 Archivos creados: {len(created_files)}")
        print(f"📂 Ubicación: {processor.output_dir}")

    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
