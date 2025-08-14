from pdf_core import PDFProcessor
from logger_config import setup_default_logging, get_logger
from pdf_cli import main

# Mantener la interfaz original
__all__ = ["PDFProcessor", "setup_default_logging", "get_logger", "main"]

if __name__ == "__main__":
    main()
