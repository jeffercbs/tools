import logging


def get_logger(name: str = "pdf_processor") -> logging.Logger:
    """Obtiene un logger configurado para el procesador PDF"""
    logger = logging.getLogger(name)
    if not logger.handlers:
        logger.addHandler(logging.NullHandler())
    return logger


def setup_default_logging():
    """Configura logging por defecto para CLI"""
    lg = logging.getLogger("pdf_processor")
    lg.setLevel(logging.INFO)

    if lg.handlers:
        return

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"
    )

    sh = logging.StreamHandler()
    sh.setLevel(logging.INFO)
    sh.setFormatter(fmt)
    lg.addHandler(sh)
