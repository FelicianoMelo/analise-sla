import logging
from pathlib import Path


def setup_logger() -> logging.Logger:
    """
    Configura logger da aplicação.

    Returns:
        logging.Logger
    """

    Path("logs").mkdir(exist_ok=True)

    logger = logging.getLogger("SLA")

    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(funcName)s | %(message)s"
    )

    file_handler = logging.FileHandler(
        "logs/analise_sla.log",
        encoding="utf-8"
    )

    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger