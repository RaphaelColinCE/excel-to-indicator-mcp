"""Configuration du logging"""

import logging
import logging.handlers
from src.config.settings import settings


def setup_logging() -> logging.Logger:
    """Configurer le logging du serveur MCP"""
    # Créer le répertoire logs s'il n'existe pas
    settings.LOG_DIR.mkdir(parents=True, exist_ok=True)

    # Créer un logger
    logger = logging.getLogger("mcp_server")
    logger.setLevel(getattr(logging, settings.LOG_LEVEL))

    # Formatter
    formatter = logging.Formatter(
        fmt="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Handler fichier (rotation)
    file_handler = logging.handlers.RotatingFileHandler(
        settings.LOG_FILE,
        maxBytes=10 * 1024 * 1024,  # 10 MB
        backupCount=5,
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Handler console (en debug)
    if settings.DEBUG:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    return logger


# Logger global
logger = setup_logging()
