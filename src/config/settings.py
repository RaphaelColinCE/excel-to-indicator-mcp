"""Configuration centralisée du serveur MCP"""

import os
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv

# Charger les variables d'environnement
load_dotenv()


class Settings:
    """Configuration du serveur MCP"""

    # ===== Chemins =====
    PROJECT_ROOT: Path = Path(__file__).parent.parent.parent
    DATA_DIR: Path = PROJECT_ROOT / "data"
    EXCEL_SOURCES_DIR: Path = Path(
        os.getenv("MCP_EXCEL_SOURCES_DIR", str(DATA_DIR / "actual" / "excel_sources"))
    )
    MAPPING_DIR: Path = Path(
        os.getenv("MCP_MAPPING_DIR", str(DATA_DIR / "actual" / "mappings"))
    )
    MAPPING_FILE: Path = Path(
        os.getenv("MCP_MAPPING_FILE", str(MAPPING_DIR / "mapping.json"))
    )
    INDICATORS_DIR: Path = Path(
        os.getenv("MCP_INDICATORS_DIR", str(DATA_DIR / "actual" / "indicators"))
    )
    OUTPUT_DIR: Path = Path(
        os.getenv("MCP_OUTPUT_DIR", str(DATA_DIR / "actual" / "output"))
    )
    LOG_DIR: Path = PROJECT_ROOT / "logs"
    LOG_FILE: Path = Path(os.getenv("MCP_LOG_FILE", str(LOG_DIR / "mcp.log")))

    # ===== Logging =====
    LOG_LEVEL: str = os.getenv("MCP_LOG_LEVEL", "INFO")

    # ===== Résolution de formules =====
    MAX_DEPTH_RESOLUTION: int = int(
        os.getenv("MCP_MAX_DEPTH_RESOLUTION", "5")
    )
    CACHE_ENABLED: bool = os.getenv("MCP_CACHE_ENABLED", "true").lower() == "true"

    # ===== Serveur MCP (optionnel) =====
    SERVER_HOST: str = os.getenv("MCP_SERVER_HOST", "localhost")
    SERVER_PORT: int = int(os.getenv("MCP_SERVER_PORT", "8000"))

    # ===== Debug =====
    DEBUG: bool = os.getenv("DEBUG", "false").lower() == "true"

    @classmethod
    def ensure_directories(cls) -> None:
        """Créer les répertoires s'ils n'existent pas"""
        cls.EXCEL_SOURCES_DIR.mkdir(parents=True, exist_ok=True)
        cls.MAPPING_DIR.mkdir(parents=True, exist_ok=True)
        cls.INDICATORS_DIR.mkdir(parents=True, exist_ok=True)
        cls.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        cls.LOG_DIR.mkdir(parents=True, exist_ok=True)

    @classmethod
    def get_excel_files(cls) -> list:
        """Récupérer tous les fichiers Excel disponibles"""
        if not cls.EXCEL_SOURCES_DIR.exists():
            return []
        return list(cls.EXCEL_SOURCES_DIR.glob("*.xlsx"))

    @classmethod
    def get_mapping_file(cls) -> Optional[Path]:
        """Récupérer le fichier de mapping s'il existe"""
        if cls.MAPPING_FILE.exists():
            return cls.MAPPING_FILE
        return None


# Instance globale
settings = Settings()
