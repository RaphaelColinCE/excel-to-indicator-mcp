"""Exceptions personnalisées du serveur MCP"""


class MCPException(Exception):
    """Exception de base pour le serveur MCP"""
    pass


# ===== Exceptions Excel =====

class ExcelFileNotFoundError(MCPException):
    """Le fichier Excel n'a pas été trouvé"""
    pass


class ExcelParsingError(MCPException):
    """Erreur lors du parsing d'un fichier Excel"""
    pass


class ExcelFormulaError(MCPException):
    """Erreur lors de l'analyse d'une formule Excel"""
    pass


# ===== Exceptions Mapping =====

class MappingFileNotFoundError(MCPException):
    """Le fichier de mapping n'a pas été trouvé"""
    pass


class MappingParsingError(MCPException):
    """Erreur lors du parsing du fichier de mapping"""
    pass


class MappingNotFoundError(MCPException):
    """Un mapping demandé n'a pas été trouvé"""
    pass


# ===== Exceptions Conversion =====

class ConversionError(MCPException):
    """Erreur lors de la conversion"""
    pass


class FormulaParsingError(MCPException):
    """Erreur lors du parsing d'une formule"""
    pass


class ReferenceResolutionError(MCPException):
    """Erreur lors de la résolution d'une référence"""
    pass


# ===== Exceptions Validation =====

class ValidationError(MCPException):
    """Erreur lors de la validation"""
    pass
