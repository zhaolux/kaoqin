class KaoqinError(Exception):
    """Base exception for the kaoqin project."""


class ConfigError(KaoqinError):
    """Raised when configuration files are missing or invalid."""


class InputFileError(KaoqinError):
    """Raised when the DingTalk input file cannot be found or read."""


class SheetDetectionError(KaoqinError):
    """Raised when the DingTalk worksheet cannot be detected."""


class GenerationError(KaoqinError):
    """Raised when attendance workbook generation fails."""
