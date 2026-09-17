class SourceGapError(RuntimeError):
    """Raised when the supplied original source references code that is not present in the archive."""


__all__ = ["SourceGapError"]
