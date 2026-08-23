from .base import OCRBackend, OCRResult
from .backends import BACKENDS, LegacyBackend, PaddleOCRBackend, QwenVLBackend, get_backend

__all__ = [
    "OCRBackend",
    "OCRResult",
    "BACKENDS",
    "LegacyBackend",
    "PaddleOCRBackend",
    "QwenVLBackend",
    "get_backend",
]
