"""Backend-agnostic OCR interface.

The check pipeline talks to this interface only, so an engine can be swapped
(and benchmarked) without touching the extraction or Excel-writing code.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional

import numpy as np


@dataclass
class OCRResult:
    """A single OCR read.

    confidence is 0.0-1.0 where the engine reports one, and None where it does
    not. Callers must treat None as "unknown", not as "bad".
    """

    text: str
    confidence: Optional[float] = None
    engine: str = ""
    raw: object = field(default=None, repr=False)

    @property
    def is_empty(self) -> bool:
        return not self.text or not self.text.strip()


class OCRBackend(ABC):
    """Base class for an OCR engine.

    Two entry points because checks mix two very different kinds of text:
    the payer block is machine-printed, the amount and payee are handwritten.
    Engines that do not distinguish the two may point both at the same model.
    """

    name: str = "base"

    @abstractmethod
    def read_printed(self, image: np.ndarray) -> OCRResult:
        """Read machine-printed text from a BGR image region."""

    @abstractmethod
    def read_handwriting(self, image: np.ndarray) -> OCRResult:
        """Read handwritten text from a BGR image region."""

    def read_check(self, image: "np.ndarray"):
        """Read a whole check in one pass, returning {"payer", "amount"}.

        Only vision-language backends can do this - they are told what the
        document is, so they read the amount as an amount rather than
        transcribing glyphs. Backends that cannot return None, and the caller
        falls back to reading each field separately.

        A field the model is not confident about comes back None, so the
        report cell is left blank for a volunteer rather than guessed at.
        """
        return None

    def warmup(self) -> None:
        """Optional: load weights ahead of the first real call."""

    def __repr__(self) -> str:
        return f"<{type(self).__name__} name={self.name!r}>"
