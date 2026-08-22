"""Concrete OCR engines.

All heavy dependencies (torch, easyocr, transformers, paddle) are imported
lazily inside the backend that needs them, so importing this module - and the
benchmark harness - stays cheap on a machine that only has some of them.
"""

import os
from typing import Optional

import numpy as np

from .base import OCRBackend, OCRResult


def _to_pil(image: np.ndarray):
    """BGR ndarray -> RGB PIL.Image."""
    import cv2
    from PIL import Image

    return Image.fromarray(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))


class LegacyBackend(OCRBackend):
    """The engine pair the app has shipped with: EasyOCR + TrOCR-handwritten.

    Kept as the default so behaviour does not change until a replacement is
    benchmarked against it on real scans. This is the baseline to beat.
    """

    name = "legacy"

    def __init__(self, languages=("en",), gpu: Optional[bool] = None,
                 model_storage_directory: Optional[str] = None,
                 trocr_model: str = "microsoft/trocr-base-handwritten"):
        self._languages = list(languages)
        self._gpu = gpu
        # Was hardcoded to one operator's home directory, which breaks on any
        # other PC. Honour an env override, else let EasyOCR use its default.
        self._model_dir = model_storage_directory or os.environ.get("EASYOCR_MODEL_DIR")
        self._trocr_model = trocr_model
        self._reader = None
        self._processor = None
        self._model = None

    def _ensure_gpu_flag(self) -> bool:
        if self._gpu is not None:
            return self._gpu
        try:
            import torch

            self._gpu = bool(torch.cuda.is_available())
        except Exception:
            self._gpu = False
        return self._gpu

    def warmup(self) -> None:
        self._get_reader()
        self._get_trocr()

    def _get_reader(self):
        if self._reader is None:
            import easyocr

            kwargs = {"gpu": self._ensure_gpu_flag()}
            if self._model_dir:
                kwargs["model_storage_directory"] = self._model_dir
            self._reader = easyocr.Reader(self._languages, **kwargs)
        return self._reader

    def _get_trocr(self):
        if self._model is None:
            from transformers import TrOCRProcessor, VisionEncoderDecoderModel

            self._processor = TrOCRProcessor.from_pretrained(self._trocr_model, use_fast=True)
            self._model = VisionEncoderDecoderModel.from_pretrained(self._trocr_model)
        return self._processor, self._model

    def read_printed(self, image: np.ndarray) -> OCRResult:
        results = self._get_reader().readtext(image)
        if not results:
            return OCRResult(text="", confidence=None, engine=self.name, raw=results)
        text = ", ".join(res[1] for res in results)
        confs = [float(res[2]) for res in results if len(res) > 2 and res[2] is not None]
        return OCRResult(
            text=text,
            confidence=(sum(confs) / len(confs)) if confs else None,
            engine=self.name,
            raw=results,
        )

    def read_handwriting(self, image: np.ndarray) -> OCRResult:
        processor, model = self._get_trocr()
        pixel_values = processor(images=_to_pil(image), return_tensors="pt").pixel_values
        generated_ids = model.generate(pixel_values, max_new_tokens=20)
        text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]
        # TrOCR's generate() gives no usable per-sequence confidence here.
        return OCRResult(text=text, confidence=None, engine=self.name)


class PaddleOCRBackend(OCRBackend):
    """PaddleOCR (PP-OCRv5 / PaddleOCR-VL).

    Runs fully on-premise, reports per-line confidence - which the legacy
    TrOCR path cannot - and handles printed and handwritten text with one
    model. Candidate replacement; must be benchmarked before it becomes default.
    """

    name = "paddleocr"

    def __init__(self, lang: str = "en", use_gpu: Optional[bool] = None, **kwargs):
        self._lang = lang
        self._use_gpu = use_gpu
        self._kwargs = kwargs
        self._ocr = None

    def _get_ocr(self):
        if self._ocr is None:
            from paddleocr import PaddleOCR

            kwargs = dict(self._kwargs)
            kwargs.setdefault("lang", self._lang)
            if self._use_gpu is not None:
                kwargs.setdefault("use_gpu", self._use_gpu)
            self._ocr = PaddleOCR(**kwargs)
        return self._ocr

    def warmup(self) -> None:
        self._get_ocr()

    def _read(self, image: np.ndarray) -> OCRResult:
        raw = self._get_ocr().ocr(image)
        lines = []
        confs = []
        # PaddleOCR's return shape has moved around between majors; accept both
        # the nested [[ [box, (text, conf)], ... ]] and flat variants.
        for page in (raw or []):
            for entry in (page or []):
                try:
                    payload = entry[1]
                    text, conf = payload[0], float(payload[1])
                except (TypeError, IndexError, ValueError):
                    continue
                lines.append(text)
                confs.append(conf)
        return OCRResult(
            text=", ".join(lines),
            confidence=(sum(confs) / len(confs)) if confs else None,
            engine=self.name,
            raw=raw,
        )

    def read_printed(self, image: np.ndarray) -> OCRResult:
        return self._read(image)

    def read_handwriting(self, image: np.ndarray) -> OCRResult:
        return self._read(image)


class QwenVLBackend(OCRBackend):
    """Qwen3-VL / Qwen2.5-VL run locally through transformers.

    Best open-weights accuracy on handwriting as of 2026, at the cost of
    needing a GPU. Prompted per field, which is why it takes a prompt hint.
    """

    name = "qwen-vl"

    PRINTED_PROMPT = "Transcribe the printed text in this image exactly. Output only the text."
    HANDWRITING_PROMPT = "Transcribe the handwriting in this image exactly. Output only the text."

    def __init__(self, model_id: str = "Qwen/Qwen3-VL-8B-Instruct",
                 device_map: str = "auto", max_new_tokens: int = 64):
        self._model_id = model_id
        self._device_map = device_map
        self._max_new_tokens = max_new_tokens
        self._model = None
        self._processor = None

    def _load(self):
        if self._model is None:
            from transformers import AutoModelForImageTextToText, AutoProcessor

            self._processor = AutoProcessor.from_pretrained(self._model_id)
            self._model = AutoModelForImageTextToText.from_pretrained(
                self._model_id, device_map=self._device_map
            )
        return self._processor, self._model

    def warmup(self) -> None:
        self._load()

    def _run(self, image: np.ndarray, prompt: str) -> OCRResult:
        processor, model = self._load()
        messages = [{
            "role": "user",
            "content": [{"type": "image"}, {"type": "text", "text": prompt}],
        }]
        chat = processor.apply_chat_template(messages, add_generation_prompt=True, tokenize=False)
        inputs = processor(text=[chat], images=[_to_pil(image)], return_tensors="pt")
        inputs = inputs.to(model.device)
        generated = model.generate(**inputs, max_new_tokens=self._max_new_tokens)
        trimmed = generated[:, inputs["input_ids"].shape[1]:]
        text = processor.batch_decode(trimmed, skip_special_tokens=True)[0].strip()
        return OCRResult(text=text, confidence=None, engine=self.name)

    def read_printed(self, image: np.ndarray) -> OCRResult:
        return self._run(image, self.PRINTED_PROMPT)

    def read_handwriting(self, image: np.ndarray) -> OCRResult:
        return self._run(image, self.HANDWRITING_PROMPT)


BACKENDS = {
    LegacyBackend.name: LegacyBackend,
    PaddleOCRBackend.name: PaddleOCRBackend,
    QwenVLBackend.name: QwenVLBackend,
}


def get_backend(name: str = "legacy", **kwargs) -> OCRBackend:
    """Instantiate a backend by name. Raises KeyError with the valid names."""
    try:
        cls = BACKENDS[name]
    except KeyError:
        raise KeyError(f"Unknown OCR backend {name!r}. Available: {sorted(BACKENDS)}") from None
    return cls(**kwargs)
