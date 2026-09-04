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

    # 3B in bf16 is about 6 GB of weights, which fits the 8 GB card the
    # counting PC has. The crops are small, so activation memory is minor.
    # Override with the OFFERING_VLM_MODEL environment variable.
    DEFAULT_MODEL = os.environ.get("OFFERING_VLM_MODEL", "Qwen/Qwen2.5-VL-3B-Instruct")

    def __init__(self, model_id: str = None,
                 device_map: str = None, max_new_tokens: int = 64,
                 min_tokens: int = 256, max_tokens: int = 768):
        self._model_id = model_id or self.DEFAULT_MODEL
        self._device_map = device_map
        self._max_new_tokens = max_new_tokens
        self._min_tokens = min_tokens
        self._max_tokens = max_tokens
        self._model = None
        self._processor = None

    def _load(self):
        if self._model is None:
            import torch
            from transformers import AutoModelForImageTextToText, AutoProcessor

            on_gpu = torch.cuda.is_available()
            device_map = self._device_map or ("cuda:0" if on_gpu else "cpu")
            # Vision cost scales with the number of image patches, and a check
            # is a wide, low-detail document - capping the pixel budget is the
            # single biggest lever on per-check latency. 28x28 is the patch
            # size, so these are token budgets, not pixels on screen.
            self._processor = AutoProcessor.from_pretrained(
                self._model_id,
                min_pixels=self._min_tokens * 28 * 28,
                max_pixels=self._max_tokens * 28 * 28,
            )
            self._model = AutoModelForImageTextToText.from_pretrained(
                self._model_id,
                dtype=torch.bfloat16 if on_gpu else torch.float32,
                device_map=device_map,
            )
            self._model.eval()
        return self._processor, self._model

    def warmup(self) -> None:
        self._load()

    def _run(self, image: np.ndarray, prompt: str, max_new_tokens: int = None) -> OCRResult:
        processor, model = self._load()
        messages = [{
            "role": "user",
            "content": [{"type": "image"}, {"type": "text", "text": prompt}],
        }]
        chat = processor.apply_chat_template(messages, add_generation_prompt=True, tokenize=False)
        inputs = processor(text=[chat], images=[_to_pil(image)], return_tensors="pt")
        inputs = inputs.to(model.device)
        generated = model.generate(**inputs,
                                   max_new_tokens=max_new_tokens or self._max_new_tokens,
                                   do_sample=False)
        trimmed = generated[:, inputs["input_ids"].shape[1]:]
        text = processor.batch_decode(trimmed, skip_special_tokens=True)[0].strip()
        return OCRResult(text=text, confidence=None, engine=self.name)

    def read_printed(self, image: np.ndarray) -> OCRResult:
        return self._run(image, self.PRINTED_PROMPT)

    def read_handwriting(self, image: np.ndarray) -> OCRResult:
        return self._run(image, self.HANDWRITING_PROMPT)

    # One pass over the whole check yields both fields. The model is told the
    # amount is written twice so it can corroborate the two against each other
    # internally - which is what the separate courtesy/legal OCR reads were
    # doing badly - and to answer null instead of guessing.
    CHECK_PROMPT = (
        "Read this US personal check and reply with strict JSON only, no prose:"
        + chr(10)
        + '{"payer": <the person or people the account belongs to, printed at the '
          "top-left; names only, never the street address, city, state, ZIP or phone>, "
          '"amount": <the dollar amount as a plain number, e.g. 20 or 12.50>}'
        + chr(10)
        + "The amount appears twice: as digits in the box after the $ sign, and "
          "written in words on the line below. Use them to confirm each other."
        + chr(10)
        + "If you cannot read a field with confidence, set it to null. Never guess."
    )

    def read_check(self, image: np.ndarray):
        """Whole-check read: {"payer": str|None, "amount": float|None, "raw": str}."""
        return self.read_check_batch([image])[0]

    def read_check_batch(self, images, batch_size: int = 4):
        """Read several checks per forward pass.

        A single check leaves the GPU mostly idle, so batching is close to free
        throughput: the whole batch costs little more than the slowest member.
        """
        import torch

        processor, model = self._load()
        results = []
        for start in range(0, len(images), batch_size):
            chunk = images[start:start + batch_size]
            messages = [{
                "role": "user",
                "content": [{"type": "image"}, {"type": "text", "text": self.CHECK_PROMPT}],
            }]
            chat = processor.apply_chat_template(messages, add_generation_prompt=True,
                                                 tokenize=False)
            inputs = processor(text=[chat] * len(chunk),
                               images=[_to_pil(img) for img in chunk],
                               return_tensors="pt", padding=True)
            inputs = inputs.to(model.device)
            with torch.inference_mode():
                generated = model.generate(**inputs, max_new_tokens=96, do_sample=False)
            trimmed = generated[:, inputs["input_ids"].shape[1]:]
            for raw in processor.batch_decode(trimmed, skip_special_tokens=True):
                results.append(_parse_check_reply(raw.strip()))
        return results


def _parse_check_reply(raw: str):
    """Turn one model reply into {"payer", "amount", "raw"}."""
    import json
    import re

    payer = amount = None
    match = re.search(r"\{.*\}", raw, re.S)
    if match:
        try:
            payload = json.loads(match.group(0))
        except ValueError:
            payload = {}
        payer = payload.get("payer")
        if isinstance(payer, str) and re.fullmatch(r"\s*(null|none|)\s*", payer, re.I):
            payer = None
        amount = _parse_model_amount(payload.get("amount"))
    return {"payer": payer, "amount": amount, "raw": raw}


def _parse_model_amount(value):
    """Turn a model's amount reply into a float, or None when it declined."""
    import re

    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) if 0 < float(value) < 100000 else None
    text = str(value)
    if re.search(r"null|none|unreadable|unclear", text, re.I):
        return None
    match = re.search(r"(\d[\d,]*)(?:\.(\d{1,2}))?", text.replace(",", ""))
    if not match:
        return None
    try:
        parsed = float(match.group(1)) + (int(match.group(2)) / 100 if match.group(2) else 0)
    except ValueError:
        return None
    return parsed if 0 < parsed < 100000 else None


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
