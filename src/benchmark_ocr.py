"""Score OCR backends against hand-labelled checks.

Usage:
    python src/benchmark_ocr.py --img_dir <scans> --truth truth.csv \
        --backend legacy --backend paddleocr

truth.csv needs a filename column plus the correct values:

    filename,name,amount
    IMG_1043_Front.tif,John Smith,125.00

Without this, "is the new model better?" cannot be answered - which is the
whole reason the reported 30% / 10-20% error rates have been hard to act on.
"""

import argparse
import csv
import os
import time
from typing import Dict, List, Optional, Sequence

from check_fields import load_rois
from check_scan import extract_check
from ocr import get_backend


def levenshtein(a: str, b: str) -> int:
    """Edit distance, iterative two-row version."""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    previous = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        current = [i]
        for j, cb in enumerate(b, start=1):
            current.append(min(previous[j] + 1, current[j - 1] + 1,
                               previous[j - 1] + (ca != cb)))
        previous = current
    return previous[-1]


def cer(reference: str, hypothesis: str) -> Optional[float]:
    """Character error rate; None when there is no reference to score against."""
    reference = (reference or "").strip()
    hypothesis = (hypothesis or "").strip()
    if not reference:
        return None
    return levenshtein(reference, hypothesis) / len(reference)


def normalise(text: Optional[str]) -> str:
    return " ".join((text or "").lower().split())


def load_truth(path: str) -> Dict[str, dict]:
    """Map filename -> {'name': str, 'amount': float|None}."""
    truth = {}
    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            filename = (row.get("filename") or row.get("file") or "").strip()
            if not filename:
                continue
            raw_amount = (row.get("amount") or "").replace("$", "").replace(",", "").strip()
            try:
                amount = float(raw_amount) if raw_amount else None
            except ValueError:
                amount = None
            truth[filename] = {"name": (row.get("name") or "").strip(), "amount": amount}
    return truth


def evaluate(backend_name: str, image_dir: str, truth: Dict[str, dict],
             roi_path: Optional[str] = None, roster_path: Optional[str] = None,
             prefer_amount: str = "courtesy") -> dict:
    """Run one backend over the labelled set and score it."""
    from check_scan import load_roster

    rois = load_rois(roi_path)
    roster = load_roster(roster_path)
    backend = get_backend(backend_name)
    backend.warmup()

    rows: List[dict] = []
    started = time.time()
    for index, (filename, expected) in enumerate(sorted(truth.items()), start=1):
        path = os.path.join(image_dir, filename)
        if not os.path.exists(path):
            continue
        record = extract_check(path, backend, rois, sequence=index, roster=roster,
                               prefer_amount=prefer_amount)
        amount_ok = (expected["amount"] is not None and record.amount is not None
                     and abs(record.amount - expected["amount"]) < 0.005)
        rows.append({
            "filename": filename,
            "expected_name": expected["name"],
            "got_name": record.payer_name or "",
            "name_cer": cer(expected["name"], record.payer_name or ""),
            "name_exact": normalise(expected["name"]) == normalise(record.payer_name),
            "expected_amount": expected["amount"],
            "got_amount": record.amount,
            "amount_ok": amount_ok,
            "amount_status": record.amount_status,
            "courtesy": record.courtesy_amount,
            "legal": record.legal_amount,
            "needs_review": record.needs_review,
        })
    elapsed = time.time() - started

    total = len(rows) or 1
    cers = [r["name_cer"] for r in rows if r["name_cer"] is not None]
    flagged = [r for r in rows if r["needs_review"]]
    missed = [r for r in rows if not r["amount_ok"]]
    # The point of the review flag is to catch the wrong ones: what share of
    # actual errors does it surface?
    caught = sum(1 for r in missed if r["needs_review"])

    return {
        "backend": backend_name,
        "checks": len(rows),
        "amount_accuracy": round(sum(r["amount_ok"] for r in rows) / total, 4),
        "name_exact_accuracy": round(sum(r["name_exact"] for r in rows) / total, 4),
        "name_cer": round(sum(cers) / len(cers), 4) if cers else None,
        "flagged_rate": round(len(flagged) / total, 4),
        "error_catch_rate": round(caught / len(missed), 4) if missed else None,
        "seconds_per_check": round(elapsed / total, 2),
        "_rows": rows,
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--img_dir", required=True)
    parser.add_argument("--truth", required=True, help="CSV of filename,name,amount")
    parser.add_argument("--backend", action="append", default=None,
                        help="repeatable; defaults to legacy")
    parser.add_argument("--roi-config", default=None)
    parser.add_argument("--roster", default=None)
    parser.add_argument("--prefer-amount", default="courtesy", choices=("courtesy", "legal"))
    parser.add_argument("--out", default="benchmark_results.csv")
    args = parser.parse_args(argv)

    truth = load_truth(args.truth)
    print(f"Loaded {len(truth)} labelled checks from {args.truth}\n")

    summaries = []
    detail_rows = []
    for backend_name in (args.backend or ["legacy"]):
        print(f"--- {backend_name} ---")
        result = evaluate(backend_name, args.img_dir, truth, args.roi_config,
                          args.roster, args.prefer_amount)
        for row in result.pop("_rows"):
            row["backend"] = backend_name
            detail_rows.append(row)
        summaries.append(result)
        for key, value in result.items():
            print(f"  {key:22} {value}")
        print()

    if detail_rows:
        with open(args.out, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=list(detail_rows[0]))
            writer.writeheader()
            writer.writerows(detail_rows)
        print(f"Per-check detail written to {args.out}")

    if len(summaries) > 1:
        best = max(summaries, key=lambda s: s["amount_accuracy"])
        print(f"\nBest amount accuracy: {best['backend']} at {best['amount_accuracy']:.1%}")


if __name__ == "__main__":
    main()
