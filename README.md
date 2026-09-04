# Cash Counting Application

Python tooling for the parish offering count: a GUI that drives the BC-40 bill
counter and fills the offering report, plus a separate command-line tool that
OCRs scanned checks into the same report.

## Requirements

- Python 3.x
- `pip install -r requirements.txt`

## Cash counting (GUI)

```
python src/cash_count_ui.py
```

1. Select the mass time (7시반, 9시, 11시, 17시).
2. Answer whether there is a second offering (2차 헌금).
3. Click "현금 카운팅 시작". The app launches UpperMonitor, configures the
   connection, and waits for you to confirm the count is done.
4. It reads the generated PDF and writes the report to
   `E:\헌금보고서\<date>\헌금보고서_<date>_<mass>미사.xlsx`.

### Second offering

The BC-40 reports cumulative totals within a session, so the 2차 figures are
derived by subtracting the 1차 counts per denomination. The 1차 counts are
also written to a snapshot file (`.offering_1차_<date>_<mass>.json`) beside the
report, so closing the app between offerings does not lose them.

The app stops and warns rather than writing suspect numbers when:

- the subtraction produces negative quantities (the machine was cleared
  between offerings, so the second report is standalone, not cumulative);
- the 2차 total comes out as zero;
- the 2차 PDF is the same file as the 1차 one (no new count was produced).

Errors are shown in a dialog and appended to `E:\헌금보고서\cash_count.log`.
The packaged exe runs without a console, so the log is the only record.

## Check scanning (command line)

Run separately, after the checks have been scanned:

```
python src/check_scan.py --img_dir <scan folder> --report_file <report.xlsx>
```

| Option | Purpose |
|---|---|
| `--order` | `scan` (default), `filename`, or `checknum`. `scan` writes rows in the order the checks physically went through the scanner. |
| `--backend` | `qwen-vl` (default), `legacy` (EasyOCR + TrOCR), `paddleocr`. |
| `--roster` | CSV/TXT of known donor names; OCR output is fuzzy-matched against it. |
| `--roi-config` | JSON of region boxes, defaults to `src/check_rois.json`. |
| `--allow-single-read` | Write an amount even when the two readings disagree (default: prefer the words). |
| `--summary-anchor` | `row,col` anchor for the 수표정리 summary block. |
| `--review-file` | Where to write the review sheet (defaults beside the report). |
| `--no-print-layout` | Skip the print layout (margins, scale, centring). Applied by default. |
| `--no-highlight` | Do not shade rows that need review. |

### How the reading works

The default `qwen-vl` backend shows the whole check to a vision-language model
(Qwen2.5-VL-3B, run locally on the GPU) and asks for the payer and both amount
readings in one pass. It is told what the document is, so it reads the amount
as an amount rather than transcribing glyph by glyph - which is what the
character-level OCR did badly, turning a handwritten `10` into `/0` or `(0`.

The model reports the **digits** (the box after the `$`) and the **words** (the
line below) separately and is explicitly told not to reconcile them. The words
win: measured on 19 real checks, the two disagreed twice and the words were
right both times - the digits had read the *check number* as the amount. Taking
the words scored 19/19; taking the digits wrote two wrong numbers. The digits
are kept as a cross-check, and a disagreement shades the row and is explained
in the review sheet (`<report>_검토.xlsx`) rather than being hidden.

A field the model will not commit to comes back empty, so a volunteer fills a
blank cell instead of a plausible wrong number reaching a financial record.

Measured against human-corrected reports for 2026-08-16 and 2026-08-30:

| | old `legacy` OCR | `qwen-vl` |
|---|---|---|
| amounts correct | 6/10 | **19/19** |
| silently wrong | 1 | **0** |
| payer names | 7/10 | 18/19 |
| per check | 8.3s | 3.3s |
| startup | ~60s | 10.7s |

Requires a CUDA GPU and a CUDA build of torch; on CPU it runs but is far too
slow to be worth using. Check numbers are never read by OCR - they come from
the scanner's filename, which was correct on all 19.

### Check summary (수표정리)

`--summary-anchor` writes a tally by amount: predefined rows ($5, $10, $20,
$25, $50, $100) always appear, and any other amount is appended below them in
ascending order.

## Comparing OCR models

```
python src/benchmark_ocr.py --img_dir <scans> --truth truth.csv \
    --backend legacy --backend paddleocr
```

`truth.csv` is `filename,name,amount`. Reports amount accuracy, name exact-match
rate, name CER, and how many actual errors the review flag caught.

## Print layout

`src/report_layout.py` shrinks the page margins to 0.25", centres the table
vertically, and re-scales to fill the page (68% -> 83% on the production
template) while respecting the manual column break that puts the cash summary
on page 1 and the check listing on page 2. Applied automatically by both the
cash app and the check scanner.

## How volunteers run it

`HKMCC_CheckScan_v3.exe` in `E:\Check Scanner execution V3\` is unchanged and
still runs `python3 .\check_scan_v3.py --img_dir ... --report_file ...`. That
file is now a shim: it re-executes the same arguments under `E:\CashCounting\.venv`
and `src/check_scan.py`, because the `python3` on PATH is the Windows Store one
with different torch and transformers versions than this was tested against.

To change behaviour, edit this repo - not the copy in that folder. The previous
OCR code is kept there as `check_scan_v3.py.backup_20260904`; restoring it is
the rollback.

## On-site testing

Several things can only be verified with the counting machine, the scanner and the
printer present. `docs/CHURCH_TEST_RUNBOOK.md` is the step-by-step procedure, the
list of files to collect, and a handoff table mapping each open question to the
exact constant that answers it. `docs/church-test-runbook.html` is the same
content as a standalone page with tick-off checkboxes, for offline use.

## Self-test

```
python src/selftest.py
```

Covers everything that does not need the machine, the scanner, or the OCR
models, and prints the list of things that still require a real test on site.

## File structure

| Path | Purpose |
|---|---|
| `src/cash_count_ui.py` | Cash counting GUI |
| `src/cash_data.py` | Cash parsing, alignment, 2차 subtraction (no GUI) |
| `src/check_scan.py` | Check OCR pipeline and CLI |
| `src/check_fields.py` | Check region crops and amount parsing |
| `src/check_summary.py` | 수표정리 summary table |
| `src/check_rois.json` | Region-of-interest boxes, tunable without code changes |
| `src/report_layout.py` | Print margins and row heights |
| `src/benchmark_ocr.py` | OCR backend comparison |
| `src/selftest.py` | Self-test |
| `src/ocr/` | Pluggable OCR backends |
| `src/utils/coordinate_finder.py` | Live mouse-coordinate readout |
| `src/utils/coordinate_capture.py` | Interactive coordinate capture |
| `src/cash_count.py` | Superseded CLI version, kept for reference only |

## Notes

- The column offsets in `cash_count_ui.py` (`FIRST_OFFERING_COL`,
  `SECOND_OFFERING_COL`, `DATE_CELL`, `TIME_CELL`) and the check block columns
  in `check_scan.py` **have been confirmed correct** against the production
  template `E:\헌금보고서\헌금보고서_양식.xlsx` (2026-08-22, direct openpyxl
  read). The copy in this repo (`Cash_Table_Formatter.xlsx`) is a stale,
  different file — it puts the denomination labels in column A and is missing
  the 수표정리 summary block entirely — do not use it as a reference beyond
  row 14.
- `src/check_rois.json` was measured against the 08-16 and 08-30 batches, but
  the default `qwen-vl` backend does not use it - it reads the whole check. The
  boxes only matter for the `legacy` and `paddleocr` backends.
- `check_summary.py`'s `DEFAULT_PREDEFINED_AMOUNTS` ($5/10/20/25/50/100) do not
  match the production template's 수표정리 block, which has 11 predefined rows
  ($5/10/15/20/25/30/45/50/100/200/600) that already auto-tally via `COUNTIF`
  formulas once check amounts land in column L. Do not pass `--summary-anchor`
  against the production template until this is fixed — it would overwrite
  those working formulas with values computed from the wrong list.
