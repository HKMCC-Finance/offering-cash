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
| `--backend` | `legacy` (default, EasyOCR + TrOCR), `paddleocr`, `qwen-vl`. |
| `--roster` | CSV/TXT of known donor names; OCR output is fuzzy-matched against it. |
| `--roi-config` | JSON of region boxes, defaults to `src/check_rois.json`. |
| `--prefer-amount` | `courtesy` (default) or `legal` when the two reads disagree. |
| `--summary-anchor` | `row,col` anchor for the 수표정리 summary block. |
| `--review-file` | Where to write the review sheet (defaults beside the report). |
| `--print-layout` | Apply the print layout after writing. |
| `--no-highlight` | Do not shade rows that need review. |

### Amount reading

Every check carries the amount twice: the courtesy box (numerals, upper right)
and the legal line (handwritten words). Both are read and reconciled. Rows
where the two disagree, where only one could be read, or where the name did not
match the roster are shaded in the report and listed in the review sheet
(`<report>_검토.xlsx`) alongside the raw OCR text.

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

`src/report_layout.py` shrinks the page margins to 0.25" and grows row heights
so the report fills the page instead of leaving a wide band of white space.
Applied automatically by the cash app, and by `--print-layout` for checks.

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
  in `check_scan.py` have **not** been verified against the production template
  `E:\헌금보고서\헌금보고서_양식.xlsx`. The copy in this repo
  (`Cash_Table_Formatter.xlsx`) puts the denomination labels in column A, which
  would make those offsets one column too far right. Confirm before changing.
- `src/check_rois.json` holds standard-layout estimates, not measured values.
  Retune against real scans.
