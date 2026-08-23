# On-site test runbook

Everything testable away from the hardware already passes (`python src/selftest.py`,
47 checks). This covers what only the BC-40, the check scanner and the printer can
answer. Work through it in order — step 1 decides whether the rest is meaningful.

Branch: `fix/check-ocr-and-report-layout`. Estimated 60–90 minutes.

A formatted version of this document with tick-off checkboxes:
<https://claude.ai/code/artifact/79d961cc-14f1-471a-9af6-8630f2a24c0f>
(also committed here as `docs/church-test-runbook.html` for offline use).

---

## 0. Get the code — 2 min

On the church PC, in the existing repo folder:

```
git fetch origin
git checkout fix/check-ocr-and-report-layout
git pull

python src/selftest.py    # expect: 47 passed, 0 failed
```

`main` is untouched. If anything goes wrong mid-count, `git checkout main` restores
the current working version immediately.

---

## 1. Confirm the template columns — 2 min — BLOCKING

Open `E:\헌금보고서\헌금보고서_양식.xlsx` and find the column holding the
denomination labels (1, 2, 5, 10, 20, 50, 100).

| Labels in | Meaning | Action |
|---|---|---|
| Column B | Shipped offsets are correct | Nothing to change; continue |
| Column A | Every write is one column too far right | **Stop.** Cash figures have been landing in the wrong columns and the Sub Total formulas have been summing the wrong cells |

While the file is open, also record:

- the row the check block starts on;
- the columns used for 순번 / CHECK # / 발행자 / 금액;
- the row and column where the **수표정리** summary block begins (needed in step 5).

This is the one thing that could not be verified remotely, and it changes the
meaning of every number in steps 2 and 3.

---

## 2. Cash — a real two-offering count — 20 min

The goal is not only "does it work" but capturing both PDFs so the machine's actual
behaviour is on record.

1. Launch `python src/cash_count_ui.py`, pick a mass time, answer **예** to 2차 헌금.
2. Count a small first offering across several denominations. Note the counts on paper.
3. Complete 1차 and check the 1차 columns in the report.
4. **Copy the 1차 PDF out of the Data folder** before counting again — it may be overwritten.
5. Count a second, different offering. Note those counts.
6. Complete 2차 and copy the 2차 PDF out too.
7. Compare the 2차 columns against the physical count.

**The key question:** open both PDFs side by side — is the second one cumulative
(1차 + 2차) or standalone? The whole subtraction rests on it being cumulative.

If the app stops with `기계가 초기화된 것으로 보입니다`, that is **not a failure**.
It means the second report is standalone, the cumulative assumption is wrong, and
`subtract_offering()` needs inverting. Record it and move on.

Then the restart case, the main suspect for the original bug:

8. Run 1차 again, then **close the app completely**.
9. Reopen, pick the same mass time, run 2차. The subtraction should still happen —
   it now reloads the 1차 counts from the snapshot file on disk.

---

## 3. Checks — scan order — 15 min

Before scanning, **write down the check numbers in the order you feed them**. That
handwritten list is the ground truth.

```
python src/check_scan.py --img_dir <scan folder> --report_file <report.xlsx>
```

Report rows should match the handwritten list top to bottom. If not, the scanner
isn't stamping file times in scan order — try the alternatives and record which
one matches:

```
python src/check_scan.py ... --order filename
python src/check_scan.py ... --order checknum
```

Also record the scanner's actual filename format (e.g. `DEP_1043_Front.tif`).

---

## 4. Checks — accuracy baseline — 30 min

Use a real batch: 100 checks if available, 20 if not. This is what turns "roughly
30% wrong" into something measurable.

After the run, open the review sheet written beside the report
(`<report>_검토.xlsx`). It lists every check with the raw OCR text, both amount
reads, and why each row was flagged.

### Building the truth file

Correct the 발행자 and 금액 columns by eye, then save as `truth.csv` with three
columns, renamed:

```
filename,name,amount
DEP_1043_Front.tif,John Smith,125.00
DEP_207_Front.tif,Mary Kim,20.00
```

This file is the durable asset — every future model comparison runs against it:

```
python src/benchmark_ocr.py --img_dir <scans> --truth truth.csv --backend legacy
```

If amount accuracy comes back near zero, the region boxes in `src/check_rois.json`
are cropping the wrong part of the scans — that is a cropping problem, not a model
problem. Open one `.tif`, note where the `$` box and the written-amount line fall as
a fraction of image width/height, and send the numbers.

Also record:

- Does the run print "Using GPU"? Determines whether a local Qwen3-VL swap is realistic.
- Is a donor roster (교적 / 봉헌자 명단) available as a file? Biggest single lever
  on name accuracy — see `--roster`.

---

## 5. 수표정리 summary and printing — 10 min

Using the summary block position recorded in step 1 (example: row 20, column I):

```
python src/check_scan.py ... --summary-anchor 20,9 --print-layout
```

Predefined rows ($5 / $10 / $20 / $25 / $50 / $100) should always appear even at
zero, with any other amount appended below the $100 row in ascending order.

- Confirm the predefined list matches the paper form. 5/10/20/25/50/100 is an
  assumption — correct it if it differs.
- Print a finished report; top and bottom margins should be about ¼ inch.
- If it spills to a second page, the row heights need pulling back.

---

## 6. Bring back

| File | Why |
|---|---|
| `헌금보고서_양식.xlsx` | Production template. Settles the column question and the summary anchor permanently |
| 1차 PDF + 2차 PDF | Settles whether the BC-40 reports cumulatively |
| 20–30 check scans (`.tif`) | Lets the region boxes be retuned and models tested without another trip |
| `truth.csv` | The scoring set; without it model changes are guesswork |
| `<report>_검토.xlsx` | Shows which failures are OCR vs parsing vs cropping |
| `cash_count.log` | In `E:\헌금보고서\`. Records anything that failed silently |
| A printed report (photo is fine) | Confirms margins on real paper |

---

## If something breaks mid-count

The count still has to be finished and correct — that comes first.

```
git checkout main
```

Restores the current working version immediately. Note what happened and which step
you were on; the branch stays on GitHub.

The report the parish keeps is a financial record. If any number looks wrong, trust
the paper count over the screen.

---

# Handoff: where the answers plug in

For whoever picks this up after the test. Each open question maps to one place.

| Question from the test | Change here |
|---|---|
| Denomination labels in column A or B? | `src/cash_count_ui.py:55-59` (`DATA_START_ROW`, `DATE_CELL`, `TIME_CELL`, `FIRST_OFFERING_COL`, `SECOND_OFFERING_COL`) |
| Check block row/columns in the template | `src/check_scan.py:45-49` (`CHECK_START_ROW`, `SEQ_COL`, `CHECK_NUM_COL`, `PAYER_COL`, `AMOUNT_COL`) |
| Is the 2차 PDF cumulative or standalone? | `src/cash_data.py:102` `subtract_offering()` — if standalone, skip the subtraction and write the frame directly |
| Which `--order` matched the physical stack? | `src/check_scan.py` `list_check_images()` — make the winner the default |
| Where do the `$` box and written line actually sit? | `src/check_rois.json` (defaults in `src/check_fields.py:27`) |
| Correct predefined check amounts | `src/check_summary.py:14` `DEFAULT_PREDEFINED_AMOUNTS` |
| Did a different OCR backend win? | `src/ocr/backends.py` — `LegacyBackend:24`, `PaddleOCRBackend:101`, `QwenVLBackend:160`; change the default in `get_backend():221` and in `check_scan.py`'s `--backend` |
| Report spilled to a second page | `src/report_layout.py` `apply_print_layout()` — lower `max_row_height` or `fill_ratio` |
| Blind clicks misfiring on app launch | `src/cash_count_ui.py:47` `APP_LAUNCH_DELAY` (currently 0.5s; the older CLI used 3s) |

## What is deliberately unresolved

- **Column offsets were left exactly as the original code had them.** The repo's
  `Cash_Table_Formatter.xlsx` disagrees with them by one column, but the production
  template was never available to confirm against. Do not change them on the basis
  of the repo copy alone.
- **`src/check_rois.json` holds standard-check-layout estimates, not measurements.**
  Until step 4 runs, treat the courtesy-box amount read as unproven.
- **`src/cash_count.py` is superseded** by `cash_count_ui.py` + `cash_data.py`. It
  has a different output path and different click coordinates. Kept for reference;
  do not develop against it.

## Invariant

`python src/selftest.py` must stay at 0 failures. It covers amount parsing,
courtesy/legal reconciliation, scan ordering, roster matching, the 수표정리 summary,
print layout, 2차 subtraction (including the machine-reset and duplicate-PDF cases),
snapshot round-tripping, and workbook writing — all without hardware.
