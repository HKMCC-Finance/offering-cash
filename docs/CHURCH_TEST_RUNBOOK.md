# On-site test runbook

Everything testable away from the hardware already passes (`python src/selftest.py`,
47 checks). This covers what only the BC-40, the check scanner and the printer can
answer. Work through it in order — step 1 decides whether the rest is meaningful.

Merged to `main` 2026-08-22. Estimated 60–90 minutes.

A formatted version of this document with tick-off checkboxes:
<https://claude.ai/code/artifact/79d961cc-14f1-471a-9af6-8630f2a24c0f>
(also committed here as `docs/church-test-runbook.html` for offline use).

---

## 0. Get the code — 2 min

This fix was merged straight to `main` on 2026-08-22 (after the column-offset check in
step 1 below was confirmed against the production template) rather than waiting on
these on-site steps, so `main` itself is **not** a safe rollback target anymore — it
already has the fix. On the church PC, in the existing repo folder:

```
git fetch origin
git checkout main
git pull

python src/selftest.py    # expect: 47 passed, 0 failed
```

The last known-working commit before this fix is tagged `pre-check-ocr-fix`. See
"If something breaks mid-count" below for how to fall back to it.

---

## 1. Confirm the template columns — 2 min — already checked, re-verify on paper

This was checked directly against the live file on the church PC before the merge
(openpyxl read of `E:\헌금보고서\헌금보고서_양식.xlsx`, cell by cell): denomination
labels are in **column B**, so the shipped offsets are correct, and the check block
header row (`COUNT`/`CHECK #`/`발행자`/`금액`) is at row 3, columns I/J/K/L, data
from row 4. That is why this was merged to `main` ahead of the rest of this runbook.

Worth a quick eyes-on confirmation anyway before a real count — open
`E:\헌금보고서\헌금보고서_양식.xlsx` and check the same things:

| Labels in | Meaning | Action |
|---|---|---|
| Column B | Shipped offsets are correct (expected) | Nothing to change; continue |
| Column A | Every write is one column too far right | **Stop**, revert to `pre-check-ocr-fix` (see "If something breaks mid-count"), and report back — the file must have changed since it was checked |

While the file is open, also record the row and column where the **수표정리**
summary block begins (needed in step 5) — that part was not double-checked against
the live formulas there in the same pass.

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

> **Largely answered already, from the PDFs already on this PC (2026-08-22).**
> Parsing the archived reports under
> `BC-40 UpperMonitor v13\Release\Data\<date>\` with the new `cash_data.py`:
>
> - **2026-08-16** — 6 consecutive reports, 12:24 → 12:35. Every denomination is
>   monotonically non-decreasing across all six ($1,854 → $3,802 → $4,707 →
>   $5,158 → $6,709 → $7,389). Within an uninterrupted session the BC-40 is
>   **cumulative**, so `subtract_offering()` has the right sign.
> - **2026-08-09** — 15 reports, and **4 of the transitions go negative**
>   (e.g. `[-308, -1, -280, -139, -64, 0, 0]`). The machine *does* get cleared
>   during a working day.
>
> So the cumulative assumption is correct, and the machine-reset guard is not
> theoretical — it will fire in real use. Still worth confirming live, but treat
> `기계가 초기화된 것으로 보입니다` as "the machine was cleared, re-pair the PDFs",
> **not** as evidence the subtraction needs inverting.

If the app stops with `기계가 초기화된 것으로 보입니다`, that is **not a failure** —
see the box above. Note which two PDFs were involved and move on.

Related risk worth watching: on a day with many reports (2026-08-09 had 15), the
1차 PDF is picked by "most recent file", so an extra test count between offerings
can pair the wrong two reports. If the 2차 numbers look wrong, check the file
times in the Data folder before suspecting the arithmetic.

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

Also record the scanner's actual filename format.

> **Partly answered already, from the scans already on this PC (2026-08-22).**
> Production scans live beside each report, in
> `E:\헌금보고서\<date>\헌금보고서_<date>_<mass>미사_checks\`.
> The real filename format is `<account>_<check#>.Front.tif`
> (e.g. `1010143133072_0308.Front.tif`) — note the **dot** before `Front`,
> not the underscore the examples assume. `parse_check_number()` handles it:
> all 17 files in the sample folder parsed correctly.
>
> - **File times are sequential within a scanning session** — 6–12 s apart,
>   consistent with sheets going through the feeder — so `--order scan`
>   (the default) is well founded.
> - **Scan order genuinely differs from filename order**, confirming the
>   original complaint. On the sample folder the two orderings disagreed
>   almost completely.
>
> Still worth confirming against a handwritten list, since only that proves
> the *direction* is right.

> **Watch out — a scan folder can hold more than one Sunday.**
> `헌금보고서_08-09-2026_11시미사_checks` contains 17 scans: 8 written
> 2026-08-09 12:44–12:45 and 9 written **2026-08-16** 12:44–12:45. A whole
> second week's checks were scanned into the previous week's folder. Every
> other folder checked held exactly one date, so this is a workflow slip
> rather than the norm — but `check_scan.py` will happily write all 17 rows
> into the one report. Before running it, confirm the folder holds only the
> current batch (`dir` sorted by date is enough).

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

**Do not pass `--summary-anchor` against this template.** Direct inspection of
`E:\헌금보고서\헌금보고서_양식.xlsx` (rows 19–29) found it already has a working
수표정리 block: 11 predefined amount rows — $5/10/15/20/25/30/45/50/100/200/600, not
the $5/10/20/25/50/100 that `check_summary.py`'s `DEFAULT_PREDEFINED_AMOUNTS` assumes
— that auto-tally via `COUNTIF($L$4:$L$30, ...)` formulas. Those formulas fire on
their own as soon as check amounts land in column L; `--summary-anchor` would write
Python-computed static values over them, using the wrong predefined list. Just run:

```
python src/check_scan.py --img_dir <scans> --report_file <report.xlsx> --print-layout
```

then confirm the 수표정리 block on the printed/opened report tallied correctly by
itself. If `check_summary.py`'s standalone summary is ever wanted for a different
template, fix `DEFAULT_PREDEFINED_AMOUNTS` (and `write_check_summary`'s column
mapping — production uses amount/label/count/total = A/B/C/D, not three columns
starting at one anchor) first; that is a separate follow-up, not part of this round.

### Printing — what to expect

Measured on the 2026-08-22 17시 report before the volunteers' test:

| | Before | After |
|---|---|---|
| Margins | 1.00" all round | 0.25" all round |
| Print scale | 68% | **83%** |
| Pages | 2 | **2** (unchanged) |
| Height on paper | 7.15" | 8.73" |
| Vertical fill | 79% | 83% |

The two-page split is deliberate and must stay: the template carries a manual
column break after column H, so **page 1 is the cash/check summary (B–H) and
page 2 is the check listing (I–L)**. Row heights are left alone — at 0.25"
margins the rows already fill the height, so the gain comes from the margins
and from replacing the template's fixed 68% scale.

Check on real paper:

- **Two pages, not one and not three.** Three means a page is spilling
  sideways; one means the column break was overridden.
- Page 1 ends cleanly after column H (금액 column of the cash block).
- Nothing clipped at the outer edges — 0.25" is tighter than some printers
  allow. If it clips, that is the margin to raise, not the scale.
- No extra column of numbers down the left of page 1 (that would be column A,
  the duplicate 수표정리 amounts).

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

**If you are running the desktop shortcut (the usual case)**, the code is inside
the packaged exe, so git will not change what you are running. Swap the build back:

1. rename `E:\CashCounting\dist\cash_count_ui` to `cash_count_ui_new`
2. rename `E:\CashCounting\dist\cash_count_ui_prefix_backup` to `cash_count_ui`

The shortcut points at that folder, so it picks up the restored build with no
shortcut edit. Reverse the two renames to go back to the new build.

**If you are running from source** (`python src/cash_count_ui.py`), `main` already
has this fix on it, so `git checkout main` will **not** help. Instead:

```
git checkout pre-check-ocr-fix -- .
```

Restores every source file to the last known-working version from before this fix
(tagged `pre-check-ocr-fix`), without switching branches or losing history. Note what
happened and which step you were on, then run `git checkout main -- .` afterward to
put the fix back once the count is safely finished on paper.

The report the parish keeps is a financial record. If any number looks wrong, trust
the paper count over the screen.

---

# Handoff: where the answers plug in

For whoever picks this up after the test. Each open question maps to one place.

| Question from the test | Change here |
|---|---|
| Denomination labels in column A or B? | ~~`src/cash_count_ui.py:55-59`~~ — **confirmed column B, offsets correct**, see step 1 |
| Check block row/columns in the template | ~~`src/check_scan.py:45-49`~~ — **confirmed row 3/cols I-L, offsets correct**, see step 1 |
| Is the 2차 PDF cumulative or standalone? | **Cumulative — confirmed** from archived PDFs (see step 2); `subtract_offering()` sign is correct. Resets between reports do occur, which is what the negative guard is for. No change expected in `src/cash_data.py:102` |
| Which `--order` matched the physical stack? | `src/check_scan.py` `list_check_images()` — make the winner the default |
| Where do the `$` box and written line actually sit? | `src/check_rois.json` (defaults in `src/check_fields.py:27`) |
| Correct predefined check amounts | `src/check_summary.py:14` `DEFAULT_PREDEFINED_AMOUNTS` — **already known to be wrong**: production is $5/10/15/20/25/30/45/50/100/200/600, not $5/10/20/25/50/100; not yet fixed, and `--summary-anchor` should not be used until it is (see step 5) |
| Did a different OCR backend win? | `src/ocr/backends.py` — `LegacyBackend:24`, `PaddleOCRBackend:101`, `QwenVLBackend:160`; change the default in `get_backend():221` and in `check_scan.py`'s `--backend` |
| Report spilled to a second page | `src/report_layout.py` `apply_print_layout()` — lower `max_row_height` or `fill_ratio` |
| Blind clicks misfiring on app launch | `src/cash_count_ui.py:47` `APP_LAUNCH_DELAY` (currently 0.5s; the older CLI used 3s) |

## What is deliberately unresolved

- **Column offsets were confirmed against the production template** (2026-08-22,
  direct openpyxl read of `E:\헌금보고서\헌금보고서_양식.xlsx`) and are correct as
  shipped — this is why the fix went to `main` ahead of the rest of this runbook.
  The repo's `Cash_Table_Formatter.xlsx` is a stale, different copy (labels one
  column left, and missing the 수표정리 block entirely) — do not use it as a
  reference for anything beyond row 14.
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
