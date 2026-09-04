"""Compatibility shim - the real check-scanning pipeline lives in the git repo.

HKMCC_CheckScan_v3.exe runs:

    python3 .\\check_scan_v3.py --img_dir <scans> --report_file <report.xlsx>

That "python3" is the Windows Store Python, which carries its own torch and
transformers at versions the pipeline was never tested against. So this file
does two things and nothing else:

  1. forwards the exe's arguments unchanged, and
  2. re-executes them under E:\\CashCounting\\.venv, which is the interpreter
     the accuracy and latency numbers were measured on.

Nothing about the operator's workflow changes: same exe, same command line.
To change behaviour, edit the repo - not this file.

The previous OCR code that lived here is kept as check_scan_v3.py.backup_20260904.
"""
import datetime
import os
import subprocess
import sys

REPO = r"E:\CashCounting"
VENV_PYTHON = os.path.join(REPO, ".venv", "Scripts", "python.exe")
PIPELINE = os.path.join(REPO, "src", "check_scan.py")


def fail(message):
    log(f"FAILED: {message.splitlines()[0] if message else ''}")
    # The launcher shows a console, so a plain message is enough - but make it
    # unmistakable, because the volunteer sees this mid-count.
    print("=" * 70)
    print("CHECK SCANNING COULD NOT START")
    print("=" * 70)
    print(message)
    print()
    print("The counting can still be finished by entering the checks by hand.")
    sys.exit(1)


LOG = r"E:\헌금보고서\check_scan.log"


def log(message):
    """Append to a log beside the reports.

    HKMCC_CheckScan_v3.exe is a GUI application, so anything printed here may
    never reach a visible console. Without this, a failure mid-count looks like
    "nothing happened" and leaves nothing to diagnose afterwards.
    """
    try:
        os.makedirs(os.path.dirname(LOG), exist_ok=True)
        with open(LOG, "a", encoding="utf-8") as handle:
            handle.write(f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} {message}{os.linesep}")
    except OSError:
        pass


def main():
    log(f"launched with {' '.join(sys.argv[1:])}")
    if not os.path.isfile(VENV_PYTHON):
        fail(f"Python environment not found:{os.linesep}  {VENV_PYTHON}{os.linesep}"
             f"The {REPO} repo and its .venv must both be present.")
    if not os.path.isfile(PIPELINE):
        fail(f"Check-scanning code not found:{os.linesep}  {PIPELINE}")

    command = [VENV_PYTHON, PIPELINE] + sys.argv[1:]
    print(f"Running the offering check pipeline from {REPO}")
    # Inherit stdio so the operator sees progress and any error as it happens.
    completed = subprocess.run(command, cwd=REPO)
    log(f"finished with exit code {completed.returncode}")
    if completed.returncode != 0:
        print()
        print("Check scanning did not finish cleanly. The checks can still be "
              "entered by hand.")
        print(f"Details were written to {LOG}")
    sys.exit(completed.returncode)


if __name__ == "__main__":
    main()
