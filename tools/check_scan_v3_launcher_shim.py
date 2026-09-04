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
import os
import subprocess
import sys

REPO = r"E:\CashCounting"
VENV_PYTHON = os.path.join(REPO, ".venv", "Scripts", "python.exe")
PIPELINE = os.path.join(REPO, "src", "check_scan.py")


def fail(message):
    # The launcher shows a console, so a plain message is enough - but make it
    # unmistakable, because the volunteer sees this mid-count.
    print("=" * 70)
    print("CHECK SCANNING COULD NOT START")
    print("=" * 70)
    print(message)
    print()
    print("The counting can still be finished by entering the checks by hand.")
    sys.exit(1)


def main():
    if not os.path.isfile(VENV_PYTHON):
        fail(f"Python environment not found:{os.linesep}  {VENV_PYTHON}{os.linesep}"
             f"The {REPO} repo and its .venv must both be present.")
    if not os.path.isfile(PIPELINE):
        fail(f"Check-scanning code not found:{os.linesep}  {PIPELINE}")

    command = [VENV_PYTHON, PIPELINE] + sys.argv[1:]
    print(f"Running the offering check pipeline from {REPO}")
    # Inherit stdio so the operator sees progress and any error as it happens.
    completed = subprocess.run(command, cwd=REPO)
    sys.exit(completed.returncode)


if __name__ == "__main__":
    main()
