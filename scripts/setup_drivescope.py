#!/usr/bin/env python3
"""Prepare DriveScope runtime folders and patch the ODRIV workbook."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PATCH_SCRIPT = ROOT / "scripts" / "patch_drivescope.py"


def main() -> int:
    command = [sys.executable, str(PATCH_SCRIPT), "--backup"]
    result = subprocess.run(command, cwd=ROOT)
    if result.returncode != 0:
        return result.returncode

    print()
    print("DriveScope setup complete.")
    print("Next step: place _OdrivDB.accdb in data/db/ before opening the workbook.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
