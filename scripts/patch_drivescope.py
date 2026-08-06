#!/usr/bin/env python3
"""Apply DriveScope fixes to the ODRIV workbook and local runtime layout."""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from openpyxl import load_workbook
from pyopenvba import ExcelFile, VBAModuleKind

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WORKBOOK = ROOT / "ODRIV_v29_2_1_AT (1).xlsm"
VBA_EXPORT = ROOT / "vba-export"
DATA_DB = ROOT / "data" / "db"

MODULES_TO_PATCH = [
    "ThisWorkbook.cls",
    "ClasseBdD.cls",
    "Module2.bas",
    "ModuleUtils.bas",
    "Outil_Boutons.bas",
    "form.frm",
    "DriveScopePaths.bas",
]


def patch_workbook_data(workbook_path: Path) -> None:
    workbook = load_workbook(workbook_path, keep_vba=True)
    cfg = workbook["CFG"]
    cfg["B1"] = str(DATA_DB)
    cfg["B2"] = str(DATA_DB)
    if cfg["B3"].value in (None, ""):
        cfg["B3"] = 2026
    workbook.save(workbook_path)


def ensure_runtime_files() -> None:
    year = 2026
    (DATA_DB / str(year)).mkdir(parents=True, exist_ok=True)
    version_file = DATA_DB / "ODRIV_DEV_VERSION.txt"
    if not version_file.exists():
        version_file.write_text("VERSION 1.0.0\n", encoding="utf-8")

    readme = DATA_DB / "README.txt"
    if not readme.exists():
        readme.write_text(
            "DriveScope local database folder\n\n"
            "Place _OdrivDB.accdb in this directory.\n"
            "Year-specific project databases (_OdrivDB_1.accdb, etc.) belong in the "
            f"{year}/ subfolder.\n",
            encoding="utf-8",
        )


def load_vba_source(module_name: str) -> str:
    path = VBA_EXPORT / module_name
    if not path.exists():
        raise FileNotFoundError(f"Missing exported VBA module: {path}")
    return path.read_text(encoding="utf-8").replace("\n", "\r\n")


def module_kind(module_name: str) -> VBAModuleKind:
    if module_name.endswith(".bas"):
        return VBAModuleKind.standard
    return VBAModuleKind.other


def patch_vba(workbook_path: Path) -> None:
    with ExcelFile(str(workbook_path)) as workbook:
        project = workbook.vba_project()
        existing = {module.name for module in project.modules}

        for module_name in MODULES_TO_PATCH:
            source = load_vba_source(module_name)
            vba_name = module_name.rsplit(".", 1)[0]
            kind = module_kind(module_name)

            if vba_name in existing:
                workbook.set_module(vba_name, source)
            else:
                project.add_module(vba_name, source, kind=kind)

        workbook.save(str(workbook_path))


def main() -> int:
    parser = argparse.ArgumentParser(description="Patch the DriveScope ODRIV workbook.")
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK,
        help="Path to the ODRIV .xlsm workbook",
    )
    parser.add_argument(
        "--backup",
        action="store_true",
        help="Create a .bak copy before patching",
    )
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 1

    if args.backup:
        backup_path = workbook_path.with_suffix(workbook_path.suffix + ".bak")
        shutil.copy2(workbook_path, backup_path)
        print(f"Backup created: {backup_path}")

    ensure_runtime_files()
    patch_workbook_data(workbook_path)
    patch_vba(workbook_path)
    print(f"Patched workbook: {workbook_path}")
    print(f"Local database folder: {DATA_DB}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
