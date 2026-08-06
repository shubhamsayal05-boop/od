# DriveScope (ODRIV)

DriveScope is the portable distribution of the ODRIV driveability rating workbook. It replaces brittle corporate network paths with a local `data/db` layout and adds resilient startup behavior.

## What was fixed

- **Portable database paths**: CFG no longer depends on `\\besn01\...` network shares.
- **Auto-provisioned runtime folders**: `data/db/` and `data/db/2026/` are created during setup.
- **Clear startup errors**: Missing `_OdrivDB.accdb` now shows the exact expected path.
- **Version check hardening**: `checkDevVersion` no longer crashes when `ODRIV_DEV_VERSION.txt` is absent.
- **Sheet hiding bug**: `HideFeuille` now uses `xlSheetVeryHidden` instead of the invalid `-1` constant.
- **Debug halt removed**: The `Stop` statement in fatal error handling was removed.
- **Consistent sheet naming**: `TARGET VEHICLE` casing is normalized.

## Quick start

```bash
pip install -r requirements.txt
python scripts/setup_drivescope.py
```

Then copy your `_OdrivDB.accdb` file into `data/db/` and open `ODRIV_v29_2_1_AT (1).xlsm` in Excel with macros enabled.

## Project layout

```
.
├── ODRIV_v29_2_1_AT (1).xlsm   # Patched workbook
├── data/db/                     # Local database root
├── scripts/
│   ├── setup_drivescope.py      # One-command setup
│   └── patch_drivescope.py      # Workbook + VBA patcher
├── vba-export/                  # Human-readable VBA sources
└── tool-tutorial-app/             # DriveScope tutorial dashboard
```

## Tutorial dashboard

```bash
cd tool-tutorial-app
python main.py
```

On Linux, install Tkinter first: `sudo apt install python3-tk`.

## Requirements

- Microsoft Excel with macros enabled
- Microsoft Access Database Engine (ACE OLEDB 12.0 provider)
- `_OdrivDB.accdb` from your ODRIV deployment package

## Re-applying patches

If you update VBA under `vba-export/`:

```bash
python scripts/patch_drivescope.py --backup
```
