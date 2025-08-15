TO# weekly-csv-backup-to-dropbox

Back up recent CSVs to Dropbox as a dated zip, with optional retention.

## Setup
1. Install Python 3.9+.
2. `python -m venv .venv && source .venv/bin/activate` (Windows: `. .venv/Scripts/activate`)
3. `pip install -r requirements.txt`
4. Copy `.env.example` → `.env`, fill values.
5. Test: `python backup.py`

## Config notes
- `SOURCE_DIRS`: comma‑separated folders. Supports network paths (e.g., `\\\\SERVER\\Share`).
- `FILE_GLOB`: defaults to `**/*.csv` (recursive). Use `*.csv` for flat directories.
- `LOOKBACK_DAYS`: files with modified time ≥ now‑N days.
- `DROPBOX_DEST_FOLDER`: must begin with `/`. If your app has **App folder** access, `/` refers to the app‑scoped root.
- `RETAIN_DAYS`: set `0` to disable deletion.
- `DRY_RUN`: `true` to simulate.

## Security
Prefer storing `DROPBOX_ACCESS_TOKEN` in your OS secret store or CI secret rather than plain `.env`.

## Troubleshooting
- **Auth failed**: confirm token scopes and that the token hasn’t expired.
- **Nothing to back up**: widen `LOOKBACK_DAYS` or verify paths/glob.
- **Permission denied**: ensure the Dropbox app has rights to `DROPBOX_DEST_FOLDER`.
