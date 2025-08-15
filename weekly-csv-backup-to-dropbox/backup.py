import os
import sys
import io
import time
import zipfile
import hashlib
from pathlib import Path
from datetime import datetime, timedelta
from typing import List

from dotenv import load_dotenv
import dropbox
from dropbox.files import WriteMode

load_dotenv()

# --- Config ---
SOURCE_DIRS = [p.strip() for p in os.getenv("SOURCE_DIRS", "").split(",") if p.strip()]
FILE_GLOB = os.getenv("FILE_GLOB", "**/*.csv")
LOOKBACK_DAYS = int(os.getenv("LOOKBACK_DAYS", "7"))
DROPBOX_DEST_FOLDER = os.getenv("DROPBOX_DEST_FOLDER", "/weekly_csv_backups")
DROPBOX_ACCESS_TOKEN = os.getenv("DROPBOX_ACCESS_TOKEN", "")
ZIP_NAME_PREFIX = os.getenv("ZIP_NAME_PREFIX", "csv_backup")
RETAIN_DAYS = int(os.getenv("RETAIN_DAYS", "30"))
DRY_RUN = os.getenv("DRY_RUN", "false").lower() == "true"

# --- Helpers ---
def log(msg: str):
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{ts}] {msg}")


def fail(msg: str, code: int = 1):
    log(f"ERROR: {msg}")
    sys.exit(code)


def find_files() -> List[Path]:
    if not SOURCE_DIRS:
        fail("No SOURCE_DIRS provided in .env")
    cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)
    files = []
    for root in SOURCE_DIRS:
        base = Path(root)
        if not base.exists():
            log(f"WARN: Source dir not found: {base}")
            continue
        for p in base.glob(FILE_GLOB):
            try:
                if p.is_file():
                    mtime = datetime.fromtimestamp(p.stat().st_mtime)
                    if mtime >= cutoff:
                        files.append(p)
            except Exception as e:
                log(f"WARN: Skipping {p}: {e}")
    files.sort()
    return files


def make_zip(files: List[Path]) -> Path:
    if not files:
        fail("No CSV files found in lookback window; nothing to back up.", code=2)
    ts = datetime.now().strftime('%Y-%m-%d_%H%M%S')
    zip_name = f"{ZIP_NAME_PREFIX}_{ts}.zip"
    out_dir = Path.cwd() / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    zip_path = out_dir / zip_name

    log(f"Creating archive: {zip_path}")
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        manifest_lines = []
        for f in files:
            arcname = f.relative_to(Path(SOURCE_DIRS[0]).anchor if f.is_absolute() else Path.cwd())
            zf.write(f, arcname=str(f))  # keep original relative tree for clarity
            h = hashlib.sha256()
            with open(f, 'rb') as fh:
                for chunk in iter(lambda: fh.read(1024 * 1024), b''):
                    h.update(chunk)
            manifest_lines.append(f"{h.hexdigest()}  {f}")
        # Add a manifest with hashes
        manifest = "\n".join(manifest_lines) + "\n"
        zf.writestr("MANIFEST.sha256", manifest)
    return zip_path


def get_dbx() -> dropbox.Dropbox:
    if not DROPBOX_ACCESS_TOKEN:
        fail("DROPBOX_ACCESS_TOKEN is missing. Set it in .env or your env vars.")
    dbx = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)
    # light ping
    try:
        dbx.users_get_current_account()
    except Exception as e:
        fail(f"Dropbox auth failed: {e}")
    return dbx


def upload_file(dbx: dropbox.Dropbox, local_path: Path, dest_folder: str) -> str:
    dest_path = f"{dest_folder.rstrip('/')}/{local_path.name}"
    log(f"Uploading to Dropbox: {dest_path}")
    if DRY_RUN:
        log("DRY_RUN=true → skipping upload")
        return dest_path
    with open(local_path, 'rb') as f:
        dbx.files_upload(f.read(), dest_path, mode=WriteMode.add)
    return dest_path


def apply_retention(dbx: dropbox.Dropbox, dest_folder: str):
    if RETAIN_DAYS <= 0:
        return
    cutoff = datetime.now() - timedelta(days=RETAIN_DAYS)
    log(f"Retention: deleting files in {dest_folder} older than {RETAIN_DAYS} days (before {cutoff.date()})")
    try:
        res = dbx.files_list_folder(dest_folder)
    except dropbox.exceptions.ApiError as e:
        if isinstance(e.error, dropbox.files.ListFolderError):
            log(f"WARN: Destination folder may not exist yet: {dest_folder} ({e})")
            return
        raise
    entries = list(res.entries)
    while res.has_more:
        res = dbx.files_list_folder_continue(res.cursor)
        entries.extend(res.entries)

    for ent in entries:
        if isinstance(ent, dropbox.files.FileMetadata) and ent.name.lower().endswith('.zip'):
            # ent.server_modified is timezone-aware
            if ent.server_modified.replace(tzinfo=None) < cutoff:
                path = ent.path_lower
                log(f"Deleting old backup: {path}")
                if not DRY_RUN:
                    dbx.files_delete_v2(path)


def main():
    log("Starting weekly CSV backup → Dropbox")
    log(f"Sources: {SOURCE_DIRS} | glob={FILE_GLOB} | lookback={LOOKBACK_DAYS}d | dest={DROPBOX_DEST_FOLDER}")
    files = find_files()
    log(f"Discovered {len(files)} CSV files to back up")
    zip_path = make_zip(files)
    dbx = get_dbx()
    dest = upload_file(dbx, zip_path, DROPBOX_DEST_FOLDER)
    log(f"Uploaded: {dest}")
    apply_retention(dbx, DROPBOX_DEST_FOLDER)
    log("Backup completed successfully ✅")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        raise
    except Exception as e:
        fail(f"Unhandled error: {e}")
