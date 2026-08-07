#!/usr/bin/env python3
"""
Finance Tracker — Backup & Restore
Usage:
    python backup_restore.py backup          # create a timestamped backup
    python backup_restore.py restore <file>   # restore from a backup zip
    python backup_restore.py list             # list available backups
"""

import sys, os, shutil, json, zipfile, glob
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
BACKUP_DIR = BASE / "backups"
DB = BASE / "finance.db"
CATEGORIES = BASE / "categories.json"
PEOPLE = BASE / "people.json"          # only for reference; DB is authoritative
EXCEL = BASE / "finance_tracker.xlsx"  # if it exists
EXTRA_FILES = [CATEGORIES, PEOPLE, EXCEL]


def backup():
    BACKUP_DIR.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    name = f"finance_backup_{ts}.zip"
    path = BACKUP_DIR / name

    if not DB.exists():
        print(f"❌ Database not found: {DB}")
        sys.exit(1)

    files_to_zip = [DB]
    for f in EXTRA_FILES:
        if f.exists():
            files_to_zip.append(f)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files_to_zip:
            z.write(f, f.name)
            print(f"  📦 {f.name}")

    size = os.path.getsize(path)
    print(f"\n✅ Backup saved: {path} ({size / 1024:.1f} KB)")
    return path


def restore(path_str):
    path = Path(path_str)
    if not path.exists():
        # try inside backups/
        alt = BACKUP_DIR / path_str
        if alt.exists():
            path = alt
        else:
            print(f"❌ File not found: {path}")
            print(f"   Also checked: {alt}")
            sys.exit(1)

    # Confirm
    print(f"⚠️  This will OVERWRITE your current data with contents from:")
    print(f"   {path}")
    confirm = input("   Continue? (y/N): ").strip().lower()
    if confirm != "y":
        print("Cancelled.")
        return

    restored = []
    with zipfile.ZipFile(path, "r") as z:
        members = z.namelist()
        for member in members:
            dest = BASE / member
            # safety: only restore known files
            if dest.name in ("finance.db", "categories.json", "people.json", "finance_tracker.xlsx"):
                z.extract(member, BASE)
                restored.append(member)
                print(f"  📄 {member}")

    if "finance.db" in members:
        print("\n✅ Restore complete! Restart your Flask app to pick up the restored DB.")
    else:
        print("\n⚠️  Restore finished, but no database was found in the backup.")


def list_backups():
    BACKUP_DIR.mkdir(exist_ok=True)
    all_zips = sorted(BACKUP_DIR.glob("*.zip"), reverse=True)
    if not all_zips:
        print("No backups found in:", BACKUP_DIR)
        return

    print(f"Backups in {BACKUP_DIR}:\n")
    for z in all_zips:
        size = os.path.getsize(z)
        modified = datetime.fromtimestamp(os.path.getmtime(z)).strftime("%Y-%m-%d %H:%M")
        kb = size / 1024
        print(f"  {z.name}  ({kb:.1f} KB)  [{modified}]")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1]
    if cmd == "backup":
        backup()
    elif cmd == "restore":
        if len(sys.argv) < 3:
            print("Usage: python backup_restore.py restore <file>")
            sys.exit(1)
        restore(sys.argv[2])
    elif cmd == "list":
        list_backups()
    else:
        print(f"Unknown command: {cmd}")
        print(__doc__)
