"""Promuove un Excel da staging private a public quando il video YouTube va live.

Usage:
    python tools/release_excel.py --slug tassazione_etf
    python tools/release_excel.py --slug pac           # alias accettato anche senza _2026

Cosa fa:
1. Legge .xlsx da repo staging (e:/sviluppo/marco-fire-simulatori-staging/simulatori/<slug>_2026.xlsx)
2. Copia in repo public (e:/sviluppo/marco-fire-simulatori/simulatori/<slug>_2026.xlsx)
3. git add + commit + push nel repo public
4. git rm + commit + push nel repo staging (rimuove il file da staging)

Pre-requisiti: entrambi i repo gia' clonati con remote configurato (PAT marcofire-it).
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

PUBLIC_ROOT = Path("e:/sviluppo/marco-fire-simulatori")
STAGING_ROOT = Path("e:/sviluppo/marco-fire-simulatori-staging")
SUBDIR = "simulatori"


def run(cmd: list[str], cwd: Path, check: bool = True) -> subprocess.CompletedProcess:
    print(f"  $ {' '.join(cmd)}  (cwd: {cwd.name})")
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if r.stdout.strip():
        print(r.stdout.rstrip())
    if r.stderr.strip():
        print(r.stderr.rstrip(), file=sys.stderr)
    if check and r.returncode != 0:
        sys.exit(f"FAIL: exit {r.returncode}")
    return r


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--slug", required=True, help="es. tassazione_etf, pac, pensione_2056")
    p.add_argument("--overwrite", action="store_true",
                   help="se public xlsx esiste gia' lo sovrascrive (re-release, es. dopo MEF day update)")
    p.add_argument("--keep-staging", action="store_true",
                   help="non rimuovere xlsx da staging dopo release (utile per re-release ripetuti)")
    args = p.parse_args()

    slug = args.slug.removesuffix("_2026")
    filename = f"{slug}_2026.xlsx"

    staging_xlsx = STAGING_ROOT / SUBDIR / filename
    public_xlsx = PUBLIC_ROOT / SUBDIR / filename

    if not staging_xlsx.exists():
        sys.exit(f"FAIL: {staging_xlsx} non trovato in staging")
    if public_xlsx.exists() and not args.overwrite:
        sys.exit(f"FAIL: {public_xlsx} esiste gia' in public. Usa --overwrite per re-release.")
    if not PUBLIC_ROOT.exists() or not STAGING_ROOT.exists():
        sys.exit("FAIL: uno dei due repo locali non esiste")

    print(f"[RELEASE] {filename}  staging -> public")
    print(f"  size: {staging_xlsx.stat().st_size / 1024:.1f} KB")

    # 1. Copy staging -> public
    print("\n[1/4] Copy staging -> public")
    shutil.copy2(staging_xlsx, public_xlsx)
    print(f"  copied -> {public_xlsx}")

    # 2. Commit + push public
    print("\n[2/4] Commit + push PUBLIC")
    run(["git", "add", f"{SUBDIR}/{filename}"], cwd=PUBLIC_ROOT)
    run(["git", "commit", "-m", f"Release {filename}: video YouTube live"], cwd=PUBLIC_ROOT)
    run(["git", "push", "origin", "main"], cwd=PUBLIC_ROOT)

    # 3. Remove from staging
    print("\n[3/4] Remove from STAGING")
    run(["git", "rm", f"{SUBDIR}/{filename}"], cwd=STAGING_ROOT)
    run(["git", "commit", "-m", f"Remove {filename}: promosso a public"], cwd=STAGING_ROOT)
    run(["git", "push", "origin", "main"], cwd=STAGING_ROOT)

    # 4. Done
    print("\n[4/4] DONE")
    public_url = f"https://github.com/marcofire-it/marco-fire-simulatori/raw/main/{SUBDIR}/{filename}"
    print(f"  Public URL: {public_url}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
