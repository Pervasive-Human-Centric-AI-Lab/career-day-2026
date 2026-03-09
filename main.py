#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Dict, List

import pandas as pd


def norm(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("–", "-").replace("—", "-")
    return s


def find_col(columns: List[str], patterns: List[str]) -> Optional[str]:
    ncols = {c: norm(c) for c in columns}
    for pat in patterns:
        rx = re.compile(pat, re.IGNORECASE)
        for orig, n in ncols.items():
            if rx.search(n):
                return orig
    return None


def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s


def slugify(name: str) -> str:
    """
    Simple slugify for filenames:
    - lowercases
    - replaces spaces with hyphens
    - removes non alnum/hyphen
    """
    s = name.strip().lower()
    s = s.replace("&", " and ")
    s = re.sub(r"[^\w\s-]", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"-{2,}", "-", s)
    return s.strip("-") or "azienda"


def unique_path(base_dir: Path, slug: str, suffix: str = ".md") -> Path:
    """
    Ensure unique filename in base_dir by appending -2, -3...
    """
    p = base_dir / f"{slug}{suffix}"
    if not p.exists():
        return p
    i = 2
    while True:
        p = base_dir / f"{slug}-{i}{suffix}"
        if not p.exists():
            return p
        i += 1


def main() -> int:
    # Fixed input/output locations
    in_path = Path("data") / "schede.xlsx"
    out_dir = Path("output")
    companies_dir = out_dir / "companies"
    index_path = out_dir / "index.md"

    companies_dir.mkdir(parents=True, exist_ok=True)

    if not in_path.exists():
        raise FileNotFoundError(f"Input file not found: {in_path}")

    df = pd.read_excel(in_path, engine="openpyxl")
    cols = list(df.columns)

    # Detect columns
    colmap: Dict[str, Optional[str]] = {
        "nome_azienda": find_col(cols, [
            r"\bnome\b.*\bazienda\b",
            r"\bazienda\b.*\bnome\b",
            r"\bragione\s*sociale\b",
            r"\bcompany\b.*\bname\b",
        ]),
        "descrizione": find_col(cols, [
            r"\bdescrizione\b.*\bazienda\b",
            r"\bazienda\b.*\bdescrizione\b",
            r"\bcompany\b.*\bdescription\b",
        ]),
        "cosa_cercate": find_col(cols, [
            r"\bcosa\b.*\bcercate\b",
            r"\boffrite\b",
            r"\bposizion",
            r"\bfigure\b",
            r"\bprofili\b",
            r"\bwe are looking\b",
            r"\bwhat\b.*\blooking\b",
        ]),
        "livello_studenti": find_col(cols, [
            r"\blivello\b.*\bstudent",
            r"\btriennal",
            r"\bmagistral",
            r"\blevel\b.*\bstudent",
        ]),
        "indirizzo_contatto": find_col(cols, [
            r"\bindirizzo\b.*\bcontatt",
            r"\bcontatt",
            r"\bemail\b",
            r"\btelefono\b",
            r"\baddress\b",
            r"\bcontact\b",
        ]),
    }

    missing = [k for k, v in colmap.items() if v is None]
    if missing:
        print("WARNING: Some fields could not be auto-detected:", ", ".join(missing))
        print("Available columns:")
        for c in cols:
            print(" -", c)
        print("Continuing: missing fields will be left blank.\n")

    records = []
    for _, row in df.iterrows():
        nome = clean_text(row.get(colmap["nome_azienda"], "")) if colmap["nome_azienda"] else ""
        if not nome:
            continue

        rec = {
            "Nome azienda": nome,
            "Descrizione azienda": clean_text(row.get(colmap["descrizione"], "")) if colmap["descrizione"] else "",
            "Cosa cercate/offrite": clean_text(row.get(colmap["cosa_cercate"], "")) if colmap["cosa_cercate"] else "",
            "Per che livello di studenti": clean_text(row.get(colmap["livello_studenti"], "")) if colmap["livello_studenti"] else "",
            "Indirizzo e contatto": clean_text(row.get(colmap["indirizzo_contatto"], "")) if colmap["indirizzo_contatto"] else "",
        }
        records.append(rec)

    # Sort by company name
    records.sort(key=lambda r: r["Nome azienda"].strip().lower())

    # Write one page per company
    written = []
    for r in records:
        slug = slugify(r["Nome azienda"])
        page_path = unique_path(companies_dir, slug, ".md")

        md = []
        md.append(f"# {r['Nome azienda']}\n")

        md.append("## Descrizione azienda\n")
        md.append((r["Descrizione azienda"] or "-") + "\n")

        md.append("## Cosa cercate/offrite\n")
        md.append((r["Cosa cercate/offrite"] or "-") + "\n")

        md.append("## Per che livello di studenti\n")
        md.append((r["Per che livello di studenti"] or "-") + "\n")

        md.append("## Indirizzo e contatto\n")
        md.append((r["Indirizzo e contatto"] or "-") + "\n")

        page_path.write_text("\n".join(md), encoding="utf-8")
        written.append((r["Nome azienda"], page_path))

    # Write index with links
    idx = []
    idx.append("# Aziende partecipanti – Career Day\n")
    idx.append("Elenco aziende (clicca per aprire la scheda):\n")
    for name, p in written:
        rel = p.relative_to(out_dir).as_posix()
        idx.append(f"- [{name}]({rel})")
    idx.append("")  # trailing newline

    out_dir.mkdir(parents=True, exist_ok=True)
    index_path.write_text("\n".join(idx), encoding="utf-8")

    print(f"OK: created {len(written)} pagine in {companies_dir}")
    print(f"OK: index written to {index_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())