#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Generate a GitHub Pages (Jekyll) site from Google Forms responses exported to Excel.

Input:
  data/schede.xlsx

Output (ready to publish from GitHub Pages -> Deploy from branch -> /docs):
  docs/index.md
  docs/companies/<slug>.md
  docs/_config.yml

Notes:
- Jekyll only processes Markdown into HTML when front matter is present.
  This script writes front matter in every page.
- Do NOT create docs/.nojekyll (it would disable Jekyll).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Dict, List, Tuple

import pandas as pd


# --------------------------
# Helpers
# --------------------------

def norm(s: str) -> str:
    """Normalise strings for column matching."""
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("–", "-").replace("—", "-")
    return s


def find_col(columns: List[str], patterns: List[str]) -> Optional[str]:
    """
    Find the first column whose normalised name matches any regex in patterns.
    Returns the original column name or None.
    """
    ncols = {c: norm(c) for c in columns}
    for pat in patterns:
        rx = re.compile(pat, re.IGNORECASE)
        for orig, n in ncols.items():
            if rx.search(n):
                return orig
    return None


def clean_text(x) -> str:
    """Clean cell text for Markdown."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s


def slugify(name: str) -> str:
    """
    Slugify for filenames/URLs:
    - lowercase
    - spaces -> hyphens
    - remove punctuation
    """
    s = (name or "").strip().lower()
    s = s.replace("&", " and ")
    # Keep unicode letters/numbers/underscore/hyphen/space
    s = re.sub(r"[^\w\s-]", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"-{2,}", "-", s)
    s = s.strip("-")
    return s or "azienda"


def ensure_unique_slug(slug: str, used: Dict[str, int]) -> str:
    """
    Ensure unique slugs: if already used, append -2, -3, ...
    """
    if slug not in used:
        used[slug] = 1
        return slug
    used[slug] += 1
    return f"{slug}-{used[slug]}"


def escape_yaml(s: str) -> str:
    """Escape a string for safe inclusion in YAML double quotes."""
    return (s or "").replace("\\", "\\\\").replace('"', '\\"')


def front_matter(title: Optional[str] = None, permalink: Optional[str] = None) -> str:
    """
    Minimal Jekyll front matter that triggers Markdown processing.
    Add optional title/permalink.
    """
    fm = ["---"]
    if title:
        fm.append(f'title: "{escape_yaml(title)}"')
    if permalink:
        fm.append(f"permalink: {permalink}")
    fm.append("---\n")
    return "\n".join(fm)


# --------------------------
# Main
# --------------------------

def main() -> int:
    # Fixed input/output locations
    in_path = Path("data") / "schede.xlsx"

    # Publish from /docs in GitHub Pages settings
    out_dir = Path("docs")
    companies_dir = out_dir / "companies"
    index_path = out_dir / "index.md"
    config_path = out_dir / "_config.yml"

    companies_dir.mkdir(parents=True, exist_ok=True)

    if not in_path.exists():
        raise FileNotFoundError(f"Input file not found: {in_path}")

    # Read Excel
    df = pd.read_excel(in_path, engine="openpyxl")
    cols = list(df.columns)

    # Detect columns (robust to small wording changes)
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
            r"\bcosa\b.*\boffrite\b",
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
        print("Available columns in the spreadsheet:")
        for c in cols:
            print(" -", c)
        print("Continuing: missing fields will be left blank.\n")

    # Build records
    records: List[Dict[str, str]] = []
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

    # Write Jekyll config (pretty URLs)
    out_dir.mkdir(parents=True, exist_ok=True)
    config_path.write_text("permalink: pretty\n", encoding="utf-8")

    # Write one page per company
    used_slugs: Dict[str, int] = {}
    written: List[Tuple[str, str]] = []  # (company name, slug)

    for r in records:
        name = r["Nome azienda"]
        slug = slugify(name)
        slug = ensure_unique_slug(slug, used_slugs)

        page_path = companies_dir / f"{slug}.md"
        permalink = f"/companies/{slug}/"

        md: List[str] = []
        md.append(front_matter(title=name, permalink=permalink))
        md.append(f"# {name}\n")

        md.append("## Descrizione azienda\n")
        md.append((r["Descrizione azienda"] or "-") + "\n")

        md.append("## Cosa cercate/offrite\n")
        md.append((r["Cosa cercate/offrite"] or "-") + "\n")

        md.append("## Per che livello di studenti\n")
        md.append((r["Per che livello di studenti"] or "-") + "\n")

        md.append("## Indirizzo e contatto\n")
        md.append((r["Indirizzo e contatto"] or "-") + "\n")

        page_path.write_text("\n".join(md), encoding="utf-8")
        written.append((name, slug))

    # Write index page (root permalink)
    idx: List[str] = []
    idx.append(front_matter(title="Aziende partecipanti – Career Day", permalink="/"))
    idx.append("## Aziende partecipanti\n\n# Career Day Dipartimento di Informatica, Università di Torino\n")
    idx.append("## 17 marzo 2026\n")
    idx.append("Clicca per aprire la scheda:\n")

    for name, slug in written:
        idx.append(f"- [{name}](companies/{slug}/)")

    idx.append("")
    index_path.write_text("\n".join(idx), encoding="utf-8")

    print(f"OK: created {len(written)} pages in {companies_dir}")
    print(f"OK: index written to {index_path}")
    print(f"OK: Jekyll config written to {config_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())