#!/usr/bin/env python3
"""
Render card images from an Excel workbook + Jinja2 HTML templates.

- Reads multiple sheets from Cards.xlsx
- For each row, renders HTML via the appropriate template
- Uses headless Chrome (Selenium) to screenshot the #card element
- Crops and saves PNGs to disk

Assumptions:
- Each HTML template contains an element with id="card"
- Excel has at least: Name, Deck, Effect (Effect optional depending on template)
- Some sheets are skipped (Adversary Backs, Land Cards)
- "Adversaries" renders twice: front + back (back filename gets "back-" prefix)
"""

from __future__ import annotations

from jinja2 import Template

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import argparse
import base64
import copy
import io
import os
import re
import time
import urllib.parse
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
from PIL import Image


# -----------------------
# Card sheet -> template(s)
# -----------------------
# You can change these paths if you keep templates elsewhere.
CARD_SHEETS: Dict[str, Dict] = {
    # sheet_name: { "outputs": [(card_type_name, template_path, filename_prefix), ...] }
    "Lanes": {"outputs": [("lane", "card_basic.html", "")]},
    "Balls": {"outputs": [("ball", "card_ball.html", "")]},
    "Stamps": {"outputs": [("resource", "card_ball.html", "")]},
    "Trials": {"outputs": [("trial", "card_basic.html", "")]},
    "Judgement": {"outputs": [("judgement", "card_basic.html", "")]},
}

# -----------------------
# Utilities
# -----------------------
_slug_re = re.compile(r"[^a-z0-9]+")


def slugify(s: str) -> str:
    s = s.strip().lower()
    s = _slug_re.sub("-", s)
    s = s.strip("-")
    return s or "unnamed"


def is_nonempty_str(x) -> bool:
    return isinstance(x, str) and x.strip() != ""


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename columns so that spaces are removed: "Land Type" -> "LandType"
    Also strips surrounding whitespace.
    """
    rename = {}
    for c in df.columns:
        if isinstance(c, str):
            rename[c] = c.strip().replace(" ", "")
        else:
            rename[c] = c
    return df.rename(columns=rename)


def nan_to_none(x):
    # pandas uses NaN for missing values; templates usually prefer None/"".
    if pd.isna(x):
        return None
    return x


# -----------------------
# Text formatting hook
# -----------------------

TOKEN_ELM = '<img class="sym" src="./components/tokens/{}"></img>'

format_map = {}

def format_text(text: Optional[str]) -> Optional[str]:
    if not text:
        return text

    for key, value in format_map.items():
        text = text.replace(key, value)

    return text

# -----------------------
# Template handling
# -----------------------
def load_templates(template_dir: Path) -> Dict[str, Template]:
    """
    Loads all templates referenced by CARD_SHEETS into memory.
    Keyed by filename (e.g. "card_island.html").
    """
    templates: Dict[str, Template] = {}
    needed = set()
    for cfg in CARD_SHEETS.values():
        for _, template_name, _ in cfg["outputs"]:
            needed.add(template_name)

    for template_name in needed:
        p = template_dir / template_name
        if not p.exists():
            raise FileNotFoundError(f"Template not found: {p}")
        with p.open("r", encoding="utf-8") as f:
            templates[template_name] = Template(f.read())

    return templates


def render_card(template: Template, row: Dict) -> str:
    # Increase size if no </br> found in text AND text is not too long
    if row["Effect"] and "</br>" not in row["Effect"] and len(row["Effect"].split(" ")) < 8:
        row["Effect"] = f'<div class="large">{row["Effect"]}</div>'

    # Apply formatting rules
    if "Effect" in row:
        row["Effect"] = format_text(row.get("Effect"))
    return template.render(row)


# -----------------------
# Selenium screenshotting
# -----------------------
def build_driver(device_scale_factor: float) -> webdriver.Chrome:
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--hide-scrollbars")
    options.add_argument("--force-device-scale-factor=%s" % device_scale_factor)
    options.add_argument("--window-size=1600,1200")
    options.add_argument("--allow-file-access-from-files")
    options.add_argument("--disable-web-security")
    return webdriver.Chrome(options=options)

def wait_dom_ready(driver: webdriver.Chrome, timeout: int = 10) -> None:
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

def print_card_element(
    driver: webdriver.Chrome,
    html: str,
    out_path: Path,
    device_scale_factor: float,
    extra_sleep_sec: float = 0.10,
) -> None:

    tmp_html = Path(__file__).parent.resolve() / "_tmp_render.html"
    tmp_html.write_text(html, encoding="utf-8")

    encoded_html = urllib.parse.quote(html)
    driver.get(tmp_html.resolve().as_uri())

    wait_dom_ready(driver, timeout=10)

    # Ensure the element exists
    card_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "card"))
    )

    # Give the browser a beat to finish layout/fonts/images
    if extra_sleep_sec and extra_sleep_sec > 0:
        time.sleep(extra_sleep_sec)

    # Selenium reports location/size in CSS pixels.
    loc = card_element.location
    size = card_element.size

    # Screenshot is in device pixels -> multiply by scale factor.
    scale = float(device_scale_factor)

    # Take full-page screenshot
    png_bytes = driver.get_screenshot_as_png()
    image = Image.open(io.BytesIO(png_bytes))

    left = int(loc["x"] * scale)
    top = int(loc["y"] * scale)
    right = int((loc["x"] + size["width"]) * scale)
    bottom = int((loc["y"] + size["height"]) * scale)

    # Crop and save
    card_image = image.crop((left, top, right, bottom))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    card_image.save(out_path, format="PNG", dpi=(300, 300))

# -----------------------
# Main flow
# -----------------------
def iter_rows(df: pd.DataFrame) -> Iterable[Dict]:
    """
    Yield row dicts with NaNs converted to None.
    """
    for _, row in df.iterrows():
        d = {k: nan_to_none(v) for k, v in row.to_dict().items()}
        yield d


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", default="Cards.xlsx", help="Path to workbook")
    parser.add_argument(
        "--templates",
        default=".",
        help="Directory containing HTML templates (e.g. card_island.html)",
    )
    parser.add_argument("--out", default="cards", help="Output directory")
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing PNG files (default: False)",
    )
    parser.add_argument(
        "--scale",
        type=float,
        default=4.0,
        help="Chrome device scale factor (used to get sharper screenshots)",
    )
    parser.add_argument(
        "--sleep",
        type=float,
        default=0.10,
        help="Extra delay after render (seconds) before screenshot",
    )
    args = parser.parse_args()

    excel_path = Path(args.excel)
    template_dir = Path(args.templates)
    out_dir = Path(args.out)
    overwrite = bool(args.overwrite)
    device_scale_factor = float(args.scale)

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    templates = load_templates(template_dir)

    # Load workbook once
    xls = pd.ExcelFile(excel_path)

    # Track per-deck numbering (deck -> count)
    counts: Dict[str, int] = {}

    driver = None
    try:
        driver = build_driver(device_scale_factor=device_scale_factor)

        # Iterate configured sheets in a stable order
        for sheet_name, cfg in CARD_SHEETS.items():
            if sheet_name not in xls.sheet_names:
                # Ignore missing sheets (common during early development)
                continue

            df = pd.read_excel(xls, sheet_name=sheet_name)
            df = normalize_columns(df)

            for row in iter_rows(df):
                # Must have Deck to route output
                deck_raw = row.get("Deck")
                if not is_nonempty_str(deck_raw):
                    continue
                deck = slugify(deck_raw)

                # Name for slugging (optional)
                name_raw = row.get("Name")
                name_slug = slugify(name_raw) if is_nonempty_str(name_raw) else None

                # Default per-deck counter
                if deck not in counts:
                    counts[deck] = 0
                counts[deck] += 1
                deck_index = counts[deck]

                # Render one or more outputs for this sheet
                for card_type, template_name, prefix in cfg["outputs"]:
                    row_for_render = copy.deepcopy(row)
                    row_for_render["CardType"] = card_type

                    # Optional: if your templates expect a normalized "NameSlug"
                    if name_slug:
                        row_for_render["NameSlug"] = name_slug

                    html = render_card(templates[template_name], row_for_render)

                    # Choose filename:
                    # - if name exists: cards/<deck>/<prefix><name>.png
                    # - else:          cards/<deck>/<prefix><index>.png
                    if name_slug:
                        filename = f"{prefix}{name_slug}.png"
                    else:
                        filename = f"{prefix}{deck_index}.png"

                    out_path = out_dir / deck / filename

                    if out_path.exists() and not overwrite:
                        continue

                    print_card_element(
                        driver=driver,
                        html=html,
                        out_path=out_path,
                        device_scale_factor=device_scale_factor,
                        extra_sleep_sec=float(args.sleep),
                    )

                    print(out_path.as_posix())

    finally:
        if driver is not None:
            driver.quit()

def create_pcio_decks(
    excel: str = "Cards.xlsx",
    out: str = "decks",
    project_name: str = "",
) -> None:
    """
    Create one CSV per deck from the Excel workbook.

    For every sheet in the workbook:
    - Read rows
    - Group rows by Deck
    - Write one CSV per deck

    CSV columns:
    - label
    - Deck
    - image
    - item-count
    - item-key

    Output filename:
    - decks/<deckslug>_<sheetslug>.csv

    Note:
    - I use the sheet slug as the second part of the filename so decks with the
      same name coming from different sheets do not overwrite each other.
    """
    excel_path = Path(excel)
    out_dir = Path(out)

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    if not project_name or not str(project_name).strip():
        raise ValueError("project_name must be a non-empty string")

    xls = pd.ExcelFile(excel_path)
    out_dir.mkdir(parents=True, exist_ok=True)

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        df = normalize_columns(df)

        if "Deck" not in df.columns:
            continue

        sheet_slug = slugify(sheet_name)
        decks: Dict[str, List[Dict]] = {}

        for row in iter_rows(df):
            deck_raw = row.get("Deck")
            name_raw = row.get("Name")

            if not is_nonempty_str(deck_raw):
                continue
            if not is_nonempty_str(name_raw):
                continue

            deck_slug = slugify(deck_raw)
            name_slug = slugify(name_raw)

            record = {
                "label": name_slug,
                "Deck": deck_slug,
                "image": (
                    f"https://api.bitbucket.org/2.0/repositories/"
                    f"skaffel/{project_name}/src/main/cards/{deck_slug}/{name_slug}.png"
                ),
                "item-count": row.get("Copies"),
                "item-key": name_slug,
            }

            if deck_slug not in decks:
                decks[deck_slug] = []
            decks[deck_slug].append(record)

        for deck_slug, records in decks.items():
            out_path = out_dir / f"{deck_slug}_{sheet_slug}.csv"
            pd.DataFrame(
                records,
                columns=["label", "Deck", "image", "item-count", "item-key"],
            ).to_csv(out_path, index=False)
            print(out_path.as_posix())

if __name__ == "__main__":
    main()
    create_pcio_decks(project_name="ballers")
