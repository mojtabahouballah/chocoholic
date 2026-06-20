
# Script name for labbook: chocoholic_single_file_v11.py
# Title: Chocoholic single-file business calculator
# - One Python file only.
# - Embedded ingredient and recipe defaults are decoded from base64 CSV.
# - Ingredient table follows the Excel model.
# - Recipe builder supports add, edit, and remove recipes.
# - Ingredients are added with a simple + button.
# - Profit summary calculates cost per piece, suggested price, selling price, and profit.

from pathlib import Path
import base64
import io
import re
import hmac

import numpy as np
import pandas as pd
import streamlit as st


APP_VERSION = "v15 GitHub-safe login version"
DEFAULT_EXCHANGE_RATE = 90000.0

# Embedded defaults. Do not edit these long strings manually.
# Use the website tables to edit data after the app starts.
INGREDIENTS_CSV_B64 = "SW5ncmVkaWVudCAvIHByb2R1Y3QsU291cmNlLFBhY2thZ2Ugc2l6ZSBwZXIgaXRlbSxVbml0LE51bWJlciBib3VnaHQsVG90YWwgcHJpY2UgTEJQLFRvdGFsIHByaWNlIFVTRCxOb3RlcwpQb3BwaW5zIGNob2NvIGJpdHMsQWwgTW9raHRhciByZWNlaXB0LDM1MC4wLGcsMS4wLDMyOTk5NC4wLDMuNjY2NiwKRmxvdXIsQWwgTW9raHRhciByZWNlaXB0LDkwMC4wLGcsNC4wLDQwMDAwMC4wLDQuNDQ0NDQ0NDQ0NDQ0NDQ1LApOdXRlbGxhLEFsIE1va2h0YXIgcmVjZWlwdCw3NTAuMCxnLDEuMCw3NjUwMDAuMCw4LjUsCkdhbmRvdXIgNTU1IGJpc2N1aXQsQWwgTW9raHRhciByZWNlaXB0LDUyMC4wLGcsMy4wLDg0MDAwMC4wLDkuMzMzMzMzMzMzMzMzMzM0LApDYWpvdSxBbCBNb2todGFyIHJlY2VpcHQsMTAwLjAsZywyLjAsMTk5OTgwLjAsMi4yMjIsCkFsbW9uZCxBbCBNb2todGFyIHJlY2VpcHQsMTAwLjAsZywyLjAsMjM5OTQwLjAsMi42NjYsCk1hZ2lrIHdoaXBwaW5nIGNyZWFtLEFsIE1va2h0YXIgcmVjZWlwdCwxMDAwLjAsbWwsMS4wLDMxMDAwMC4wLDMuNDQ0NDQ0NDQ0NDQ0NDQ0NiwKTWlsayAxIEwsQWwgTW9raHRhciByZWNlaXB0LDEuMCxMLDEuMCwxNzAwMDAuMCwxLjg4ODg4ODg4ODg4ODg4ODgsCkNhbmRpYSBtaWxrLEFsIE1va2h0YXIgcmVjZWlwdCwxLjAsTCwxLjAsNjQwMDAwLjAsNy4xMTExMTExMTExMTExMTEsClB1Y2sgY29va2luZyBjcmVhbSxBbCBNb2todGFyIHJlY2VpcHQsMS4wLEwsMS4wLDU5OTk5NC4wLDYuNjY2NiwKTHVycGFrIGJ1dHRlcixBbCBNb2todGFyIHJlY2VpcHQsNDAwLjAsZyw3LjAsMzQzMDAwMC4wLDM4LjExMTExMTExMTExMTExNCwKRWdnIGNhcnRvbixBbCBNb2todGFyIHJlY2VpcHQsMzAuMCx1bml0LDEuMCwzMjk5OTQuMCwzLjY2NjYsCkRvbW8gdmFuaWxsYSBjdXN0YXJkLEFsIE1va2h0YXIgcmVjZWlwdCwzMDAuMCxnLDEuMCwyMDAwMDcuMCwyLjIyMjMsCk5lc3RsZSBjb25kZW5zZWQgbWlsayxBbCBNb2todGFyIHJlY2VpcHQsMzcwLjAsZywyLjAsNDYwMDAwLjAsNS4xMTExMTExMTExMTExMTEsCk9pbCxBbCBNb2todGFyIHJlY2VpcHQsNC4wLEwsMS4wLDgzMDAwMC4wLDkuMjIyMjIyMjIyMjIyMjIxLApCdWVubyBzcHJlYWQsTXVuY2hpZXMgY2FydCwxLjAsa2csMi4wLCwsVW5jbGVhciBVU0QgcHJpY2UgaW4gRXhjZWw6ID8/Pz8/Pz8KV2hpcHBlZCBjcmVhbSBwb3dkZXIgdmFuaWxsYSxNdW5jaGllcyBjYXJ0LDEuMCxrZywyLjAsLDE0LjAsClBpc3RhY2hpbyBzcHJlYWQsTXVuY2hpZXMgY2FydCwxLjAsa2csMi4wLCwzNy4wLApXaGl0ZSBjaG9jb2xhdGUgY2hpcHMsTXVuY2hpZXMgY2FydCw1MDAuMCxnLDEuMCwsMy41LApDaGlhIHNlZWRzLE11bmNoaWVzIGNhcnQsNDAwLjAsZywxLjAsLDQuNCwKSWNpbmcgc3VnYXIsTXVuY2hpZXMgY2FydCwxLjAsa2csMS4wLCwxLjUsCkNvcm4gZmxvdXIsTXVuY2hpZXMgY2FydCwxLjAsa2csMi4wLCwzLjAsCkJha2luZyBwb3dkZXIsTXVuY2hpZXMgY2FydCw1MDAuMCxnLDEuMCwsMi41LApWYW5pbGxhIHN1Z2FyIHBvd2RlcixNdW5jaGllcyBjYXJ0LDUwMC4wLGcsMS4wLCwyLjksClJlZCBnZWwgZm9vZCBjb2xvcmluZyxNdW5jaGllcyBjYXJ0LDIwLjAsZywxLjAsLCxVbmNsZWFyIFVTRCBwcmljZSBpbiBFeGNlbDogPz8/Pz8/Pz8/PwpSaWNlIGNyaXNweSBtaWxrIHRvcHBpbmdzLE11bmNoaWVzIGNhcnQsNTAwLjAsZywxLjAsLDMuOSwKTWlsayBjaG9jb2xhdGUgY2hpcHMsTXVuY2hpZXMgY2FydCwxLjAsa2csMS4wLCw2LjUsCk9hdCBmbGFrZXMsTXVuY2hpZXMgY2FydCwyNTAuMCxnLDQuMCwsMy42LApEYXJrIGNob2NvbGF0ZSBibG9jayA0MCUgbm8gYWRkZWQgc3VnYXIsTXVuY2hpZXMgY2FydCwyMjUuMCxnLDIuMCwsNy41LApEcmllZCBmcnVpdCBjdWJlcyxNdW5jaGllcyBjYXJ0LDMwMC4wLGcsMS4wLCw0Ljc1LApSaWNlIGNyaXNweSByYWluYm93LE11bmNoaWVzIGNhcnQsNDUwLjAsZywyLjAsLDcuOCwKU3BlY3Vsb29zIGJpc2N1aXQgY3JlYW15IHNwcmVhZCxNdW5jaGllcyBjYXJ0LDEuMCxrZywyLjAsLDE2LjAsCkJpc2N1aXQgc3ByZWFkLE11bmNoaWVzIGNhcnQsMS4wLGtnLDIuMCwsMTMuMCwKQ29jb2EgcG93ZGVyIHVuc3dlZXRlbmVkIDEwLTEyJSxNdW5jaGllcyBjYXJ0LDEuMCxrZywxLjAsLCxVbmNsZWFyIFVTRCBwcmljZSBpbiBFeGNlbDogPz8/Pz8/Pz8/PwpSZWQgdmVsdmV0IGNha2UgbWl4LE11bmNoaWVzIGNhcnQsNDAwLjAsZywxLjAsLDIuNzUsCkRhcmsgY2hvY29sYXRlIGJsb2NrIGZvciBiYWtpbmcsTXVuY2hpZXMgY2FydCwzLjAsa2csMi4wLCw0NC4wLApNaWxrIGNob2NvbGF0ZSBibG9jayBmb3IgYmFraW5nLE11bmNoaWVzIGNhcnQsMy4wLGtnLDMuMCwsNjQuNSwKV2hpdGUgc3VnYXIsTXVuY2hpZXMgY2FydCw5MDAuMCxnLDEwLjAsLDkuMCwKQnV0dGVyY3JlYW0gdmFuaWxsYSxNdW5jaGllcyBjYXJ0LDEwMC4wLGcsMS4wLCwxLjUsCkJlZWYgZ2VsYXRpbiBwb3dkZXIgaGFsYWwsTXVuY2hpZXMgY2FydCw1MC4wLGcsMi4wLCwzLjAsClNvZGl1bSBiaWNhcmJvbmF0ZSxNdW5jaGllcyBjYXJ0LDUwMC4wLGcsMS4wLCwxLjUsClBlYW51dCBidXR0ZXIgMTAwJSxNdW5jaGllcyBjYXJ0LDEuMCxrZywyLjAsLDE1LjAsClhYIC0gcmVuYW1lIHJvdyA0NyxNdW5jaGllcyBjYXJ0LDEuMCxrZywyLjAsLDE1LjAsT3JpZ2luYWwgRXhjZWwgbmFtZSB3YXMgWFguIFJlbmFtZSB0aGlzIGluZ3JlZGllbnQuClhYIC0gcmVuYW1lIHJvdyA0OCxNdW5jaGllcyBjYXJ0LDEuMCxrZywyLjAsLDE1LjAsT3JpZ2luYWwgRXhjZWwgbmFtZSB3YXMgWFguIFJlbmFtZSB0aGlzIGluZ3JlZGllbnQuClhYIC0gcmVuYW1lIHJvdyA0OSxNdW5jaGllcyBjYXJ0LDEuMCxrZywyLjAsLDE1LjAsT3JpZ2luYWwgRXhjZWwgbmFtZSB3YXMgWFguIFJlbmFtZSB0aGlzIGluZ3JlZGllbnQuClhYIC0gcmVuYW1lIHJvdyA1MCxNdW5jaGllcyBjYXJ0LDEuMCxrZywyLjAsLDE1LjAsT3JpZ2luYWwgRXhjZWwgbmFtZSB3YXMgWFguIFJlbmFtZSB0aGlzIGluZ3JlZGllbnQuClhYIC0gcmVuYW1lIHJvdyA1MSxNdW5jaGllcyBjYXJ0LDEuMCxrZywyLjAsLDE1LjAsT3JpZ2luYWwgRXhjZWwgbmFtZSB3YXMgWFguIFJlbmFtZSB0aGlzIGluZ3JlZGllbnQuClhYIC0gcmVuYW1lIHJvdyA1MixNdW5jaGllcyBjYXJ0LDEuMCxrZywyLjAsLDE1LjAsT3JpZ2luYWwgRXhjZWwgbmFtZSB3YXMgWFguIFJlbmFtZSB0aGlzIGluZ3JlZGllbnQuCkNyZWFtIGNoZWVzZSxNYW51YWwgcGxhY2Vob2xkZXIsLGcsLCwsTWlzc2luZyBwcmljZS9zaXplOyBjb21wbGV0ZSB0aGlzIHJvdyBiZWZvcmUgZmluYWwgcHJpY2luZy4KQ29va2llIGNyZWFtIC8gT3JlbyBjb29raWUgY3JlYW0sTWFudWFsIHBsYWNlaG9sZGVyLCxnLCwsLE1pc3NpbmcgcHJpY2Uvc2l6ZTsgY29tcGxldGUgdGhpcyByb3cgYmVmb3JlIGZpbmFsIHByaWNpbmcuClBpbmsgY2hvY29sYXRlLE1hbnVhbCBwbGFjZWhvbGRlciwsZywsLCxNaXNzaW5nIHByaWNlL3NpemU7IGNvbXBsZXRlIHRoaXMgcm93IGJlZm9yZSBmaW5hbCBwcmljaW5nLgpNaXggLyBmbGF2b3IgcG93ZGVyLE1hbnVhbCBwbGFjZWhvbGRlciwsZywsLCxNaXNzaW5nIHByaWNlL3NpemU7IGNvbXBsZXRlIHRoaXMgcm93IGJlZm9yZSBmaW5hbCBwcmljaW5nLgpEYXRlcyxNYW51YWwgcGxhY2Vob2xkZXIsLGcsLCwsTWlzc2luZyBwcmljZS9zaXplOyBjb21wbGV0ZSB0aGlzIHJvdyBiZWZvcmUgZmluYWwgcHJpY2luZy4KTWl4ZWQgbnV0cyxNYW51YWwgcGxhY2Vob2xkZXIsLGcsLCwsTWlzc2luZyBwcmljZS9zaXplOyBjb21wbGV0ZSB0aGlzIHJvdyBiZWZvcmUgZmluYWwgcHJpY2luZy4K"
RECIPE_SETTINGS_CSV_B64 = "cmVjaXBlLHBpZWNlc19wZXJfYmF0Y2gscGFja2FnaW5nX2Nvc3RfcGVyX3BpZWNlLGxhYm9yX2Nvc3RfcGVyX2JhdGNoLG90aGVyX2Nvc3RfcGVyX2JhdGNoLHdhc3RlX3BlcmNlbnQsc2VsbF9wcmljZV9wZXJfcGllY2UscGxhbm5lZF9iYXRjaGVzX3NvbGQKRXhjZWwgaW1wb3J0ZWQgcmVjaXBlLDI2LjAsMC4wLDAuMCwwLjAsMC4wLDAuMCwxLjAKRmVjaHdldCBSYW1sIC0gY2hlY2sgbmFtZSwxNi4wLDAuMCwwLjAsMC4wLDAuMCwwLjAsMS4wCkNob2NvbGF0ZSBSYW1sIC0gY2hlY2sgbmFtZSwxMC4wLDAuMCwwLjAsMC4wLDAuMCwwLjAsMS4wCkhlYWx0aHkgU25pY2tlcnMsMTAuMCwwLjAsMC4wLDAuMCwwLjAsMC4wLDEuMAo="
RECIPE_LINES_CSV_B64 = "cmVjaXBlLGluZ3JlZGllbnQscXVhbnRpdHlfdXNlZApFeGNlbCBpbXBvcnRlZCByZWNpcGUsRmxvdXIsMzcwLjAKRXhjZWwgaW1wb3J0ZWQgcmVjaXBlLE51dGVsbGEsNTIwLjAKRXhjZWwgaW1wb3J0ZWQgcmVjaXBlLEx1cnBhayBidXR0ZXIsMjAwLjAKRXhjZWwgaW1wb3J0ZWQgcmVjaXBlLEVnZyBjYXJ0b24sMi4wCkV4Y2VsIGltcG9ydGVkIHJlY2lwZSxDb3JuIGZsb3VyLDEwLjAKRXhjZWwgaW1wb3J0ZWQgcmVjaXBlLEJha2luZyBwb3dkZXIsOC4wCkV4Y2VsIGltcG9ydGVkIHJlY2lwZSxWYW5pbGxhIHN1Z2FyIHBvd2RlciwxMC4wCkV4Y2VsIGltcG9ydGVkIHJlY2lwZSxNaWxrIGNob2NvbGF0ZSBjaGlwcywyMDAuMApFeGNlbCBpbXBvcnRlZCByZWNpcGUsV2hpdGUgc3VnYXIsODguMApFeGNlbCBpbXBvcnRlZCByZWNpcGUsU29kaXVtIGJpY2FyYm9uYXRlLDIuMApGZWNod2V0IFJhbWwgLSBjaGVjayBuYW1lLE1hZ2lrIHdoaXBwaW5nIGNyZWFtLDIwMC4wCkZlY2h3ZXQgUmFtbCAtIGNoZWNrIG5hbWUsTmVzdGxlIGNvbmRlbnNlZCBtaWxrLDE0MC4wCkZlY2h3ZXQgUmFtbCAtIGNoZWNrIG5hbWUsQ3JlYW0gY2hlZXNlLDMwMC4wCkNob2NvbGF0ZSBSYW1sIC0gY2hlY2sgbmFtZSxDb29raWUgY3JlYW0gLyBPcmVvIGNvb2tpZSBjcmVhbSwyMDAuMApDaG9jb2xhdGUgUmFtbCAtIGNoZWNrIG5hbWUsTWlsayBjaG9jb2xhdGUgYmxvY2sgZm9yIGJha2luZywyMDAuMApDaG9jb2xhdGUgUmFtbCAtIGNoZWNrIG5hbWUsUGluayBjaG9jb2xhdGUsNzUuMApDaG9jb2xhdGUgUmFtbCAtIGNoZWNrIG5hbWUsTWl4IC8gZmxhdm9yIHBvd2RlciwxLjAKSGVhbHRoeSBTbmlja2VycyxEYXRlcywzNTAuMApIZWFsdGh5IFNuaWNrZXJzLE1peGVkIG51dHMsMTAwLjAKSGVhbHRoeSBTbmlja2VycyxEYXJrIGNob2NvbGF0ZSBibG9jayA0MCUgbm8gYWRkZWQgc3VnYXIsMjUuMAo="

DATA_DIR = Path("chocoholic_data_v11")
DATA_DIR.mkdir(exist_ok=True)

INGREDIENTS_FILE = DATA_DIR / "ingredients.csv"
RECIPE_SETTINGS_FILE = DATA_DIR / "recipe_settings.csv"
RECIPE_LINES_FILE = DATA_DIR / "recipe_lines.csv"

INGREDIENT_INPUT_COLUMNS = [
    "Ingredient / product",
    "Source",
    "Package size per item",
    "Unit",
    "Number bought",
    "Total price LBP",
    "Total price USD",
    "Notes",
]

INGREDIENT_DISPLAY_COLUMNS = [
    "Ingredient / product",
    "Source",
    "Package size per item",
    "Unit",
    "Number bought",
    "Total price LBP",
    "Total price USD",
    "Total bought in base unit",
    "Base unit",
    "Price per kg/L/unit $",
    "Cost per g/ml/unit $",
    "Notes",
    "Status",
]

RECIPE_SETTINGS_COLUMNS = [
    "recipe",
    "pieces_per_batch",
    "packaging_cost_per_piece",
    "labor_cost_per_batch",
    "other_cost_per_batch",
    "waste_percent",
    "sell_price_per_piece",
    "planned_batches_sold",
]

RECIPE_LINE_COLUMNS = [
    "recipe",
    "ingredient",
    "quantity_used",
]


def embedded_csv_to_df(b64_text, columns):
    csv_text = base64.b64decode(b64_text.encode("ascii")).decode("utf-8")
    df = pd.read_csv(io.StringIO(csv_text))

    for col in columns:
        if col not in df.columns:
            df[col] = None

    return df[columns].copy()


def default_ingredients_df():
    return embedded_csv_to_df(INGREDIENTS_CSV_B64, INGREDIENT_INPUT_COLUMNS)


def default_recipe_settings_df():
    return embedded_csv_to_df(RECIPE_SETTINGS_CSV_B64, RECIPE_SETTINGS_COLUMNS)


def default_recipe_lines_df():
    return embedded_csv_to_df(RECIPE_LINES_CSV_B64, RECIPE_LINE_COLUMNS)


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_number(value, default=0.0):
    try:
        if pd.isna(value) or value == "":
            return default
        return float(value)
    except Exception:
        return default


def numeric_or_nan(value):
    try:
        if pd.isna(value) or value == "":
            return np.nan
        return float(value)
    except Exception:
        return np.nan


def safe_key(text):
    text = clean_text(text)
    text = re.sub(r"[^A-Za-z0-9_]+", "_", text)
    if text == "":
        return "empty"
    return text[:80]


def save_csv(df, path, columns):
    df = df[columns].copy()
    df.to_csv(path, index=False)


def load_csv(path, columns):
    if path.exists():
        df = pd.read_csv(path)
    else:
        df = pd.DataFrame(columns=columns)

    for col in columns:
        if col not in df.columns:
            df[col] = None

    return df[columns].copy()


def initialize_files():
    if not INGREDIENTS_FILE.exists():
        save_csv(default_ingredients_df(), INGREDIENTS_FILE, INGREDIENT_INPUT_COLUMNS)

    if not RECIPE_SETTINGS_FILE.exists():
        save_csv(default_recipe_settings_df(), RECIPE_SETTINGS_FILE, RECIPE_SETTINGS_COLUMNS)

    if not RECIPE_LINES_FILE.exists():
        save_csv(default_recipe_lines_df(), RECIPE_LINES_FILE, RECIPE_LINE_COLUMNS)


def reset_all_data():
    save_csv(default_ingredients_df(), INGREDIENTS_FILE, INGREDIENT_INPUT_COLUMNS)
    save_csv(default_recipe_settings_df(), RECIPE_SETTINGS_FILE, RECIPE_SETTINGS_COLUMNS)
    save_csv(default_recipe_lines_df(), RECIPE_LINES_FILE, RECIPE_LINE_COLUMNS)

    for key in list(st.session_state.keys()):
        if key.startswith("recipe_cart_") or key.startswith("pending_ingredient_"):
            del st.session_state[key]


def package_to_base_unit(unit):
    unit = clean_text(unit).lower()

    if unit in ["kg", "g"]:
        return "g"

    if unit in ["l", "liter", "litre", "ml"]:
        return "ml"

    if unit in ["unit", "piece", "pc"]:
        return "unit"

    return "g"


def unit_conversion_factor(unit):
    unit = clean_text(unit).lower()

    if unit == "kg":
        return 1000.0

    if unit in ["l", "liter", "litre"]:
        return 1000.0

    return 1.0


def compute_ingredient_table(ingredients_input, exchange_rate):
    df = ingredients_input.copy()

    for col in INGREDIENT_INPUT_COLUMNS:
        if col not in df.columns:
            df[col] = None

    for col in ["Ingredient / product", "Source", "Unit", "Notes"]:
        df[col] = df[col].apply(clean_text)

    for col in ["Package size per item", "Number bought", "Total price LBP", "Total price USD"]:
        df[col] = df[col].apply(numeric_or_nan)

    df["Base unit"] = df["Unit"].apply(package_to_base_unit)
    df["unit_conversion"] = df["Unit"].apply(unit_conversion_factor)

    df["Total bought in base unit"] = df.apply(
        lambda row: row["Package size per item"] * row["Number bought"] * row["unit_conversion"]
        if pd.notna(row["Package size per item"]) and pd.notna(row["Number bought"])
        else np.nan,
        axis=1,
    )

    df["final_total_price_usd"] = df.apply(
        lambda row: row["Total price LBP"] / exchange_rate
        if pd.notna(row["Total price LBP"]) and row["Total price LBP"] > 0 and exchange_rate > 0
        else (
            row["Total price USD"]
            if pd.notna(row["Total price USD"]) and row["Total price USD"] > 0
            else np.nan
        ),
        axis=1,
    )

    df["Cost per g/ml/unit $"] = df.apply(
        lambda row: row["final_total_price_usd"] / row["Total bought in base unit"]
        if pd.notna(row["final_total_price_usd"])
        and pd.notna(row["Total bought in base unit"])
        and row["Total bought in base unit"] > 0
        else 0.0,
        axis=1,
    )

    df["Price per kg/L/unit $"] = df.apply(
        lambda row: row["Cost per g/ml/unit $"] * 1000
        if row["Base unit"] in ["g", "ml"]
        else row["Cost per g/ml/unit $"],
        axis=1,
    )

    def status(row):
        if clean_text(row["Ingredient / product"]) == "":
            return "EMPTY"

        if pd.isna(row["Package size per item"]) or clean_number(row["Package size per item"]) <= 0:
            return "MISSING SIZE"

        if pd.isna(row["Number bought"]) or clean_number(row["Number bought"]) <= 0:
            return "MISSING QUANTITY"

        if clean_number(row["Cost per g/ml/unit $"]) <= 0:
            return "MISSING PRICE"

        return "OK"

    df["Status"] = df.apply(status, axis=1)

    return df[INGREDIENT_DISPLAY_COLUMNS].copy()


def compute_recipe_summary(ingredients_input, recipe_lines, recipe_settings, exchange_rate, target_margin_percent):
    ingredients = compute_ingredient_table(ingredients_input, exchange_rate)

    ingredient_lookup = ingredients[
        [
            "Ingredient / product",
            "Source",
            "Base unit",
            "Cost per g/ml/unit $",
            "Status",
        ]
    ].rename(columns={"Ingredient / product": "ingredient"})

    lines = recipe_lines.copy()
    settings = recipe_settings.copy()

    for col in RECIPE_LINE_COLUMNS:
        if col not in lines.columns:
            lines[col] = None

    for col in RECIPE_SETTINGS_COLUMNS:
        if col not in settings.columns:
            settings[col] = None

    lines["recipe"] = lines["recipe"].apply(clean_text)
    lines["ingredient"] = lines["ingredient"].apply(clean_text)
    lines["quantity_used"] = lines["quantity_used"].apply(clean_number)

    settings["recipe"] = settings["recipe"].apply(clean_text)

    for col in RECIPE_SETTINGS_COLUMNS:
        if col != "recipe":
            settings[col] = settings[col].apply(clean_number)

    merged = lines.merge(ingredient_lookup, on="ingredient", how="left")

    if "Cost per g/ml/unit $" not in merged.columns:
        merged["Cost per g/ml/unit $"] = 0.0

    merged["Cost per g/ml/unit $"] = merged["Cost per g/ml/unit $"].fillna(0.0)
    merged["line_cost_usd"] = merged["quantity_used"] * merged["Cost per g/ml/unit $"]

    if merged.empty:
        ingredient_cost = pd.DataFrame(columns=["recipe", "ingredient_cost_per_batch"])
        missing_used = pd.DataFrame(columns=["recipe", "missing_cost_ingredients"])
    else:
        ingredient_cost = (
            merged.groupby("recipe", as_index=False)["line_cost_usd"]
            .sum()
            .rename(columns={"line_cost_usd": "ingredient_cost_per_batch"})
        )

        missing_used = (
            merged[
                (merged["quantity_used"] > 0)
                & (merged["Cost per g/ml/unit $"] <= 0)
            ]
            .groupby("recipe", as_index=False)["ingredient"]
            .apply(lambda values: ", ".join(sorted(set(values))))
            .rename(columns={"ingredient": "missing_cost_ingredients"})
        )

    summary = settings.merge(ingredient_cost, on="recipe", how="left")
    summary = summary.merge(missing_used, on="recipe", how="left")

    summary["ingredient_cost_per_batch"] = summary["ingredient_cost_per_batch"].fillna(0.0)
    summary["missing_cost_ingredients"] = summary["missing_cost_ingredients"].fillna("")

    summary["packaging_cost_per_batch"] = summary["pieces_per_batch"] * summary["packaging_cost_per_piece"]

    summary["raw_cost_per_batch"] = (
        summary["ingredient_cost_per_batch"]
        + summary["packaging_cost_per_batch"]
        + summary["labor_cost_per_batch"]
        + summary["other_cost_per_batch"]
    )

    summary["total_cost_per_batch"] = summary["raw_cost_per_batch"] * (1 + summary["waste_percent"] / 100)

    summary["cost_per_piece"] = summary.apply(
        lambda row: row["total_cost_per_batch"] / row["pieces_per_batch"]
        if row["pieces_per_batch"] > 0
        else 0.0,
        axis=1,
    )

    target_margin = max(min(target_margin_percent / 100, 0.95), 0.0)

    summary["suggested_price_per_piece"] = summary.apply(
        lambda row: row["cost_per_piece"] / (1 - target_margin)
        if row["cost_per_piece"] > 0 and target_margin < 1
        else 0.0,
        axis=1,
    )

    summary["revenue_per_batch"] = summary["sell_price_per_piece"] * summary["pieces_per_batch"]
    summary["profit_per_piece"] = summary["sell_price_per_piece"] - summary["cost_per_piece"]
    summary["profit_per_batch"] = summary["revenue_per_batch"] - summary["total_cost_per_batch"]
    summary["total_revenue"] = summary["revenue_per_batch"] * summary["planned_batches_sold"]
    summary["total_profit"] = summary["profit_per_batch"] * summary["planned_batches_sold"]

    summary["real_margin_percent"] = summary.apply(
        lambda row: row["profit_per_piece"] / row["sell_price_per_piece"] * 100
        if row["sell_price_per_piece"] > 0
        else 0.0,
        axis=1,
    )

    return summary, merged, ingredients


def delete_recipe(recipe_name, recipe_settings, recipe_lines):
    recipe_name = clean_text(recipe_name)

    new_settings = recipe_settings[recipe_settings["recipe"] != recipe_name].copy()
    new_lines = recipe_lines[recipe_lines["recipe"] != recipe_name].copy()

    save_csv(new_settings, RECIPE_SETTINGS_FILE, RECIPE_SETTINGS_COLUMNS)
    save_csv(new_lines, RECIPE_LINES_FILE, RECIPE_LINE_COLUMNS)

    for key in list(st.session_state.keys()):
        if key.startswith("recipe_cart_") or key.startswith("pending_ingredient_"):
            del st.session_state[key]


def show_ingredients_page(ingredients_input):
    st.header("1. Ingredients")

    st.write(
        "This table follows your Excel model. Edit the purchase columns; the cost columns calculate automatically."
    )

    col1, col2, col3 = st.columns([1, 1, 2])

    with col1:
        exchange_rate = st.number_input(
            "Exchange rate LBP/USD",
            min_value=1.0,
            value=DEFAULT_EXCHANGE_RATE,
            step=1000.0,
            key="ingredients_exchange_rate",
        )

    with col2:
        if st.button("Reset all data from embedded Excel import"):
            reset_all_data()
            st.success("Data reset from the embedded Excel import.")
            st.rerun()

    with col3:
        st.info("For receipt rows, fill Total price LBP. For cart/USD rows, fill Total price USD.")

    ingredient_display = compute_ingredient_table(ingredients_input, exchange_rate)

    edited = st.data_editor(
        ingredient_display,
        num_rows="dynamic",
        use_container_width=True,
        height=650,
        key="ingredients_table",
        disabled=[
            "Total bought in base unit",
            "Base unit",
            "Price per kg/L/unit $",
            "Cost per g/ml/unit $",
            "Status",
        ],
        column_config={
            "Ingredient / product": st.column_config.TextColumn("Ingredient / product", required=True),
            "Source": st.column_config.TextColumn("Source"),
            "Package size per item": st.column_config.NumberColumn("Package size per item", min_value=0.0, step=1.0),
            "Unit": st.column_config.SelectboxColumn("Unit", options=["g", "kg", "ml", "L", "unit"]),
            "Number bought": st.column_config.NumberColumn("Number bought", min_value=0.0, step=1.0),
            "Total price LBP": st.column_config.NumberColumn("Total price LBP", min_value=0.0, step=1000.0),
            "Total price USD": st.column_config.NumberColumn("Total price USD", min_value=0.0, step=0.1),
            "Total bought in base unit": st.column_config.NumberColumn("Total bought in base unit", format="%.1f"),
            "Base unit": st.column_config.TextColumn("Base unit"),
            "Price per kg/L/unit $": st.column_config.NumberColumn("Price per kg/L/unit $", format="%.2f"),
            "Cost per g/ml/unit $": st.column_config.NumberColumn("Cost per g/ml/unit $", format="%.5f"),
            "Notes": st.column_config.TextColumn("Notes"),
            "Status": st.column_config.TextColumn("Status"),
        },
    )

    if st.button("Save ingredient table"):
        to_save = edited[INGREDIENT_INPUT_COLUMNS].copy()
        to_save["Ingredient / product"] = to_save["Ingredient / product"].apply(clean_text)
        to_save = to_save[to_save["Ingredient / product"] != ""]

        if to_save["Ingredient / product"].duplicated().any():
            st.error("Some ingredient names are duplicated. Keep ingredient names unique.")
        else:
            save_csv(to_save, INGREDIENTS_FILE, INGREDIENT_INPUT_COLUMNS)
            st.success("Ingredient table saved.")
            st.rerun()

    incomplete = ingredient_display[ingredient_display["Status"] != "OK"]

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("Ingredients", len(ingredient_display))
    metric2.metric("Incomplete rows", len(incomplete))
    metric3.metric("Exchange rate", f"{exchange_rate:,.0f} LBP/USD")

    if len(incomplete) > 0:
        with st.expander("Show incomplete ingredient rows"):
            st.dataframe(incomplete, use_container_width=True)


def show_recipe_builder(ingredients_input, recipe_settings, recipe_lines):
    st.header("2. Recipe Builder")

    st.write("Add or edit a recipe. Press + beside an ingredient, enter the quantity used, then save the recipe.")

    exchange_rate = st.number_input(
        "Exchange rate used for recipe cost",
        min_value=1.0,
        value=DEFAULT_EXCHANGE_RATE,
        step=1000.0,
        key="builder_exchange_rate",
    )

    recipe_options = sorted(
        [recipe for recipe in recipe_settings["recipe"].dropna().unique() if clean_text(recipe) != ""]
    )

    mode = st.radio(
        "What do you want to do?",
        ["Edit existing recipe", "Add new recipe"],
        horizontal=True,
    )

    if mode == "Edit existing recipe" and recipe_options:
        selected_recipe = st.selectbox("Recipe", recipe_options)
    else:
        selected_recipe = st.text_input("New recipe name", placeholder="Example: Pistachio cookie")

    selected_recipe = clean_text(selected_recipe)

    if selected_recipe == "":
        st.warning("Choose an existing recipe or type a new recipe name.")
        return

    recipe_safe = safe_key(selected_recipe)
    cart_key = f"recipe_cart_{recipe_safe}"

    if cart_key not in st.session_state:
        saved_lines = recipe_lines[recipe_lines["recipe"] == selected_recipe][["ingredient", "quantity_used"]].copy()
        st.session_state[cart_key] = {
            row["ingredient"]: float(row["quantity_used"])
            for _, row in saved_lines.iterrows()
        }

    existing_setting = recipe_settings[recipe_settings["recipe"] == selected_recipe]

    if existing_setting.empty:
        current_setting = {
            "recipe": selected_recipe,
            "pieces_per_batch": 0.0,
            "packaging_cost_per_piece": 0.0,
            "labor_cost_per_batch": 0.0,
            "other_cost_per_batch": 0.0,
            "waste_percent": 0.0,
            "sell_price_per_piece": 0.0,
            "planned_batches_sold": 1.0,
        }
    else:
        current_setting = existing_setting.iloc[0].to_dict()

    if mode == "Edit existing recipe" and selected_recipe in recipe_options:
        with st.expander("Remove this recipe"):
            st.warning("This permanently removes the recipe and its ingredient quantities.")
            confirm_delete = st.checkbox(f"I confirm I want to remove {selected_recipe}")

            if st.button("Remove recipe", disabled=not confirm_delete):
                delete_recipe(selected_recipe, recipe_settings, recipe_lines)
                st.success(f"Removed recipe: {selected_recipe}")
                st.rerun()

    st.subheader("A. Recipe information")

    info1, info2, info3, info4 = st.columns(4)

    with info1:
        pieces_per_batch = st.number_input(
            "Pieces obtained from this recipe",
            min_value=0.0,
            value=float(clean_number(current_setting.get("pieces_per_batch", 0.0))),
            step=1.0,
        )

    with info2:
        packaging_cost_per_piece = st.number_input(
            "Packaging $/piece",
            min_value=0.0,
            value=float(clean_number(current_setting.get("packaging_cost_per_piece", 0.0))),
            step=0.01,
        )

    with info3:
        labor_cost_per_batch = st.number_input(
            "Labor $/batch",
            min_value=0.0,
            value=float(clean_number(current_setting.get("labor_cost_per_batch", 0.0))),
            step=0.5,
        )

    with info4:
        other_cost_per_batch = st.number_input(
            "Other cost $/batch",
            min_value=0.0,
            value=float(clean_number(current_setting.get("other_cost_per_batch", 0.0))),
            step=0.5,
        )

    info5, info6 = st.columns(2)

    with info5:
        waste_percent = st.number_input(
            "Waste/safety % added to cost",
            min_value=0.0,
            value=float(clean_number(current_setting.get("waste_percent", 0.0))),
            step=1.0,
        )

    with info6:
        desired_margin_percent = st.number_input(
            "Desired profit margin for suggested price (%)",
            min_value=0.0,
            max_value=95.0,
            value=50.0,
            step=5.0,
            help="Only used to suggest a price. You can manually choose your real selling price.",
        )

    st.divider()

    st.subheader("B. Add ingredients with +")

    ingredients_for_calc = compute_ingredient_table(ingredients_input, exchange_rate)

    catalog = ingredients_for_calc[
        [
            "Ingredient / product",
            "Source",
            "Base unit",
            "Cost per g/ml/unit $",
            "Status",
        ]
    ].copy()

    search = st.text_input("Search ingredient", placeholder="Example: butter, sugar, chocolate")

    if search.strip():
        query = search.strip().lower()
        catalog = catalog[
            catalog["Ingredient / product"].str.lower().str.contains(query, na=False)
            | catalog["Source"].str.lower().str.contains(query, na=False)
        ]

    catalog = catalog.sort_values(["Source", "Ingredient / product"]).head(50)

    header = st.columns([0.5, 3, 1.6, 0.9, 1.2, 1.2])
    header[0].markdown("**Add**")
    header[1].markdown("**Ingredient**")
    header[2].markdown("**Source**")
    header[3].markdown("**Unit**")
    header[4].markdown("**Cost/unit**")
    header[5].markdown("**Status**")

    pending_key = f"pending_ingredient_{recipe_safe}"

    for idx, row in catalog.iterrows():
        ingredient_name = row["Ingredient / product"]
        row_key = safe_key(f"{selected_recipe}_{ingredient_name}_{idx}")
        columns = st.columns([0.5, 3, 1.6, 0.9, 1.2, 1.2])

        if columns[0].button("+", key=f"plus_{row_key}"):
            st.session_state[pending_key] = ingredient_name

        columns[1].write(ingredient_name)
        columns[2].write(row["Source"])
        columns[3].write(row["Base unit"])
        columns[4].write(f"{row['Cost per g/ml/unit $']:.5f} $")
        columns[5].write(row["Status"])

    if pending_key in st.session_state:
        pending = st.session_state[pending_key]

        pending_info = ingredients_for_calc[ingredients_for_calc["Ingredient / product"] == pending]

        unit = pending_info.iloc[0]["Base unit"] if not pending_info.empty else "g/ml/unit"

        st.info(f"Adding: {pending}")

        add1, add2, add3 = st.columns([1, 1, 3])

        with add1:
            qty_used = st.number_input(
                f"Quantity used ({unit})",
                min_value=0.0,
                value=float(st.session_state[cart_key].get(pending, 0.0)),
                step=1.0,
                key=f"qty_{safe_key(selected_recipe)}_{safe_key(pending)}",
            )

        with add2:
            if st.button("Confirm add/update"):
                if qty_used > 0:
                    st.session_state[cart_key][pending] = float(qty_used)
                elif pending in st.session_state[cart_key]:
                    del st.session_state[cart_key][pending]

                del st.session_state[pending_key]
                st.rerun()

        with add3:
            if st.button("Cancel"):
                del st.session_state[pending_key]
                st.rerun()

    st.divider()

    st.subheader("C. Current recipe")

    cart = st.session_state[cart_key]

    if not cart:
        st.warning("No ingredients added yet.")
        current_recipe_df = pd.DataFrame(
            columns=[
                "ingredient",
                "Source",
                "Base unit",
                "quantity_used",
                "Cost per g/ml/unit $",
                "line_cost_usd",
                "Status",
            ]
        )
    else:
        cart_df = pd.DataFrame(
            [{"ingredient": name, "quantity_used": qty} for name, qty in cart.items()]
        )

        current_recipe_df = cart_df.merge(
            ingredients_for_calc[
                [
                    "Ingredient / product",
                    "Source",
                    "Base unit",
                    "Cost per g/ml/unit $",
                    "Status",
                ]
            ].rename(columns={"Ingredient / product": "ingredient"}),
            on="ingredient",
            how="left",
        )

        current_recipe_df["Cost per g/ml/unit $"] = current_recipe_df["Cost per g/ml/unit $"].fillna(0.0)
        current_recipe_df["line_cost_usd"] = (
            current_recipe_df["quantity_used"] * current_recipe_df["Cost per g/ml/unit $"]
        )

        cart_header = st.columns([3, 1, 0.9, 1.2, 1.2, 1.1, 0.8])
        cart_header[0].markdown("**Ingredient**")
        cart_header[1].markdown("**Qty**")
        cart_header[2].markdown("**Unit**")
        cart_header[3].markdown("**Cost/unit**")
        cart_header[4].markdown("**Line cost**")
        cart_header[5].markdown("**Status**")
        cart_header[6].markdown("**Remove**")

        for idx, row in current_recipe_df.iterrows():
            columns = st.columns([3, 1, 0.9, 1.2, 1.2, 1.1, 0.8])
            columns[0].write(row["ingredient"])
            columns[1].write(f"{row['quantity_used']:.1f}")
            columns[2].write(row["Base unit"])
            columns[3].write(f"{row['Cost per g/ml/unit $']:.5f} $")
            columns[4].write(f"{row['line_cost_usd']:.2f} $")
            columns[5].write(row["Status"])

            remove_key = safe_key(f"remove_{selected_recipe}_{row['ingredient']}_{idx}")
            if columns[6].button("Remove", key=remove_key):
                if row["ingredient"] in st.session_state[cart_key]:
                    del st.session_state[cart_key][row["ingredient"]]
                st.rerun()

    ingredient_cost = float(current_recipe_df["line_cost_usd"].sum()) if not current_recipe_df.empty else 0.0
    packaging_cost = pieces_per_batch * packaging_cost_per_piece
    raw_cost = ingredient_cost + packaging_cost + labor_cost_per_batch + other_cost_per_batch
    total_cost = raw_cost * (1 + waste_percent / 100)

    cost_per_piece = total_cost / pieces_per_batch if pieces_per_batch > 0 else 0.0

    margin = max(min(desired_margin_percent / 100, 0.95), 0.0)
    suggested_price = cost_per_piece / (1 - margin) if cost_per_piece > 0 and margin < 1 else 0.0

    previous_sell_price = float(clean_number(current_setting.get("sell_price_per_piece", 0.0)))

    if previous_sell_price <= 0 and suggested_price > 0:
        previous_sell_price = round(suggested_price, 2)

    price_col1, price_col2 = st.columns(2)

    with price_col1:
        sell_price_per_piece = st.number_input(
            "Your selling price per piece $",
            min_value=0.0,
            value=float(previous_sell_price),
            step=0.1,
        )

    with price_col2:
        planned_batches_sold = st.number_input(
            "Planned batches sold",
            min_value=0.0,
            value=float(clean_number(current_setting.get("planned_batches_sold", 1.0))),
            step=1.0,
        )

    revenue_per_batch = sell_price_per_piece * pieces_per_batch
    profit_per_piece = sell_price_per_piece - cost_per_piece
    profit_per_batch = revenue_per_batch - total_cost
    total_profit = profit_per_batch * planned_batches_sold
    real_margin = profit_per_piece / sell_price_per_piece * 100 if sell_price_per_piece > 0 else 0.0

    st.subheader("D. Live calculation")

    metric1, metric2, metric3, metric4, metric5 = st.columns(5)
    metric1.metric("Ingredient cost", f"{ingredient_cost:.2f} $")
    metric2.metric("Total recipe cost", f"{total_cost:.2f} $")
    metric3.metric("Cost per piece", f"{cost_per_piece:.2f} $")
    metric4.metric("Suggested price", f"{suggested_price:.2f} $")
    metric5.metric("Profit per batch", f"{profit_per_batch:.2f} $")

    metric6, metric7, metric8 = st.columns(3)
    metric6.metric("Profit per piece", f"{profit_per_piece:.2f} $")
    metric7.metric("Real margin", f"{real_margin:.1f} %")
    metric8.metric("Total planned profit", f"{total_profit:.2f} $")

    if not current_recipe_df.empty:
        missing = current_recipe_df[
            (current_recipe_df["quantity_used"] > 0)
            & (current_recipe_df["Cost per g/ml/unit $"] <= 0)
        ]

        if not missing.empty:
            st.error(
                "This recipe uses ingredients with missing price/size: "
                + ", ".join(missing["ingredient"].tolist())
            )

    if st.button("Add Recipe / Save Recipe", type="primary"):
        if pieces_per_batch <= 0:
            st.error("Fill how many pieces the recipe gives.")
        elif not cart:
            st.error("Add at least one ingredient.")
        else:
            new_setting = pd.DataFrame(
                [
                    {
                        "recipe": selected_recipe,
                        "pieces_per_batch": pieces_per_batch,
                        "packaging_cost_per_piece": packaging_cost_per_piece,
                        "labor_cost_per_batch": labor_cost_per_batch,
                        "other_cost_per_batch": other_cost_per_batch,
                        "waste_percent": waste_percent,
                        "sell_price_per_piece": sell_price_per_piece,
                        "planned_batches_sold": planned_batches_sold,
                    }
                ]
            )[RECIPE_SETTINGS_COLUMNS]

            settings_without = recipe_settings[recipe_settings["recipe"] != selected_recipe].copy()
            settings_new = pd.concat([settings_without, new_setting], ignore_index=True)

            new_lines = pd.DataFrame(
                [
                    {
                        "recipe": selected_recipe,
                        "ingredient": ingredient,
                        "quantity_used": quantity,
                    }
                    for ingredient, quantity in cart.items()
                    if quantity > 0
                ]
            )[RECIPE_LINE_COLUMNS]

            lines_without = recipe_lines[recipe_lines["recipe"] != selected_recipe].copy()
            lines_new = pd.concat([lines_without, new_lines], ignore_index=True)

            save_csv(settings_new, RECIPE_SETTINGS_FILE, RECIPE_SETTINGS_COLUMNS)
            save_csv(lines_new, RECIPE_LINES_FILE, RECIPE_LINE_COLUMNS)

            st.success(f"Saved recipe: {selected_recipe}")
            st.rerun()


def show_profit_summary(ingredients_input, recipe_settings, recipe_lines):
    st.header("3. Profit Summary")

    exchange_rate = st.number_input(
        "Exchange rate used for summary",
        min_value=1.0,
        value=DEFAULT_EXCHANGE_RATE,
        step=1000.0,
        key="summary_exchange_rate",
    )

    target_margin = st.number_input(
        "Desired profit margin for suggested price (%)",
        min_value=0.0,
        max_value=95.0,
        value=50.0,
        step=5.0,
        key="summary_margin",
    )

    summary, breakdown, _ = compute_recipe_summary(
        ingredients_input,
        recipe_lines,
        recipe_settings,
        exchange_rate,
        target_margin,
    )

    if summary.empty:
        st.warning("No recipes yet.")
        return

    display = summary[
        [
            "recipe",
            "pieces_per_batch",
            "ingredient_cost_per_batch",
            "total_cost_per_batch",
            "cost_per_piece",
            "suggested_price_per_piece",
            "sell_price_per_piece",
            "planned_batches_sold",
            "revenue_per_batch",
            "profit_per_batch",
            "real_margin_percent",
            "total_profit",
            "missing_cost_ingredients",
        ]
    ].copy()

    edited = st.data_editor(
        display,
        use_container_width=True,
        height=420,
        key="summary_table",
        disabled=[
            "recipe",
            "pieces_per_batch",
            "ingredient_cost_per_batch",
            "total_cost_per_batch",
            "cost_per_piece",
            "suggested_price_per_piece",
            "revenue_per_batch",
            "profit_per_batch",
            "real_margin_percent",
            "total_profit",
            "missing_cost_ingredients",
        ],
        column_config={
            "recipe": st.column_config.TextColumn("Recipe"),
            "pieces_per_batch": st.column_config.NumberColumn("Pieces/batch"),
            "ingredient_cost_per_batch": st.column_config.NumberColumn("Ingredient cost $/batch", format="%.2f"),
            "total_cost_per_batch": st.column_config.NumberColumn("Total cost $/batch", format="%.2f"),
            "cost_per_piece": st.column_config.NumberColumn("Cost $/piece", format="%.2f"),
            "suggested_price_per_piece": st.column_config.NumberColumn("Suggested price $/piece", format="%.2f"),
            "sell_price_per_piece": st.column_config.NumberColumn("Your sell price $/piece", min_value=0.0, step=0.1),
            "planned_batches_sold": st.column_config.NumberColumn("Planned batches sold", min_value=0.0, step=1.0),
            "revenue_per_batch": st.column_config.NumberColumn("Revenue $/batch", format="%.2f"),
            "profit_per_batch": st.column_config.NumberColumn("Profit $/batch", format="%.2f"),
            "real_margin_percent": st.column_config.NumberColumn("Margin %", format="%.1f"),
            "total_profit": st.column_config.NumberColumn("Total profit $", format="%.2f"),
            "missing_cost_ingredients": st.column_config.TextColumn("Missing-cost ingredients"),
        },
    )

    if st.button("Save prices / planned batches from summary"):
        updated_settings = recipe_settings.copy()
        price_map = edited.set_index("recipe")["sell_price_per_piece"].to_dict()
        batch_map = edited.set_index("recipe")["planned_batches_sold"].to_dict()

        updated_settings["sell_price_per_piece"] = updated_settings["recipe"].map(price_map).fillna(
            updated_settings["sell_price_per_piece"]
        )

        updated_settings["planned_batches_sold"] = updated_settings["recipe"].map(batch_map).fillna(
            updated_settings["planned_batches_sold"]
        )

        save_csv(updated_settings, RECIPE_SETTINGS_FILE, RECIPE_SETTINGS_COLUMNS)
        st.success("Summary prices and planned batches saved.")
        st.rerun()

    total_revenue = float(summary["total_revenue"].sum())
    total_cost = float((summary["total_cost_per_batch"] * summary["planned_batches_sold"]).sum())
    total_profit = float(summary["total_profit"].sum())

    total1, total2, total3 = st.columns(3)
    total1.metric("Total revenue", f"{total_revenue:.2f} $")
    total2.metric("Total cost", f"{total_cost:.2f} $")
    total3.metric("Total profit", f"{total_profit:.2f} $")

    st.subheader("Remove a saved recipe")

    recipe_names = sorted(summary["recipe"].tolist())
    recipe_to_delete = st.selectbox("Recipe to remove", recipe_names)
    confirm_delete = st.checkbox(f"I confirm I want to remove {recipe_to_delete}", key="summary_delete_confirm")

    if st.button("Remove selected recipe", disabled=not confirm_delete):
        delete_recipe(recipe_to_delete, recipe_settings, recipe_lines)
        st.success(f"Removed recipe: {recipe_to_delete}")
        st.rerun()

    st.subheader("Recipe details")

    selected_detail = st.selectbox("Choose recipe to inspect", summary["recipe"].tolist())

    detail = breakdown[
        (breakdown["recipe"] == selected_detail)
        & (breakdown["quantity_used"] > 0)
    ].copy()

    if detail.empty:
        st.info("No ingredients saved for this recipe yet.")
    else:
        st.dataframe(
            detail[
                [
                    "ingredient",
                    "Source",
                    "Base unit",
                    "quantity_used",
                    "Cost per g/ml/unit $",
                    "line_cost_usd",
                    "Status",
                ]
            ].style.format(
                {
                    "quantity_used": "{:.1f}",
                    "Cost per g/ml/unit $": "{:.5f} $",
                    "line_cost_usd": "{:.2f} $",
                }
            ),
            use_container_width=True,
        )



def get_login_credentials():
    """
    Read login credentials from Streamlit secrets.
    Do not put the real password inside app.py if the GitHub repository is public.
    """
    try:
        username = str(st.secrets["auth"]["username"])
        password = str(st.secrets["auth"]["password"])
        return username, password
    except Exception:
        return None, None


def require_login():
    username_expected, password_expected = get_login_credentials()

    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return

    if username_expected is None or password_expected is None:
        st.title("Chocoholic")
        st.error("Login is not configured. Add username and password in Streamlit Cloud Secrets.")
        st.code('[auth]\nusername = "chahra"\npassword = "YOUR_PASSWORD"', language="toml")
        st.stop()

    st.title("Chocoholic")
    st.subheader("Login")

    username_entered = st.text_input("Username")
    password_entered = st.text_input("Password", type="password")

    if st.button("Login"):
        username_ok = hmac.compare_digest(username_entered, username_expected)
        password_ok = hmac.compare_digest(password_entered, password_expected)

        if username_ok and password_ok:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect username or password.")

    st.stop()

def main():
    initialize_files()

    ingredients_input = load_csv(INGREDIENTS_FILE, INGREDIENT_INPUT_COLUMNS)
    recipe_settings = load_csv(RECIPE_SETTINGS_FILE, RECIPE_SETTINGS_COLUMNS)
    recipe_lines = load_csv(RECIPE_LINES_FILE, RECIPE_LINE_COLUMNS)

    st.set_page_config(
        page_title="Chocoholic",
        page_icon="🍫",
        layout="wide",
    )

    require_login()

    st.title("Chocoholic")
    st.caption(f"Running version: {APP_VERSION}")

    logout_col1, logout_col2 = st.columns([6, 1])
    with logout_col2:
        if st.button("Logout"):
            st.session_state["authenticated"] = False
            st.rerun()

    st.write("Ingredient prices -> recipe quantities -> cost per piece -> selling price -> profit.")

    tab_ingredients, tab_builder, tab_profit = st.tabs(
        ["1. Ingredients", "2. Recipe Builder", "3. Profit Summary"]
    )

    with tab_ingredients:
        show_ingredients_page(ingredients_input)

    with tab_builder:
        show_recipe_builder(ingredients_input, recipe_settings, recipe_lines)

    with tab_profit:
        show_profit_summary(ingredients_input, recipe_settings, recipe_lines)


if __name__ == "__main__":
    main()
