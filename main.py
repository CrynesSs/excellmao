from pathlib import Path
from copy import copy
import configparser

import pandas as pd
from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config" / "init.ini"

config = configparser.ConfigParser()
read_files = config.read(CONFIG_PATH)
if not read_files:
    raise FileNotFoundError(f"Could not read config file: {CONFIG_PATH}")


def resolve_path(raw_path: str) -> Path:
    path = Path(raw_path)
    if not path.is_absolute():
        path = BASE_DIR / path
    return path


# --- Paths ---
csv_path = resolve_path(config["PATHS"]["csv_path"])
mapping_excel_path = resolve_path(config["PATHS"]["mapping_excel_path"])
target_excel_path = resolve_path(config["PATHS"]["target_excel_path"])
output_path = resolve_path(config["PATHS"]["output_path"])

# --- Settings ---
update_existing = config.getboolean("SETTINGS", "update_existing_file")

# --- CSV config ---
csv_key_col = config["CSV"]["key_col"]
csv_value_1_col = config["CSV"]["value_1_col"]
csv_value_2_col = config["CSV"]["value_2_col"]

# --- Mapping workbook config ---
mapping_sheet_name = config["MAPPING_EXCEL"]["sheet_name"]
mapping_key_col = config["MAPPING_EXCEL"]["key_col"]
mapping_place_1_col = config["MAPPING_EXCEL"]["place_1_col"]
mapping_place_2_col = config["MAPPING_EXCEL"]["place_2_col"]

# --- Target workbook config ---
target_default_sheet_name = config["TARGET_EXCEL"]["default_sheet_name"]

# --- Style config ---
font_color = config["STYLE"].get("font_color", "FFFF0000").replace("#", "").upper()
if len(font_color) == 6:
    font_color = "FF" + font_color


if not update_existing and target_excel_path.resolve() == output_path.resolve():
    raise AssertionError(
        "Target input and output are the same, while update_existing_file=false"
    )

save_path = target_excel_path if update_existing else output_path
save_path.parent.mkdir(parents=True, exist_ok=True)


# --- Load CSV write data ---
df = pd.read_csv(csv_path, sep=",")

required_csv_cols = [csv_key_col, csv_value_1_col, csv_value_2_col]
missing_csv_cols = [col for col in required_csv_cols if col not in df.columns]
if missing_csv_cols:
    raise ValueError(f"Missing CSV columns: {missing_csv_cols}")

df[csv_key_col] = df[csv_key_col].astype(str).str.strip()

if df[csv_key_col].duplicated().any():
    duplicates = df.loc[df[csv_key_col].duplicated(), csv_key_col].tolist()
    raise ValueError(f"Duplicate IDs found in CSV: {duplicates}")

csv_lookup = df.set_index(csv_key_col)[[csv_value_1_col, csv_value_2_col]].to_dict(
    orient="index"
)


# --- Load mapping workbook, which contains ID -> target cell locations ---
mapping_wb = load_workbook(mapping_excel_path, data_only=True, read_only=True)
if mapping_sheet_name not in mapping_wb.sheetnames:
    raise ValueError(f"Mapping sheet not found: {mapping_sheet_name}")

mapping_ws = mapping_wb[mapping_sheet_name]

header_row = 1
header_values = [
    cell.value for cell in next(mapping_ws.iter_rows(min_row=header_row, max_row=header_row))
]
headers = {
    str(value).strip(): idx
    for idx, value in enumerate(header_values)
    if value is not None
}

required_mapping_cols = [mapping_key_col, mapping_place_1_col, mapping_place_2_col]
missing_mapping_cols = [col for col in required_mapping_cols if col not in headers]
if missing_mapping_cols:
    raise ValueError(f"Missing mapping workbook columns: {missing_mapping_cols}")

mapping_key_idx = headers[mapping_key_col]
place_1_idx = headers[mapping_place_1_col]
place_2_idx = headers[mapping_place_2_col]

mapping_lookup = {}
for values in mapping_ws.iter_rows(min_row=header_row + 1, values_only=True):
    if not values or values[mapping_key_idx] is None:
        continue

    mapping_id = str(values[mapping_key_idx]).strip()

    if mapping_id in mapping_lookup:
        raise ValueError(f"Duplicate ID found in mapping workbook: {mapping_id}")

    mapping_lookup[mapping_id] = {
        mapping_place_1_col: values[place_1_idx],
        mapping_place_2_col: values[place_2_idx],
    }

mapping_wb.close()


# --- Load target workbook, which receives the values ---
keep_vba = target_excel_path.suffix.lower() == ".xlsm"
target_wb = load_workbook(target_excel_path, keep_vba=keep_vba, rich_text=True)

if target_default_sheet_name not in target_wb.sheetnames:
    raise ValueError(f"Default sheet not found: {target_default_sheet_name}")


def clean_excel_value(value):
    if pd.isna(value):
        return None
    if hasattr(value, "item"):
        return value.item()
    return value


def write_red_value(workbook, default_sheet_name, cell_reference, value, color):
    """
    Write value into the target workbook using either:
      - plain cell address: E7
      - sheet-qualified address: Output!E7
    """
    if cell_reference is None:
        return None

    raw_reference = str(cell_reference).strip()
    if not raw_reference:
        return None

    if "!" in raw_reference:
        sheet_name, address = raw_reference.split("!", 1)
        sheet_name = sheet_name.strip().strip("'")
        address = address.strip()
    else:
        sheet_name = default_sheet_name
        address = raw_reference

    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found for target reference '{raw_reference}'")

    ws = workbook[sheet_name]
    cell = ws[address]
    cell.value = clean_excel_value(value)

    new_font = copy(cell.font)
    new_font.color = color
    cell.font = new_font

    return f"{sheet_name}!{address}"


matched = 0
csv_without_mapping = []
written_values = 0
written_cells = []

for csv_id, csv_values in csv_lookup.items():
    if csv_id not in mapping_lookup:
        csv_without_mapping.append(csv_id)
        continue

    place_1 = mapping_lookup[csv_id][mapping_place_1_col]
    place_2 = mapping_lookup[csv_id][mapping_place_2_col]

    value_1 = csv_values[csv_value_1_col]
    value_2 = csv_values[csv_value_2_col]

    written_1 = write_red_value(
        target_wb,
        target_default_sheet_name,
        place_1,
        value_1,
        font_color,
    )
    if written_1:
        written_values += 1
        written_cells.append((csv_id, csv_value_1_col, written_1))

    written_2 = write_red_value(
        target_wb,
        target_default_sheet_name,
        place_2,
        value_2,
        font_color,
    )
    if written_2:
        written_values += 1
        written_cells.append((csv_id, csv_value_2_col, written_2))

    matched += 1

mapping_without_csv = sorted(set(mapping_lookup) - set(csv_lookup))

target_wb.save(save_path)

print("Done.")
print(f"Matched IDs: {matched}")
print(f"CSV IDs without mapping: {len(csv_without_mapping)}")
if csv_without_mapping:
    print(f"  {csv_without_mapping}")
print(f"Mapping IDs without CSV value: {len(mapping_without_csv)}")
if mapping_without_csv:
    print(f"  {mapping_without_csv}")
print(f"Written values: {written_values}")
for csv_id, value_col, address in written_cells:
    print(f"  {csv_id} / {value_col} -> {address}")
print(f"Saved to: {save_path}")
