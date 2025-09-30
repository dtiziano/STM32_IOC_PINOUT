#!/usr/bin/env python3
"""
kicad_label_replace.py

Usage:
    python kicad_label_replace.py <schematic.kicad_sch> <pins.xlsx>

Reads the Excel file (expects a sheet with columns 'PinNumber' and 'Signal'),
parses the KiCad schematic (.kicad_sch), finds labels containing the placeholder
'XXXXXXXXXXXXXXXXX', locates which wire each label is on, follows that wire to
the connected MCU pin, looks up the PinNumber in the Excel and replaces the
label text with the Signal from Excel.

Added debug prints for pin, wire, and label locations.
"""

import sys
import re
import math
import shutil
import pandas as pd
from pathlib import Path
import math

COORD_TOLERANCE = 1e-3  # tolerance for coordinate matches (adjust if needed)


def read_pin_map_from_excel(excel_path):
    """Reads Excel and returns PinNumber -> Signal mapping."""
    df_raw = pd.read_excel(excel_path, sheet_name=0, header=None, dtype=str)
    header_row_idx = None

    # Find header row by scanning for 'PinNumber' and 'Signal'
    for i, row in df_raw.iterrows():
        row_lower = [str(cell).strip().lower() for cell in row]
        if "pinnumber" in row_lower and "signal" in row_lower:
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError(
            "Could not find header row containing 'PinNumber' and 'Signal' in Excel."
        )

    df = pd.read_excel(excel_path, sheet_name=0, header=header_row_idx, dtype=str)
    col_map = {c.lower(): c for c in df.columns}
    pin_col = col_map["pinname"]
    sig_col = col_map["signal"]

    pin_map = {}
    for _, row in df.iterrows():
        pin = str(row[pin_col]).strip()
        signal = str(row[sig_col]).strip()
        if pin and pin.lower() not in ("nan", "none"):
            pin_map[pin] = signal
    return pin_map


def read_pin_map_from_excel_by_name(excel_path):
    """
    Reads the Excel file and returns a dictionary mapping PinName -> Signal.
    Automatically detects the header row by scanning for the required columns.
    """
    df_raw = pd.read_excel(excel_path, sheet_name=0, header=None, dtype=str)
    header_row_idx = None

    # Look for the row that contains both 'PinName' and 'Signal'
    for i, row in df_raw.iterrows():
        row_lower = [str(cell).strip().lower() for cell in row]
        if "pinname" in row_lower and "signal" in row_lower:
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError(
            "Could not find header row containing 'PinName' and 'Signal' in Excel."
        )

    # Read Excel again using detected header row
    df = pd.read_excel(excel_path, sheet_name=0, header=header_row_idx, dtype=str)

    # Normalize column names
    col_map = {c.lower(): c for c in df.columns}
    pin_col = col_map["pinname"]
    sig_col = col_map["signal"]

    # Build mapping dictionary
    pin_map = {}
    for _, row in df.iterrows():
        pin = str(row[pin_col]).strip()
        signal = str(row[sig_col]).strip()
        if pin and pin.lower() not in ("nan", "none"):
            pin_map[pin] = signal

    return pin_map


def float_equal(a, b, tol=COORD_TOLERANCE):
    return abs(float(a) - float(b)) <= tol


def coord_equal(a, b, tol=COORD_TOLERANCE):
    return float_equal(a[0], b[0], tol) and float_equal(a[1], b[1], tol)


def parse_wire_blocks(sch_text):
    """
    Return list of wires, each wire is dict with list of pts (x,y).
    """
    wires = []
    pos = 0
    while True:
        m = re.search(r"\(wire\b", sch_text[pos:])
        if not m:
            break
        start = pos + m.start()
        block, end_pos = extract_parentheses_block(sch_text, start)
        # find the pts block inside
        pts_m = re.search(r"\(pts\b", block)
        pts_f = []
        if pts_m:
            pts_start = pts_m.start()
            pts_block, _ = extract_parentheses_block(block, pts_start)
            # extract all (xy x y)
            pts = re.findall(r"\(xy\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)\)", pts_block)
            pts_f = [(float(x), float(y)) for x, y in pts]
        wires.append({"pts": pts_f, "wire_start": start})
        pos = end_pos
    return wires


def parse_label_blocks(sch_text, placeholder="XXXXXXXXXXXXXXXXX"):
    """Return list of labels with xy and full span."""
    labels = []
    for m in re.finditer(r'\(label\s+"([^"]+)"(.*?)\)', sch_text, flags=re.DOTALL):
        text = m.group(1)
        body = m.group(2)
        if text != placeholder:
            continue
        at_m = re.search(r"\(at\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)", body)
        if not at_m:
            full = sch_text[m.start() : m.end()]
            at_m2 = re.search(r"\(at\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)", full)
            if not at_m2:
                print(
                    "Warning: label without (at ...) coords, skipping", m.group(0)[:80]
                )
                continue
            x, y = float(at_m2.group(1)), float(at_m2.group(2))
        else:
            x, y = float(at_m.group(1)), float(at_m.group(2))
        labels.append(
            {
                "text": text,
                "xy": (x, y),
                "start": m.start(1),
                "end": m.end(1),
                "full_span": (m.start(), m.end()),
                "full_text": m.group(0),
            }
        )
    return labels


import math


def compute_pin_endpoint(symbol_origin, pin_at, length, rotation_deg):
    """
    Compute absolute anchor and endpoint of a pin.
    """
    x0, y0 = symbol_origin
    x_rel, y_rel = pin_at
    abs_anchor = (x0 + x_rel, y0 + y_rel)

    theta = math.radians(180 - rotation_deg)
    dx = length * math.cos(theta)
    dy = length * math.sin(theta)
    abs_endpoint = (abs_anchor[0] + dx, abs_anchor[1] + dy)
    print(
        "origin:",
        symbol_origin,
        "pin_at:",
        pin_at,
        "length:",
        length,
        "rot:",
        rotation_deg,
    )
    return abs_anchor, abs_endpoint


import re
import math


def extract_parentheses_block(text, start_idx):
    if text[start_idx] != "(":
        raise ValueError("Expected '(' at start_idx")

    depth = 1
    idx = start_idx + 1
    in_string = False

    while idx < len(text):
        c = text[idx]

        if c == '"' and text[idx - 1] != "\\":  # toggle string state
            in_string = not in_string

        elif not in_string:
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
                if depth == 0:
                    return text[start_idx : idx + 1], idx + 1

        idx += 1

    raise ValueError("No matching closing parenthesis found")


def parse_ic_pins(sch_text, symbol_name="MCU_ST_STM32H5:STM32H503CBTx"):
    """
    Parse the MCU symbol inside lib_symbols and return pins with absolute positions.
    Returns list of dicts: { 'pin_name', 'abs_anchor', 'abs_endpoint' }
    """
    pins = []

    # 1. Extract lib_symbols block
    lib_m = re.search(r"\(lib_symbols", sch_text)
    if not lib_m:
        print("ERROR: Could not find (lib_symbols block")
        return []
    lib_block, _ = extract_parentheses_block(sch_text, lib_m.start())
    if lib_block is None:
        print("ERROR: Could not extract lib_symbols block")
        return []

    # 2. Find top-level MCU symbol inside lib_symbols
    top_match = re.search(r'\(symbol\s+"{}"'.format(re.escape(symbol_name)), lib_block)
    if not top_match:
        print(f"ERROR: MCU symbol '{symbol_name}' not found inside lib_symbols")
        return []

    top_block, _ = extract_parentheses_block(lib_block, top_match.start())
    if top_block is None:
        print("ERROR: Could not extract top-level MCU symbol block")
        return []

    # 3. Get top-level origin
    at_m = re.search(r"\(at\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)", top_block)
    top_origin = (float(at_m.group(1)), float(at_m.group(2))) if at_m else (0.0, 0.0)
    print("Top-level MCU origin:", top_origin)

    # 4. Find all nested symbols
    nested_matches = list(re.finditer(r'\(symbol\s+"[^"]+"', top_block))
    for nm in nested_matches:
        nested_start = nm.start()
        nested_block, _ = extract_parentheses_block(top_block, nested_start)
        if nested_block is None:
            continue

        # nested origin
        at_m = re.search(
            r"\(at\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s*(-?\d+(?:\.\d+)?)?\)",
            nested_block,
        )
        nested_origin = (
            (float(at_m.group(1)), float(at_m.group(2))) if at_m else (0.0, 0.0)
        )
        abs_nested_origin = (
            top_origin[0] + nested_origin[0],
            top_origin[1] + nested_origin[1],
        )
        print("Nested symbol origin:", abs_nested_origin)

        # parse pins
        pin_matches = re.finditer(r"\(pin\b(.*?)\)\s*\)", nested_block, flags=re.DOTALL)
        for pm in pin_matches:
            pin_text = pm.group(1)
            name_m = re.search(r'\(name\s+"([^"]+)"', pin_text)
            at_m = re.search(
                r"\(at\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)\)",
                pin_text,
            )
            length_m = re.search(r"\(length\s+([-\d\.]+)\)", pin_text)
            if not name_m or not at_m or not length_m:
                continue
            pin_name = name_m.group(1)
            rel_x, rel_y = float(at_m.group(1)), float(at_m.group(2))
            rot_deg = float(at_m.group(3)) if at_m.group(3) else 0.0
            length = float(length_m.group(1))
            abs_anchor = (abs_nested_origin[0] + rel_x, abs_nested_origin[1] + rel_y)
            rot_rad = math.radians(rot_deg)
            dx = length * math.cos(rot_rad)
            dy = length * math.sin(rot_rad)
            abs_endpoint = (abs_anchor[0] + dx, abs_anchor[1] + dy)
            pins.append(
                {
                    "pin_name": pin_name,
                    "abs_anchor": abs_anchor,
                    "abs_endpoint": abs_endpoint,
                }
            )
            print(f"Pin: {pin_name}, anchor: {abs_anchor}, endpoint: {abs_endpoint}")

    print(f"Total MCU pins parsed: {len(pins)}")
    return pins


def find_wire_for_point(wires, point):
    for wi, w in enumerate(wires):
        for pi, pt in enumerate(w["pts"]):
            if coord_equal(pt, point):
                return wi, pi
    return None, None


def find_pin_on_wire(wires, all_pins, wire_idx):
    w = wires[wire_idx]
    for pt in w["pts"]:
        for p in all_pins:
            if coord_equal(p["abs_endpoint"], pt) or coord_equal(p["abs_anchor"], pt):
                return p
    return None


def replace_label_in_text(sch_text, label_full_span, new_text):
    block = sch_text[label_full_span[0] : label_full_span[1]]
    new_block = block.replace('"XXXXXXXXXXXXXXXXX"', f'"{new_text}"', 1)
    return sch_text[: label_full_span[0]] + new_block + sch_text[label_full_span[1] :]


def main(sch_path, excel_path, placeholder="XXXXXXXXXXXXXXXXX", backup=True):
    sch_path = Path(sch_path)
    excel_path = Path(excel_path)
    if not sch_path.exists():
        raise FileNotFoundError(f"Schematic file not found: {sch_path}")
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    pin_map = read_pin_map_from_excel(excel_path)
    print(f"Loaded {len(pin_map)} pin mappings from Excel.")

    with open(sch_path, "r", encoding="utf-8") as f:
        sch_text = f.read()

    wires = parse_wire_blocks(sch_text)
    print(f"Found {len(wires)} wires (with pts).")

    labels = parse_label_blocks(sch_text, placeholder=placeholder)
    print(f"Found {len(labels)} labels with placeholder '{placeholder}'.")

    all_pins = parse_ic_pins(sch_text)
    print(f"Found {len(all_pins)} parsed pins (with computed endpoints).")

    # Debug prints
    print("\n=== MCU Pins ===")
    for p in all_pins:
        print(
            f"Pin ({p['pin_name']}): anchor={p['abs_anchor']}, endpoint={p['abs_endpoint']}"
        )

    print("\n=== Wires ===")
    for wi, w in enumerate(wires):
        print(f"Wire {wi}: pts={w['pts']}")

    print("\n=== Labels ===")
    for lab in labels:
        print(f"Label at {lab['xy']} text={lab['text']}")

    errors = []
    replacements = []

    for label in labels:
        lab_xy = label["xy"]
        print(f"Processing label at {lab_xy}")

        # find wire that contains this label
        wire_idx, pt_idx = find_wire_for_point(wires, lab_xy)
        if wire_idx is None:
            errors.append(f"Label at {lab_xy} not on any wire point.")
            continue

        # find pin connected to that wire
        pin = find_pin_on_wire(wires, all_pins, wire_idx)
        if pin is None:
            errors.append(f"No pin found connected to wire for label at {lab_xy}.")
            continue

        pin_name = pin["pin_name"]
        print(
            f"  Label connects to pin_name: {pin_name} at endpoint {pin['abs_endpoint']}"
        )

        if pin_name not in pin_map:
            errors.append(
                f"Pin name {pin_name} (found for label at {lab_xy}) not in Excel map."
            )
            continue

        signal = pin_map[pin_name]

        # replace label text in schematic
        try:
            sch_text = replace_label_in_text(sch_text, label["full_span"], signal)
            replacements.append(
                {"label_xy": lab_xy, "pin_name": pin_name, "signal": signal}
            )
            print(f"  Replaced label at {lab_xy} with '{signal}'")
        except Exception as e:
            errors.append(f"Failed to replace label at {lab_xy}: {e}")
            continue

        wires = parse_wire_blocks(sch_text)
        labels = parse_label_blocks(sch_text, placeholder=placeholder)
        all_pins = parse_ic_pins(sch_text)
        replacements.append({"label_xy": lab_xy, "pin": pin_num, "signal": signal})

    if errors:
        if backup:
            bak = sch_path.with_suffix(sch_path.suffix + ".bak")
            shutil.copy2(sch_path, bak)
            print(f"Backup written to {bak}")
        with open(sch_path, "w", encoding="utf-8") as f:
            f.write(sch_text)
        err_text = "\n".join(errors)
        raise RuntimeError(
            f"Completed with errors:\n{err_text}\nReplaced labels: {replacements}"
        )
    else:
        if backup:
            bak = sch_path.with_suffix(sch_path.suffix + ".bak")
            shutil.copy2(sch_path, bak)
            print(f"Backup written to {bak}")
        with open(sch_path, "w", encoding="utf-8") as f:
            f.write(sch_text)
        print(f"All labels updated successfully. Replacements: {replacements}")


if __name__ == "__main__":
    # if len(sys.argv) != 3:
    #     print("Usage: python kicad_label_replace.py <schematic.kicad_sch> <pins.xlsx>")
    #     sys.exit(1)
    # sch_file = sys.argv[1]
    # excel_file = sys.argv[2]
    sch_file = (
        "/Users/tiziano/Documents/GitHub/STM32_IOC_PINOUT/SDIO_Multiplexer.kicad_sch"
    )
    excel_file = (
        "/Users/tiziano/Documents/GitHub/STM32_IOC_PINOUT/STM32H562VGT6_config.xlsx"
    )

    main(sch_file, excel_file)
