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


def find_symbol_instance(sch_text, lib_id):
    """
    Find the first instance of a symbol with the given lib_id in the schematic.
    Returns its absolute position and rotation: (x, y, rotation_deg)
    """
    # Match a symbol block containing the lib_id
    m = re.search(
        r'\(symbol\s*\(\s*lib_id\s+"{}"\)\s*\(at\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)?\)'.format(
            re.escape(lib_id)
        ),
        sch_text,
        flags=re.DOTALL,
    )
    if not m:
        print(f"Warning: Symbol with lib_id '{lib_id}' not found in schematic.")
        return None, None, None

    x = float(m.group(1))
    y = float(m.group(2))
    rotation = float(m.group(3)) if m.group(3) else 0.0
    return x, y, rotation


import math


def compute_pin_points(symbol_points, pin_rel_points, pin_length, pin_rot):
    """
    Given symbol absolute position, pin relative position, pin length and pin rotation,
    compute absolute anchor and endpoint of the pin.
    """
    symbol_x, symbol_y = symbol_points
    pin_rel_x, pin_rel_y = pin_rel_points

    # Compute absolute anchor point
    abs_anchor_x = symbol_x + pin_rel_x
    abs_anchor_y = symbol_y + pin_rel_y
    abs_anchor = (abs_anchor_x, abs_anchor_y)

    # Compute endpoint based on rotation and length
    theta = math.radians(180 - pin_rot)  # Adjust for KiCad's coordinate system
    dx = pin_length * math.cos(theta)
    dy = pin_length * math.sin(theta)
    abs_endpoint = (abs_anchor_x + dx, abs_anchor_y + dy)

    return abs_anchor, abs_endpoint


def parse_ic_pins(symbol_x, symbol_y, symbol_rotation, sch_text, pin_names=None):
    """
    Parse the schematic for pins by searching for pin names directly.
    Returns a list of dicts: { 'pin_name', 'abs_anchor', 'length', 'pin_type' }
    If pin_names is provided, only looks for those pins; otherwise, finds all pins.
    Warnings are printed if a pin is missing or duplicated.
    """
    pins = []
    if pin_names is None:
        pin_names = set(re.findall(r'\(name\s+"([^"]+)"', sch_text))

    for pin_name in pin_names:
        matches = [
            m.start()
            for m in re.finditer(r'\(name\s+"{}"'.format(re.escape(pin_name)), sch_text)
        ]
        if len(matches) == 0:
            print(f"Warning: Pin '{pin_name}' not found in schematic!")
            continue
        if len(matches) > 1:
            print(f"Warning: Pin '{pin_name}' found more than once in schematic!")
            continue

        start_idx = sch_text.rfind("(pin", 0, matches[0])
        if start_idx == -1:
            print(f"Warning: Could not find '(pin' start for pin '{pin_name}'")
            continue

        try:
            pin_block, _ = extract_parentheses_block(sch_text, start_idx)
        except ValueError:
            print(f"Warning: Could not extract pin block for '{pin_name}'")
            continue

        at_m = re.search(
            r"\(at\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)\)",
            pin_block,
        )
        if not at_m:
            print(f"Warning: Could not find (at ...) for pin '{pin_name}'")
            continue
        x, y, rot = map(float, at_m.groups())

        length_m = re.search(r"\(length\s+([-+]?\d*\.?\d+)\)", pin_block)
        if not length_m:
            print(f"Warning: Could not find (length ...) for pin '{pin_name}'")
            continue
        length = float(length_m.group(1))

        type_m = re.search(r"\(pin\s+(\S+)\s+line", pin_block)
        if not type_m:
            print(f"Warning: Could not find pin type for pin '{pin_name}'")
            continue
        pin_type = type_m.group(1)

        abs_anchor, abs_endpoint = compute_pin_points(
            (symbol_x, symbol_y), (x, y), length, rot + symbol_rotation
        )
        pins.append(
            {
                "pin_name": pin_name,
                "rel_anchor": (x, y),
                "rot": rot,
                "length": length,
                "pin_type": pin_type,
                "abs_anchor": abs_anchor,
                "abs_endpoint": abs_endpoint,
            }
        )
        print(
            f"Pin parsed: {pin_name}, anchor: {abs_anchor}, endpoint: {abs_endpoint}, length: {length}, type: {pin_type}"
        )

    print(f"Total pins parsed: {len(pins)}")
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

    symbol_x, symbol_y, symbol_rotation = find_symbol_instance(
        sch_text, lib_id="MCU_ST_STM32H5:STM32H503CBTx"
    )
    print(f"MCU symbol position and rotation: {(symbol_x, symbol_y, symbol_rotation)}")
    if symbol_rotation != 0.0:
        print(symbol_rotation)
        raise NotImplementedError("Rotation handling not implemented yet.")

    all_pins = parse_ic_pins(
        symbol_x, symbol_y, symbol_rotation, sch_text, pin_names=pin_map.keys()
    )
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
        all_pins = parse_ic_pins(
            symbol_x, symbol_y, symbol_rotation, sch_text, pin_names=pin_map.keys()
        )
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
