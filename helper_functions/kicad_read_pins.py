from kiutils.board import Board
from kiutils.libraries import LibTable
from kiutils.schematic import Schematic
from kiutils.footprint import Footprint
from kiutils.symbol import SymbolLib
from kiutils.wks import WorkSheet
from kiutils.dru import DesignRules
from collections import defaultdict, deque
import pandas as pd


def get_symbol_by_entry_name(symbols, entry_name):
    for sym in symbols:
        if sym.entryName == entry_name:
            return sym
    return None


def get_pin_list_from_symbol(symbol):
    # Convert keys to int and find the maximum pin number
    try:
        pins_int = {int(number): uuid for number, uuid in symbol.pins.items()}
    except ValueError:
        raise ValueError("Pin numbers must be integers.")
    max_pin = max(pins_int.keys())

    # Create a list with None for missing pins
    pin_list = [None] * (max_pin + 1)

    # Fill the list so that index == pin number
    for number, uuid in pins_int.items():
        pin_list[number] = {"number": number, "uuid": uuid}

    # # Example: print the pins
    # for idx, pin in enumerate(pin_list):
    #     if idx == 18:
    #         print(f"Pin number: {pin['number']}, UUID: {pin['uuid']}")
    return pin_list


def get_pins_from_library(lib_file, _unit):
    if not _unit:
        raise ValueError("Unit must be specified.")
    # Load the library
    library = SymbolLib().from_file(lib_file)

    if len(library.symbols) != 1:
        raise ValueError("Library should contain exactly one symbol.")

    pin_list = []

    for sym in library.symbols:
        # Iterate over all units
        for unit in sym.units:
            if unit.unitId != _unit:
                continue
            for pin in unit.pins:  # pins are SymbolPin objects
                pin_list.append(pin)
                # print(f"    Pin number: {pin.number}, name: {pin.name}, position: ({pin.position})")
    return pin_list


def is_close(p1, p2, tol=0.01):
    return abs(p1[0] - p2[0]) < tol and abs(p1[1] - p2[1]) < tol


def build_wire_graph(wires):
    """
    Build adjacency graph of wires. Wires are connected if they share a coordinate.
    Returns:
        wire_graph: dict wire_uuid -> set of connected wire_uuids
        coord_to_wires: dict coord -> list of wire_uuids touching it
    """
    coord_to_wires = defaultdict(list)
    for wire in wires:
        coord_to_wires[wire["start"]].append(wire["uuid"])
        coord_to_wires[wire["end"]].append(wire["uuid"])

    wire_graph = defaultdict(set)
    for wire in wires:
        wire_id = wire["uuid"]
        for neighbor in coord_to_wires[wire["start"]] + coord_to_wires[wire["end"]]:
            if neighbor != wire_id:
                wire_graph[wire_id].add(neighbor)
    return wire_graph, coord_to_wires


def _kicad_label_to_pin_map(pins, labels, wires):
    """
    Map labels to pins through wire connectivity (multi-segment nets and junctions supported)
    Returns: dict label_text -> list of SymbolPin objects
    """
    wire_graph, coord_to_wires = build_wire_graph(wires)

    # Map pins for easy lookup
    pin_id_map = {id(pin): pin for pin in pins}

    # Map pins -> wires they touch
    pin_to_wires = defaultdict(set)
    for pin in pins:
        for coord, wire_uuids in coord_to_wires.items():
            if is_close(pin.position, coord):
                pin_to_wires[id(pin)].update(wire_uuids)
    # Map labels -> wires they touch
    label_to_wires = defaultdict(set)
    for label in labels:
        label_pos = (label.position.X, label.position.Y)
        for coord, wire_uuids in coord_to_wires.items():
            if is_close(label_pos, coord):
                label_to_wires[label.text].update(wire_uuids)

    # BFS function to find all reachable wires from a set of starting wires
    def reachable_wires(start_wires):
        visited = set()
        queue = deque(start_wires)
        while queue:
            w = queue.popleft()
            if w in visited:
                continue
            visited.add(w)
            for neighbor in wire_graph[w]:
                if neighbor not in visited:
                    queue.append(neighbor)
        return visited

    # Build label -> pins mapping
    label_to_pins = defaultdict(list)
    for label_text, start_wires in label_to_wires.items():
        all_reachable = reachable_wires(start_wires)
        for pin_id, pin_wires in pin_to_wires.items():
            if all_reachable & pin_wires:  # if any wire overlaps
                label_to_pins[label_text].append(pin_id_map[pin_id])
    return label_to_pins


def _load_and_parse_schematic(lib_file: str, schematic_file: str, unit: int):

    # load schematic from file
    schematic = Schematic().from_file(schematic_file)

    # extract all symbols from schematic
    symbols = schematic.schematicSymbols

    # find the symbol based on its entryName
    symbol = get_symbol_by_entry_name(symbols, "STM32H562VGT6")
    if symbol.position.angle != 0:
        raise ValueError("Rotation not implemented yet")

    pin_list = get_pin_list_from_symbol(symbol)

    # Load the symbol library to the the pins location
    pin_list = get_pins_from_library(lib_file, unit)

    # Adjust pin positions based on symbol position
    for pin in pin_list:
        pin.position = (
            symbol.position.X + pin.position.X,
            symbol.position.Y - pin.position.Y,
        )

    # find all labels
    labels = schematic.labels

    # for label in labels:
    #     print(label.uuid, label.text, label.position)

    wires_raw = [item for item in schematic.graphicalItems if item.type == "wire"]
    wires = []
    for wire_raw in wires_raw:
        # unpack start and end points
        points = wire_raw.points
        # Extract start coordinates (first point)
        start_x = points[0].X
        start_y = points[0].Y

        # Extract end coordinates (second point)
        end_x = points[1].X
        end_y = points[1].Y
        wires.append(
            {"uuid": wire_raw.uuid, "start": (start_x, start_y), "end": (end_x, end_y)}
        )
    # delete wires_raw  # free memory
    del wires_raw, wire_raw, points, start_x, start_y, end_x, end_y
    # for wire in wires:
    #     print(f"Wire UUID: {wire['uuid']}, Start: {wire['start']}, End: {wire['end']}")
    return pin_list, labels, wires


def kicad_pins_to_labels_map(lib_file: str, schematic_file: str, unit: int):
    """
    Wrapper for kicad_pins_to_labels_map to handle SymbolPin objects
    """
    # lib_file = "STM32H562VGT6.kicad_sym"
    # schematic_file = "SDIO_Multiplexer.kicad_sch"
    if not lib_file or not schematic_file:
        raise ValueError("Both lib_file and schematic_file must be provided.")
    pins, labels, wires = _load_and_parse_schematic(lib_file, schematic_file, unit)
    pins_to_labels = _kicad_label_to_pin_map(pins, labels, wires)

    # Build list of dictionaries
    data = []
    for label, connected_pins in pins_to_labels.items():
        for pin in connected_pins:
            data.append({"Label": label, "PinNumber": pin.number, "PinName": pin.name})

    # Convert to DataFrame
    df = pd.DataFrame(data)

    # Optional: sort by Label or PinNumber
    pins_to_labels = df.sort_values(by=["PinNumber"]).reset_index(drop=True)
    return pins_to_labels
