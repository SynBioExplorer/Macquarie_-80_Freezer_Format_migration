#!/usr/bin/env python3
"""
Freezer Path Parser
===================
Helper functions to parse freezer, shelf, and rack information from file paths.

Author: Claude Code
Date: 2025-10-30
"""

import re
from pathlib import Path
from typing import Optional, Tuple


def parse_freezer_info_from_path(file_path: Path) -> Tuple[Optional[int], Optional[str], Optional[str]]:
    """
    Parse freezer number, shelf, and rack from file path.

    Args:
        file_path: Path to Excel file

    Returns:
        Tuple of (freezer_num, shelf, rack) or (None, None, None) if can't parse

    Examples:
        "older freezer log format/-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A1 -80C freezer 2.xlsx"
        -> (2, 'A', 'A1')

        "older freezer log format/-80 freezer3 log/rack B2 -80C freezer 3_VictoriaBarja.xlsx"
        -> (3, None, 'B2')
    """
    path_str = str(file_path)
    filename = file_path.stem  # filename without extension

    # Extract freezer number
    freezer_num = None

    # Try from path
    freezer_match = re.search(r'freezer\s*([123])', path_str, re.IGNORECASE)
    if freezer_match:
        freezer_num = int(freezer_match.group(1))

    # Extract shelf
    shelf = None

    # Try from path: "shelf A", "shelf B", etc.
    shelf_match = re.search(r'shelf\s+([A-E])', path_str, re.IGNORECASE)
    if shelf_match:
        shelf = shelf_match.group(1).upper()

    # Extract rack
    rack = None

    # Try various rack patterns
    rack_patterns = [
        r'[Rr]ack\s+(\d+[A-E])',    # "rack 2A", "rack 3B" (number+letter)
        r'[Rr]ack\s+([A-E]\d+)',    # "rack A1", "rack B2" (letter+number)
        r'_([A-E]\d+)[\s_-]',       # "_A1 ", "_B2_"
        r'[Rr]ack\s+([A-E][A-E]\d+)', # "rack AA1"
    ]

    for pattern in rack_patterns:
        rack_match = re.search(pattern, filename)
        if rack_match:
            rack = rack_match.group(1).upper()
            break

    # If rack not found from filename, try from parent directory
    if not rack:
        parent_name = file_path.parent.name
        for pattern in rack_patterns:
            rack_match = re.search(pattern, parent_name)
            if rack_match:
                rack = rack_match.group(1).upper()
                break

    return freezer_num, shelf, rack


def normalize_rack_name(rack: str) -> str:
    """
    Normalize rack name to consistent format.

    Examples:
        "A1" -> "A1"
        "2A" -> "A2"  (flip order)
        "B2" -> "B2"
        "1" -> "1"
    """
    if not rack:
        return rack

    rack = rack.strip().upper()

    # If format is "2A" (number+letter), flip to "A2" (letter+number)
    match = re.match(r'^(\d+)([A-E])$', rack)
    if match:
        return f"{match.group(2)}{match.group(1)}"

    return rack


def extract_shelf_from_rack(rack: str) -> Optional[str]:
    """
    Extract shelf letter from rack name.

    Examples:
        "A1" -> "A"
        "B2" -> "B"
        "C3" -> "C"
    """
    if not rack:
        return None

    match = re.match(r'^([A-E])', rack.upper())
    if match:
        return match.group(1)

    return None


def map_old_sheet_to_new_structure(file_path: Path, sheet_name: str) -> Tuple[Optional[int], Optional[str], Optional[str]]:
    """
    Map old format file/sheet to new format structure (freezer, shelf, rack).

    Args:
        file_path: Path to old format Excel file
        sheet_name: Sheet name within the file

    Returns:
        Tuple of (freezer_num, shelf, rack)
    """
    # Parse from file path
    freezer_num, shelf, rack = parse_freezer_info_from_path(file_path)

    # Try to extract additional info from sheet name
    if not freezer_num:
        sheet_freezer_match = re.search(r'freezer\s*([123])', sheet_name, re.IGNORECASE)
        if sheet_freezer_match:
            freezer_num = int(sheet_freezer_match.group(1))

    if not shelf:
        sheet_shelf_match = re.search(r'shelf\s+([A-E])', sheet_name, re.IGNORECASE)
        if sheet_shelf_match:
            shelf = sheet_shelf_match.group(1).upper()

    if not rack:
        # Try to extract rack from sheet name
        rack_patterns = [
            r'[Rr]ack\s+([A-E]?\d+)',
            r'Box\s+(\d+)',
        ]
        for pattern in rack_patterns:
            rack_match = re.search(pattern, sheet_name)
            if rack_match:
                rack = rack_match.group(1).upper()
                break

    # Normalize rack name
    if rack:
        rack = normalize_rack_name(rack)

    # If we have a rack but no shelf, try to extract shelf from rack
    if rack and not shelf:
        shelf = extract_shelf_from_rack(rack)

    return freezer_num, shelf, rack


if __name__ == '__main__':
    # Test cases
    test_paths = [
        Path("older freezer log format/-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A1 -80C freezer 2.xlsx"),
        Path("older freezer log format/-80 freezer3 log/rack B2 -80C freezer 3_VictoriaBarja.xlsx"),
        Path("older freezer log format/-80 freezer 1 log/-80C Freezer1 Rack C1 ArdenL&NickY.xlsx"),
        Path("older freezer log format/-80 freezer3 log/Rack 2A_-80 Freezer 3_ FP.xlsx"),
    ]

    print("Testing path parser:")
    print("="*80)
    for path in test_paths:
        freezer, shelf, rack = map_old_sheet_to_new_structure(path, "")
        print(f"File: {path.name}")
        print(f"  -> Freezer: {freezer}, Shelf: {shelf}, Rack: {rack}")
        print()
