#!/usr/bin/env python3
"""
Extract all unique Lab Group values from Excel files in the older freezer log format directory.
"""

import pandas as pd
import os
from collections import defaultdict
from pathlib import Path

# Directory containing the Excel files
base_dir = "/Users/felix/Library/CloudStorage/OneDrive-SharedLibraries-MacquarieUniversity/Australian Genome Foundry - AWS cloud infrastructure/09_Freezer_Log_formatting/older freezer log format"

# Excel files to process
excel_files = [
    "BLANK TEMPLATE - PC2 Lab Sample Recording Template for -80 Freezers_May 2025.xlsx",
    "Freezer allocations Walk-in freezer PC2 Synbio.xlsx",
    "-80 freezer2 log/-80C freezer2 log_box level.xlsx",
    "-80 freezer3 log/Rack C2 -80oC freezer 3-Dayane-Costa-Williams.xlsx",
    "-80 freezer3 log/rack 2C -80C freezer 3_ArturSawiki.xlsx",
    "-80 freezer3 log/rack 2A -80C freezer 3.xlsx",
    "-80 freezer3 log/Rack C3 -80C freezer3_JordiPerez.xlsx",
    "-80 freezer3 log/rack A4 -80C freezer 3_AfrinTalukder&VivianBonacker.xlsx",
    "-80 freezer3 log/rack 2D -80C, freezer 3, 1st column SP.xlsx",
    "-80 freezer3 log/rack B2 -80C freezer 3_VictoriaBarja.xlsx",
    "-80 freezer3 log/Rack 2E -80oC freezer 3-Josh-Carla.xlsx",
    "-80 freezer3 log/Rack 2A_-80 Freezer 3_ FP.xlsx",
    "-80 freezer3 log/rack A1 -80C freezer 3 .xlsx",
    "-80 freezer 1 log/-80 freezer 1_rack D1_EA&JP.xlsx",
    "-80 freezer 1 log/-80C Freezer1 Rack B2 DanielPascoe.xlsx",
    "-80 freezer 1 log/-80C Freezers_1_Rack_D3_KM.csv.xlsx",
    "-80 freezer 1 log/-80C freezer1 log_box level.xlsx",
    "-80 freezer 1 log/-80C Freezer1 Rack C1 ArdenL&NickY.xlsx",
    "-80 freezer 1 log/-80C freezer 1_rack D2_VictoriaBarja.xlsx",
    "-80 freezer 1 log/-80 Freezer1 log_tube level/RackB1 -80C Freezer 1.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C6 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf D/rack D6 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf D/rack D1 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf D/2023-11-30 rack D1 -80C freezer 2 - backup.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C3 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf B/rack B1 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C1 -80C freezer 2_LV.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf D/rack D4 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf D/rack D2 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf B/rack B2 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf D/rack D5 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C4 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf B/rack B4 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf B/rack B5 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf B/rack B3 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C5 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf B/rack B6 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A5 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A1 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A4 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A3 -80C freezer 2 (SC).xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A2 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A6 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf E/rack E1 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf E/rack E5 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf E/rack E4 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf E/rack E2 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf E/rack E6 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf E/rack E3 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf D/rack D3 -80C freezer 2.xlsx",
    "-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C2 -80C freezer 2.xlsx",
]

# Dictionary to store lab group values and their occurrences
lab_group_data = defaultdict(lambda: {"count": 0, "files": []})

# Process each file
files_processed = 0
files_with_lab_group = 0
files_with_errors = []

print("=" * 80)
print("EXTRACTING LAB GROUP VALUES FROM EXCEL FILES")
print("=" * 80)
print()

for file_path in excel_files:
    full_path = os.path.join(base_dir, file_path)

    try:
        # Try reading the Excel file
        # First, try reading all sheets to find the right one
        xl_file = pd.ExcelFile(full_path)

        lab_group_found = False

        for sheet_name in xl_file.sheet_names:
            try:
                df = pd.read_excel(full_path, sheet_name=sheet_name)

                # Look for "Lab group" or "Lab Group" column (case-insensitive)
                lab_group_col = None
                for col in df.columns:
                    if isinstance(col, str) and col.lower() == "lab group":
                        lab_group_col = col
                        break

                if lab_group_col:
                    lab_group_found = True
                    files_with_lab_group += 1

                    # Extract non-empty values
                    values = df[lab_group_col].dropna()
                    # Remove empty strings and whitespace-only values
                    values = values[values.astype(str).str.strip() != ""]

                    # Count each unique value
                    for value in values:
                        value_str = str(value).strip()
                        if value_str and value_str.lower() != "nan":
                            lab_group_data[value_str]["count"] += 1
                            if file_path not in lab_group_data[value_str]["files"]:
                                lab_group_data[value_str]["files"].append(file_path)

                    break  # Found the column, no need to check other sheets

            except Exception as e:
                # Skip sheets that can't be read
                continue

        files_processed += 1

        if not lab_group_found:
            print(f"No 'Lab group' column found in: {file_path}")

    except Exception as e:
        files_with_errors.append((file_path, str(e)))
        print(f"ERROR reading {file_path}: {e}")

print()
print("=" * 80)
print("SUMMARY")
print("=" * 80)
print(f"Total files processed: {files_processed}")
print(f"Files with 'Lab group' column: {files_with_lab_group}")
print(f"Files with errors: {len(files_with_errors)}")
print()

if files_with_errors:
    print("Files with errors:")
    for file_path, error in files_with_errors:
        print(f"  - {file_path}")
        print(f"    Error: {error}")
    print()

print("=" * 80)
print("UNIQUE LAB GROUP VALUES")
print("=" * 80)
print()

if not lab_group_data:
    print("No lab group values found!")
else:
    # Sort by count (descending)
    sorted_lab_groups = sorted(lab_group_data.items(), key=lambda x: x[1]["count"], reverse=True)

    print(f"Total unique lab group values found: {len(sorted_lab_groups)}")
    print()

    for lab_group, data in sorted_lab_groups:
        print(f"Lab Group: '{lab_group}'")
        print(f"  Frequency: {data['count']}")
        print(f"  Found in {len(data['files'])} file(s):")
        for file in data['files']:
            print(f"    - {file}")
        print()

print("=" * 80)
print("STANDARDIZATION MAPPING")
print("=" * 80)
print()
print("Target standard values: Paulsen, Cain, Jaschke, Llorente")
print()

# Group similar values
if lab_group_data:
    print("Suggested mappings:")
    for lab_group, data in sorted_lab_groups:
        lower_value = lab_group.lower()

        if "paulsen" in lower_value:
            print(f"  '{lab_group}' -> 'Paulsen'")
        elif "cain" in lower_value:
            print(f"  '{lab_group}' -> 'Cain'")
        elif "jaschke" in lower_value or "jasch" in lower_value:
            print(f"  '{lab_group}' -> 'Jaschke'")
        elif "llorente" in lower_value:
            print(f"  '{lab_group}' -> 'Llorente'")
        else:
            print(f"  '{lab_group}' -> ??? (UNKNOWN - needs manual mapping)")

print()
print("=" * 80)
