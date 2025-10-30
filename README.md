# Macquarie -80°C Freezer Format Migration Tool

A Python-based tool for migrating legacy freezer log Excel files to a new standardized format for the PC2 Facility at Macquarie University.

## Overview

This project automates the migration of 51 legacy freezer log files (with 166 different column variations) into 13 standardized Excel files organized by Freezer → Shelf → Rack hierarchy.

### Key Features

- **Parallel Processing**: Utilizes multiprocessing (10 CPU cores) for fast migration
- **Data Validation**: Comprehensive validation with detailed error reporting
- **Column Mapping**: Automatically maps 166 unique old column names to 23 standardized columns
- **Lab Group Standardization**: Normalizes lab group names across files
- **Hierarchical Organization**: Organizes data by Freezer (1-3) → Shelf (A-E) → Rack
- **Progress Tracking**: Real-time progress bars with tqdm

## Migration Results

- **67,420 rows** successfully migrated
- **758 sheets** processed from 49 files
- **13 output files** created (organized by Freezer/Shelf)
- **0 errors** (all validation issues documented)
- **~20 seconds** processing time with parallel execution

## Project Structure

```
.
├── migrate_freezer_logs.py         # Main migration script
├── freezer_path_parser.py          # Path parsing utilities
├── new freezer log format/         # Template files for new format
├── analysis_scripts/               # Analysis and exploration tools
│   ├── analyze_columns.py
│   ├── extract_lab_groups.py
│   └── lab_group_mapping.csv
└── documentation/                  # Project documentation
    ├── COLUMN_ANALYSIS_SUMMARY.md
    ├── EXECUTIVE_SUMMARY.md
    └── FIXES_APPLIED.md
```

## Requirements

```bash
pip install pandas openpyxl tqdm
```

## Usage

### Basic Migration

```bash
python3 migrate_freezer_logs.py
```

The script will:
1. Process all Excel files in `older freezer log format/` directory
2. Create migrated files in `migrated_logs/` directory
3. Generate validation reports

### Output Files

**Migrated Data:**
- `6WW 326 PC2 Facility - Freezer {N} Shelf {X}.xlsx` (13 files)

**Validation Reports:**
- `validation_report.txt` - Human-readable report with file/sheet/row details
- `validation_issues.json` - Machine-readable validation data

## Column Mapping

The tool maps 166 unique old column names to 23 standardized columns:

| New Format Column | Old Format Variations |
|-------------------|----------------------|
| Sample ID | Sample ID, ID, Sample, Code |
| Sample name | Sample name, Sample Name, Sample description |
| Lab group | Lab group, Lab Group, PI |
| Sample type | Sample type, Type, Sample Type |
| ... | ... |

See [documentation/COLUMN_ANALYSIS_SUMMARY.md](documentation/COLUMN_ANALYSIS_SUMMARY.md) for complete mapping.

## Lab Group Standardization

Lab groups are automatically standardized to PI names:
- Paulsen
- Cain
- Jaschke
- Llorente
- Williams
- AGF
- Packer

## Data Validation

The migration includes comprehensive validation:

- **Invalid tube formats**: Tube positions not matching expected format (e.g., A1, B5)
- **Unknown lab groups**: Lab group values not in standardization mapping
- **Invalid date formats**: Dates that couldn't be parsed
- **Missing shelf/rack**: Files where location couldn't be determined

All validation issues are logged with:
- Source file name
- Sheet name
- Row number
- Issue description

## Architecture

### Parallel Processing

Uses Python's `multiprocessing.Pool` with 10 worker processes for optimal performance on multi-core systems.

### Hierarchical Data Structure

Data is organized as: `freezer_data[freezer_num][shelf][rack] = DataFrame`

### Validation System

Comprehensive validation logging tracks:
- File processing errors
- Column mapping issues
- Data format warnings
- Missing or invalid values

## Performance

- **Original**: Single-threaded, ~2 hours for large files
- **Optimized**: 10 parallel processes, ~20 seconds total

## Contributing

This tool was developed for the Australian Genome Foundry at Macquarie University.

## License

This project is intended for internal use at Macquarie University PC2 Facility.

## Contact

Australian Genome Foundry
Macquarie University

---

**Note**: This repository contains only the migration scripts and templates. Actual freezer log data is excluded for privacy.
