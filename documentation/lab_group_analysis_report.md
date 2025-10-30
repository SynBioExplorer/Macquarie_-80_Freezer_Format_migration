# Lab Group Value Analysis Report

**Date:** 2025-10-30
**Directory Analyzed:** `/Users/felix/Library/CloudStorage/OneDrive-SharedLibraries-MacquarieUniversity/Australian Genome Foundry - AWS cloud infrastructure/09_Freezer_Log_formatting/older freezer log format/`

## Executive Summary

- **Total Excel files processed:** 51
- **Files with "Lab group" column:** 43 files
- **Files without "Lab group" column:** 8 files
- **Unique lab group values found:** 18
- **Total occurrences:** 1,048,364

## Key Findings

### 1. Legitimate Lab Group Values (Need Standardization)

| Current Value | Frequency | Should Map To | Files |
|--------------|-----------|---------------|-------|
| **Ian Paulsen** | 1,046,056 | Paulsen | 1 |
| **Paulsen** | 1,251 | Paulsen | 15 |
| **PAULSEN** | 24 | Paulsen | 1 |
| **paulsen** | 9 | Paulsen | 2 |
| **Paulsen group** | 81 | Paulsen | 1 |
| **Cain** | 56 | Cain | 2 |
| **Jaschke lab group** | 162 | Jaschke | 2 |
| **Jaschke lab** | 81 | Jaschke | 1 |
| **Jaschke group** | 40 | Jaschke | 1 |

**Total legitimate entries:** 1,047,760

### 2. Values Requiring Manual Mapping

| Value | Frequency | Files | Notes |
|-------|-----------|-------|-------|
| **Briardo's group** | 529 | 7 | Likely "Llorente" (Briardo Llorente) |
| **Briardo's gruop** | 81 | 1 | Typo of "Briardo's group" |
| **AGF** | 67 | 1 | Australian Genome Foundry - needs mapping |
| **Packer** | 2 | 1 | Unknown lab group |

**Total entries needing manual mapping:** 679

### 3. Invalid Values (Not Lab Groups)

These values appear to be data entry errors or status notes mistakenly entered in the Lab Group column:

| Value | Frequency | File | Issue |
|-------|-----------|------|-------|
| **PCR didn't work** | 1 | rack A4 -80C freezer 3 | Status note |
| **unsure if PCR is ok** | 1 | rack A4 -80C freezer 3 | Status note |
| **PCR ok** | 1 | rack A4 -80C freezer 3 | Status note |
| **sample thrown away** | 1 | rack A4 -80C freezer 3 | Status note |
| **Waiting for sequencing** | 1 | rack B2 -80C freezer 3 | Status note |

**Total invalid entries:** 5

## Files Without "Lab Group" Column

The following 8 files do not have a "Lab group" column:

1. `BLANK TEMPLATE - PC2 Lab Sample Recording Template for -80 Freezers_May 2025.xlsx`
2. `Freezer allocations Walk-in freezer PC2 Synbio.xlsx`
3. `-80 freezer3 log/Rack C3 -80C freezer3_JordiPerez.xlsx`
4. `-80 freezer3 log/Rack 2A_-80 Freezer 3_ FP.xlsx`
5. `-80 freezer 1 log/-80 freezer 1_rack D1_EA&JP.xlsx`
6. `-80 freezer 1 log/-80C Freezer1 Rack B2 DanielPascoe.xlsx`
7. `-80 freezer 1 log/-80C Freezers_1_Rack_D3_KM.csv.xlsx`
8. `-80 freezer 1 log/-80C Freezer1 Rack C1 ArdenL&NickY.xlsx`

## Recommended Standardization Mapping

### Confirmed Mappings

```
"Ian Paulsen"       -> "Paulsen"
"Paulsen"           -> "Paulsen"
"PAULSEN"           -> "Paulsen"
"paulsen"           -> "Paulsen"
"Paulsen group"     -> "Paulsen"

"Cain"              -> "Cain"

"Jaschke lab group" -> "Jaschke"
"Jaschke lab"       -> "Jaschke"
"Jaschke group"     -> "Jaschke"
```

### Requires Confirmation

```
"Briardo's group"   -> "Llorente" (assumed - Briardo Llorente)
"Briardo's gruop"   -> "Llorente" (typo correction)
"AGF"               -> ??? (Australian Genome Foundry - which lab?)
"Packer"            -> ??? (Unknown lab)
```

### Should Be Removed/Cleaned

These are not lab groups and should be removed or moved to appropriate columns:
- "PCR didn't work"
- "unsure if PCR is ok"
- "PCR ok"
- "sample thrown away"
- "Waiting for sequencing"

## Detailed File Breakdown

### Files with Paulsen Variations

**"Ian Paulsen" (1,046,056 occurrences):**
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf B/rack B2 -80C freezer 2.xlsx`

**"Paulsen" (1,251 occurrences in 15 files):**
- `-80 freezer2 log/-80C freezer2 log_box level.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C6 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C3 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf B/rack B1 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C4 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf B/rack B5 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C5 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A1 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A3 -80C freezer 2 (SC).xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf A/rack A2 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf E/rack E1 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf E/rack E4 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf E/rack E2 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf D/rack D3 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf C/rack C2 -80C freezer 2.xlsx`

### Files with Briardo's Group (529 occurrences in 7 files)

- `-80 freezer3 log/Rack C2 -80oC freezer 3-Dayane-Costa-Williams.xlsx`
- `-80 freezer3 log/rack 2C -80C freezer 3_ArturSawiki.xlsx`
- `-80 freezer3 log/rack A4 -80C freezer 3_AfrinTalukder&VivianBonacker.xlsx`
- `-80 freezer3 log/rack 2D -80C, freezer 3, 1st column SP.xlsx`
- `-80 freezer3 log/rack B2 -80C freezer 3_VictoriaBarja.xlsx`
- `-80 freezer3 log/Rack 2E -80oC freezer 3-Josh-Carla.xlsx`
- `-80 freezer 1 log/-80C freezer 1_rack D2_VictoriaBarja.xlsx`

### Files with Jaschke Variations (283 occurrences in 4 files)

**"Jaschke lab group" (162 occurrences):**
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf D/rack D1 -80C freezer 2.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf D/2023-11-30 rack D1 -80C freezer 2 - backup.xlsx`

**"Jaschke lab" (81 occurrences):**
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf D/rack D2 -80C freezer 2.xlsx`

**"Jaschke group" (40 occurrences):**
- `-80 freezer2 log/-80C freezer2 log_box level.xlsx`

### Files with Cain (56 occurrences in 2 files)

- `-80 freezer2 log/-80C freezer2 log_box level.xlsx`
- `-80 freezer2 log/-80 freezer2 log_tube level/shelf E/rack E3 -80C freezer 2.xlsx`

## Priority Actions

1. **Confirm Briardo mapping:** Verify that "Briardo's group" should map to "Llorente"
2. **Resolve AGF:** Determine which of the four lab groups AGF samples belong to
3. **Resolve Packer:** Determine which lab group "Packer" samples belong to
4. **Clean invalid entries:** Remove or relocate the 5 status notes from the Lab Group column
5. **Add Lab Group column:** Consider adding the column to the 8 files that don't have it

## Statistics Summary

| Category | Count | Percentage |
|----------|-------|------------|
| Paulsen variations | 1,047,421 | 99.91% |
| Cain | 56 | 0.01% |
| Jaschke variations | 283 | 0.03% |
| Briardo variations | 610 | 0.06% |
| Unknown/Invalid | 74 | 0.01% |
| **TOTAL** | **1,048,444** | **100%** |

## Notes

- The extremely high count for "Ian Paulsen" (1,046,056) in a single file suggests this file contains the bulk of the freezer log data
- There is a typo in one file: "Briardo's gruop" should be "Briardo's group"
- Several PCR status notes were incorrectly entered in the Lab Group column and should be moved to a notes/status column
- No instances of "Llorente" were found in any files, but "Briardo's group" likely refers to Briardo Llorente's lab
