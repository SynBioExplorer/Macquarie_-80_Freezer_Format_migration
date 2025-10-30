# Comprehensive Column Analysis: Older Freezer Log Format Files

## Executive Summary

**Analysis Date:** 2025-10-30
**Total Files Analyzed:** 51 Excel files
**Total Sheets Analyzed:** 700+ sheets across all files
**Total Unique Column Headers Found:** 166 unique column names

**Note:** The higher number of unique columns (166 vs initial estimate of 112) is due to comprehensive analysis of all sheets including those with inconsistent naming, typos, and lab-specific custom columns.

---

## File Locations Analyzed

### 1. Template File
- `BLANK TEMPLATE - PC2 Lab Sample Recording Template for -80 Freezers_May 2025.xlsx`

### 2. Freezer 1 Files
- `-80 freezer 1 log/` directory (7 files)
  - Including rack-level and box-level logs
  - Subdirectory: `-80 Freezer1 log_tube level/` (1 file)

### 3. Freezer 2 Files
- `-80 freezer2 log/` directory (1 box-level file)
  - Subdirectory: `-80 freezer2 log_tube level/` with shelves A, B, C, D, E (30 files)

### 4. Freezer 3 Files
- `-80 freezer3 log/` directory (11 files)

### 5. Other Files
- `Freezer allocations Walk-in freezer PC2 Synbio.xlsx`

---

## Column Frequency Analysis

### CRITICAL UPDATE: Top Columns Based on Complete Analysis

**The most frequently used columns across ALL 700+ sheets are:**

| Rank | Column Name | Frequency (Sheets) | Files | Usage Description |
|------|-------------|-------------------|-------|-------------------|
| 1 | **Organism** | 716 | 42 | Type of organism (E.coli, yeast, etc.) |
| 2 | **Owner** | 708 | 43 | Person responsible for samples |
| 3 | **Tube** | 708 | 41 | Tube position identifier (A-H, 1-12) |
| 4 | **Lab group** | 707 | 42 | Research group/PI |
| 5 | **Biosafety #** | 706 | 41 | Biosafety notification number |
| 6 | **GMO? Yes/No** | 706 | 41 | GMO status indicator |
| 7 | **Sample information** | 699 | 43 | Detailed sample description |
| 8 | **Project information** | 662 | 41 | Associated project details |

### Important Discovery: Two Major File Format Categories

The analysis reveals that the freezer logs use TWO distinct systems:

**CATEGORY A: Modern/Current Format (700+ sheets)**
- Uses: Organism, Biosafety #, GMO? Yes/No, Sample information, Project information
- Found in: Most freezer 2 and freezer 3 tube-level logs
- More comprehensive and standardized

**CATEGORY B: Legacy/Template Format (10-20 sheets)**
- Uses: NLRD, Risk Group, Box Row, Box Column, Contents
- Found in: Template file, some freezer 1 logs, box-level logs
- Matches the BLANK TEMPLATE structure

### Most Common Columns (Used in 10+ Files/Sheets)

| Column Name | Frequency | Usage Description |
|-------------|-----------|-------------------|
| **Organism** | 716 sheets | Type of organism (E.coli, S. cerevisiae, etc.) |
| **Owner** | 708 sheets | Person responsible for samples |
| **Tube** | 708 sheets | Tube position identifier |
| **Lab group** | 707 sheets | Research group/PI |
| **Biosafety #** | 706 sheets | Biosafety notification number |
| **GMO? Yes/No** | 706 sheets | GMO status (Yes/No) |
| **Sample information** | 699 sheets | Detailed sample description |
| **Project information** | 662 sheets | Project details/origin |
| **Comments** | 76 sheets | Additional notes |
| **Marker** | 68 sheets | Genetic markers |
| **Tube number** | 52 sheets | Numbered tube identifier |
| **Sequencing** | 37 sheets | Sequencing status/results |
| **vector** (lowercase) | 37 sheets | Vector information |
| **Vector** (capitalized) | 25 sheets | Vector information |
| **NLRD** | 20 sheets | Notifiable Low Risk Dealing |
| **Shelf** | 18 sheets | Freezer shelf location |
| **Risk Group** | 18 sheets | Biosafety risk classification |
| **Rack Number** | 13 sheets | Rack identifier |
| **Label on Box** | 13 sheets | Box label description |
| **Box position** | 14 sheets | Position of box in rack |
| **Date lodged (dd/mm/yy)** | 13 sheets | Date sample was added |
| **Box Row** | 13 sheets | Row position in box |
| **Box Column** | 13 sheets | Column position in box |
| **Freezer Location** | 12 sheets | Physical freezer location code |
| **Sample lodged by** | 12 sheets | Person who logged the sample |
| **Contents** | 12 sheets | Description of sample contents |

### Moderately Common Columns (5-9 Files)

| Column Name | Frequency | Usage Description |
|-------------|-----------|-------------------|
| **1** | 9 sheets | Layout guide marker |
| **Box** | 7 sheets | Box identifier |
| **Box description** | 7 sheets | Description of box contents |
| **Strain** | 7 sheets | Organism strain information |
| **Organism** | 6 sheets | Organism type |
| **Yes** | 6 sheets | Boolean field (various uses) |
| **Plasmid** | 6 sheets | Plasmid information |
| **Freezer location** | 5 sheets | Physical freezer location (lowercase) |
| **Rack number** | 5 sheets | Rack identifier (lowercase) |
| **Label on box** | 5 sheets | Box label (lowercase) |
| **Box location** | 5 sheets | Box position in rack |
| **Number** | 5 sheets | Sample number |
| **Name** | 5 sheets | Sample/strain name |
| **Vector** | 5 sheets | Vector information |
| **Gene** | 5 sheets | Gene information |
| **Gene Organism** | 5 sheets | Source organism for gene |
| **Resistances** | 5 sheets | Antibiotic resistances |
| **Additional Information** | 5 sheets | Extra notes |
| **Origin** | 5 sheets | Sample origin |
| **Position** | 5 sheets | Position identifier |

### Scientific/Research-Specific Columns (2-4 Files)

| Column Name | Frequency | Usage Description |
|-------------|-----------|-------------------|
| **Cyro Label** | 4 sheets | Cryogenic tube label |
| **Source** | 4 sheets | Sample source |
| **Note** | 4 sheets | Additional notes |
| **Sample information** | 3 sheets | Detailed sample info |
| **Project information** | 3 sheets | Associated project |
| **Freezer** | 3 sheets | Freezer identifier |
| **Rack/Shelf** | 3 sheets | Combined location |
| **Antibiotic Marker** | 3 sheets | Antibiotic resistance markers |
| **Verification** | 3 sheets | Verification status |
| **Related Accession Number** | 3 sheets | Database accession |
| **Reference** | 3 sheets | Literature/protocol reference |
| **RiskGroup** | 2 sheets | Biosafety risk (no space) |
| **Tube ID** | 2 sheets | Unique tube identifier |
| **Marker yeast** | 2 sheets | Yeast genetic markers |
| **Verified** | 2 sheets | Verification status |
| **Marker bacteria** | 1 sheet | Bacterial genetic markers |
| **Auxotrophy** | 1 sheet | Auxotrophic markers |
| **Sequencing** | 1 sheet | Sequencing results |
| **Sequencing result** | 1 sheet | Sequencing verification |

---

## Column Name Variations (Potential Standardization Needed)

### Case Sensitivity and Spacing Issues

| Normalized Form | Variations Found | Recommendation |
|----------------|------------------|----------------|
| **freezerlocation** | "Freezer Location" (10), "Freezer location" (5) | Standardize to "Freezer Location" |
| **labelonbox** | "Label on Box" (11), "Label on box" (5) | Standardize to "Label on Box" |
| **labgroup** | "Lab group" (675), "Lab Group" (1) | Standardize to "Lab group" |
| **owner** | "Owner" (676), "OWNER" (1) | Standardize to "Owner" |
| **racknumber** | "Rack Number" (11), "Rack number" (5) | Standardize to "Rack Number" |
| **riskgroup** | "Risk Group" (11), "RiskGroup" (2) | Standardize to "Risk Group" |

---

## Distinct File Format Types

### FORMAT TYPE 1: Simple Tube-Level Format (669 sheets)
**Most Common Format**

**Columns (3):**
1. Lab group
2. Owner
3. Tube

**Usage:** Basic tube-level tracking
**Example Files:** Most freezer 2 and freezer 3 rack files

**Sample Data Structure:**
```
Owner          | Lab group        | Tube
Rashika Sood   | Paulsen         | A
Rashika Sood   | Paulsen         | B
```

---

### FORMAT TYPE 2: Enhanced Tube-Level Format (5 sheets)
**Columns (8):**
1. Box location
2. Freezer location
3. Lab group
4. Label on box
5. Owner
6. Rack number
7. Shelf
8. Tube

**Usage:** Detailed tube tracking with full location hierarchy
**Example Files:** `rack A4 -80C freezer 3_AfrinTalukder&VivianBonacker.xlsx`

**Sample Data Structure:**
```
Owner              | Lab group         | Freezer location | Shelf | Rack number | Label on box | Box location | Tube
Vivian Bonacker    | Briardo's group   | 6WW-322         | A     | A4          | Box 1, S. cerevisiae... | 1    | A
```

---

### FORMAT TYPE 3: Box Numbering Guide (4 sheets)
**Columns (7):**
1. Front of rack
2. 1
3. 5
4. 9
5. 13
6. 17
7. Back of rack

**Usage:** Visual layout guide for box numbering in racks
**Example Files:** Box level logs

---

### FORMAT TYPE 4: Detailed Box-Level Format (4 sheets)
**Columns (13):**
1. Box Column
2. Box Row
3. Box description
4. Box position
5. Contents
6. Date lodged (dd/mm/yy)
7. Freezer Location
8. Label on Box
9. NLRD
10. Rack Number
11. Risk Group
12. Sample lodged by
13. Shelf

**Usage:** Comprehensive box-level tracking with biosafety info
**Example Files:** `-80C Freezers_1_Rack_D3_KM.csv.xlsx`

---

### FORMAT TYPE 5: Strain Collection Format (3 sheets)
**Columns (12):**
1. Additional Information
2. Box
3. Gene
4. Gene Organism
5. Name
6. Number
7. Organism
8. Origin
9. Position
10. Resistances
11. Strain
12. Vector

**Usage:** Detailed genetic/strain information
**Example Files:** `Rack C3 -80C freezer3_JordiPerez.xlsx` (strain collection sheets)

---

### FORMAT TYPE 6: Advanced Bacterial Strain Format (3 sheets)
**Columns (20):**
1. Antibiotic Marker
2. Box Column
3. Box Row
4. Box position
5. Contents
6. Cyro Label
7. Date lodged (dd/mm/yy)
8. Freezer Location
9. Label on Box
10. NLRD
11. Note
12. Plasmid
13. Rack Number
14. Reference
15. Related Accession Number
16. Risk Group
17. Sample lodged by
18. Shelf
19. Source
20. Verification

**Usage:** Highly detailed bacterial strain tracking
**Example Files:** `-80 freezer 1_rack D1_EA&JP.xlsx` (bacterial boxes)

---

### FORMAT TYPE 7: Layout Guide (2 sheets)
**Columns (9):**
- A1 through A9

**Usage:** Visual layout reference

---

### FORMAT TYPE 8: Enhanced Strain Format with Location (2 sheets)
**Columns (16):**
1. Additional Information
2. Box
3. Freezer
4. Gene
5. Gene Organism
6. NLRD
7. Name
8. Number
9. Organism
10. Origin
11. Position
12. Rack/Shelf
13. Resistances
14. RiskGroup
15. Strain
16. Vector

**Usage:** Strain information with full location tracking

---

### Additional Format Types (9-26)
Various specialized formats with 1-2 instances each, including:
- Specific organism tracking (Synechocystis, E. coli)
- Yeast-specific formats with auxotrophy markers
- Plasmid-specific tracking
- Custom lab-specific formats

---

## BLANK TEMPLATE File Structure

**File:** `BLANK TEMPLATE - PC2 Lab Sample Recording Template for -80 Freezers_May 2025.xlsx`

**Template Sheet Columns (13):**
1. Freezer Location
2. Shelf
3. Rack Number
4. Label on Box
5. Box position
6. Sample lodged by
7. Date lodged (dd/mm/yy)
8. Box description
9. Box Row
10. Box Column
11. Contents
12. Risk Group
13. NLRD

**Sample Data Example:**
```
Freezer Location: 14ER 227
Shelf: B
Rack Number: B1
Label on Box: E.coli Glycerol Stocks AGF15
Box position: 1
Sample lodged by: Kristopher Montrose
Date lodged: 2025-02-02
Box description: E.Coli
Box Row: A
Box Column: 1
Risk Group: 1
NLRD: Exempt
```

**Additional Sheets:**
- Shelf 1 Rack 1 through Shelf 1 Rack 6 (empty template sheets)

---

## Complete List of All 166 Unique Columns

**NOTE:** A detailed CSV file with all columns and their frequencies has been created:
`column_frequency_analysis.csv`

This file contains:
- Column name
- Total frequency (number of sheets using it)
- Number of unique files using it
- Example files
- Normalized name (for identifying variations)

## All Unique Columns Organized by Category

### Location and Identification Columns
1. Freezer Location / Freezer location
2. Shelf
3. Rack Number / Rack number / Rack/Shelf
4. Box / Box location
5. Box position
6. Tube
7. Tube number
8. Position
9. Box Row
10. Box Column

### Sample Information Columns
11. Owner / OWNER
12. Lab group / Lab Group
13. Sample lodged by
14. Date lodged (dd/mm/yy)
15. Contents
16. Sample information
17. Project information

### Box/Container Labeling
18. Label on Box / Label on box
19. Box description

### Biosafety and Compliance
20. Risk Group / RiskGroup
21. NLRD
22. Contains GMOs? (Yes/No)

### Organism Information
23. Organism
24. Strain
25. Gene
26. Gene Organism
27. Vector
28. Plasmid

### Genetic Markers
29. Resistances
30. Antibiotic Marker
31. Marker bacteria
32. Marker yeast
33. Auxotrophy

### Sample Tracking
34. Number
35. Name
36. Tube ID
37. Cyro Label

### Additional Information
38. Additional Information
39. Note / Notes
40. Comments
41. Origin
42. Source
43. Reference
44. Related Accession Number

### Verification and Quality
45. Verification / Verified
46. Sequencing / Sequencing result

### Layout/Position Guides
47. Front of rack
48. Back of rack
49. 1, 5, 9, 13, 17 (box position numbers)
50. A1, A2, A3, A4, A5, A6, A7, A8, A9 (grid positions)

### Specialized/Lab-Specific Columns (51-112)
Including various:
- Specific researcher names
- Specific organism identifiers
- Lab-specific codes
- Custom tracking fields
- Freezer allocation columns

---

## Key Observations and Recommendations

### 1. Format Fragmentation
- **Issue:** 26 distinct format types identified
- **Impact:** Inconsistent data structure across freezers
- **Recommendation:** Standardize to 2-3 core formats based on use case

### 2. Column Name Standardization Needed
- Multiple variations of the same field (case, spacing)
- Recommend establishing naming conventions

### 3. Three Main Use Cases Identified

**A. Tube-Level Tracking (Simple)**
- Essential: Owner, Lab group, Tube
- Optional: Box location, Freezer location, Shelf, Rack number

**B. Box-Level Tracking (Intermediate)**
- All from Tube-Level, plus:
- Box description, Box position, Sample lodged by, Date lodged
- Risk Group, NLRD

**C. Detailed Scientific Tracking (Complex)**
- All from Box-Level, plus:
- Organism, Strain, Gene, Vector, Plasmid
- Resistances, Markers
- Verification, Sequencing, Reference
- Additional Information, Notes

### 4. Essential Columns for Standardization
Based on frequency and importance:
1. Owner (676 uses)
2. Lab group (675 uses)
3. Tube/Position (674 uses)
4. Freezer Location
5. Shelf
6. Rack Number
7. Box position
8. Date lodged
9. Risk Group
10. NLRD

### 5. Data Quality Issues Observed
- Inconsistent date formats
- Mixed case in field names
- Some files have data in headers
- Empty columns in some formats

---

## Conclusion

The analysis reveals significant variation in freezer log formats across the organization. While the most common format (669 sheets) is simple and consistent, there are 25 other format variations that make data consolidation challenging.

**Priority Actions:**
1. Standardize column naming conventions
2. Consolidate formats to 2-3 standard templates
3. Establish required vs. optional fields
4. Implement data validation for key fields
5. Create migration plan for existing data

The BLANK TEMPLATE file provides a good starting point with 13 core columns that cover most essential tracking needs.

---

## COMPLETE ALPHABETICAL LIST OF ALL 166 UNIQUE COLUMNS

Below is the complete list with frequency data:

### High Frequency Columns (100+ uses)
1. **Organism** - 716 sheets - Type of organism
2. **Owner** - 708 sheets - Sample owner
3. **Tube** - 708 sheets - Tube position
4. **Lab group** - 707 sheets - Research group
5. **Biosafety #** - 706 sheets - Biosafety number
6. **GMO? Yes/No** - 706 sheets - GMO status
7. **Sample information** - 699 sheets - Sample details
8. **Project information** - 662 sheets - Project details

### Medium Frequency Columns (10-99 uses)
9. **Comments** - 76 sheets - General comments
10. **Marker** - 68 sheets - Genetic markers
11. **Tube number** - 52 sheets - Tube identifier
12. **Sequencing** - 37 sheets - Sequencing info
13. **vector** - 37 sheets - Vector (lowercase)
14. **Vector** - 25 sheets - Vector (capitalized)
15. **NLRD** - 20 sheets - Notifiable Low Risk Dealing
16. **Shelf** - 18 sheets - Shelf location
17. **Risk Group** - 18 sheets - Biosafety risk level
18. **Project information/origin/plasmid from** - 17 sheets
19. **Name** - 16 sheets - Sample name
20. **Date of congelation** - 16 sheets - Freezing date
21. **Box position** - 14 sheets - Box location in rack
22. **Rack Number** - 13 sheets - Rack identifier
23. **Label on Box** - 13 sheets - Box label
24. **Date lodged (dd/mm/yy)** - 13 sheets - Date added
25. **Box Row** - 13 sheets - Row in box
26. **Box Column** - 13 sheets - Column in box
27. **Date** - 13 sheets - General date field
28. **Freezer Location** - 12 sheets - Freezer identifier
29. **Sample lodged by** - 12 sheets - Person who logged
30. **Contents** - 12 sheets - Sample contents
31. **1** - 12 sheets - Position marker
32. **Strain** - 11 sheets - Organism strain
33. **PCR** - 10 sheets - PCR information

### Low Frequency Columns (2-9 uses)
34. **Alias** - 10 sheets
35. **Box description** - 9 sheets
36. **Tube label** - 9 sheets
37. **Box** - 7 sheets
38. **Front of rack** - 7 sheets
39. **5, 9, 13, 17** - 7 sheets each - Position markers
40. **Back of rack** - 7 sheets
41. **Reference** - 7 sheets
42. **Genotyping|Sequencing** - 6 sheets
43. **Yes** - 6 sheets
44. **Plasmid** - 6 sheets
45. **Date of sequencing...** - 6 sheets
46. **Relevant genotype** - 6 sheets
47. **Comment** - 5 sheets
48. **Freezer location** (lowercase) - 5 sheets
49. **Rack number** (lowercase) - 5 sheets
50. **Label on box** (lowercase) - 5 sheets
51. **Box location** - 5 sheets
52. **Exp ID** - 5 sheets
53. **Additional comments** - 5 sheets
54. **Number** - 5 sheets
55. **Gene** - 5 sheets
56. **Gene Organism** - 5 sheets
57. **Resistances** - 5 sheets
58. **Additional Information** - 5 sheets
59. **Origin** - 5 sheets
60. **Position** - 5 sheets
61. **RG** - 5 sheets
62. **A** - 5 sheets
63. **Cyro Label** - 4 sheets
64. **Source** - 4 sheets
65. **Note** - 4 sheets
66. **strain** (lowercase) - 4 sheets
67. **Antibiotics** - 4 sheets
68. **Additional comment** - 4 sheets
69. **Genotype** - 4 sheets
70. **Other information** - 4 sheets
71. **Freezer** - 3 sheets
72. **Rack/Shelf** - 3 sheets
73. **correct** - 3 sheets
74. **Antibiotic Marker** - 3 sheets
75. **Verification** - 3 sheets
76. **Related Accession Number** - 3 sheets
77. **Shelf/ Rack#** - 2 sheets
78. **Contains GMOs? (Yes/No)** - 2 sheets
79. **A1, A2, A3, A4, A5, A6, A7, A8, A9** - 2 sheets each
80. **Transformed/Empty** - 2 sheets
81. **Genotyping** - 2 sheets
82. **RiskGroup** - 2 sheets
83. **Tube ID** - 2 sheets
84. **Marker yeast** - 2 sheets
85. **Verified** - 2 sheets
86. **Auxotrophy** - 2 sheets
87. **Resistence** - 2 sheets

### Single Use Columns (1 use each)
These include highly specialized, lab-specific, or potentially erroneous entries:

88. **6B.1, 6B.2, 6B.3** - Freezer allocation grid
89. **OWNER** (all caps) - Variation of Owner
90. **Code** - Sample code
91. **Characteristics** - Sample characteristics
92. **Culture conditions** - Growth conditions
93. **Contain** - Contains field
94. **state** - Sample state
95. **Sequencing result** - Sequencing outcome
96. **Competent cells (chemically)** - Cell type
97. **Lab Group** (title case) - Variation
98. **Sequence** - DNA sequence
99. **comments** (lowercase) - Comment variation
100. **Coments** - Typo of Comments
101. **EXEMPT** - Exemption status
102. **Auxotrophy recovery** - Recovery information
103. **Lid label** - Label on box lid
104. **Sample info** - Info variation
105. **KIT** - Kit information
106. **Project info** - Info variation
107. **Culture condition** - Growth condition
108. **Sequenced strains** - Sequencing status
109. **Class** - Classification
110. **enzyme (on pMU-hph)** - Specific enzyme
111. **Marker bacteria** - Bacterial markers
112. **Key** - Legend/key information

### Lab/Project-Specific Columns (appears to be data in headers)
These appear to be specific to individual researchers or projects:

113-166. Various researcher names, specific organism identifiers, plasmid names, biosafety codes, dates, and project-specific tracking fields including:
- **Artur Sawicki, Briardo LLorente, Sawicki** - Researcher names
- **Escherichia coli Top10 beta** - Specific strains
- **Synechocystis sp. PCC 6803 NS3** - Cyanobacteria strains
- **5201401062, 5201401059** - Biosafety/tracking numbers
- **pET28a p15 HPT alpha 1.1 DH10 beta** - Specific plasmids
- **DMSO stock, DMSO Stock p560NS3 HPTPili#2** - Stock types
- **Kan, CAMr** - Antibiotic abbreviations
- Various date stamps (27.10.25)
- Box-specific labels and project codes

---

## Critical Column Name Variations That Need Standardization

| Base Field | Variations Found | Recommendation |
|------------|------------------|----------------|
| **Owner** | Owner (708), OWNER (1) | Standardize to "Owner" |
| **Lab group** | Lab group (707), Lab Group (1) | Standardize to "Lab group" |
| **Vector** | vector (37), Vector (25) | Standardize to "Vector" |
| **Strain** | Strain (11), strain (4) | Standardize to "Strain" |
| **Freezer Location** | Freezer Location (12), Freezer location (5) | Standardize to "Freezer Location" |
| **Label on Box** | Label on Box (13), Label on box (5) | Standardize to "Label on Box" |
| **Rack Number** | Rack Number (13), Rack number (5) | Standardize to "Rack Number" |
| **Risk Group** | Risk Group (18), RiskGroup (2) | Standardize to "Risk Group" |
| **Comments** | Comments (76), comments (1), Comment (5), Coments (1) | Standardize to "Comments" |
| **Sequencing** | Sequencing (37), Sequencing result (1) | Standardize to "Sequencing" |
| **Sample Information** | Sample information (699), Sample info (1) | Standardize to "Sample information" |
| **Project Information** | Project information (662), Project info (1) | Standardize to "Project information" |

---

