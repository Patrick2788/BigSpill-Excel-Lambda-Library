## BigSpill User Manual

## Table of Contents
1. Overview
2. Installation & Setup
3. Function Reference
4. Errors
5. Usage Patterns
6. Multi-Step Examples
7. Performance Notes
8. Versioning Notes
   Appendix
   	A. Onboarding Workbooks
    B. Additional Resources


---
## 1. Overview
content

---
## 2. Installation & Setup

### 2.1 Install via BigSpill Template (Recommended)
The [BigSpill Template](<templates/BigSpill_Template.xlsx>) is the fastest way to install the full library.

Save it to a safe location on your computer and double‑click it to create a new workbook. This new workbook will come preloaded with the entire BigSpill function catalog.

### 2.2 Verifying BigSpill is Installed
1. Double-click BigSpill_Template.xltx to open a new workbook.
2. In any cell enter: `=Shiftλ(123,1)`.
3. If the functions spills a 3 x 1 array of: {2; 3; 1} then BigSpill is installed. `#NAME?` would indicate the workbook was not created from the BigSpill Template.

---
## 3. Function Reference
This section provides a complete, module‑ordered catalog of all BigSpill functions, along with short descriptions and structural groupings.

---
### Gridwork
The foundational coordinate systems for 2D arrays. Each function produces a centered, symmetrical geometric description of the grid. 

| Function | Description |
|---------|-------------|
| **DGridλ**   |    Direction grid (0°-360°) using centered Cartesian axes |
| **CGridλ**   |    Cartesian grid (row or column coordinates, centered) |
| **MGridλ**   |    Mesh grid (1‑based row or column indices) |
| **TGridλ**   |    Thresholded mesh grid (indices wrap after a limit) |
| **PolarGridλ**	| Polar coordinates: radial distance and angle (`ATAN2`) |

---
### Grid Geometry
A family of geometric extraction operators for 2D arrays. Each function returns a region of the grid defined by a specific shape, preserving the geometry of the extracted area.

| Function | Description |
|---------|-------------|
| **Circleλ** | Circular region by target or {row, col} center |
| **Diamondλ** | Manhattan‑radius diamond |
| **Plusλ** | Cross‑shaped region (von Neumann neighborhood) |
| **Pyramidλ** | Directional half‑diamond (four pieces) |
| **Ringλ** | Layer‑based perimeter extraction |
| **Squareλ** | Chebyshev‑radius square |
| **Triangleλ** | Directional triangular region (above or below) |

---
### 2D Array Shaping
Shaping functions reshape, resize, wrap, pad, traverse, or reflow grids while preserving the structure of the underlying data.

#### Core Shaping Primitives

| Function | Description |
| --- | --- |
| **ReShape2Dλ** | Reshapes a 2D array into a new column width using row‑major order |
| **ReSizeλ** | Repeats a 2D array vertically and/or horizontally by integer repeat counts |
| **Scan2Dλ** | Traversal‑aware 2D scanning operator extending Excel’s native `SCAN` |
| **Traverseλ** | Remaps a 2D array into a new traversal order using diagonal directions |
| **TakeBlockλ** | Extracts a rectangular block using start_at and optional stop_when |
| **DStackλ** | Stacks multiple 2D arrays along a depth dimension |

#### Reflow

| Function | Description |
| --- | --- |
| **WrapRows2Dλ** | Wraps into row blocks of specified height |
| **WrapCols2Dλ** | Wraps into column blocks of specified width |
| **Staircaseλ** | Constructs diagonal staircase patterns |
| **UnPivotλ** | Converts a matrix into a flattened row‑wise table |
| **Zipλ** | Zips 1–4 arrays by flattening and aligning by length |
| **SliceByDegλ** | Extracts angular slices using degree bounds (0°–360°) |
		
#### Padding

| Function | Description |
| --- | --- |
| **Padλ** | Applies row and/or column padding to any array |
| **InsulateRowsλ** | Adds rows of padding before and/or after an input array |
| **InsulateColsλ** | Adds columns of padding before and/or after an input array |

#### Interpolation

| Function | Description |
| --- | --- |
| **Magnifyλ** | Expands each element into a block of specified height × width using nearest‑neighbor interpolation |
| **Zoomλ** | Expands each element into a block whose size matches the height × width of `zoom_kernel` |

---
### Grid Algebra
A toolkit for structural editing of 2D arrays. These operators support deletion, selection, filtering, rolling, expansion, and structural mapping.

#### Deletion

| Function | Description |
| --- | --- |
| **DeleteRowsλ** | Deletes rows from a 2D array based on contiguous subgroup runs |
| **DeleteColsλ** | Deletes columns from a 2D array based on contiguous subgroup runs |
| **DeleteWhereλ** | Removes vertical and horizontal vectors based on a predicate Lambda |
| **Pinchλ** | Removes rows/columns whose cells all have `LEN(cell) `= `0` |
| **Removeλ** | Removes specific rows and/or columns |
| **Squeezeλ** | Removes null and straight‑zero vectors |
		
#### Expansion

| Function | Description |
| --- | --- |
| **KroneckerProdλ** | Computes the Kronecker product of two arrays |
		
#### Filtering

| Function | Description |
| --- | --- |
| **Excludeλ** | Filters an array by excluding any row containing a specified criterion |
| **Snapλ** | Selects rows containing at least one value from the include‑criteria |
		
#### Rolling

| Function | Description |
| --- | --- |
| **Shiftλ** | Circular shift on scalar, 1D, or 2D arrays |
| **ShiftRowsλ** | Circular horizontal shift per row |
| **ShiftColsλ** | Circular vertical shift per column |
		
#### Selection

| Function | Description |
| --- | --- |
| **Embedλ** | Embeds an array into a 2D host at a specified row or column index |
| **KeepRowsλ** | Retains rows based on contiguous subgroup runs |
| **KeepColsλ** | Retains columns based on contiguous subgroup runs |
		
#### Structural Mapping

| Function | Description |
| --- | --- |
| **BlockMapλ** | Reshapes a matrix row‑wise or column‑wise into blocks of size depth × width |
| **Convolveλ** | Performs sliding‑window convolution using a 2D kernel |
| **Foldλ** | Groups identical rows and aggregates values from `values_array` |
| **Join2Dλ** | Joins non‑empty values across 2–4 identically‑shaped arrays using a delimiter |

---
### Repetition
Functions that replicate arrays along rows, columns, or both.

| Function | Description |
| --- | --- |
| **Echoλ** | Repeats each element in a 1D array according to repeat counts |
| **RepeatRowsλ** | Repeats each row according to repeat counts |
| **RepeatColsλ** | Repeats each column according to repeat counts |
| **Tessellateλ** | Generates a modular tessellation of a 2D pattern defined by `new_height `× `new_width` |

---
### Grid Analytics
A collection of functions for elegant sequencing, categorical analysis, neighborhood‑based aggregation, and statistics. 

| Function | Description |
| --- | --- |
| **Streakλ** | Computes streak counts of consecutive identical values |
| **Countdownλ** | Generates countdown indices for contiguous runs |

#### Categorical Analysis

| Function | Description |
| --- | --- |
| **ByDiagλ** | Applies an ETA function or Lambda to each diagonal |
| **DiagMapλ** | Extracts all diagonals and returns them as a 2D matrix |
| **DiagIndexλ** | Extracts a single diagonal or anti‑diagonal vector |
| **Pairwiseλ** | Creates vertical pairs from 1D and scalar inputs |
| **GroupbyBinλ** | Aggregates values into fixed‑width numeric bins |
| **GroupbyDateλ** | Aggregates values by date/time interval (minute → year) |
| **PivotbyCatλ** | Pivots data by categories and values |
| **Grainλ** | Downsamples a matrix by aggregating non‑overlapping spatial blocks |

#### Neighborhoods

| Function | Description |
| --- | --- |
| **MooreAggλ** | Aggregates over Moore or von Neumann neighborhoods using ETA functions |
| **MooreSelectλ** | Selects a specific neighbor (1–9) from the Moore neighborhood |
		
#### Statistics

| Function | Description |
| --- | --- |
| **Histogramλ** | Generates a histogram using automatic binning rules |
| **Modeλ** | Returns the most frequently occurring values or text entries |
| **ZoneStatλ** | Aggregates values by zones determined by `Histogramλ` |

---
### Text
This category provides operators that overcome common limitations in Excel's native text engine (e.g., `TEXTSPLIT` cannot spill 2D results; `BYROW` cannot deploy `SORT` to align columns). These functions enable structural alignment, extraction, and transformation of text arrays, extending Excel's capabilities for both 1D and 2D inputs.

| Function | Description |
| --- | --- |
| **Alignλ** | Sorts each row (or column when `by_col=TRUE`) alphabetically |
| **AlignDistinctλ** | Sorted distinct values |
| **AlignUniqueλ** | Sorted unique values |
| **Coalesceλ** | Returns the first non‑empty value across 2–4 arrays |
| **Explodeλ** | Converts scalar, 1D, or 2D input into exploded character array |
| **NumbersOnlyλ** | Extracts numeric characters (0–9) |
| **TextOnlyλ** | Extracts non‑numeric characters |
| **Splitλ** | Splits text arrays into tokens and can return a 2D spill |

---
### Engineering
Binary, Gray‑code, and bit‑level utilities designed to operate well beyond Excel's native `DEC2BIN` limit of 511. All functions support concise binary strings and shape‑preserving array behavior, with optional 2D exploded bit‑matrices where applicable.

| Function | Description |
| --- | --- |
| **Bin2Decλ** | Converts concise binary → decimal |
| **Bin2Grayλ** | Converts binary → Gray code |
| **BitCountλ** | Counts number of 1‑bits |
| **Dec2Binλ** | Converts non‑negative integers → binary |
| **Dec2Grayλ** | Converts non‑negative integers → Gray code |

---
### Combinatorics
High‑performance generators for permutations, combinations, subsets, and related structures. 

| Function | Description |
| --- | --- |
| **PermRλ** | Permutations with replacement |
| **Permλ** | Permutations without replacement |
| **CombinationsRλ** | Combinations with replacement |
| **Combinationsλ** | Combinations without replacement |
| **Derangementsλ** | Permutations where no element remains in its original position |
| **KnapSackλ** | Solves subset‑sum / knapsack feasibility |
| **SubsetGenλ** | Generates all subsets (power set) in binary‑mask order |
| **SubsetSumλ** | Returns all subsets whose elements sum to a target |

---
### Diagnostics
Functions whose purpose is to expose, annotate, or explain the internal structure of a grid. Diagnostics functions are used for debugging, introspection, visualization, and structural analysis.

| Function | Description |
| --- | --- |
| **Revealλ** | Returns either a semantic summary (Numbers, Texts, Errors, etc.) or a verbose cell‑by‑cell narration |

---
## 5. Errors
content

---
## 6. Usage Patterns
content

---
## 7. Multi-Step Examples
content

---
## 8. Performance Notes
content

---
## 9. Versioning Notes
content

---
## Appendix

### Onboarding Workbooks
## Appendix

### A. Onboarding Workbooks
The following workbooks provide demonstrations of all functions in BigSpill. 
They are listed in the recommended learning order:

- [00 – BigSpill – Starters](<onboarding/00 - BigSpill - Starters.xlsx>)
- [01 – BigSpill – Gridwork and Grid Geometry](<onboarding/01 - BigSpill - Gridwork and Grid Geometry.xlsx>)
- [02 – BigSpill – 2D Array Shaping](<onboarding/02 - BigSpill - 2D Array Shaping.xlsx>)
- [03 – BigSpill – Grid Algebra](<onboarding/03 - BigSpill - Grid Algebra.xlsx>)
- [04 – BigSpill – Repetition](<onboarding/04 - BigSpill - Repetition.xlsx>)
- [05 – BigSpill – Grid Analytics](<onboarding/05 - BigSpill - Grid Analytics.xlsx>)
- [06 – BigSpill – Text](<onboarding/06 - BigSpill - Text.xlsx>)
- [07 – BigSpill – Engineering](<onboarding/07 - BigSpill - Engineering.xlsx>)
- [08 – BigSpill – Combinatorics](<onboarding/08 - BigSpill - Combinatorics.xlsx>)
- [09 – BigSpill – Diagnostics](<onboarding/09 - BigSpill - Diagnostics.xlsx>)







