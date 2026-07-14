<p align="left">
  <img src="../assets/BigSpillLogo.png" width="320">
</p>

# User Manual

## Table of Contents
1. Welcome
2. Installation & Setup
3. Function Reference
4. Errors
5. Usage Patterns
6. Multi-Step Examples
7. Performance Notes
8. Versioning Notes
9. Appendix

---
## 1. Welcome
BigSpill builds on the innovations Excel introduced between 2018 and 2021 (Dynamic Arrays, LAMBDA, LET) and the plethora of new functions added in the past 8 years. This library exists to bring the power of Excel, where most anything is possible (within reason), to the user without the complexity.

BigSpill's audience is everyone from beginners who want easy-to-use and reliable functions to developers that want building blocks. Whether you learn by exploring or by reading step‑by‑step, you’re in the right place.

---
## 2. Installation & Setup

### 2.1 Install via BigSpill Template (Recommended)
The [BigSpill Template](<templates/BigSpill_Template.xlsx>) is the fastest way to install the full library.

Save it to a safe location on your computer and double‑click it to create a new workbook. This new workbook will come preloaded with the entire BigSpill function catalog.

### 2.2 Verifying BigSpill is Installed
1. Double-click BigSpill_Template.xltx to open a new workbook.
2. In any cell enter: `=Shiftλ(123,1)`.
3. If the functions spills a 3 x 1 array of: `{2; 3; 1}` then BigSpill is installed. `#NAME?` would indicate the workbook was not created from the BigSpill Template.

### 2.3 Adding BigSpill to an Existing Workbook
If you want to use BigSpill inside an existing .xlsx, the easiest method is to copy a sheet from a BigSpill‑enabled workbook. This automatically transfers all BigSpill modules and supporting names.

Steps:
1. Open a new workbook created from BigSpill_Template.xltx.
2. Open the workbook that will receive BigSpill.
3. Right‑click the sheet tab in the BigSpill workbook → Move or Copy…
4. Choose the destination workbook.
5. Check Create a copy → OK.
6. To verify the transfer, enter: `=Shiftλ(123,1)`. If `{2;3;1}` spills, BigSpill is installed.

Note:
- Always copy from the BigSpill workbook into the workbook that will receive BigSpill (Once you've verified installation, the blank sheet can be deleted from your workbook).
- All BigSpill functions include a `λ` suffix, which prevents name conflicts when the above steps are performed.

### 2.4 Requesting a Gist import URL (Advanced)
BigSpill can also be installed using the Advanced Formula Environment (AFE) by importing a private Gist that contains the full module. This installation method is intended for developers.

The Gist URL is not included in this manual. Access may be granted upon request.

---
## 3. Function Reference
This section provides a complete, module‑ordered catalog of all BigSpill functions, along with short descriptions and structural groupings.

---
### Gridwork
The foundational coordinate systems for 2D arrays. Each function produces a centered, symmetrical geometric description of the grid. 

#### Dependency Hierarchy
```
Resizeλ
├── CGridλ
│   ├── DGridλ
│   └── PolarGridλ
├── MGridλ
└── TGridλ
```
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

#### Dependency Hierarchy
```
Squeezeλ
├── Circleλ
├── Diamondλ
├── Plusλ
├── Squareλ
└── Triangleλ

CGridλ
└── (shared dependency with Squeezeλ)

MGridλ
├── Pyramidλ
└── Ringλ
```
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

#### Dependency Hierarchy
```
ReShape2Dλ
└── Scan2Dλ

Traverseλ
└── Scan2Dλ

ReSizeλ
├── TakeBlockλ
├── UnPivotλ
└── Zoomλ

MGridλ
└── TakeBlockλ

Echoλ
├── WrapRows2Dλ
└── WrapCols2Dλ

ValidateStaircaseλ
└── Staircaseλ

RepeatRowsλ
└── UnPivotλ

InsulateRowsλ
├── UnPivotλ
└── Padλ

InsulateColsλ
└── Padλ

DGridλ
└── SliceByDegλ

DeleteWhereλ
└── SliceByDegλ

Padλ
└── (depends on InsulateRowsλ and InsulateColsλ)

Magnifyλ
└── (no dependencies)

Zipλ
└── (no dependencies)
```
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

#### Dependency Hierarchy
```
Streakλ
├── DeleteRowsλ
├── DeleteColsλ
└── KeepRowsλ
    └── KeepColsλ

Countdownλ
├── DeleteRowsλ
├── DeleteColsλ
└── KeepRowsλ
    └── KeepColsλ

DeleteWhereλ
└── SliceByDegλ
    └── (depends on DGridλ)

Pinchλ
└── (no dependencies)

Removeλ
└── (no dependencies)

Squeezeλ
└── (no dependencies)

Magnifyλ
└── KroneckerProdλ

ReSizeλ
├── KroneckerProdλ
├── ShiftRowsλ
├── ShiftColsλ
└── Convolveλ

Explodeλ
└── Shiftλ
    ├── ShiftRowsλ
    └── ShiftColsλ

MGridλ
├── ShiftRowsλ
└── ShiftColsλ

Excludeλ
└── (no dependencies)

Snapλ
└── (no dependencies)

Embedλ
└── (no dependencies)

ValidateBlockMapλ
└── BlockMapλ

Alignλ
└── Foldλ

Join2Dλ
└── (no dependencies)

```
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

#### Dependency Hierarchy
```
Echoλ
└── (no dependencies)

RepeatRowsλ
└── (no dependencies)

RepeatColsλ
└── (no dependencies)

TGridλ
└── Tessellateλ
```

| Function | Description |
| --- | --- |
| **Echoλ** | Repeats each element in a 1D array according to repeat counts |
| **RepeatRowsλ** | Repeats each row according to repeat counts |
| **RepeatColsλ** | Repeats each column according to repeat counts |
| **Tessellateλ** | Generates a modular tessellation of a 2D pattern defined by `new_height `× `new_width` |

---
### Grid Analytics
A collection of functions for elegant sequencing, categorical analysis, neighborhood‑based aggregation, and statistics.

#### Dependency Hierarchy
```
Streakλ
├── DeleteRowsλ
├── DeleteColsλ
└── KeepRowsλ
    └── KeepColsλ

Countdownλ
├── DeleteRowsλ
├── DeleteColsλ
└── KeepRowsλ
    └── KeepColsλ

ValidateDiagλ
├── ByDiagλ
├── DiagMapλ
└── DiagIndexλ

ByDiagλ
├── DiagMapλ
└── DiagIndexλ

Echoλ
└── Pairwiseλ

Resizeλ
└── Pairwiseλ

ValidateGDλ
└── GroupbyDateλ

RepeatRowsλ
└── PivotbyCatλ

Streakλ
└── PivotbyCatλ

Histogramλ
└── ZoneStatλ

Grainλ
└── (no dependencies)

MooreAggλ
└── (no dependencies)

MooreSelectλ
└── (no dependencies)

Modeλ
└── (no dependencies)

GroupbyBinλ
└── (no dependencies)
```

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

#### Dependency Hierarchy
```
RegexSafeλ
└── Splitλ
    └── Alignλ
        ├── AlignDistinctλ
        └── AlignUniqueλ

Explodeλ
├── NumbersOnlyλ
└── TextOnlyλ

MGridλ
└── Explodeλ

Resizeλ
└── Explodeλ

Coalesceλ
└── (no dependencies)

```

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

#### Dependency Hierarchy
```
Explodeλ
├── Bin2Decλ
└── Bin2Grayλ

Dec2Binλ
├── BitCountλ
└── Dec2Grayλ

Bin2Grayλ
└── Dec2Grayλ

Dec2Grayλ
└── (depends on Bin2Grayλ and Dec2Binλ)
```

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

#### Dependency Hierarchy
```
Enumerateλ
├── PermRλ
├── Permλ
├── CombinationsRλ
├── Combinationsλ
└── Derangementsλ

Dec2Binλ
├── KnapSackλ
├── SubsetGenλ
└── SubsetSumλ
```

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
## 4. Errors
BigSpill functions return text‑based error codes when inputs are invalid. These codes are designed to be more informative than Excel’s native errors and help users quickly identify the source of a problem.

| Error | Description |
|---------|-------------|
| #AXIS! | Invalid direction or axis code. |
| #BIN-WIDTH! | Bin width is negative or out of range. |
| #BLOCK-OFFSET! | Block offset is invalid. |
| #BLOCK-SIZE! | Block size is invalid. |
| #COLUMN-INDEX! | Invalid column index. |
| #DEGREE! | Invalid degree parameter. |
| #DELIMITER-ARRAY! | Multiple delimiters supplied where one was expected. |
| #DIMENSIONS! | Array shapes do not match, or height/width parameters are invalid. |
| #EMPTY-ARRAY! | A required array is empty. |
| #EMPTY-INPUT! | No data supplied. |
| #EMBED! | Invalid host array for embedding. |
| #FUNCTION! | Reducer or function is not a valid LAMBDA. |
| #HEIGHT! | Invalid height parameter. |
| #INPUT-MISMATCH! | Item and weight arrays differ in length. |
| #INTERVAL! | Invalid interval specification. |
| #INVALID-COL! | Column index out of bounds. |
| #INVALID-FUNCTION! | Invalid reducer or function. |
| #INVALID-ROW! | Row index out of bounds. |
| #KERNEL! | Kernel contains non‑numeric values or is larger than the input array. |
| #MATRIX! | Input is not a valid matrix. |
| #NEIGHBOR! | Invalid neighbor code (must be 1–9). |
| #NO-ANCHOR! | start_at anchor not found. |
| #NO-DATES! | No dates found in input. |
| #NO-REPEATS! | No repeated values found. |
| #NO-SLICE! | stop_when occurs before start_at. |
| #NO-SOLUTION! | No subset matches the target. |
| #NO-TIMES! | No times found in input. |
| #NOT-ARRAY! | Input is not an array. |
| #RADIUS! | Invalid radius parameter. |
| #ROW-LIMIT! | Spill exceeds Excel’s row limit. |
| #SCALAR! | A scalar was supplied where an array was required. |
| #SPILL-RISK! | Spill size exceeds safe limits. |
| #TARGET! | Target value not found. |
| #TEXT! | A text value was supplied where a number was required. |
| #TEXT-PAD! | Padding count is negative or text. |
| #TEXT-SET! | Set contains non‑numeric values. |
| #TEXT-SHIFT! | Invalid text shift input. |
| #TEXT-TARGET! | Target contains non‑numeric values. |
| #TOO-MANY-ITEMS! | Combinatorial explosion (too many items). |
| #VALUE-CAPACITY! | Capacity is invalid. |
| #VALUE-ITEM! | Item values are invalid. |
| #VALUE-WEIGHTS! | Weights are invalid. |
| #WIDTH! | Invalid width parameter. |




---
## 5. Usage Patterns
content

#### 5.1 Split a 1D array into a 2D array and sort row-wise

 `Strings` =

```text
{"1, 10, 2, 1, 5, 9, 7, 5, 8, 1";
"1, 9, 2, 6, 7, 10, 3, 1, 9, 1";
"3, 3, 10, 6, 4, 9, 4, 8, 3, 6";
"6, 4, 7, 6, 1, 2, 10, 10, 8, 7";
"6, 2, 1, 6, 5, 3, 10, 8, 1, 2"}
```

Formula:

```excel
=LET(
    arr, --Splitλ(Strings, ", "),
    sorted, Alignλ(arr),
    sorted
)
```
Output:

```text
={"1","1","1","2","5","5","7","8","9","10";
"1","1","1","2","3","6","7","9","9","10";
"3","3","3","4","4","6","6","8","9","10";
"1","2","4","6","6","7","7","8","10","10";
"1","1","2","2","3","5","6","6","8","10"}
```

#### 5.2 Diamond Sum

`Grid` =
```text
{9,6,9,4,5;
6,8,5,10,1;
3,1,1,10,7;
7,3,3,3,7;
2,2,2,5,1}
```
Formula:
```Excel
=SUM(+Diamondλ(Grid,{3,3},2))
```

Output:
`65`

The `+` is used to coerce the result of `Diamondλ`into an array before being evaluated by `SUM`.

#### 5.3 Generate a deck of cards from two vectors

`Rank` =
```text
{"A";"K";"Q";"J";10;9;8;7;6;5;4;3;2}
```
`Suit` =
```text
{"♠";"♥";"♦";"♣"}
```
Formula:
```Excel
=Pairwiseλ(Rank,Suit)
```
Output:
```text
{"A" , "♠";
 "A" , "♥";
 "A" , "♦";
 "A" , "♣";
 "K" , "♠";
 "K" , "♥";
 "K" , "♦";
 "K" , "♣";
 "Q" , "♠";
 "Q" , "♥";
 "Q" , "♦";
 "Q" , "♣";
 "J" , "♠";
 "J" , "♥";
 "J" , "♦";
 "J" , "♣";
 10  , "♠";
 10  , "♥";
 10  , "♦";
 10  , "♣";
 9   , "♠";
 9   , "♥";
 9   , "♦";
 9   , "♣";
 8   , "♠";
 8   , "♥";
 8   , "♦";
 8   , "♣";
 7   , "♠";
 7   , "♥";
 7   , "♦";
 7   , "♣";
 6   , "♠";
 6   , "♥";
 6   , "♦";
 6   , "♣";
 5   , "♠";
 5   , "♥";
 5   , "♦";
 5   , "♣";
 4   , "♠";
 4   , "♥";
 4   , "♦";
 4   , "♣";
 3   , "♠";
 3   , "♥";
 3   , "♦";
 3   , "♣";
 2   , "♠";
 2   , "♥";
 2   , "♦";
 2   , "♣"}

```

#### 5.4 Perform a circular shift on a tuple to create distinct shifts

Shiftλ rotates a tuple by a given number of positions. Supplying a vector of shifts produces all rotations.

Formula:
```Excel
=Shiftλ(12345,SEQUENCE(5))
```

Output:
```text
{"2","3","4","5","1";
 "3","4","5","1","2";
 "4","5","1","2","3";
 "5","1","2","3","4";
 "1","2","3","4","5"}
```

`{1,2,3,4,5}` may also be supplied to yield the same results.

#### 5.5 Split and Fold

Splitλ converts the 1D array to a 2D array of colors. Foldλ provides an aggregate COUNTA of colors where order of colors does not matter.

Colors=
```Text
{"red,green,blue";
"blue,green,red";
"yellow,yellow,yellow";
"green,red,blue";
"yellow,yellow,yellow";
"navy,white,black"}
```
Formula:
```Excel
=Foldλ(Splitλ(colors,","))
```

Output:
```Text
{"black","navy","white",	1;
 "blue","green",  "red",		3;
 "yellow","yellow","yellow",2}
```
#### 5.6 Snap Filtering

Snapλ is designed to be a quick row-level filter alternative to FILTER. The user only needs to specify the include criteria (No booleans required).

Grades=
````
{"Bobby","History","B";
 "Deb","Physics","A";
"Charlie","History","A";
"Charlie","Chemistry","C";
"Bobby","Math","A";
"Willy","Chemistry","B";
"Jose","Math","A";
"Bobby","English","B"}
````
Formula:
```Excel
=Snapλ(Grades,"A")
```
Output:
```Text
{"Deb","Physics","A";
"Charlie","History","A";
"Bobby","Math","A";
"Jose","Math","A"}
```
#### 5.7 Downsample a matrix

Grainλ operates on blocks (specified by `depth` and `width`) and aggregates using an ETA function or custom LAMBDA.

Matrix=
```Text
{345,459,931,924,206,711;
 997,357,425,778,844,839;
 927,675,277,646,843,533;
 489,647,253,657,141,193;
 137,290,266,423,546,651;
 903,455,234,357,257,870}
````
Formula:
```Excel
=Grainλ(Matrix,2,3,MAX)
```
Output:
```Text
{997,924;
 927,843;
 903,870}
```
Formula:
```Excel
=Grainλ(Matrix,2,3,LAMBDA(v,MAX(v)-MIN(v)))
```
Output:
```Text
{652,718;
 674,702;
 766,613}
```

---
## 6. Multi-Step Examples
content

---
## 7. Performance Notes
content

---
## 8. Versioning Notes

### 8.1 Compatability
#### Excel 365 (Windows & Mac) (Full Support - Recommended)

All BigSpill functions run on:
- Current Channel
- Monthly Enterprise Channel
- Semi‑Annual Enterprise Channel

Excel 365 receives ongoing feature updates, including new dynamic array functions and Lambda helpers. Future BigSpill versions may rely on these updates.

#### Excel 2024 - (Full Support as of July 2026)
Excel 2024 includes all Lambda and dynamic array features available at the time of its release, and BigSpill is fully compatible with those features.

Note:
- Excel 2024 does not receive new functions after release.  
- Future BigSpill versions may require features available only in Excel 365.

#### Excel for the Web - (Not Supported) 
Excel for the web includes partial LAMBDA functionality, but it has significant limitations:
- ETA is not supported  
- Implicit aggregation often fails  
- Certain nested or recursive Lambdas do not evaluate correctly  
- AFE modules may not load or resolve reliably  

#### Other Versions and Applications - (Not Supported)
BigSpill is designed for the modern Excel 365. Versions not listed above are not compatible. BigSpill is not intended for use with other spreadsheet applications, and there are no plans to support them.

---
## 9. Appendix

### Onboarding Workbooks

The following workbooks provide demonstrations of all functions in BigSpill. All workbooks include the entire BigSpill module.
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








