<p align="left">
  <img src="../assets/BigSpillLogo.png" width="320">
</p>

# User Manual
**BigSpill** is a library designed to harness the power of modern Excel to make everyday tasks easier and to provide tools for tackling almost any problem.
It brings clarity and reusability to grid‑based work by replacing long, brittle formulas with sustainable, dynamic array solutions.

The library provides reliable, easy‑to‑use functions for everyday tasks and structured building blocks for advanced development. BigSpill is designed to support users at all skill levels.

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
10. Closing Notes

---
## 1. Welcome
**BigSpill** builds on the innovations Excel introduced between 2018 and 2021 (Dynamic Arrays, `LAMBDA`, `LET`) along with the wealth of new functions added in recent years. It exists to bring Excel’s power, where most anything is possible within reason, to the user without the complexity.

The library provides reliable, easy‑to‑use functions for everyday tasks and building blocks for advanced development. BigSpill is designed to support users at all skill levels.

---
## 2. Installation & Setup

### 2.1 Install via BigSpill Template (Recommended)
The [BigSpill Template](../templates/BigSpill_Template.xltx) is the fastest way to install the full library.

Save it to a safe location on your computer and double‑click it to create a new workbook. This new workbook will come preloaded with the entire BigSpill function catalog.

### 2.2 Verifying BigSpill is Installed
1. Double-click `BigSpill_Template.xltx` to open a new workbook.
2. In any cell, enter: `=Shiftλ(123,1)`.
3. If the function spills a 3x1 array `{2; 3; 1}`, then BigSpill is installed. `#NAME?` would indicate the workbook was not created from the BigSpill Template.

### 2.3 Adding BigSpill to an Existing Workbook
If you want to use BigSpill inside an existing `.xlsx`, the easiest method is to copy a sheet from a BigSpill‑enabled workbook. This automatically transfers all BigSpill modules.

Steps:
1. Open a new workbook created from `BigSpill_Template.xltx`.
2. Open the workbook that will receive BigSpill.
3. Right‑click the sheet tab in the BigSpill workbook → Move or Copy...
4. Choose the destination workbook.
5. Check Create a copy → OK.
6. To verify the transfer, enter: `=Shiftλ(123,1)`. If `{2;3;1}` spills, BigSpill is installed.
7. After verification, you may delete the copied sheet.

Note:
- Always copy from the BigSpill workbook into the workbook that will receive BigSpill (Once you've verified installation, the blank sheet can be deleted from your workbook).
- All BigSpill functions include a `λ` suffix, which prevents name conflicts when the above steps are performed.

### 2.4 Requesting a Gist import URL
BigSpill can also be installed using the Advanced Formula Environment (AFE) by importing a private Gist that contains the full module. This installation method is intended for developers.

The Gist URL is not included in this manual. Access may be granted upon request.

---
## 3. Function Reference
This section provides a complete, module‑ordered catalog of all **BigSpill** functions, along with short descriptions and structural groupings. Dependency hierarchies are shown for each category where applicable.

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
A family of geometric extraction functions for 2D arrays. Each function returns a region of the grid defined by a specific shape, preserving the geometry of the extracted area.

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
| **TakeBlockλ** | Extracts a rectangular block using `start_at` and optional `stop_when` |
| **DStackλ** | Stacks multiple 2D arrays along a depth dimension |

#### Reflow

| Function | Description |
| --- | --- |
| **WrapRows2Dλ** | Wraps into row blocks of specified height |
| **WrapCols2Dλ** | Wraps into column blocks of specified width |
| **Staircaseλ** | Constructs diagonal staircase patterns |
| **UnPivotλ** | Converts a matrix into a flattened row‑wise table |
| **Zipλ** | Zips 1–4 arrays by flattening and aligning by length |
| **SliceByDegλ** | Extracts angular slices using degree bounds (0°-360°) |
		
#### Padding

| Function | Description |
| --- | --- |
| **Padλ** | Applies row and/or column padding to any array |
| **InsulateRowsλ** | Adds rows of padding before and/or after an input array |
| **InsulateColsλ** | Adds columns of padding before and/or after an input array |

#### Interpolation

| Function | Description |
| --- | --- |
| **Magnifyλ** | Expands each element into a block of specified `height` × `width` using nearest‑neighbor interpolation |
| **Zoomλ** | Expands each element into a block whose size matches the height × width of `zoom_kernel` |

---

### Grid Algebra
A toolkit for structural editing of 2D arrays. These functions support deletion, selection, filtering, rolling, expansion, and structural mapping.

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
| **BlockMapλ** | Reshapes a matrix row‑wise or column‑wise into blocks of size `depth` × `width` |
| **Convolveλ** | Performs sliding‑window convolution using a 2D `kernel` |
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
| **ByDiagλ** | Applies an ETA function or `LAMBDA` to each diagonal |
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
This category provides functions that overcome common limitations in Excel's native text engine (e.g., `TEXTSPLIT` cannot spill 2D results; `BYROW` cannot deploy `SORT` to align columns). These functions enable structural alignment, extraction, and transformation of text arrays, extending Excel's capabilities for both 1D and 2D inputs.

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

**Note:** `NumbersOnlyλ` is strict - it extracts only characters whose `CODE()` is numeric. Mixed values such as "30.25A" become "3025" because non‑numeric characters (like "." or "A") are removed rather than interpreted.

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
| **KnapSackλ** | Solves the 0-1 Knapsack problem |
| **SubsetGenλ** | Generates all subsets (power set) in binary‑mask order |
| **SubsetSumλ** | Returns all subsets whose elements sum to a `target` |

---

### Diagnostics
Functions whose purpose is to expose, annotate, or explain the internal structure of a grid. Diagnostics functions are used for debugging, visualization, and structural analysis.

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
| #NO-ANCHOR! | `start_at` anchor not found. |
| #NO-DATES! | No dates found in input. |
| #NO-REPEATS! | No repeated values found. |
| #NO-SLICE! | `stop_when` occurs before `start_at`. |
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

Included are several examples illustrating the possibilities available with **BigSpill**.

#### 5.1 Split and Sort

`Splitλ` converts a 1D array of strings into a 2D array. `Alignλ` then sorts each row alphabetically.

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
    sorted)
```

Output:
```text
={"1","1","1","2","5","5","7","8","9","10";
"1","1","1","2","3","6","7","9","9","10";
"3","3","3","4","4","6","6","8","9","10";
"1","2","4","6","6","7","7","8","10","10";
"1","1","2","2","3","5","6","6","8","10"}
```
┈┈┈┈┈┈┈┈┈┈┈
#### 5.2 Easy Filtering

`Snapλ` provides a quick row-level filter alternative to `FILTER`. Only the include‑criteria value is required.

`Grades` =
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
┈┈┈┈┈┈┈┈┈┈┈
#### 5.3 Split and Fold

`Splitλ` converts each row into a 2D array of colors. `Foldλ` aggregates indentical rows with `COUNTA`.

Colors =
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
{"black", "navy",  "white",	1;
 "blue",  "green", "red",	3;
 "yellow","yellow","yellow",2}
```
┈┈┈┈┈┈┈┈┈┈┈
#### 5.4 Cartesian Product

`Pairwiseλ` creates a deck of card cards from rank and suit.

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
┈┈┈┈┈┈┈┈┈┈┈
#### 5.5 Circular Shift

`Shiftλ` rotates a tuple by a specified number of positions. Supplying a vector of shifts produces all rotations.

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

`{1,2,3,4,5}` may also be supplied to produce the same results.

┈┈┈┈┈┈┈┈┈┈┈
#### 5.6 Diamond Sum

`Diamondλ` extracts a diamond-shaped region from the grid. `SUM` then aggregates the values.

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

The unary `+` is used to coerce the result of `Diamondλ`into an array before evaluation by `SUM`.

┈┈┈┈┈┈┈┈┈┈┈
#### 5.7 Downsample a matrix

`Grainλ` operates on blocks (specified by `depth` and `width`) and aggregates using an ETA function or custom `LAMBDA`.

`Matrix` =
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

#### 6.1 Clean-up and Re-Map

`Pinchλ` removes blank rows and columns. `BlockMapλ` then re-tiles the cleaned data using blocks defined by `depth` x `width`.

`Data` =
```Text
{"Jon Doe",        "", "Deb Smith",        "", "Ned Johnson";
 "Senior VP",      "", "HR manager",       "", "Data Analyst";
 "jdoe@example.com","", "dsmith@example.com","", "njohnson@example.com";
 "", "", "", "", "";
 "", "", "", "", "";
 "Peg Bailey",     "", "Randolph Wilkens", "", "Jose Alvarado";
 "Data Analyst",   "", "Legal Advisor",    "", "CEO";
 "pbailey@example.com","", "rwilkens@example.com","", "Jose@example.com"}
```

Formula:
```Excel
=LET(
     cleaned, Pinchλ(Data),
     remapped, BlockMapλ(cleaned, 3, 1),
     remapped)
```

Output:
```Text
{"Jon Doe",        "Senior VP",      "jdoe@example.com";
 "Deb Smith",      "HR manager",     "dsmith@example.com";
 "Ned Johnson",    "Data Analyst",   "njohnson@example.com";
 "Peg Bailey",     "Data Analyst",   "pbailey@example.com";
 "Randolph Wilkens","Legal Advisor", "rwilkens@example.com";
 "Jose Alvarado",  "CEO",            "Jose@example.com"}
```
┈┈┈┈┈┈┈┈┈┈┈

#### 6.2 Inspect, Extract, and Analyze

The goal is to obtain a binned analysis of numerical data with `Histogramλ`.

The dataset contains footnote characters and cannot be altered directly.

`MessyData` =
```Text
{ 2,   "89₁", 45, 31, 20;
 17,    36,  "9₂",  6, 15;
 44,   100,   78, 63, 15;
 75,    41,   80, 80, 91;
 99,   1.2, "52₃", 74, 10;
 13,    66,   84, 61, "62₄";
 15,     4,   85, 12, 81;
"77₅",  47,   66, 90, "99†";
 75,    66,   89, 12, 35;
 13,    18,  "1‡", 28, 53}
```

`Revealλ` confirms the dataset contains non-numeric text.
```Excel
=Revealλ(MessyData)
```

Output:
```Text
{"Number",43;
"Text",7;
"Contains special characters",7;
"Number-stored-as-text",0;
"Error",0;
"Contains spaces",0;
"Formula",0;
"Blank",0;
"Logical",0}
```

To obtain a clean numeric analysis, extract only the numbers using `NumbersOnlyλ`, then compute the histogram with a bin width of 25.

```Excel
=LET(
     numbers, NumbersOnlyλ(MessyData),
     hist, Histogramλ(numbers, 25),
     hist)
```

Output:
```
{"Start", "End",  "Total";
 "<",     "0",    0;
 "0",     "25",   17;
 "25",    "50",   8;
 "50",    "75",   11;
 "75",    "100",  14}
```
┈┈┈┈┈┈┈┈┈┈┈

#### 6.3 Create an Auto-Number and Switch Existing Columns

This example generates auto‑numbered company codes using `Streakλ` and restructures an existing dataset using `Removeλ` and `Embedλ`.

`Products` =
```Text
{"Nexora Solutions","NS","Sprocket",12.99;
"Nexora Solutions","NS","Widget",15.49;
"Nexora Solutions","NS","Gadget",9.99;
"Vibrantix Corp","VC","Cog",14.99;
"Vibrantix Corp","VC","Gear",19.99;
"AetherWave Inc","AW","Bolt",7.99;
"AetherWave Inc","AW","Nut",6.49;
"AetherWave Inc","AW","Screw",5.99;
"AetherWave Inc","AW","Washer",3.99;
"QuantumLeap LLC","QL","Lever",11.49;
"QuantumLeap LLC","QL","Pulley",13.99}
```

`CompanyCode`=
```Text
{"NS";
 "NS";
 "NS";
 "VC";
 "VC";
 "AW";
 "AW";
 "AW";
 "AW";
 "QL";
 "QL"}
```

Formula:
```Excel
=LET(
     autonum, Streakλ(CompanyCode, , TRUE),
     rev_products, Removeλ(Products, , 2),
     new_products, Embedλ(rev_products, autonum, , 2),
     new_products)
```

Output:
```
{"Nexora Solutions","NS·1","Sprocket",12.99;
"Nexora Solutions","NS·2","Widget",15.49;
"Nexora Solutions","NS·3","Gadget",9.99;
"Vibrantix Corp","VC·1","Cog",14.99;
"Vibrantix Corp","VC·2","Gear",19.99;
"AetherWave Inc","AW·1","Bolt",7.99;
"AetherWave Inc","AW·2","Nut",6.49;
"AetherWave Inc","AW·3","Screw",5.99;
"AetherWave Inc","AW·4","Washer",3.99;
"QuantumLeap LLC","QL·1","Lever",11.49;
"QuantumLeap LLC","QL·2","Pulley",13.99}
```
┈┈┈┈┈┈┈┈┈┈┈

#### 6.4 Party of the First Party

In the classic role-playing game Final Fantasy, there are six classes available to form a party of four at the start of the game.

The goal is to list all possible parties without repetition that do not include a Thief

`Classes`=
```Text
{"Fighter";"Black Belt";"Thief";"Red Mage";"White Mage";"Black Mage"}
```

Formula:
```Excel
=LET(
    parties, Combinationsλ(Classes, 4),
    filtered, Excludeλ(parties, "Thief"),
    filtered)
```

Output:
```Text
{"Fighter","Black Belt","Red Mage","White Mage";
"Fighter","Black Belt","Red Mage","Black Mage";
"Fighter","Black Belt","White Mage","Black Mage";
"Fighter","Red Mage","White Mage","Black Mage";
"Black Belt","Red Mage","White Mage","Black Mage"}
```


---
## 7. Performance Notes
All **BigSpill** functions were developed and tested using Excel 365 for PC, 64‑bit. The functions are designed to leverage Excel’s spill and array‑native calculation engine, minimizing formula inputs and optimizing recalculation speed.

#### Early Termination via Dual LET Blocks
Many BigSpill functions use a two‑stage `LET` structure.
The first `LET` block performs lightweight validation and checks for simple cases. This allows a function to terminate early and quickly when an error is detected.  Additionally, some functions perform a simpler, trivial calculation when the situation warrants it. For example, `Convolveλ` will route to a simpler `MAP` when the supplied `kernel` is a scalar.  Functions will only proceed to the second `LET` block when necessary.

#### Deferred Evaluation (“Thunks”)
Several functions use deferred evaluation to avoid Excel’s eager calculation of intermediate indices.
By wrapping index expressions in `LAMBDA`, BigSpill prevents Excel from computing values until they are actually needed.
This technique is applied only where it measurably improves performance. Functions that showed no benefit from deferred evaluation use direct evaluation instead.

For an example where deferred evaluation is advantageous, please see `Spiralλ`: https://gist.github.com/Patrick2788/f89ce80c7410bd30eef8adb949b088b0.

#### Reuse of Lower‑Level Operators
As shown in the dependency hierarchies in **Section 3: Function Reference**, most BigSpill functions are composed from smaller, efficient primitives such as `Resizeλ`, `Streakλ` and `Echoλ`, for example. This modular design reduces duplication, improves maintainability, and ensures that complex operators inherit the performance characteristics of the optimized lower‑level components.


---
## 8. Versioning Notes

### 8.1 Compatability
#### Excel 365 (Windows & Mac) - Full Support (Recommended)

All **BigSpill** functions run on:
- Current Channel
- Monthly Enterprise Channel
- Semi‑Annual Enterprise Channel

Excel 365 receives ongoing feature updates, including new dynamic array functions and Lambda helpers. Future BigSpill versions may rely on these updates.

#### Excel 2024 - Full Support as of July 2026
Excel 2024 includes all `LAMBDA` and dynamic array features available at the time of its release, and BigSpill is fully compatible with those features.

Note:
- Excel 2024 does not receive new functions after release.
- Future BigSpill versions may require features available only in Excel 365.

#### Excel for the Web - Not Supported 
Excel for the web includes partial LAMBDA functionality, but it has significant limitations:
- ETA is not supported  
- Implicit aggregation often fails  
- Certain nested or recursive Lambdas do not evaluate correctly  
- AFE modules may not load or resolve reliably

#### Excel Insider Channels (Beta / Preview) - Use at Your Own Risk
Excel Insider builds receive frequent updates that introduce or remove experimental features. These changes may affect the calculation engine, spill behavior, or Lambda evaluation. Full compatibility with BigSpill cannot be guaranteed.

#### Other Versions and Applications - Not Supported
BigSpill is designed for the modern Excel 365. Versions not listed above are not compatible. BigSpill is not intended for use with other spreadsheet applications, and there are no plans to support them.

---
## 9. Appendix

### Onboarding Workbooks

The following workbooks provide demonstrations of all functions in **BigSpill**. All workbooks include the entire BigSpill module.
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

---
## 10. Closing Notes

**BigSpill** is the result of nearly two decades of study, including over one thousand hours of Lambda research and development work. The goal has always been to make complex calculations easy, elegant, and sustainable. BigSpill is a living library that will update as Excel continues to evolve with new dynamic array functions, Lambda helpers, and ETA innovations.

Thank you for taking the time to explore BigSpill and enjoy creating with this library!





