# VBA Array & Combinatorics Toolkit

**Author:** Martín Roldán ([@M-Roldan](https://github.com/M-Roldan))

A modular Excel VBA library for array manipulation, text processing, math operations, and combinatorics. All modules follow consistent design conventions and are built to work together as a cohesive foundation for complex spreadsheet automation.

---

## Modules

### `ERRORS`
Centralized registry of custom error codes shared across the library. Each module is assigned a reserved numeric range so that any error can be traced back to its origin by code alone.

| Range | Module |
|---|---|
| 1000–1999 | ARRAYS |
| 2000–2999 | TEXT |
| 3000–3999 | MATH |
| 4000–4099 | COMBINATORICS |
| 100000–100999 | Portable modules |

---

### `ARRAYS`
Structural foundation of the library. Provides utilities for inspecting, normalizing, transforming, and cleaning VBA arrays. Used internally by `MATH` and `COMBINATORICS`.

| Function | Description |
|---|---|
| `ARR_NDIM` | Returns the number of dimensions of an array (0 if not an array) |
| `ARR_TO_2D` | Normalizes any input (scalar, 1D, 2D, Range) into a 2D array |
| `ARR_SHAPE` | Returns the size of each dimension as a 1D array |
| `ARR_2D_GET_COL` | Extracts a single column from a 2D array as a 1D array |
| `ARR_2D_SPLIT_COLS` | Splits a 2D array into a 1-based array of cleaned column arrays |
| `ARR_1D_REMOVE_EMPTY` | Removes empty strings, Empty variants, and error values from a 1D array |
| `ARR_TO_1D_LONG` | Converts any scalar or vector input into a 1-based 1D Long array |
| `ARR_TO_1D_STRING` | Converts any scalar or vector input into a 1-based 1D String array |

**Depends on:** `ERRORS`

---

### `TEXT`
String utilities for extracting and transforming character subsets. Also serves as the numeric-to-digits bridge used by the Radix Engine in `MATH`.

| Function | Description |
|---|---|
| `TXT_STRIP_DIGITS` | Removes all digit characters (0–9) from a string |
| `TXT_EXTRACT_DIGITS` | Extracts all digit characters concatenated |
| `TXT_EXTRACT_ALPHA` | Extracts all ASCII alphabetic characters (A–Z, a–z) |
| `TXT_EXTRACT_ALPHANUM` | Extracts all ASCII alphanumeric characters |
| `TXT_EXTRACT_FIRSTNUM` | Extracts the first contiguous sequence of digits |
| `TXT_NUM_TO_DIGITS` | Converts a non-negative integer into a 1-based digit array (MSD → LSD) |

**Depends on:** `ERRORS`

---

### `MATH`
Two independent mathematical engines: a GCD engine and a Radix engine for base conversion.

> **Note:** `MATH_GCD` requires **64-bit Office (Win64)**. On 32-bit installations it returns `#N/A`.

| Function | Description |
|---|---|
| `MATH_GCD` | Computes the GCD of any combination of scalars, arrays, and ranges |
| `MATH_RADIX_CONVERT` | Converts a number between any two bases (pure or mixed) |
| `MATH_RADIX_TO_DECIMAL` | Converts a digit array in any base to its decimal value |
| `MATH_RADIX_FROM_DECIMAL` | Converts a decimal value to any base representation |

**Pure base:** a single integer ≥ 2 applied uniformly to all digit positions.  
**Mixed base:** a 1D array where each element defines the base of one position.  
Digit arrays always use MSD → LSD order.

**Depends on:** `ARRAYS`, `TEXT`, `ERRORS`

---

### `COMBINATORICS`
Generates Cartesian products and variable-length radix combinations from multiple input ranges or arrays.

| Function | Description |
|---|---|
| `COMB_CARTESIAN` | Returns the full Cartesian product as a 2D array (rows = combinations, columns = dimensions) |
| `COMB_CARTESIAN_UNRANK` | Returns a single combination at a 1-based index without computing the full product |
| `COMB_UNRANK_VARLEN` | Returns a single combination at a 1-based index from an unbounded variable-length radix space |

The default row limit is controlled by `MAX_COMBINATIONS = 200000`. Increase with caution — large values may cause Excel to become unresponsive.

**Depends on:** `ARRAYS`, `MATH`, `ERRORS`

---

### Portable modules

Self-contained drop-in modules with no external dependencies. Each one bundles all required utilities internally using a module-specific prefix. Designed to be copied into any VBA project as a single file.

| Module | Equivalent function | Description |
|---|---|---|
| `ZP001_COMBINE_RANGES_PORTABLE` | `P001_COMBINE_RANGES` | Full Cartesian product (same as `COMB_CARTESIAN`) |
| `ZP002_UNRANK_VARLEN_PORTABLE` | `P002_UNRANK_VARLEN` | Variable-length radix unranking (same as `COMB_UNRANK_VARLEN`) |

---

## Design conventions

- **1-based indexing.** All output arrays use `LBound = 1`, matching Excel's native range convention.
- **Canonical 2D shape.** `ARR_TO_2D` is the standard normalization entry point. Scalars become `(1 × 1)`, 1D arrays become `(N × 1)` columns.
- **Range transparency.** Any function that accepts `Variant` input detects `Range` objects and extracts `.Value` before processing. Downstream logic never sees `Range` objects.
- **Error signaling.** Internal violations raise `Err.Raise ERR_*`. Public worksheet UDFs catch all errors and return `CVErr(...)` so they degrade gracefully in cells.
- **Empty-value semantics.** Cleaning routines discard empty strings (`""`), `Empty` variants, and error values. A fully empty result returns `Empty` rather than a zero-length array.

---

## Installation

1. Open the Excel workbook where you want to use the library.
2. Open the VBA Editor: `Alt + F11`.
3. In the menu, go to **File → Import File...**.
4. Import the modules in this order (to respect dependencies):
   1. `ERRORS.bas`
   2. `ARRAYS.bas`
   3. `TEXT.bas`
   4. `MATH.bas`
   5. `COMBINATORICS.bas`
5. If you only need a single self-contained function, import the corresponding portable module instead (`ZP001_*.bas` or `ZP002_*.bas`).
6. Make sure **Tools → References** has no missing references flagged.
7. Save the workbook as `.xlsm` (macro-enabled) or `.xlsb`.

---

## Quick examples

### Cartesian product of two lists

```vba
' Returns a 6×2 array: all combinations of {A,B,C} × {1,2}
Dim result As Variant
result = COMB_CARTESIAN(Array("A", "B", "C"), Array(1, 2))
```

### Base conversion

```vba
' Convert 255 from decimal (base 10) to binary (base 2)
Dim bits As Variant
bits = MATH_RADIX_CONVERT(Array(2, 5, 5), 10, 2)
' Returns: Array(1, 1, 1, 1, 1, 1, 1, 1)  → 11111111

' Convert from mixed base (hours:minutes:seconds)
Dim hms As Variant
hms = MATH_RADIX_CONVERT(Array(1, 30, 0), Array(24, 60, 60), 10)
' Returns: 5400 (total seconds)
```

### GCD of a range

```vba
' In a worksheet cell:
=MATH_GCD(A1:A10)
```

### Variable-length unranking

```vba
' Get the 10th combination of any length over {a, b, c}
Dim combo As Variant
combo = COMB_UNRANK_VARLEN(10, Array("a", "b", "c"))
```

### Normalize an array to 2D

```vba
Dim arr As Variant
arr = ARR_TO_2D(Array(10, 20, 30))
' Returns: (3×1) column array
```

---

## Requirements

- Microsoft Excel with macros enabled
- VBA editor access (not restricted by IT policy)
- `Option Explicit` is used throughout — all variables must be declared
- `MATH_GCD` requires **64-bit Office** (Win64 compilation flag)

---

## License

MIT License — Copyright (c) 2025 M. Roldán — see [LICENSE](LICENSE) for details.
