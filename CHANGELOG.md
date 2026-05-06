# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.0] — 2025

### Added

#### ERRORS
- Centralized `LibError` enum with reserved error code ranges for all modules.

#### ARRAYS
- `ARR_NDIM` — returns the number of dimensions of any array or Range input.
- `ARR_TO_2D` — canonical normalization of scalars, 1D arrays, and Ranges into a 2D array.
- `ARR_SHAPE` — returns the size of each dimension as a 1-based 1D array.
- `ARR_2D_GET_COL` — extracts a single column from a 2D array as a 1D array.
- `ARR_2D_SPLIT_COLS` — splits a 2D array into a 1-based array of cleaned column arrays.
- `ARR_1D_REMOVE_EMPTY` — removes empty strings, Empty variants, and error values from a 1D array.
- `ARR_TO_1D_LONG` — converts any scalar or vector input into a 1-based 1D Long array.
- `ARR_TO_1D_STRING` — converts any scalar or vector input into a 1-based 1D String array.

#### TEXT
- `TXT_STRIP_DIGITS` — removes all digit characters from a string.
- `TXT_EXTRACT_DIGITS` — extracts all digit characters concatenated.
- `TXT_EXTRACT_ALPHA` — extracts all ASCII alphabetic characters.
- `TXT_EXTRACT_ALPHANUM` — extracts all ASCII alphanumeric characters.
- `TXT_EXTRACT_FIRSTNUM` — extracts the first contiguous sequence of digits.
- `TXT_NUM_TO_DIGITS` — converts a non-negative integer into a 1-based digit array (MSD → LSD).

#### MATH
- `MATH_GCD` — GCD of any combination of scalars, arrays, and ranges (Win64 only).
- `MATH_RADIX_CONVERT` — converts a number between any two bases, pure or mixed.
- `MATH_RADIX_TO_DECIMAL` — converts a digit array in any base to its decimal value.
- `MATH_RADIX_FROM_DECIMAL` — converts a decimal value to any base representation.

#### COMBINATORICS
- `COMB_CARTESIAN` — generates the full Cartesian product of multiple inputs as a 2D array.
- `COMB_CARTESIAN_UNRANK` — returns a single combination at a 1-based index without computing the full product.
- `COMB_UNRANK_VARLEN` — returns a single combination at a 1-based index from an unbounded variable-length radix space.
- `MAX_COMBINATIONS` public constant (default: 200 000) to cap output size.

#### Portable modules
- `ZP001_COMBINE_RANGES_PORTABLE` — self-contained Cartesian product (`P001_COMBINE_RANGES`), no dependencies.
- `ZP002_UNRANK_VARLEN_PORTABLE` — self-contained variable-length unranking (`P002_UNRANK_VARLEN`), no dependencies.
