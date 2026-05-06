Attribute VB_Name = "ERRORS"
Option Explicit

' ============================================================
' MODULE: ERRORS
' CATEGORY: Library Infrastructure / Error Codes
'
' PURPOSE:
' Centralized registry of custom error codes raised across
' the library. Each module is assigned a reserved numeric
' range to prevent collisions and to make error origin
' traceable from the code alone.
'
' ------------------------------------------------------------
' RANGE ALLOCATION
' ------------------------------------------------------------
'
'   1000 - 1999   ARRAYS
'   2000 - 2999   TEXT
'   3000 - 3999   MATH
'   4000 - 4099   COMBINATORICS
'   5000 - 5999   reserved for future modules
'   NA            loose modules
'
' All codes are offsets to be added to vbObjectError.
' Use the ERR_* enum values directly; do not hardcode numbers.
'
' ============================================================

Public Enum LibError

    ' --- ARRAYS (1000-1999) ---
    ERR_ARR_NOT_VECTOR = vbObjectError + 1001
    ERR_ARR_INVALID_DIMENSION = vbObjectError + 1002
    ERR_ARR_COLINDEX_OUT_OF_RANGE = vbObjectError + 1003
    ERR_ARR_INVALID_VALUE = vbObjectError + 1004
    
    ' --- TEXT (2000-2999) ---
    ERR_TXT_INVALID_NUMERIC = vbObjectError + 2001

    ' --- MATH (3000-3999) ---
    ERR_RADIX_BASE_TOO_SMALL = vbObjectError + 3001
    ERR_RADIX_DIGIT_OUT_OF_RANGE = vbObjectError + 3002
    ERR_RADIX_SIZE_MISMATCH = vbObjectError + 3003
    ERR_RADIX_MIXED_BASE_INVALID = vbObjectError + 3004
    ERR_RADIX_MIXED_DIGIT_INVALID = vbObjectError + 3005
    ERR_RADIX_NEGATIVE_VALUE = vbObjectError + 3006
    ERR_RADIX_NON_INTEGER = vbObjectError + 3007
    ERR_RADIX_DIGITS_NOT_ARRAY = vbObjectError + 3008
    ERR_RADIX_BASE_NOT_1D = vbObjectError + 3009
    ERR_RADIX_OVERFLOW = vbObjectError + 3010

    ' --- COMBINATORICS (4000-4999) ---
    ERR_COMB_TOO_MANY = vbObjectError + 4001
    ERR_COMB_INDEX_OUT_OF_RANGE = vbObjectError + 4002
    ERR_COMB_INVALID_DIMENSIONS = vbObjectError + 4003
    

End Enum
