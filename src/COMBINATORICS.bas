Attribute VB_Name = "COMBINATORICS"
Option Explicit

' User-adjustable performance limit for COMB_CARTESIAN and
' P001_COMBINE_RANGES. Increase with caution -- large values
' may cause Excel to become unresponsive.
Public Const MAX_COMBINATIONS As Long = 200000

' Internal arithmetic guard. Long overflows above 2^31 - 1.
' This limit is not adjustable -- it is a property of the
' Long data type used for total combination count.
Private Const MAX_LONG As Long = 2147483647

' Double loses integer precision above 2^53.
' Beyond this point, index validation is unreliable.
Private Const MAX_SAFE_DOUBLE As Double = 9.00719925474099E+15 ' 2^53


' ============================================================
' MODULE: COMBINATORICS
' CATEGORY: Math / Combinatorics
'
' PURPOSE:
' This module provides functions for generating Cartesian
' products and variable-length radix combinations from
' multiple input ranges or arrays. Each column of each input
' is treated as an independent dimension, and the output
' enumerates all possible combinations across those dimensions.
'
' ------------------------------------------------------------
' PUBLIC API
' ------------------------------------------------------------
'
'   COMB_CARTESIAN
'     Generates the full Cartesian product of all inputs and
'     returns it as a 2D array (rows = combinations,
'     columns = dimensions).
'
'   COMB_CARTESIAN_UNRANK
'     Returns a single combination at a specified 1-based
'     index without computing the full product. Useful when
'     only one row is needed at a time.
'
'   COMB_UNRANK_VARLEN
'     Returns a single combination at a specified 1-based
'     index from a variable-length radix space. Higher indices
'     map to longer combinations. Useful for enumerating all
'     possible strings or sequences of any length over a
'     fixed set of values.
'
' ------------------------------------------------------------
' INTERNAL FUNCTIONS
' ------------------------------------------------------------
'
'   Cartesian_Build_Full
'     Core engine. Builds the full Cartesian product using
'     positional weight (mixed radix) decomposition. Does not
'     call the Radix Engine -- uses its own direct arithmetic
'     for performance across up to MAX_COMBINATIONS iterations.
'
'   Cartesian_Unrank
'     Index engine for fixed-length Cartesian products.
'     Returns one combination by delegating index decomposition
'     to MATH_RADIX_CONVERT (MATH module). Called once per
'     COMB_CARTESIAN_UNRANK invocation.
'
'   Cartesian_Build_Dimensions
'     Normalization helper. Converts raw ParamArray inputs
'     into a clean 1-based array of 1D dimension arrays,
'     ready for consumption by the Cartesian engines above.
'
'   VarlenRadix_Unrank
'     Index engine for variable-length radix spaces. Returns
'     one combination by locating the block containing the
'     target index and delegating decomposition to
'     MATH_RADIX_CONVERT. Called once per COMB_UNRANK_VARLEN
'     invocation.
'
' ------------------------------------------------------------
' DESIGN CONVENTIONS
' ------------------------------------------------------------
'
' 1. DIMENSIONS
'    Each input is split into columns via ARR_2D_SPLIT_COLS.
'    Each non-empty column becomes one independent dimension.
'    The number of output columns equals the total number of
'    non-empty columns across all inputs.
'
' 2. COMBINATION COUNT
'    Total combinations = product of all dimension sizes.
'    Raises ERR_COMB_TOO_MANY if this exceeds MAX_COMBINATIONS.
'    This limit does not apply to COMB_UNRANK_VARLEN, which
'    operates on an unbounded space.
'
' 3. INDEXING
'    Combinations are numbered 1-based (1 = first combination).
'    Internally, the index is converted to zero-based before
'    mixed radix decomposition.
'
' 4. ERROR SIGNALING
'    Internal violations raise Err.Raise ERR_*.
'    Public UDFs catch all errors and return CVErr(...):
'      ERR_COMB_TOO_MANY            --> xlErrNum
'      ERR_COMB_INDEX_OUT_OF_RANGE  --> xlErrNum
'      ERR_COMB_INVALID_DIMENSIONS  --> xlErrNum
'      ERR_ARR_NOT_VECTOR           --> xlErrNum
'      ERR_ARR_INVALID_VALUE        --> xlErrNum
'      anything else                --> xlErrValue
'
' ------------------------------------------------------------
' DEPENDENCIES
' ------------------------------------------------------------
'
'   ARRAYS module:
'     ARR_TO_2D, ARR_2D_SPLIT_COLS (used by Cartesian engine)
'     ARR_TO_1D_STRING (used by COMB_UNRANK_VARLEN)
'
'   MATH module:
'     MATH_RADIX_CONVERT (used by Cartesian_Unrank and
'     VarlenRadix_Unrank)
'
'   ERRORS module:
'     ERR_COMB_TOO_MANY, ERR_COMB_INDEX_OUT_OF_RANGE,
'     ERR_COMB_INVALID_DIMENSIONS, ERR_ARR_NOT_VECTOR,
'     ERR_ARR_INVALID_VALUE
'
' ------------------------------------------------------------
' RELATED MODULES
' ------------------------------------------------------------
'
'   TEXT module:
'     Not a direct dependency. TEXT is consumed by MATH,
'     which COMBINATORICS depends on transitively.
'
' ============================================================















' ============================================================
' FUNCTION: Cartesian_Build_Full
' CATEGORY: Combinatorics / Cartesian Product (Internal)
'
' DESCRIPTION:
' The Cartesian_Build_Full function generates the full
' Cartesian product of a set of dimensions using positional
' weight (mixed radix) decomposition. It is the core engine
' called by COMB_CARTESIAN.
'
' Each row of the output corresponds to one unique combination.
' The row index (0-based internally) is decomposed into per-
' dimension indices using the same positional weight formula
' used in mixed radix arithmetic:
'
'   indexValue = (row \ divisor[i]) Mod size[i]
'
' This avoids nested loops entirely -- each dimension is filled
' independently in a single pass over all rows.
'
' PERFORMANCE NOTE:
' This function does NOT delegate to MATH_RADIX_CONVERT. The
' decomposition is implemented directly to avoid normalization
' overhead across up to MAX_COMBINATIONS iterations. For a
' single-row lookup, see Cartesian_Unrank.
'
' PRECONDITIONS (caller must guarantee):
'   dimensions must be a 1-based 1D array of 1D arrays.
'   Each inner array must be non-empty.
'   Use Cartesian_Build_Dimensions to produce a conformant
'   dimensions array from raw user input.
'
' STRATEGY:
' 1. Compute sizes and validate total
'    For each dimension, compute its size and accumulate the
'    running product. Raises ERR_COMB_TOO_MANY if the total
'    would exceed MAX_COMBINATIONS.
'
' 2. Compute positional weights (divisors)
'    divisors(i) = product of sizes of all dimensions to the
'    right of i. The rightmost dimension has weight 1.
'    Built right to left: divisors(i) = divisors(i+1) * size(i+1).
'
' 3. Allocate result
'    A 2D array (total x numDims) is allocated once.
'
' 4. Fill using positional weight formula
'    For each row index (0 to total-1) and each dimension i:
'      indexValue = (row \ divisors(i)) Mod sizes(i)
'    This gives the 0-based position within dimension i for
'    that row. The actual value is looked up as:
'      dimensions(i)(LBound(dimensions(i)) + indexValue)
'    The LBound offset makes the lookup safe for dimensions
'    with any lower bound.
'
' PARAMETERS:
'   dimensions  Variant   1-based 1D array of 1D arrays.
'                         Each element is one dimension.
'
' RETURN VALUE:
' A 1-based 2D Variant array (total x numDims) where each
' row is one combination and each column is one dimension.
' Returns Empty if any dimension has size 0.
'
' EXAMPLE:
' ' D1 = ("A","B"), D2 = (1,2,3)
' ' Returns 6 rows:
' '   A 1
' '   A 2
' '   A 3
' '   B 1
' '   B 2
' '   B 3
'
' ============================================================
Private Function Cartesian_Build_Full( _
    ByVal dimensions As Variant _
) As Variant

    Dim sizes() As Long
    Dim divisors() As Double
    Dim i As Long

    Dim total As Double: total = 1
    Dim numDims As Long

    ' --------------------------------------------------------
    ' STEP 0: Read dimension bounds
    ' dimensions may use any lower bound, so lb and ub are
    ' captured once and used throughout to avoid repeated
    ' LBound/UBound calls.
    ' --------------------------------------------------------
    Dim lb As Long, ub As Long
    lb = LBound(dimensions)
    ub = UBound(dimensions)

    numDims = ub - lb + 1

    ReDim sizes(lb To ub)
    ReDim divisors(lb To ub)

    ' --------------------------------------------------------
    ' STEP 1: Compute sizes and validate total
    ' Accumulate the product of all dimension sizes into total.
    ' The overflow guard uses floating-point division to avoid
    ' Long overflow before the comparison.
    ' --------------------------------------------------------
    For i = lb To ub

        sizes(i) = UBound(dimensions(i)) - LBound(dimensions(i)) + 1

        ' An empty dimension makes the Cartesian product empty
        If sizes(i) = 0 Then
            Cartesian_Build_Full = Empty
            Exit Function
        End If

        ' Guard 1: arithmetic -- prevent Long overflow before multiplying.
        ' Uses integer division to avoid the overflow in the comparison itself.
        If total > CDbl(MAX_LONG) / CDbl(sizes(i)) Then
            Err.Raise ERR_COMB_TOO_MANY, , _
                "Combination space exceeds Long capacity"
        End If

        total = total * sizes(i)
        
        ' Guard 2: performance -- enforce the user-configurable limit.
        ' Checked after multiplying because total is now the updated product.
        If total > MAX_COMBINATIONS Then
            Err.Raise ERR_COMB_TOO_MANY, , _
                "Too many combinations (" & total & "). Limit: " & MAX_COMBINATIONS
        End If

    Next i

    ' --------------------------------------------------------
    ' STEP 2: Compute positional weights (right to left)
    ' divisors(i) holds the product of all sizes to the right
    ' of position i. This is the weight of dimension i in the
    ' mixed radix decomposition:
    '   divisors(ub)   = 1
    '   divisors(ub-1) = size(ub)
    '   divisors(ub-2) = size(ub) * size(ub-1)
    '   ...
    ' --------------------------------------------------------
    divisors(ub) = 1

    For i = ub - 1 To lb Step -1
        divisors(i) = divisors(i + 1) * sizes(i + 1)
    Next i

    ' --------------------------------------------------------
    ' STEP 3: Allocate result array (1-based, rows x cols)
    ' --------------------------------------------------------
    Dim result() As Variant
    ReDim result(1 To total, 1 To numDims)

    ' --------------------------------------------------------
    ' STEP 4: Fill result using positional weight formula
    ' For each 0-based row index:
    '   indexValue = (row \ divisors(i)) Mod sizes(i)
    ' gives the 0-based position within dimension i.
    ' The LBound offset maps it to the actual array index.
    ' --------------------------------------------------------
    Dim row As Long
    Dim indexValue As Double
    Dim colIndex As Long

    For row = 0 To total - 1

        colIndex = 1

        For i = lb To ub

            ' Decompose row index for this dimension
            indexValue = (row \ divisors(i)) Mod sizes(i)

            ' Map position index to actual dimension value
            result(row + 1, colIndex) = _
                dimensions(i)(LBound(dimensions(i)) + indexValue)

            colIndex = colIndex + 1

        Next i

    Next row

    Cartesian_Build_Full = result

End Function



' ============================================================
' FUNCTION: COMB_CARTESIAN
' CATEGORY: Combinatorics / Cartesian Product
'
' DESCRIPTION:
' The COMB_CARTESIAN function generates the full Cartesian
' product of a variable number of inputs and returns it as a
' 2D array where each row is one unique combination and each
' column corresponds to one dimension.
'
' Each input may contain multiple columns. Each non-empty
' column is treated as an independent dimension. Empty cells
' within a column are silently removed before processing.
'
' STRATEGY:
' 1. Normalize inputs into dimensions
'    Delegates to Cartesian_Build_Dimensions, which converts
'    each input to a 2D array, splits it into columns, removes
'    empty values, and collects the non-empty columns into a
'    1-based array of dimensions. Raises
'    ERR_COMB_INVALID_DIMENSIONS if no valid dimensions are
'    found.
'
' 2. Delegate to the Cartesian engine
'    Passes the normalized dimensions array to
'    Cartesian_Build_Full, which produces the full
'    result using positional weight decomposition.
'
' PARAMETERS:
'   inputs    ParamArray   Any combination of:
'                            - Range objects (single or multi-column)
'                            - 1D or 2D arrays
'                            - Scalar values
'                          Each column becomes one dimension.
'                          Empty cells are ignored.
'
' RETURN VALUE:
' A 1-based 2D Variant array (combinations x dimensions).
' Returns a CVErr on failure:
'   xlErrNum   if inputs produce no valid dimensions or if the
'              total combinations exceed MAX_COMBINATIONS
'   xlErrValue for any other unexpected error
'
' EXAMPLE:
' ' Worksheet usage:
' =COMB_CARTESIAN(A1:A3, B1:B2)
' ' Returns all 6 combinations of column A x column B
'
' ' VBA usage:
' Dim result As Variant
' result = COMB_CARTESIAN(Array("A","B"), Array(1,2,3))
' ' Returns:
' '   A 1
' '   A 2
' '   A 3
' '   B 1
' '   B 2
' '   B 3
'
' ============================================================
Public Function COMB_CARTESIAN(ParamArray inputs() As Variant) As Variant

    ' Copy ParamArray to a plain Variant before passing.
    ' VBA does not allow passing a ParamArray directly as
    ' an argument to another function.
    Dim inputsCopy As Variant
    inputsCopy = inputs

    Dim dimensions As Variant

    On Error GoTo HandleError

    ' --------------------------------------------------------
    ' STEP 1: Normalize inputs --> dimensions
    ' Raises ERR_COMB_INVALID_DIMENSIONS if no valid
    ' dimensions are found across all inputs.
    ' --------------------------------------------------------
    dimensions = Cartesian_Build_Dimensions(inputsCopy)

    ' --------------------------------------------------------
    ' STEP 2: Delegate to math-based Cartesian engine
    ' --------------------------------------------------------
    COMB_CARTESIAN = Cartesian_Build_Full(dimensions)

    Exit Function

HandleError:

    Select Case Err.Number
        Case ERR_COMB_TOO_MANY, ERR_COMB_INVALID_DIMENSIONS
            COMB_CARTESIAN = CVErr(xlErrNum)
        Case Else
            COMB_CARTESIAN = CVErr(xlErrValue)
    End Select

End Function







' ============================================================
' FUNCTION: Cartesian_Unrank
' CATEGORY: Combinatorics / Cartesian Product (Internal)
'
' DESCRIPTION:
' The Cartesian_Unrank function returns the
' single combination at a given 1-based index from the
' Cartesian product of a set of dimensions, without computing
' the full product.
'
' The index is interpreted as a position in the enumeration
' produced by COMB_CARTESIAN: index 1 = first row,
' index 2 = second row, and so on.
'
' Internally, the 1-based index is converted to zero-based
' and then decomposed into per-dimension positions using
' MATH_RADIX_CONVERT, treating the dimension sizes as a
' mixed base system. Each position is then mapped to the
' actual value in its corresponding dimension array.
'
' PERFORMANCE NOTE:
' Unlike Cartesian_Build_Full, this function calls
' MATH_RADIX_CONVERT once per invocation. This is acceptable
' because the function is called once per COMB_CARTESIAN_UNRANK
' call, not in a loop over all combinations.
'
' PRECONDITIONS (caller must guarantee):
'   dimensions must be a 1-based 1D array of 1D arrays.
'   Each inner array must be non-empty.
'   targetIndex must be >= 1.
'
' STRATEGY:
' 1. Compute sizes of each dimension
'    For each dimension, compute its size as
'    UBound - LBound + 1. Exits silently if any size is 0.
'
' 2. Compute total combinations (overflow-safe)
'    Accumulates the product of all sizes using Double to
'    avoid Long overflow. Raises ERR_COMB_TOO_MANY if the
'    total would exceed MAX_COMBINATIONS.
'
' 3. Validate index bounds
'    Raises ERR_COMB_INDEX_OUT_OF_RANGE if targetIndex is
'    outside [1, total].
'
' 4. Convert index to zero-based
'    Subtracts 1 from targetIndex. MATH_RADIX_CONVERT expects
'    a zero-based value since digit positions start at 0.
'
' 5. Decompose index into per-dimension positions
'    Calls MATH_RADIX_CONVERT(zeroIndex, 10, sizes) to convert
'    the scalar index from base 10 into a mixed base
'    representation where each base is the corresponding
'    dimension size. The result is a 1-based array of
'    0-based positions, one per dimension.
'
' 6. Map positions to actual values
'    For each dimension i, retrieves the value at:
'      dimensions(i)(LBound(dimensions(i)) + digits(colIndex))
'    The LBound offset makes the lookup safe for dimensions
'    with any lower bound.
'
' PARAMETERS:
'   dimensions    Variant   1-based 1D array of 1D arrays.
'                           Each element is one dimension.
'
'   targetIndex   Long      1-based index of the combination
'                           to retrieve. Must be in [1, total].
'
' RETURN VALUE:
' A 1-based (1 x numDims) 2D Variant array representing the
' combination at the specified index. Returns Empty if any
' dimension has size 0.
'
' EXAMPLE:
' ' D1 = ("A","B"), D2 = (1,2,3) --> 6 combinations
' ' Index 4 = ("B", 1)
' ' Index 5 = ("B", 2)
' ' Index 6 = ("B", 3)
'
' ============================================================
Private Function Cartesian_Unrank( _
    ByVal dimensions As Variant, _
    ByVal targetIndex As Long _
) As Variant

    Dim lb As Long, ub As Long
    lb = LBound(dimensions)
    ub = UBound(dimensions)

    Dim numDims As Long
    numDims = ub - lb + 1

    Dim sizes() As Long
    ReDim sizes(lb To ub)

    Dim i As Long

    ' --------------------------------------------------------
    ' STEP 1: Compute sizes of each dimension
    ' An empty dimension makes the result undefined -- exit
    ' silently and return Empty implicitly.
    ' --------------------------------------------------------
    For i = lb To ub
        sizes(i) = UBound(dimensions(i)) - LBound(dimensions(i)) + 1
        If sizes(i) = 0 Then Exit Function
    Next i

    ' --------------------------------------------------------
    ' STEP 2: Compute total combinations (overflow-safe)
    ' Uses Double and floating-point division in the guard to
    ' avoid Long overflow before the multiplication.
    ' --------------------------------------------------------
    Dim total As Double
    total = 1

    For i = lb To ub
        If total > MAX_SAFE_DOUBLE / CDbl(sizes(i)) Then
            Err.Raise ERR_COMB_INDEX_OUT_OF_RANGE, , _
                "Combination space exceeds Double precision limit"
        End If
        total = total * sizes(i)
    Next i

    ' --------------------------------------------------------
    ' STEP 3: Validate index bounds
    ' --------------------------------------------------------
    If targetIndex < 1 Or targetIndex > total Then
        Err.Raise ERR_COMB_INDEX_OUT_OF_RANGE, , "Index out of range"
    End If

    ' --------------------------------------------------------
    ' STEP 4: Convert index to zero-based
    ' MATH_RADIX_CONVERT produces digits in [0, size-1], so
    ' the index must be zero-based before decomposition.
    ' --------------------------------------------------------
    Dim zeroIndex As Long
    zeroIndex = targetIndex - 1

    ' --------------------------------------------------------
    ' STEP 5: Decompose index into per-dimension positions
    ' Treats sizes as a mixed base system: each dimension size
    ' is the base for that position. The result is a 1-based
    ' array of 0-based position indices, one per dimension.
    ' --------------------------------------------------------
    Dim digits As Variant
    digits = MATH_RADIX_CONVERT(zeroIndex, 10, sizes)

    ' --------------------------------------------------------
    ' STEP 6: Map position indices to actual dimension values
    ' digits(colIndex) is the 0-based position within
    ' dimension i. The LBound offset makes the lookup safe
    ' for dimensions with any lower bound.
    ' --------------------------------------------------------
    Dim result() As Variant
    ReDim result(1 To 1, 1 To numDims)

    Dim colIndex As Long
    colIndex = 1

    Dim dimArray As Variant
    Dim dimLB As Long
    For i = lb To ub

        dimArray = dimensions(i)

        dimLB = LBound(dimArray)

        result(1, colIndex) = dimArray(dimLB + digits(colIndex))

        colIndex = colIndex + 1

    Next i

    Cartesian_Unrank = result

End Function




' ============================================================
' FUNCTION: COMB_CARTESIAN_UNRANK
' CATEGORY: Combinatorics / Cartesian Product
'
' DESCRIPTION:
' The COMB_CARTESIAN_UNRANK function returns a single
' combination at a specified 1-based index from the Cartesian
' product of a variable number of inputs, without computing
' the full product.
'
' This is the indexed counterpart of COMB_CARTESIAN. Both
' functions produce the same enumeration order -- index 1
' corresponds to the first row of COMB_CARTESIAN, index 2
' to the second row, and so on.
'
' Each input may contain multiple columns. Each non-empty
' column is treated as an independent dimension. Empty cells
' within a column are silently removed before processing.
'
' STRATEGY:
' 1. Normalize inputs into dimensions
'    Delegates to Cartesian_Build_Dimensions, which converts
'    each input to a 2D array, splits it into columns, removes
'    empty values, and collects the non-empty columns into a
'    1-based array of dimensions. Raises
'    ERR_COMB_INVALID_DIMENSIONS if no valid dimensions are
'    found.
'
' 2. Retrieve combination at targetIndex
'    Delegates to Cartesian_Unrank, which
'    validates the index, decomposes it using MATH_RADIX_CONVERT,
'    and maps the result to actual dimension values.
'
' PARAMETERS:
'   targetIndex   Long         1-based index of the combination
'                              to retrieve. Must be in
'                              [1, product of all dimension sizes].
'
'   inputs        ParamArray   Any combination of:
'                                - Range objects (single or multi-column)
'                                - 1D or 2D arrays
'                                - Scalar values
'                              Each column becomes one dimension.
'                              Empty cells are ignored.
'
' RETURN VALUE:
' A 1-based (1 x numDims) 2D Variant array representing the
' combination at targetIndex.
' Returns a CVErr on failure:
'   xlErrNum   if inputs produce no valid dimensions, if the
'              total combinations exceed MAX_COMBINATIONS, or
'              if targetIndex is out of range
'   xlErrValue for any other unexpected error
'
' EXAMPLE:
' ' Worksheet usage:
' =COMB_CARTESIAN_UNRANK(4, A1:A2, B1:B3)
' ' Returns the 4th combination of column A x column B
'
' ' VBA usage:
' Dim result As Variant
' result = COMB_CARTESIAN_UNRANK(4, Array("A","B"), Array(1,2,3))
' ' D1=("A","B"), D2=(1,2,3) --> 6 combinations
' ' Index 4 returns: ("B", 1)
'
' ============================================================
Public Function COMB_CARTESIAN_UNRANK( _
    ByVal targetIndex As Long, _
    ParamArray inputs() As Variant _
) As Variant

    On Error GoTo HandleError

    ' Copy ParamArray to a plain Variant before passing.
    ' VBA does not allow passing a ParamArray directly as
    ' an argument to another function.
    Dim inputsCopy As Variant
    inputsCopy = inputs

    Dim dimensions As Variant

    ' --------------------------------------------------------
    ' STEP 1: Normalize inputs --> dimensions
    ' Raises ERR_COMB_INVALID_DIMENSIONS if no valid
    ' dimensions are found across all inputs.
    ' --------------------------------------------------------
    dimensions = Cartesian_Build_Dimensions(inputsCopy)

    ' --------------------------------------------------------
    ' STEP 2: Retrieve combination at targetIndex
    ' Raises ERR_COMB_TOO_MANY or ERR_COMB_INDEX_OUT_OF_RANGE
    ' if the index or combination count is invalid.
    ' --------------------------------------------------------
    COMB_CARTESIAN_UNRANK = _
        Cartesian_Unrank(dimensions, targetIndex)

    Exit Function

HandleError:

    Select Case Err.Number
        Case ERR_COMB_TOO_MANY, ERR_COMB_INDEX_OUT_OF_RANGE, ERR_COMB_INVALID_DIMENSIONS
            COMB_CARTESIAN_UNRANK = CVErr(xlErrNum)
        Case Else
            COMB_CARTESIAN_UNRANK = CVErr(xlErrValue)
    End Select

End Function




' ============================================================
' FUNCTION: Cartesian_Build_Dimensions
' CATEGORY: Combinatorics / Cartesian Product (Internal)
'
' DESCRIPTION:
' The Cartesian_Build_Dimensions function converts a raw
' ParamArray-derived input collection into a clean 1-based
' array of dimensions, ready for consumption by the Cartesian
' product engines.
'
' Each input is normalized to a 2D array and then split into
' columns. Each non-empty column becomes one independent
' dimension. Empty cells within a column are removed by
' ARR_2D_SPLIT_COLS via ARRAY_REMOVE_EMPTY_1D.
'
' This function is the single normalization point shared by
' both COMB_CARTESIAN and COMB_CARTESIAN_UNRANK, ensuring
' both public functions produce the same dimension structure
' from identical inputs.
'
' PRECONDITIONS (caller must guarantee):
'   inputs must be a plain Variant array produced by copying
'   a ParamArray. Do not pass a ParamArray directly -- VBA
'   does not allow passing ParamArray as a function argument.
'
' STRATEGY:
' 1. Iterate over each input
'    For each element in inputs:
'      a. Normalize to 2D via ARR_TO_2D (handles scalars,
'         1D arrays, 2D arrays, and Range objects).
'      b. Split into columns via ARR_2D_SPLIT_COLS. Each
'         non-empty column becomes a 1D array of values with
'         empty cells already removed.
'      c. Collect each non-empty column into dimensions,
'         growing the array incrementally via ReDim Preserve.
'
' 2. Validate result
'    If no valid dimensions were collected, raises
'    ERR_COMB_INVALID_DIMENSIONS. Callers rely on this error
'    rather than checking the return value.
'
' PARAMETERS:
'   inputs    Variant   Plain Variant array containing the
'                       original ParamArray values. Produced
'                       by: inputsCopy = inputs inside a
'                       ParamArray function.
'
' RETURN VALUE:
' A 1-based 1D array of 1D arrays. Each element is one
' dimension (a cleaned column from the inputs).
' Never returns Empty -- raises ERR_COMB_INVALID_DIMENSIONS
' instead if no valid dimensions are found.
'
' ============================================================
Private Function Cartesian_Build_Dimensions(ByRef inputs As Variant) As Variant

    Dim dimensions() As Variant
    Dim dimCount As Long
    Dim i As Long
    Dim j As Long

    dimCount = 0

    ' --------------------------------------------------------
    ' STEP 1: Normalize each input and collect dimensions
    ' --------------------------------------------------------
    For i = LBound(inputs) To UBound(inputs)

        Dim splitColumns As Variant

        ' Normalize to 2D, then split into cleaned columns.
        ' ARR_2D_SPLIT_COLS removes empty cells from each
        ' column internally via ARRAY_REMOVE_EMPTY_1D.
        splitColumns = ARR_TO_2D(inputs(i))
        splitColumns = ARR_2D_SPLIT_COLS(splitColumns)

        If Not IsEmpty(splitColumns) Then

            For j = LBound(splitColumns) To UBound(splitColumns)

                ' Slots left Empty by ARR_2D_SPLIT_COLS
                ' correspond to fully-empty columns -- skip them.
                If Not IsEmpty(splitColumns(j)) Then

                    dimCount = dimCount + 1
                    ReDim Preserve dimensions(1 To dimCount)
                    dimensions(dimCount) = splitColumns(j)

                End If

            Next j

        End If

    Next i

    ' --------------------------------------------------------
    ' STEP 2: Validate that at least one dimension was found.
    ' Signal via Err.Raise so callers can handle through their
    ' existing error handler without checking the return value.
    ' --------------------------------------------------------
    If dimCount = 0 Then
        Err.Raise ERR_COMB_INVALID_DIMENSIONS, , "No valid dimensions found in inputs"
    End If

    Cartesian_Build_Dimensions = dimensions

End Function








' ============================================================
' FUNCTION: VarlenRadix_Unrank
' CATEGORY: Combinatorics / Variable-Length Radix (Internal)
'
' DESCRIPTION:
' The VarlenRadix_Unrank function returns the single
' combination at a given 1-based index from the variable-
' length radix space defined by valueSet, without computing
' the full enumeration.
'
' Unlike Cartesian_Unrank, where the number of output columns
' is fixed by the number of dimensions, here the output length
' is determined by the index itself: higher indices map to
' longer combinations. The enumeration order is:
'
'   Index 1         --> all 1-column combinations
'   Index N+1       --> all 2-column combinations
'   Index N^2+N+1   --> all 3-column combinations
'   ...
'
' where N is the size of valueSet.
'
' PRECONDITIONS (caller must guarantee):
'   valueSet must be a non-empty 1-based 1D String array.
'   targetIndex must be >= 1.
'   Use ARR_TO_1D_STRING to normalize inputs before calling.
'
' STRATEGY:
' 1. Compute the size of valueSet
'    The number of available values (valueSetSize) determines
'    the base of the radix system and the block sizes.
'    Exits silently if valueSet is empty.
'
' 2. Validate the lower index bound
'    Raises ERR_COMB_INDEX_OUT_OF_RANGE if targetIndex < 1.
'
' 3. Locate the block containing targetIndex
'    Combinations are grouped into blocks by output length:
'    block 1 has N combinations (length 1), block 2 has N^2
'    (length 2), etc. The loop advances through blocks until
'    the cumulative total reaches targetIndex. An overflow
'    guard raises ERR_COMB_INDEX_OUT_OF_RANGE if the space
'    exceeds Double precision.
'
' 4. Compute the local index within the block
'    Subtracts the cumulative total of all preceding blocks
'    to obtain the 1-based position within the current block.
'
' 5. Convert local index to zero-based
'    Subtracts 1 before passing to MATH_RADIX_CONVERT, which
'    expects a zero-based value.
'
' 6. Decompose index into per-position digits
'    Calls MATH_RADIX_CONVERT with a uniform base (valueSetSize)
'    repeated once per output column. The result is a 1-based
'    array of 0-based position indices.
'
' 7. Map digits to actual values
'    Each digit is used to index into valueSet, offset by
'    LBound to support any lower bound.
'
' PARAMETERS:
'   valueSet      Variant   1-based 1D String array defining
'                           the set of values for each position.
'
'   targetIndex   Long      1-based index of the combination
'                           to retrieve. Must be >= 1.
'
' RETURN VALUE:
' A 1-based (1 x comboLen) 2D Variant array representing the
' combination at targetIndex. Returns Empty if valueSet is
' empty.
' Note: comboLen is the length of the combination at targetIndex.
'
' EXAMPLE:
' ' valueSet = ("A", "B"), combinations in order:
' '   Index 1 --> ("A")
' '   Index 2 --> ("B")
' '   Index 3 --> ("A", "A")
' '   Index 4 --> ("A", "B")
' '   Index 5 --> ("B", "A")
' '   Index 6 --> ("B", "B")
' '   Index 7 --> ("A", "A", "A")
' '   ...
'
' ============================================================
Private Function VarlenRadix_Unrank( _
    ByVal valueSet As Variant, _
    ByVal targetIndex As Long _
) As Variant

    Dim lb As Long, ub As Long
    lb = LBound(valueSet)
    ub = UBound(valueSet)

    Dim valueSetSize As Long
    valueSetSize = ub - lb + 1

    ' --------------------------------------------------------
    ' STEP 1: Compute size of valueSet
    ' An empty valueSet makes the result undefined --> exit
    ' silently and return Empty implicitly.
    ' --------------------------------------------------------
    If valueSetSize = 0 Then Exit Function

    ' --------------------------------------------------------
    ' STEP 2: Validate lower index bound
    ' --------------------------------------------------------
    If targetIndex < 1 Then
        Err.Raise ERR_COMB_INDEX_OUT_OF_RANGE, , "Index out of range"
    End If

    ' --------------------------------------------------------
    ' STEP 3: Locate the block containing targetIndex
    ' Combinations are grouped by output length. Block k
    ' contains all combinations of length k (valueSetSize^k
    ' combinations). The loop advances until the running total
    ' reaches or exceeds targetIndex.
    ' An overflow guard raises ERR_COMB_INDEX_OUT_OF_RANGE if
    ' the space exceeds Double precision before targetIndex
    ' is reached.
    ' --------------------------------------------------------
    Dim comboLen As Long:       comboLen = 1
    Dim prevBlocksTotal As Double: prevBlocksTotal = 0
    Dim blockSize As Double:    blockSize = 1
    Dim runningTotal As Double

    Do While targetIndex > runningTotal
        If blockSize > MAX_SAFE_DOUBLE / valueSetSize Then
            Err.Raise ERR_COMB_INDEX_OUT_OF_RANGE, , _
                "Combination space exceeds Double precision limit"
        End If
        blockSize = blockSize * valueSetSize
        runningTotal = blockSize + prevBlocksTotal
        prevBlocksTotal = runningTotal
        comboLen = comboLen + 1
    Loop

    ' --------------------------------------------------------
    ' STEP 4: Compute local index within the block (1-based)
    ' Subtracts the total of all preceding blocks.
    ' --------------------------------------------------------
    Dim blockIndex As Long
    blockIndex = targetIndex - (runningTotal - blockSize)

    ' --------------------------------------------------------
    ' STEP 5: Build the uniform radix sizes array
    ' MATH_RADIX_CONVERT expects a (1 x N) 2D array where
    ' each element is the base for that position.
    ' --------------------------------------------------------
    Dim i As Long
    Dim radixSizes() As Variant
    ReDim radixSizes(1 To 1, 1 To comboLen - 1)
    For i = 1 To comboLen - 1
        radixSizes(1, i) = valueSetSize
    Next i

    ' --------------------------------------------------------
    ' STEP 6: Convert local index to zero-based
    ' MATH_RADIX_CONVERT produces digits in [0, size-1], so
    ' the index must be zero-based before decomposition.
    ' --------------------------------------------------------
    Dim zeroIndex As Long
    zeroIndex = blockIndex - 1

    ' --------------------------------------------------------
    ' STEP 7: Decompose index into per-position digits
    ' Treats radixSizes as a uniform base system. The result
    ' is a 1-based array of 0-based position indices, one
    ' per output column.
    ' --------------------------------------------------------
    Dim digits As Variant
    digits = MATH_RADIX_CONVERT(zeroIndex, 10, radixSizes)

    ' --------------------------------------------------------
    ' STEP 8: Map digits to actual values in valueSet
    ' Each digit is the 0-based position within valueSet.
    ' The lb offset makes the lookup safe for any lower bound.
    ' --------------------------------------------------------
    Dim result() As Variant
    ReDim result(1 To 1, 1 To comboLen - 1)

    Dim dimLB As Long
    dimLB = LBound(valueSet)
    For i = 1 To comboLen - 1
        result(1, i) = valueSet(dimLB + digits(i))
    Next i

    VarlenRadix_Unrank = result

End Function





' ============================================================
' FUNCTION: COMB_UNRANK_VARLEN
' CATEGORY: Combinatorics / Variable-Length Radix
'
' DESCRIPTION:
' The COMB_UNRANK_VARLEN function returns the single
' combination at a given 1-based index from the variable-
' length radix space defined by the input values, without
' computing the full enumeration.
'
' The enumeration groups combinations by length: all 1-element
' combinations come first, then all 2-element combinations,
' and so on. Within each group, combinations are ordered by
' their position in the radix system defined by the input
' values.
'
' STRATEGY:
' 1. Normalize input --> valueSet
'    Delegates to ARR_TO_1D_STRING, which converts any
'    supported input into a 1-based 1D String array. Raises
'    ERR_ARR_NOT_VECTOR if input is not a scalar or 1D vector.
'    Raises ERR_ARR_INVALID_VALUE if input contains error
'    values (e.g. #N/A, #DIV/0!).
'
' 2. Retrieve combination at targetIndex
'    Delegates to VarlenRadix_Unrank, which locates the block
'    containing targetIndex, decomposes the local index using
'    MATH_RADIX_CONVERT, and maps the result to actual values
'    in valueSet.
'
' PARAMETERS:
'   targetIndex   Long      1-based index of the combination
'                           to retrieve. Must be >= 1.
'
'   inputs        Variant   Any combination of:
'                             - Scalar value
'                             - 1D array
'                             - Single-column or single-row
'                               Range
'                           Each element becomes one possible
'                           value in the output. Error values
'                           are rejected.
'
' RETURN VALUE:
' A 1-based (1 x comboLen [combinatios length]) 2D Variant
' array representing the combination at targetIndex.
' Returns a CVErr on failure:
'   xlErrNum   if input is not a valid 1D vector, if it
'              contains error values, if targetIndex < 1,
'              or if the combination space exceeds Double
'              precision
'   xlErrValue for any other unexpected error
'
' EXAMPLE:
' ' Worksheet usage:
' =COMB_UNRANK_VARLEN(4, A1:A2)
' ' Returns the 4th combination of the values in A1:A2
'
' ' VBA usage:
' Dim result As Variant
' result = COMB_UNRANK_VARLEN(4, Array("A", "B"))
' ' valueSet = ("A", "B"), combinations in order:
' '   Index 1 --> ("A")
' '   Index 2 --> ("B")
' '   Index 3 --> ("A", "A")
' '   Index 4 --> ("A", "B")  <-- returns this
'
' ============================================================
Public Function COMB_UNRANK_VARLEN( _
    ByVal targetIndex As Long, _
    ByVal inputs As Variant _
) As Variant

    On Error GoTo HandleError

    Dim valueSet As Variant

    ' --------------------------------------------------------
    ' STEP 1: Normalize input --> valueSet
    ' Raises ERR_ARR_NOT_VECTOR if input is not a scalar or
    ' 1D vector. Raises ERR_ARR_INVALID_VALUE if input contains
    ' error values (e.g. #N/A, #DIV/0!).
    ' --------------------------------------------------------
    valueSet = ARR_TO_1D_STRING(inputs)

    ' --------------------------------------------------------
    ' STEP 2: Retrieve combination at targetIndex
    ' Raises ERR_COMB_INDEX_OUT_OF_RANGE if targetIndex < 1
    ' or if the combination space exceeds Double precision.
    ' --------------------------------------------------------
    COMB_UNRANK_VARLEN = _
        VarlenRadix_Unrank(valueSet, targetIndex)

    Exit Function

HandleError:

    Select Case Err.Number
        Case ERR_COMB_TOO_MANY, ERR_COMB_INDEX_OUT_OF_RANGE, _
             ERR_ARR_NOT_VECTOR, ERR_ARR_INVALID_VALUE
            COMB_UNRANK_VARLEN = CVErr(xlErrNum)
        Case Else
            COMB_UNRANK_VARLEN = CVErr(xlErrValue)
    End Select

End Function

