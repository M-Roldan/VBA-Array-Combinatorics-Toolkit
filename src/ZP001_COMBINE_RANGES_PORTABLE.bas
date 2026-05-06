Attribute VB_Name = "ZP001_COMBINE_RANGES_PORTABLE"
Option Explicit

' ============================================================
' MODULE: ZP001_COMBINE_RANGES_PORTABLE
' CATEGORY: Combinatorics / Cartesian Product
'
' PURPOSE:
' This module provides a self-contained, portable implementation
' of the Cartesian product function P001_COMBINE_RANGES.
' It is designed to be copied and pasted into any VBA project
' without requiring any other module from the library.
'
' ------------------------------------------------------------
' PUBLIC API
' ------------------------------------------------------------
'
'   P001_COMBINE_RANGES
'     Generates the full Cartesian product of all inputs and
'     returns it as a 2D array (rows = combinations,
'     columns = dimensions).
'
' ------------------------------------------------------------
' INTERNAL FUNCTIONS
' ------------------------------------------------------------
'
'   P001_Cartesian_Build_Full
'     Core engine. Builds the full Cartesian product using
'     a block repetition algorithm. Each dimension fills its
'     column by repeating values in blocks whose size decreases
'     as dimensions are processed left to right.
'
'   P001_ARR_NDIM
'     Returns the number of dimensions of an array.
'
'   P001_ARR_TO_2D
'     Normalizes any input into a canonical 2D array.
'
'   P001_ARR_SHAPE
'     Returns the size of each dimension as a 1D array.
'
'   P001_ARR_2D_GET_COL
'     Extracts a single column from a 2D array as a 1D array.
'
'   P001_ARR_1D_REMOVE_EMPTY
'     Removes empty strings, Empty variants, and error values
'     from a 1D array.
'
'   P001_ARR_2D_SPLIT_COLS
'     Splits a 2D array into a 1-based array of cleaned columns.
'
' ------------------------------------------------------------
' DESIGN CONVENTIONS
' ------------------------------------------------------------
'
' 1. SELF-CONTAINED
'    This module has no dependencies on any other module in
'    the library. All required utilities are implemented
'    locally with the P001_ prefix.
'
' 2. ERROR SIGNALING
'    This module does not participate in the centralized error
'    system (LibError enum in ERRORS.bas). Overflow is signaled
'    via vbObjectError + 1 and caught by P001_COMBINE_RANGES,
'    which returns CVErr(xlErrNum).
'
' 3. COMBINATION LIMITS
'    Two independent limits are enforced in P001_Cartesian_Build_Full:
'
'      P001_MAX_LONG        -- hard arithmetic limit. The Long
'                              data type cannot exceed 2^31 - 1.
'                              This limit is not adjustable.
'
'      P001_MAX_COMBINATIONS -- soft performance limit. Controls
'                              how many rows the engine will
'                              produce. Increase with caution --
'                              large values may cause Excel to
'                              become unresponsive.
'
' 4. INDEXING
'    All arrays produced by this module are 1-based,
'    consistent with the library-wide convention.
'
' 5. ALGORITHM
'    The Cartesian product engine uses block repetition rather
'    than positional weight decomposition. This algorithm is
'    more intuitive to read in isolation, which is appropriate
'    for a portable module intended to be understood without
'    the broader library context.
'
' ------------------------------------------------------------
' PORTABLE MODULE NOTICE
' ------------------------------------------------------------
'
'    This module is a portable, self-contained version of the
'    functionality provided by COMB_CARTESIAN in the
'    COMBINATORICS module. It is intended for use in projects
'    where deploying the full library is not practical.
'
'    If the full library is available, prefer COMB_CARTESIAN
'    over P001_COMBINE_RANGES, as it uses a more efficient
'    engine (positional weight decomposition) and participates
'    in the centralized error system.
'
' ============================================================

' User-adjustable performance limit.
' Controls the maximum number of combinations P001_COMBINE_RANGES
' will produce. Increase with caution -- large values may cause
' Excel to become unresponsive.
Private Const P001_MAX_COMBINATIONS As Long = 200000

' Internal arithmetic guard. Long overflows above 2^31 - 1.
' This limit is not adjustable -- it is a property of the
' Long data type used for the total combination count.
Private Const P001_MAX_LONG As Long = 2147483647


' ============================================================
' FUNCTION: P001_ARR_NDIM
' CATEGORY: Combinatorics / Array Utilities (Internal)
'
' DESCRIPTION:
' The P001_ARR_NDIM function returns the number of dimensions
' of an array by probing successive dimensions using On Error.
' Returns 0 for non-arrays and scalars.
'
' STRATEGY:
' 1. If input is a Range, extract .Value first.
' 2. If input is not an array, return 0.
' 3. Probe dimensions 1 to 60 (VBA maximum) using UBound.
'    The loop exits via the error handler when a dimension
'    does not exist. The last successfully probed dimension
'    is returned.
'
' PARAMETERS:
'   inputArray  Variant   The value to inspect. Accepted forms:
'                           - Array of any dimensionality
'                           - Range (value is extracted via .Value)
'                           - Scalar or non-array (returns 0)
'
' RETURN VALUE:
' A Long representing the number of dimensions:
'   0 --> Not an array
'   1 --> 1D array
'   2 --> 2D array
'   N --> N-dimensional array
'
' ============================================================
Private Function P001_ARR_NDIM(ByVal inputArray As Variant) As Long

    ' Extract value from Range before inspection
    If TypeName(inputArray) = "Range" Then
        inputArray = inputArray.value
    End If

    ' Non-arrays and scalars have 0 dimensions
    If Not IsArray(inputArray) Then
        P001_ARR_NDIM = 0
        Exit Function
    End If

    ' --------------------------------------------------------
    ' Probe dimensions using On Error.
    ' UBound raises an error when the dimension does not exist.
    ' The last successfully probed index is the dimension count.
    ' --------------------------------------------------------
    Dim dimensionIndex As Long
    Dim tmp As Long

    On Error GoTo ExitFunction

    For dimensionIndex = 1 To 60
        tmp = UBound(inputArray, dimensionIndex)
        P001_ARR_NDIM = dimensionIndex
    Next dimensionIndex

ExitFunction:
    On Error GoTo 0

End Function


' ============================================================
' FUNCTION: P001_ARR_TO_2D
' CATEGORY: Combinatorics / Array Utilities (Internal)
'
' DESCRIPTION:
' The P001_ARR_TO_2D function normalizes any input into a
' canonical 1-based 2D array:
'   - Scalars         --> (1 x 1)
'   - 1D arrays       --> (N x 1) column vector
'   - 2D arrays       --> returned as-is
'   - Range           --> .Value extracted; single-cell Ranges
'                         are wrapped into (1 x 1)
'
' STRATEGY:
' 1. Handle Range input
'    Extracts .Value. Single-cell Ranges return a scalar and
'    are treated as the scalar case below.
'
' 2. Handle 1D array
'    Reshapes into a (N x 1) 2D column vector using 1-based
'    indexing regardless of the original lower bound.
'
' 3. Handle 2D or higher array
'    Returned as-is. No reshaping is performed.
'
' 4. Handle scalar
'    Wrapped into a (1 x 1, 1-based) 2D array.
'
' PARAMETERS:
'   inputValue  Variant   The value to normalize. Accepted forms:
'                           - Range (single cell or multi-cell)
'                           - 1D array
'                           - 2D array (returned as-is)
'                           - Scalar
'
' RETURN VALUE:
' A 1-based 2D Variant array.
'
' ============================================================
Private Function P001_ARR_TO_2D(ByVal inputValue As Variant) As Variant

    ' --------------------------------------------------------
    ' CASE 1: Range input
    ' Extract .Value. A single-cell Range returns a scalar,
    ' which falls through to the scalar case below.
    ' A multi-cell Range returns a 1-based 2D array directly.
    ' --------------------------------------------------------
    If TypeName(inputValue) = "Range" Then
        Dim rangeVal As Variant
        rangeVal = inputValue.value
        If IsArray(rangeVal) Then
            P001_ARR_TO_2D = rangeVal
            Exit Function
        Else
            ' Single-cell Range: wrap scalar into (1 x 1)
            Dim singleCell(1 To 1, 1 To 1) As Variant
            singleCell(1, 1) = rangeVal
            P001_ARR_TO_2D = singleCell
            Exit Function
        End If
    End If

    ' --------------------------------------------------------
    ' CASE 2: Array input
    ' --------------------------------------------------------
    If IsArray(inputValue) Then

        If P001_ARR_NDIM(inputValue) = 1 Then

            ' ------------------------------------------------
            ' CASE 2a: 1D array -- reshape to (N x 1)
            ' 1-based indexing regardless of original LBound.
            ' ------------------------------------------------
            Dim i As Long
            Dim result() As Variant
            Dim lb As Long: lb = LBound(inputValue)
            Dim ub As Long: ub = UBound(inputValue)

            ReDim result(1 To ub - lb + 1, 1 To 1)

            For i = lb To ub
                result(i - lb + 1, 1) = inputValue(i)
            Next i

            P001_ARR_TO_2D = result
            Exit Function

        End If

        ' ----------------------------------------------------
        ' CASE 2b: Already 2D or higher -- return as-is
        ' ----------------------------------------------------
        P001_ARR_TO_2D = inputValue
        Exit Function

    End If

    ' --------------------------------------------------------
    ' CASE 3: Scalar -- wrap into (1 x 1)
    ' --------------------------------------------------------
    Dim temp(1 To 1, 1 To 1) As Variant
    temp(1, 1) = inputValue
    P001_ARR_TO_2D = temp

End Function


' ============================================================
' FUNCTION: P001_ARR_SHAPE
' CATEGORY: Combinatorics / Array Utilities (Internal)
'
' DESCRIPTION:
' The P001_ARR_SHAPE function returns the size of each
' dimension of an array as a 1-based 1D Long array.
' Returns Empty for non-arrays.
'
' PARAMETERS:
'   arr   Variant   The array to inspect. Any dimensionality
'                   is supported. Non-arrays return Empty.
'
' RETURN VALUE:
' A 1-based 1D Long array where each element is the size of
' the corresponding dimension. Returns Empty for non-arrays.
'
' ============================================================
Private Function P001_ARR_SHAPE(ByVal arr As Variant) As Variant

    Dim dims As Long
    dims = P001_ARR_NDIM(arr)

    ' Non-arrays return Empty implicitly
    If dims = 0 Then Exit Function

    Dim result() As Long
    ReDim result(1 To dims)

    Dim i As Long
    For i = 1 To dims
        result(i) = UBound(arr, i) - LBound(arr, i) + 1
    Next i

    P001_ARR_SHAPE = result

End Function


' ============================================================
' FUNCTION: P001_ARR_2D_GET_COL
' CATEGORY: Combinatorics / Array Utilities (Internal)
'
' DESCRIPTION:
' The P001_ARR_2D_GET_COL function extracts a single column
' from a 2D array and returns it as a 1-based 1D array.
'
' PRECONDITIONS (caller must guarantee):
'   arr must be a 2D array with 1-based indexing.
'   columnIndex must be >= 1 and <= number of columns.
'
' PARAMETERS:
'   arr           Variant   A 1-based 2D array to extract from.
'   columnIndex   Long      1-based index of the column to extract.
'
' RETURN VALUE:
' A 1-based 1D Variant array containing the values of the
' specified column. Returns Empty if columnIndex is out of range.
'
' ============================================================
Private Function P001_ARR_2D_GET_COL( _
    ByVal arr As Variant, _
    ByVal columnIndex As Long _
) As Variant

    Dim shape As Variant
    shape = P001_ARR_SHAPE(arr)

    Dim rowCount As Long
    rowCount = shape(1)

    ' Return Empty silently for out-of-range column index
    If columnIndex < 1 Or columnIndex > shape(2) Then Exit Function

    Dim result() As Variant
    Dim r As Long

    ReDim result(1 To rowCount)

    ' Extract values from the specified column
    For r = 1 To rowCount
        result(r) = arr(r, columnIndex)
    Next r

    P001_ARR_2D_GET_COL = result

End Function


' ============================================================
' FUNCTION: P001_ARR_1D_REMOVE_EMPTY
' CATEGORY: Combinatorics / Array Utilities (Internal)
'
' DESCRIPTION:
' The P001_ARR_1D_REMOVE_EMPTY function removes empty strings,
' Empty variants, and error values from a 1D array and returns
' a compacted 1-based result.
'
' Uses a two-pass strategy: first counts valid elements to
' enable exact allocation, then copies them into the result.
'
' PARAMETERS:
'   arr   Variant   A 1D array to clean. May use any lower bound.
'                   Error values, empty strings, and Empty
'                   variants are all discarded.
'
' RETURN VALUE:
' A 1-based 1D Variant array containing only valid elements.
' Returns Empty if no valid elements remain.
'
' ============================================================
Private Function P001_ARR_1D_REMOVE_EMPTY(ByVal arr As Variant) As Variant

    Dim i As Long
    Dim count As Long

    ' --------------------------------------------------------
    ' PASS 1: Count valid (non-empty, non-error) elements
    ' --------------------------------------------------------
    For i = LBound(arr) To UBound(arr)
        If Not IsError(arr(i)) Then
            If arr(i) <> "" And Not IsEmpty(arr(i)) Then
                count = count + 1
            End If
        End If
    Next i

    ' Return Empty implicitly if no valid elements found
    If count = 0 Then Exit Function

    Dim temp() As Variant
    ReDim temp(1 To count)

    count = 0

    ' --------------------------------------------------------
    ' PASS 2: Copy valid elements into result array
    ' --------------------------------------------------------
    For i = LBound(arr) To UBound(arr)
        If Not IsError(arr(i)) Then
            If arr(i) <> "" And Not IsEmpty(arr(i)) Then
                count = count + 1
                temp(count) = arr(i)
            End If
        End If
    Next i

    P001_ARR_1D_REMOVE_EMPTY = temp

End Function


' ============================================================
' FUNCTION: P001_ARR_2D_SPLIT_COLS
' CATEGORY: Combinatorics / Array Utilities (Internal)
'
' DESCRIPTION:
' The P001_ARR_2D_SPLIT_COLS function splits a 2D array into
' a 1-based array of cleaned 1D column arrays.
'
' Each column is extracted via P001_ARR_2D_GET_COL and cleaned
' via P001_ARR_1D_REMOVE_EMPTY. Fully empty columns produce
' Empty slots in the result array rather than being omitted,
' preserving the correspondence between slot index and original
' column number.
'
' PARAMETERS:
'   arr   Variant   A 1-based 2D array to split.
'
' RETURN VALUE:
' A 1-based 1D array of 1D arrays. Each element corresponds
' to one column. Empty slots indicate fully-empty columns.
'
' ============================================================
Private Function P001_ARR_2D_SPLIT_COLS(ByVal arr As Variant) As Variant

    Dim shape As Variant
    shape = P001_ARR_SHAPE(arr)

    Dim colCount As Long
    colCount = shape(2)

    Dim result() As Variant
    ReDim result(1 To colCount)

    Dim c As Long
    Dim colData As Variant

    For c = 1 To colCount

        ' Extract column and remove empty values
        colData = P001_ARR_2D_GET_COL(arr, c)
        colData = P001_ARR_1D_REMOVE_EMPTY(colData)

        ' Empty slots correspond to fully-empty columns.
        ' Callers check IsEmpty(result(c)) to skip them.
        If Not IsEmpty(colData) Then
            result(c) = colData
        End If

    Next c

    P001_ARR_2D_SPLIT_COLS = result

End Function


' ============================================================
' FUNCTION: P001_Cartesian_Build_Full
' CATEGORY: Combinatorics / Cartesian Product (Internal)
'
' DESCRIPTION:
' The P001_Cartesian_Build_Full function generates the full
' Cartesian product of a set of dimensions using a block
' repetition algorithm.
'
' Each dimension fills its column in the result by repeating
' its values in consecutive blocks. The block size starts at
' total / size(D1) and decreases with each dimension, producing
' the correct enumeration order.
'
' EXAMPLE (D1={"A","B"}, D2={1,2}, D3={"#","$"}):
'
'   repeatBlock progression:
'     D1: block = 4  --> AAAABBBB
'     D2: block = 2  --> 11221122
'     D3: block = 1  --> #$#$#$#$
'
'   Result:
'     A 1 #
'     A 1 $
'     A 2 #
'     A 2 $
'     B 1 #
'     B 1 $
'     B 2 #
'     B 2 $
'
' KEY FORMULA:
'   repeatBlock = total \ size(current dimension)
'   It decreases at each step:
'   total --> total/n1 --> total/(n1*n2) --> ...
'
' NOTE ON ALGORITHM CHOICE:
'   This module uses block repetition rather than the positional
'   weight (divisor) algorithm used by COMBINATORICS. Block
'   repetition is more intuitive to read in isolation, which
'   is appropriate for a portable module intended to be
'   understood without the broader library context.
'
' COMBINATION LIMITS:
'   Two independent guards are applied during size computation:
'
'   Guard 1 (arithmetic): prevents Long overflow before any
'   multiplication. Uses P001_MAX_LONG (2^31 - 1). Not
'   adjustable -- determined by the Long data type.
'
'   Guard 2 (performance): enforces P001_MAX_COMBINATIONS after
'   each multiplication. Adjustable by the caller via the
'   module constant. Prevents Excel from becoming unresponsive
'   on excessively large products.
'
' PRECONDITIONS (caller must guarantee):
'   dimensions must be a 1-based 1D array of 1D arrays.
'   Each inner array must be non-empty.
'
' PARAMETERS:
'   dimensions  Variant   1-based 1D array of 1D arrays.
'                         Each element is one dimension.
'
' RETURN VALUE:
' A 1-based 2D Variant array (total x numDims) where each
' row is one combination and each column is one dimension.
' Returns Empty if any dimension has size 0.
' Raises vbObjectError + 1 if any guard is exceeded.
'
' ============================================================
Private Function P001_Cartesian_Build_Full( _
    ByVal dimensions As Variant _
) As Variant

    Dim lb As Long, ub As Long
    lb = LBound(dimensions)
    ub = UBound(dimensions)

    Dim numDims As Long
    numDims = ub - lb + 1

    Dim sizes() As Long
    ReDim sizes(lb To ub)

    Dim total As Long
    total = 1

    Dim i As Long

    ' --------------------------------------------------------
    ' STEP 1: Compute sizes and validate total
    '
    ' Two guards are applied on each iteration:
    '   Guard 1: arithmetic -- uses integer division to check
    '            if multiplying would overflow Long capacity.
    '            Must run BEFORE the multiplication.
    '   Guard 2: performance -- checks if the updated total
    '            exceeds the user-configurable limit.
    '            Runs AFTER the multiplication on the updated
    '            value of total.
    ' --------------------------------------------------------
    For i = lb To ub

        sizes(i) = UBound(dimensions(i)) - LBound(dimensions(i)) + 1

        ' An empty dimension makes the Cartesian product empty
        If sizes(i) = 0 Then
            P001_Cartesian_Build_Full = Empty
            Exit Function
        End If

        ' Guard 1: prevent Long overflow before multiplying.
        ' Equivalent to checking total * sizes(i) > MAX_LONG
        ' without executing the multiplication.
        If total > P001_MAX_LONG \ sizes(i) Then
            Err.Raise vbObjectError + 1, , _
                "P001_Cartesian_Build_Full: combination space " & _
                "exceeds Long capacity"
        End If

        total = total * sizes(i)

        ' Guard 2: enforce performance limit after multiplying.
        ' P001_MAX_COMBINATIONS is user-adjustable via the
        ' module constant at the top of this file.
        If total > P001_MAX_COMBINATIONS Then
            Err.Raise vbObjectError + 1, , _
                "P001_Cartesian_Build_Full: too many combinations " & _
                "(" & total & "). Limit: " & P001_MAX_COMBINATIONS
        End If

    Next i

    ' --------------------------------------------------------
    ' STEP 2: Allocate result (1-based, total rows x numDims cols)
    ' --------------------------------------------------------
    Dim result() As Variant
    ReDim result(1 To total, 1 To numDims)

    ' --------------------------------------------------------
    ' STEP 3: Fill using block repetition pattern.
    '
    ' repeatBlock controls how many consecutive rows each value
    ' occupies before advancing to the next value in the
    ' dimension. It starts at total and is divided by each
    ' dimension size as we move left to right.
    '
    ' The outer Do While loop cycles through all values of the
    ' current dimension repeatedly until all rows are filled.
    ' --------------------------------------------------------
    Dim repeatBlock As Long
    Dim rowIndex As Long
    Dim repeatIndex As Long
    Dim j As Long

    repeatBlock = total

    For i = lb To ub

        ' Reduce block size for this dimension
        repeatBlock = repeatBlock \ sizes(i)

        rowIndex = 1

        Do While rowIndex <= total

            For j = LBound(dimensions(i)) To UBound(dimensions(i))

                ' Repeat this value for repeatBlock consecutive rows
                For repeatIndex = 1 To repeatBlock
                    result(rowIndex, i - lb + 1) = dimensions(i)(j)
                    rowIndex = rowIndex + 1
                Next repeatIndex

            Next j

        Loop

    Next i

    P001_Cartesian_Build_Full = result

End Function


' ============================================================
' FUNCTION: P001_COMBINE_RANGES
' CATEGORY: Combinatorics / Cartesian Product
'
' DESCRIPTION:
' The P001_COMBINE_RANGES function generates the full Cartesian
' product of a variable number of inputs and returns it as a
' 2D array where each row is one unique combination and each
' column corresponds to one dimension.
'
' Each input may contain multiple columns. Each non-empty
' column is treated as an independent dimension. Empty cells
' within a column are silently removed before processing.
'
' NOTE:
' This is the portable equivalent of COMB_CARTESIAN in the
' COMBINATORICS module. If the full library is available,
' prefer COMB_CARTESIAN as it uses a more efficient engine.
'
' STRATEGY:
' 1. Normalize inputs into dimensions
'    For each input: normalize to 2D via P001_ARR_TO_2D,
'    split into columns via P001_ARR_2D_SPLIT_COLS, and
'    collect each non-empty column as one dimension.
'    Returns CVErr(xlErrValue) if no valid dimensions are found.
'
' 2. Delegate to the Cartesian engine
'    Passes the normalized dimensions array to
'    P001_Cartesian_Build_Full, which produces the full result
'    using block repetition.
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
'   xlErrNum   if total combinations exceed P001_MAX_COMBINATIONS
'              or if the combination space exceeds Long capacity
'   xlErrValue for invalid inputs or any other unexpected error
'
' EXAMPLE:
' ' Worksheet usage:
' =P001_COMBINE_RANGES(A1:A3, B1:B2)
' ' Returns all 6 combinations of column A x column B
'
' ' VBA usage:
' Dim result As Variant
' result = P001_COMBINE_RANGES(Array("A","B"), Array(1,2,3))
' ' Returns:
' '   A 1
' '   A 2
' '   A 3
' '   B 1
' '   B 2
' '   B 3
'
' ============================================================
Public Function P001_COMBINE_RANGES(ParamArray inputs() As Variant) As Variant

    On Error GoTo HandleError

    Dim dimensions() As Variant
    Dim dimCount As Long
    Dim i As Long
    Dim j As Long

    dimCount = 0

    ' --------------------------------------------------------
    ' STEP 1: Normalize each input into dimensions.
    ' Each non-empty column of each input becomes one dimension.
    ' --------------------------------------------------------
    For i = LBound(inputs) To UBound(inputs)

        Dim cols As Variant

        ' Normalize to 2D, then split into cleaned columns
        cols = P001_ARR_TO_2D(inputs(i))
        cols = P001_ARR_2D_SPLIT_COLS(cols)

        If Not IsEmpty(cols) Then

            For j = LBound(cols) To UBound(cols)

                ' Slots left Empty by P001_ARR_2D_SPLIT_COLS
                ' correspond to fully-empty columns -- skip them
                If Not IsEmpty(cols(j)) Then

                    dimCount = dimCount + 1
                    ReDim Preserve dimensions(1 To dimCount)
                    dimensions(dimCount) = cols(j)

                End If

            Next j

        End If

    Next i

    ' --------------------------------------------------------
    ' STEP 2: Validate that at least one dimension was found
    ' --------------------------------------------------------
    If dimCount = 0 Then
        P001_COMBINE_RANGES = CVErr(xlErrValue)
        Exit Function
    End If

    ' --------------------------------------------------------
    ' STEP 3: Delegate to block repetition Cartesian engine
    ' --------------------------------------------------------
    P001_COMBINE_RANGES = P001_Cartesian_Build_Full(dimensions)

    Exit Function

HandleError:

    Select Case Err.Number
        Case vbObjectError + 1
            ' Arithmetic overflow or combination limit exceeded
            P001_COMBINE_RANGES = CVErr(xlErrNum)
        Case Else
            P001_COMBINE_RANGES = CVErr(xlErrValue)
    End Select

End Function


