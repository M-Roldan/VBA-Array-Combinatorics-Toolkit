Attribute VB_Name = "ARRAYS"
Option Explicit

' ============================================================
' MODULE: ARRAYS
' CATEGORY: Array Utilities / Structural Normalization
'
' PURPOSE:
' This module provides utilities for inspecting, normalizing,
' transforming and cleaning VBA arrays. It serves as the
' structural foundation used by higher-level modules such as
' MATH (Radix engine) and COMBINATORICS (Cartesian product).
'
' ------------------------------------------------------------
' DESIGN CONVENTIONS
' ------------------------------------------------------------
'
' 1. INDEXING: 1-BASED
'    All arrays produced by this module use 1-based indexing.
'    This matches Excel's native range convention and avoids
'    off-by-one coordination between functions.
'
'      • ARR_TO_2D           --> (1 To rows, 1 To cols)
'      • ARR_SHAPE           --> (1 To dims)
'      • ARR_2D_GET_COL      --> (1 To rows)
'      • ARR_2D_SPLIT_COLS   --> (1 To cols)
'      • ARR_TO_1D_LONG      --> (1 To n)
'      • ARR_TO_1D_STRING    --> (1 To n)
'
'    Input arrays may use any lower bound. Output arrays
'    ALWAYS use lower bound 1.
'
' 2. CANONICAL SHAPE: 2D
'    ARR_TO_2D is the canonical normalization routine of the
'    library. Scalars become (1 x 1), 1D arrays become (N x 1)
'    columns, and 2D arrays are returned as-is.
'
'    Functions that require 2D input (ARR_2D_GET_COL,
'    ARR_2D_SPLIT_COLS) document this as a PRECONDITION in
'    their header.
'
' 3. RANGE HANDLING
'    Any function that accepts raw user input (Variant)
'    detects Range objects via TypeName and extracts .Value
'    before processing. Downstream logic never sees Range
'    objects.
'
' 4. EMPTY-VALUE SEMANTICS
'    Cleaning routines (ARR_1D_REMOVE_EMPTY) discard:
'
'      • Empty strings ("")
'      • Empty variants (IsEmpty)
'      • Error values (IsError)
'
'    When a cleaning routine produces no valid values, it
'    implicitly returns Empty rather than a zero-length array.
'
' 5. ERROR SIGNALING
'    Internal invariant violations raise errors via
'    Err.Raise ERR_*. Public routines intended
'    as worksheet UDFs return CVErr(...) instead.
'
' ------------------------------------------------------------
' DEPENDENCIES
' ------------------------------------------------------------
'
'   ERRORS module:
'     ERR_ARR_NOT_VECTOR, ERR_ARR_INVALID_DIMENSION,
'     ERR_ARR_COLINDEX_OUT_OF_RANGE, ERR_ARR_INVALID_VALUE
'
' ------------------------------------------------------------
' RELATED MODULES
' ------------------------------------------------------------
'
'   MATH module:
'     Uses ARR_TO_2D, ARR_SHAPE, ARR_TO_1D_LONG in
'     the Radix Engine (Normalize_Base, Normalize_Digits).
'
'   COMBINATORICS module:
'     Uses ARR_TO_2D, ARR_2D_SPLIT_COLS in
'     Cartesian_Build_Dimensions.
'     Uses ARR_TO_1D_STRING in COMB_UNRANK_VARLEN.
'
' ============================================================




' ============================================================
' FUNCTION: ARR_NDIM
'
' DESCRIPTION:
' The ARR_NDIM function returns the number of dimensions
' of a VBA array. It is conceptually similar to the ndim attribute
' in numerical libraries such as NumPy (Python).
'
' If the input is a Range, its value is extracted before analysis.
' If the input is not an array, the function returns 0.
'
' STRATEGY:
' 1. Normalize Range input
'    If the input is a Range, its .Value is extracted so the
'    analysis operates on the underlying array structure.
'
' 2. Detect non-array inputs
'    If the input is not an array, the function returns 0.
'    This covers scalars and single-cell ranges.
'
' 3. Probe each dimension safely
'    The function attempts to access UBound on increasing
'    dimension indices. When a dimension does not exist, VBA
'    raises an error, which is captured by the error handler.
'
' 4. Return the highest valid dimension
'    The last dimension successfully probed is returned as the
'    total number of dimensions. The upper probe limit is 60,
'    matching the VBA maximum.
'
' PARAMETERS:
'   inputArray  Variant   The value to inspect. Accepted forms:
'                           - Array of any dimensionality
'                           - Range (value is extracted via .Value)
'                           - Scalar or non-array (returns 0)
'
' RETURN VALUE:
' A Long representing the number of dimensions:
'   • 0 --> Not an array
'   • 1 --> 1D array
'   • 2 --> 2D array
'   • N --> N-dimensional array
'
' EXAMPLE:
' Dim arr(1 To 3, 0 To 4) As Variant
' ARR_NDIM(arr)   ' Returns: 2
'
'
' NOTES:
' - Works with arrays from ranges (convert Range to array) and literals
' - Safe for use in validation logic
' - A single-cell range returns a scalar value
'
' ============================================================

Public Function ARR_NDIM(ByVal inputArray As Variant) As Long
    
    If TypeName(inputArray) = "Range" Then
        inputArray = inputArray.value
    End If
    
    ' --------------------------------------------------------
    ' CASE 1: Not an array
    ' --------------------------------------------------------
    If Not IsArray(inputArray) Then
        ARR_NDIM = 0
        Exit Function
    End If
    
    ' --------------------------------------------------------
    ' Detect number of dimensions safely
    ' Strategy:
    ' Try accessing UBound for increasing dimensions
    ' Stop when an error occurs
    ' --------------------------------------------------------
    Dim dimensionIndex As Long
    
    On Error GoTo ExitFunction
    
    For dimensionIndex = 1 To 60
        
        Dim tmp As Long
        
        ' Attempt to access current dimension
        ' If it exists, no error is thrown
        tmp = UBound(inputArray, dimensionIndex)
        
        ' Update detected dimension count
        ARR_NDIM = dimensionIndex
        
    Next dimensionIndex

ExitFunction:
    On Error GoTo 0
    
End Function


' ============================================================
' FUNCTION: ARR_TO_2D
'
' DESCRIPTION:
' The ARR_TO_2D function normalizes any supported input into a
' 2D array and 1-based indexing. It is the canonical normalization
' routine of the library, ensuring that downstream functions
' always operate on a predictable 2D structure.
'
' Ranges, scalars and 1D arrays are wrapped or reshaped. Arrays
' that are already 2D (or higher) are returned unchanged.
'
' STRATEGY:
' 1. Handle Range input
'    If the input is a Range, its .Value is returned directly.
'    Excel ranges always expose their content as a 2D array.
'
' 2. Handle array input
'    If the input is already an array, its dimensionality is
'    inspected via ARR_NDIM:
'
'      • 1D arrays are reshaped into a (N x 1) column
'      • 2D arrays are returned as-is
'
' 3. Handle scalar input
'    If the input is neither a Range nor an array, it is wrapped
'    inside a (1 x 1) 2D array.
'
' 4. Return the normalized 2D array
'    The output is always a 2D array suitable for generic
'    downstream processing.
'
' PARAMETERS:
'   inputValue  Variant   The value to normalize. Accepted forms:
'                           - Range (single cell or multi-cell)
'                           - 1D array
'                           - 2D array (returned as-is)
'                           - Scalar
'
' RETURN VALUE:
' A Variant containing a 2D array.
'
' EXAMPLE:
' ARR_TO_2D(42)              ' Returns: (1 x 1) containing 42
' ARR_TO_2D(Array(1, 2, 3))  ' Returns: (3 x 1) column
' ARR_TO_2D(Range("A1:B3"))  ' Returns: (3 x 2) array
'
' ============================================================

Public Function ARR_TO_2D(ByVal inputValue As Variant) As Variant
    
    ' --------------------------------------------------------
    ' CASE 1: Input is a Range
    ' - Excel ranges always return a 2D array via .Value
    ' --------------------------------------------------------
    If TypeName(inputValue) = "Range" Then
        Dim rangeVal As Variant
        rangeVal = inputValue.value
        If IsArray(rangeVal) Then
            ARR_TO_2D = rangeVal
        Else
            Dim singleCell(1 To 1, 1 To 1) As Variant
            singleCell(1, 1) = rangeVal
            ARR_TO_2D = singleCell
        End If
        Exit Function
    End If
    
    ' --------------------------------------------------------
    ' CASE 2: Input is already an array
    ' --------------------------------------------------------
    If IsArray(inputValue) Then
        
        ' ----------------------------------------------------
        ' Subcase: 1D array
        ' Convert to 2D column (n x 1)
        ' ----------------------------------------------------
        If ARR_NDIM(inputValue) = 1 Then
            
            Dim i As Long
            Dim result() As Variant
            
            ' Create 2D array with one column
            ReDim result(1 To UBound(inputValue) - LBound(inputValue) + 1, 1 To 1)
            
            ' Copy values from 1D array into column
            For i = LBound(inputValue) To UBound(inputValue)
                result(i - LBound(inputValue) + 1, 1) = inputValue(i)
            Next i
            
            ARR_TO_2D = result
            Exit Function
            
        End If
        
        ' ----------------------------------------------------
        ' Subcase: already 2D (or higher)
        ' Return as-is
        ' ----------------------------------------------------
        ARR_TO_2D = inputValue
        Exit Function
        
    End If
    
    ' --------------------------------------------------------
    ' CASE 3: Scalar value
    ' Wrap into a 2D array (1 x 1)
    ' --------------------------------------------------------
    Dim temp(1 To 1, 1 To 1) As Variant
    temp(1, 1) = inputValue
    
    ARR_TO_2D = temp
    
End Function


' ============================================================
' FUNCTION: ARR_IS_1D
'
' DESCRIPTION:
' The ARR_IS_1D function returns True when the input is a
' one-dimensional array, and False otherwise. It is a thin
' semantic wrapper over ARR_NDIM.
'
' The check is performed on the ORIGINAL input structure, not
' on the normalized version produced by ARR_TO_2D.
'
' STRATEGY:
' 1. Retrieve the number of dimensions
'    The function delegates to ARR_NDIM(arr).
'
' 2. Compare against 1
'    The function returns True when the result is exactly 1.
'    Scalars, Ranges and multi-dimensional arrays return False.
'
' PARAMETERS:
'   arr   Variant   The value to check. Any input is accepted;
'                   non-arrays and scalars return False.
'
' RETURN VALUE:
' A Boolean indicating whether the input is a 1D array.
'
' EXAMPLE:
' Dim v() As Variant: ReDim v(1 To 5)
' ARR_IS_1D(v)                 ' Returns: True
' ARR_IS_1D(Range("A1:B3"))    ' Returns: False
'
' ============================================================

Public Function ARR_IS_1D(arr As Variant) As Boolean
    
    ' --------------------------------------------------------
    ' Check if the number of dimensions equals 1
    ' ARR_NDIM returns:
    ' 0 = not an array
    ' 1 = 1D array
    ' 2 = 2D array, etc.
    ' --------------------------------------------------------
    ARR_IS_1D = (ARR_NDIM(arr) = 1)
    
    ' --------------------------------------------------------
    ' Important:
    ' This evaluates the ORIGINAL structure of the input,
    ' not the normalized version produced by ARR_TO_2D
    ' --------------------------------------------------------
    
End Function


' ============================================================
' FUNCTION: ARR_IS_2D
'
' DESCRIPTION:
' The ARR_IS_2D function returns True when the input is a
' two-dimensional array, and False otherwise. It is a thin
' semantic wrapper over ARR_NDIM.
'
' The check is performed on the ORIGINAL input structure, not
' on the normalized version produced by ARR_TO_2D.
'
' STRATEGY:
' 1. Retrieve the number of dimensions
'    The function delegates to ARR_NDIM(arr).
'
' 2. Compare against 2
'    The function returns True when the result is exactly 2.
'    Scalars, 1D arrays and higher-dimensional arrays return False.
'
' PARAMETERS:
'   arr   Variant   The value to check. Any input is accepted;
'                   non-arrays and scalars return False.
'
' RETURN VALUE:
' A Boolean indicating whether the input is a 2D array.
'
' EXAMPLE:
' ARR_IS_2D(Range("A1:B3"))    ' Returns: True
' ARR_IS_2D(Array(1, 2, 3))    ' Returns: False
'
' ============================================================

Public Function ARR_IS_2D(arr As Variant) As Boolean
    
    ' --------------------------------------------------------
    ' Check if the input has exactly 2 dimensions
    ' Uses ARR_NDIM to determine structure
    ' --------------------------------------------------------
    ARR_IS_2D = (ARR_NDIM(arr) = 2)
    
    ' --------------------------------------------------------
    ' Important:
    ' This evaluates the ORIGINAL structure of the input,
    ' not the normalized version produced by ARR_TO_2D
    ' --------------------------------------------------------
    
End Function



' ============================================================
' FUNCTION: ARR_SHAPE
'
' DESCRIPTION:
' The ARR_SHAPE function returns the shape of a VBA array,
' meaning the size of each dimension. It is conceptually similar
' to the shape attribute in numerical libraries such as NumPy.
'
' If the input is not an array, the function implicitly returns
' Empty.
'
' STRATEGY:
' 1. Determine the number of dimensions
'    The function calls ARR_NDIM(arr) to identify how
'    many dimensions the input has.
'
' 2. Handle non-array inputs
'    If the number of dimensions is 0, the function exits and
'    implicitly returns Empty.
'
' 3. Initialize the result array
'    A 1-based Long array is created with one element per
'    dimension.
'
' 4. Calculate the size of each dimension
'    Size is computed as:
'
'      • Size = UBound(arr, i) - LBound(arr, i) + 1
'
'    This formula supports arrays with arbitrary lower bounds.
'
' 5. Return the shape
'    The function returns a Long array containing the size of
'    each dimension.
'
' PARAMETERS:
'   arr   Variant   The array to inspect. Any dimensionality
'                   is supported. Non-arrays return Empty.
'
' RETURN VALUE:
' A 1-based 1D Long array. Each element represents the size of
' a dimension.
'
' EXAMPLE:
' Dim arr(1 To 3, 0 To 4) As Variant
' ARR_SHAPE(arr)   ' Returns: (3, 5)
'
' ============================================================


Public Function ARR_SHAPE(arr As Variant) As Variant
    
    Dim dims As Long
    
    ' --------------------------------------------------------
    ' Get number of dimensions
    ' --------------------------------------------------------
    dims = ARR_NDIM(arr)
    
    ' --------------------------------------------------------
    ' If not an array, return Empty
    ' --------------------------------------------------------
    If dims = 0 Then Exit Function
    
    Dim result() As Long
    
    ' --------------------------------------------------------
    ' Initialize result array
    ' One element per dimension
    ' --------------------------------------------------------
    ReDim result(1 To dims)
    
    Dim i As Long
    
    ' --------------------------------------------------------
    ' Loop through each dimension
    ' Calculate size using:
    ' size = UBound - LBound + 1
    ' --------------------------------------------------------
    For i = 1 To dims
        result(i) = UBound(arr, i) - LBound(arr, i) + 1
    Next i
    
    ' --------------------------------------------------------
    ' Return shape array
    ' --------------------------------------------------------
    ARR_SHAPE = result
    
End Function





' ============================================================
' FUNCTION: ARR_SIZE
'
' DESCRIPTION:
' The ARR_SIZE function returns the total number of elements
' contained in a VBA array, regardless of the number of dimensions.
'
' The total size is calculated as the product of the sizes of
' all dimensions, making the function conceptually similar to
' the size or prod(shape) operations in numerical libraries.
'
' STRATEGY:
' 1. Retrieve the shape of the array
'    The function calls ARR_SHAPE(arr), which returns a one-
'    dimensional array where each element represents the size
'    of a corresponding dimension.
'
' 2. Initialize the total element counter
'    The counter is initialized to 1 so it can be used as a
'    multiplicative accumulator.
'
' 3. Multiply all dimensions
'    The function iterates over the shape array and multiplies
'    all dimension sizes together using:
'
'      • Total = dim1 × dim2 × ... × dimN
'
'    This approach works for arrays of any dimensionality.
'
' 4. Return the total number of elements
'    The final product is returned as a Long value.
'
' PARAMETERS:
'   arr   Variant   The array whose total element count is
'                   needed. Non-arrays return 0.
'
' RETURN VALUE:
' A Long representing the total number of elements in the array.
'
' EXAMPLE:
' Dim arr(1 To 3, 0 To 4) As Variant
' ARR_SIZE(arr)   ' Returns: 15
'
' ============================================================


Public Function ARR_SIZE(arr As Variant) As Long
    
    Dim shape As Variant
    
    ' --------------------------------------------------------
    ' Retrieve the shape of the array
    ' shape is a 1D array where each element represents
    ' the size of a dimension
    ' --------------------------------------------------------
    shape = ARR_SHAPE(arr)

    ' No-array case
    If IsEmpty(shape) Then
        ARR_SIZE = 0
        Exit Function
    End If
    
    Dim i As Long
    Dim total As Long
    
    ' --------------------------------------------------------
    ' Initialize total element counter
    ' --------------------------------------------------------
    total = 1
    
    ' --------------------------------------------------------
    ' Multiply all dimensions to get total number of elements
    ' Works for n-dimensional arrays
    ' --------------------------------------------------------
    For i = LBound(shape) To UBound(shape)
        total = total * shape(i)
    Next i
    
    ' --------------------------------------------------------
    ' Return total number of elements
    ' --------------------------------------------------------
    ARR_SIZE = total
    
End Function


' ============================================================
' FUNCTION: ARR_2D_GET_COL
'
' DESCRIPTION:
' The ARR_2D_GET_COL function extracts a single column from
' a 2D array and returns it as a 1D array. It is a building block
' for column-based transformations such as ARR_2D_SPLIT_COLS.
'
' PRECONDITION:
' The input must be a 2D array. Scalars and 1D arrays are not
' supported and will raise a runtime error. Use ARR_TO_2D to
' normalize inputs before calling.
'
' STRATEGY:
' 1. Retrieve the shape of the array
'    The function calls ARR_SHAPE once to obtain the full
'    structural description of the input. The result is used
'    for both validation and extraction.
'
' 2. Validate: must be a 2D array
'    If ARR_SHAPE returns Empty, the input is not an array.
'    If UBound(shape) <> 2, the input does not have exactly
'    two dimensions. Both cases raise ERR_ARR_INVALID_DIMENSION.
'
' 3. Validate the column index
'    The requested column must be >= 1 and <= shape(2).
'    Out-of-range indices raise ERR_ARR_COLINDEX_OUT_OF_RANGE.
'
' 4. Delegate extraction to ARR_2D_Get_Col_Raw
'    Once all preconditions are satisfied, the function
'    delegates to the internal implementation, passing the
'    array, column index, and row count already obtained
'    from the shape.
'
' RETURN VALUE:
' A 1-based 1D array containing the values of the specified
' column.
'
' EXAMPLE:
' Dim arr(1 To 3, 1 To 2) As Variant
' arr(1,1)="A": arr(2,1)="B": arr(3,1)="C"
' ARR_2D_GET_COL(arr, 1)   ' Returns: ("A", "B", "C")
'
' ============================================================

Public Function ARR_2D_GET_COL(arr As Variant, columnIndex As Long) As Variant
    
    Dim shape As Variant
    
    ' --------------------------------------------------------
    ' Get array shape (rows, columns)
    ' --------------------------------------------------------
    shape = ARR_SHAPE(arr)
    
    
    ' --------------------------------------------------------
    ' Validate: must be a 2D array
    ' --------------------------------------------------------
    If IsEmpty(shape) Then
        Err.Raise ERR_ARR_INVALID_DIMENSION, , "Input must be a 2D array"
    End If
    
    If UBound(shape) <> 2 Then
        Err.Raise ERR_ARR_INVALID_DIMENSION, , "Input must be a 2D array"
    End If
    ' --------------------------------------------------------
    
    
    Dim rowCount As Long
    rowCount = shape(1)
    
    ' --------------------------------------------------------
    ' Validate column index (1-based)
    ' --------------------------------------------------------
    If columnIndex < 1 Or columnIndex > shape(2) Then
        Err.Raise ERR_ARR_COLINDEX_OUT_OF_RANGE, , "Column index out of Range"
    End If
    

    ARR_2D_GET_COL = ARR_2D_Get_Col_Raw(arr, columnIndex, rowCount)
    
End Function


' ============================================================
' FUNCTION: ARR_2D_Get_Col_Raw
' CATEGORY: Array Utilities / Extraction (Internal)
'
' DESCRIPTION:
' The ARR_2D_Get_Col_Raw function extracts a single column
' from a 2D array and returns it as a 1-based 1D array.
'
' This is the raw internal implementation used by
' ARR_2D_GET_COL and ARR_2D_SPLIT_COLS. It performs
' no validation and assumes all preconditions are already
' satisfied by the caller.
'
' PRECONDITIONS (caller must guarantee):
'    arr must be a 2D array
'    arr must use 1-based indexing
'    columnIndex must be >= 1 and <= number of columns
'    rowCount must equal UBound(arr, 1)
'
' STRATEGY:
' 1. Allocate the result array
'    A 1-based 1D array is created with rowCount elements.
'
' 2. Extract values from the specified column
'    Iterates from row 1 to rowCount and copies each value
'    from arr(r, columnIndex) into result(r).
'
' RETURN VALUE:
' A 1-based 1D Variant array containing the values of the
' specified column.
'
' EXAMPLE:
' ' Called internally after validation:
' Dim col As Variant
' col = ARR_2D_Get_Col_Raw(arr, 2, 5)
' ' Returns column 2 of arr as a 1D array of 5 elements
'
' ============================================================

Private Function ARR_2D_Get_Col_Raw( _
    ByVal arr As Variant, _
    ByVal columnIndex As Long, _
    ByVal rowCount As Long _
) As Variant

    Dim result() As Variant
    Dim r As Long
    
    ' --------------------------------------------------------
    ' Initialize result array (1D)
    ' --------------------------------------------------------
    ReDim result(1 To rowCount)
    
    ' --------------------------------------------------------
    ' Extract values from specified column
    ' --------------------------------------------------------
    For r = 1 To rowCount
        result(r) = arr(r, columnIndex)
    Next r

    ARR_2D_Get_Col_Raw = result

End Function





' ============================================================
' FUNCTION: ARR_1D_REMOVE_EMPTY
'
' DESCRIPTION:
' The ARR_1D_REMOVE_EMPTY function removes empty values from
' a 1D array and returns a compact 1D array containing only the
' valid entries.
'
' Empty strings, Empty values and error values are all
' discarded. If no valid values remain, the function implicitly
' returns Empty.
'
' STRATEGY:
' 1. First pass: count valid values
'    The function iterates over the array once to count how
'    many elements are neither errors, empty strings, nor
'    Empty. This enables an exact allocation in the next step.
'
' 2. Handle the all-empty case
'    If no valid values were found, the function exits and
'    implicitly returns Empty.
'
' 3. Allocate the result array
'    A 1-based 1D array is created with the exact size needed.
'
' 4. Second pass: copy valid values
'    The function iterates again and copies only the valid
'    entries into the result, preserving their relative order.
'
' PARAMETERS:
'   arr   Variant   A 1D array to clean. May use any lower
'                   bound. Error values, empty strings and
'                   Empty variants are all discarded.
'
' RETURN VALUE:
' A 1-based 1D array with only the non-empty values, or Empty
' if no valid entries remain.
'
' EXAMPLE:
' Dim v: v = Array("A", "", Empty, "B")
' ARR_1D_REMOVE_EMPTY(v)   ' Returns: ("A", "B")
'
' ============================================================

Public Function ARR_1D_REMOVE_EMPTY(arr As Variant) As Variant
    
    Dim temp() As Variant
    Dim i As Long, count As Long
    
    ' --------------------------------------------------------
    ' Count non-empty values (exclude empty strings "")
    ' --------------------------------------------------------
    For i = LBound(arr) To UBound(arr)
        If Not IsError(arr(i)) Then
            If arr(i) <> "" And Not IsEmpty(arr(i)) Then
                count = count + 1
            End If
        End If
    Next i
    
    ' --------------------------------------------------------
    ' If no valid values, return Empty
    ' --------------------------------------------------------
    If count = 0 Then Exit Function
    
    ' --------------------------------------------------------
    ' Initialize result array with exact size
    ' --------------------------------------------------------
    ReDim temp(1 To count)
    
    count = 0
    
    ' --------------------------------------------------------
    ' Copy non-empty values into result array
    ' --------------------------------------------------------
    For i = LBound(arr) To UBound(arr)
        If Not IsError(arr(i)) Then
            If arr(i) <> "" And Not IsEmpty(arr(i)) Then
                count = count + 1
                temp(count) = arr(i)
            End If
        End If
    Next i
    
    ' --------------------------------------------------------
    ' Return cleaned array
    ' --------------------------------------------------------
    ARR_1D_REMOVE_EMPTY = temp
    
End Function


' ============================================================
' FUNCTION: ARR_2D_SPLIT_COLS
'
' DESCRIPTION:
' The ARR_2D_SPLIT_COLS function splits a 2D array into a
' collection of 1D arrays, one per column. It is used as a
' preprocessing step for routines that treat each column as an
' independent dimension (e.g., the Cartesian product engine).
'
' Empty values inside each column are removed via
' ARR_1D_REMOVE_EMPTY.
'
' PRECONDITION:
' The input must be a 2D array. Use ARR_TO_2D to normalize
' inputs before calling.
'
' STRATEGY:
' 1. Retrieve the shape of the array
'    The function calls ARR_SHAPE to determine the column
'    count.
'
' 2. Allocate the result container
'    A 1-based 1D array is created with one slot per column.
'
' 3. Process each column
'    For every column:
'
'      • Extract the column with ARR_2D_Get_Col_Raw
'      • Remove empty values with ARR_1D_REMOVE_EMPTY
'      • Store the cleaned column in the corresponding slot
'
'    Columns that become fully empty are left as Empty in the
'    result.
'
' PARAMETERS:
'   arr   Variant   A 2D array to split. Must use 1-based
'                   indexing. Use ARR_TO_2D to normalize
'                   before calling.
'
' RETURN VALUE:
' A 1-based 1D array of 1D arrays. Each element represents one
' cleaned column of the input.
'
' EXAMPLE:
' Dim arr(1 To 2, 1 To 2) As Variant
' arr(1,1)="A": arr(2,1)="B"
' arr(1,2)=1:   arr(2,2)=2
' ARR_2D_SPLIT_COLS(arr)
' ' Returns: array containing ("A","B") and (1,2)
'
' ============================================================

Public Function ARR_2D_SPLIT_COLS(arr As Variant) As Variant
    
    Dim shape As Variant
    
    ' --------------------------------------------------------
    ' Get array shape (rows, columns)
    ' --------------------------------------------------------
    shape = ARR_SHAPE(arr)
    
    
    ' --------------------------------------------------------
    ' Validate: must be a 2D array
    ' --------------------------------------------------------
    If IsEmpty(shape) Then
        Err.Raise ERR_ARR_INVALID_DIMENSION, , "Input must be a 2D array"
    End If
    
    If UBound(shape) <> 2 Then
        Err.Raise ERR_ARR_INVALID_DIMENSION, , "Input must be a 2D array"
    End If
    ' --------------------------------------------------------
    
    
    Dim colCount As Long
    colCount = shape(2)
    
    Dim result() As Variant
    
    ' --------------------------------------------------------
    ' Initialize result array (1-based)
    ' Each element will store a column (1D array)
    ' --------------------------------------------------------
    ReDim result(1 To colCount)
    
    Dim c As Long
    
    ' --------------------------------------------------------
    ' Process each column
    ' --------------------------------------------------------
    For c = 1 To colCount
        
        Dim colData As Variant
        
        ' Extract column as 1D array
        colData = ARR_2D_Get_Col_Raw(arr, c, shape(1))
        
        ' Remove empty values
        colData = ARR_1D_REMOVE_EMPTY(colData)
        
        ' Only store non-empty columns
        If Not IsEmpty(colData) Then
            result(c) = colData
        End If
        
    Next c
    
    ' --------------------------------------------------------
    ' Return array of column arrays
    ' --------------------------------------------------------
    ARR_2D_SPLIT_COLS = result
    
End Function


' ============================================================
' FUNCTION: ARR_TO_1D_LONG
'
' DESCRIPTION:
' The ARR_TO_1D_LONG function converts any supported input
' into a 1-based 1D array of Long values. It provides a single
' canonical representation for routines that require a flat
' numeric sequence (e.g., the Radix engine).
'
' Supported inputs:
'   • Scalar --> single-element array
'   • Row vector (1 x N) --> N elements
'   • Column vector (N x 1) --> N elements
'
' Rectangular 2D matrices are rejected with a runtime error,
' since they cannot be unambiguously flattened.
'
' STRATEGY:
' 1. Normalize input to 2D
'    The function calls ARR_TO_2D so that all inputs share a
'    common structural representation.
'
' 2. Retrieve the shape
'    Rows and columns are obtained via ARR_SHAPE.
'
' 3. Dispatch by shape
'    The function branches on the shape of the normalized
'    input:
'
'      • (1 x N) --> row vector case (covers scalar when N = 1)
'      • (N x 1) --> column vector case
'      • otherwise --> raise error
'
' 4. Copy values with CLng coercion
'    Each element is converted to Long before being stored,
'    enforcing numeric consistency for downstream logic.
'
' RETURN VALUE:
' A 1-based 1D Long array.
'
' EXAMPLE:
' ARR_TO_1D_LONG(42)                    ' Returns: (42)
' ARR_TO_1D_LONG(Array(1, 2, 3))        ' Returns: (1, 2, 3)
' ARR_TO_1D_LONG(Range("A1:A5"))        ' Returns: 1D Long array
'
' ============================================================

Public Function ARR_TO_1D_LONG(ByVal inputValue As Variant) As Long()

    Dim result() As Long
    Dim i As Long
    
    Dim arr As Variant
    Dim shape As Variant
    
    ' --------------------------------------------------------
    ' Normalize to 2D array always
    ' --------------------------------------------------------
    arr = ARR_TO_2D(inputValue)
    
    shape = ARR_SHAPE(arr)
    
    Dim rows As Long, cols As Long
    rows = shape(1)
    cols = shape(2)
    
    
    ' --------------------------------------------------------
    ' CASE 1: Row vector (covers scalar when cols = 1)
    ' --------------------------------------------------------
    If rows = 1 Then
        ReDim result(1 To cols)
        
        For i = 1 To cols
            result(i) = CLng(arr(1, i))
        Next i
        
        ARR_TO_1D_LONG = result
        Exit Function
    End If
    
    ' --------------------------------------------------------
    ' CASE 2: Column vector
    ' --------------------------------------------------------
    If cols = 1 Then
        ReDim result(1 To rows)
        
        For i = 1 To rows
            result(i) = CLng(arr(i, 1))
        Next i
        
        ARR_TO_1D_LONG = result
        Exit Function
    End If
    
    Err.Raise ERR_ARR_NOT_VECTOR, , "Value must be scalar or 1D"

End Function




' ============================================================
' FUNCTION: ARR_TO_1D_STRING
'
' DESCRIPTION:
' The ARR_TO_1D_STRING function converts any supported input
' into a 1-based 1D array of String values. It provides a single
' canonical representation for routines that require a flat
' string sequence (e.g., COMB_UNRANK_VARLEN).
'
' Supported inputs:
'   • Scalar --> single-element array
'   • Row vector (1 x N) --> N elements
'   • Column vector (N x 1) --> N elements
'
' Rectangular 2D matrices are rejected with a runtime error,
' since they cannot be unambiguously flattened.
' Error values (e.g. #N/A, #DIV/0!) are rejected with a
' runtime error, since CStr would silently convert them to
' strings like "Error 2042".
'
' STRATEGY:
' 1. Normalize input to 2D
'    The function calls ARR_TO_2D so that all inputs share a
'    common structural representation.
'
' 2. Retrieve the shape
'    Rows and columns are obtained via ARR_SHAPE.
'
' 3. Dispatch by shape
'    The function branches on the shape of the normalized
'    input:
'
'      • (1 x N) --> row vector case (covers scalar when N = 1)
'      • (N x 1) --> column vector case
'      • otherwise --> raise error
'
' 4. Copy values with CStr coercion
'    Each element is validated against IsError before
'    conversion, then stored as String.
'
' RETURN VALUE:
' A 1-based 1D String array.
'
' EXAMPLE:
' ARR_TO_1D_STRING("hello")             ' Returns: ("hello")
' ARR_TO_1D_STRING(Array("a","b","c"))  ' Returns: ("a", "b", "c")
' ARR_TO_1D_STRING(Range("A1:A3"))      ' Returns: 1D String array
'
' ============================================================

Public Function ARR_TO_1D_STRING(ByVal inputValue As Variant) As String()

    Dim result() As String
    Dim i As Long
    
    Dim arr As Variant
    Dim shape As Variant
    
    ' --------------------------------------------------------
    ' Normalize to 2D array always
    ' --------------------------------------------------------
    arr = ARR_TO_2D(inputValue)
    
    shape = ARR_SHAPE(arr)
    
    Dim rows As Long, cols As Long
    rows = shape(1)
    cols = shape(2)
    
    ' --------------------------------------------------------
    ' CASE 1: Row vector (covers scalar when cols = 1)
    ' --------------------------------------------------------
    If rows = 1 Then
        ReDim result(1 To cols)
        
        For i = 1 To cols
            If IsError(arr(1, i)) Then Err.Raise ERR_ARR_INVALID_VALUE, , _
                "Array contains error value"
            result(i) = CStr(arr(1, i))
        Next i
        
        ARR_TO_1D_STRING = result
        Exit Function
    End If
    
    ' --------------------------------------------------------
    ' CASE 2: Column vector
    ' --------------------------------------------------------
    If cols = 1 Then
        ReDim result(1 To rows)
        
        For i = 1 To rows
            If IsError(arr(i, 1)) Then Err.Raise ERR_ARR_INVALID_VALUE, , _
                "Array contains error value"
            result(i) = CStr(arr(i, 1))
        Next i
        
        ARR_TO_1D_STRING = result
        Exit Function
    End If
    
    Err.Raise ERR_ARR_NOT_VECTOR, , "Value must be scalar or 1D"

End Function
